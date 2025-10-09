"""Utilities to collect Outlook data files from Windows hosts on a LAN.

The script uses SMB to connect to remote machines and copy PST/OST files
into a local directory for archiving or analysis. See README for usage.
"""
from __future__ import annotations

import argparse
import json
import logging
import os
import platform
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Iterator, List, Optional

try:
    from smb.SMBConnection import SMBConnection
    from smb.base import NotConnectedError, OperationFailure
except ImportError as exc:  # pragma: no cover - handled at runtime
    raise SystemExit(
        "The pysmb package is required to run this script. Install it with 'pip install pysmb'."
    ) from exc

OUTLOOK_EXTENSIONS = (".pst", ".ost")


def parse_args(argv: Optional[List[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Collect Outlook PST/OST files from Windows hosts over SMB",
    )
    parser.add_argument(
        "--config",
        required=True,
        type=Path,
        help="Path to a JSON configuration file describing the target hosts.",
    )
    parser.add_argument(
        "--destination",
        required=True,
        type=Path,
        help="Local directory where copied Outlook files will be stored.",
    )
    parser.add_argument(
        "--log-level",
        default="INFO",
        choices=["CRITICAL", "ERROR", "WARNING", "INFO", "DEBUG"],
        help="Verbosity of logging output.",
    )
    return parser.parse_args(argv)


@dataclass
class HostUser:
    """Definition of a Windows user whose Outlook files should be collected."""

    username: str
    search_roots: List[str] = field(
        default_factory=lambda: [
            r"Users/{user}/AppData/Local/Microsoft/Outlook",
            r"Users/{user}/Documents/Outlook Files",
        ]
    )


@dataclass
class HostConfig:
    """Configuration describing how to connect to a Windows host via SMB."""

    host: str
    server_name: str
    username: str
    password: str
    shares: List[str] = field(default_factory=lambda: ["C$"])
    domain: str = ""
    port: int = 445
    users: List[HostUser] = field(default_factory=list)

    @classmethod
    def from_dict(cls, payload: Dict) -> "HostConfig":
        users_payload = payload.get("users", [])
        users: List[HostUser] = []
        for user_payload in users_payload:
            user = HostUser(username=user_payload["username"])
            if user_payload.get("search_roots"):
                user.search_roots = user_payload["search_roots"]
            users.append(user)
        return cls(
            host=payload["host"],
            server_name=payload["server_name"],
            username=payload["username"],
            password=payload["password"],
            shares=payload.get("shares", ["C$"]),
            domain=payload.get("domain", ""),
            port=int(payload.get("port", 445)),
            users=users,
        )


def load_config(path: Path) -> Dict:
    with path.open("r", encoding="utf-8") as f:
        return json.load(f)


def connect(host: HostConfig, client_name: str) -> SMBConnection:
    connection = SMBConnection(
        host.username,
        host.password,
        client_name,
        host.server_name,
        use_ntlm_v2=True,
        domain=host.domain,
        is_direct_tcp=True,
    )
    logging.debug("Connecting to %s (%s:%s)", host.server_name, host.host, host.port)
    if not connection.connect(host.host, host.port):
        raise ConnectionError(f"Unable to connect to host {host.server_name} at {host.host}:{host.port}")
    return connection


def remote_walk(connection: SMBConnection, share: str, base_path: str) -> Iterator[str]:
    normalized = base_path.strip("/\\").replace("\\", "/")
    path_for_listing = normalized or "/"
    try:
        entries = connection.listPath(share, path_for_listing)
    except OperationFailure as exc:
        logging.warning("Failed to list path %s on share %s: %s", path_for_listing, share, exc)
        return

    for entry in entries:
        if entry.filename in {".", ".."}:
            continue
        remote_path = f"{normalized}/{entry.filename}" if normalized else entry.filename
        if entry.isDirectory:
            yield from remote_walk(connection, share, remote_path)
        elif entry.filename.lower().endswith(OUTLOOK_EXTENSIONS):
            yield remote_path


def collect_from_host(connection: SMBConnection, host: HostConfig, destination: Path) -> None:
    if not host.users:
        logging.warning("No users configured for host %s; skipping", host.server_name)
        return
    for share in host.shares:
        for user in host.users:
            for root in user.search_roots:
                remote_root = root.format(user=user.username)
                logging.info(
                    "Scanning %s share %s for Outlook files in %s", host.server_name, share, remote_root
                )
                for remote_path in remote_walk(connection, share, remote_root):
                    store_remote_file(connection, host, share, remote_path, destination)


def store_remote_file(
    connection: SMBConnection,
    host: HostConfig,
    share: str,
    remote_path: str,
    destination_root: Path,
) -> None:
    relative_path = Path(remote_path.strip("/").replace("\\", "/"))
    destination_dir = destination_root / host.server_name / share / relative_path.parent
    destination_dir.mkdir(parents=True, exist_ok=True)
    destination_file = destination_dir / relative_path.name

    remote_display = f"/{remote_path.lstrip('/')}"
    logging.info("Copying %s:%s%s -> %s", host.server_name, share, remote_display, destination_file)
    try:
        with destination_file.open("wb") as fh:
            connection.retrieveFile(share, remote_path, fh)
    except (OperationFailure, NotConnectedError) as exc:
        logging.error("Failed to copy %s from %s: %s", remote_path, host.server_name, exc)


def configure_logging(level: str) -> None:
    logging.basicConfig(
        level=getattr(logging, level.upper()),
        format="%(asctime)s [%(levelname)s] %(message)s",
    )


def main(argv: Optional[List[str]] = None) -> None:
    args = parse_args(argv)
    configure_logging(args.log_level)

    config_payload = load_config(args.config)
    client_name = config_payload.get("client_name", platform.node() or os.getenv("COMPUTERNAME", "collector"))
    hosts = [HostConfig.from_dict(item) for item in config_payload.get("hosts", [])]
    if not hosts:
        raise SystemExit("No hosts defined in configuration file")

    destination_root = args.destination
    destination_root.mkdir(parents=True, exist_ok=True)

    for host_config in hosts:
        logging.info("Processing host %s", host_config.server_name)
        try:
            conn = connect(host_config, client_name)
        except Exception as exc:
            logging.error("Failed to connect to host %s: %s", host_config.server_name, exc)
            continue

        try:
            collect_from_host(conn, host_config, destination_root)
        except Exception as exc:
            logging.error("Failed to process host %s: %s", host_config.server_name, exc)
        finally:
            try:
                conn.close()
            except Exception:
                logging.debug("Error closing connection to %s", host_config.server_name, exc_info=True)


if __name__ == "__main__":
    main()
