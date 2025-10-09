# recordings_automation

This repository contains tooling for collecting Outlook data files from Windows
machines across a LAN. The automation uses SMB to scan the standard Outlook
folders for each configured user and copies any `.pst` or `.ost` archives into a
local directory for safekeeping.

## Requirements

* Python 3.9+
* [`pysmb`](https://pysmb.readthedocs.io/) (install via `pip install -r requirements.txt`)

## Configuration

Use the provided [`configs/sample_config.json`](configs/sample_config.json) as a
starting point. The configuration file describes the controller hostname along
with each remote Windows workstation you need to scan:

```json
{
  "client_name": "automation-controller",
  "hosts": [
    {
      "host": "192.168.1.50",
      "server_name": "WS-ACCOUNTING",
      "username": "domain\\\\backupadmin",
      "password": "REPLACE_WITH_PASSWORD",
      "domain": "DOMAIN",
      "shares": ["C$"],
      "users": [
        { "username": "jane.doe" },
        {
          "username": "john.smith",
          "search_roots": [
            "Users/{user}/Documents/Outlook Files",
            "Users/{user}/AppData/Local/Microsoft/Outlook"
          ]
        }
      ]
    }
  ]
}
```

* `client_name` – NetBIOS name of the machine running the collector.
* `hosts` – List of remote machines to inspect.
  * `host` – IP address or DNS name of the remote machine.
  * `server_name` – NetBIOS name of the remote server.
  * `username`/`password`/`domain` – Credentials with permission to read the
    desired shares (typically an administrator account).
  * `shares` – SMB shares to inspect. Defaults to `C$` if omitted.
  * `users` – Windows users whose Outlook files should be collected. Each user
    can override the list of search directories with `search_roots` if needed.

All search roots support `{user}` placeholders that are replaced with the user's
name at runtime. The defaults cover the standard Outlook directories under
`AppData` and `Documents`.

## Usage

After configuring your hosts, run the collector with:

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python src/outlook_collector.py --config configs/sample_config.json --destination ./outlook_backups
```

The destination folder will contain subdirectories for each host, share, and
relative Outlook path. Subsequent runs will overwrite existing files, so archive
or move collected data as needed.

## Safety considerations

* Use dedicated backup credentials with read-only permissions where possible.
* Ensure that copying large OST/PST files does not overwhelm your storage
  capacity or network bandwidth.
* Collected mail archives can contain sensitive data; secure the destination
  directory with appropriate access controls.
