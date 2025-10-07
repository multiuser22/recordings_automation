#!/usr/bin/env python3
"""Compress PDF files to a target size using pikepdf.

The script mimics the behaviour of online PDF compressors by
progressively lowering the quality of embedded images until the
requested file size is achieved (within a tolerance).
"""
from __future__ import annotations

import argparse
import math
import re
import shutil
import sys
import tempfile
from pathlib import Path
from typing import Optional

import pikepdf

DEFAULT_TOLERANCE = 0.05  # 5%
DEFAULT_MIN_QUALITY = 20
DEFAULT_MAX_QUALITY = 95

SIZE_PATTERN = re.compile(r"^(?P<value>\d+(?:\.\d+)?)(?P<unit>[KMG]?B)?$", re.IGNORECASE)


class SizeParseError(ValueError):
    """Raised when the provided size string cannot be parsed."""


def parse_size(size_str: str) -> int:
    """Convert a human readable size string to bytes.

    Examples
    --------
    "500KB" -> 512000
    "0.5MB" -> 524288
    "100" -> 100 (bytes)
    """

    match = SIZE_PATTERN.match(size_str.strip())
    if not match:
        raise SizeParseError(f"Unable to parse size value: {size_str!r}")

    value = float(match.group("value"))
    unit = (match.group("unit") or "B").upper()

    multiplier = {
        "B": 1,
        "KB": 1024,
        "MB": 1024 ** 2,
        "GB": 1024 ** 3,
    }[unit]

    bytes_value = int(value * multiplier)
    if bytes_value <= 0:
        raise SizeParseError("Target size must be positive")

    return bytes_value


def human_readable_size(num_bytes: int) -> str:
    if num_bytes <= 0:
        return "0 B"
    units = ["B", "KB", "MB", "GB"]
    magnitude = min(int(math.log(num_bytes, 1024)), len(units) - 1)
    value = num_bytes / (1024 ** magnitude)
    return f"{value:.2f} {units[magnitude]}"


def compress_pdf(
    input_path: Path,
    output_path: Path,
    target_size: int,
    tolerance: float = DEFAULT_TOLERANCE,
    min_quality: int = DEFAULT_MIN_QUALITY,
    max_quality: int = DEFAULT_MAX_QUALITY,
    max_iterations: int = 8,
) -> tuple[Path, int, bool]:
    """Compress ``input_path`` into ``output_path``.

    Returns a tuple with (final_output_path, final_size_in_bytes, reached_target).
    """

    if not input_path.exists():
        raise FileNotFoundError(f"Input PDF not found: {input_path}")

    if tolerance <= 0 or tolerance >= 1:
        raise ValueError("tolerance must be between 0 and 1 (exclusive)")

    if min_quality < 1 or max_quality > 100 or min_quality >= max_quality:
        raise ValueError("Invalid quality range")

    input_size = input_path.stat().st_size
    if input_size <= target_size:
        # Just copy the file if it's already below the target size.
        shutil.copy2(input_path, output_path)
        return output_path, input_size, True

    best_candidate: Optional[Path] = None
    best_candidate_size: Optional[int] = None
    fallback_result: Optional[Path] = None
    fallback_result_size: Optional[int] = None

    low, high = min_quality, max_quality
    reached_target = False

    iterations = 0
    while low <= high and iterations < max_iterations:
        iterations += 1
        quality = (low + high) // 2
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp_file:
            temp_output = Path(tmp_file.name)

        try:
            with pikepdf.open(input_path) as pdf:
                pdf.save(
                    temp_output,
                    linearize=True,
                    optimize_images=True,
                    image_quality=quality,
                )
        except pikepdf.PdfError as exc:
            temp_output.unlink(missing_ok=True)
            raise RuntimeError(f"Failed to process PDF: {exc}") from exc

        size = temp_output.stat().st_size

        keep_temp = False

        if size <= target_size * (1 + tolerance):
            if best_candidate is None or size > (best_candidate_size or 0):
                if best_candidate is not None and best_candidate.exists():
                    best_candidate.unlink(missing_ok=True)
                best_candidate = temp_output
                best_candidate_size = size
                keep_temp = True
            reached_target = size <= target_size
            low = quality + 1
        else:
            high = quality - 1

        if not keep_temp and best_candidate is None:
            if fallback_result_size is None or size < fallback_result_size:
                if fallback_result is not None and fallback_result.exists():
                    fallback_result.unlink(missing_ok=True)
                fallback_result = temp_output
                fallback_result_size = size
                keep_temp = True

        if not keep_temp:
            temp_output.unlink(missing_ok=True)

    result_path: Optional[Path]
    result_size: Optional[int]
    if best_candidate is not None:
        result_path = best_candidate
        result_size = best_candidate_size
    else:
        result_path = fallback_result
        result_size = fallback_result_size

    if result_path is None or not result_path.exists():
        raise RuntimeError("Unable to generate compressed PDF")

    shutil.move(result_path, output_path)
    return output_path, result_size or output_path.stat().st_size, reached_target


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Reduce PDF file size using pikepdf")
    parser.add_argument("input", type=Path, help="Path to the input PDF file")
    parser.add_argument("output", type=Path, help="Path to the output PDF file")
    parser.add_argument(
        "--target",
        required=True,
        help="Desired maximum size (e.g. 500KB, 1.5MB)",
    )
    parser.add_argument(
        "--tolerance",
        type=float,
        default=DEFAULT_TOLERANCE,
        help="Acceptable relative difference between the result size and the target (default: 0.05 for 5%)",
    )
    parser.add_argument(
        "--min-quality",
        type=int,
        default=DEFAULT_MIN_QUALITY,
        help="Lower bound for image quality used during compression",
    )
    parser.add_argument(
        "--max-quality",
        type=int,
        default=DEFAULT_MAX_QUALITY,
        help="Upper bound for image quality used during compression",
    )

    args = parser.parse_args(argv)

    try:
        target_size = parse_size(args.target)
    except SizeParseError as exc:
        parser.error(str(exc))
        return 2

    output_path = args.output
    if output_path.exists() and output_path.is_dir():
        output_path = output_path / args.input.name

    output_path.parent.mkdir(parents=True, exist_ok=True)

    final_path, final_size, reached_target = compress_pdf(
        input_path=args.input,
        output_path=output_path,
        target_size=target_size,
        tolerance=args.tolerance,
        min_quality=args.min_quality,
        max_quality=args.max_quality,
    )

    print(
        "Compressed PDF saved to",
        final_path,
        f"(size: {human_readable_size(final_size)})",
    )
    if not reached_target:
        print(
            "Warning: Unable to reach the requested target size exactly,",
            "but the output is within the tolerance window.",
            file=sys.stderr,
        )
    return 0 if reached_target else 1


if __name__ == "__main__":
    sys.exit(main())
