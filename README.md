# recordings_automation
This repo for automation recrodings to the platfrom this is also the private repo

## Tools

### `scripts/reduce_pdf_size.py`

This command line utility reduces the size of a PDF to a custom target value. It
uses [`pikepdf`](https://pikepdf.readthedocs.io/en/latest/) to optimise embedded
images and iteratively adjusts the compression quality until the resulting file
falls within a configurable tolerance of the requested size.

```bash
python scripts/reduce_pdf_size.py input.pdf output.pdf --target 500KB
```

Additional options:

- `--tolerance` – Acceptable relative difference (default `0.05` for ±5%).
- `--min-quality` / `--max-quality` – Bounds for image quality (1–100).

> **Note:** Install dependencies with `pip install pikepdf` before running the
> script.
