"""Entrypoint for `python -m docx_renderer`."""

from .cli import main

if __name__ == "__main__":
    raise SystemExit(main())
