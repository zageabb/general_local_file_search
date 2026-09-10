#!/usr/bin/env python3
"""
Bootstrap launcher for Generic Local File Search v2.2.0.

The exact v2.2.0 source package is stored in
`generic_local_file_search_v2.2.0.zip`. This launcher executes the packaged
`generic_file_search_gui.py` directly from that archive so the repository can
be cloned and run immediately while preserving the original package bytes.
"""
from pathlib import Path
from zipfile import ZipFile

PACKAGE = Path(__file__).with_name("generic_local_file_search_v2.2.0.zip")
MEMBER = "generic_file_search_gui.py"


def main() -> None:
    if not PACKAGE.exists():
        raise SystemExit(f"Missing package: {PACKAGE}")

    with ZipFile(PACKAGE, "r") as archive:
        source = archive.read(MEMBER)

    namespace = {
        "__name__": "__main__",
        "__file__": str(Path(__file__).resolve()),
        "__package__": None,
    }
    exec(compile(source, MEMBER, "exec"), namespace, namespace)


if __name__ == "__main__":
    main()
