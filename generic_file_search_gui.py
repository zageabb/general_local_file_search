#!/usr/bin/env python3
"""Launcher for Generic Local File Search v2.3.0.

Loads the preserved v2.2.0 application from the release ZIP, installs the v2.3
structured LLM search layer, then starts the Tkinter GUI.
"""
from pathlib import Path
from types import ModuleType
from zipfile import ZipFile
import sys

from v2_3_search_logic import install

PACKAGE = Path(__file__).with_name("generic_local_file_search_v2.2.0.zip")
MEMBER = "generic_file_search_gui.py"


def load_core() -> ModuleType:
    if not PACKAGE.exists():
        raise SystemExit(f"Missing preserved core package: {PACKAGE}")
    with ZipFile(PACKAGE, "r") as archive:
        source = archive.read(MEMBER)

    module = ModuleType("generic_file_search_core")
    module.__file__ = str(PACKAGE) + "/" + MEMBER
    module.__package__ = None
    sys.modules[module.__name__] = module
    exec(compile(source, module.__file__, "exec"), module.__dict__, module.__dict__)
    return module


def main() -> int:
    core = load_core()
    install(core)
    return core.main()


if __name__ == "__main__":
    raise SystemExit(main())
