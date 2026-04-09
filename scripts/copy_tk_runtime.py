from __future__ import annotations

import shutil
import sys
from pathlib import Path


def copy_file(source: Path, target: Path) -> None:
    if not source.exists():
        raise FileNotFoundError(f"Required Tk runtime file is missing: {source}")
    target.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, target)
    print(f"Copied {source.name} -> {target}")


def copy_tree(source: Path, target: Path) -> None:
    if not source.exists():
        raise FileNotFoundError(f"Required Tk runtime directory is missing: {source}")
    shutil.copytree(source, target, dirs_exist_ok=True)
    print(f"Copied {source} -> {target}")


def main() -> int:
    project_root = Path(__file__).resolve().parents[1]
    bundle_root = project_root / "build" / "octopusdetool" / "windows" / "app" / "src"

    if not bundle_root.exists():
        raise FileNotFoundError(
            "Briefcase Windows support bundle was not found. Run `briefcase create windows` first."
        )

    python_root = Path(sys.base_prefix)
    dll_root = python_root / "DLLs"
    lib_root = python_root / "Lib"
    tcl_root = python_root / "tcl"

    copy_file(dll_root / "_tkinter.pyd", bundle_root / "_tkinter.pyd")
    copy_file(dll_root / "tcl86t.dll", bundle_root / "tcl86t.dll")
    copy_file(dll_root / "tk86t.dll", bundle_root / "tk86t.dll")
    copy_tree(lib_root / "tkinter", bundle_root / "tkinter")
    copy_tree(tcl_root / "tcl8.6", bundle_root / "tcl" / "tcl8.6")
    copy_tree(tcl_root / "tk8.6", bundle_root / "tcl" / "tk8.6")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
