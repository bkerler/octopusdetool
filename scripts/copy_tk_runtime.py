from __future__ import annotations

import shutil
import sys
from pathlib import Path

APP_NAME = "octopusdetool"


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


def copy_windows_tk_runtime(project_root: Path) -> None:
    bundle_root = project_root / "build" / APP_NAME / "windows" / "app" / "src"

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
    copy_file(dll_root / "zlib1.dll", bundle_root / "zlib1.dll")
    copy_tree(lib_root / "tkinter", bundle_root / "tkinter")
    copy_tree(tcl_root / "tcl8.6", bundle_root / "tcl" / "tcl8.6")
    copy_tree(tcl_root / "tk8.6", bundle_root / "tcl" / "tk8.6")
    copy_tree(tcl_root / "dde1.4", bundle_root / "tcl" / "dde1.4")
    copy_tree(tcl_root / "reg1.3", bundle_root / "tcl" / "reg1.3")


def find_macos_tkinter_package() -> Path:
    try:
        import tkinter
    except ImportError as exc:
        raise FileNotFoundError(
            "The current Python interpreter does not provide `tkinter`. "
            "Use a macOS Python build with Tcl/Tk support before packaging."
        ) from exc

    return Path(tkinter.__file__).resolve().parent


def find_macos_tkinter_extension() -> Path:
    try:
        import _tkinter
    except ImportError as exc:
        raise FileNotFoundError(
            "The current Python interpreter does not provide `_tkinter`. "
            "Use a macOS Python build with Tcl/Tk support before packaging."
        ) from exc

    return Path(_tkinter.__file__).resolve()


def find_macos_stdlib_targets(project_root: Path) -> list[Path]:
    build_root = project_root / "build" / APP_NAME / "macos" / "xcode"

    if not build_root.exists():
        raise FileNotFoundError(
            "Briefcase macOS Xcode project was not found. Run `briefcase create macOS Xcode` first."
        )

    stdlib_roots = sorted(
        {
            path.resolve()
            for path in build_root.glob("**/python-stdlib/lib/python*")
            if path.is_dir()
        }
    )

    if not stdlib_roots:
        raise FileNotFoundError(
            "No macOS python-stdlib folders were found. Run `briefcase create macOS Xcode` first."
        )

    return stdlib_roots


def copy_macos_tk_runtime(project_root: Path) -> None:
    tkinter_package = find_macos_tkinter_package()
    tkinter_extension = find_macos_tkinter_extension()

    for stdlib_root in find_macos_stdlib_targets(project_root):
        copy_tree(tkinter_package, stdlib_root / "tkinter")
        copy_file(
            tkinter_extension,
            stdlib_root / "lib-dynload" / tkinter_extension.name,
        )


def main() -> int:
    project_root = Path(__file__).resolve().parents[1]

    if sys.platform == "win32":
        copy_windows_tk_runtime(project_root)
    elif sys.platform == "darwin":
        copy_macos_tk_runtime(project_root)
    else:
        raise RuntimeError(
            "copy_tk_runtime.py only supports the Windows and macOS packaging flows."
        )

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
