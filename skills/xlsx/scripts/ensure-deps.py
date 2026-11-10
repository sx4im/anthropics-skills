"""
Check the xlsx skill's Python dependencies and, outside the sandbox,
install the missing ones into a virtual environment.

SKILL.md says openpyxl, pandas and markitdown are preinstalled, but that
only holds inside the sandbox. On a stock machine (e.g. macOS system
Python) an import can fail, and `pip install` into the system interpreter
is unreliable there (externally managed environment). This script checks
the three imports and, if anything is missing, creates a venv and installs
only the missing packages into it -- never into the system Python.

Usage (paths are relative to this skill's directory):
    python scripts/ensure-deps.py [--check] [--venv PATH] [--packages ...]

    --check     report only; exit 1 if anything is missing, install nothing
    --venv      where to create the venv (default: <skill dir>/.venv)
    --packages  which packages to check/install
                (default: openpyxl pandas markitdown)
"""

import argparse
import shutil
import subprocess
import sys
from pathlib import Path

SKILL_DIR = Path(__file__).resolve().parent.parent
DEFAULT_PACKAGES = ("openpyxl", "pandas", "markitdown")


def check_imports(python, packages):
    """Return the packages that cannot be imported by the given interpreter."""
    missing = []
    for package in packages:
        result = subprocess.run(
            [str(python), "-c", f"import {package}"],
            capture_output=True,
        )
        if result.returncode != 0:
            missing.append(package)
    return missing


def check_markitdown_command(python):
    """True if the `markitdown` command is usable from this interpreter."""
    if shutil.which("markitdown"):
        return True
    result = subprocess.run(
        [str(python), "-m", "markitdown", "--help"],
        capture_output=True,
    )
    return result.returncode == 0


def venv_python(venv_dir):
    return venv_dir / ("Scripts" if sys.platform == "win32" else "bin") / "python"


def ensure_venv(venv_dir):
    """Create the venv if needed; return its interpreter, or None on failure."""
    python = venv_python(venv_dir)
    if python.exists():
        return python
    result = subprocess.run(
        [sys.executable, "-m", "venv", str(venv_dir)],
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        print(result.stderr.strip(), file=sys.stderr)
        return None
    return python


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--check", action="store_true",
                        help="report only; install nothing")
    parser.add_argument("--venv", type=Path, default=SKILL_DIR / ".venv",
                        help="where to create the venv (default: %(default)s)")
    parser.add_argument("--packages", nargs="+", default=list(DEFAULT_PACKAGES),
                        help="packages to check/install (default: %(default)s)")
    args = parser.parse_args(argv)

    missing = check_imports(sys.executable, args.packages)
    if "markitdown" in args.packages and "markitdown" not in missing:
        if not check_markitdown_command(sys.executable):
            missing.append("markitdown")

    if not missing:
        print("All dependencies present:", ", ".join(args.packages))
        return 0

    if args.check:
        print("Missing:", ", ".join(missing))
        return 1

    # The venv must be self-contained: a script run with the venv's
    # interpreter needs every package in it, not just the ones that were
    # missing from the system interpreter.
    if args.venv.exists():
        python = venv_python(args.venv)
        to_install = check_imports(python, args.packages)
        if "markitdown" in args.packages and "markitdown" not in to_install:
            if not check_markitdown_command(python):
                to_install.append("markitdown")
        if not to_install:
            print("All dependencies present in the virtual environment:",
                  ", ".join(args.packages))
            return 0
    else:
        print(f"Creating virtual environment at {args.venv} ...")
        python = ensure_venv(args.venv)
        if python is None:
            print("Could not create a virtual environment. Set one up by hand:",
                  file=sys.stderr)
            print(f"  python -m venv {args.venv}", file=sys.stderr)
            print(f"  {venv_python(args.venv)} -m pip install "
                  + " ".join(args.packages), file=sys.stderr)
            return 1
        to_install = list(args.packages)

    result = subprocess.run(
        [str(python), "-m", "pip", "install", "--disable-pip-version-check",
         *to_install],
    )
    if result.returncode != 0:
        return 1

    still_missing = check_imports(python, args.packages)
    if "markitdown" in args.packages and "markitdown" not in still_missing:
        if not check_markitdown_command(python):
            still_missing.append("markitdown")
    if still_missing:
        print("Still missing after install:", ", ".join(still_missing),
              file=sys.stderr)
        return 1

    print("Installed into the virtual environment. Run your scripts with:")
    print(f"  {python} scripts/recalc.py ...")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
