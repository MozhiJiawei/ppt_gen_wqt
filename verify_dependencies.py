#!/usr/bin/env python3
from __future__ import annotations

import argparse


def pass_check(name: str, detail: str = "") -> None:
    print(f"PASS {name}{': ' + detail if detail else ''}")


def fail_check(name: str, detail: str) -> None:
    print(f"FAIL {name}: {detail}")


def import_module(module_name: str, package_name: str) -> bool:
    try:
        module = __import__(module_name)
    except Exception as exc:
        fail_check(package_name, str(exc))
        return False

    version = getattr(module, "__version__", "")
    pass_check(package_name, version or "import ok")
    return True


def main() -> int:
    argparse.ArgumentParser(description="Verify dependencies for tech-report-ppt-safe-layout.").parse_args()

    checks = [
        import_module("pptx", "python-pptx"),
        import_module("matplotlib", "matplotlib"),
        import_module("numpy", "numpy"),
        import_module("PIL", "pillow"),
    ]
    return 0 if all(checks) else 1


if __name__ == "__main__":
    raise SystemExit(main())
