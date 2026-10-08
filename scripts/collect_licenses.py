from __future__ import annotations

import importlib.metadata as metadata
from pathlib import Path
import re
import shutil
import sys

PACKAGES = [
    "mael",
    "openpyxl",
    "et-xmlfile",
    "PyYAML",
    "pyinstaller",
    "altgraph",
    "packaging",
    "pefile",
    "pyinstaller-hooks-contrib",
    "pywin32-ctypes",
    "setuptools",
]


def safe_name(value: str) -> str:
    return re.sub(r"[^A-Za-z0-9_.-]+", "_", value)


def main() -> None:
    output = Path(sys.argv[1]).resolve()
    licenses_root = output / "ThirdPartyLicenses"
    licenses_root.mkdir(parents=True, exist_ok=True)
    notices = [
        "Third-party software notices",
        "",
        "This application bundles or is built using the packages listed below.",
        "Retain these notices when redistributing the application.",
        "",
    ]

    for package in PACKAGES:
        try:
            distribution = metadata.distribution(package)
        except metadata.PackageNotFoundError:
            notices.append(f"- {package}: not installed in the build environment")
            continue

        name = distribution.metadata.get("Name", package)
        notices.append(
            f"- {name} {distribution.version}: "
            f"{distribution.metadata.get('License-Expression') or distribution.metadata.get('License') or 'See included license files'}"
        )
        destination = licenses_root / safe_name(name)
        copied = 0
        for relative_path in distribution.files or []:
            file_name = Path(str(relative_path)).name
            if not re.search(r"(?i)(license|licence|copying|notice|copyright)", file_name):
                continue
            source = Path(distribution.locate_file(relative_path))
            target = destination / file_name
            if source.is_file():
                destination.mkdir(parents=True, exist_ok=True)
                if not target.exists():
                    shutil.copy2(source, target)
                copied += 1
        if copied == 0:
            notices.append(f"  License text not found in package metadata; review {distribution.metadata.get('Home-page', 'project metadata')}.")

    python_root = Path(sys.base_prefix)
    for license_file in (python_root / "LICENSE.txt", python_root / "LICENSE"):
        if license_file.is_file():
            shutil.copy2(license_file, licenses_root / f"Python-{license_file.name}")
            notices.append("- Python runtime license text: ThirdPartyLicenses/Python-LICENSE.txt")
            break

    (output / "THIRD_PARTY_NOTICES.txt").write_text("\n".join(notices) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
