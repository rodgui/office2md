#!/usr/bin/env python3
"""
Bump version script for office2md.

Usage:
    python scripts/bump_version.py patch   # 0.1.0 -> 0.1.1
    python scripts/bump_version.py minor   # 0.1.0 -> 0.2.0
    python scripts/bump_version.py major   # 0.1.0 -> 1.0.0
    python scripts/bump_version.py 0.2.0   # Set specific version
"""

import re
import sys
from pathlib import Path


VERSION_FILE = Path(__file__).parent.parent / "office2md" / "__version__.py"


def get_current_version() -> str:
    """Read current version from __version__.py."""
    content = VERSION_FILE.read_text()
    match = re.search(r'__version__\s*=\s*["\']([^"\']+)["\']', content)
    if not match:
        raise ValueError("Could not find version in __version__.py")
    return match.group(1)


def parse_version(version: str) -> tuple:
    """Parse version string to tuple of ints."""
    parts = version.split(".")
    if len(parts) != 3:
        raise ValueError(f"Invalid version format: {version}")
    return tuple(int(p) for p in parts)


def bump_version(current: str, bump_type: str) -> str:
    """Calculate new version based on bump type."""
    major, minor, patch = parse_version(current)
    
    if bump_type == "major":
        return f"{major + 1}.0.0"
    elif bump_type == "minor":
        return f"{major}.{minor + 1}.0"
    elif bump_type == "patch":
        return f"{major}.{minor}.{patch + 1}"
    else:
        # Assume it's a specific version
        parse_version(bump_type)  # Validate format
        return bump_type


def update_version_file(new_version: str):
    """Update __version__.py with new version."""
    content = VERSION_FILE.read_text()
    
    # Update __version__
    content = re.sub(
        r'__version__\s*=\s*["\'][^"\']+["\']',
        f'__version__ = "{new_version}"',
        content
    )
    
    VERSION_FILE.write_text(content)


def main():
    if len(sys.argv) != 2:
        print(__doc__)
        sys.exit(1)
    
    bump_type = sys.argv[1]
    
    current = get_current_version()
    new = bump_version(current, bump_type)
    
    print(f"Current version: {current}")
    print(f"New version:     {new}")
    
    confirm = input("Proceed? [y/N] ").strip().lower()
    if confirm != "y":
        print("Aborted.")
        sys.exit(1)
    
    update_version_file(new)
    print(f"✅ Updated {VERSION_FILE}")
    
    print("\nNext steps:")
    print(f"  1. Update CHANGELOG.md with changes for v{new}")
    print(f"  2. git add -A")
    print(f"  3. git commit -m 'Bump version to {new}'")
    print(f"  4. git tag -a v{new} -m 'Release {new}'")
    print(f"  5. git push origin main --tags")


if __name__ == "__main__":
    main()