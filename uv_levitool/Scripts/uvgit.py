"""Install or upgrade a package from a git repository using uv.

Attempts to upgrade the package first. If upgrade fails (e.g. the package is
not yet installed), falls back to a fresh add. All errors are logged to the
console via loguru.

Usage:
    python uvgit.py <git_repo_url> [branch]
    main(["https://github.com/owner/pkg.git", "branch"])

Args:
    git_repo_url (str): Full git URL, e.g. https://github.com/owner/pkg.git (required).
    branch (str): Branch, tag, or commit to pin, e.g. main (optional).

Returns:
    int: Exit code.
        0 = success (upgrade succeeded).
        1 = upgrade failed, add succeeded.
        2 = both upgrade and add failed.
        3 = bad arguments / precondition failure.
"""

import re
import shutil
import subprocess
import sys
from pathlib import Path
from typing import List, Optional, Tuple

from loguru import logger

try:
    from uvbekutils import setup_loguru
    _HAS_SETUP_LOGURU = True
except ImportError:
    _HAS_SETUP_LOGURU = False

SUBPROCESS_TIMEOUT = 120  # seconds


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _parse_package_name(repo_url: str) -> Optional[str]:
    """Extract the package name from the last segment of a git URL.

    Handles both .git-suffixed and bare paths, e.g.:
        https://github.com/owner/uvbekutils.git  ->  uvbekutils
        https://github.com/owner/uvbekutils      ->  uvbekutils

    Args:
        repo_url (str): Full git repository URL.

    Returns:
        str: Package name, or None if it cannot be derived.

    Examples:
        >>> _parse_package_name("https://github.com/owner/uvbekutils.git")
        'uvbekutils'
        >>> _parse_package_name("https://github.com/owner/uvbekutils")
        'uvbekutils'
        >>> _parse_package_name("git@github.com:owner/mypackage.git")
        'mypackage'
        >>> _parse_package_name("https://github.com/owner/") is None
        True
    """
    match = re.search(r"/([^/]+?)(?:\.git)?$", repo_url.rstrip("/"))
    return match.group(1) if match else None


def _find_uv() -> Optional[str]:
    """Locate the uv executable, checking PATH then common install locations.

    PyCharm and other IDEs may not inherit the full shell PATH, so uv may be
    present on the system but invisible to shutil.which. This function falls
    back to well-known installation directories used by uv's own installer,
    Homebrew, and Cargo.

    Returns:
        str: Full path to the uv executable, or None if not found anywhere.
    """
    found = shutil.which("uv")
    if found:
        return found

    common = [
        Path.home() / ".local" / "bin" / "uv",       # uv installer default
        Path.home() / ".cargo" / "bin" / "uv",        # cargo install
        Path("/opt/homebrew/bin/uv"),                  # Homebrew on Apple Silicon
        Path("/usr/local/bin/uv"),                     # Homebrew on Intel / manual
    ]
    for path in common:
        if path.exists():
            return str(path)
    return None


def _build_add_spec(repo_url: str, branch: Optional[str]) -> str:
    """Build the package specification string for `uv add`.

    Args:
        repo_url (str): Full git repository URL.
        branch (str): Branch, tag, or commit ref, or None to use the default branch.

    Returns:
        str: Spec in the form ``pkg@git+url`` or ``pkg@git+url@branch``.

    Examples:
        >>> _build_add_spec("https://github.com/owner/uvbekutils.git", None)
        'uvbekutils@git+https://github.com/owner/uvbekutils.git'
        >>> _build_add_spec("https://github.com/owner/uvbekutils.git", "main")
        'uvbekutils@git+https://github.com/owner/uvbekutils.git@main'
        >>> _build_add_spec("https://github.com/owner/uvbekutils.git", "optiona_showbuttons")
        'uvbekutils@git+https://github.com/owner/uvbekutils.git@optiona_showbuttons'
    """
    pkg = _parse_package_name(repo_url)
    spec = f"{pkg}@git+{repo_url}"
    if branch:
        spec += f"@{branch}"
    return spec


# ---------------------------------------------------------------------------
# Pre-flight checks
# ---------------------------------------------------------------------------

def preflight(repo_url: str) -> Tuple[bool, str]:
    """Validate inputs and environment before running any uv commands.

    Checks performed:
        1. URL is non-empty.
        2. URL starts with a recognised git scheme (https, http, git@, git://).
        3. Package name can be derived from the URL path.
        4. ``uv`` is present on PATH.

    Args:
        repo_url (str): Git repository URL to validate.

    Returns:
        Tuple[bool, str]: ``(True, "")`` if all checks pass, or
            ``(False, error_message)`` on the first failing check.
    """
    if not repo_url or not repo_url.strip():
        return False, "git_repo_url must not be empty"

    if not repo_url.startswith(("https://", "http://", "git@", "git://")):
        return False, f"git_repo_url does not look like a valid git URL: {repo_url!r}"

    pkg = _parse_package_name(repo_url)
    if not pkg:
        return False, f"Cannot derive package name from URL: {repo_url!r}"

    if not _find_uv():
        return False, "'uv' was not found on PATH or common install locations — install uv first"

    return True, ""


# ---------------------------------------------------------------------------
# Core operations
# ---------------------------------------------------------------------------

def upgrade_package(repo_url: str, pkg_name: str, branch: Optional[str]) -> bool:
    """Upgrade an already-installed package to the latest commit on a git ref.

    Runs: ``uv add pkg@git+url[@branch] --upgrade-package pkg_name``

    Args:
        repo_url (str): Full git repository URL.
        pkg_name (str): Package name as known to uv (derived from the URL).
        branch (str): Branch, tag, or commit ref, or None for the default branch.

    Returns:
        bool: True if uv exited successfully, False on error or timeout.
    """
    uv = _find_uv()
    add_spec = _build_add_spec(repo_url, branch)
    cmd = [uv, "add", add_spec, "--upgrade-package", pkg_name]
    logger.info("Upgrading: {}", " ".join(cmd))
    try:
        subprocess.run(cmd, check=True, timeout=SUBPROCESS_TIMEOUT)
        return True
    except subprocess.CalledProcessError as exc:
        logger.error("Upgrade failed (exit {})", exc.returncode)
        return False
    except subprocess.TimeoutExpired:
        logger.error("Upgrade timed out after {}s", SUBPROCESS_TIMEOUT)
        return False


def add_package(repo_url: str, branch: Optional[str]) -> bool:
    """Install a package fresh from a git URL.

    Runs: ``uv add pkg@git+url[@branch]``

    Args:
        repo_url (str): Full git repository URL.
        branch (str): Branch, tag, or commit ref, or None for the default branch.

    Returns:
        bool: True if uv exited successfully, False on error or timeout.
    """
    uv = _find_uv()
    add_spec = _build_add_spec(repo_url, branch)
    cmd = [uv, "add", add_spec]
    logger.info("Adding: {}", " ".join(cmd))
    try:
        subprocess.run(cmd, check=True, timeout=SUBPROCESS_TIMEOUT)
        return True
    except subprocess.CalledProcessError as exc:
        logger.error("Add failed (exit {})", exc.returncode)
        return False
    except subprocess.TimeoutExpired:
        logger.error("Add timed out after {}s", SUBPROCESS_TIMEOUT)
        return False


# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------

def main(argv: Optional[List[str]] = None) -> int:
    """Run the upgrade-or-add workflow and return an exit code.

    Tries upgrade first; falls back to add if upgrade fails. Configures
    loguru logging before doing anything else.

    Args:
        argv (List[str]): Argument list: [git_repo_url, branch(optional)].
            Defaults to sys.argv[1:] when None.

    Returns:
        int: Exit code (see module docstring for values).
    """
    if _HAS_SETUP_LOGURU:
        setup_loguru("DEBUG", "DEBUG")
    else:
        logger.remove()
        logger.add(sys.stderr, level="DEBUG")

    if argv is None:
        argv = sys.argv[1:]

    if not argv or argv[0] in ("-h", "--help"):
        print(__doc__)
        return 0

    repo_url = argv[0]
    branch = argv[1] if len(argv) > 1 else None

    # --- Pre-flight ---
    ok, err = preflight(repo_url)
    if not ok:
        logger.error("Precondition failed: {}", err)
        return 3

    pkg_name = _parse_package_name(repo_url)  # guaranteed non-None after preflight

    # --- Try upgrade first ---
    if upgrade_package(repo_url, pkg_name, branch):
        logger.info("Package '{}' upgraded successfully.", pkg_name)
        return 0

    logger.warning("Upgrade unsuccessful; attempting fresh add...")

    # --- Fall back to add ---
    if add_package(repo_url, branch):
        logger.info("Package '{}' added successfully.", pkg_name)
        return 1  # upgrade failed but add succeeded — caller may care

    logger.error("Both upgrade and add failed for '{}'.", pkg_name)
    return 2


# ---------------------------------------------------------------------------
# Tests — discovered and run by pytest
# ---------------------------------------------------------------------------

URL = "https://github.com/owner/uvbekutils.git"


# --- _parse_package_name ---

def test_parse_url_with_git_suffix():
    assert _parse_package_name("https://github.com/owner/uvbekutils.git") == "uvbekutils"

def test_parse_url_without_git_suffix():
    assert _parse_package_name("https://github.com/owner/uvbekutils") == "uvbekutils"

def test_parse_ssh_url():
    assert _parse_package_name("git@github.com:owner/mypackage.git") == "mypackage"

def test_parse_trailing_slash_returns_parent_segment():
    # rstrip("/") strips the slash so "owner" becomes the last segment
    assert _parse_package_name("https://github.com/owner/") == "owner"

def test_parse_bare_scheme_returns_none():
    # no usable path segment — cannot derive a name
    assert _parse_package_name("https://") is None


# --- _build_add_spec ---

def test_build_spec_no_branch():
    assert _build_add_spec(URL, None) == "uvbekutils@git+https://github.com/owner/uvbekutils.git"

def test_build_spec_with_branch():
    assert _build_add_spec(URL, "main") == "uvbekutils@git+https://github.com/owner/uvbekutils.git@main"

def test_build_spec_with_feature_branch():
    assert _build_add_spec(URL, "optiona_showbuttons") == \
        "uvbekutils@git+https://github.com/owner/uvbekutils.git@optiona_showbuttons"


# --- preflight (bad input only — no network or uv needed) ---

def test_preflight_empty_url():
    ok, msg = preflight("")
    assert not ok
    assert "empty" in msg

def test_preflight_bad_scheme():
    ok, msg = preflight("not-a-url")
    assert not ok
    assert "valid git URL" in msg

def test_preflight_empty_string_url():
    ok, msg = preflight("   ")
    assert not ok
    assert "empty" in msg


if __name__ == "__main__":
    sys.exit(main())