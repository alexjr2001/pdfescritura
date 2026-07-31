import json
import os
import shutil
import tempfile
import urllib.error
import urllib.request
from dataclasses import dataclass
from pathlib import Path

from version import __version__


GITHUB_OWNER = "TU_USUARIO"
GITHUB_REPO = "TU_REPOSITORIO"
GITHUB_ASSET_NAME = "GeneradorTestimonios.exe"


@dataclass(frozen=True)
class ReleaseInfo:
    version: str
    asset_url: str


def _version_tuple(version: str):
    version = version.strip().lstrip("vV")
    parts = []
    for part in version.split("."):
        try:
            parts.append(int(part))
        except ValueError:
            parts.append(0)
    return tuple(parts)


def _request_json(url: str):
    request = urllib.request.Request(
        url,
        headers={
            "Accept": "application/vnd.github+json",
            "User-Agent": "pdfEscritura-updater",
        },
    )
    with urllib.request.urlopen(request, timeout=15) as response:
        return json.load(response)


def get_latest_release() -> ReleaseInfo | None:
    if GITHUB_OWNER == "TU_USUARIO" or GITHUB_REPO == "TU_REPOSITORIO":
        return None

    url = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/releases/latest"
    try:
        data = _request_json(url)
    except (urllib.error.URLError, TimeoutError, json.JSONDecodeError):
        return None

    version = str(data.get("tag_name", "")).strip()
    assets = data.get("assets", [])

    for asset in assets:
        if asset.get("name") == GITHUB_ASSET_NAME:
            return ReleaseInfo(version=version, asset_url=asset.get("browser_download_url", ""))

    return None


def has_new_version(local_version: str = __version__) -> bool:
    latest = get_latest_release()
    if latest is None:
        return False

    return _version_tuple(latest.version) > _version_tuple(local_version)


def download_latest_app(destination: Path) -> bool:
    latest = get_latest_release()
    if latest is None or not latest.asset_url:
        return False

    destination.parent.mkdir(parents=True, exist_ok=True)
    temp_destination = destination.with_suffix(destination.suffix + ".download")

    try:
        with urllib.request.urlopen(latest.asset_url, timeout=60) as response, open(temp_destination, "wb") as output:
            shutil.copyfileobj(response, output)
        os.replace(temp_destination, destination)
        return True
    except (urllib.error.URLError, TimeoutError, OSError):
        if temp_destination.exists():
            temp_destination.unlink(missing_ok=True)
        return False
