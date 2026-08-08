import json
import os
import shutil
import hashlib
import errno
import ssl
import urllib.error
import urllib.request
import zipfile
from dataclasses import dataclass
from pathlib import Path

try:
    import certifi
except ImportError:  # pragma: no cover - depende del entorno de empaquetado
    certifi = None

from version import __version__


GITHUB_OWNER = "alexjr2001"
GITHUB_REPO = "pdfescritura"
APP_ZIP_ASSET_NAME = "GeneradorTestimonios-win64.zip"
CHECKSUMS_ASSET_NAME = "checksums.txt"
APP_EXE_NAME = "GeneradorTestimonios.exe"

_last_release_error = ""


def _build_ssl_context() -> ssl.SSLContext:
    # En ejecutables empaquetados, usar certifi evita fallos por cadenas CA no
    # disponibles en el entorno embebido de OpenSSL.
    cafile = None
    if certifi is not None:
        try:
            cafile = certifi.where()
        except Exception:
            cafile = None

    if cafile:
        try:
            return ssl.create_default_context(cafile=cafile)
        except Exception:
            pass
    return ssl.create_default_context()


_SSL_CONTEXT = _build_ssl_context()


@dataclass(frozen=True)
class ReleaseInfo:
    version: str
    zip_asset_url: str
    checksums_url: str


@dataclass(frozen=True)
class UpdateResult:
    executable_path: Path | None
    status: str
    updated: bool


def _version_tuple(version: str):
    version = version.strip().lstrip("vV")
    parts = []
    for part in version.split("."):
        try:
            parts.append(int(part))
        except ValueError:
            parts.append(0)
    return tuple(parts)


def _repo_slug(repo: str) -> str:
    repo = repo.strip().rstrip("/")
    if repo.startswith("http://") or repo.startswith("https://"):
        last = repo.split("/")[-1]
        return last
    return repo


def _install_root() -> Path:
    local_app_data = os.environ.get("LOCALAPPDATA")
    if local_app_data:
        return Path(local_app_data) / "GeneradorTestimonios"
    return Path.home() / ".generador_testimonios"


def _version_file() -> Path:
    return _install_root() / "current_version.txt"


def _executable_for_version(version: str) -> Path:
    clean_version = version.strip().lstrip("vV") or "unknown"
    return _install_root() / f"app-{clean_version}" / APP_EXE_NAME


def _read_local_version() -> str:
    version_file = _version_file()
    if not version_file.exists():
        return ""
    return version_file.read_text(encoding="utf-8").strip()


def _request_json(url: str):
    request = urllib.request.Request(
        url,
        headers={
            "Accept": "application/vnd.github+json",
            "User-Agent": "pdfEscritura-updater",
        },
    )
    with urllib.request.urlopen(request, timeout=15, context=_SSL_CONTEXT) as response:
        return json.load(response)


def _format_release_error(exc: Exception) -> str:
    if isinstance(exc, urllib.error.HTTPError):
        return f"HTTP {exc.code} {exc.reason}".strip()
    if isinstance(exc, urllib.error.URLError):
        reason = getattr(exc, "reason", None)
        return f"URLError: {reason}" if reason else "URLError"
    if isinstance(exc, TimeoutError):
        return "Timeout al consultar GitHub Releases"
    if isinstance(exc, json.JSONDecodeError):
        return "Respuesta inválida de GitHub Releases (JSON)"
    return f"{type(exc).__name__}: {exc}".strip()


def get_latest_release() -> ReleaseInfo | None:
    global _last_release_error
    if GITHUB_OWNER == "TU_USUARIO" or GITHUB_REPO == "TU_REPOSITORIO":
        _last_release_error = "Repositorio sin configurar"
        return None

    repo_slug = _repo_slug(GITHUB_REPO)
    url = f"https://api.github.com/repos/{GITHUB_OWNER}/{repo_slug}/releases/latest"
    _last_release_error = ""
    try:
        data = _request_json(url)
    except (urllib.error.URLError, TimeoutError, json.JSONDecodeError, urllib.error.HTTPError) as exc:
        _last_release_error = _format_release_error(exc)
        return None

    version = str(data.get("tag_name", "")).strip()
    assets = data.get("assets", [])
    zip_asset_url = ""
    checksums_url = ""

    for asset in assets:
        name = asset.get("name")
        if name == APP_ZIP_ASSET_NAME:
            zip_asset_url = asset.get("browser_download_url", "")
        elif name == CHECKSUMS_ASSET_NAME:
            checksums_url = asset.get("browser_download_url", "")

    if zip_asset_url:
        return ReleaseInfo(
            version=version,
            zip_asset_url=zip_asset_url,
            checksums_url=checksums_url,
        )

    _last_release_error = f"No se encontró asset {APP_ZIP_ASSET_NAME} en el último release"
    return None


def has_new_version(local_version: str = "") -> bool:
    local_version = local_version or _read_local_version() or __version__
    latest = get_latest_release()
    if latest is None:
        return False

    return _version_tuple(latest.version) > _version_tuple(local_version)


def _download_to_file(url: str, destination: Path) -> bool:
    destination.parent.mkdir(parents=True, exist_ok=True)
    temp_destination = destination.with_suffix(destination.suffix + ".download")
    try:
        with urllib.request.urlopen(url, timeout=60, context=_SSL_CONTEXT) as response, open(temp_destination, "wb") as output:
            shutil.copyfileobj(response, output)
        os.replace(temp_destination, destination)
        return True
    except (urllib.error.URLError, TimeoutError, OSError):
        if temp_destination.exists():
            temp_destination.unlink(missing_ok=True)
        return False


def _parse_checksums(path: Path) -> dict[str, str]:
    values: dict[str, str] = {}
    for raw_line in path.read_text(encoding="utf-8-sig").splitlines():
        line = raw_line.strip()
        if not line:
            continue
        if "  " in line:
            digest, filename = line.split("  ", 1)
        else:
            parts = line.split()
            if len(parts) < 2:
                continue
            digest, filename = parts[0], parts[-1]

        digest = digest.strip().lower()
        filename = filename.strip().lstrip("*").replace("\\", "/")
        basename = Path(filename).name

        values[filename] = digest
        values[basename] = digest
    return values


def _sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest().lower()


def _is_access_denied_error(exc: OSError) -> bool:
    winerror = getattr(exc, "winerror", None)
    return winerror == 5 or exc.errno in {errno.EACCES, errno.EPERM}


def _install_release(latest: ReleaseInfo) -> UpdateResult:
    install_root = _install_root()
    install_root.mkdir(parents=True, exist_ok=True)

    zip_path = install_root / APP_ZIP_ASSET_NAME
    checksums_path = install_root / CHECKSUMS_ASSET_NAME

    if not _download_to_file(latest.zip_asset_url, zip_path):
        return UpdateResult(None, "No se pudo descargar el paquete de actualización.", False)

    if latest.checksums_url:
        if _download_to_file(latest.checksums_url, checksums_path):
            checksums = _parse_checksums(checksums_path)
            expected = checksums.get(APP_ZIP_ASSET_NAME)
            if not expected:
                return UpdateResult(
                    None,
                    f"El archivo {CHECKSUMS_ASSET_NAME} no contiene una entrada para {APP_ZIP_ASSET_NAME}.",
                    False,
                )

            current = _sha256_file(zip_path)
            if expected != current:
                return UpdateResult(
                    None,
                    (
                        "La verificación del paquete falló (SHA256 inválido). "
                        f"Esperado: {expected[:12]}..., Actual: {current[:12]}..."
                    ),
                    False,
                )

    target_dir = _executable_for_version(latest.version).parent
    temp_extract_dir = target_dir.with_name(target_dir.name + "-tmp")

    if temp_extract_dir.exists():
        shutil.rmtree(temp_extract_dir, ignore_errors=True)

    temp_extract_dir.mkdir(parents=True, exist_ok=True)

    try:
        with zipfile.ZipFile(zip_path, "r") as archive:
            archive.extractall(temp_extract_dir)
    except zipfile.BadZipFile:
        shutil.rmtree(temp_extract_dir, ignore_errors=True)
        return UpdateResult(None, "El archivo ZIP de actualización está dañado.", False)

    extracted_exe = temp_extract_dir / APP_EXE_NAME
    if not extracted_exe.exists():
        nested_match = list(temp_extract_dir.rglob(APP_EXE_NAME))
        if not nested_match:
            shutil.rmtree(temp_extract_dir, ignore_errors=True)
            return UpdateResult(None, "El paquete no contiene el ejecutable esperado.", False)
        extracted_exe = nested_match[0]

    try:
        if target_dir.exists():
            shutil.rmtree(target_dir)
        target_dir.parent.mkdir(parents=True, exist_ok=True)
        temp_extract_dir.replace(target_dir)
    except OSError as exc:
        shutil.rmtree(temp_extract_dir, ignore_errors=True)
        if _is_access_denied_error(exc):
            return UpdateResult(
                None,
                "No se pudo actualizar porque la aplicación está abierta o bloqueada. "
                "Ciérrala y vuelve a intentar.",
                False,
            )
        return UpdateResult(None, f"No se pudo instalar la actualización: {exc}", False)

    _version_file().write_text(latest.version.strip(), encoding="utf-8")
    final_exe = target_dir / extracted_exe.relative_to(extracted_exe.parents[0])
    # Si la estructura fue anidada, usa búsqueda directa para garantizar ruta válida.
    if not final_exe.exists():
        found = list(target_dir.rglob(APP_EXE_NAME))
        if found:
            final_exe = found[0]

    return UpdateResult(final_exe if final_exe.exists() else None, "Actualización instalada.", True)


def _get_local_executable() -> Path | None:
    version = _read_local_version()
    if version:
        exe = _executable_for_version(version)
        if exe.exists():
            return exe

    install_root = _install_root()
    candidates = sorted(install_root.rglob(APP_EXE_NAME), key=lambda p: p.stat().st_mtime, reverse=True)
    return candidates[0] if candidates else None


def ensure_app_up_to_date() -> UpdateResult:
    local_exe = _get_local_executable()
    local_version = _read_local_version() or __version__
    latest = get_latest_release()

    if latest is None:
        detalle = f" Detalle: {_last_release_error}." if _last_release_error else ""
        if local_exe and local_exe.exists():
            return UpdateResult(
                local_exe,
                "Sin conexión o sin Release disponible. Se usa la versión local." + detalle,
                False,
            )
        return UpdateResult(None, "No hay versión local y no se pudo consultar GitHub Releases." + detalle, False)

    needs_update = (not local_exe or not local_exe.exists()) or (
        _version_tuple(latest.version) > _version_tuple(local_version)
    )

    if needs_update:
        updated = _install_release(latest)
        if updated.executable_path and updated.executable_path.exists():
            return updated
        if local_exe and local_exe.exists():
            return UpdateResult(local_exe, updated.status, False)
        return updated

    return UpdateResult(local_exe, "Ya tienes la versión más reciente.", False)
