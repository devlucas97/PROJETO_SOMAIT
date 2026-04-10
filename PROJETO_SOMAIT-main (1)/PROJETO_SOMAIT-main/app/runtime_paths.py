from pathlib import Path
import sys


def get_runtime_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


def get_bundle_root() -> Path:
    if getattr(sys, "frozen", False):
        meipass = getattr(sys, "_MEIPASS", None)
        if meipass:
            return Path(meipass)
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


def get_runtime_path(*parts: str) -> str:
    return str(get_runtime_root().joinpath(*parts))


def get_bundle_path(*parts: str) -> str:
    return str(get_bundle_root().joinpath(*parts))


def ensure_runtime_dir(*parts: str) -> str:
    path = get_runtime_root().joinpath(*parts)
    path.mkdir(parents=True, exist_ok=True)
    return str(path)


def iter_config_paths():
    runtime_config = get_runtime_root() / "config.json"
    yield str(runtime_config)

    bundle_config = get_bundle_root() / "config.json"
    if bundle_config != runtime_config:
        yield str(bundle_config)