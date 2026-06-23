from __future__ import annotations

import json
import os
import sys
from functools import lru_cache
from pathlib import Path
from typing import Any, TYPE_CHECKING

from core.utils import GeradorSalmosError

if TYPE_CHECKING:
    from core.scripture import LeituraBiblica

from ui.theme import (
    MODEL_NAME,
    OUTPUT_DIR_DEFAULT,
    PLACEHOLDER_LIVRO_CAP,
    PLACEHOLDER_TITULO,
    PLACEHOLDER_VERSICULO,
    PLACEHOLDERS_LIVRO_CAP,
    PLACEHOLDERS_TITULO,
    PLACEHOLDERS_VERSICULO,
    PP_ALIGN_JUSTIFY,
    PP_SAVE_AS_OPENXML,
    TEMPLATE_FINAL_INDEX,
    TEMPLATE_IGREJA_INDEX,
    TEMPLATE_PASTOR_INDEX,
    TEMPLATE_TITLE_INDEX,
)

APP_DIR = Path(__file__).resolve().parent.parent


@lru_cache(maxsize=1)
def app_dir() -> Path:
    """Return the runtime application directory.

    Returns:
        Path: Directory of the frozen executable or source root.
    """
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return APP_DIR


@lru_cache(maxsize=32)
def resource_path(name: str) -> Path:
    """Resolve a resource path from external or bundled location.

    Args:
        name: Relative resource filename.

    Returns:
        Path: Absolute path to the resolved resource.
    """
    external = app_dir() / name
    if external.exists():
        return external.resolve()
    bundle_dir = Path(getattr(sys, "_MEIPASS", app_dir()))
    return (bundle_dir / name).resolve()


USER_DATA_DIR = Path(os.getenv("LOCALAPPDATA")) / "LeituraBiblica"
USER_DATA_DIR.mkdir(parents=True, exist_ok=True)

MODEL_FILE = resource_path(MODEL_NAME)
LOG_FILE = USER_DATA_DIR / "logs.log"
CONFIG_FILE = USER_DATA_DIR / "config.json"


def output_dir(config: dict[str, Any] | None = None) -> Path:
    """Resolve and ensure the output directory exists.

    Args:
        config: Optional user configuration mapping.

    Returns:
        Path: Writable output directory path.
    """
    cfg = config or carregar_configuracao()
    raw = cfg.get("output_dir", "").strip()
    folder = Path(raw) if raw else OUTPUT_DIR_DEFAULT
    try:
        folder = validar_pasta_saida(folder)
    except (GeradorSalmosError, OSError, RuntimeError, ValueError):
        from core.logging import logger

        logger.exception("Could not create output directory %s — falling back to default", folder)
        folder = validar_pasta_saida(OUTPUT_DIR_DEFAULT)
    return folder


def validar_pasta_saida(caminho: str | Path) -> Path:
    """Validate and prepare a folder path for output.

    Args:
        caminho: Candidate folder path selected by the user or config.

    Returns:
        Path: Normalized writable directory path.

    Raises:
        GeradorSalmosError: If the path is invalid, not a directory, or not writable.
    """
    texto = str(caminho).strip()
    if not texto:
        raise GeradorSalmosError("Selecione uma pasta de destino válida.")

    pasta = Path(texto).expanduser().resolve()
    if pasta.exists() and not pasta.is_dir():
        raise GeradorSalmosError("O caminho selecionado não é uma pasta válida.")

    pasta.mkdir(parents=True, exist_ok=True)
    if not os.access(pasta, os.W_OK):
        raise GeradorSalmosError("A pasta selecionada não possui permissão de escrita.")
    return pasta


def carregar_configuracao() -> dict[str, str]:
    """Load persisted user configuration from disk.

    Returns:
        dict[str, str]: Stored settings, or an empty dictionary on failure.
    """
    try:
        if CONFIG_FILE.exists():
            data = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                return {str(k): str(v) for k, v in data.items()}
    except (OSError, json.JSONDecodeError, TypeError, ValueError) as exc:
        from core.logging import logger

        logger.exception("Could not load config from %s", CONFIG_FILE)
    return {}


def salvar_configuracao(config: dict[str, str]) -> None:
    """Persist user configuration to disk.

    Args:
        config: Configuration mapping to serialize.
    """
    try:
        CONFIG_FILE.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")
    except (OSError, TypeError, ValueError) as exc:
        from core.logging import logger

        logger.exception("Could not save config to %s", CONFIG_FILE)


def caminho_saida_para(leitura: LeituraBiblica, config: dict[str, Any] | None = None) -> Path:
    """Build the output file path for a parsed reading.

    Args:
        leitura: Parsed scripture reading information.
        config: Optional user configuration mapping.

    Returns:
        Path: Absolute path for the generated presentation file.
    """
    return output_dir(config) / leitura.nome_arquivo
