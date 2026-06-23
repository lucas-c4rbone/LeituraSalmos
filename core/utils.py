from __future__ import annotations

import re
from pathlib import Path


INVALID_FILENAME_CHARS_RE = re.compile(r'[<>:"/\\|?*]')
MULTISPACE_RE = re.compile(r"\s+")
WINDOWS_RESERVED_FILENAMES = {
    "CON",
    "PRN",
    "AUX",
    "NUL",
    "COM1",
    "COM2",
    "COM3",
    "COM4",
    "COM5",
    "COM6",
    "COM7",
    "COM8",
    "COM9",
    "LPT1",
    "LPT2",
    "LPT3",
    "LPT4",
    "LPT5",
    "LPT6",
    "LPT7",
    "LPT8",
    "LPT9",
}


class GeradorSalmosError(Exception):
    """Represent an expected, user-facing domain error.

    This exception is raised for validation and workflow issues that should be
    shown to the user as friendly messages instead of internal tracebacks.
    """


def limpar_nome_arquivo(nome: str) -> str:
    """Sanitize a filename so it is valid on Windows.

    Args:
        nome: Raw filename text.

    Returns:
        str: Sanitized filename with illegal characters removed.

    Raises:
        GeradorSalmosError: If the resulting name becomes empty.
    """
    nome_limpo = INVALID_FILENAME_CHARS_RE.sub("", nome).strip()
    nome_limpo = MULTISPACE_RE.sub(" ", nome_limpo)
    nome_limpo = nome_limpo.rstrip(" .")
    if not nome_limpo:
        raise GeradorSalmosError("Não foi possível montar o nome do arquivo final.")
    stem = Path(nome_limpo).stem.casefold().upper()
    if stem in WINDOWS_RESERVED_FILENAMES:
        raise GeradorSalmosError(
            "O nome do arquivo gerado é inválido no Windows. "
            "Escolha um texto com título/livro diferentes."
        )
    return nome_limpo
