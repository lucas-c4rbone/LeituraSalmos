from __future__ import annotations

import pytest

from core.utils import GeradorSalmosError, limpar_nome_arquivo


def test_limpar_nome_arquivo_remove_invalid_chars_and_collapse_spaces() -> None:
    """Ensure invalid filename characters are removed and spaces normalized."""
    raw = '  Salmos<>:"/\\|?*   23   '
    assert limpar_nome_arquivo(raw) == "Salmos 23"


def test_limpar_nome_arquivo_empty_after_cleanup_raises() -> None:
    """Ensure an error is raised when cleanup results in an empty filename."""
    with pytest.raises(GeradorSalmosError, match="Não foi possível montar o nome do arquivo final"):
        limpar_nome_arquivo('   <>:"/\\|?*   ')


def test_limpar_nome_arquivo_windows_reserved_name_raises() -> None:
    """Ensure Windows reserved filenames are rejected."""
    with pytest.raises(GeradorSalmosError, match="inválido no Windows"):
        limpar_nome_arquivo("con")
