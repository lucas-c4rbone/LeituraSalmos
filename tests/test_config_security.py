from __future__ import annotations

from pathlib import Path

import pytest

from core.config import validar_pasta_saida
from core.utils import GeradorSalmosError


def test_validar_pasta_saida_rejects_file_path(tmp_path: Path) -> None:
    """Ensure folder validation fails when pointing to a file."""
    arquivo = tmp_path / "nao_e_pasta.txt"
    arquivo.write_text("x", encoding="utf-8")

    with pytest.raises(GeradorSalmosError, match="não é uma pasta válida"):
        validar_pasta_saida(arquivo)


def test_validar_pasta_saida_creates_missing_folder(tmp_path: Path) -> None:
    """Ensure folder validation creates non-existing directories."""
    destino = tmp_path / "nova" / "pasta"

    validada = validar_pasta_saida(destino)

    assert validada.exists()
    assert validada.is_dir()
