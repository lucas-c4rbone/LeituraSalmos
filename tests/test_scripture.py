from __future__ import annotations

import pytest

from core.scripture import LeituraBiblica, ler_texto_entrada
from core.utils import GeradorSalmosError


def test_leitura_biblica_properties() -> None:
    """Verify computed reading properties return expected values."""
    leitura = LeituraBiblica(livro="Salmos", capitulo="23", versiculos=["1 - O Senhor..."])
    assert leitura.livro_capitulo == "Salmos 23"
    assert leitura.nome_arquivo == "Salmos 23.pptx"


def test_ler_texto_entrada_valid_basic_text() -> None:
    """Verify parsing succeeds for a valid scripture block."""
    texto = """\
Livro: Salmos
Capítulo: 23
1 O Senhor é o meu pastor
2 Nada me faltará
"""
    leitura = ler_texto_entrada(texto)

    assert leitura.livro == "Salmos"
    assert leitura.capitulo == "23"
    assert leitura.versiculos == [
        "1 - O Senhor é o meu pastor",
        "2 - Nada me faltará",
    ]


def test_ler_texto_entrada_accepts_capitulo_without_accent() -> None:
    """Verify parser accepts both accented and unaccented chapter keys."""
    texto = """\
Livro: João
Capitulo: 3
1 Porque Deus amou o mundo
"""
    leitura = ler_texto_entrada(texto)
    assert leitura.capitulo == "3"


def test_ler_texto_entrada_invalid_verse_line_raises_with_line_number() -> None:
    """Verify invalid verse formatting reports the original line number."""
    texto = """\
Livro: Salmos
Capítulo: 23
linha inválida
"""
    with pytest.raises(GeradorSalmosError, match="linha 3"):
        ler_texto_entrada(texto)


def test_ler_texto_entrada_missing_livro_raises() -> None:
    """Verify parser raises when the book field is missing."""
    texto = """\
Capítulo: 23
1 O Senhor é o meu pastor
"""
    with pytest.raises(GeradorSalmosError, match="Campo 'Livro:' não encontrado"):
        ler_texto_entrada(texto)


def test_ler_texto_entrada_missing_capitulo_raises() -> None:
    """Verify parser raises when the chapter field is missing."""
    texto = """\
Livro: Salmos
1 O Senhor é o meu pastor
"""
    with pytest.raises(GeradorSalmosError, match="Campo 'Capítulo:' não encontrado"):
        ler_texto_entrada(texto)


def test_ler_texto_entrada_missing_verses_raises() -> None:
    """Verify parser raises when no numbered verses are provided."""
    texto = """\
Livro: Salmos
Capítulo: 23
"""
    with pytest.raises(GeradorSalmosError, match="Nenhum versículo numerado encontrado"):
        ler_texto_entrada(texto)
