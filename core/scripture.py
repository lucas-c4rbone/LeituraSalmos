from __future__ import annotations

import re
from dataclasses import dataclass

from core.utils import GeradorSalmosError, limpar_nome_arquivo


VERSICULO_LINHA_RE = re.compile(r"^(\d+)\s+(.+)$")


@dataclass(frozen=True)
class LeituraBiblica:
    """Store parsed scripture content used to generate slides.

    Attributes:
        livro: Bible book name.
        capitulo: Chapter number as text.
        versiculos: Formatted verse lines.
    """

    livro: str
    capitulo: str
    versiculos: list[str]

    @property
    def livro_capitulo(self) -> str:
        """Return the canonical display label for book and chapter.

        Returns:
            str: String in the format "<livro> <capitulo>".
        """
        return f"{self.livro} {self.capitulo}"

    @property
    def nome_arquivo(self) -> str:
        """Build the output filename for this scripture reading.

        Returns:
            str: Sanitized PowerPoint filename ending with .pptx.
        """
        return f"{limpar_nome_arquivo(self.livro_capitulo)}.pptx"


def ler_texto_entrada(texto: str) -> LeituraBiblica:
    """Parse raw clipboard text into a structured reading model.

    Args:
        texto: Source text containing book, chapter, and numbered verses.

    Returns:
        LeituraBiblica: Parsed reading ready for slide generation.

    Raises:
        GeradorSalmosError: If required fields are missing or a line format is invalid.
    """
    livro = ""
    capitulo = ""
    versiculos: list[str] = []

    for numero_linha, linha_original in enumerate(texto.splitlines(), start=1):
        linha = linha_original.strip()
        if not linha:
            continue

        chave, separador, valor = linha.partition(":")
        if separador and chave.strip().casefold() == "livro":
            livro = valor.strip()
            continue
        if separador and chave.strip().casefold() in {"capitulo", "capítulo"}:
            capitulo = valor.strip()
            continue

        match = VERSICULO_LINHA_RE.match(linha)
        if match:
            numero, texto_versiculo = match.groups()
            versiculos.append(f"{numero} - {texto_versiculo.strip()}")
            continue

        raise GeradorSalmosError(
            f"Formato inválido na linha {numero_linha}:\n\n"
            f"    {linha_original}\n\n"
            "Cada versículo deve começar com seu número seguido de espaço.\n"
            "Exemplo:  1 O Senhor é o meu pastor…"
        )

    if not livro:
        raise GeradorSalmosError(
            "Campo 'Livro:' não encontrado.\n\n"
            "Certifique-se de que o texto começa com:\n    Livro: Salmos"
        )
    if not capitulo:
        raise GeradorSalmosError(
            "Campo 'Capítulo:' não encontrado.\n\n"
            "Certifique-se de incluir:\n    Capítulo: 23"
        )
    if not versiculos:
        raise GeradorSalmosError(
            "Nenhum versículo numerado encontrado.\n\n"
            "Cada versículo deve começar com seu número:\n    1 O Senhor é o meu pastor…"
        )

    return LeituraBiblica(livro=livro, capitulo=capitulo, versiculos=versiculos)
