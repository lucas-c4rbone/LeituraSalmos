from __future__ import annotations

import re
import time
from functools import lru_cache
from pathlib import Path
from typing import Any, Callable, Iterator

try:
    import pythoncom
    import pywintypes
    import win32com.client
except ImportError:
    pythoncom = None
    pywintypes = None
    win32com = None

from core.config import (
    MODEL_FILE,
    PP_ALIGN_JUSTIFY,
    PP_SAVE_AS_OPENXML,
    PLACEHOLDER_LIVRO_CAP,
    PLACEHOLDER_TITULO,
    PLACEHOLDER_VERSICULO,
    PLACEHOLDERS_LIVRO_CAP,
    PLACEHOLDERS_TITULO,
    PLACEHOLDERS_VERSICULO,
    TEMPLATE_FINAL_INDEX,
    TEMPLATE_IGREJA_INDEX,
    TEMPLATE_PASTOR_INDEX,
    TEMPLATE_TITLE_INDEX,
    caminho_saida_para,
    validar_pasta_saida,
)
from core.logging import logger
from core.scripture import LeituraBiblica, ler_texto_entrada
from core.utils import GeradorSalmosError, limpar_nome_arquivo


RECOVERABLE_ERRORS = (AttributeError, OSError, RuntimeError, TypeError, ValueError)


def com_error_types() -> tuple[type[BaseException], ...]:
    """Return COM exception types available in the current environment.

    Returns:
        tuple[type[BaseException], ...]: Tuple used in ``except`` clauses.
    """
    if pywintypes is None:
        return ()
    return (pywintypes.com_error,)


class ResultadoGeracao:
    """Represent the outcome metadata from a generation run.

    Attributes:
        caminho: Absolute path of the generated presentation.
        quantidade_versiculos: Number of verses rendered.
        duracao_segundos: Elapsed generation time in seconds.
    """

    caminho: Path
    quantidade_versiculos: int
    duracao_segundos: float

    def __init__(self, caminho: Path, quantidade_versiculos: int, duracao_segundos: float) -> None:
        self.caminho = caminho
        self.quantidade_versiculos = quantidade_versiculos
        self.duracao_segundos = duracao_segundos


def iterar_formas(forma: Any) -> Iterator[Any]:
    """Yield a shape and all nested shapes for grouped objects.

    Args:
        forma: Root PowerPoint shape object.

    Yields:
        Any: The current shape and child shapes if grouped.
    """
    yield forma
    try:
        if forma.Type == 6:  # msoGroup
            for i in range(1, forma.GroupItems.Count + 1):
                yield from iterar_formas(forma.GroupItems(i))
    except (AttributeError, TypeError, ValueError):
        return


def normalizar_placeholder(texto: str) -> str:
    """Normalize placeholder text for robust matching.

    Args:
        texto: Placeholder text extracted from slide shapes.

    Returns:
        str: Normalized, case-insensitive placeholder representation.
    """
    texto = texto.replace("\r", "").replace("\n", " ").strip()
    texto = texto.replace("–", "-").replace("—", "-")
    texto = re.sub(r"\s+", " ", texto)
    return texto.casefold()


@lru_cache(maxsize=64)
def _placeholders_normalizados(placeholders: tuple[str, ...]) -> frozenset[str]:
    """Cache normalized placeholder sets for repeated lookups.

    Args:
        placeholders: Placeholder variants accepted for replacement.

    Returns:
        frozenset[str]: Normalized placeholder tokens.
    """
    return frozenset(normalizar_placeholder(p) for p in placeholders)


def _formas_texto_do_slide(slide: Any, lookup_cache: dict[int, list[Any]] | None = None) -> list[Any]:
    """Return text-capable shapes for a slide, using optional cache.

    Args:
        slide: PowerPoint slide COM object.
        lookup_cache: Optional cache keyed by slide object id.

    Returns:
        list[Any]: Shapes that expose a non-empty text frame.
    """
    cache_key = id(slide)
    if lookup_cache is not None and cache_key in lookup_cache:
        return lookup_cache[cache_key]

    formas: list[Any] = []
    for forma_raiz in slide.Shapes:
        for forma in iterar_formas(forma_raiz):
            try:
                if forma.HasTextFrame and forma.TextFrame.HasText:
                    formas.append(forma)
            except (AttributeError, TypeError, ValueError):
                continue

    if lookup_cache is not None:
        lookup_cache[cache_key] = formas
    return formas


def centralizar_caixa_no_slide(slide: Any, forma: Any) -> None:
    """Center a shape inside the current slide area.

    Args:
        slide: PowerPoint slide COM object.
        forma: Shape COM object to reposition.
    """
    try:
        ps = slide.Parent.PageSetup
        forma.Left = (ps.SlideWidth - forma.Width) / 2
        forma.Top = (ps.SlideHeight - forma.Height) / 2
    except (AttributeError, TypeError, ValueError, ZeroDivisionError) as exc:
        logger.debug("Could not center shape on slide: %s", exc, exc_info=True)


def justificar_texto(forma: Any) -> None:
    """Set paragraph alignment to justify when supported.

    Args:
        forma: Shape COM object that may expose text frames.
    """
    for attr in ("TextFrame", "TextFrame2"):
        try:
            getattr(forma, attr).TextRange.ParagraphFormat.Alignment = PP_ALIGN_JUSTIFY
        except (AttributeError, TypeError, ValueError) as exc:
            logger.debug("Could not justify text for %s: %s", attr, exc, exc_info=True)


def substituir_placeholder(
    slide: Any,
    placeholders: tuple[str, ...],
    novo_texto: str,
    centralizar_caixa: bool = False,
    justificar: bool = False,
    lookup_cache: dict[int, list[Any]] | None = None,
) -> bool:
    """Replace placeholders on a slide with the provided text.

    Args:
        slide: PowerPoint slide COM object.
        placeholders: Accepted placeholder variants.
        novo_texto: Replacement text.
        centralizar_caixa: Whether to center the updated shape.
        justificar: Whether to justify updated text paragraphs.
        lookup_cache: Optional cache of text-capable slide shapes.

    Returns:
        bool: ``True`` if at least one replacement was applied.
    """
    substituiu = False
    placeholders_norm = _placeholders_normalizados(placeholders)

    for forma in _formas_texto_do_slide(slide, lookup_cache):
        try:
            tr = forma.TextFrame.TextRange
            texto_atual = tr.Text
            texto_norm = normalizar_placeholder(texto_atual)

            if texto_norm in placeholders_norm:
                tr.Text = novo_texto
                if justificar:
                    justificar_texto(forma)
                if centralizar_caixa:
                    centralizar_caixa_no_slide(slide, forma)
                substituiu = True
                continue

            texto_novo = texto_atual
            for p in placeholders:
                if p in texto_novo:
                    texto_novo = texto_novo.replace(p, novo_texto)

            if texto_novo != texto_atual:
                tr.Text = texto_novo
                if justificar:
                    justificar_texto(forma)
                if centralizar_caixa:
                    centralizar_caixa_no_slide(slide, forma)
                substituiu = True
        except (AttributeError, TypeError, ValueError):
            continue

    return substituiu


def exigir_placeholder(
    slide: Any,
    placeholders: tuple[str, ...],
    novo_texto: str,
    descricao_slide: str,
    nome_placeholder: str,
    centralizar_caixa: bool = False,
    justificar: bool = False,
    lookup_cache: dict[int, list[Any]] | None = None,
) -> None:
    """Require placeholder replacement and raise a friendly error if missing.

    Args:
        slide: PowerPoint slide COM object.
        placeholders: Accepted placeholder variants.
        novo_texto: Replacement text.
        descricao_slide: Human-readable slide context.
        nome_placeholder: Placeholder name shown in error messages.
        centralizar_caixa: Whether to center the updated shape.
        justificar: Whether to justify updated text paragraphs.
        lookup_cache: Optional cache of text-capable slide shapes.

    Raises:
        GeradorSalmosError: If no placeholder match is found.
    """
    if not substituir_placeholder(
        slide,
        placeholders,
        novo_texto,
        centralizar_caixa=centralizar_caixa,
        justificar=justificar,
        lookup_cache=lookup_cache,
    ):
        raise GeradorSalmosError(
            f"Placeholder '{nome_placeholder}' não encontrado no {descricao_slide}.\n\n"
            "Verifique se o modelo.pptx contém os marcadores esperados."
        )


def validar_modelo(caminho_modelo: Path) -> None:
    """Validate model file existence and type.

    Args:
        caminho_modelo: Path to the PowerPoint model file.

    Raises:
        GeradorSalmosError: If the model path is missing or invalid.
    """
    if not caminho_modelo.exists():
        raise GeradorSalmosError(
            f"Arquivo modelo.pptx não encontrado:\n{caminho_modelo}\n\n"
            "Coloque o arquivo modelo.pptx na mesma pasta do programa."
        )
    if not caminho_modelo.is_file():
        raise GeradorSalmosError(f"O caminho não aponta para um arquivo válido:\n{caminho_modelo}")


def escolher_indice_template(indice_versiculo: int, total: int) -> int:
    """Choose the template slide index for a verse position.

    Args:
        indice_versiculo: Zero-based verse index.
        total: Total number of verses to generate.

    Returns:
        int: Template slide index to duplicate.
    """
    if indice_versiculo == total - 1:
        return TEMPLATE_FINAL_INDEX
    if indice_versiculo % 2 == 0:
        return TEMPLATE_PASTOR_INDEX
    return TEMPLATE_IGREJA_INDEX


def abrir_powerpoint() -> Any:
    """Launch a new PowerPoint COM application instance.

    Returns:
        Any: PowerPoint application COM object.

    Raises:
        GeradorSalmosError: If pywin32 is unavailable or PowerPoint cannot start.
    """
    if win32com is None or pythoncom is None:
        raise GeradorSalmosError(
            "A biblioteca pywin32 não está instalada.\n\n"
            "Instale com:\n    py -m pip install -r requirements.txt"
        )
    try:
        app = win32com.client.DispatchEx("PowerPoint.Application")
        app.Visible = True
        return app
    except (AttributeError, OSError, RuntimeError) as exc:
        raise GeradorSalmosError(
            "Não foi possível abrir o Microsoft PowerPoint.\n\n"
            "Verifique se o PowerPoint está instalado corretamente neste computador."
        ) from exc


def gerar_powerpoint(
    texto_leitura: str,
    titulo: str,
    atualizar_progresso: Callable[[int], None] | None = None,
    substituir_existente: bool = False,
    config: dict[str, str] | None = None,
    verso_inicio: int | None = None,
    verso_fim: int | None = None,
) -> ResultadoGeracao:
    """Generate a PowerPoint presentation from scripture text.

    Args:
        texto_leitura: Raw scripture text input.
        titulo: User-facing reading title.
        atualizar_progresso: Optional callback receiving progress percentages.
        substituir_existente: Whether an existing output file can be overwritten.
        config: Optional user configuration mapping.
        verso_inicio: Optional 1-based first verse index.
        verso_fim: Optional 1-based final verse index (inclusive in UI intent).

    Returns:
        ResultadoGeracao: Metadata describing the generated file.

    Raises:
        GeradorSalmosError: For expected generation and validation failures.
        PermissionError: Propagated as user-facing errors where applicable.
    """
    inicio = time.perf_counter()

    titulo = titulo.strip()
    if not titulo:
        raise GeradorSalmosError("Informe o título antes de gerar a apresentação.")

    validar_modelo(MODEL_FILE)
    leitura = ler_texto_entrada(texto_leitura)

    total_original = len(leitura.versiculos)
    idx_inicio = (verso_inicio - 1) if verso_inicio is not None else 0
    idx_fim = verso_fim if verso_fim is not None else total_original
    idx_inicio = max(0, min(idx_inicio, total_original))
    idx_fim = max(idx_inicio, min(idx_fim, total_original))

    versiculos_selecionados = leitura.versiculos[idx_inicio:idx_fim]
    if not versiculos_selecionados:
        raise GeradorSalmosError(
            "O intervalo de versículos selecionado não contém versículos.\n\n"
            "Verifique os campos 'De' e 'Até' e tente novamente."
        )

    leitura = LeituraBiblica(livro=leitura.livro, capitulo=leitura.capitulo, versiculos=versiculos_selecionados)
    caminho_saida = caminho_saida_para(leitura, config).resolve()
    validar_pasta_saida(caminho_saida.parent)

    if caminho_saida.suffix.casefold() != ".pptx":
        raise GeradorSalmosError("O arquivo de saída deve ter extensão .pptx.")

    nome_limpo_esperado = f"{limpar_nome_arquivo(caminho_saida.stem)}.pptx"
    if caminho_saida.name != nome_limpo_esperado:
        raise GeradorSalmosError("O nome do arquivo de saída contém caracteres inválidos.")

    if caminho_saida.exists() and not caminho_saida.is_file():
        raise GeradorSalmosError("O caminho de saída selecionado não aponta para um arquivo válido.")

    logger.info(
        "Generation started | book=%s | chapter=%s | verses=%d (of %d) | title=%r | output=%s",
        leitura.livro,
        leitura.capitulo,
        len(leitura.versiculos),
        total_original,
        titulo,
        caminho_saida,
    )

    powerpoint: Any | None = None
    apresentacao: Any | None = None
    com_inicializado = False
    lookup_cache: dict[int, list[Any]] = {}

    def _progresso(valor: int) -> None:
        if atualizar_progresso:
            atualizar_progresso(valor)

    try:
        _progresso(10)
        if pythoncom is None or win32com is None:
            raise GeradorSalmosError(
                "A biblioteca pywin32 não está instalada.\n\n"
                "Instale com:\n    py -m pip install -r requirements.txt"
            )
        pythoncom.CoInitialize()
        com_inicializado = True
        logger.debug("COM initialised")

        powerpoint = abrir_powerpoint()
        logger.debug("PowerPoint instance created")

        _progresso(20)
        apresentacao = powerpoint.Presentations.Open(str(MODEL_FILE), ReadOnly=False, Untitled=False, WithWindow=False)
        slides = apresentacao.Slides
        logger.debug("Template opened | slides=%d", slides.Count)

        if slides.Count < 4:
            raise GeradorSalmosError(
                f"O modelo tem apenas {slides.Count} slide(s).\n\n"
                "São necessários pelo menos 4 slides: título, pastor, igreja e final."
            )

        slide_titulo = slides(TEMPLATE_TITLE_INDEX)
        exigir_placeholder(
            slide_titulo,
            PLACEHOLDERS_LIVRO_CAP,
            leitura.livro_capitulo,
            "slide de título",
            PLACEHOLDER_LIVRO_CAP,
            lookup_cache=lookup_cache,
        )
        exigir_placeholder(
            slide_titulo,
            PLACEHOLDERS_TITULO,
            titulo,
            "slide de título",
            PLACEHOLDER_TITULO,
            lookup_cache=lookup_cache,
        )
        logger.debug("Title slide filled")

        template_slides = {
            TEMPLATE_PASTOR_INDEX: slides(TEMPLATE_PASTOR_INDEX),
            TEMPLATE_IGREJA_INDEX: slides(TEMPLATE_IGREJA_INDEX),
            TEMPLATE_FINAL_INDEX: slides(TEMPLATE_FINAL_INDEX),
        }

        total = len(leitura.versiculos)
        for indice, versiculo in enumerate(leitura.versiculos):
            template_index = escolher_indice_template(indice, total)
            slide_modelo = template_slides[template_index]
            novo_slide = slide_modelo.Duplicate().Item(1)
            novo_slide.MoveTo(slides.Count)
            exigir_placeholder(
                novo_slide,
                PLACEHOLDERS_VERSICULO,
                versiculo,
                f"slide modelo {template_index}",
                PLACEHOLDER_VERSICULO,
                centralizar_caixa=True,
                justificar=True,
                lookup_cache=lookup_cache,
            )
            progresso = 25 + int(((indice + 1) / total) * 55)
            _progresso(progresso)

        logger.debug("Verse slides generated | count=%d", total)

        for _ in range(3):
            slides(TEMPLATE_PASTOR_INDEX).Delete()
        logger.debug("Template slides removed")

        _progresso(90)
        if caminho_saida.exists():
            if not substituir_existente:
                raise GeradorSalmosError(
                    f"O arquivo já existe:\n{caminho_saida.name}\n\n"
                    "Confirme a substituição antes de gerar novamente."
                )
            try:
                caminho_saida.unlink()
                logger.debug("Existing file removed | path=%s", caminho_saida)
            except PermissionError as exc:
                raise GeradorSalmosError(
                    f"Não foi possível substituir '{caminho_saida.name}'.\n\n"
                    "Feche o arquivo no PowerPoint e tente novamente."
                ) from exc

        apresentacao.SaveAs(str(caminho_saida), FileFormat=PP_SAVE_AS_OPENXML)
        _progresso(100)

        duracao = time.perf_counter() - inicio
        logger.info("Generation complete | file=%s | verses=%d | seconds=%.2f", caminho_saida.name, len(leitura.versiculos), duracao)
        return ResultadoGeracao(caminho=caminho_saida, quantidade_versiculos=len(leitura.versiculos), duracao_segundos=duracao)

    except GeradorSalmosError:
        logger.warning("Generation failed (expected error)", exc_info=True)
        raise
    except PermissionError as exc:
        logger.exception("Permission error while saving presentation")
        raise GeradorSalmosError(
            "Erro de permissão ao salvar o arquivo.\n\n"
            "Feche a apresentação no PowerPoint se estiver aberta e tente novamente."
        ) from exc
    except com_error_types() as exc:
        logger.exception("PowerPoint COM error")
        raise GeradorSalmosError(
            "O PowerPoint retornou um erro durante a geração.\n\n"
            "Possíveis causas:\n"
            "  • modelo.pptx aberto em modo somente leitura\n"
            "  • modelo.pptx não contém os marcadores esperados\n"
            "  • Arquivo corrompido ou protegido por senha"
        ) from exc
    except RECOVERABLE_ERRORS:
        logger.exception("Unexpected generation error")
        raise

    finally:
        if apresentacao is not None:
            try:
                apresentacao.Saved = True
                apresentacao.Close()
                logger.debug("Presentation closed")
            except RECOVERABLE_ERRORS + com_error_types():
                logger.exception("Could not close presentation")
            finally:
                del apresentacao

        if powerpoint is not None:
            try:
                powerpoint.Quit()
                logger.debug("PowerPoint quit")
            except RECOVERABLE_ERRORS + com_error_types():
                logger.exception("Could not quit PowerPoint")
            finally:
                del powerpoint

        if pythoncom is not None and com_inicializado:
            try:
                pythoncom.CoUninitialize()
                logger.debug("COM uninitialised")
            except RECOVERABLE_ERRORS + com_error_types():
                logger.exception("Could not uninitialise COM")
