from __future__ import annotations

from datetime import datetime
import sys
from pathlib import Path

import json
import logging
import logging.handlers
import re
import subprocess
import sys
import threading
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

try:
    import customtkinter as ctk
    CTK_AVAILABLE = True
except ImportError:
    CTK_AVAILABLE = False

import tkinter as tk
from tkinter import messagebox, filedialog

try:
    from PIL import Image, ImageDraw, ImageFilter, ImageTk
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

try:
    import pythoncom
    import pywintypes
    import win32com.client
except ImportError:
    pythoncom = None
    pywintypes = None
    win32com = None


# ── Application identity ───────────────────────────────────────────────────────
APP_NAME    = "Gerador de Leitura Bíblica"
APP_VERSION = "2.0"
MODEL_NAME  = "modelo.pptx"
APP_AUTHOR  = "Igreja Tabernáculo do Senhor - Barra do Garças"


# ── Path resolution ────────────────────────────────────────────────────────────

def app_dir() -> Path:
    """Directory that contains the executable (frozen) or this script (source)."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def resource_path(name: str) -> Path:
    """Return the absolute path to *name*.

    Resolution order:
    1. Alongside the executable / script — lets users replace modelo.pptx
       without rebuilding the bundle.
    2. Inside the PyInstaller _MEIPASS bundle directory.
    """
    external = app_dir() / name
    if external.exists():
        return external.resolve()
    bundle_dir = Path(getattr(sys, "_MEIPASS", app_dir()))
    return (bundle_dir / name).resolve()


import os

APP_DIR = app_dir()

# Arquivos graváveis do usuário
USER_DATA_DIR = Path(os.getenv("LOCALAPPDATA")) / "LeituraBiblica"
USER_DATA_DIR.mkdir(parents=True, exist_ok=True)

MODEL_FILE = resource_path(MODEL_NAME)
LOG_FILE = USER_DATA_DIR / "logs.log"
CONFIG_FILE = USER_DATA_DIR / "config.json"

# Default output folder — overridable via config.json key "output_dir".
OUTPUT_DIR_DEFAULT = Path.home() / "Documents" / "LeituraBiblica"


# ── PowerPoint constants ───────────────────────────────────────────────────────
TEMPLATE_TITLE_INDEX  = 1
TEMPLATE_PASTOR_INDEX = 2
TEMPLATE_IGREJA_INDEX = 3
TEMPLATE_FINAL_INDEX  = 4

PP_ALIGN_JUSTIFY  = 4
PP_SAVE_AS_OPENXML = 24   # ppSaveAsOpenXMLPresentation (.pptx)

# Placeholder aliases recognised in the template
PLACEHOLDER_LIVRO_CAP = "{livro_cap}"
PLACEHOLDER_TITULO    = "{titulo}"
PLACEHOLDER_VERSICULO = "{versiculo}"

PLACEHOLDERS_LIVRO_CAP = (PLACEHOLDER_LIVRO_CAP, "livro cap", "livro_cap")
PLACEHOLDERS_TITULO    = (PLACEHOLDER_TITULO, "titulo", "título")
PLACEHOLDERS_VERSICULO = (
    PLACEHOLDER_VERSICULO,
    "versiculo",
    "versículo",
    "versImpar",
    "versPar",
    "versLast",
    "1 - versImpar",
    "2 - versPar",
    "3 - versLast",
    "1 – versImpar",
    "2 – versPar",
    "3 – versLast",
)


# ── Design tokens ──────────────────────────────────────────────────────────────
# Fluent Blue Dashboard — dark surfaces with blue accent palette.
BG_APP        = "#0a0e1a"   # window backdrop — near-black, faint blue tint
BG_CARD       = "#111827"   # main floating card surface (lighter = "raised")
BG_CARD_INNER = "#0d1520"   # recessed surface inside the card (kept for status/fade math)
BG_TILE       = "#162033"   # nested "widget tile" surface (range box, progress, button row)
BG_INPUT      = "#0f1a2e"   # input fields / textbox well
BG_BTN        = "#3b82f6"
BG_BTN_HOVER  = "#2563eb"
BG_PROG_TRACK = "#1a2540"
BG_PROG_FILL  = "#3b82f6"

# Gradient accent stops — sky blue → blue → deep blue
GRAD_START    = "#60a5fa"   # sky blue
GRAD_MID      = "#3b82f6"   # blue-500
GRAD_END      = "#1d4ed8"   # blue-700

# Borders — kept faint; elevation reads from surface-shade + drop shadow.
CARD_BORDER      = "#1d2d4a"
CARD_BORDER_TOP  = "#2a3f6a"
INPUT_BORDER     = "#1d2d4a"
INPUT_BORDER_FOC = "#3b82f6"

FG_TITLE   = "#eaf0ff"
FG_LABEL   = "#7f9ab3"
FG_INPUT   = "#d9eaff"
FG_HINT    = "#5a7a9a"      # improved contrast (was #3d5070, ~2.1:1 → now ~3.5:1)
FG_SUCCESS = "#34d399"
FG_ERROR   = "#f87171"
FG_WARNING = "#fbbf24"
FG_WHITE   = "#ffffff"
ACCENT_ICON = "#3b82f6"

FONT_FAMILY = "Segoe UI"
RADIUS       = 22   # primary floating widgets (main card)
RADIUS_TILE  = 18   # nested widget tiles
RADIUS_PILL  = 20   # pill-shaped controls (buttons, toggle, entries)
SHADOW_COLOR = "#02050d"   # colour the floating drop-shadow is rendered in

# Icons (Unicode)
ICON_BOOK  = "📖"
ICON_PLAY  = "▶"
ICON_CHECK = "✓"
ICON_CROSS = "✗"
ICON_WARN  = "⚠"
ICON_FOLDER = "📂"


# ── Logging ────────────────────────────────────────────────────────────────────

def _setup_logging() -> logging.Logger:
    """Configure rotating-file + stderr logging and return the app logger."""
    fmt = logging.Formatter(
        "%(asctime)s | %(levelname)-8s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    # Rotate at 1 MB, keep 3 backups so logs don't grow forever.
    file_handler = logging.handlers.RotatingFileHandler(
        LOG_FILE, maxBytes=1_000_000, backupCount=3, encoding="utf-8"
    )
    file_handler.setFormatter(fmt)
    file_handler.setLevel(logging.DEBUG)

    root = logging.getLogger()
    root.setLevel(logging.DEBUG)
    root.addHandler(file_handler)

    log = logging.getLogger(APP_NAME)
    log.info("=" * 60)
    log.info("Session started | version=%s | python=%s", APP_VERSION, sys.version.split()[0])
    log.info("app_dir=%s", APP_DIR)
    log.info("model=%s | exists=%s", MODEL_FILE, MODEL_FILE.exists())
    return log


logger = _setup_logging()


# ── Domain errors ──────────────────────────────────────────────────────────────

class GeradorSalmosError(Exception):
    """Expected, user-facing error displayed as a friendly dialog."""


# ── Domain models ──────────────────────────────────────────────────────────────

@dataclass(frozen=True)
class LeituraBiblica:
    livro: str
    capitulo: str
    versiculos: list[str]

    @property
    def livro_capitulo(self) -> str:
        return f"{self.livro} {self.capitulo}"

    @property
    def nome_arquivo(self) -> str:
        return f"{limpar_nome_arquivo(self.livro_capitulo)}.pptx"


@dataclass(frozen=True)
class ResultadoGeracao:
    caminho: Path
    quantidade_versiculos: int
    duracao_segundos: float


# ── Utility functions ──────────────────────────────────────────────────────────

def limpar_nome_arquivo(nome: str) -> str:
    """Strip characters illegal in Windows filenames."""
    nome_limpo = re.sub(r'[<>:"/\\|?*]', "", nome).strip()
    nome_limpo = re.sub(r"\s+", " ", nome_limpo)
    if not nome_limpo:
        raise GeradorSalmosError("Não foi possível montar o nome do arquivo final.")
    return nome_limpo


def output_dir(config: dict[str, str] | None = None) -> Path:
    """Return the resolved output directory, creating it when necessary."""
    cfg = config or carregar_configuracao()
    raw = cfg.get("output_dir", "").strip()
    folder = Path(raw) if raw else OUTPUT_DIR_DEFAULT
    try:
        folder = folder.resolve()
        folder.mkdir(parents=True, exist_ok=True)
    except Exception:
        logger.exception("Could not create output directory %s — falling back to default", folder)
        folder = OUTPUT_DIR_DEFAULT.resolve()
        folder.mkdir(parents=True, exist_ok=True)
    return folder


def caminho_saida_para(leitura: LeituraBiblica, config: dict[str, str] | None = None) -> Path:
    return output_dir(config) / leitura.nome_arquivo


# ── Configuration persistence ──────────────────────────────────────────────────

def carregar_configuracao() -> dict[str, str]:
    try:
        if CONFIG_FILE.exists():
            data = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                return {str(k): str(v) for k, v in data.items()}
    except Exception:
        logger.exception("Could not load config from %s", CONFIG_FILE)
    return {}


def salvar_configuracao(config: dict[str, str]) -> None:
    try:
        CONFIG_FILE.write_text(
            json.dumps(config, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
    except Exception:
        logger.exception("Could not save config to %s", CONFIG_FILE)


# ── Input parsing ──────────────────────────────────────────────────────────────

def ler_texto_entrada(texto: str) -> LeituraBiblica:
    """Parse the raw scripture text and return a LeituraBiblica.

    Expected format::

        Livro: Salmos
        Capítulo: 23
        1 O Senhor é o meu pastor…
        2 Ele me faz repousar…
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

        match = re.match(r"^(\d+)\s+(.+)$", linha)
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


# ── PowerPoint shape helpers ───────────────────────────────────────────────────

def iterar_formas(forma):
    """Yield *forma* and, recursively, every shape inside a group."""
    yield forma
    try:
        if forma.Type == 6:  # msoGroup
            for i in range(1, forma.GroupItems.Count + 1):
                yield from iterar_formas(forma.GroupItems(i))
    except Exception:
        return


def normalizar_placeholder(texto: str) -> str:
    texto = texto.replace("\r", "").replace("\n", " ").strip()
    texto = texto.replace("–", "-").replace("—", "-")
    texto = re.sub(r"\s+", " ", texto)
    return texto.casefold()


def centralizar_caixa_no_slide(slide, forma) -> None:
    try:
        ps = slide.Parent.PageSetup
        forma.Left = (ps.SlideWidth  - forma.Width)  / 2
        forma.Top  = (ps.SlideHeight - forma.Height) / 2
    except Exception:
        pass


def justificar_texto(forma) -> None:
    for attr in ("TextFrame", "TextFrame2"):
        try:
            getattr(forma, attr).TextRange.ParagraphFormat.Alignment = PP_ALIGN_JUSTIFY
        except Exception:
            pass


def substituir_placeholder(
    slide,
    placeholders: tuple[str, ...],
    novo_texto: str,
    centralizar_caixa: bool = False,
    justificar: bool = False,
) -> bool:
    """Replace all occurrences of *placeholders* in *slide* with *novo_texto*.

    Returns True if at least one substitution was made.
    """
    substituiu = False
    placeholders_norm = {normalizar_placeholder(p) for p in placeholders}

    for forma_raiz in slide.Shapes:
        for forma in iterar_formas(forma_raiz):
            try:
                if not forma.HasTextFrame or not forma.TextFrame.HasText:
                    continue

                tr = forma.TextFrame.TextRange
                texto_atual = tr.Text
                texto_norm  = normalizar_placeholder(texto_atual)

                # Exact match — replace the whole text range.
                if texto_norm in placeholders_norm:
                    tr.Text = novo_texto
                    if justificar:
                        justificar_texto(forma)
                    if centralizar_caixa:
                        centralizar_caixa_no_slide(slide, forma)
                    substituiu = True
                    continue

                # Partial match — replace only the token inside the text.
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

            except Exception:
                continue

    return substituiu


def exigir_placeholder(
    slide,
    placeholders: tuple[str, ...],
    novo_texto: str,
    descricao_slide: str,
    nome_placeholder: str,
    centralizar_caixa: bool = False,
    justificar: bool = False,
) -> None:
    if not substituir_placeholder(
        slide,
        placeholders,
        novo_texto,
        centralizar_caixa=centralizar_caixa,
        justificar=justificar,
    ):
        raise GeradorSalmosError(
            f"Placeholder '{nome_placeholder}' não encontrado no {descricao_slide}.\n\n"
            "Verifique se o modelo.pptx contém os marcadores esperados."
        )


# ── Template validation ────────────────────────────────────────────────────────

def validar_modelo(caminho_modelo: Path) -> None:
    if not caminho_modelo.exists():
        raise GeradorSalmosError(
            f"Arquivo modelo.pptx não encontrado:\n{caminho_modelo}\n\n"
            "Coloque o arquivo modelo.pptx na mesma pasta do programa."
        )
    if not caminho_modelo.is_file():
        raise GeradorSalmosError(
            f"O caminho não aponta para um arquivo válido:\n{caminho_modelo}"
        )


# ── Slide template selection ───────────────────────────────────────────────────

def escolher_indice_template(indice_versiculo: int, total: int) -> int:
    if indice_versiculo == total - 1:
        return TEMPLATE_FINAL_INDEX
    if indice_versiculo % 2 == 0:
        return TEMPLATE_PASTOR_INDEX
    return TEMPLATE_IGREJA_INDEX


# ── PowerPoint generation ──────────────────────────────────────────────────────

def abrir_powerpoint():
    """Launch a new PowerPoint instance and return the COM application object."""
    if win32com is None or pythoncom is None:
        raise GeradorSalmosError(
            "A biblioteca pywin32 não está instalada.\n\n"
            "Instale com:\n    py -m pip install -r requirements.txt"
        )
    try:
        app = win32com.client.DispatchEx("PowerPoint.Application")
        app.Visible = True
        return app
    except Exception as exc:
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
    """Generate the presentation and return a ResultadoGeracao on success.

    *verso_inicio* and *verso_fim* are 1-based verse numbers.  Both are
    optional:
      • Neither set → generate all verses.
      • Only *verso_fim* set → generate from verse 1 to *verso_fim*.
      • Both set → generate [verso_inicio .. verso_fim] inclusive.
    """
    inicio = time.perf_counter()

    titulo = titulo.strip()
    if not titulo:
        raise GeradorSalmosError("Informe o título antes de gerar a apresentação.")

    validar_modelo(MODEL_FILE)
    leitura = ler_texto_entrada(texto_leitura)

    # ── Apply verse range ─────────────────────────────────────────────────────
    total_original = len(leitura.versiculos)
    idx_inicio = (verso_inicio - 1) if verso_inicio is not None else 0
    idx_fim    = verso_fim if verso_fim is not None else total_original  # exclusive

    # Guard against out-of-range values (should already be validated in the UI
    # layer, but we double-check here for safety).
    idx_inicio = max(0, min(idx_inicio, total_original))
    idx_fim    = max(idx_inicio, min(idx_fim, total_original))

    versiculos_selecionados = leitura.versiculos[idx_inicio:idx_fim]
    if not versiculos_selecionados:
        raise GeradorSalmosError(
            "O intervalo de versículos selecionado não contém versículos.\n\n"
            "Verifique os campos 'De' e 'Até' e tente novamente."
        )

    # Rebuild leitura with the filtered verses so the rest of the function
    # stays unchanged.
    leitura = LeituraBiblica(
        livro=leitura.livro,
        capitulo=leitura.capitulo,
        versiculos=versiculos_selecionados,
    )

    caminho_saida = caminho_saida_para(leitura, config)

    logger.info(
        "Generation started | book=%s | chapter=%s | verses=%d (of %d) | title=%r | output=%s",
        leitura.livro,
        leitura.capitulo,
        len(leitura.versiculos),
        total_original,
        titulo,
        caminho_saida,
    )

    powerpoint = None
    apresentacao = None
    com_inicializado = False

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
        # Open as read-write; WithWindow=False keeps the window hidden.
        # The application is already Visible=True so PowerPoint won't hang.
        apresentacao = powerpoint.Presentations.Open(
            str(MODEL_FILE), ReadOnly=False, Untitled=False, WithWindow=False
        )
        logger.debug("Template opened | slides=%d", apresentacao.Slides.Count)

        if apresentacao.Slides.Count < 4:
            raise GeradorSalmosError(
                f"O modelo tem apenas {apresentacao.Slides.Count} slide(s).\n\n"
                "São necessários pelo menos 4 slides: título, pastor, igreja e final."
            )

        # ── Slide 1: title ────────────────────────────────────────────────────
        slide_titulo = apresentacao.Slides(TEMPLATE_TITLE_INDEX)
        exigir_placeholder(
            slide_titulo, PLACEHOLDERS_LIVRO_CAP,
            leitura.livro_capitulo, "slide de título", PLACEHOLDER_LIVRO_CAP,
        )
        exigir_placeholder(
            slide_titulo, PLACEHOLDERS_TITULO,
            titulo, "slide de título", PLACEHOLDER_TITULO,
        )
        logger.debug("Title slide filled")

        # ── Verse slides ──────────────────────────────────────────────────────
        total = len(leitura.versiculos)
        for indice, versiculo in enumerate(leitura.versiculos):
            template_index = escolher_indice_template(indice, total)
            slide_modelo = apresentacao.Slides(template_index)
            novo_slide = slide_modelo.Duplicate().Item(1)
            novo_slide.MoveTo(apresentacao.Slides.Count)
            exigir_placeholder(
                novo_slide, PLACEHOLDERS_VERSICULO,
                versiculo, f"slide modelo {template_index}", PLACEHOLDER_VERSICULO,
                centralizar_caixa=True, justificar=True,
            )
            progresso = 25 + int(((indice + 1) / total) * 55)
            _progresso(progresso)

        logger.debug("Verse slides generated | count=%d", total)

        # ── Remove template slides ────────────────────────────────────────────
        # Slides 2, 3, 4 are the templates. Deleting at index 2 three times
        # removes all three because the deck shifts down after each deletion.
        for _ in range(3):
            apresentacao.Slides(TEMPLATE_PASTOR_INDEX).Delete()
        logger.debug("Template slides removed")

        # ── Save ──────────────────────────────────────────────────────────────
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

        # FileFormat=24 forces .pptx (OpenXML) regardless of the template format.
        apresentacao.SaveAs(str(caminho_saida), FileFormat=PP_SAVE_AS_OPENXML)
        _progresso(100)

        duracao = time.perf_counter() - inicio
        logger.info(
            "Generation complete | file=%s | verses=%d | seconds=%.2f",
            caminho_saida.name,
            len(leitura.versiculos),
            duracao,
        )
        return ResultadoGeracao(
            caminho=caminho_saida,
            quantidade_versiculos=len(leitura.versiculos),
            duracao_segundos=duracao,
        )

    except GeradorSalmosError:
        logger.warning("Generation failed (expected error)", exc_info=True)
        raise
    except PermissionError as exc:
        logger.exception("Permission error while saving presentation")
        raise GeradorSalmosError(
            "Erro de permissão ao salvar o arquivo.\n\n"
            "Feche a apresentação no PowerPoint se estiver aberta e tente novamente."
        ) from exc
    except Exception as exc:
        if pywintypes is not None and isinstance(exc, pywintypes.com_error):
            logger.exception("PowerPoint COM error")
            raise GeradorSalmosError(
                "O PowerPoint retornou um erro durante a geração.\n\n"
                "Possíveis causas:\n"
                "  • modelo.pptx aberto em modo somente leitura\n"
                "  • modelo.pptx não contém os marcadores esperados\n"
                "  • Arquivo corrompido ou protegido por senha"
            ) from exc
        logger.exception("Unexpected generation error")
        raise

    finally:
        # Release COM objects before CoUninitialize to avoid dangling references.
        # Setting Saved=True suppresses the "save changes?" dialog on error paths.
        if apresentacao is not None:
            try:
                apresentacao.Saved = True
                apresentacao.Close()
                logger.debug("Presentation closed")
            except Exception:
                logger.exception("Could not close presentation")
            finally:
                del apresentacao

        if powerpoint is not None:
            try:
                powerpoint.Quit()
                logger.debug("PowerPoint quit")
            except Exception:
                logger.exception("Could not quit PowerPoint")
            finally:
                del powerpoint

        if pythoncom is not None and com_inicializado:
            try:
                pythoncom.CoUninitialize()
                logger.debug("COM uninitialised")
            except Exception:
                logger.exception("Could not uninitialise COM")


# ── Application window ─────────────────────────────────────────────────────────

class Aplicacao(ctk.CTk if CTK_AVAILABLE else tk.Tk):

    # Placeholder text shown in the scripture field when empty.
    _SCRIPTURE_PLACEHOLDER = "Cole aqui o texto copiado do Busca."

    # ══════════════════════════════════════════════════════════════════════════
    # ── Fluent Design Animation Engine ───────────────────────────────────────
    # All animation primitives live here. Zero business logic.
    # ══════════════════════════════════════════════════════════════════════════

    # ── Easing curves (t → t′) ───────────────────────────────────────────────

    @staticmethod
    def _ease_out_cubic(t: float) -> float:
        """Deceleration — objects arrive smoothly."""
        return 1 - (1 - t) ** 3

    @staticmethod
    def _ease_in_cubic(t: float) -> float:
        """Acceleration — objects leave quickly."""
        return t ** 3

    @staticmethod
    def _ease_in_out_quint(t: float) -> float:
        """Fluent standard ease — symmetric, smooth peak."""
        if t < 0.5:
            return 16 * t ** 5
        return 1 - (-2 * t + 2) ** 5 / 2

    @staticmethod
    def _ease_spring(t: float) -> float:
        """Spring overshoot — matches Fluent spring physics."""
        c4 = (2 * 3.14159265) / 3
        if t == 0:
            return 0
        if t == 1:
            return 1
        import math
        return 2 ** (-8 * t) * math.sin((t * 10 - 0.75) * c4) + 1

    # ── Core tween scheduler ─────────────────────────────────────────────────

    def _tween(
        self,
        duration_ms: int,
        on_update,            # callable(progress: float)
        ease=None,
        on_done=None,         # callable() | None
        fps: int = 60,
        cancel_token: list | None = None,  # [False] — set to [True] to cancel mid-flight
    ) -> None:
        """Drive *on_update* from 0.0→1.0 over *duration_ms* milliseconds.

        Pass a ``cancel_token=[False]`` list and set it to ``[True]`` from
        outside to abort the tween early (on_done will NOT be called).
        """
        if ease is None:
            ease = self._ease_out_cubic
        interval = max(1, 1000 // fps)
        steps = max(1, duration_ms // interval)
        step_ref = [0]

        def _tick():
            if cancel_token is not None and cancel_token[0]:
                return
            step_ref[0] += 1
            raw = step_ref[0] / steps
            t = min(raw, 1.0)
            try:
                on_update(ease(t))
            except Exception:
                pass
            if t < 1.0:
                self.after(interval, _tick)
            elif on_done:
                try:
                    on_done()
                except Exception:
                    pass

        self.after(0, _tick)

    # ── 1. Fade In / Fade Out ─────────────────────────────────────────────────
    # Tkinter has no native alpha per-widget; we simulate fade by blending
    # the widget's fg_color toward the parent bg colour.

    @staticmethod
    def _blend_hex(c1: str, c2: str, t: float) -> str:
        """Linear blend c1 → c2 at fraction t."""
        r1, g1, b1 = int(c1[1:3], 16), int(c1[3:5], 16), int(c1[5:7], 16)
        r2, g2, b2 = int(c2[1:3], 16), int(c2[3:5], 16), int(c2[5:7], 16)
        r = int(r1 + (r2 - r1) * t)
        g = int(g1 + (g2 - g1) * t)
        b = int(b1 + (b2 - b1) * t)
        return f"#{r:02x}{g:02x}{b:02x}"

    def _fade_widget(
        self,
        widget,
        from_color: str,
        to_color: str,
        duration_ms: int = 220,
        attr: str = "text_color",
        on_done=None,
        cancel_token: list | None = None,
    ) -> None:
        """Fade a widget's *attr* from *from_color* to *to_color*."""
        def _update(t):
            try:
                widget.configure(**{attr: self._blend_hex(from_color, to_color, t)})
            except Exception:
                pass
        self._tween(duration_ms, _update, ease=self._ease_in_out_quint,
                    on_done=on_done, cancel_token=cancel_token)

    # ── 2. Fluent Entrance Animation (Fade + slide-up + slide-in) ─────────────
    # Since Tkinter widgets can't be made transparent, we animate the pady/padx
    # offsets to simulate a slide-up + slide-in-from-the-right entrance,
    # cascading like the staggered card reveal in the Fluent reference clip.

    def _entrance_animation(
        self,
        widget,
        parent_grid_kw: dict,
        delay_ms: int = 0,
        slide_px: int = 18,
        slide_x_px: int = 0,
        ease=None,
        on_done=None,
    ) -> None:
        """Animate a grid widget sliding up + in from the right, fading-in feel."""
        original_pady = parent_grid_kw.get("pady", 0)
        original_padx = parent_grid_kw.get("padx", 0)

        if isinstance(original_pady, tuple):
            base_top, base_bot = original_pady
        else:
            base_top = base_bot = original_pady

        if isinstance(original_padx, tuple):
            base_left, base_right = original_padx
        else:
            base_left = base_right = original_padx

        def _update(t):
            off_y = int(slide_px * (1 - t))
            off_x = int(slide_x_px * (1 - t))
            try:
                if slide_x_px:
                    widget.grid_configure(
                        pady=(base_top + off_y, base_bot),
                        padx=(base_left + off_x, max(0, base_right - off_x)),
                    )
                else:
                    widget.grid_configure(pady=(base_top + off_y, base_bot))
            except Exception:
                pass

        self.after(delay_ms, lambda: self._tween(
            260, _update, ease=ease or self._ease_out_cubic,
            on_done=on_done,
        ))

    # ── 3. Scale-In / Scale-Out (window geometry pulse) ───────────────────────
    # Simulated via brief geometry shrink/expand on the root window on launch.

    def _scale_in_window(self) -> None:
        """Fluent scale-in: window grows from 96 % to 100 % on startup.

        Guards against being called before the window is fully mapped —
        ``winfo_width()`` returns 1 in that state, which would cause the
        window to collapse to a 1×1 pixel black square mid-animation.
        """
        try:
            self.update_idletasks()
            w = self.winfo_width()
            h = self.winfo_height()
            x = self.winfo_x()
            y = self.winfo_y()
        except Exception:
            return

        # If the window hasn't been drawn yet tkinter reports 1×1 — retry.
        if w <= 1 or h <= 1:
            self.after(40, self._scale_in_window)
            return

        start_w = int(w * 0.94)
        start_h = int(h * 0.94)

        def _update(t):
            cw = int(start_w + (w - start_w) * t)
            ch = int(start_h + (h - start_h) * t)
            cx = x + (w - cw) // 2
            cy = y + (h - ch) // 2
            try:
                self.geometry(f"{cw}x{ch}+{cx}+{cy}")
            except Exception:
                pass

        self._tween(300, _update, ease=self._ease_spring)

    # ── 4. Reveal Animation — border glow on hover ───────────────────────────

    def _attach_reveal(
        self,
        widget,
        normal_border: str = CARD_BORDER,
        reveal_border: str = GRAD_START,
        duration_ms: int = 180,
    ) -> None:
        """Animate border color toward accent on hover (Fluent Reveal).

        Uses a shared cancel_token so an in-flight tween is always cancelled
        before the reverse tween starts — prevents colour jumps on fast
        enter/leave sequences.
        """
        _token: list = [False]   # mutable so both closures share it
        _current_border: list = [normal_border]
        _inside: list = [False]  # guard: run enter-anim only once per hover

        def _on_enter(_e):
            nonlocal _token
            if _inside[0]:
                return           # cursor is already inside — don't re-trigger
            _inside[0] = True
            _token[0] = True          # cancel any running leave-fade
            _token = [False]
            from_color = _current_border[0]
            _current_border[0] = reveal_border

            def _update(t):
                c = self._blend_hex(from_color, reveal_border, t)
                _current_border[0] = c
                try:
                    widget.configure(border_color=c)
                except Exception:
                    pass

            self._tween(duration_ms, _update, ease=self._ease_out_cubic,
                        cancel_token=_token)

        def _on_leave(_e):
            nonlocal _token
            if not _inside[0]:
                return           # wasn't inside — don't re-trigger
            _inside[0] = False
            _token[0] = True          # cancel any running enter-fade
            _token = [False]
            from_color = _current_border[0]
            _current_border[0] = normal_border

            def _update(t):
                c = self._blend_hex(from_color, normal_border, t)
                _current_border[0] = c
                try:
                    widget.configure(border_color=c)
                except Exception:
                    pass

            self._tween(duration_ms * 2, _update, ease=self._ease_out_cubic,
                        cancel_token=_token)

        try:
            widget.bind("<Enter>", _on_enter, add="+")
            widget.bind("<Leave>", _on_leave, add="+")
        except Exception:
            pass

    # ── 4b. Focus-only Reveal — border glow on focus, never on hover ─────────
    # Used for widgets (e.g. the Scripture textbox) where a hover-driven
    # border animation indirectly disturbs the caret (insertwidth /
    # insertbackground) on every <Enter>/<Leave>. Only FocusIn/FocusOut
    # may run here — hover must never call this.

    def _attach_focus_reveal(
        self,
        widget,
        normal_border: str = CARD_BORDER,
        reveal_border: str = GRAD_START,
        duration_ms: int = 200,
    ) -> None:
        _token: list = [False]
        _current: list = [normal_border]

        def _on_focus_in(_e):
            nonlocal _token
            _token[0] = True
            _token = [False]
            from_color = _current[0]

            def _update(t):
                c = self._blend_hex(from_color, reveal_border, t)
                _current[0] = c
                try:
                    widget.configure(border_color=c)
                except Exception:
                    pass

            self._tween(duration_ms, _update, ease=self._ease_out_cubic, cancel_token=_token)

        def _on_focus_out(_e):
            nonlocal _token
            _token[0] = True
            _token = [False]
            from_color = _current[0]

            def _update(t):
                c = self._blend_hex(from_color, normal_border, t)
                _current[0] = c
                try:
                    widget.configure(border_color=c)
                except Exception:
                    pass

            self._tween(duration_ms, _update, ease=self._ease_out_cubic, cancel_token=_token)

        try:
            widget.bind("<FocusIn>",  _on_focus_in,  add="+")
            widget.bind("<FocusOut>", _on_focus_out, add="+")
        except Exception:
            pass

    # ── 5. Hover Lift — card/button lifts by adjusting padding ───────────────

    def _attach_hover_lift(
        self,
        widget,
        lift_px: int = 3,
        duration_ms: int = 140,
    ) -> None:
        """Simulate a 'lift' on hover by padding-top shrink (shadow illusion).

        Base pady is captured once at bind time and never re-read at event
        time, so repeated hovers cannot accumulate drift.
        Cancel tokens prevent concurrent up/down tweens from fighting each other.
        """
        # Capture the *original* grid padding once, before any animation touches it.
        try:
            info = widget.grid_info()
            pady = info.get("pady", (0, 0))
            if isinstance(pady, int):
                pady = (pady, pady)
            _base_top, _base_bot = int(pady[0]), int(pady[1])
        except Exception:
            _base_top, _base_bot = 0, 0

        _token: list = [False]
        _lifted = [False]

        def _on_enter(_e):
            nonlocal _token
            if _lifted[0]:
                return
            _lifted[0] = True
            _token[0] = True
            _token = [False]

            def _up(t):
                off = int(lift_px * t)
                try:
                    widget.grid_configure(pady=(max(0, _base_top - off), _base_bot + off))
                except Exception:
                    pass

            self._tween(duration_ms, _up, ease=self._ease_out_cubic, cancel_token=_token)

        def _on_leave(_e):
            nonlocal _token
            if not _lifted[0]:
                return
            _lifted[0] = False
            _token[0] = True
            _token = [False]

            def _down(t):
                off = int(lift_px * (1 - t))
                try:
                    widget.grid_configure(pady=(max(0, _base_top - off), _base_bot + off))
                except Exception:
                    pass

            self._tween(duration_ms * 2, _down, ease=self._ease_out_cubic, cancel_token=_token)

        def _ignore_motion(_e):
            # Motion must never (re)start the lift — only Enter/Leave may.
            # The _lifted guard above already makes _on_enter a no-op while
            # inside, but we also swallow Motion explicitly so no other code
            # path (e.g. CTk's own internal hover bookkeeping) can be
            # mistaken for a fresh Enter.
            return

        try:
            widget.bind("<Enter>", _on_enter, add="+")
            widget.bind("<Leave>", _on_leave, add="+")
            widget.bind("<Motion>", _ignore_motion, add="+")
        except Exception:
            pass

    # ── 6. Press Animation — brief scale-down on click ───────────────────────

    def _attach_press_animation(self, widget) -> None:
        """Fluent press: fill darkens and the button settles inward slightly,
        then springs back on release — a tactile press, not just a color
        change."""
        _token: list = [False]

        try:
            info = widget.grid_info()
            pady = info.get("pady", (0, 0))
            if isinstance(pady, int):
                pady = (pady, pady)
            _base_top, _base_bot = int(pady[0]), int(pady[1])
        except Exception:
            _base_top, _base_bot = 0, 0

        def _on_press(_e):
            try:
                widget.configure(fg_color=GRAD_MID)
                widget.grid_configure(pady=(_base_top + 2, _base_bot + 2))
            except Exception:
                pass

        def _on_release(_e):
            nonlocal _token
            _token[0] = True
            _token = [False]

            def _restore(t):
                try:
                    c = self._blend_hex(GRAD_MID, GRAD_START, t)
                    widget.configure(fg_color=c)
                    settle = int(2 * (1 - t))
                    widget.grid_configure(pady=(_base_top + settle, _base_bot + settle))
                except Exception:
                    pass
            self._tween(180, _restore, ease=self._ease_spring, cancel_token=_token)

        try:
            widget.bind("<ButtonPress-1>",   _on_press,   add="+")
            widget.bind("<ButtonRelease-1>", _on_release, add="+")
        except Exception:
            pass

    # ── 7. Acrylic Reveal — input border shimmer on focus ────────────────────

    def _attach_acrylic_focus(self, widget) -> None:
        """Animate border from INPUT_BORDER → cyan glow on focus (acrylic effect).

        Uses a cancel token so a focus-in tween is cleanly aborted when
        focus-out fires before the tween finishes (e.g. Tab-key navigation).
        """
        _token: list = [False]
        _current: list = [INPUT_BORDER]

        def _on_focus_in(_e):
            nonlocal _token
            _token[0] = True
            _token = [False]
            from_color = _current[0]

            def _up(t):
                c = self._blend_hex(from_color, INPUT_BORDER_FOC, t)
                _current[0] = c
                try:
                    widget.configure(border_color=c)
                except Exception:
                    pass

            self._tween(200, _up, ease=self._ease_out_cubic, cancel_token=_token)

        def _on_focus_out(_e):
            nonlocal _token
            _token[0] = True
            _token = [False]
            from_color = _current[0]

            def _down(t):
                c = self._blend_hex(from_color, INPUT_BORDER, t)
                _current[0] = c
                try:
                    widget.configure(border_color=c)
                except Exception:
                    pass

            self._tween(350, _down, ease=self._ease_out_cubic, cancel_token=_token)

        try:
            widget.bind("<FocusIn>",  _on_focus_in,  add="+")
            widget.bind("<FocusOut>", _on_focus_out, add="+")
        except Exception:
            pass

    # ── 8. Content Transition — status label cross-fade ──────────────────────

    def _status_transition(self, new_text: str, new_color: str) -> None:
        """Fade status label out, swap content, fade back in.

        CTkLabel.cget("text_color") can return a list of two colours
        (light/dark theme tuple) in some CustomTkinter builds.  We always
        resolve it to a plain hex string before blending.
        """
        raw = self.status_lbl.cget("text_color")
        if isinstance(raw, (list, tuple)):
            # CTk returns [light_color, dark_color]; pick the active one.
            # In dark mode the second element is used.
            current_color = str(raw[1]) if len(raw) > 1 else str(raw[0])
        else:
            current_color = str(raw)

        # Ensure it looks like a valid hex colour; fall back to BG_CARD.
        if not current_color.startswith("#") or len(current_color) not in (4, 7):
            current_color = BG_CARD

        def _fade_out_done():
            self.status_var.set(new_text)
            self._fade_widget(
                self.status_lbl,
                BG_CARD,           # 'invisible' (matches bg)
                new_color,
                duration_ms=200,
                attr="text_color",
            )

        self._fade_widget(
            self.status_lbl,
            current_color,
            BG_CARD,               # fade toward bg = invisible
            duration_ms=200,
            attr="text_color",
            on_done=_fade_out_done,
        )

    # ── 9. Connected Animation — progress bar entrance ───────────────────────
    # When generation starts, the progress bar 'connects' from 0 with a spring.

    def _progress_connected_entrance(self) -> None:
        """Spring the progress bar from 0 → 10% on generation start (connected animation)."""
        def _update(t):
            valor = 10 * t  # 0 → 10%
            try:
                self._ctk_progress.set(valor / 100.0)
                self._ctk_progress.configure(progress_color=GRAD_START)
            except Exception:
                pass

        def _finish():
            try:
                self._ctk_progress.set(0.1)
                self._ctk_progress.configure(progress_color=GRAD_START)
            except Exception:
                pass

        # Spring easing for a satisfying 'snap' entrance
        self._tween(300, _update, ease=self._ease_spring, on_done=_finish)

    # ── 10. Smooth Resize — collapsible range section ─────────────────────────
    # Animate height of _range_frame when showing/hiding by padding tween.

    def _animate_range_show(self) -> None:
        """Fade/reveal the range content in without moving the container."""
        try:
            self._range_frame.grid(row=1, column=0, sticky="ew", padx=14, pady=(0, 8))
        except Exception:
            return

        # Fade content in via text_color animation on child labels
        def _update(t):
            try:
                for child in self._range_frame.winfo_children():
                    try:
                        # Fade text widgets in
                        child.configure(text_color=self._blend_hex(BG_TILE, FG_LABEL, t))
                    except Exception:
                        pass
            except Exception:
                pass

        self._tween(220, _update, ease=self._ease_out_cubic)

    def _animate_range_hide(self, on_done=None) -> None:
        """Fade range content out then remove from grid — container stays static."""
        def _update(t):
            try:
                for child in self._range_frame.winfo_children():
                    try:
                        child.configure(text_color=self._blend_hex(FG_LABEL, BG_TILE, t))
                    except Exception:
                        pass
            except Exception:
                pass

        def _done():
            try:
                self._range_frame.grid_remove()
            except Exception:
                pass
            if on_done:
                on_done()

        self._tween(220, _update, ease=self._ease_in_cubic, on_done=_done)

    # ── 11. Elevation Animation — card border brightens on generate ───────────

    def _elevation_pulse(self, widget, duration_ms: int = 400) -> None:
        """Pulse border from dim → bright → dim (elevation glow)."""
        half = duration_ms // 2

        def _rise(t):
            c = self._blend_hex(CARD_BORDER, GRAD_START, t)
            try:
                widget.configure(border_color=c)
            except Exception:
                pass

        def _fall_start():
            def _fall(t):
                c = self._blend_hex(GRAD_START, CARD_BORDER, t)
                try:
                    widget.configure(border_color=c)
                except Exception:
                    pass
            self._tween(half, _fall, ease=self._ease_out_cubic)

        self._tween(half, _rise, ease=self._ease_out_cubic, on_done=_fall_start)

    # ── 12. Opacity Transition — fade entire card during generation ───────────

    # Token shared across dim/restore calls so they can cancel each other.
    _card_opacity_token: list = [False]

    def _card_opacity_transition(self, widget, dim: bool) -> None:
        """Dim the card border while generating (subtle depth cue).

        Reads the widget's *actual* current border color as the start point
        so the transition is smooth even when called mid-animation
        (e.g. immediately after _elevation_pulse finishes).
        """
        # Cancel any running card opacity tween.
        self._card_opacity_token[0] = True
        self._card_opacity_token = [False]

        try:
            raw = widget.cget("border_color")
            if isinstance(raw, (list, tuple)):
                from_color = str(raw[1]) if len(raw) > 1 else str(raw[0])
            else:
                from_color = str(raw)
            if not from_color.startswith("#"):
                from_color = CARD_BORDER
        except Exception:
            from_color = CARD_BORDER

        dim_color = BG_CARD_INNER   # almost invisible against the card bg
        to_color  = dim_color if dim else CARD_BORDER

        self._fade_widget(
            widget, from_color, to_color,
            duration_ms=300 if dim else 400,
            attr="border_color",
            cancel_token=self._card_opacity_token,
        )

    # ── 13. Gradient Image — soft vertical gradient for flat-color widgets ───
    # CTkFrame/CTkLabel only take a flat fg_color; this renders a real
    # gradient (used by the header accent pill) as a CTkImage instead.

    def _make_vertical_gradient_image(
        self,
        width: int,
        height: int,
        top_rgb: tuple[int, int, int],
        bottom_rgb: tuple[int, int, int],
        radius: int = 0,
    ):
        """Render a soft vertical gradient, rounded if *radius* > 0.

        Returns a ``CTkImage`` ready to pass to a ``CTkLabel``, or ``None``
        if Pillow/CustomTkinter aren't available — callers must fall back
        to a flat-color widget in that case.
        """
        if not (PIL_AVAILABLE and CTK_AVAILABLE):
            return None
        try:
            width = max(1, int(width))
            height = max(1, int(height))
            scale = 4  # supersample then downscale for smooth edges
            w, h = width * scale, height * scale
            img = Image.new("RGB", (w, h))
            px = img.load()
            for y in range(h):
                t = y / max(h - 1, 1)
                r = int(top_rgb[0] + (bottom_rgb[0] - top_rgb[0]) * t)
                g = int(top_rgb[1] + (bottom_rgb[1] - top_rgb[1]) * t)
                b = int(top_rgb[2] + (bottom_rgb[2] - top_rgb[2]) * t)
                for x in range(w):
                    px[x, y] = (r, g, b)
            if radius:
                mask = Image.new("L", (w, h), 0)
                ImageDraw.Draw(mask).rounded_rectangle(
                    (0, 0, w - 1, h - 1), radius=radius * scale, fill=255,
                )
                img.putalpha(mask)
            img = img.resize((width, height), Image.LANCZOS)
            return ctk.CTkImage(light_image=img, dark_image=img, size=(width, height))
        except Exception:
            logger.exception("Could not render gradient image")
            return None

    # ── 14. Floating Shadow — soft blurred drop-shadow behind a card ─────────
    # Tkinter has no native elevation/shadow. This renders a blurred rounded
    # rectangle as an RGBA image and places it behind *widget*, tracking its
    # size on resize (debounced so a window drag doesn't re-render every
    # pixel). Mirrors the reference dashboard's floating, shadowed widgets.

    def _make_shadow_image(self, width: int, height: int, radius: int, blur: int, alpha: int):
        if not PIL_AVAILABLE:
            return None
        try:
            width = max(1, int(width))
            height = max(1, int(height))
            pad = blur * 2
            img = Image.new("RGBA", (width + pad * 2, height + pad * 2), (0, 0, 0, 0))
            draw = ImageDraw.Draw(img)
            sr, sg, sb = int(SHADOW_COLOR[1:3], 16), int(SHADOW_COLOR[3:5], 16), int(SHADOW_COLOR[5:7], 16)
            draw.rounded_rectangle(
                (pad, pad, pad + width, pad + height),
                radius=radius, fill=(sr, sg, sb, alpha),
            )
            return img.filter(ImageFilter.GaussianBlur(blur))
        except Exception:
            logger.exception("Could not render shadow image")
            return None

    def _attach_floating_shadow(
        self,
        widget,
        parent,
        radius: int = 26,
        blur: int = 18,
        alpha: int = 130,
        offset: tuple[int, int] = (0, 10),
    ) -> None:
        """Place a soft blurred shadow behind *widget* to simulate elevation.

        Entirely best-effort: any failure (missing Pillow, unusual Tk build,
        widget not yet mapped, etc.) leaves the widget exactly as it already
        was — bordered, no shadow — rather than raising.
        """
        if not PIL_AVAILABLE:
            return

        try:
            shadow_lbl = tk.Label(parent, bg=BG_APP, bd=0, highlightthickness=0)
        except Exception:
            return

        state = {"size": (0, 0), "pending": None, "photo": None}

        def _reposition(pad: int) -> None:
            try:
                x = widget.winfo_x() - pad + offset[0]
                y = widget.winfo_y() - pad + offset[1]
                shadow_lbl.place(x=x, y=y)
                shadow_lbl.lower(widget)
                widget.lift()  # belt-and-suspenders: guarantee card stays on top
            except Exception:
                pass

        def _redraw() -> None:
            state["pending"] = None
            try:
                w = widget.winfo_width()
                h = widget.winfo_height()
                if w <= 1 or h <= 1:
                    return
                lw, lh = state["size"]
                # Skip re-rendering for tiny deltas — avoids re-blurring on
                # every intermediate frame while a window is being dragged.
                if abs(w - lw) < 6 and abs(h - lh) < 6 and state["photo"] is not None:
                    _reposition(blur * 2)
                    return
                img = self._make_shadow_image(w, h, radius, blur, alpha)
                if img is None:
                    return
                state["size"] = (w, h)
                photo = ImageTk.PhotoImage(img)
                state["photo"] = photo          # keep a reference alive
                shadow_lbl.configure(image=photo)
                _reposition(blur * 2)
            except Exception:
                pass

        def _on_configure(_event=None) -> None:
            if state["pending"] is not None:
                try:
                    self.after_cancel(state["pending"])
                except Exception:
                    pass
            state["pending"] = self.after(90, _redraw)

        try:
            widget.bind("<Configure>", _on_configure, add="+")
            parent.bind("<Configure>", _on_configure, add="+")
        except Exception:
            pass
        self.after(140, _redraw)

    # ── Fluent animation wiring — called once after interface is built ─────────

    def _wire_fluent_animations(self) -> None:
        """Attach all Fluent animations to UI widgets."""
        # 1. Window scale-in on launch
        self.after(60, self._scale_in_window)

        # 2. Entrance animations — staggered slide-up + slide-in-from-right,
        #    cascading header → divider → card → nested tiles, like the
        #    Fluent reference's sequential card-deck reveal.
        delays = [0, 50, 110, 190, 240, 290]
        entrance_targets = [
            (self._header_ref,         {"pady": (36, 0), "padx": 36}, delays[0], 14),
            (self._div_ref,            {"pady": 22, "padx": 36},      delays[1], 0),
            (self._card_ref,           {"pady": (0, 12), "padx": 32}, delays[2], 14),
            (self._range_outer_ref,    {"pady": (0, 26)},             delays[3], 10),
            (self._progress_frame_ref, {},                            delays[4], 10),
            (self._btn_row_ref,        {"pady": (24, 0)},             delays[5], 10),
        ]
        for idx, (widget, kw, delay, slide_x) in enumerate(entrance_targets):
            if widget is not None:
                self._entrance_animation(
                    widget,
                    kw,
                    delay_ms=delay,
                    slide_x_px=slide_x,
                    on_done=self._mark_startup_ready if idx == len(entrance_targets) - 1 else None,
                )

        # 3. Reveal on the main card only — no lift/reveal on containers/tiles.
        self._attach_reveal(self._card_ref)

        # 5. Press animation on button only (the button itself, not its container)
        self._attach_press_animation(self.botao_gerar)
        # Hover lift only on the button itself
        self._attach_hover_lift(self.botao_gerar, lift_px=3)
        # Hover lift on the expand/collapse button only
        self._attach_hover_lift(self._range_toggle_btn, lift_px=2)

        # 6. Acrylic focus on all entries (beyond the existing border switch)
        for entry in (self.entry_titulo, self.entry_verso_de, self.entry_verso_ate):
            self._attach_acrylic_focus(entry)
        self._attach_focus_reveal(self.scripture_text,
                                   normal_border=INPUT_BORDER,
                                   reveal_border=INPUT_BORDER_FOC,
                                   duration_ms=200)

        # 7. Floating drop-shadow under the main card — the headline
        #    "elevation" cue from the reference design.
        if self._card_ref is not None:
            self._attach_floating_shadow(self._card_ref, self._card_ref.master)

    # ══════════════════════════════════════════════════════════════════════════
    # End of Fluent Animation Engine
    # ══════════════════════════════════════════════════════════════════════════

    def __init__(self) -> None:
        if CTK_AVAILABLE:
            ctk.set_appearance_mode("dark")
            ctk.set_default_color_theme("blue")

        super().__init__()

        self.withdraw()  # hide until fully rendered

        self.title("Gerador de Leitura Bíblica")
        self.geometry("860x740")
        self.minsize(700, 640)
        self.configure(fg_color=BG_APP)
        self.resizable(True, True)

        self._apply_dark_titlebar()

        # State
        self.titulo_var            = tk.StringVar()
        self.status_var            = tk.StringVar(value="")
        self.progresso_var         = tk.DoubleVar(value=0)
        self.verso_de_var          = tk.StringVar()
        self.verso_ate_var         = tk.StringVar()
        self.pasta_saida_var       = tk.StringVar()
        self._gerando              = False
        self._config               = carregar_configuracao()
        self.pasta_saida_var.set(str(output_dir(self._config)))
        self._target_progress      = 0.0
        self._progress_animation_running = False
        self._card_opacity_token   = [False]      # instance variable for thread-safety
        self._original_window_height = None
        self._app_icon             = None
        self._ultimo_resultado: ResultadoGeracao | None = None

        self._aplicar_icone()

        # ── Animation widget refs (populated by _montar_interface) ────────────
        self._header_ref: ctk.CTkFrame | None = None
        self._div_ref:    tk.Canvas    | None = None
        self._card_ref:   ctk.CTkFrame | None = None
        self._range_outer_ref:    ctk.CTkFrame | None = None
        self._progress_frame_ref: ctk.CTkFrame | None = None
        self._btn_row_ref:        ctk.CTkFrame | None = None

        self._startup_ready = False
        self._splash_start_time = time.perf_counter()
        self._startup_splash = self._create_splash_screen()

        self._montar_interface()
        self._restaurar_geometria()
        self.protocol("WM_DELETE_WINDOW", self._fechar)
        self.bind_all("<Control-Return>", self._atalho_gerar)
        self.after(250, self._perguntar_clipboard_inicial)
        self._schedule_splash_close(self._startup_splash)

        logger.info("Window initialised | geometry=%s", self.geometry())

    # ── Dark title bar (Windows 11) ───────────────────────────────────────────

    def _apply_dark_titlebar(self) -> None:
        try:
            from ctypes import windll, c_int, byref, sizeof
            value = c_int(1)
            windll.dwmapi.DwmSetWindowAttribute(
                windll.user32.GetParent(self.winfo_id()),
                20,  # DWMWA_USE_IMMERSIVE_DARK_MODE
                byref(value),
                sizeof(value),
            )
        except Exception:
            pass

    # ── Icon ──────────────────────────────────────────────────────────────────

    def _aplicar_icone(self) -> None:
        try:
            icon = tk.PhotoImage(width=32, height=32)
            icon.put("#0a0e1a", to=(0, 0, 32, 32))
            icon.put("#1d4ed8", to=(5, 4, 27, 28))
            icon.put("#1e40af", to=(7, 6, 25, 26))
            icon.put("#eaf0ff", to=(10, 9, 15, 23))
            icon.put("#eaf0ff", to=(17, 9, 22, 23))
            icon.put("#60a5fa", to=(15, 10, 17, 24))
            icon.put("#60a5fa", to=(9, 8, 23, 10))
            icon.put("#60a5fa", to=(9, 23, 23, 25))
            self.iconphoto(True, icon)
            self._app_icon = icon
        except Exception:
            logger.exception("Could not apply app icon")

    def _create_splash_screen(self) -> tk.Toplevel:
        splash = tk.Toplevel(self)
        splash.overrideredirect(True)
        try:
            splash.attributes("-topmost", True)
        except Exception:
            pass

        splash.configure(bg=BG_APP)

        container = ctk.CTkFrame(
            splash,
            fg_color=BG_APP,
            border_width=1,
            border_color=CARD_BORDER,
        )
        container.pack(expand=True, fill="both", padx=2, pady=2)

        content = ctk.CTkFrame(container, fg_color=BG_APP)
        content.place(relx=0.5, rely=0.5, anchor="center", relwidth=0.92, relheight=0.92)

        ctk.CTkLabel(
            content,
            text=ICON_BOOK,
            font=(FONT_FAMILY, 38),
            fg_color=BG_APP,
            text_color=GRAD_START,
        ).pack(pady=(18, 6))

        ctk.CTkLabel(
            content,
            text="Gerador de Leitura Bíblica",
            font=(FONT_FAMILY, 17, "bold"),
            fg_color=BG_APP,
            text_color=FG_TITLE,
        ).pack(pady=(0, 4))

        ctk.CTkLabel(
            content,
            text="Carregando...",
            font=(FONT_FAMILY, 12),
            fg_color=BG_APP,
            text_color=FG_LABEL,
        ).pack(pady=(0, 18))

        progress = ctk.CTkProgressBar(
            content,
            mode="indeterminate",
            progress_color=GRAD_MID,
            fg_color=BG_APP,
        )
        progress.pack(fill="x", padx=36, pady=(0, 16), ipady=3)
        progress.start()

        splash.update_idletasks()
        width = 420
        height = 220
        x = (splash.winfo_screenwidth() - width) // 2
        y = (splash.winfo_screenheight() - height) // 2
        splash.geometry(f"{width}x{height}+{x}+{y}")
        splash.deiconify()

        return splash

    def _schedule_splash_close(self, splash: tk.Toplevel) -> None:
        min_visible_ms = 500

        def _hide_splash() -> None:
            try:
                splash.destroy()
            except Exception:
                pass
            self._fade_in_main_window()

        def _check_ready() -> None:
            if self._startup_ready:
                _hide_splash()
            else:
                self.after(80, _check_ready)

        def _close_splash() -> None:
            elapsed = int((time.perf_counter() - self._splash_start_time) * 1000)
            remaining = max(min_visible_ms - elapsed, 0)
            if remaining > 0:
                self.after(remaining, _check_ready)
            else:
                _check_ready()

        self.after(0, _close_splash)

    def _fade_in_main_window(self) -> None:
        try:
            self.attributes("-alpha", 0.0)
        except Exception:
            self.deiconify()
            return

        self.deiconify()

        def _update(alpha: float) -> None:
            try:
                self.attributes("-alpha", alpha)
            except Exception:
                pass

        def _finish() -> None:
            try:
                self.attributes("-alpha", 1.0)
            except Exception:
                pass

        self._tween(260, _update, ease=self._ease_out_cubic, on_done=_finish)

    def _mark_startup_ready(self) -> None:
        self._startup_ready = True

    # ── Layout ────────────────────────────────────────────────────────────────

    def _montar_interface(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        outer = ctk.CTkFrame(self, fg_color=BG_APP)
        outer.grid(row=0, column=0, sticky="nsew")
        outer.grid_columnconfigure(0, weight=1)
        outer.grid_rowconfigure(2, weight=1)

        # ── Header ────────────────────────────────────────────────────────────
        header = ctk.CTkFrame(outer, fg_color=BG_APP)
        header.grid(row=0, column=0, sticky="ew", padx=36, pady=(36, 0))
        header.grid_columnconfigure(1, weight=1)
        self._header_ref = header

        # Gradient accent pill left of title — soft vertical sky→teal gradient,
        # echoing the reference dashboard's sidebar gradient. Falls back to a
        # flat-color pill if Pillow isn't available for any reason.
        accent_pill_img = self._make_vertical_gradient_image(
            10, 54, (0x93, 0xc5, 0xfd), (0x1d, 0x4e, 0xd8), radius=5,
        )
        if accent_pill_img is not None:
            accent_pill = ctk.CTkLabel(
                header, image=accent_pill_img, text="", fg_color=BG_APP,
                width=10, height=54,
            )
            accent_pill.image = accent_pill_img  # keep a reference alive
        else:
            accent_pill = ctk.CTkFrame(
                header, fg_color=GRAD_START, corner_radius=6, width=6, height=54,
            )
            accent_pill.grid_propagate(False)
        accent_pill.grid(row=0, column=0, rowspan=2, sticky="ns", padx=(0, 14))

        title_block = ctk.CTkFrame(header, fg_color=BG_APP)
        title_block.grid(row=0, column=1, sticky="ew")
        title_block.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            title_block, text=f"{ICON_BOOK}  Gerador de Leitura Bíblica",
            fg_color=BG_APP, text_color=FG_TITLE,
            font=(FONT_FAMILY, 26, "bold"),
        ).grid(row=0, column=0, sticky="w")

        ctk.CTkLabel(
            title_block,
            text="Gera apresentações .pptx preservando fontes, cores, animações e transições do modelo.",
            fg_color=BG_APP, text_color=FG_LABEL, font=(FONT_FAMILY, 12),
        ).grid(row=1, column=0, sticky="w", pady=(5, 0))

        # ── Gradient divider bar (sky → mint → deep teal) ─────────────────────
        div_canvas = tk.Canvas(outer, height=2, bg=BG_APP, bd=0, highlightthickness=0)
        div_canvas.grid(row=1, column=0, sticky="ew", padx=36, pady=22)
        self._div_ref = div_canvas

        def _draw_divider(event=None):
            w = div_canvas.winfo_width()
            h = max(1, div_canvas.winfo_height())
            if w < 2:
                return
            # Always clear first — a prior version of this redraw never did,
            # so every <Configure> (i.e. every resize) silently piled more
            # line items onto the canvas forever. Fixed here.
            div_canvas.delete("all")

            stops = [
                (0.0,   0x93, 0xc5, 0xfd),   # blue-300
                (0.5,   0x3b, 0x82, 0xf6),   # blue-500
                (1.0,   0x1d, 0x4e, 0xd8),   # blue-700
            ]

            def _stop_color(t: float) -> tuple[int, int, int]:
                if t < 0.5:
                    s = t / 0.5
                    a, b = stops[0], stops[1]
                else:
                    s = (t - 0.5) / 0.5
                    a, b = stops[1], stops[2]
                r = int(a[1] + s * (b[1] - a[1]))
                g = int(a[2] + s * (b[2] - a[2]))
                bl = int(a[3] + s * (b[3] - a[3]))
                return r, g, bl

            if PIL_AVAILABLE:
                try:
                    img = Image.new("RGB", (w, 1))
                    px = img.load()
                    for i in range(w):
                        px[i, 0] = _stop_color(i / max(w - 1, 1))
                    img = img.resize((w, h))
                    photo = ImageTk.PhotoImage(img)
                    div_canvas.create_image(0, 0, anchor="nw", image=photo)
                    div_canvas._gradient_img = photo  # keep reference alive
                    return
                except Exception:
                    logger.exception("Could not render divider gradient via PIL")

            # Fallback (no Pillow): per-pixel line segments, same colours.
            for i in range(max(w, 1)):
                r, g, bl = _stop_color(i / max(w, 1))
                div_canvas.create_line(i, 0, i, h, fill=f"#{r:02x}{g:02x}{bl:02x}")

        div_canvas.bind("<Configure>", lambda e: _draw_divider())

        # ── Card — floating glass panel ───────────────────────────────────────
        card = ctk.CTkFrame(
            outer,
            fg_color=BG_CARD,
            corner_radius=RADIUS,
            border_width=1,
            border_color=CARD_BORDER,
        )
        card.grid(row=2, column=0, sticky="nsew", padx=32, pady=(0, 12))
        card.grid_columnconfigure(0, weight=1)
        card.grid_rowconfigure(0, weight=1)
        self._card_ref = card

        inner = ctk.CTkFrame(card, fg_color=BG_CARD, corner_radius=RADIUS - 4)
        inner.pack(fill="both", expand=True, padx=32, pady=32)
        inner.grid_columnconfigure(0, weight=1)
        inner.grid_rowconfigure(1, weight=1)

        # Scripture
        self._field_label(inner, "Texto Bíblico", row=0)

        self.scripture_text = ctk.CTkTextbox(
            inner,
            fg_color=BG_INPUT, text_color=FG_HINT,
            font=(FONT_FAMILY, 13),
            wrap="word", height=100,
            corner_radius=RADIUS_TILE,
            border_width=1,
            border_color=INPUT_BORDER,
        )
        self.scripture_text.grid(row=1, column=0, sticky="nsew", pady=(8, 26))
        self.scripture_text.insert("1.0", self._SCRIPTURE_PLACEHOLDER)
        self.scripture_text.bind("<FocusIn>",  self._limpar_placeholder_scripture)
        self.scripture_text.bind("<FocusOut>", self._restaurar_placeholder_scripture)
        # Ensure the text cursor (caret) is visible only while the field has focus.
        # Only FocusIn/FocusOut may change insertwidth — never Enter/Leave.
        def _caret_show(_e=None):
            try:
                self.scripture_text._textbox.configure(insertwidth=2, insertbackground=FG_INPUT)
            except Exception:
                pass
        def _caret_hide(_e=None):
            try:
                self.scripture_text._textbox.configure(insertwidth=0)
            except Exception:
                pass
        self.scripture_text.bind("<FocusIn>",  _caret_show, add="+")
        self.scripture_text.bind("<FocusOut>", _caret_hide, add="+")

        # Title
        self._field_label(inner, "Título da leitura", row=2)

        self.entry_titulo = self._entry(inner, self.titulo_var, inner_ipady=8)
        self.entry_titulo.grid(row=3, column=0, sticky="ew", pady=(8, 26))

        # ── Salvar em (output folder) ───────────────────────────────────────────
        self._field_label(inner, "Salvar em", row=4)

        pasta_row = ctk.CTkFrame(inner, fg_color=BG_CARD, corner_radius=0)
        pasta_row.grid(row=5, column=0, sticky="ew", pady=(8, 26))
        pasta_row.grid_columnconfigure(0, weight=1)

        self.entry_pasta_saida = self._entry(pasta_row, self.pasta_saida_var, inner_ipady=8)
        self.entry_pasta_saida.grid(row=0, column=0, sticky="ew", padx=(0, 10))

        self._botao_escolher_pasta = ctk.CTkButton(
            pasta_row,
            text=f"{ICON_FOLDER}  Procurar",
            command=self._escolher_pasta_saida,
            fg_color=BG_TILE, text_color=FG_LABEL,
            hover_color=BG_TILE,
            font=(FONT_FAMILY, 13), cursor="hand2",
            corner_radius=RADIUS_PILL,
            width=110, height=44,
        )
        self._botao_escolher_pasta.grid(row=0, column=1)

        # ── Verse range (optional) — its own floating tile ────────────────────
        range_outer = ctk.CTkFrame(
            inner, fg_color=BG_TILE, corner_radius=RADIUS_TILE,
            border_width=1, border_color=CARD_BORDER,
        )
        range_outer.grid(row=6, column=0, sticky="ew", pady=(0, 26))
        range_outer.grid_columnconfigure(0, weight=1)
        self._range_outer_ref = range_outer

        # CTk can cache a stale corner mask if _draw() runs before the first
        # child is gridded, leaving the top-left corner square while the
        # other three round correctly. Forcing one extra redraw once layout
        # has settled guarantees all four corners render identically.
        def _force_uniform_corners() -> None:
            try:
                range_outer.update_idletasks()
                range_outer._draw(no_color_updates=True)
            except Exception:
                pass

        self.after(60, _force_uniform_corners)

        # Collapsible toggle
        self._range_visible = False
        self._range_toggle_btn = ctk.CTkButton(
            range_outer,
            text="▶  Intervalo de versículos  (opcional)",
            command=self._toggle_range_section,
            fg_color=BG_TILE, text_color=FG_LABEL,
            hover_color=BG_TILE,
            font=(FONT_FAMILY, 12, "bold"), cursor="hand2", anchor="w",
            corner_radius=RADIUS_TILE,
        )
        # sticky="ew" + column weight ensures the button fills the container so
        # all four corners are clipped by the parent's rounded border — without
        # this the left side is left square because the button doesn't reach
        # the edges of the rounded frame.
        self._range_toggle_btn.grid(row=0, column=0, sticky="ew")

        self._range_frame = ctk.CTkFrame(range_outer, fg_color=BG_TILE, corner_radius=RADIUS_TILE)
        self._range_frame.grid_columnconfigure(1, weight=0)

        ctk.CTkLabel(
            self._range_frame, text="De:", fg_color=BG_TILE, text_color=FG_LABEL,
            font=(FONT_FAMILY, 12),
        ).grid(row=0, column=0, sticky="w", padx=(0, 8), pady=(6, 0))

        self.entry_verso_de = self._entry(self._range_frame, self.verso_de_var, inner_ipady=4)
        self.entry_verso_de.configure(width=80, height=32)
        self.entry_verso_de.grid(row=0, column=1, sticky="w", pady=(6, 0))

        ctk.CTkLabel(
            self._range_frame, text="Até:", fg_color=BG_TILE, text_color=FG_LABEL,
            font=(FONT_FAMILY, 12),
        ).grid(row=0, column=2, sticky="w", padx=(24, 8), pady=(6, 0))

        self.entry_verso_ate = self._entry(self._range_frame, self.verso_ate_var, inner_ipady=4)
        self.entry_verso_ate.configure(width=80, height=32)
        self.entry_verso_ate.grid(row=0, column=3, sticky="w", pady=(6, 0))

        ctk.CTkLabel(
            self._range_frame,
            text="Deixe em branco para gerar todos os versículos.",
            fg_color=BG_TILE, text_color=FG_HINT,
            font=(FONT_FAMILY, 11),
        ).grid(row=1, column=0, columnspan=4, sticky="w", pady=(4, 4))

        # Status
        self.status_lbl = ctk.CTkLabel(
            inner, textvariable=self.status_var,
            fg_color=BG_CARD, text_color=FG_HINT,
            font=(FONT_FAMILY, 12), anchor="w",
            wraplength=680, justify="left",
        )
        self.status_lbl.grid(row=7, column=0, sticky="ew", pady=(0, 12))

        # Progress bar — its own floating tile
        self._progress_frame(inner, row=8)

        # Buttons — its own floating tile
        # btn_row must NOT stretch — it wraps tightly around the button.
        # Place it right-aligned inside the card with explicit right padding.
        btn_row = ctk.CTkFrame(
            inner, fg_color=BG_TILE, corner_radius=RADIUS_TILE,
            border_width=1, border_color=CARD_BORDER,
        )
        btn_row.grid(row=9, column=0, sticky="e", pady=(24, 0), padx=(0, 0))
        self._btn_row_ref = btn_row

        self.botao_gerar = self._primary_button(
            btn_row, f"  {ICON_PLAY}  Gerar PowerPoint  (Ctrl+Enter)", self.gerar,
        )
        self.botao_gerar.grid(row=0, column=0, pady=10, padx=10)

        # Footer
        self._footer_lbl = ctk.CTkLabel(
            outer,
            text=self._footer_text(),
            fg_color=BG_APP, text_color="#5a649a",
            font=(FONT_FAMILY, 11),
        )
        self._footer_lbl.grid(row=3, column=0, pady=(2, 6))

        # ── Wire all Fluent animations ────────────────────────────────────────
        self.after(10, self._wire_fluent_animations)

    def _footer_text(self) -> str:
        return f"© {datetime.now().year}  |  {APP_NAME} v{APP_VERSION}  |  {APP_AUTHOR}"

    # ── Widget factories ──────────────────────────────────────────────────────

    def _field_label(self, parent, text: str, row: int) -> ctk.CTkLabel:
        lbl = ctk.CTkLabel(
            parent, text=text.upper(), fg_color=BG_CARD, text_color=GRAD_START,
            font=(FONT_FAMILY, 11, "bold"), anchor="w",
        )
        lbl.grid(row=row, column=0, sticky="w")
        return lbl

    def _entry(self, parent, var: tk.Variable, inner_ipady: int = 6) -> ctk.CTkEntry:
        entry = ctk.CTkEntry(
            parent, textvariable=var,
            fg_color=BG_INPUT, text_color=FG_INPUT,
            border_color=INPUT_BORDER,
            border_width=1,
            font=(FONT_FAMILY, 13),
            corner_radius=RADIUS_PILL,
            height=44,
        )

        def _on_focus_in(_e):
            entry.configure(border_color=INPUT_BORDER_FOC)
        def _on_focus_out(_e):
            entry.configure(border_color=INPUT_BORDER)

        entry.bind("<FocusIn>",  _on_focus_in)
        entry.bind("<FocusOut>", _on_focus_out)

        return entry

    def _primary_button(self, parent, text: str, command) -> ctk.CTkButton:
        btn = ctk.CTkButton(
            parent, text=text, command=command,
            fg_color=GRAD_START, text_color=BG_APP,
            hover_color=GRAD_MID,
            font=(FONT_FAMILY, 14, "bold"), cursor="hand2",
            corner_radius=RADIUS_PILL,
            height=50,
        )
        return btn

    def _progress_frame(self, parent, row: int) -> None:
        frame = ctk.CTkFrame(
            parent, fg_color=BG_TILE, corner_radius=RADIUS_TILE,
            border_width=1, border_color=CARD_BORDER,
        )
        frame.grid(row=row, column=0, sticky="ew")
        frame.grid_columnconfigure(0, weight=1)
        self._progress_frame_ref = frame

        self._ctk_progress = ctk.CTkProgressBar(
            frame,
            fg_color=BG_PROG_TRACK,
            progress_color=GRAD_START,
            corner_radius=8,
            height=12,
            mode="determinate",
        )
        self._ctk_progress.set(0)
        # Idle state: no filled segment should be visible at all.
        self._ctk_progress.configure(progress_color=BG_PROG_TRACK)
        self._ctk_progress.grid(row=0, column=0, sticky="ew", padx=14, pady=14)
        # Disable any hover/pointer cursor — progress bar is display-only
        try:
            self._ctk_progress.configure(cursor="")
        except Exception:
            pass
        # Keep CTkProgressBar in sync with progresso_var (0–100 → 0.0–1.0)
        self.progresso_var.trace_add("write", self._sync_ctk_progress)

    def _sync_ctk_progress(self, *_) -> None:
        try:
            valor = self.progresso_var.get()
            self._ctk_progress.set(valor / 100.0)
            # At idle (0%) the rounded-corner fill cap can still render a
            # small visible sliver at the left edge even at value 0.
            # Match the fill color to the track so nothing is visible
            # until generation actually starts.
            # At 100%, also hide — smooth completion.
            if valor <= 0 or valor >= 99.5:
                self._ctk_progress.configure(progress_color=BG_PROG_TRACK)
            else:
                self._ctk_progress.configure(progress_color=GRAD_START)
        except Exception:
            pass

    def _redraw_progress(self, event=None) -> None:
        try:
            valor = self.progresso_var.get()
            self._ctk_progress.set(valor / 100.0)
            if valor <= 0 or valor >= 99.5:
                self._ctk_progress.configure(progress_color=BG_PROG_TRACK)
            else:
                self._ctk_progress.configure(progress_color=GRAD_START)
        except Exception:
            pass

    # ── Progress animation ────────────────────────────────────────────────────

    def _set_progresso_imediato(self, valor: float) -> None:
        self.progresso_var.set(valor)
        self._target_progress = float(valor)
        self._redraw_progress()
        self.update_idletasks()

    def _set_progresso(self, valor: float) -> None:
        self._target_progress = max(0.0, min(100.0, float(valor)))
        if not self._progress_animation_running:
            self._progress_animation_running = True
            self._animar_progresso()

    def _animar_progresso(self) -> None:
        # Always read _target_progress fresh — it may have been updated by
        # the worker thread between ticks, so we should track the latest goal.
        atual    = float(self.progresso_var.get())
        destino  = self._target_progress   # re-read every tick
        diferenca = destino - atual

        if abs(diferenca) < 0.3:
            self.progresso_var.set(destino)
            self._redraw_progress()
            self._progress_animation_running = False
            return

        passo = max(0.5, abs(diferenca) * 0.18)
        novo = atual + (passo if diferenca > 0 else -passo)
        self.progresso_var.set(novo)
        self._redraw_progress()
        self.after(16, self._animar_progresso)   # ~60 fps

    # ── Status helpers ────────────────────────────────────────────────────────

    def _set_status(self, msg: str, color: str = FG_HINT) -> None:
        try:
            self._status_transition(msg, color)
        except Exception:
            # Fallback if animation engine not ready yet
            self.status_var.set(msg)
            self.status_lbl.configure(text_color=color)

    def _set_status_gerando(self, msg: str) -> None:
        self._set_status(msg, FG_LABEL)

    def _set_status_ok(self, msg: str) -> None:
        self._set_status(f"{ICON_CHECK}  {msg}", FG_SUCCESS)

    def _set_status_erro(self, msg: str) -> None:
        self._set_status(f"{ICON_CROSS}  {msg}", FG_ERROR)

    # ── Progress reporting ────────────────────────────────────────────────────

    _PROGRESS_MESSAGES: dict[int, str] = {
        10:  "Iniciando o PowerPoint…",
        20:  "Abrindo o modelo…",
        25:  "Preenchendo slide de título…",
        90:  "Salvando apresentação…",
        100: "Concluído!",
    }

    def atualizar_progresso(self, valor: int) -> None:
        """Called from the worker thread — schedules UI update on the main thread."""
        self.after(0, self._atualizar_progresso_ui, valor)

    def _atualizar_progresso_ui(self, valor: int) -> None:
        if valor in self._PROGRESS_MESSAGES:
            self._set_status_gerando(self._PROGRESS_MESSAGES[valor])
        elif 25 < valor < 90:
            # Compute how many verses have been processed.
            pct = int(((valor - 25) / 65) * 100)
            self._set_status_gerando(f"Gerando versículos… {pct}%")
        self._set_progresso(valor)

    # ── Output folder ("Salvar em") ─────────────────────────────────────────────

    def _escolher_pasta_saida(self) -> None:
        """Open a folder picker, store the choice, and persist it to
        config.json immediately so it's restored automatically next time
        and never asked for again."""
        pasta_atual = self.pasta_saida_var.get().strip() or str(output_dir(self._config))
        escolhida = filedialog.askdirectory(
            initialdir=pasta_atual,
            title="Selecionar pasta de destino",
            parent=self,
        )
        if not escolhida:
            return  # user cancelled — keep current folder
        self.pasta_saida_var.set(escolhida)
        self._config["output_dir"] = escolhida
        salvar_configuracao(self._config)
        logger.info("Output folder updated | output_dir=%s", escolhida)
        try:
            self._footer_lbl.configure(text=self._footer_text())
        except Exception:
            pass

    # ── Geometry persistence ──────────────────────────────────────────────────

    def _restaurar_geometria(self) -> None:
        geometria = self._config.get("geometry", "")
        if re.fullmatch(r"\d+x\d+[+-]\d+[+-]\d+", geometria):
            self.geometry(geometria)
            return
        # Centre on screen.
        self.update_idletasks()
        w = self.winfo_width()
        h = self.winfo_height()
        x = max((self.winfo_screenwidth()  - w) // 2, 0)
        y = max((self.winfo_screenheight() - h) // 2, 0)
        self.geometry(f"{w}x{h}+{x}+{y}")

    # ── Window close ─────────────────────────────────────────────────────────

    def _fechar(self) -> None:
        if self._gerando:
            return
        self._config["geometry"] = self.geometry()
        salvar_configuracao(self._config)
        logger.info("Window closed | geometry=%s", self._config["geometry"])
        self.destroy()

    # ── Keyboard shortcut ─────────────────────────────────────────────────────

    def _atalho_gerar(self, _event=None) -> str:
        self.gerar()
        return "break"

    # ── Clipboard auto-paste ──────────────────────────────────────────────────

    def _perguntar_clipboard_inicial(self) -> None:
        if self._gerando or not self._scripture_tem_placeholder():
            return
        try:
            texto = self.clipboard_get().strip()
        except Exception:
            return
        if not texto:
            return
        try:
            leitura = ler_texto_entrada(texto)
        except GeradorSalmosError:
            return
        if messagebox.askyesno(
            "Texto bíblico detectado",
            f"Foi detectado um texto bíblico na área de transferência:\n\n"
            f"{leitura.livro_capitulo} — {len(leitura.versiculos)} versículos\n\n"
            "Deseja colar automaticamente?",
        ):
            self.scripture_text.configure(state="normal", text_color=FG_INPUT)
            self.scripture_text.delete("1.0", "end")
            self.scripture_text.insert("1.0", texto)
            self.scripture_text.focus_set()
            logger.info("Scripture pasted from clipboard | %s", leitura.livro_capitulo)

    # ── UI state helpers ──────────────────────────────────────────────────────

    def _set_ui_gerando(self, gerando: bool) -> None:
        estado = "disabled" if gerando else "normal"
        self.scripture_text.configure(state=estado)
        self.entry_titulo.configure(state=estado)
        self.entry_verso_de.configure(state=estado)
        self.entry_verso_ate.configure(state=estado)
        self.botao_gerar.configure(
            state=estado,
            fg_color="#163060" if gerando else GRAD_START,
        )
        # Fluent: elevation pulse on card when generation starts/ends
        if self._card_ref is not None:
            if gerando:
                self._elevation_pulse(self._card_ref, duration_ms=400)
                self._progress_connected_entrance()
            else:
                self._card_opacity_transition(self._card_ref, dim=False)

    def _limpar_inputs_apos_sucesso(self) -> None:
        self.scripture_text.configure(state="normal", text_color=FG_HINT)
        self.scripture_text.delete("1.0", "end")
        self.scripture_text.insert("1.0", self._SCRIPTURE_PLACEHOLDER)
        self.titulo_var.set("")
        self.verso_de_var.set("")
        self.verso_ate_var.set("")
        self.scripture_text.focus_set()

    # ── Verse range toggle ────────────────────────────────────────────────────

    def _toggle_range_section(self) -> None:
        self._range_visible = not self._range_visible
        if self._range_visible:
            self._range_toggle_btn.configure(
                text="▼  Intervalo de versículos  (opcional)"
            )
            self._grow_window_for_range()
            self._animate_range_show()
        else:
            self._range_toggle_btn.configure(
                text="▶  Intervalo de versículos  (opcional)"
            )
            self._animate_range_hide()
            self._shrink_window_for_range()

    def _grow_window_for_range(self) -> None:
        """Grow the window to fit the range panel, anchored to the
        original (pre-expansion) height.

        The Scripture textbox is the only row in the card with grid
        weight. Without extra window height, expanding the range panel
        would force the grid to shrink that weighted row to make room —
        which is what made Scripture disappear. Growing the window keeps
        every widget, including Scripture, fully visible.

        The target height is always computed from the saved original
        height plus the panel's required height, never from the window's
        *current* height — so repeated expand/collapse cycles can never
        accumulate extra height.

        This only applies in windowed (normal) mode. When the window is
        maximized there is already room for the panel, so forcing a
        geometry change there is what made the window balloon to an
        oversized height — skip it entirely in that case.
        """
        try:
            if self.state() != "normal":
                return
        except Exception:
            pass
        try:
            self.update_idletasks()
            if self._original_window_height is None:
                self._original_window_height = self.winfo_height()
            needed = self._range_frame.winfo_reqheight() + 4
        except Exception:
            return
        try:
            w = self.winfo_width()
            x, y = self.winfo_x(), self.winfo_y()
            target_h = self._original_window_height + max(0, needed)
            if target_h > self.winfo_height():
                self.geometry(f"{w}x{target_h}+{x}+{y}")
                self.update_idletasks()
        except Exception:
            pass

    def _shrink_window_for_range(self) -> None:
        """Reverse of `_grow_window_for_range` — restores the window to
        its saved original height exactly, regardless of how many
        expand/collapse cycles have happened, so height never drifts.

        Skipped when maximized, mirroring `_grow_window_for_range`,
        since no growth happens in that state either."""
        if self._original_window_height is None:
            return
        try:
            if self.state() != "normal":
                return
        except Exception:
            pass
        try:
            w = self.winfo_width()
            x, y = self.winfo_x(), self.winfo_y()
            min_h = self.minsize()[1]
            new_h = max(min_h, self._original_window_height)
            self.geometry(f"{w}x{new_h}+{x}+{y}")
            self.update_idletasks()
        except Exception:
            pass

    # ── Scripture field placeholder ───────────────────────────────────────────

    def _limpar_placeholder_scripture(self, _event=None) -> None:
        if self._scripture_tem_placeholder():
            self.scripture_text.delete("1.0", "end")
            self.scripture_text.configure(text_color=FG_INPUT)

    def _restaurar_placeholder_scripture(self, _event=None) -> None:
        if not self.scripture_text.get("1.0", "end").strip():
            self.scripture_text.configure(text_color=FG_HINT)
            self.scripture_text.delete("1.0", "end")
            self.scripture_text.insert("1.0", self._SCRIPTURE_PLACEHOLDER)

    def _scripture_tem_placeholder(self) -> bool:
        return self.scripture_text.get("1.0", "end").strip() == self._SCRIPTURE_PLACEHOLDER

    def _obter_texto_scripture(self) -> str:
        if self._scripture_tem_placeholder():
            return ""
        return self.scripture_text.get("1.0", "end").strip()

    # ── Main action: generate ─────────────────────────────────────────────────

    def gerar(self) -> None:
        if self._gerando:
            return

        # ── Validate scripture ────────────────────────────────────────────────
        texto_leitura = self._obter_texto_scripture()
        if not texto_leitura:
            self._set_status_erro("Cole o texto bíblico antes de gerar.")
            self.scripture_text.focus_set()
            return

        try:
            leitura = ler_texto_entrada(texto_leitura)
        except GeradorSalmosError as exc:
            logger.warning("Validation failed — invalid scripture: %s", exc)
            self._set_progresso_imediato(0)
            self._set_status_erro("Texto bíblico incompleto ou com formato inválido.")
            messagebox.showerror(
                "Texto inválido",
                str(exc),
                parent=self,
            )
            self.scripture_text.focus_set()
            return

        # ── Validate title ────────────────────────────────────────────────────
        titulo_leitura = self.titulo_var.get().strip()
        if not titulo_leitura:
            logger.warning("Validation failed — missing title")
            self._set_progresso_imediato(0)
            self._set_status_erro("Informe o título da leitura.")
            messagebox.showerror(
                "Título obrigatório",
                "Digite o título da leitura antes de gerar a apresentação.\n\n"
                "Exemplo: Salmo da Misericórdia",
                parent=self,
            )
            self.entry_titulo.focus_set()
            return

        # ── Validate verse range ──────────────────────────────────────────────
        total_versiculos = len(leitura.versiculos)
        verso_de: int | None = None
        verso_ate: int | None = None

        raw_de  = self.verso_de_var.get().strip()
        raw_ate = self.verso_ate_var.get().strip()

        if raw_de or raw_ate:
            # Parse "De"
            if raw_de:
                if not raw_de.isdigit():
                    self._set_status_erro("O campo 'De' deve ser um número inteiro.")
                    messagebox.showerror(
                        "Intervalo inválido",
                        f"O campo 'De' contém um valor inválido: \"{raw_de}\"\n\n"
                        "Digite apenas um número inteiro positivo.",
                        parent=self,
                    )
                    self.entry_verso_de.focus_set()
                    return
                verso_de = int(raw_de)
                if verso_de < 1:
                    self._set_status_erro("O campo 'De' deve ser pelo menos 1.")
                    messagebox.showerror(
                        "Intervalo inválido",
                        "O versículo inicial ('De') deve ser pelo menos 1.",
                        parent=self,
                    )
                    self.entry_verso_de.focus_set()
                    return
                if verso_de > total_versiculos:
                    self._set_status_erro(
                        f"O campo 'De' ({verso_de}) excede o total de versículos ({total_versiculos})."
                    )
                    messagebox.showerror(
                        "Intervalo inválido",
                        f"O versículo inicial ({verso_de}) é maior que o total de\n"
                        f"versículos no texto ({total_versiculos}).",
                        parent=self,
                    )
                    self.entry_verso_de.focus_set()
                    return

            # Parse "Até"
            if raw_ate:
                if not raw_ate.isdigit():
                    self._set_status_erro("O campo 'Até' deve ser um número inteiro.")
                    messagebox.showerror(
                        "Intervalo inválido",
                        f"O campo 'Até' contém um valor inválido: \"{raw_ate}\"\n\n"
                        "Digite apenas um número inteiro positivo.",
                        parent=self,
                    )
                    self.entry_verso_ate.focus_set()
                    return
                verso_ate = int(raw_ate)
                if verso_ate < 1:
                    self._set_status_erro("O campo 'Até' deve ser pelo menos 1.")
                    messagebox.showerror(
                        "Intervalo inválido",
                        "O versículo final ('Até') deve ser pelo menos 1.",
                        parent=self,
                    )
                    self.entry_verso_ate.focus_set()
                    return
                if verso_ate > total_versiculos:
                    self._set_status_erro(
                        f"O campo 'Até' ({verso_ate}) excede o total de versículos ({total_versiculos})."
                    )
                    messagebox.showerror(
                        "Intervalo inválido",
                        f"O versículo final ({verso_ate}) é maior que o total de\n"
                        f"versículos no texto ({total_versiculos}).\n\n"
                        f"O texto tem {total_versiculos} versículo(s).",
                        parent=self,
                    )
                    self.entry_verso_ate.focus_set()
                    return

            # Cross-field check
            ef_de  = verso_de  if verso_de  is not None else 1
            ef_ate = verso_ate if verso_ate is not None else total_versiculos
            if ef_de > ef_ate:
                self._set_status_erro(
                    f"Intervalo inválido: 'De' ({ef_de}) é maior que 'Até' ({ef_ate})."
                )
                messagebox.showerror(
                    "Intervalo inválido",
                    f"O versículo inicial ({ef_de}) não pode ser maior\n"
                    f"que o versículo final ({ef_ate}).",
                    parent=self,
                )
                self.entry_verso_de.focus_set()
                return

        # ── Confirm overwrite ─────────────────────────────────────────────────
        caminho_saida = caminho_saida_para(leitura, self._config)
        substituir_existente = False
        if caminho_saida.exists():
            substituir_existente = messagebox.askyesno(
                "Substituir apresentação?",
                f"Já existe uma apresentação com este nome:\n\n"
                f"    {caminho_saida.name}\n\n"
                "Deseja substituí-la?",
                icon="warning",
                parent=self,
            )
            if not substituir_existente:
                logger.info("Generation cancelled — user kept existing file: %s", caminho_saida.name)
                self._set_status("Geração cancelada. Arquivo existente mantido.", FG_HINT)
                return

        # ── Start generation ──────────────────────────────────────────────────
        logger.info(
            "Generation requested | %s | verses=%d | title=%r | range=[%s..%s]",
            leitura.livro_capitulo,
            len(leitura.versiculos),
            titulo_leitura,
            verso_de or "início",
            verso_ate or "fim",
        )
        self._gerando = True
        self._set_ui_gerando(True)
        self._set_progresso_imediato(0)
        self._set_status_gerando("Iniciando geração…")
        self.update_idletasks()

        def _worker() -> None:
            try:
                resultado = gerar_powerpoint(
                    texto_leitura,
                    titulo_leitura,
                    self.atualizar_progresso,
                    substituir_existente=substituir_existente,
                    config=self._config,
                    verso_inicio=verso_de,
                    verso_fim=verso_ate,
                )
            except GeradorSalmosError as exc:
                self.after(0, self._on_erro, str(exc))
            except Exception:
                logger.exception("Unexpected error in generation worker")
                self.after(0, self._on_erro_inesperado)
            else:
                self.after(0, self._on_sucesso, resultado)

        threading.Thread(target=_worker, daemon=True, name="pptx-worker").start()

    # ── Generation callbacks ──────────────────────────────────────────────────

    def _on_sucesso(self, resultado: ResultadoGeracao) -> None:
        self._ultimo_resultado = resultado
        caminho = resultado.caminho
        verses  = resultado.quantidade_versiculos
        secs    = resultado.duracao_segundos

        self._set_progresso(100)
        self._set_status_ok(f"{caminho.name} — {verses} versículos — {secs:.1f}s")
        self._gerando = False
        self._set_ui_gerando(False)

        # Open Explorer with the file pre-selected.
        try:
            subprocess.Popen(f'explorer.exe /select,"{caminho}"')
            logger.info("Explorer opened at %s", caminho)
        except Exception:
            logger.exception("Could not open Explorer")

        messagebox.showinfo(
            "Apresentação gerada!",
            f"Arquivo:      {caminho.name}\n"
            f"Versículos:   {verses}\n"
            f"Tempo:         {secs:.1f} segundos\n"
            f"Pasta:          {caminho.parent}",
            parent=self,
        )
        self._limpar_inputs_apos_sucesso()

    def _on_erro(self, mensagem: str) -> None:
        self._set_progresso_imediato(0)
        self._set_status_erro("Não foi possível gerar a apresentação.")
        self._gerando = False
        self._set_ui_gerando(False)
        messagebox.showerror(
            "Erro ao gerar PowerPoint",
            mensagem,
            parent=self,
        )
        self.scripture_text.focus_set()

    def _on_erro_inesperado(self) -> None:
        self._set_progresso_imediato(0)
        self._set_status_erro("Ocorreu um erro inesperado.")
        self._gerando = False
        self._set_ui_gerando(False)
        messagebox.showerror(
            "Erro inesperado",
            "Ocorreu um erro inesperado ao gerar a apresentação.\n\n"
            f"Os detalhes foram gravados em:\n{LOG_FILE}",
            parent=self,
        )
        self.scripture_text.focus_set()


# ── Entry point ────────────────────────────────────────────────────────────────

def main() -> int:
    if not CTK_AVAILABLE:
        import warnings
        warnings.warn(
            "CustomTkinter não encontrado. Instale com: pip install customtkinter",
            stacklevel=1,
        )
    app = Aplicacao()
    app.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())