from __future__ import annotations

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
from tkinter import messagebox

try:
    import pythoncom
    import pywintypes
    import win32com.client
except ImportError:
    pythoncom = None
    pywintypes = None
    win32com = None


# ── Application identity ───────────────────────────────────────────────────────
APP_NAME    = "GeradorSalmos"
APP_VERSION = "2.0"
MODEL_NAME  = "modelo.pptx"


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
BG_APP        = "#0b0d14"          # deeper, richer background
BG_CARD       = "#141720"          # glass card base — very dark navy
BG_CARD_INNER = "#0f1219"          # inset panel (slightly darker than card)
BG_INPUT      = "#1c2035"          # input background, muted navy
BG_BTN        = "#2563eb"
BG_BTN_HOVER  = "#1d4ed8"
BG_PROG_TRACK = "#1c2035"
BG_PROG_FILL  = "#3b82f6"

# Glass / acrylic accents
CARD_BORDER      = "#252d45"       # subtle border for floating effect
CARD_BORDER_TOP  = "#2e3a55"       # brighter top edge (light catch)
INPUT_BORDER     = "#252d45"
INPUT_BORDER_FOC = "#2563eb"

FG_TITLE   = "#f0f4ff"
FG_LABEL   = "#9baac4"
FG_INPUT   = "#dde4f0"
FG_HINT    = "#4a5878"
FG_SUCCESS = "#34d399"
FG_ERROR   = "#f87171"
FG_WARNING = "#fbbf24"
FG_WHITE   = "#ffffff"

FONT_FAMILY = "Segoe UI"
RADIUS      = 14                   # larger radius → rounder, softer cards

# Icons (Unicode)
ICON_BOOK  = "📖"
ICON_PLAY  = "▶"
ICON_CHECK = "✓"
ICON_CROSS = "✗"
ICON_WARN  = "⚠"


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
    _SCRIPTURE_PLACEHOLDER = "Cole aqui o texto copiado do Busca Bíblica."

    def __init__(self) -> None:
        if CTK_AVAILABLE:
            ctk.set_appearance_mode("dark")
            ctk.set_default_color_theme("blue")

        super().__init__()

        self.title("Gerador de Leitura Bíblica")
        self.geometry("780x680")
        self.minsize(660, 600)
        self.configure(bg=BG_APP)
        self.resizable(True, True)

        self._apply_dark_titlebar()

        # State
        self.titulo_var            = tk.StringVar()
        self.status_var            = tk.StringVar(value="")
        self.progresso_var         = tk.DoubleVar(value=0)
        self.verso_de_var          = tk.StringVar()
        self.verso_ate_var         = tk.StringVar()
        self._gerando              = False
        self._config               = carregar_configuracao()
        self._target_progress      = 0.0
        self._progress_animation_running = False
        self._app_icon             = None
        self._ultimo_resultado: ResultadoGeracao | None = None

        self._aplicar_icone()
        self._montar_interface()
        self._restaurar_geometria()
        self.protocol("WM_DELETE_WINDOW", self._fechar)
        self.bind_all("<Control-Return>", self._atalho_gerar)
        self.after(250, self._perguntar_clipboard_inicial)

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
            icon.put("#0f1117", to=(0, 0, 32, 32))
            icon.put("#2563eb", to=(5, 4, 27, 28))
            icon.put("#1d4ed8", to=(7, 6, 25, 26))
            icon.put("#f0f4ff", to=(10, 9, 15, 23))
            icon.put("#f0f4ff", to=(17, 9, 22, 23))
            icon.put("#93c5fd", to=(15, 10, 17, 24))
            icon.put("#93c5fd", to=(9, 8, 23, 10))
            icon.put("#93c5fd", to=(9, 23, 23, 25))
            self.iconphoto(True, icon)
            self._app_icon = icon
        except Exception:
            logger.exception("Could not apply app icon")

    # ── Layout ────────────────────────────────────────────────────────────────

    def _montar_interface(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        outer = tk.Frame(self, bg=BG_APP)
        outer.grid(row=0, column=0, sticky="nsew")
        outer.grid_columnconfigure(0, weight=1)
        outer.grid_rowconfigure(2, weight=1)

        # ── Header ────────────────────────────────────────────────────────────
        header = tk.Frame(outer, bg=BG_APP)
        header.grid(row=0, column=0, sticky="ew", padx=36, pady=(32, 0))
        header.grid_columnconfigure(0, weight=1)

        tk.Label(
            header, text=ICON_BOOK, bg=BG_APP,
            font=(FONT_FAMILY, 30), fg="#93c5fd",
        ).grid(row=0, column=0, sticky="w")

        tk.Label(
            header, text="Gerador de Leitura Bíblica", bg=BG_APP,
            fg=FG_TITLE, font=(FONT_FAMILY, 22, "bold"),
        ).grid(row=1, column=0, sticky="w", pady=(2, 0))

        tk.Label(
            header,
            text="Gera apresentações .pptx preservando fontes, cores, animações e transições do modelo.",
            bg=BG_APP, fg=FG_HINT, font=(FONT_FAMILY, 10),
        ).grid(row=2, column=0, sticky="w", pady=(4, 0))

        # ── Divider ───────────────────────────────────────────────────────────
        tk.Frame(outer, bg=CARD_BORDER, height=1).grid(
            row=1, column=0, sticky="ew", padx=36, pady=24
        )

        # ── Card — floating glass panel ───────────────────────────────────────
        # Outer border frame gives the "raised card" illusion on dark bg
        card_border = tk.Frame(outer, bg=CARD_BORDER_TOP, bd=0)
        card_border.grid(row=2, column=0, sticky="nsew", padx=32, pady=(0, 12))
        card_border.grid_columnconfigure(0, weight=1)
        card_border.grid_rowconfigure(0, weight=1)

        card = tk.Frame(card_border, bg=BG_CARD, bd=0)
        card.grid(row=0, column=0, sticky="nsew", padx=1, pady=1)
        card.grid_columnconfigure(0, weight=1)
        card.grid_rowconfigure(0, weight=1)

        inner = tk.Frame(card, bg=BG_CARD)
        inner.pack(fill="both", expand=True, padx=32, pady=28)
        inner.grid_columnconfigure(0, weight=1)
        inner.grid_rowconfigure(1, weight=1)

        # Scripture
        self._field_label(inner, "Scripture", row=0)

        scripture_frame = tk.Frame(inner, bg=CARD_BORDER, bd=0)
        scripture_frame.grid(row=1, column=0, sticky="nsew", pady=(8, 22))
        scripture_frame.grid_columnconfigure(0, weight=1)
        scripture_frame.grid_rowconfigure(0, weight=1)

        self.scripture_text = tk.Text(
            scripture_frame,
            bg=BG_INPUT, fg=FG_HINT,
            insertbackground=FG_INPUT,
            relief="flat", font=(FONT_FAMILY, 10),
            bd=0, highlightthickness=0,
            wrap="word", height=12, padx=14, pady=12,
        )
        self.scripture_text.grid(row=0, column=0, sticky="nsew", padx=1, pady=1)
        self.scripture_text.insert("1.0", self._SCRIPTURE_PLACEHOLDER)
        self.scripture_text.bind("<FocusIn>",  self._limpar_placeholder_scripture)
        self.scripture_text.bind("<FocusOut>", self._restaurar_placeholder_scripture)

        scrollbar = tk.Scrollbar(scripture_frame, command=self.scripture_text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.scripture_text.configure(yscrollcommand=scrollbar.set)

        # Title
        self._field_label(inner, "Título da leitura", row=2)

        self.entry_titulo = self._entry(inner, self.titulo_var, inner_ipady=8)
        self.entry_titulo.grid(row=3, column=0, sticky="ew", pady=(8, 22))

        # ── Verse range (optional) ────────────────────────────────────────────
        range_outer = tk.Frame(inner, bg=BG_CARD)
        range_outer.grid(row=4, column=0, sticky="ew", pady=(0, 22))
        range_outer.grid_columnconfigure(0, weight=1)

        # Collapsible toggle
        self._range_visible = False
        self._range_toggle_btn = tk.Button(
            range_outer,
            text="▶  Intervalo de versículos  (opcional)",
            command=self._toggle_range_section,
            bg=BG_CARD, fg=FG_LABEL,
            activebackground=BG_CARD, activeforeground=FG_TITLE,
            relief="flat", bd=0, padx=0, pady=0,
            font=(FONT_FAMILY, 9, "bold"), cursor="hand2", anchor="w",
        )
        self._range_toggle_btn.grid(row=0, column=0, sticky="w")

        self._range_frame = tk.Frame(range_outer, bg=BG_CARD)
        self._range_frame.grid_columnconfigure(1, weight=0)

        tk.Label(
            self._range_frame, text="De:", bg=BG_CARD, fg=FG_LABEL,
            font=(FONT_FAMILY, 9),
        ).grid(row=0, column=0, sticky="w", padx=(0, 8), pady=(12, 0))

        self.entry_verso_de = self._entry(self._range_frame, self.verso_de_var, inner_ipady=6)
        self.entry_verso_de.configure(width=8)
        self.entry_verso_de.grid(row=0, column=1, sticky="w", pady=(12, 0))

        tk.Label(
            self._range_frame, text="Até:", bg=BG_CARD, fg=FG_LABEL,
            font=(FONT_FAMILY, 9),
        ).grid(row=0, column=2, sticky="w", padx=(24, 8), pady=(12, 0))

        self.entry_verso_ate = self._entry(self._range_frame, self.verso_ate_var, inner_ipady=6)
        self.entry_verso_ate.configure(width=8)
        self.entry_verso_ate.grid(row=0, column=3, sticky="w", pady=(12, 0))

        tk.Label(
            self._range_frame,
            text="Deixe em branco para gerar todos os versículos.",
            bg=BG_CARD, fg=FG_HINT,
            font=(FONT_FAMILY, 8),
        ).grid(row=1, column=0, columnspan=4, sticky="w", pady=(6, 0))

        # Status
        self.status_lbl = tk.Label(
            inner, textvariable=self.status_var,
            bg=BG_CARD, fg=FG_HINT,
            font=(FONT_FAMILY, 9), anchor="w",
            wraplength=640, justify="left",
        )
        self.status_lbl.grid(row=5, column=0, sticky="ew", pady=(0, 10))

        # Progress bar
        self._progress_frame(inner, row=6)

        # Buttons
        btn_row = tk.Frame(inner, bg=BG_CARD)
        btn_row.grid(row=7, column=0, sticky="ew", pady=(20, 0))
        btn_row.grid_columnconfigure(0, weight=1)

        self.botao_gerar = self._primary_button(
            btn_row, f"  {ICON_PLAY}  Gerar PowerPoint  (Ctrl+Enter)", self.gerar,
        )
        self.botao_gerar.grid(row=0, column=1)

        # Footer
        self._footer_lbl = tk.Label(
            outer,
            text=self._footer_text(),
            bg=BG_APP, fg="#2a3550",
            font=(FONT_FAMILY, 8),
        )
        self._footer_lbl.grid(row=3, column=0, pady=(6, 16))

    def _footer_text(self) -> str:
        return f"modelo: {MODEL_FILE.name}  ·  saída: {output_dir(self._config)}"

    # ── Widget factories ──────────────────────────────────────────────────────

    def _field_label(self, parent, text: str, row: int) -> tk.Label:
        lbl = tk.Label(
            parent, text=text, bg=BG_CARD, fg=FG_LABEL,
            font=(FONT_FAMILY, 9, "bold"), anchor="w",
        )
        lbl.grid(row=row, column=0, sticky="w")
        return lbl

    def _entry(self, parent, var: tk.Variable, inner_ipady: int = 6) -> tk.Entry:
        # Wrap in a 1px border frame to simulate a rounded inset border
        border = tk.Frame(parent, bg=INPUT_BORDER, bd=0)
        entry = tk.Entry(
            border, textvariable=var,
            bg=BG_INPUT, fg=FG_INPUT,
            insertbackground=FG_INPUT,
            relief="flat", font=(FONT_FAMILY, 10),
            bd=0, highlightthickness=0,
        )
        entry.pack(fill="both", expand=True, padx=1, pady=1, ipady=inner_ipady)

        def _on_focus_in(_e):
            border.configure(bg=INPUT_BORDER_FOC)
        def _on_focus_out(_e):
            border.configure(bg=INPUT_BORDER)

        entry.bind("<FocusIn>",  _on_focus_in)
        entry.bind("<FocusOut>", _on_focus_out)

        # Forward key methods from border frame to inner entry
        border.focus_set = entry.focus_set

        _orig_configure = border.configure
        def _border_configure(**kw):
            # Keys that belong to Entry (state, width, etc.) go to inner entry
            entry_keys = {k: kw.pop(k) for k in list(kw) if k in ("state", "width")}
            if entry_keys:
                entry.configure(**entry_keys)
            if kw:
                _orig_configure(**kw)
        border.configure = _border_configure

        return border

    def _primary_button(self, parent, text: str, command) -> tk.Button:
        btn = tk.Button(
            parent, text=text, command=command,
            bg=BG_BTN, fg=FG_WHITE,
            activebackground=BG_BTN_HOVER, activeforeground=FG_WHITE,
            relief="flat", bd=0, padx=28, pady=12,
            font=(FONT_FAMILY, 11, "bold"), cursor="hand2",
        )
        btn.bind("<Enter>", lambda _: btn.configure(bg=BG_BTN_HOVER))
        btn.bind("<Leave>", lambda _: btn.configure(bg=BG_BTN))
        return btn

    def _progress_frame(self, parent, row: int) -> None:
        frame = tk.Frame(parent, bg=BG_CARD)
        frame.grid(row=row, column=0, sticky="ew")
        frame.grid_columnconfigure(0, weight=1)

        self._prog_canvas = tk.Canvas(
            frame, height=8, bg=BG_PROG_TRACK, bd=0, highlightthickness=0,
        )
        self._prog_canvas.grid(row=0, column=0, sticky="ew")
        self._prog_canvas.bind("<Configure>", self._redraw_progress)
        self._prog_fill = self._prog_canvas.create_rectangle(
            0, 0, 0, 8, fill=BG_PROG_FILL, outline="",
        )

    def _redraw_progress(self, event=None) -> None:
        w     = self._prog_canvas.winfo_width()
        pct   = self.progresso_var.get() / 100.0
        fill_w = int(w * pct)
        self._prog_canvas.coords(self._prog_fill, 0, 0, fill_w, 8)

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
        atual    = float(self.progresso_var.get())
        destino  = self._target_progress
        diferenca = destino - atual

        if abs(diferenca) < 0.5:
            self.progresso_var.set(destino)
            self._redraw_progress()
            self._progress_animation_running = False
            return

        passo = max(0.8, abs(diferenca) * 0.22)
        self.progresso_var.set(atual + passo if diferenca > 0 else atual - passo)
        self._redraw_progress()
        self.after(20, self._animar_progresso)

    # ── Status helpers ────────────────────────────────────────────────────────

    def _set_status(self, msg: str, color: str = FG_HINT) -> None:
        self.status_var.set(msg)
        self.status_lbl.configure(fg=color)

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
            self.scripture_text.configure(state="normal", fg=FG_INPUT)
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
            bg="#1a3a8f" if gerando else BG_BTN,
        )

    def _limpar_inputs_apos_sucesso(self) -> None:
        self.scripture_text.configure(state="normal", fg=FG_HINT)
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
            self._range_frame.grid(row=1, column=0, sticky="ew")
            self._range_toggle_btn.configure(
                text="▼  Intervalo de versículos  (opcional)"
            )
        else:
            self._range_frame.grid_remove()
            self._range_toggle_btn.configure(
                text="▶  Intervalo de versículos  (opcional)"
            )
            # Clear values when hidden so they don't silently filter on next run.
            self.verso_de_var.set("")
            self.verso_ate_var.set("")

    # ── Scripture field placeholder ───────────────────────────────────────────

    def _limpar_placeholder_scripture(self, _event=None) -> None:
        if self._scripture_tem_placeholder():
            self.scripture_text.delete("1.0", "end")
            self.scripture_text.configure(fg=FG_INPUT)

    def _restaurar_placeholder_scripture(self, _event=None) -> None:
        if not self.scripture_text.get("1.0", "end").strip():
            self.scripture_text.configure(fg=FG_HINT)
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