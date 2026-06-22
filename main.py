from __future__ import annotations

import json
import logging
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


APP_NAME = "GeradorSalmos"
MODEL_NAME = "modelo.pptx"


def app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def resource_path(name: str) -> Path:
    external = app_dir() / name
    if external.exists():
        return external

    bundle_dir = Path(getattr(sys, "_MEIPASS", app_dir()))
    return bundle_dir / name


APP_DIR = app_dir()
MODEL_FILE = resource_path(MODEL_NAME)
LOG_FILE = APP_DIR / "logs.log"
CONFIG_FILE = APP_DIR / "config.json"

PLACEHOLDER_LIVRO_CAP = "{livro_cap}"
PLACEHOLDER_TITULO = "{titulo}"
PLACEHOLDER_VERSICULO = "{versiculo}"

PLACEHOLDERS_LIVRO_CAP = (PLACEHOLDER_LIVRO_CAP, "livro cap", "livro_cap")
PLACEHOLDERS_TITULO = (PLACEHOLDER_TITULO, "titulo", "título")
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

TEMPLATE_TITLE_INDEX = 1
TEMPLATE_PASTOR_INDEX = 2
TEMPLATE_IGREJA_INDEX = 3
TEMPLATE_FINAL_INDEX = 4

PP_ALIGN_JUSTIFY = 4

# ── Design tokens ──────────────────────────────────────────────────────────────
BG_APP        = "#0f1117"
BG_CARD       = "#1a1d27"
BG_INPUT      = "#22263a"
BG_BTN        = "#2563eb"
BG_BTN_HOVER  = "#1d4ed8"
BG_PROG_TRACK = "#252a3d"
BG_PROG_FILL  = "#3b82f6"

FG_TITLE      = "#f0f4ff"
FG_LABEL      = "#a8b3cf"
FG_INPUT      = "#e2e8f0"
FG_HINT       = "#5b6a8a"
FG_SUCCESS    = "#34d399"
FG_ERROR      = "#f87171"
FG_WHITE      = "#ffffff"

FONT_FAMILY   = "Segoe UI"
RADIUS        = 10

# ── Icons (Unicode) ────────────────────────────────────────────────────────────
ICON_BOOK     = "📖"
ICON_PLAY     = "▶"
ICON_CHECK    = "✓"
ICON_CROSS_ALT= "✗"


logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    encoding="utf-8",
)
logger = logging.getLogger(APP_NAME)


class GeradorSalmosError(Exception):
    """Erro esperado, exibido de forma amigavel na interface."""


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


def limpar_nome_arquivo(nome: str) -> str:
    nome_limpo = re.sub(r'[<>:"/\\|?*]', "", nome).strip()
    nome_limpo = re.sub(r"\s+", " ", nome_limpo)
    if not nome_limpo:
        raise GeradorSalmosError("Nao foi possivel montar o nome do arquivo final.")
    return nome_limpo


def caminho_saida_para(leitura: LeituraBiblica) -> Path:
    return APP_DIR / leitura.nome_arquivo


def carregar_configuracao() -> dict[str, str]:
    try:
        if CONFIG_FILE.exists():
            data = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                return {str(key): str(value) for key, value in data.items()}
    except Exception:
        logger.exception("Could not load configuration")
    return {}


def salvar_configuracao(config: dict[str, str]) -> None:
    try:
        CONFIG_FILE.write_text(
            json.dumps(config, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
    except Exception:
        logger.exception("Could not save configuration")


def ler_texto_entrada(texto: str) -> LeituraBiblica:
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
            "Texto da leitura invalido.\n\n"
            f"Linha {numero_linha} nao segue o formato esperado:\n{linha_original}"
        )

    if not livro:
        raise GeradorSalmosError("Texto da leitura invalido: a linha 'Livro:' nao foi encontrada.")
    if not capitulo:
        raise GeradorSalmosError("Texto da leitura invalido: a linha 'Capitulo:' nao foi encontrada.")
    if not versiculos:
        raise GeradorSalmosError("Texto da leitura invalido: nenhum versiculo numerado foi encontrado.")

    return LeituraBiblica(livro=livro, capitulo=capitulo, versiculos=versiculos)


def iterar_formas(forma):
    yield forma

    try:
        if forma.Type == 6:  # msoGroup
            for indice in range(1, forma.GroupItems.Count + 1):
                yield from iterar_formas(forma.GroupItems(indice))
    except Exception:
        return


def normalizar_placeholder(texto: str) -> str:
    texto = texto.replace("\r", "").replace("\n", " ").strip()
    texto = texto.replace("–", "-").replace("—", "-")
    texto = re.sub(r"\s+", " ", texto)
    return texto.casefold()


def centralizar_caixa_no_slide(slide, forma) -> None:
    try:
        page_setup = slide.Parent.PageSetup
        forma.Left = (page_setup.SlideWidth - forma.Width) / 2
        forma.Top = (page_setup.SlideHeight - forma.Height) / 2
    except Exception:
        pass


def justificar_texto(forma) -> None:
    try:
        forma.TextFrame.TextRange.ParagraphFormat.Alignment = PP_ALIGN_JUSTIFY
    except Exception:
        pass

    try:
        forma.TextFrame2.TextRange.ParagraphFormat.Alignment = PP_ALIGN_JUSTIFY
    except Exception:
        pass


def substituir_placeholder(
    slide,
    placeholders: tuple[str, ...],
    novo_texto: str,
    centralizar_caixa: bool = False,
    justificar: bool = False,
) -> bool:
    substituiu = False
    placeholders_norm = {normalizar_placeholder(item) for item in placeholders}

    for forma_raiz in slide.Shapes:
        for forma in iterar_formas(forma_raiz):
            try:
                if not forma.HasTextFrame or not forma.TextFrame.HasText:
                    continue

                text_range = forma.TextFrame.TextRange
                texto_atual = text_range.Text
                texto_norm = normalizar_placeholder(texto_atual)

                if texto_norm in placeholders_norm:
                    text_range.Text = novo_texto
                    if justificar:
                        justificar_texto(forma)
                    if centralizar_caixa:
                        centralizar_caixa_no_slide(slide, forma)
                    substituiu = True
                    continue

                substituiu_na_forma = False
                texto_novo = texto_atual
                for placeholder in placeholders:
                    if placeholder in texto_novo:
                        texto_novo = texto_novo.replace(placeholder, novo_texto)
                        substituiu_na_forma = True

                if substituiu_na_forma:
                    text_range.Text = texto_novo
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
            f"Placeholder {nome_placeholder} nao encontrado no {descricao_slide} do modelo."
        )


def validar_modelo(caminho_modelo: Path) -> None:
    if not caminho_modelo.exists():
        raise GeradorSalmosError(f"Arquivo modelo.pptx nao encontrado:\n{caminho_modelo}")
    if not caminho_modelo.is_file():
        raise GeradorSalmosError(f"modelo.pptx nao e um arquivo valido:\n{caminho_modelo}")


def escolher_indice_template(indice_versiculo: int, total: int) -> int:
    if indice_versiculo == total - 1:
        return TEMPLATE_FINAL_INDEX
    if indice_versiculo % 2 == 0:
        return TEMPLATE_PASTOR_INDEX
    return TEMPLATE_IGREJA_INDEX


def abrir_powerpoint():
    if win32com is None or pythoncom is None:
        raise GeradorSalmosError(
            "A biblioteca pywin32 nao esta instalada.\n\n"
            "Instale com: py -m pip install -r requirements.txt"
        )

    try:
        app = win32com.client.DispatchEx("PowerPoint.Application")
        app.Visible = True
        return app
    except Exception as exc:
        raise GeradorSalmosError(
            "Nao foi possivel abrir o Microsoft PowerPoint.\n\n"
            "Verifique se o PowerPoint esta instalado corretamente neste Windows."
        ) from exc


def gerar_powerpoint(
    texto_leitura: str,
    titulo: str,
    atualizar_progresso: Callable[[int], None] | None = None,
    substituir_existente: bool = False,
) -> ResultadoGeracao:
    inicio = time.perf_counter()
    titulo = titulo.strip()
    if not titulo:
        raise GeradorSalmosError("Informe o titulo do Salmo antes de gerar o PowerPoint.")

    validar_modelo(MODEL_FILE)
    leitura = ler_texto_entrada(texto_leitura)
    caminho_saida = caminho_saida_para(leitura)

    powerpoint = None
    apresentacao = None
    com_inicializado = False

    try:
        atualizar_progresso and atualizar_progresso(10)
        if win32com is None or pythoncom is None:
            raise GeradorSalmosError(
                "A biblioteca pywin32 nao esta instalada.\n\n"
                "Instale com: py -m pip install -r requirements.txt"
            )
        pythoncom.CoInitialize()
        com_inicializado = True
        powerpoint = abrir_powerpoint()

        atualizar_progresso and atualizar_progresso(20)
        apresentacao = powerpoint.Presentations.Open(str(MODEL_FILE), WithWindow=True)

        if apresentacao.Slides.Count < 4:
            raise GeradorSalmosError("O modelo precisa conter exatamente os 4 slides descritos.")

        slide_titulo = apresentacao.Slides(TEMPLATE_TITLE_INDEX)
        exigir_placeholder(
            slide_titulo,
            PLACEHOLDERS_LIVRO_CAP,
            leitura.livro_capitulo,
            "slide 1",
            PLACEHOLDER_LIVRO_CAP,
        )
        exigir_placeholder(
            slide_titulo,
            PLACEHOLDERS_TITULO,
            titulo,
            "slide 1",
            PLACEHOLDER_TITULO,
        )

        total = len(leitura.versiculos)
        for indice, versiculo in enumerate(leitura.versiculos):
            template_index = escolher_indice_template(indice, total)
            slide_modelo = apresentacao.Slides(template_index)
            novo_slide = slide_modelo.Duplicate().Item(1)
            novo_slide.MoveTo(apresentacao.Slides.Count)
            exigir_placeholder(
                novo_slide,
                PLACEHOLDERS_VERSICULO,
                versiculo,
                f"slide modelo {template_index}",
                PLACEHOLDER_VERSICULO,
                centralizar_caixa=True,
                justificar=True,
            )

            progresso = 25 + int(((indice + 1) / total) * 55)
            atualizar_progresso and atualizar_progresso(progresso)

        for _ in range(3):
            apresentacao.Slides(TEMPLATE_PASTOR_INDEX).Delete()

        atualizar_progresso and atualizar_progresso(90)
        if caminho_saida.exists():
            if not substituir_existente:
                raise GeradorSalmosError(
                    f"O arquivo {caminho_saida.name} ja existe.\n\n"
                    "Confirme a substituicao antes de gerar novamente."
                )
            try:
                caminho_saida.unlink()
            except PermissionError as exc:
                raise GeradorSalmosError(
                    "Nao foi possivel substituir o arquivo final existente.\n\n"
                    f"Feche o arquivo {caminho_saida.name}, se ele estiver aberto, e tente novamente."
                ) from exc

        apresentacao.SaveAs(str(caminho_saida))
        atualizar_progresso and atualizar_progresso(100)

        duracao = time.perf_counter() - inicio
        logger.info(
            "Generated | file=%s | verses=%s | seconds=%.2f",
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
        logger.exception("Generation failed with expected error")
        raise
    except PermissionError as exc:
        logger.exception("Permission error while generating presentation")
        raise GeradorSalmosError(
            "Erro de permissao ao salvar o arquivo final.\n\n"
            "Feche o PowerPoint gerado anteriormente, se estiver aberto, e tente novamente."
        ) from exc
    except Exception as exc:
        if pywintypes is not None and isinstance(exc, pywintypes.com_error):
            logger.exception("PowerPoint COM error")
            raise GeradorSalmosError(
                "O PowerPoint retornou um erro durante a geracao.\n\n"
                "Confira se modelo.pptx nao esta protegido, aberto como somente leitura "
                "ou sem os placeholders esperados."
            ) from exc
        logger.exception("Unexpected generation error")
        raise
    finally:
        if apresentacao is not None:
            try:
                apresentacao.Close()
            except Exception:
                logger.exception("Could not close PowerPoint presentation")
        if powerpoint is not None:
            try:
                powerpoint.Quit()
            except Exception:
                logger.exception("Could not quit PowerPoint")
        if pythoncom is not None and com_inicializado:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                logger.exception("Could not uninitialize COM")


# ── Application (CustomTkinter) ────────────────────────────────────────────────

class Aplicacao(ctk.CTk if CTK_AVAILABLE else tk.Tk):
    def __init__(self) -> None:
        if CTK_AVAILABLE:
            ctk.set_appearance_mode("dark")
            ctk.set_default_color_theme("blue")

        super().__init__()

        self.title("Gerador de Salmos")
        self.geometry("780x680")
        self.minsize(660, 600)
        self.configure(bg=BG_APP)
        self.resizable(True, True)

        # Try to set a nice window icon color on Windows 11 (dark title bar)
        try:
            from ctypes import windll, c_int, byref, sizeof
            DWMWA_USE_IMMERSIVE_DARK_MODE = 20
            value = c_int(1)
            windll.dwmapi.DwmSetWindowAttribute(
                windll.user32.GetParent(self.winfo_id()),
                DWMWA_USE_IMMERSIVE_DARK_MODE,
                byref(value),
                sizeof(value),
            )
        except Exception:
            pass

        self.titulo_var  = tk.StringVar()
        self.status_var  = tk.StringVar(value="")
        self.progresso_var = tk.DoubleVar(value=0)
        self._gerando = False
        self._scripture_placeholder = "Paste here the scripture copied from Busca."
        self._config = carregar_configuracao()
        self._target_progress = 0.0
        self._progress_animation_running = False
        self._app_icon = None

        self._aplicar_icone()
        self._montar_interface()
        self._restaurar_geometria()
        self.protocol("WM_DELETE_WINDOW", self._fechar)
        self.bind_all("<Control-Return>", self._atalho_gerar)
        self.after(250, self._perguntar_clipboard_inicial)

    # ── Layout ────────────────────────────────────────────────────────────────

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

    def _montar_interface(self) -> None:
        # Root grid
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        outer = tk.Frame(self, bg=BG_APP)
        outer.grid(row=0, column=0, sticky="nsew")
        outer.grid_columnconfigure(0, weight=1)
        outer.grid_rowconfigure(1, weight=1)

        # ── Header ────────────────────────────────────────────────────────────
        header = tk.Frame(outer, bg=BG_APP)
        header.grid(row=0, column=0, sticky="ew", padx=32, pady=(28, 0))
        header.grid_columnconfigure(0, weight=1)

        icon_lbl = tk.Label(header, text=ICON_BOOK, bg=BG_APP,
                            font=(FONT_FAMILY, 28), fg="#93c5fd")
        icon_lbl.grid(row=0, column=0, sticky="w")

        tk.Label(
            header,
            text="Gerador de Leitura Bíblica",
            bg=BG_APP,
            fg=FG_TITLE,
            font=(FONT_FAMILY, 22, "bold"),
        ).grid(row=1, column=0, sticky="w", pady=(2, 0))

        tk.Label(
            header,
            text="Gera apresentações .pptx preservando fontes, cores, animações e transições do modelo.",
            bg=BG_APP,
            fg=FG_HINT,
            font=(FONT_FAMILY, 10),
        ).grid(row=2, column=0, sticky="w", pady=(3, 0))

        # ── Divider ───────────────────────────────────────────────────────────
        tk.Frame(outer, bg="#252a3d", height=1).grid(
            row=1, column=0, sticky="ew", padx=32, pady=20
        )

        # ── Card ──────────────────────────────────────────────────────────────
        card = tk.Frame(outer, bg=BG_CARD, bd=0)
        card.grid(row=2, column=0, sticky="nsew", padx=32, pady=(0, 8))
        card.grid_columnconfigure(0, weight=1)
        outer.grid_rowconfigure(2, weight=1)

        # Add rounded feel via inner padding
        inner = tk.Frame(card, bg=BG_CARD)
        inner.pack(fill="both", expand=True, padx=28, pady=24)
        inner.grid_columnconfigure(0, weight=1)

        # ── Scripture input ───────────────────────────────────────────────────
        self._field_label(inner, "Scripture", row=0)

        scripture_frame = tk.Frame(inner, bg=BG_CARD)
        scripture_frame.grid(row=1, column=0, sticky="nsew", pady=(6, 20))
        scripture_frame.grid_columnconfigure(0, weight=1)
        scripture_frame.grid_rowconfigure(0, weight=1)
        inner.grid_rowconfigure(1, weight=1)

        self.scripture_text = tk.Text(
            scripture_frame,
            bg=BG_INPUT,
            fg=FG_HINT,
            insertbackground=FG_INPUT,
            relief="flat",
            font=(FONT_FAMILY, 10),
            bd=0,
            highlightthickness=1,
            highlightbackground="#2e3a52",
            highlightcolor=BG_BTN,
            wrap="word",
            height=12,
            padx=10,
            pady=10,
        )
        self.scripture_text.grid(row=0, column=0, sticky="nsew")
        self.scripture_text.insert("1.0", self._scripture_placeholder)
        self.scripture_text.bind("<FocusIn>", self._limpar_placeholder_scripture)
        self.scripture_text.bind("<FocusOut>", self._restaurar_placeholder_scripture)

        scrollbar = tk.Scrollbar(scripture_frame, command=self.scripture_text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.scripture_text.configure(yscrollcommand=scrollbar.set)

        # ── Title input ───────────────────────────────────────────────────────
        self._field_label(inner, "Reading title", row=2)

        self.entry_titulo = self._entry(
            inner,
            self.titulo_var,
            placeholder="Ex: Salmo da Misericórdia",
        )
        self.entry_titulo.grid(row=3, column=0, sticky="ew", pady=(6, 20), ipady=6)

        # ── Status ────────────────────────────────────────────────────────────
        self.status_lbl = tk.Label(
            inner,
            textvariable=self.status_var,
            bg=BG_CARD,
            fg=FG_HINT,
            font=(FONT_FAMILY, 9),
            anchor="w",
            wraplength=640,
            justify="left",
        )
        self.status_lbl.grid(row=4, column=0, sticky="ew", pady=(0, 8))

        # ── Progress bar ──────────────────────────────────────────────────────
        self._progress_frame(inner, row=5)

        # ── Buttons ───────────────────────────────────────────────────────────
        btn_row = tk.Frame(inner, bg=BG_CARD)
        btn_row.grid(row=6, column=0, sticky="ew", pady=(18, 0))
        btn_row.grid_columnconfigure(0, weight=1)

        self.botao_gerar = self._primary_button(
            btn_row,
            f"  {ICON_PLAY}  Generate PowerPoint",
            self.gerar,
        )
        self.botao_gerar.grid(row=0, column=1)

        # ── Footer ────────────────────────────────────────────────────────────
        tk.Label(
            outer,
            text=f"model: {MODEL_FILE.name}  ·  output: same folder as the program",
            bg=BG_APP,
            fg="#2e3a52",
            font=(FONT_FAMILY, 8),
        ).grid(row=3, column=0, pady=(4, 14))

    # ── Widget factories ──────────────────────────────────────────────────────

    def _field_label(self, parent, text: str, row: int) -> tk.Label:
        lbl = tk.Label(
            parent,
            text=text,
            bg=BG_CARD,
            fg=FG_LABEL,
            font=(FONT_FAMILY, 9, "bold"),
            anchor="w",
        )
        lbl.grid(row=row, column=0, sticky="w")
        return lbl

    def _entry(self, parent, var: tk.Variable, placeholder: str = "") -> tk.Entry:
        e = tk.Entry(
            parent,
            textvariable=var,
            bg=BG_INPUT,
            fg=FG_INPUT,
            insertbackground=FG_INPUT,
            relief="flat",
            font=(FONT_FAMILY, 10),
            bd=0,
            highlightthickness=1,
            highlightbackground="#2e3a52",
            highlightcolor=BG_BTN,
        )
        return e

    def _primary_button(self, parent, text: str, command) -> tk.Button:
        btn = tk.Button(
            parent,
            text=text,
            command=command,
            bg=BG_BTN,
            fg=FG_WHITE,
            activebackground=BG_BTN_HOVER,
            activeforeground=FG_WHITE,
            relief="flat",
            bd=0,
            padx=22,
            pady=10,
            font=(FONT_FAMILY, 11, "bold"),
            cursor="hand2",
        )
        btn.bind("<Enter>", lambda _: btn.configure(bg=BG_BTN_HOVER))
        btn.bind("<Leave>", lambda _: btn.configure(bg=BG_BTN))
        return btn

    def _progress_frame(self, parent, row: int) -> None:
        """Canvas-based progress bar with rounded fill."""
        frame = tk.Frame(parent, bg=BG_CARD)
        frame.grid(row=row, column=0, sticky="ew")
        frame.grid_columnconfigure(0, weight=1)

        self._prog_canvas = tk.Canvas(
            frame,
            height=6,
            bg=BG_PROG_TRACK,
            bd=0,
            highlightthickness=0,
        )
        self._prog_canvas.grid(row=0, column=0, sticky="ew")
        self._prog_canvas.bind("<Configure>", self._redraw_progress)

        # Draw initial empty bar
        self._prog_fill = self._prog_canvas.create_rectangle(
            0, 0, 0, 6, fill=BG_PROG_FILL, outline=""
        )

    def _redraw_progress(self, event=None) -> None:
        w = self._prog_canvas.winfo_width()
        pct = self.progresso_var.get() / 100.0
        fill_w = int(w * pct)
        self._prog_canvas.coords(self._prog_fill, 0, 0, fill_w, 6)

    # ── Callbacks ─────────────────────────────────────────────────────────────

    def _restaurar_geometria(self) -> None:
        geometria = self._config.get("geometry", "")
        if re.fullmatch(r"\d+x\d+[+-]\d+[+-]\d+", geometria):
            self.geometry(geometria)
            return

        self.update_idletasks()
        largura = self.winfo_width()
        altura = self.winfo_height()
        x = max((self.winfo_screenwidth() - largura) // 2, 0)
        y = max((self.winfo_screenheight() - altura) // 2, 0)
        self.geometry(f"{largura}x{altura}+{x}+{y}")

    def _fechar(self) -> None:
        if self._gerando:
            return
        self._config["geometry"] = self.geometry()
        salvar_configuracao(self._config)
        self.destroy()

    def _atalho_gerar(self, _event=None) -> str:
        self.gerar()
        return "break"

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
            ler_texto_entrada(texto)
        except GeradorSalmosError:
            return

        if messagebox.askyesno(
            "Scripture detected",
            "Scripture detected in clipboard. Paste it?",
        ):
            self.scripture_text.configure(state="normal", fg=FG_INPUT)
            self.scripture_text.delete("1.0", "end")
            self.scripture_text.insert("1.0", texto)
            self.scripture_text.focus_set()

    def _set_ui_gerando(self, gerando: bool) -> None:
        estado = "disabled" if gerando else "normal"
        self.scripture_text.configure(state=estado)
        self.entry_titulo.configure(state=estado)
        self.botao_gerar.configure(
            state=estado,
            bg="#1a3a8f" if gerando else BG_BTN,
        )

    def _limpar_inputs_apos_sucesso(self) -> None:
        self.scripture_text.configure(state="normal", fg=FG_HINT)
        self.scripture_text.delete("1.0", "end")
        self.scripture_text.insert("1.0", self._scripture_placeholder)
        self.titulo_var.set("")
        self.scripture_text.focus_set()

    def _limpar_placeholder_scripture(self, _event=None) -> None:
        if self._scripture_tem_placeholder():
            self.scripture_text.delete("1.0", "end")
            self.scripture_text.configure(fg=FG_INPUT)

    def _restaurar_placeholder_scripture(self, _event=None) -> None:
        if not self.scripture_text.get("1.0", "end").strip():
            self.scripture_text.configure(fg=FG_HINT)
            self.scripture_text.delete("1.0", "end")
            self.scripture_text.insert("1.0", self._scripture_placeholder)

    def _scripture_tem_placeholder(self) -> bool:
        return self.scripture_text.get("1.0", "end").strip() == self._scripture_placeholder

    def _obter_texto_scripture(self) -> str:
        if self._scripture_tem_placeholder():
            return ""
        return self.scripture_text.get("1.0", "end").strip()

    def _set_status(self, msg: str, color: str = FG_HINT) -> None:
        self.status_var.set(msg)
        self.status_lbl.configure(fg=color)

    def _set_progresso_imediato(self, valor: int) -> None:
        self.progresso_var.set(valor)
        self._target_progress = float(valor)
        self._redraw_progress()
        self.update_idletasks()

    def _set_progresso(self, valor: int) -> None:
        self._target_progress = max(0.0, min(100.0, float(valor)))
        if not self._progress_animation_running:
            self._progress_animation_running = True
            self._animar_progresso()

    def _animar_progresso(self) -> None:
        atual = float(self.progresso_var.get())
        destino = self._target_progress
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

    def atualizar_progresso(self, valor: int) -> None:
        self.after(0, self._atualizar_progresso_ui, valor)

    def _atualizar_progresso_ui(self, valor: int) -> None:
        msgs = {
            10: "Abrindo PowerPoint…",
            20: "Carregando modelo…",
            90: "Salvando apresentação…",
            100: "Concluído!",
        }
        if valor in msgs:
            self._set_status(msgs[valor])
        elif valor > 20:
            self._set_status(f"Gerando slides… {valor}%")
        self._set_progresso(valor)

    def gerar(self) -> None:
        if self._gerando:
            return

        texto_leitura = self._obter_texto_scripture()
        try:
            leitura = ler_texto_entrada(texto_leitura)
        except GeradorSalmosError as exc:
            logger.warning("Invalid scripture: %s", exc)
            self._set_progresso_imediato(0)
            self._set_status(f"{ICON_CROSS_ALT}  Scripture text is incomplete or invalid.", FG_ERROR)
            messagebox.showerror("Invalid scripture", str(exc))
            return

        titulo_leitura = self.titulo_var.get().strip()
        if not titulo_leitura:
            logger.warning("Generation blocked: missing title")
            self._set_progresso_imediato(0)
            self._set_status(f"{ICON_CROSS_ALT}  Informe o titulo da leitura.", FG_ERROR)
            messagebox.showerror(
                "Missing title",
                "Please type the reading title before generating the presentation.",
            )
            return

        caminho_saida = caminho_saida_para(leitura)
        substituir_existente = False
        if caminho_saida.exists():
            substituir_existente = messagebox.askyesno(
                "Replace existing file?",
                f"The file already exists:\n\n{caminho_saida.name}\n\n"
                "Do you want to replace it?",
            )
            if not substituir_existente:
                logger.info("Generation canceled: existing file kept | file=%s", caminho_saida.name)
                self._set_status("Generation canceled. Existing file was kept.")
                return

        self._gerando = True
        self._set_ui_gerando(True)
        self._set_progresso_imediato(0)
        self._set_status("Starting generation…")
        self.update_idletasks()

        def _worker():
            try:
                resultado = gerar_powerpoint(
                    texto_leitura,
                    titulo_leitura,
                    self.atualizar_progresso,
                    substituir_existente=substituir_existente,
                )
            except GeradorSalmosError as exc:
                self.after(0, self._on_erro, str(exc))
            except Exception:
                logger.exception("Unexpected UI worker error")
                self.after(0, self._on_erro_inesperado)
            else:
                self.after(0, self._on_sucesso, resultado)

        threading.Thread(target=_worker, daemon=True).start()

    def _on_sucesso(self, resultado: ResultadoGeracao) -> None:
        caminho_saida = resultado.caminho
        self._set_progresso(100)
        self._set_status(f"{ICON_CHECK}  {caminho_saida.name} salvo com sucesso.", FG_SUCCESS)
        self._gerando = False
        self._set_ui_gerando(False)

        pasta_saida = caminho_saida.parent
        try:
            subprocess.Popen(["explorer.exe", f'/select,"{caminho_saida}"'])
        except Exception:
            logger.exception("Could not open Explorer")

        messagebox.showinfo(
            "Presentation generated successfully!",
            "Presentation generated successfully!\n\n"
            f"File: {caminho_saida.name}\n"
            f"Verses: {resultado.quantidade_versiculos}\n"
            f"Generation time: {resultado.duracao_segundos:.1f} seconds\n"
            f"Output folder: {pasta_saida}",
        )
        self._limpar_inputs_apos_sucesso()

    def _on_erro(self, mensagem: str) -> None:
        self._set_progresso_imediato(0)
        self._set_status(f"{ICON_CROSS_ALT}  Não foi possível gerar a apresentação.", FG_ERROR)
        self._gerando = False
        self._set_ui_gerando(False)
        messagebox.showerror("Erro ao gerar PowerPoint", mensagem)

    def _on_erro_inesperado(self) -> None:
        self._set_progresso_imediato(0)
        self._set_status(f"{ICON_CROSS_ALT}  Ocorreu um erro inesperado.", FG_ERROR)
        self._gerando = False
        self._set_ui_gerando(False)
        messagebox.showerror(
            "Erro inesperado",
            "Ocorreu um erro inesperado ao gerar o PowerPoint.\n\n"
            "Os detalhes foram registrados em logs.log.",
        )


def main() -> int:
    if not CTK_AVAILABLE:
        # Fallback warning — app still runs with tkinter
        import warnings
        warnings.warn(
            "CustomTkinter não encontrado. Instale com: pip install customtkinter\n"
            "A interface continuará funcionando com tkinter padrão.",
            stacklevel=1,
        )
    app = Aplicacao()
    app.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
