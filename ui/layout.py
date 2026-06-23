from __future__ import annotations

import subprocess
import threading
import tkinter as tk
from datetime import datetime
from typing import Any

from tkinter import filedialog, messagebox

import customtkinter as ctk

from core.config import LOG_FILE, caminho_saida_para, carregar_configuracao, output_dir, salvar_configuracao, validar_pasta_saida
from core.logging import logger
from core.generator import ResultadoGeracao, gerar_powerpoint
from core.scripture import ler_texto_entrada
from core.utils import GeradorSalmosError
from ui.animations import FluentAnimations
from ui.theme import (
    APP_AUTHOR,
    APP_NAME,
    APP_VERSION,
    BG_APP,
    BG_CARD,
    BG_CARD_INNER,
    BG_INPUT,
    BG_TILE,
    CARD_BORDER,
    FG_ERROR,
    FG_INPUT,
    FG_LABEL,
    FG_HINT,
    FG_SUCCESS,
    FG_TITLE,
    FONT_FAMILY,
    GRAD_MID,
    GRAD_START,
    ICON_BOOK,
    ICON_CHECK,
    ICON_CROSS,
    ICON_FOLDER,
    ICON_PLAY,
    INPUT_BORDER,
    INPUT_BORDER_FOC,
    RADIUS,
    RADIUS_PILL,
    RADIUS_TILE,
)
from ui.widgets import FluentWidgets

try:
    from PIL import Image, ImageDraw, ImageTk
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False


class Aplicacao(ctk.CTk):
    """Main application window for scripture-to-PowerPoint generation."""

    _SCRIPTURE_PLACEHOLDER = "Cole aqui o texto copiado do Busca."

    def __init__(self) -> None:
        """Initialize widgets, state, and startup animations."""
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        super().__init__()

        self.withdraw()
        self.title("Gerador de Leitura Bíblica")
        self.geometry("860x740")
        self.minsize(700, 640)
        self.configure(fg_color=BG_APP)
        self.resizable(True, True)

        self._apply_dark_titlebar()

        self.titulo_var = tk.StringVar()
        self.status_var = tk.StringVar(value="")
        self.progresso_var = tk.DoubleVar(value=0)
        self.verso_de_var = tk.StringVar()
        self.verso_ate_var = tk.StringVar()
        self.pasta_saida_var = tk.StringVar()
        self._gerando = False
        self._config = carregar_configuracao()
        self.pasta_saida_var.set(str(output_dir(self._config)))
        self._target_progress = 0.0
        self._progress_animation_running = False
        self._card_opacity_token = [False]
        self._original_window_height = None
        self._app_icon = None
        self._ultimo_resultado: ResultadoGeracao | None = None

        self._aplicar_icone()
        self._header_ref = None
        self._div_ref = None
        self._card_ref = None
        self._range_outer_ref = None
        self._progress_frame_ref = None
        self._btn_row_ref = None

        self._startup_ready = False
        self._splash_start_time = 0.0
        self._startup_splash = None

        self._montar_interface()
        self._restaurar_geometria()
        self.protocol("WM_DELETE_WINDOW", self._fechar)
        self.bind_all("<Control-Return>", self._atalho_gerar)
        self.after(250, self._perguntar_clipboard_inicial)
        self._schedule_splash_close(self._create_splash_screen())

        logger.info("Window initialised | geometry=%s", self.geometry())

    def _apply_dark_titlebar(self) -> None:
        try:
            from ctypes import windll, c_int, byref, sizeof
            value = c_int(1)
            windll.dwmapi.DwmSetWindowAttribute(
                windll.user32.GetParent(self.winfo_id()),
                20,
                byref(value),
                sizeof(value),
            )
        except (tk.TclError, AttributeError, ValueError, TypeError) as exc:
            logger.debug("Ignored UI operation error: %s", exc, exc_info=True)

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
        except (tk.TclError, RuntimeError, ValueError, TypeError) as exc:
            logger.exception("Could not apply app icon")

    def _create_splash_screen(self) -> tk.Toplevel:
        splash = tk.Toplevel(self)
        splash.overrideredirect(True)
        try:
            splash.attributes("-topmost", True)
        except (tk.TclError, AttributeError, ValueError, TypeError) as exc:
            logger.debug("Ignored UI operation error: %s", exc, exc_info=True)

        splash.configure(bg=BG_APP)
        container = ctk.CTkFrame(splash, fg_color=BG_APP, border_width=1, border_color=CARD_BORDER)
        container.pack(expand=True, fill="both", padx=2, pady=2)
        content = ctk.CTkFrame(container, fg_color=BG_APP)
        content.place(relx=0.5, rely=0.5, anchor="center", relwidth=0.92, relheight=0.92)

        ctk.CTkLabel(content, text=ICON_BOOK, font=(FONT_FAMILY, 38), fg_color=BG_APP, text_color=GRAD_START).pack(pady=(18, 6))
        ctk.CTkLabel(content, text="Gerador de Leitura Bíblica", font=(FONT_FAMILY, 17, "bold"), fg_color=BG_APP, text_color=FG_TITLE).pack(pady=(0, 4))
        ctk.CTkLabel(content, text="Carregando...", font=(FONT_FAMILY, 12), fg_color=BG_APP, text_color=FG_LABEL).pack(pady=(0, 18))

        progress = ctk.CTkProgressBar(content, mode="indeterminate", progress_color=GRAD_MID, fg_color=BG_APP)
        progress.pack(fill="x", padx=36, pady=(0, 16), ipady=3)
        progress.start()

        splash.update_idletasks()
        width = 420
        height = 220
        x = (splash.winfo_screenwidth() - width) // 2
        y = (splash.winfo_screenheight() - height) // 2
        splash.geometry(f"{width}x{height}+{x}+{y}")
        splash.deiconify()
        self._startup_splash = splash
        self._splash_start_time = datetime.now().timestamp()
        return splash

    def _schedule_splash_close(self, splash: tk.Toplevel) -> None:
        min_visible_ms = 500

        def _hide_splash() -> None:
            try:
                splash.destroy()
            except (tk.TclError, AttributeError, ValueError, TypeError) as exc:
                logger.debug("Ignored UI operation error: %s", exc, exc_info=True)
            self._fade_in_main_window()

        def _check_ready() -> None:
            if self._startup_ready:
                _hide_splash()
            else:
                self.after(80, _check_ready)

        def _close_splash() -> None:
            elapsed = int((datetime.now().timestamp() - self._splash_start_time) * 1000)
            remaining = max(min_visible_ms - elapsed, 0)
            if remaining > 0:
                self.after(remaining, _check_ready)
            else:
                _check_ready()

        self.after(0, _close_splash)

    def _fade_in_main_window(self) -> None:
        try:
            self.attributes("-alpha", 0.0)
        except (tk.TclError, AttributeError, ValueError, TypeError):
            self.deiconify()
            return

        self.deiconify()

        def _update(alpha: float) -> None:
            try:
                self.attributes("-alpha", alpha)
            except (tk.TclError, AttributeError, ValueError, TypeError) as exc:
                logger.debug("Ignored UI operation error: %s", exc, exc_info=True)

        def _finish() -> None:
            try:
                self.attributes("-alpha", 1.0)
            except (tk.TclError, AttributeError, ValueError, TypeError) as exc:
                logger.debug("Ignored UI operation error: %s", exc, exc_info=True)

        self.tween(260, _update, ease=self.ease_out_cubic, on_done=_finish)

    def _mark_startup_ready(self) -> None:
        self._startup_ready = True

    def _montar_interface(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        outer = ctk.CTkFrame(self, fg_color=BG_APP)
        outer.grid(row=0, column=0, sticky="nsew")
        outer.grid_columnconfigure(0, weight=1)
        outer.grid_rowconfigure(2, weight=1)

        header = ctk.CTkFrame(outer, fg_color=BG_APP)
        header.grid(row=0, column=0, sticky="ew", padx=36, pady=(36, 0))
        header.grid_columnconfigure(1, weight=1)
        self._header_ref = header

        accent_pill_img = self._make_vertical_gradient_image(10, 54, (0x93, 0xc5, 0xfd), (0x1d, 0x4e, 0xd8), radius=5)
        if accent_pill_img is not None:
            accent_pill = ctk.CTkLabel(header, image=accent_pill_img, text="", fg_color=BG_APP, width=10, height=54)
            accent_pill.image = accent_pill_img
        else:
            accent_pill = ctk.CTkFrame(header, fg_color=GRAD_START, corner_radius=6, width=6, height=54)
            accent_pill.grid_propagate(False)
        accent_pill.grid(row=0, column=0, rowspan=2, sticky="ns", padx=(0, 14))

        title_block = ctk.CTkFrame(header, fg_color=BG_APP)
        title_block.grid(row=0, column=1, sticky="ew")
        title_block.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(title_block, text=f"{ICON_BOOK}  Gerador de Leitura Bíblica", fg_color=BG_APP, text_color=FG_TITLE, font=(FONT_FAMILY, 26, "bold")).grid(row=0, column=0, sticky="w")
        ctk.CTkLabel(title_block, text="Gera apresentações .pptx preservando fontes, cores, animações e transições do modelo.", fg_color=BG_APP, text_color=FG_LABEL, font=(FONT_FAMILY, 12)).grid(row=1, column=0, sticky="w", pady=(5, 0))

        div_canvas = tk.Canvas(outer, height=2, bg=BG_APP, bd=0, highlightthickness=0)
        div_canvas.grid(row=1, column=0, sticky="ew", padx=36, pady=22)
        self._div_ref = div_canvas

        def _draw_divider(event=None):
            w = div_canvas.winfo_width()
            h = max(1, div_canvas.winfo_height())
            if w < 2:
                return
            div_canvas.delete("all")
            stops = [
                (0.0, 0x93, 0xc5, 0xfd),
                (0.5, 0x3b, 0x82, 0xf6),
                (1.0, 0x1d, 0x4e, 0xd8),
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
                    div_canvas._gradient_img = photo
                    return
                except (OSError, ValueError, RuntimeError, TypeError):
                    logger.exception("Could not render divider gradient via PIL")

            for i in range(max(w, 1)):
                r, g, bl = _stop_color(i / max(w, 1))
                div_canvas.create_line(i, 0, i, h, fill=f"#{r:02x}{g:02x}{bl:02x}")

        div_canvas.bind("<Configure>", lambda e: _draw_divider())

        card = ctk.CTkFrame(outer, fg_color=BG_CARD, corner_radius=RADIUS, border_width=1, border_color=CARD_BORDER)
        card.grid(row=2, column=0, sticky="nsew", padx=32, pady=(0, 12))
        card.grid_columnconfigure(0, weight=1)
        card.grid_rowconfigure(0, weight=1)
        self._card_ref = card

        inner = ctk.CTkFrame(card, fg_color=BG_CARD, corner_radius=RADIUS - 4)
        inner.pack(fill="both", expand=True, padx=32, pady=32)
        inner.grid_columnconfigure(0, weight=1)
        inner.grid_rowconfigure(1, weight=1)

        self._field_label(inner, "Texto Bíblico", row=0)
        self.scripture_text = ctk.CTkTextbox(inner, fg_color=BG_INPUT, text_color=FG_HINT, font=(FONT_FAMILY, 13), wrap="word", height=100, corner_radius=RADIUS_TILE, border_width=1, border_color=INPUT_BORDER)
        self.scripture_text.grid(row=1, column=0, sticky="nsew", pady=(8, 26))
        self.scripture_text.insert("1.0", self._SCRIPTURE_PLACEHOLDER)
        self.scripture_text.bind("<FocusIn>", self._limpar_placeholder_scripture)
        self.scripture_text.bind("<FocusOut>", self._restaurar_placeholder_scripture)

        def _caret_show(_e=None):
            try:
                self.scripture_text._textbox.configure(insertwidth=2, insertbackground=FG_INPUT)
            except (tk.TclError, AttributeError, ValueError, TypeError) as exc:
                logger.debug("Ignored UI operation error: %s", exc, exc_info=True)

        def _caret_hide(_e=None):
            try:
                self.scripture_text._textbox.configure(insertwidth=0)
            except (tk.TclError, AttributeError, ValueError, TypeError) as exc:
                logger.debug("Ignored UI operation error: %s", exc, exc_info=True)

        self.scripture_text.bind("<FocusIn>", _caret_show, add="+")
        self.scripture_text.bind("<FocusOut>", _caret_hide, add="+")

        self._field_label(inner, "Título da leitura", row=2)
        self.entry_titulo = FluentWidgets.entry(inner, self.titulo_var, inner_ipady=8)
        self.entry_titulo.grid(row=3, column=0, sticky="ew", pady=(8, 26))

        self._field_label(inner, "Salvar em", row=4)
        pasta_row = ctk.CTkFrame(inner, fg_color=BG_CARD, corner_radius=0)
        pasta_row.grid(row=5, column=0, sticky="ew", pady=(8, 26))
        pasta_row.grid_columnconfigure(0, weight=1)
        self.entry_pasta_saida = FluentWidgets.entry(pasta_row, self.pasta_saida_var, inner_ipady=8)
        self.entry_pasta_saida.grid(row=0, column=0, sticky="ew", padx=(0, 10))
        self._botao_escolher_pasta = ctk.CTkButton(pasta_row, text=f"{ICON_FOLDER}  Procurar", command=self._escolher_pasta_saida, fg_color=BG_TILE, text_color=FG_LABEL, hover_color=BG_TILE, font=(FONT_FAMILY, 13), cursor="hand2", corner_radius=RADIUS_PILL, width=110, height=44)
        self._botao_escolher_pasta.grid(row=0, column=1)

        range_outer = ctk.CTkFrame(inner, fg_color=BG_TILE, corner_radius=RADIUS_TILE, border_width=1, border_color=CARD_BORDER)
        range_outer.grid(row=6, column=0, sticky="ew", pady=(0, 26))
        range_outer.grid_columnconfigure(0, weight=1)
        self._range_outer_ref = range_outer

        def _force_uniform_corners() -> None:
            try:
                range_outer.update_idletasks()
                range_outer._draw(no_color_updates=True)
            except (tk.TclError, AttributeError, ValueError, TypeError) as exc:
                logger.debug("Ignored UI operation error: %s", exc, exc_info=True)

        self.after(60, _force_uniform_corners)

        self._range_visible = False
        self._range_toggle_btn = ctk.CTkButton(range_outer, text="▶  Intervalo de versículos  (opcional)", command=self._toggle_range_section, fg_color=BG_TILE, text_color=FG_LABEL, hover_color=BG_TILE, font=(FONT_FAMILY, 12, "bold"), cursor="hand2", anchor="w", corner_radius=RADIUS_TILE)
        self._range_toggle_btn.grid(row=0, column=0, sticky="ew")
        self._range_frame = ctk.CTkFrame(range_outer, fg_color=BG_TILE, corner_radius=RADIUS_TILE)
        self._range_frame.grid_columnconfigure(1, weight=0)

        ctk.CTkLabel(self._range_frame, text="De:", fg_color=BG_TILE, text_color=FG_LABEL, font=(FONT_FAMILY, 12)).grid(row=0, column=0, sticky="w", padx=(0, 8), pady=(6, 0))
        self.entry_verso_de = FluentWidgets.entry(self._range_frame, self.verso_de_var, inner_ipady=4)
        self.entry_verso_de.configure(width=80, height=32)
        self.entry_verso_de.grid(row=0, column=1, sticky="w", pady=(6, 0))
        ctk.CTkLabel(self._range_frame, text="Até:", fg_color=BG_TILE, text_color=FG_LABEL, font=(FONT_FAMILY, 12)).grid(row=0, column=2, sticky="w", padx=(24, 8), pady=(6, 0))
        self.entry_verso_ate = FluentWidgets.entry(self._range_frame, self.verso_ate_var, inner_ipady=4)
        self.entry_verso_ate.configure(width=80, height=32)
        self.entry_verso_ate.grid(row=0, column=3, sticky="w", pady=(6, 0))
        ctk.CTkLabel(self._range_frame, text="Deixe em branco para gerar todos os versículos.", fg_color=BG_TILE, text_color=FG_HINT, font=(FONT_FAMILY, 11)).grid(row=1, column=0, columnspan=4, sticky="w", pady=(4, 4))

        self.status_lbl = ctk.CTkLabel(inner, textvariable=self.status_var, fg_color=BG_CARD, text_color=FG_HINT, font=(FONT_FAMILY, 12), anchor="w", wraplength=680, justify="left")
        self.status_lbl.grid(row=7, column=0, sticky="ew", pady=(0, 12))
        self._progress_frame(inner, row=8)

        btn_row = ctk.CTkFrame(inner, fg_color=BG_TILE, corner_radius=RADIUS_TILE, border_width=1, border_color=CARD_BORDER)
        btn_row.grid(row=9, column=0, sticky="e", pady=(24, 0), padx=(0, 0))
        self._btn_row_ref = btn_row
        self.botao_gerar = FluentWidgets.primary_button(btn_row, f"  {ICON_PLAY}  Gerar PowerPoint  (Ctrl+Enter)", self.gerar, GRAD_START, GRAD_MID)
        self.botao_gerar.grid(row=0, column=0, pady=10, padx=10)

        self._footer_lbl = ctk.CTkLabel(outer, text=self._footer_text(), fg_color=BG_APP, text_color="#5a649a", font=(FONT_FAMILY, 11))
        self._footer_lbl.grid(row=3, column=0, pady=(2, 6))

        self.after(10, self._wire_fluent_animations)

    def _field_label(self, parent, text: str, row: int) -> None:
        FluentWidgets.field_label(parent, text, row)

    def _footer_text(self) -> str:
        return f"© {datetime.now().year}  |  {APP_NAME} v{APP_VERSION}  |  {APP_AUTHOR}"

    def _progress_frame(self, parent, row: int) -> None:
        frame = ctk.CTkFrame(parent, fg_color=BG_TILE, corner_radius=RADIUS_TILE, border_width=1, border_color=CARD_BORDER)
        frame.grid(row=row, column=0, sticky="ew")
        frame.grid_columnconfigure(0, weight=1)
        self._progress_frame_ref = frame

        self._ctk_progress = ctk.CTkProgressBar(frame, fg_color=BG_TILE, progress_color=GRAD_START, corner_radius=8, height=12, mode="determinate")
        self._ctk_progress.set(0)
        self._ctk_progress.configure(progress_color=BG_TILE)
        self._ctk_progress.grid(row=0, column=0, sticky="ew", padx=14, pady=14)
        try:
            self._ctk_progress.configure(cursor="")
        except (tk.TclError, AttributeError, ValueError, TypeError) as exc:
            logger.debug("Ignored UI operation error: %s", exc, exc_info=True)
        self.progresso_var.trace_add("write", self._sync_ctk_progress)

    def _sync_ctk_progress(self, *_):
        try:
            valor = self.progresso_var.get()
            self._ctk_progress.set(valor / 100.0)
            if valor <= 0 or valor >= 99.5:
                self._ctk_progress.configure(progress_color=BG_TILE)
            else:
                self._ctk_progress.configure(progress_color=GRAD_START)
        except (tk.TclError, AttributeError, ValueError, TypeError) as exc:
            logger.debug("Ignored UI operation error: %s", exc, exc_info=True)

    def _redraw_progress(self, event=None) -> None:
        self._sync_ctk_progress()

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
        atual = float(self.progresso_var.get())
        destino = self._target_progress
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
        self.after(16, self._animar_progresso)

    def _set_status(self, msg: str, color: str = FG_HINT) -> None:
        try:
            self._status_transition(msg, color)
        except (tk.TclError, AttributeError, ValueError, TypeError):
            self.status_var.set(msg)
            self.status_lbl.configure(text_color=color)

    def _set_status_gerando(self, msg: str) -> None:
        self._set_status(msg, FG_LABEL)

    def _set_status_ok(self, msg: str) -> None:
        self._set_status(f"{ICON_CHECK}  {msg}", FG_SUCCESS)

    def _set_status_erro(self, msg: str) -> None:
        self._set_status(f"{ICON_CROSS}  {msg}", FG_ERROR)

    _PROGRESS_MESSAGES: dict[int, str] = {
        10: "Iniciando o PowerPoint…",
        20: "Abrindo o modelo…",
        25: "Preenchendo slide de título…",
        90: "Salvando apresentação…",
        100: "Concluído!",
    }

    def atualizar_progresso(self, valor: int) -> None:
        """Schedule progress updates on the UI thread.

        Args:
            valor: Progress percentage from 0 to 100.
        """
        self.after(0, self._atualizar_progresso_ui, valor)

    def _atualizar_progresso_ui(self, valor: int) -> None:
        if valor in self._PROGRESS_MESSAGES:
            self._set_status_gerando(self._PROGRESS_MESSAGES[valor])
        elif 25 < valor < 90:
            pct = int(((valor - 25) / 65) * 100)
            self._set_status_gerando(f"Gerando versículos… {pct}%")
        self._set_progresso(valor)

    def _escolher_pasta_saida(self) -> None:
        pasta_atual = self.pasta_saida_var.get().strip() or str(output_dir(self._config))
        escolhida = filedialog.askdirectory(initialdir=pasta_atual, title="Selecionar pasta de destino", parent=self)
        if not escolhida:
            return
        try:
            pasta_validada = validar_pasta_saida(escolhida)
        except GeradorSalmosError as exc:
            messagebox.showerror("Pasta inválida", str(exc), parent=self)
            return

        self.pasta_saida_var.set(str(pasta_validada))
        self._config["output_dir"] = str(pasta_validada)
        salvar_configuracao(self._config)
        logger.info("Output folder updated | output_dir=%s", pasta_validada)
        try:
            self._footer_lbl.configure(text=self._footer_text())
        except (tk.TclError, AttributeError, ValueError, TypeError) as exc:
            logger.debug("Ignored UI operation error: %s", exc, exc_info=True)

    def _restaurar_geometria(self) -> None:
        geometria = self._config.get("geometry", "")
        if __import__("re").fullmatch(r"\d+x\d+[+-]\d+[+-]\d+", geometria):
            self.geometry(geometria)
            return
        self.update_idletasks()
        w = self.winfo_width()
        h = self.winfo_height()
        x = max((self.winfo_screenwidth() - w) // 2, 0)
        y = max((self.winfo_screenheight() - h) // 2, 0)
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _fechar(self) -> None:
        if self._gerando:
            return
        self._config["geometry"] = self.geometry()
        salvar_configuracao(self._config)
        logger.info("Window closed | geometry=%s", self._config["geometry"])
        self.destroy()

    def _atalho_gerar(self, _event=None) -> str:
        self.gerar()
        return "break"

    def _perguntar_clipboard_inicial(self) -> None:
        if self._gerando or not self._scripture_tem_placeholder():
            return
        try:
            texto = self.clipboard_get().strip()
        except tk.TclError:
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

    def _set_ui_gerando(self, gerando: bool) -> None:
        estado = "disabled" if gerando else "normal"
        self.scripture_text.configure(state=estado)
        self.entry_titulo.configure(state=estado)
        self.entry_verso_de.configure(state=estado)
        self.entry_verso_ate.configure(state=estado)
        self.botao_gerar.configure(state=estado, fg_color="#163060" if gerando else GRAD_START)
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

    def _toggle_range_section(self) -> None:
        self._range_visible = not self._range_visible
        if self._range_visible:
            self._range_toggle_btn.configure(text="▼  Intervalo de versículos  (opcional)")
            self._grow_window_for_range()
            self._animate_range_show()
        else:
            self._range_toggle_btn.configure(text="▶  Intervalo de versículos  (opcional)")
            self._animate_range_hide()
            self._shrink_window_for_range()

    def _grow_window_for_range(self) -> None:
        try:
            if self.state() != "normal":
                return
        except (tk.TclError, AttributeError, ValueError, TypeError) as exc:
            logger.debug("Ignored UI operation error: %s", exc, exc_info=True)
        try:
            self.update_idletasks()
            if self._original_window_height is None:
                self._original_window_height = self.winfo_height()
            needed = self._range_frame.winfo_reqheight() + 4
        except (tk.TclError, AttributeError, ValueError, TypeError):
            return
        try:
            w = self.winfo_width()
            x, y = self.winfo_x(), self.winfo_y()
            target_h = self._original_window_height + max(0, needed)
            if target_h > self.winfo_height():
                self.geometry(f"{w}x{target_h}+{x}+{y}")
                self.update_idletasks()
        except (tk.TclError, AttributeError, ValueError, TypeError) as exc:
            logger.debug("Ignored UI operation error: %s", exc, exc_info=True)

    def _shrink_window_for_range(self) -> None:
        if self._original_window_height is None:
            return
        try:
            if self.state() != "normal":
                return
        except (tk.TclError, AttributeError, ValueError, TypeError) as exc:
            logger.debug("Ignored UI operation error: %s", exc, exc_info=True)
        try:
            w = self.winfo_width()
            x, y = self.winfo_x(), self.winfo_y()
            min_h = self.minsize()[1]
            new_h = max(min_h, self._original_window_height)
            self.geometry(f"{w}x{new_h}+{x}+{y}")
            self.update_idletasks()
        except (tk.TclError, AttributeError, ValueError, TypeError) as exc:
            logger.debug("Ignored UI operation error: %s", exc, exc_info=True)

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

    def gerar(self) -> None:
        """Validate input and start presentation generation in a worker thread."""
        if self._gerando:
            return

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
            messagebox.showerror("Texto inválido", str(exc), parent=self)
            self.scripture_text.focus_set()
            return

        titulo_leitura = self.titulo_var.get().strip()
        if not titulo_leitura:
            logger.warning("Validation failed — missing title")
            self._set_progresso_imediato(0)
            self._set_status_erro("Informe o título da leitura.")
            messagebox.showerror("Título obrigatório", "Digite o título da leitura antes de gerar a apresentação.\n\nExemplo: Salmo da Misericórdia", parent=self)
            self.entry_titulo.focus_set()
            return

        total_versiculos = len(leitura.versiculos)
        verso_de: int | None = None
        verso_ate: int | None = None
        raw_de = self.verso_de_var.get().strip()
        raw_ate = self.verso_ate_var.get().strip()

        if raw_de or raw_ate:
            if raw_de:
                if not raw_de.isdigit():
                    self._set_status_erro("O campo 'De' deve ser um número inteiro.")
                    messagebox.showerror("Intervalo inválido", f"O campo 'De' contém um valor inválido: \"{raw_de}\"\n\nDigite apenas um número inteiro positivo.", parent=self)
                    self.entry_verso_de.focus_set()
                    return
                verso_de = int(raw_de)
                if verso_de < 1:
                    self._set_status_erro("O campo 'De' deve ser pelo menos 1.")
                    messagebox.showerror("Intervalo inválido", "O versículo inicial ('De') deve ser pelo menos 1.", parent=self)
                    self.entry_verso_de.focus_set()
                    return
                if verso_de > total_versiculos:
                    self._set_status_erro(f"O campo 'De' ({verso_de}) excede o total de versículos ({total_versiculos}).")
                    messagebox.showerror("Intervalo inválido", f"O versículo inicial ({verso_de}) é maior que o total de\nversículos no texto ({total_versiculos}).", parent=self)
                    self.entry_verso_de.focus_set()
                    return

            if raw_ate:
                if not raw_ate.isdigit():
                    self._set_status_erro("O campo 'Até' deve ser um número inteiro.")
                    messagebox.showerror("Intervalo inválido", f"O campo 'Até' contém um valor inválido: \"{raw_ate}\"\n\nDigite apenas um número inteiro positivo.", parent=self)
                    self.entry_verso_ate.focus_set()
                    return
                verso_ate = int(raw_ate)
                if verso_ate < 1:
                    self._set_status_erro("O campo 'Até' deve ser pelo menos 1.")
                    messagebox.showerror("Intervalo inválido", "O versículo final ('Até') deve ser pelo menos 1.", parent=self)
                    self.entry_verso_ate.focus_set()
                    return
                if verso_ate > total_versiculos:
                    self._set_status_erro(f"O campo 'Até' ({verso_ate}) excede o total de versículos ({total_versiculos}).")
                    messagebox.showerror("Intervalo inválido", f"O versículo final ({verso_ate}) é maior que o total de\nversículos no texto ({total_versiculos}).\n\nO texto tem {total_versiculos} versículo(s).", parent=self)
                    self.entry_verso_ate.focus_set()
                    return

            ef_de = verso_de if verso_de is not None else 1
            ef_ate = verso_ate if verso_ate is not None else total_versiculos
            if ef_de > ef_ate:
                self._set_status_erro(f"Intervalo inválido: 'De' ({ef_de}) é maior que 'Até' ({ef_ate}).")
                messagebox.showerror("Intervalo inválido", f"O versículo inicial ({ef_de}) não pode ser maior\nque o versículo final ({ef_ate}).", parent=self)
                self.entry_verso_de.focus_set()
                return

        caminho_saida = caminho_saida_para(leitura, self._config)
        substituir_existente = False
        if caminho_saida.exists():
            substituir_existente = messagebox.askyesno("Substituir apresentação?", f"Já existe uma apresentação com este nome:\n\n    {caminho_saida.name}\n\nDeseja substituí-la?", icon="warning", parent=self)
            if not substituir_existente:
                logger.info("Generation cancelled — user kept existing file: %s", caminho_saida.name)
                self._set_status("Geração cancelada. Arquivo existente mantido.", FG_HINT)
                return

        logger.info("Generation requested | %s | verses=%d | title=%r | range=[%s..%s]", leitura.livro_capitulo, len(leitura.versiculos), titulo_leitura, verso_de or "início", verso_ate or "fim")
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
            except (OSError, RuntimeError, ValueError, TypeError, AttributeError):
                logger.exception("Unexpected error in generation worker")
                self.after(0, self._on_erro_inesperado)
            else:
                self.after(0, self._on_sucesso, resultado)

        threading.Thread(target=_worker, daemon=True, name="pptx-worker").start()

    def _on_sucesso(self, resultado: ResultadoGeracao) -> None:
        self._ultimo_resultado = resultado
        caminho = resultado.caminho
        verses = resultado.quantidade_versiculos
        secs = resultado.duracao_segundos
        self._set_progresso(100)
        self._set_status_ok(f"{caminho.name} — {verses} versículos — {secs:.1f}s")
        self._gerando = False
        self._set_ui_gerando(False)

        try:
            subprocess.Popen(f'explorer.exe /select,"{caminho}"')
            logger.info("Explorer opened at %s", caminho)
        except (OSError, ValueError):
            logger.exception("Could not open Explorer")

        messagebox.showinfo("Apresentação gerada!", f"Arquivo:      {caminho.name}\nVersículos:   {verses}\nTempo:         {secs:.1f} segundos\nPasta:          {caminho.parent}", parent=self)
        self._limpar_inputs_apos_sucesso()

    def _on_erro(self, mensagem: str) -> None:
        self._set_progresso_imediato(0)
        self._set_status_erro("Não foi possível gerar a apresentação.")
        self._gerando = False
        self._set_ui_gerando(False)
        messagebox.showerror("Erro ao gerar PowerPoint", mensagem, parent=self)
        self.scripture_text.focus_set()

    def _on_erro_inesperado(self) -> None:
        self._set_progresso_imediato(0)
        self._set_status_erro("Ocorreu um erro inesperado.")
        self._gerando = False
        self._set_ui_gerando(False)
        messagebox.showerror("Erro inesperado", "Ocorreu um erro inesperado ao gerar a apresentação.\n\nOs detalhes foram gravados em:\n" + str(LOG_FILE), parent=self)
        self.scripture_text.focus_set()

    def _wire_fluent_animations(self) -> None:
        self.after(60, self.scale_in_window)
        delays = [0, 50, 110, 190, 240, 290]
        entrance_targets = [
            (self._header_ref, {"pady": (36, 0), "padx": 36}, delays[0], 14),
            (self._div_ref, {"pady": 22, "padx": 36}, delays[1], 0),
            (self._card_ref, {"pady": (0, 12), "padx": 32}, delays[2], 14),
            (self._range_outer_ref, {"pady": (0, 26)}, delays[3], 10),
            (self._progress_frame_ref, {}, delays[4], 10),
            (self._btn_row_ref, {"pady": (24, 0)}, delays[5], 10),
        ]
        for idx, (widget, kw, delay, slide_x) in enumerate(entrance_targets):
            if widget is not None:
                self.entrance_animation(widget, kw, delay_ms=delay, slide_x_px=slide_x, on_done=self._mark_startup_ready if idx == len(entrance_targets) - 1 else None)

        self.attach_reveal(self._card_ref)
        self.attach_press_animation(self.botao_gerar)
        self.attach_hover_lift(self.botao_gerar, lift_px=3)
        self.attach_hover_lift(self._range_toggle_btn, lift_px=2)
        for entry in (self.entry_titulo, self.entry_verso_de, self.entry_verso_ate):
            self.attach_acrylic_focus(entry)
        self.attach_focus_reveal(self.scripture_text, normal_border=INPUT_BORDER, reveal_border=INPUT_BORDER_FOC, duration_ms=200)
        if self._card_ref is not None:
            self._attach_floating_shadow(self._card_ref, self._card_ref.master)

    def _status_transition(self, new_text: str, new_color: str) -> None:
        raw = self.status_lbl.cget("text_color")
        if isinstance(raw, (list, tuple)):
            current_color = str(raw[1]) if len(raw) > 1 else str(raw[0])
        else:
            current_color = str(raw)
        if not current_color.startswith("#") or len(current_color) not in (4, 7):
            current_color = BG_CARD

        def _fade_out_done() -> None:
            self.status_var.set(new_text)
            self.fade_widget(self.status_lbl, BG_CARD, new_color, duration_ms=200, attr="text_color")

        self.fade_widget(self.status_lbl, current_color, BG_CARD, duration_ms=200, attr="text_color", on_done=_fade_out_done)

    def _progress_connected_entrance(self) -> None:
        def _update(t: float) -> None:
            valor = 10 * t
            try:
                self._ctk_progress.set(valor / 100.0)
                self._ctk_progress.configure(progress_color=GRAD_START)
            except (tk.TclError, AttributeError, ValueError, TypeError) as exc:
                logger.debug("Ignored UI operation error: %s", exc, exc_info=True)

        def _finish() -> None:
            try:
                self._ctk_progress.set(0.1)
                self._ctk_progress.configure(progress_color=GRAD_START)
            except (tk.TclError, AttributeError, ValueError, TypeError) as exc:
                logger.debug("Ignored UI operation error: %s", exc, exc_info=True)

        self.tween(300, _update, ease=self.ease_spring, on_done=_finish)

    def _animate_range_show(self) -> None:
        try:
            self._range_frame.grid(row=1, column=0, sticky="ew", padx=14, pady=(0, 8))
        except (tk.TclError, AttributeError, ValueError, TypeError):
            return

        def _update(t: float) -> None:
            try:
                for child in self._range_frame.winfo_children():
                    try:
                        child.configure(text_color=self.blend_hex(BG_TILE, FG_LABEL, t))
                    except (tk.TclError, AttributeError, ValueError, TypeError) as exc:
                        logger.debug("Ignored UI operation error: %s", exc, exc_info=True)
            except (tk.TclError, AttributeError, ValueError, TypeError) as exc:
                logger.debug("Ignored UI operation error: %s", exc, exc_info=True)

        self.tween(220, _update, ease=self.ease_out_cubic)

    def _animate_range_hide(self, on_done=None) -> None:
        def _update(t: float) -> None:
            try:
                for child in self._range_frame.winfo_children():
                    try:
                        child.configure(text_color=self.blend_hex(FG_LABEL, BG_TILE, t))
                    except (tk.TclError, AttributeError, ValueError, TypeError) as exc:
                        logger.debug("Ignored UI operation error: %s", exc, exc_info=True)
            except (tk.TclError, AttributeError, ValueError, TypeError) as exc:
                logger.debug("Ignored UI operation error: %s", exc, exc_info=True)

        def _done() -> None:
            try:
                self._range_frame.grid_remove()
            except (tk.TclError, AttributeError, ValueError, TypeError) as exc:
                logger.debug("Ignored UI operation error: %s", exc, exc_info=True)
            if on_done:
                on_done()

        self.tween(220, _update, ease=self.ease_in_cubic, on_done=_done)

    def _elevation_pulse(self, widget, duration_ms: int = 400) -> None:
        half = duration_ms // 2

        def _rise(t: float) -> None:
            c = self.blend_hex(CARD_BORDER, GRAD_START, t)
            try:
                widget.configure(border_color=c)
            except (tk.TclError, AttributeError, ValueError, TypeError) as exc:
                logger.debug("Ignored UI operation error: %s", exc, exc_info=True)

        def _fall_start() -> None:
            def _fall(t: float) -> None:
                c = self.blend_hex(GRAD_START, CARD_BORDER, t)
                try:
                    widget.configure(border_color=c)
                except (tk.TclError, AttributeError, ValueError, TypeError) as exc:
                    logger.debug("Ignored UI operation error: %s", exc, exc_info=True)

            self.tween(half, _fall, ease=self.ease_out_cubic)

        self.tween(half, _rise, ease=self.ease_out_cubic, on_done=_fall_start)

    _card_opacity_token: list[bool] = [False]

    def _card_opacity_transition(self, widget, dim: bool) -> None:
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
        except (tk.TclError, AttributeError, ValueError, TypeError):
            from_color = CARD_BORDER

        dim_color = BG_CARD_INNER
        to_color = dim_color if dim else CARD_BORDER
        self.fade_widget(widget, from_color, to_color, duration_ms=300 if dim else 400, attr="border_color", cancel_token=self._card_opacity_token)

    def _make_vertical_gradient_image(self, width: int, height: int, top_rgb: tuple[int, int, int], bottom_rgb: tuple[int, int, int], radius: int = 0):
        if not (PIL_AVAILABLE and ctk.CTkImage is not None):
            return None
        try:
            width = max(1, int(width))
            height = max(1, int(height))
            scale = 4
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
                ImageDraw.Draw(mask).rounded_rectangle((0, 0, w - 1, h - 1), radius=radius * scale, fill=255)
                img.putalpha(mask)
            img = img.resize((width, height), Image.LANCZOS)
            return ctk.CTkImage(light_image=img, dark_image=img, size=(width, height))
        except (OSError, ValueError, RuntimeError, TypeError):
            logger.exception("Could not render gradient image")
            return None

    def tween(self, duration_ms: int, on_update, ease=None, on_done=None, fps: int = 60, cancel_token: list | None = None) -> None:
        """Proxy to animation tween helper.

        Args:
            duration_ms: Total animation duration in milliseconds.
            on_update: Callback receiving eased progress.
            ease: Optional easing function.
            on_done: Optional callback invoked when animation ends.
            fps: Target frames per second.
            cancel_token: Optional mutable cancellation token.
        """
        return FluentAnimations.tween(self, duration_ms, on_update, ease=ease, on_done=on_done, fps=fps, cancel_token=cancel_token)

    def blend_hex(self, c1: str, c2: str, t: float) -> str:
        """Proxy to color interpolation helper.

        Args:
            c1: Start color in ``#RRGGBB`` format.
            c2: End color in ``#RRGGBB`` format.
            t: Blend factor between 0 and 1.

        Returns:
            str: Interpolated color in ``#RRGGBB`` format.
        """
        return FluentAnimations.blend_hex(c1, c2, t)

    def ease_out_cubic(self, t: float) -> float:
        """Proxy to ease-out cubic curve.

        Args:
            t: Normalized progress between 0 and 1.

        Returns:
            float: Eased progress value.
        """
        return FluentAnimations.ease_out_cubic(t)

    def ease_in_cubic(self, t: float) -> float:
        """Proxy to ease-in cubic curve.

        Args:
            t: Normalized progress between 0 and 1.

        Returns:
            float: Eased progress value.
        """
        return FluentAnimations.ease_in_cubic(t)

    def ease_in_out_quint(self, t: float) -> float:
        """Proxy to ease-in-out quintic curve.

        Args:
            t: Normalized progress between 0 and 1.

        Returns:
            float: Eased progress value.
        """
        return FluentAnimations.ease_in_out_quint(t)

    def ease_spring(self, t: float) -> float:
        """Proxy to spring-like easing curve.

        Args:
            t: Normalized progress between 0 and 1.

        Returns:
            float: Eased progress value.
        """
        return FluentAnimations.ease_spring(t)

    def fade_widget(self, widget, from_color: str, to_color: str, duration_ms: int = 220, attr: str = "text_color", on_done=None, cancel_token: list | None = None) -> None:
        """Proxy to widget color fade helper.

        Args:
            widget: Target widget.
            from_color: Start color.
            to_color: End color.
            duration_ms: Animation duration in milliseconds.
            attr: Config attribute to animate.
            on_done: Optional completion callback.
            cancel_token: Optional mutable cancellation token.
        """
        return FluentAnimations.fade_widget(self, widget, from_color, to_color, duration_ms=duration_ms, attr=attr, on_done=on_done, cancel_token=cancel_token)

    def entrance_animation(self, widget, parent_grid_kw: dict[str, Any], delay_ms: int = 0, slide_px: int = 18, slide_x_px: int = 0, ease=None, on_done=None) -> None:
        """Proxy to grid entrance animation helper.

        Args:
            widget: Target grid widget.
            parent_grid_kw: Grid padding metadata.
            delay_ms: Delay before animation start.
            slide_px: Vertical offset in pixels.
            slide_x_px: Horizontal offset in pixels.
            ease: Optional easing function.
            on_done: Optional completion callback.
        """
        return FluentAnimations.entrance_animation(self, widget, parent_grid_kw, delay_ms=delay_ms, slide_px=slide_px, slide_x_px=slide_x_px, ease=ease, on_done=on_done)

    def scale_in_window(self) -> None:
        """Proxy to top-level scale-in animation."""
        return FluentAnimations.scale_in_window(self)

    def attach_reveal(self, widget, normal_border: str = CARD_BORDER, reveal_border: str = GRAD_START, duration_ms: int = 180) -> None:
        """Proxy to hover reveal border animation.

        Args:
            widget: Target widget.
            normal_border: Default border color.
            reveal_border: Hover border color.
            duration_ms: Transition duration in milliseconds.
        """
        return FluentAnimations.attach_reveal(self, widget, normal_border=normal_border, reveal_border=reveal_border, duration_ms=duration_ms)

    def attach_focus_reveal(self, widget, normal_border: str = CARD_BORDER, reveal_border: str = GRAD_START, duration_ms: int = 200) -> None:
        """Proxy to focus border reveal animation.

        Args:
            widget: Target widget.
            normal_border: Default border color.
            reveal_border: Focus border color.
            duration_ms: Transition duration in milliseconds.
        """
        return FluentAnimations.attach_focus_reveal(self, widget, normal_border=normal_border, reveal_border=reveal_border, duration_ms=duration_ms)

    def attach_hover_lift(self, widget, lift_px: int = 3, duration_ms: int = 140) -> None:
        """Proxy to hover lift motion animation.

        Args:
            widget: Target widget.
            lift_px: Lift amount in pixels.
            duration_ms: Transition duration in milliseconds.
        """
        return FluentAnimations.attach_hover_lift(self, widget, lift_px=lift_px, duration_ms=duration_ms)

    def attach_press_animation(self, widget) -> None:
        """Proxy to press/release button feedback animation.

        Args:
            widget: Target widget.
        """
        return FluentAnimations.attach_press_animation(self, widget)

    def attach_acrylic_focus(self, widget) -> None:
        """Proxy to acrylic-like focus border animation.

        Args:
            widget: Target widget.
        """
        return FluentAnimations.attach_acrylic_focus(self, widget)

    def _attach_floating_shadow(self, widget, parent, radius: int = 26, blur: int = 18, alpha: int = 130, offset: tuple[int, int] = (0, 10)) -> None:
        return FluentAnimations._attach_floating_shadow(self, widget, parent, radius=radius, blur=blur, alpha=alpha, offset=offset)


