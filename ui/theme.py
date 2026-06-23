from __future__ import annotations

from pathlib import Path

APP_NAME = "Gerador de Leitura Bíblica"
APP_VERSION = "2.0"
MODEL_NAME = "modelo.pptx"
APP_AUTHOR = "Igreja Tabernáculo do Senhor - Barra do Garças"

OUTPUT_DIR_DEFAULT = Path.home() / "Documents" / "LeituraBiblica"

TEMPLATE_TITLE_INDEX = 1
TEMPLATE_PASTOR_INDEX = 2
TEMPLATE_IGREJA_INDEX = 3
TEMPLATE_FINAL_INDEX = 4

PP_ALIGN_JUSTIFY = 4
PP_SAVE_AS_OPENXML = 24

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

BG_APP = "#0a0e1a"
BG_CARD = "#111827"
BG_CARD_INNER = "#0d1520"
BG_TILE = "#162033"
BG_INPUT = "#0f1a2e"

GRAD_START = "#60a5fa"
GRAD_MID = "#3b82f6"

CARD_BORDER = "#1d2d4a"
INPUT_BORDER = "#1d2d4a"
INPUT_BORDER_FOC = "#3b82f6"

FG_TITLE = "#eaf0ff"
FG_LABEL = "#7f9ab3"
FG_INPUT = "#d9eaff"
FG_HINT = "#5a7a9a"
FG_SUCCESS = "#34d399"
FG_ERROR = "#f87171"

FONT_FAMILY = "Segoe UI"
RADIUS = 22
RADIUS_TILE = 18
RADIUS_PILL = 20

ICON_BOOK = "📖"
ICON_PLAY = "▶"
ICON_CHECK = "✓"
ICON_CROSS = "✗"
ICON_FOLDER = "📂"
