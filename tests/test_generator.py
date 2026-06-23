from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace
from typing import Any

import pytest

from core.utils import GeradorSalmosError
import core.generator as generator


class _ParagraphFormat:
    def __init__(self) -> None:
        self.Alignment = None


class _TextRange:
    def __init__(self, text: str) -> None:
        self.Text = text
        self.ParagraphFormat = _ParagraphFormat()


class _TextFrame:
    def __init__(self, text: str) -> None:
        self.HasText = True
        self.TextRange = _TextRange(text)


class _Shape:
    def __init__(
        self,
        text: str = "",
        has_text: bool = True,
        shape_type: int = 1,
        group_items: _GroupItems | None = None,
    ) -> None:
        self.Type = shape_type
        self.HasTextFrame = has_text
        self.TextFrame = _TextFrame(text)
        self.TextFrame2 = _TextFrame(text)
        self.GroupItems = group_items
        self.Left = 0
        self.Top = 0
        self.Width = 100
        self.Height = 20


class _GroupItems:
    def __init__(self, items: list[_Shape]) -> None:
        self._items = list(items)
        self.Count = len(self._items)

    def __call__(self, index: int) -> _Shape:
        return self._items[index - 1]


class _PageSetup:
    SlideWidth = 800
    SlideHeight = 600


class _SlideParent:
    PageSetup = _PageSetup()


class _SlideForHelpers:
    def __init__(self, shapes: list[_Shape]) -> None:
        self.Shapes = shapes
        self.Parent = _SlideParent()


def test_normalizar_placeholder_removes_linebreaks_dashes_and_normalizes_spaces() -> None:
    """Verify placeholder normalization removes formatting differences."""
    raw = "  VeRsÍculo\n  com\r   espaço — e – traço  "
    normalized = generator.normalizar_placeholder(raw)
    assert normalized == "versículo com espaço - e - traço"


def test_escolher_indice_template_selects_final_even_and_odd() -> None:
    """Verify template selection rules for final, even, and odd verses."""
    assert generator.escolher_indice_template(4, 5) == generator.TEMPLATE_FINAL_INDEX
    assert generator.escolher_indice_template(0, 5) == generator.TEMPLATE_PASTOR_INDEX
    assert generator.escolher_indice_template(1, 5) == generator.TEMPLATE_IGREJA_INDEX


def test_iterar_formas_walks_nested_groups() -> None:
    """Verify grouped PowerPoint shapes are traversed recursively."""
    nested = _Shape("nested")
    group = _Shape(shape_type=6, group_items=_GroupItems([nested]))
    root = _Shape("root")

    found = list(generator.iterar_formas(root)) + list(generator.iterar_formas(group))
    assert root in found
    assert nested in found


def test_justificar_texto_sets_alignment_on_textframes() -> None:
    """Verify text justification is applied to available text frames."""
    shape = _Shape("texto")
    generator.justificar_texto(shape)
    assert shape.TextFrame.TextRange.ParagraphFormat.Alignment == generator.PP_ALIGN_JUSTIFY
    assert shape.TextFrame2.TextRange.ParagraphFormat.Alignment == generator.PP_ALIGN_JUSTIFY


def test_substituir_placeholder_exact_match_changes_text() -> None:
    """Verify exact placeholder matches are replaced."""
    shape = _Shape("{titulo}")
    slide = _SlideForHelpers([shape])

    changed = generator.substituir_placeholder(slide, ("{titulo}",), "Novo Titulo")

    assert changed is True
    assert shape.TextFrame.TextRange.Text == "Novo Titulo"


def test_substituir_placeholder_partial_match_replaces_in_text() -> None:
    """Verify placeholders embedded in larger text are replaced."""
    shape = _Shape("Leitura: {titulo}")
    slide = _SlideForHelpers([shape])

    changed = generator.substituir_placeholder(slide, ("{titulo}",), "Salmo")

    assert changed is True
    assert shape.TextFrame.TextRange.Text == "Leitura: Salmo"


def test_substituir_placeholder_with_center_and_justify() -> None:
    """Verify replacement can also center and justify updated content."""
    shape = _Shape("{versiculo}")
    slide = _SlideForHelpers([shape])

    changed = generator.substituir_placeholder(
        slide,
        ("{versiculo}",),
        "1 - Texto",
        centralizar_caixa=True,
        justificar=True,
    )

    assert changed is True
    assert shape.Left == 350
    assert shape.Top == 290
    assert shape.TextFrame.TextRange.ParagraphFormat.Alignment == generator.PP_ALIGN_JUSTIFY


def test_exigir_placeholder_raises_when_not_found() -> None:
    """Verify missing placeholders raise a user-facing domain error."""
    slide = _SlideForHelpers([_Shape("Sem placeholder")])
    with pytest.raises(GeradorSalmosError, match="Placeholder '{titulo}' não encontrado"):
        generator.exigir_placeholder(slide, ("{titulo}",), "Novo", "slide", "{titulo}")


def test_validar_modelo_missing_file_raises(tmp_path: Path) -> None:
    """Verify model validation fails for missing files."""
    missing = tmp_path / "nao_existe.pptx"
    with pytest.raises(GeradorSalmosError, match="modelo.pptx não encontrado"):
        generator.validar_modelo(missing)


def test_validar_modelo_path_is_not_file_raises(tmp_path: Path) -> None:
    """Verify model validation fails for directory paths."""
    folder = tmp_path / "pasta"
    folder.mkdir()
    with pytest.raises(GeradorSalmosError, match="não aponta para um arquivo válido"):
        generator.validar_modelo(folder)


def test_abrir_powerpoint_without_pywin32_raises(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify opening PowerPoint fails with a friendly error when pywin32 is unavailable."""
    monkeypatch.setattr(generator, "win32com", None)
    monkeypatch.setattr(generator, "pythoncom", None)
    with pytest.raises(GeradorSalmosError, match="pywin32"):
        generator.abrir_powerpoint()


def test_abrir_powerpoint_dispatch_error_raises_friendly_message(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify COM dispatch failures are wrapped in a friendly error message."""
    class _FailClient:
        @staticmethod
        def DispatchEx(_name: str) -> Any:
            raise RuntimeError("boom")

    monkeypatch.setattr(generator, "pythoncom", object())
    monkeypatch.setattr(generator, "win32com", SimpleNamespace(client=_FailClient()))

    with pytest.raises(GeradorSalmosError, match="Não foi possível abrir o Microsoft PowerPoint"):
        generator.abrir_powerpoint()


def test_abrir_powerpoint_success_sets_visible(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify PowerPoint is made visible after successful startup."""
    class _App:
        def __init__(self) -> None:
            self.Visible = False

    app = _App()

    class _Client:
        @staticmethod
        def DispatchEx(_name: str) -> _App:
            return app

    monkeypatch.setattr(generator, "pythoncom", object())
    monkeypatch.setattr(generator, "win32com", SimpleNamespace(client=_Client()))

    returned = generator.abrir_powerpoint()
    assert returned is app
    assert app.Visible is True


def test_gerar_powerpoint_requires_title(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify generation requires a non-empty title."""
    monkeypatch.setattr(generator, "validar_modelo", lambda _p: None)
    with pytest.raises(GeradorSalmosError, match="Informe o título"):
        generator.gerar_powerpoint("Livro: Salmos\nCapítulo: 23\n1 A", "   ")


def test_gerar_powerpoint_empty_selected_range_raises(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify generation fails when the selected verse range is empty."""
    class _Leitura:
        livro = "Salmos"
        capitulo = "23"
        versiculos = ["1 - A", "2 - B"]

    monkeypatch.setattr(generator, "validar_modelo", lambda _p: None)
    monkeypatch.setattr(generator, "ler_texto_entrada", lambda _t: _Leitura())

    with pytest.raises(GeradorSalmosError, match="não contém versículos"):
        generator.gerar_powerpoint(
            "dummy",
            "Titulo",
            verso_inicio=3,
            verso_fim=2,
        )


def test_gerar_powerpoint_existing_file_without_replace_raises(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    """Verify generation blocks overwrite when replacement is not confirmed."""
    class _Leitura:
        livro = "Salmos"
        capitulo = "23"
        versiculos = ["1 - A"]

    output_file = tmp_path / "saida.pptx"
    output_file.write_text("x", encoding="utf-8")

    class _Slide:
        def __init__(self) -> None:
            self.Shapes = []

        def Duplicate(self) -> SimpleNamespace:
            return SimpleNamespace(Item=lambda _i: SimpleNamespace(MoveTo=lambda _n: None))

        def Delete(self):
            return None

    class _Slides:
        Count = 4

        def __call__(self, _index: int) -> _Slide:
            return _Slide()

    class _Presentation:
        def __init__(self) -> None:
            self.Slides = _Slides()
            self.Saved = False

        def SaveAs(self, _path: str, FileFormat: int) -> None:
            return None

        def Close(self) -> None:
            return None

    class _Presentations:
        def Open(self, *_args: Any, **_kwargs: Any) -> _Presentation:
            return _Presentation()

    class _PowerPoint:
        def __init__(self) -> None:
            self.Presentations = _Presentations()

        def Quit(self) -> None:
            return None

    monkeypatch.setattr(generator, "validar_modelo", lambda _p: None)
    monkeypatch.setattr(generator, "ler_texto_entrada", lambda _t: _Leitura())
    monkeypatch.setattr(generator, "caminho_saida_para", lambda _l, _c: output_file)
    monkeypatch.setattr(generator, "abrir_powerpoint", lambda: _PowerPoint())
    monkeypatch.setattr(generator, "exigir_placeholder", lambda *_a, **_k: None)
    monkeypatch.setattr(generator, "pythoncom", SimpleNamespace(CoInitialize=lambda: None, CoUninitialize=lambda: None))
    monkeypatch.setattr(generator, "win32com", object())

    with pytest.raises(GeradorSalmosError, match="O arquivo já existe"):
        generator.gerar_powerpoint("dummy", "Titulo", substituir_existente=False)


def test_gerar_powerpoint_wraps_unexpected_com_error(monkeypatch: pytest.MonkeyPatch) -> None:
    """Verify COM-level errors are converted into user-friendly domain errors."""
    class _Leitura:
        livro = "Salmos"
        capitulo = "23"
        versiculos = ["1 - A"]

    class _ComError(Exception):
        pass

    class _FakePywinTypes:
        com_error = _ComError

    def _boom_open() -> None:
        raise _ComError("com boom")

    monkeypatch.setattr(generator, "validar_modelo", lambda _p: None)
    monkeypatch.setattr(generator, "ler_texto_entrada", lambda _t: _Leitura())
    monkeypatch.setattr(generator, "abrir_powerpoint", _boom_open)
    monkeypatch.setattr(generator, "pythoncom", SimpleNamespace(CoInitialize=lambda: None, CoUninitialize=lambda: None))
    monkeypatch.setattr(generator, "win32com", object())
    monkeypatch.setattr(generator, "pywintypes", _FakePywinTypes())

    with pytest.raises(GeradorSalmosError, match="PowerPoint retornou um erro"):
        generator.gerar_powerpoint("dummy", "Titulo")


def test_gerar_powerpoint_wraps_permission_error_on_save(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    """Verify save permission failures are converted into user-facing errors."""
    class _Leitura:
        livro = "Salmos"
        capitulo = "23"
        versiculos = ["1 - A"]

    out = tmp_path / "saida.pptx"

    class _Slide:
        def __init__(self) -> None:
            self.Shapes = []

        def Duplicate(self) -> SimpleNamespace:
            return SimpleNamespace(Item=lambda _i: SimpleNamespace(MoveTo=lambda _n: None))

        def Delete(self):
            return None

    class _Slides:
        Count = 4

        def __call__(self, _index: int) -> _Slide:
            return _Slide()

    class _Presentation:
        def __init__(self) -> None:
            self.Slides = _Slides()
            self.Saved = False

        def SaveAs(self, _path: str, FileFormat: int) -> None:
            raise PermissionError("denied")

        def Close(self) -> None:
            return None

    class _Presentations:
        def Open(self, *_args: Any, **_kwargs: Any) -> _Presentation:
            return _Presentation()

    class _PowerPoint:
        def __init__(self) -> None:
            self.Presentations = _Presentations()

        def Quit(self) -> None:
            return None

    monkeypatch.setattr(generator, "validar_modelo", lambda _p: None)
    monkeypatch.setattr(generator, "ler_texto_entrada", lambda _t: _Leitura())
    monkeypatch.setattr(generator, "caminho_saida_para", lambda _l, _c: out)
    monkeypatch.setattr(generator, "abrir_powerpoint", lambda: _PowerPoint())
    monkeypatch.setattr(generator, "exigir_placeholder", lambda *_a, **_k: None)
    monkeypatch.setattr(generator, "pythoncom", SimpleNamespace(CoInitialize=lambda: None, CoUninitialize=lambda: None))
    monkeypatch.setattr(generator, "win32com", object())

    with pytest.raises(GeradorSalmosError, match="Erro de permissão ao salvar o arquivo"):
        generator.gerar_powerpoint("dummy", "Titulo", substituir_existente=True)


def test_gerar_powerpoint_invalid_output_target_type_raises(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    """Verify generation rejects output targets that are directories."""

    class _Leitura:
        livro = "Salmos"
        capitulo = "23"
        versiculos = ["1 - A"]

    destino = tmp_path / "saida.pptx"
    destino.mkdir()

    monkeypatch.setattr(generator, "validar_modelo", lambda _p: None)
    monkeypatch.setattr(generator, "ler_texto_entrada", lambda _t: _Leitura())
    monkeypatch.setattr(generator, "caminho_saida_para", lambda _l, _c: destino)

    with pytest.raises(GeradorSalmosError, match="não aponta para um arquivo válido"):
        generator.gerar_powerpoint("dummy", "Titulo")
