from __future__ import annotations

import logging
import math
import tkinter as tk
from typing import Any

logger = logging.getLogger(__name__)

from ui.theme import (
    CARD_BORDER,
    GRAD_MID,
    GRAD_START,
    INPUT_BORDER,
    INPUT_BORDER_FOC,
)


class FluentAnimations:
    """Provide reusable animation helpers for CustomTkinter widgets."""

    @staticmethod
    def ease_out_cubic(t: float) -> float:
        """Apply an ease-out cubic curve.

        Args:
            t: Normalized progress between 0 and 1.

        Returns:
            float: Eased progress value.
        """
        return 1 - (1 - t) ** 3

    @staticmethod
    def ease_in_cubic(t: float) -> float:
        """Apply an ease-in cubic curve.

        Args:
            t: Normalized progress between 0 and 1.

        Returns:
            float: Eased progress value.
        """
        return t ** 3

    @staticmethod
    def ease_in_out_quint(t: float) -> float:
        """Apply an ease-in-out quintic curve.

        Args:
            t: Normalized progress between 0 and 1.

        Returns:
            float: Eased progress value.
        """
        if t < 0.5:
            return 16 * t ** 5
        return 1 - (-2 * t + 2) ** 5 / 2

    @staticmethod
    def ease_spring(t: float) -> float:
        """Apply a spring-like easing curve.

        Args:
            t: Normalized progress between 0 and 1.

        Returns:
            float: Eased progress value with overshoot feel.
        """
        c4 = (2 * math.pi) / 3
        if t == 0:
            return 0
        if t == 1:
            return 1
        return 2 ** (-8 * t) * math.sin((t * 10 - 0.75) * c4) + 1

    @staticmethod
    def blend_hex(c1: str, c2: str, t: float) -> str:
        """Interpolate two hex colors.

        Args:
            c1: Start color in ``#RRGGBB`` format.
            c2: End color in ``#RRGGBB`` format.
            t: Blend factor between 0 and 1.

        Returns:
            str: Interpolated color in ``#RRGGBB`` format.
        """
        r1, g1, b1 = int(c1[1:3], 16), int(c1[3:5], 16), int(c1[5:7], 16)
        r2, g2, b2 = int(c2[1:3], 16), int(c2[3:5], 16), int(c2[5:7], 16)
        r = int(r1 + (r2 - r1) * t)
        g = int(g1 + (g2 - g1) * t)
        b = int(b1 + (b2 - b1) * t)
        return f"#{r:02x}{g:02x}{b:02x}"

    def tween(
        self,
        duration_ms: int,
        on_update,
        ease=None,
        on_done=None,
        fps: int = 60,
        cancel_token: list | None = None,
    ) -> None:
        """Animate a numeric transition by repeatedly calling ``on_update``.

        Args:
            duration_ms: Total animation duration in milliseconds.
            on_update: Callback receiving eased progress.
            ease: Optional easing function.
            on_done: Optional callback invoked when animation ends.
            fps: Target frames per second.
            cancel_token: Mutable boolean token used to cancel active animation.
        """
        if ease is None:
            ease = self.ease_out_cubic
        interval = max(1, 1000 // fps)
        steps = max(1, duration_ms // interval)
        step_ref = [0]

        def _tick() -> None:
            if cancel_token is not None and cancel_token[0]:
                return
            step_ref[0] += 1
            raw = step_ref[0] / steps
            t = min(raw, 1.0)
            try:
                on_update(ease(t))
            except (tk.TclError, AttributeError, ValueError, TypeError) as exc:
                logger.debug("Ignored animation operation error: %s", exc, exc_info=True)
            if t < 1.0:
                self.after(interval, _tick)
            elif on_done:
                try:
                    on_done()
                except (tk.TclError, AttributeError, ValueError, TypeError) as exc:
                    logger.debug("Ignored animation operation error: %s", exc, exc_info=True)

        self.after(0, _tick)

    def fade_widget(
        self,
        widget,
        from_color: str,
        to_color: str,
        duration_ms: int = 220,
        attr: str = "text_color",
        on_done=None,
        cancel_token: list | None = None,
    ) -> None:
        """Animate a widget color-like attribute between two colors.

        Args:
            widget: Target widget with ``configure`` support.
            from_color: Start hex color.
            to_color: End hex color.
            duration_ms: Animation duration in milliseconds.
            attr: Widget attribute to animate.
            on_done: Optional callback invoked when animation ends.
            cancel_token: Mutable cancellation token.
        """
        def _update(t: float) -> None:
            try:
                widget.configure(**{attr: self.blend_hex(from_color, to_color, t)})
            except (tk.TclError, AttributeError, ValueError, TypeError) as exc:
                logger.debug("Ignored animation operation error: %s", exc, exc_info=True)

        self.tween(duration_ms, _update, ease=self.ease_in_out_quint, on_done=on_done, cancel_token=cancel_token)

    def entrance_animation(
        self,
        widget,
        parent_grid_kw: dict[str, Any],
        delay_ms: int = 0,
        slide_px: int = 18,
        slide_x_px: int = 0,
        ease=None,
        on_done=None,
    ) -> None:
        """Animate a widget entrance with slide offsets.

        Args:
            widget: Widget configured via ``grid``.
            parent_grid_kw: Original grid padding metadata.
            delay_ms: Initial delay before animation starts.
            slide_px: Vertical slide offset in pixels.
            slide_x_px: Horizontal slide offset in pixels.
            ease: Optional easing function.
            on_done: Optional completion callback.
        """
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

        def _update(t: float) -> None:
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
            except (tk.TclError, AttributeError, ValueError, TypeError) as exc:
                logger.debug("Ignored animation operation error: %s", exc, exc_info=True)

        self.after(delay_ms, lambda: self.tween(260, _update, ease=ease or self.ease_out_cubic, on_done=on_done))

    def scale_in_window(self) -> None:
        """Animate a small scale-in effect for the top-level window."""
        try:
            self.update_idletasks()
            w = self.winfo_width()
            h = self.winfo_height()
            x = self.winfo_x()
            y = self.winfo_y()
        except (tk.TclError, AttributeError, ValueError, TypeError, RuntimeError):
            return

        if w <= 1 or h <= 1:
            self.after(40, self.scale_in_window)
            return

        start_w = int(w * 0.94)
        start_h = int(h * 0.94)

        def _update(t: float) -> None:
            cw = int(start_w + (w - start_w) * t)
            ch = int(start_h + (h - start_h) * t)
            cx = x + (w - cw) // 2
            cy = y + (h - ch) // 2
            try:
                self.geometry(f"{cw}x{ch}+{cx}+{cy}")
            except (tk.TclError, AttributeError, ValueError, TypeError) as exc:
                logger.debug("Ignored animation operation error: %s", exc, exc_info=True)

        self.tween(300, _update, ease=self.ease_spring)

    def attach_reveal(
        self,
        widget,
        normal_border: str = CARD_BORDER,
        reveal_border: str = GRAD_START,
        duration_ms: int = 180,
    ) -> None:
        """Add hover border reveal animation to a widget.

        Args:
            widget: Target widget.
            normal_border: Default border color.
            reveal_border: Border color used on hover.
            duration_ms: Hover-in animation duration.
        """
        _token: list[bool] = [False]
        _current_border: list[str] = [normal_border]
        _inside: list[bool] = [False]

        def _on_enter(_e) -> None:
            nonlocal _token
            if _inside[0]:
                return
            _inside[0] = True
            _token[0] = True
            _token = [False]
            from_color = _current_border[0]
            _current_border[0] = reveal_border

            def _update(t: float) -> None:
                c = self.blend_hex(from_color, reveal_border, t)
                _current_border[0] = c
                try:
                    widget.configure(border_color=c)
                except (tk.TclError, AttributeError, ValueError, TypeError) as exc:
                    logger.debug("Ignored animation operation error: %s", exc, exc_info=True)

            self.tween(duration_ms, _update, ease=self.ease_out_cubic, cancel_token=_token)

        def _on_leave(_e) -> None:
            nonlocal _token
            if not _inside[0]:
                return
            _inside[0] = False
            _token[0] = True
            _token = [False]
            from_color = _current_border[0]
            _current_border[0] = normal_border

            def _update(t: float) -> None:
                c = self.blend_hex(from_color, normal_border, t)
                _current_border[0] = c
                try:
                    widget.configure(border_color=c)
                except (tk.TclError, AttributeError, ValueError, TypeError) as exc:
                    logger.debug("Ignored animation operation error: %s", exc, exc_info=True)

            self.tween(duration_ms * 2, _update, ease=self.ease_out_cubic, cancel_token=_token)

        try:
            widget.bind("<Enter>", _on_enter, add="+")
            widget.bind("<Leave>", _on_leave, add="+")
        except (tk.TclError, AttributeError, ValueError, TypeError) as exc:
            logger.debug("Ignored animation operation error: %s", exc, exc_info=True)

    def attach_focus_reveal(
        self,
        widget,
        normal_border: str = CARD_BORDER,
        reveal_border: str = GRAD_START,
        duration_ms: int = 200,
    ) -> None:
        """Add border reveal animation for focus transitions.

        Args:
            widget: Target widget.
            normal_border: Default border color.
            reveal_border: Border color used when focused.
            duration_ms: Focus transition duration.
        """
        _token: list[bool] = [False]
        _current: list[str] = [normal_border]

        def _on_focus_in(_e) -> None:
            nonlocal _token
            _token[0] = True
            _token = [False]
            from_color = _current[0]

            def _update(t: float) -> None:
                c = self.blend_hex(from_color, reveal_border, t)
                _current[0] = c
                try:
                    widget.configure(border_color=c)
                except (tk.TclError, AttributeError, ValueError, TypeError) as exc:
                    logger.debug("Ignored animation operation error: %s", exc, exc_info=True)

            self.tween(duration_ms, _update, ease=self.ease_out_cubic, cancel_token=_token)

        def _on_focus_out(_e) -> None:
            nonlocal _token
            _token[0] = True
            _token = [False]
            from_color = _current[0]

            def _update(t: float) -> None:
                c = self.blend_hex(from_color, normal_border, t)
                _current[0] = c
                try:
                    widget.configure(border_color=c)
                except (tk.TclError, AttributeError, ValueError, TypeError) as exc:
                    logger.debug("Ignored animation operation error: %s", exc, exc_info=True)

            self.tween(duration_ms, _update, ease=self.ease_out_cubic, cancel_token=_token)

        try:
            widget.bind("<FocusIn>", _on_focus_in, add="+")
            widget.bind("<FocusOut>", _on_focus_out, add="+")
        except (tk.TclError, AttributeError, ValueError, TypeError) as exc:
            logger.debug("Ignored animation operation error: %s", exc, exc_info=True)

    def attach_hover_lift(
        self,
        widget,
        lift_px: int = 3,
        duration_ms: int = 140,
    ) -> None:
        """Add hover lift motion to a grid-managed widget.

        Args:
            widget: Target widget.
            lift_px: Vertical lift amount in pixels.
            duration_ms: Lift animation duration.
        """
        try:
            info = widget.grid_info()
            pady = info.get("pady", (0, 0))
            if isinstance(pady, int):
                pady = (pady, pady)
            _base_top, _base_bot = int(pady[0]), int(pady[1])
        except (tk.TclError, AttributeError, ValueError, TypeError):
            _base_top, _base_bot = 0, 0

        _token: list[bool] = [False]
        _lifted = [False]

        def _on_enter(_e) -> None:
            nonlocal _token
            if _lifted[0]:
                return
            _lifted[0] = True
            _token[0] = True
            _token = [False]

            def _up(t: float) -> None:
                off = int(lift_px * t)
                try:
                    widget.grid_configure(pady=(max(0, _base_top - off), _base_bot + off))
                except (tk.TclError, AttributeError, ValueError, TypeError) as exc:
                    logger.debug("Ignored animation operation error: %s", exc, exc_info=True)

            self.tween(duration_ms, _up, ease=self.ease_out_cubic, cancel_token=_token)

        def _on_leave(_e) -> None:
            nonlocal _token
            if not _lifted[0]:
                return
            _lifted[0] = False
            _token[0] = True
            _token = [False]

            def _down(t: float) -> None:
                off = int(lift_px * (1 - t))
                try:
                    widget.grid_configure(pady=(max(0, _base_top - off), _base_bot + off))
                except (tk.TclError, AttributeError, ValueError, TypeError) as exc:
                    logger.debug("Ignored animation operation error: %s", exc, exc_info=True)

            self.tween(duration_ms * 2, _down, ease=self.ease_out_cubic, cancel_token=_token)

        def _ignore_motion(_e) -> None:
            return

        try:
            widget.bind("<Enter>", _on_enter, add="+")
            widget.bind("<Leave>", _on_leave, add="+")
            widget.bind("<Motion>", _ignore_motion, add="+")
        except (tk.TclError, AttributeError, ValueError, TypeError) as exc:
            logger.debug("Ignored animation operation error: %s", exc, exc_info=True)

    def attach_press_animation(self, widget) -> None:
        """Add press/release feedback animation to a widget.

        Args:
            widget: Target widget.
        """
        _token: list[bool] = [False]

        try:
            info = widget.grid_info()
            pady = info.get("pady", (0, 0))
            if isinstance(pady, int):
                pady = (pady, pady)
            _base_top, _base_bot = int(pady[0]), int(pady[1])
        except (tk.TclError, AttributeError, ValueError, TypeError):
            _base_top, _base_bot = 0, 0

        def _on_press(_e) -> None:
            try:
                widget.configure(fg_color=GRAD_MID)
                widget.grid_configure(pady=(_base_top + 2, _base_bot + 2))
            except (tk.TclError, AttributeError, ValueError, TypeError) as exc:
                logger.debug("Ignored animation operation error: %s", exc, exc_info=True)

        def _on_release(_e) -> None:
            nonlocal _token
            _token[0] = True
            _token = [False]

            def _restore(t: float) -> None:
                try:
                    c = self.blend_hex(GRAD_MID, GRAD_START, t)
                    widget.configure(fg_color=c)
                    settle = int(2 * (1 - t))
                    widget.grid_configure(pady=(_base_top + settle, _base_bot + settle))
                except (tk.TclError, AttributeError, ValueError, TypeError) as exc:
                    logger.debug("Ignored animation operation error: %s", exc, exc_info=True)

            self.tween(180, _restore, ease=self.ease_spring, cancel_token=_token)

        try:
            widget.bind("<ButtonPress-1>", _on_press, add="+")
            widget.bind("<ButtonRelease-1>", _on_release, add="+")
        except (tk.TclError, AttributeError, ValueError, TypeError) as exc:
            logger.debug("Ignored animation operation error: %s", exc, exc_info=True)

    def attach_acrylic_focus(self, widget) -> None:
        """Animate entry border color for focus changes.

        Args:
            widget: Target widget.
        """
        _token: list[bool] = [False]
        _current: list[str] = [INPUT_BORDER]

        def _on_focus_in(_e) -> None:
            nonlocal _token
            _token[0] = True
            _token = [False]
            from_color = _current[0]

            def _up(t: float) -> None:
                c = self.blend_hex(from_color, INPUT_BORDER_FOC, t)
                _current[0] = c
                try:
                    widget.configure(border_color=c)
                except (tk.TclError, AttributeError, ValueError, TypeError) as exc:
                    logger.debug("Ignored animation operation error: %s", exc, exc_info=True)

            self.tween(200, _up, ease=self.ease_out_cubic, cancel_token=_token)

        def _on_focus_out(_e) -> None:
            nonlocal _token
            _token[0] = True
            _token = [False]
            from_color = _current[0]

            def _down(t: float) -> None:
                c = self.blend_hex(from_color, INPUT_BORDER, t)
                _current[0] = c
                try:
                    widget.configure(border_color=c)
                except (tk.TclError, AttributeError, ValueError, TypeError) as exc:
                    logger.debug("Ignored animation operation error: %s", exc, exc_info=True)

            self.tween(350, _down, ease=self.ease_out_cubic, cancel_token=_token)

        try:
            widget.bind("<FocusIn>", _on_focus_in, add="+")
            widget.bind("<FocusOut>", _on_focus_out, add="+")
        except (tk.TclError, AttributeError, ValueError, TypeError) as exc:
            logger.debug("Ignored animation operation error: %s", exc, exc_info=True)


