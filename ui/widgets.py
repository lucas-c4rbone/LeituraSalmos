from __future__ import annotations

import tkinter as tk

import customtkinter as ctk

from ui.theme import (
    BG_APP,
    BG_CARD,
    BG_INPUT,
    FG_INPUT,
    FG_TITLE,
    FONT_FAMILY,
    INPUT_BORDER,
    INPUT_BORDER_FOC,
    RADIUS_PILL,
)


class FluentWidgets:
    """Factory helpers for consistently styled CustomTkinter widgets."""

    @staticmethod
    def field_label(parent, text: str, row: int, accent_color: str = FG_TITLE) -> ctk.CTkLabel:
        """Create and place a section label in a grid row.

        Args:
            parent: Container widget.
            text: Label text.
            row: Target grid row.
            accent_color: Foreground color for emphasis.

        Returns:
            ctk.CTkLabel: Created label widget.
        """
        lbl = ctk.CTkLabel(
            parent,
            text=text.upper(),
            fg_color=BG_CARD,
            text_color=accent_color,
            font=(FONT_FAMILY, 11, "bold"),
            anchor="w",
        )
        lbl.grid(row=row, column=0, sticky="w")
        return lbl

    @staticmethod
    def entry(parent, var: tk.Variable, inner_ipady: int = 6) -> ctk.CTkEntry:
        """Create a styled entry field with focus border behavior.

        Args:
            parent: Container widget.
            var: Tk variable bound to the entry value.
            inner_ipady: Vertical inner padding used by caller layouts.

        Returns:
            ctk.CTkEntry: Created entry widget.
        """
        entry = ctk.CTkEntry(
            parent,
            textvariable=var,
            fg_color=BG_INPUT,
            text_color=FG_INPUT,
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

        entry.bind("<FocusIn>", _on_focus_in)
        entry.bind("<FocusOut>", _on_focus_out)
        return entry

    @staticmethod
    def primary_button(parent, text: str, command, fg_color: str, hover_color: str) -> ctk.CTkButton:
        """Create a primary action button using project styling.

        Args:
            parent: Container widget.
            text: Button caption.
            command: Callback invoked on click.
            fg_color: Default background color.
            hover_color: Hover background color.

        Returns:
            ctk.CTkButton: Created button widget.
        """
        return ctk.CTkButton(
            parent,
            text=text,
            command=command,
            fg_color=fg_color,
            text_color=BG_APP,
            hover_color=hover_color,
            font=(FONT_FAMILY, 14, "bold"),
            cursor="hand2",
            corner_radius=RADIUS_PILL,
            height=50,
        )
