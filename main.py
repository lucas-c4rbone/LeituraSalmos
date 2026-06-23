from __future__ import annotations

from ui.layout import Aplicacao


def main() -> int:
    """Start the desktop application.

    Returns:
        int: Process exit status code.
    """
    app = Aplicacao()
    app.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())