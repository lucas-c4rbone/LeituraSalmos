from pathlib import Path


def test_refactored_modules_exist() -> None:
    """Verify all expected refactored module files are present."""
    root = Path(__file__).resolve().parents[1]
    expected: list[Path] = [
        root / 'app.py',
        root / 'core' / '__init__.py',
        root / 'core' / 'generator.py',
        root / 'core' / 'scripture.py',
        root / 'core' / 'config.py',
        root / 'core' / 'logging.py',
        root / 'core' / 'utils.py',
        root / 'ui' / '__init__.py',
        root / 'ui' / 'layout.py',
        root / 'ui' / 'animations.py',
        root / 'ui' / 'widgets.py',
        root / 'ui' / 'theme.py',
    ]
    for path in expected:
        assert path.exists(), f'Missing expected module: {path}'
