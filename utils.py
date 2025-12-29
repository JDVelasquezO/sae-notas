from pathlib import Path

def safe_float(val):
    try:
        return float(val)
    except (ValueError, TypeError):
        return 0.0

def delete_file(path: Path):
    if path.exists():
        path.unlink()
