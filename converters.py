from pathlib import Path
import pandas as pd
from odf.opendocument import load
from odf.table import Table, TableRow, TableCell
from odf.text import P

def convert_to_csv(file_path: Path) -> Path:
    """Convierte XLSX u ODS a CSV."""
    csv_path = file_path.with_suffix(".csv")
    if file_path.suffix == ".xlsx":
        pd.read_excel(file_path, engine='openpyxl').to_csv(csv_path, index=False)
    elif file_path.suffix == ".ods":
        doc = load(str(file_path))
        sheet = doc.getElementsByType(Table)[0]
        with csv_path.open('w', encoding='utf-8') as f:
            for row in sheet.getElementsByType(TableRow):
                row_data = [
                    ''.join(p.firstChild.data for p in cell.getElementsByType(P) if p.firstChild)
                    for cell in row.getElementsByType(TableCell)
                ]
                f.write(','.join(map(str.strip, row_data)) + '\n')
    else:
        raise ValueError(f"Formato no soportado: {file_path.suffix}")
    return csv_path
