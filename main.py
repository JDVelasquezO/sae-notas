import json
from pathlib import Path
from converters import convert_to_csv
from processors import add_column, merge_columns
from utils import delete_file

def operate(file1: Path, file2: Path, config: dict):
    file1_csv = convert_to_csv(file1)
    file2_csv = convert_to_csv(file2)

    add_column(file1_csv, "Zona", config)
    add_column(file1_csv, "Final", config)
    add_column(file2_csv, "Laboratorio 100%", config)
    add_column(file2_csv, "Asistencia 100%", config)

    try:
        merge_columns(file1_csv, file2_csv, file2, "Zona", "Zona", "Número de ID", "CUI/Pasaporte")
        merge_columns(file1_csv, file2_csv, file2, "Final", "Final", "Número de ID", "CUI/Pasaporte")
    except KeyError as e:
        print(f"Error: La columna {e} no existe en el archivo. Revisar encabezados")

    delete_file(file1_csv)
    delete_file(file2_csv)

def main():
    folder = Path('./test')
    file1 = next(folder.glob('*.ods'))
    file2 = next(folder.glob('*.xlsx'))
    config = json.loads(Path('./input.json').read_text())
    operate(file1, file2, config)

if __name__ == "__main__":
    main()
