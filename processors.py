import pandas as pd
from pathlib import Path
from utils import safe_float

def calculate_zona(df, config):
    cols = config["file1"]["colZone"]["cols"]
    total = config["file1"]["totalZone"]
    suma = sum(df[c["col"]].apply(safe_float) for c in cols)
    return ((suma * 100 / total * 0.60) * 75 / 60).round(2)

def calculate_final(df, config):
    col_exam = config["file1"]["colExam"]["colA"]
    total = config["file1"]["totalFinal"]
    return ((df[col_exam].apply(safe_float) * 100 / total * 0.40) * 25 / 40).round(2)

def add_column(file_csv: Path, col_name: str, config: dict):
    df = pd.read_csv(file_csv)
    if col_name == "Zona":
        df[col_name] = calculate_zona(df, config)
    elif col_name == "Final":
        df[col_name] = calculate_final(df, config)
    else:
        df[col_name] = 100
    df.to_csv(file_csv, index=False)
    print(f"Added column {col_name} to {file_csv.name}")

def merge_columns(source_csv: Path, target_csv: Path, excel_out: Path,
                  source_col: str, target_col: str, source_id: str, target_id: str):
    df_src = pd.read_csv(source_csv)
    df_tgt = pd.read_csv(target_csv)

    # Crear un diccionario ID -> valor de columna fuente
    mapping = df_src.set_index(source_id)[source_col].to_dict()

    # Rellenar la columna destino solo donde estaba vacía o NaN
    df_tgt[target_col] = df_tgt[target_id].map(mapping).fillna(df_tgt[target_col])

    # Guardar cambios
    df_tgt.to_csv(target_csv, index=False)
    df_tgt.to_excel(excel_out, index=False, engine='openpyxl')
    print(f"Updated column {target_col} from {source_col} using IDs")
