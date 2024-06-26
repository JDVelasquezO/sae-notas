import json
import os.path

import pandas as pd
from odf.opendocument import load
from odf.table import Table, TableRow, TableCell
from odf.text import P
from functools import reduce


# Función para convertir archivo xlsx a csv
def convert_xlsx_to_csv(xlsx_file):
    csv_file = xlsx_file[:-5] + ".csv"
    df = pd.read_excel(xlsx_file, engine='openpyxl')
    df.to_csv(csv_file, index=False)
    return csv_file


# Función para convertir archivo ods a csv
def convert_ods_to_csv(ods_file):
    csv_file = ods_file[:-4] + ".csv"
    doc = load(ods_file)
    sheet = doc.getElementsByType(Table)[0]

    with open(csv_file, 'w', encoding='utf-8') as f:
        for row in sheet.getElementsByType(TableRow):
            row_data = []
            for cell in row.getElementsByType(TableCell):
                cell_text = ''
                for p in cell.getElementsByType(P):
                    for text_node in p.childNodes:
                        if hasattr(text_node, 'data'):  # Verificar si es un nodo de texto
                            cell_text += text_node.data
                row_data.append(cell_text.strip())
            f.write(','.join(row_data) + '\n')

    return csv_file


def addSomeCol(fileCSV, nameCol, objectCols):
    try:
        df = pd.read_csv(fileCSV)

        # Función para convertir a float y manejar valores no convertibles
        def convertFloat(val):
            try:
                return float(val)
            except ValueError:
                return 0.0

        file1 = objectCols.get("file1")

        if nameCol == "Zona":
            arrayColsZone = file1.get("colZone").get("cols")
            sumCols = reduce(lambda acc, item: acc + df[item.get('col', 0)].apply(convertFloat),
                             arrayColsZone, 0)
            totalRes = sumCols * 100 / file1.get("totalZone") * 75 / 100
            df[nameCol] = totalRes.round(2)

        elif nameCol == "Final":
            colExam = file1.get("colExam").get("colA")
            totalRes = df[colExam].apply(convertFloat) * 100 / file1.get("totalFinal") * 25 / 100
            df[nameCol] = totalRes.round(2)

        else:
            df[nameCol] = 100

        df.to_csv(fileCSV, index=False)
        print(f"added col {nameCol} to {fileCSV}")

    except Exception as e:
        print("Error: ", str(e))


def sendCols(sourceCSV, sourceCol, destinyCSV, destinyCol, colSourceID, colDestinationID, excelFile):
    try:
        dfSource = pd.read_csv(sourceCSV)
        dfDest = pd.read_csv(destinyCSV)

        for index, row in dfSource.iterrows():
            valSourceID = row[colSourceID]
            mask = dfDest[colDestinationID] == valSourceID
            dfDest.loc[mask, destinyCol] = row[sourceCol]

        dfDest.to_csv(destinyCSV, index=False)
        dfDest.to_excel(excelFile, index=False, engine='openpyxl')

        print(f"Added cols from source to destination from {colSourceID} to {colDestinationID}")
        print("Todo bien")

    except Exception as e:
        print("Error: ", str(e))


def operate(file1, file2, fileCols):
    oldFile2 = file2
    if file1.endswith('.xlsx'):
        file1 = convert_xlsx_to_csv(file1)
    elif file1.endswith('.ods'):
        file1 = convert_ods_to_csv(file1)

    if file2.endswith('.xlsx'):
        file2 = convert_xlsx_to_csv(file2)
    elif file2.endswith('.ods'):
        file2 = convert_ods_to_csv(file2)

    addSomeCol(file1, 'Zona', fileCols)
    addSomeCol(file1, 'Final', fileCols)

    addSomeCol(file2, 'Laboratorio 100%', fileCols)
    addSomeCol(file2, 'Asistencia 100%', fileCols)

    sendCols(file1, 'Zona', file2, 'Zona', 'Número de ID',
             'CUI/Pasaporte', oldFile2)
    sendCols(file1, 'Final', file2, 'Final', 'Número de ID',
             'CUI/Pasaporte', oldFile2)

    deleteCSV(file1)
    deleteCSV(file2)


def deleteCSV(file):
    try:
        if os.path.exists(file):
            os.remove(file)
        else:
            print("Doesn't exists")
    except Exception as e:
        print("Error: ", str(e))


def main():
    print("Escribe la ruta de tu forlder")
    folder = input()
    # folder = './test'
    files = os.listdir(folder)
    file1 = [file for file in files if file.endswith('.ods')]
    file2 = [file for file in files if file.endswith('.xlsx')]

    print("Escribe la ruta de tu json:")
    fileJson = input()

    with open(fileJson, 'r') as file:
        fileCols = json.load(file)

    operate(os.path.join(folder, file1[0]), os.path.join(folder, file2[0]), fileCols)


main()
