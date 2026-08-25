"""
Combina varios archivos CSV (formato Timestamp / Valor / Tag) en una sola tabla,
alineados por Timestamp, rellenando los huecos con forward-fill y luego backward-fill.

Requisitos:
    pip install pandas openpyxl

Uso:
    python combinar_csvs.py

Al ejecutarlo se abre el explorador de Windows/Mac para elegir la carpeta
con los CSV, y al final otro diálogo para elegir dónde guardar el resultado
(.xlsx o .csv).

Soporta dos formatos de CSV:
  1) Un solo tag por archivo: columnas Timestamp, Valor, Tag (en ese orden).
  2) Varios tags "en bloque" dentro del mismo archivo (como en el Excel
     original): grupos de 3 columnas (Timestamp, Valor, Tag) uno al lado
     del otro, separados por una columna vacía.
"""

import glob
import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox


def elegir_carpeta() -> str:
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    carpeta = filedialog.askdirectory(title="Selecciona la carpeta con los CSV")
    root.destroy()
    return carpeta


def elegir_archivo_salida(carpeta_inicial: str) -> str:
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    salida = filedialog.asksaveasfilename(
        title="Guardar archivo combinado como...",
        initialdir=carpeta_inicial,
        initialfile="combinado.xlsx",
        defaultextension=".xlsx",
        filetypes=[("Excel", "*.xlsx"), ("CSV", "*.csv")],
    )
    root.destroy()
    return salida


def leer_csv_como_bloques(path: str) -> list[pd.DataFrame]:
    """
    Lee un CSV y devuelve una lista de DataFrames (Timestamp, <NombreTag>),
    uno por cada tag encontrado, ya sea que el archivo tenga un solo tag
    o varios en bloques de 3 columnas.
    """
    raw = pd.read_csv(path, header=None, low_memory=False)

    dfs = []
    col = 0
    while col + 2 < raw.shape[1]:
        bloque = raw.iloc[:, col:col + 3].copy()
        bloque.columns = ["Timestamp", "Valor", "Tag"]
        bloque = bloque.iloc[1:].dropna(how="all")  # saltar encabezado del bloque

        if len(bloque) > 0 and bloque["Tag"].notna().any():
            tag_name = bloque["Tag"].dropna().iloc[0]
            bloque = bloque[["Timestamp", "Valor"]].dropna()
            bloque.columns = ["Timestamp", tag_name]
            bloque["Timestamp"] = pd.to_datetime(bloque["Timestamp"], errors="coerce")
            bloque = bloque.dropna(subset=["Timestamp"]).drop_duplicates(subset=["Timestamp"])
            if len(bloque) > 0:
                dfs.append(bloque)

        col += 4  # 3 columnas de datos + 1 de separación

    return dfs


def main():
    carpeta = elegir_carpeta()
    if not carpeta:
        print("No se seleccionó ninguna carpeta. Saliendo.")
        return

    archivos_csv = sorted(glob.glob(os.path.join(carpeta, "*.csv")))
    if not archivos_csv:
        messagebox.showwarning("Sin archivos", "No se encontraron archivos .csv en esa carpeta.")
        return

    print(f"Encontrados {len(archivos_csv)} archivos CSV:")
    for f in archivos_csv:
        print(f"  - {os.path.basename(f)}")

    todos_los_dfs = []
    for f in archivos_csv:
        try:
            dfs = leer_csv_como_bloques(f)
            todos_los_dfs.extend(dfs)
            print(f"  {os.path.basename(f)}: {len(dfs)} tag(s) encontrados")
        except Exception as e:
            print(f"  Error leyendo {os.path.basename(f)}: {e}")

    if not todos_los_dfs:
        messagebox.showwarning(
            "Sin datos", "No se encontraron datos válidos (Timestamp/Valor/Tag) en los CSV."
        )
        return

    resultado = todos_los_dfs[0]
    for d in todos_los_dfs[1:]:
        resultado = resultado.merge(d, on="Timestamp", how="outer")

    resultado = resultado.sort_values("Timestamp").reset_index(drop=True)
    resultado = resultado.ffill().bfill()

    salida = elegir_archivo_salida(carpeta)
    if not salida:
        print("No se guardó ningún archivo.")
        return

    if salida.lower().endswith(".xlsx"):
        resultado.to_excel(salida, index=False)
    else:
        resultado.to_csv(salida, index=False)

    print(f"Archivo combinado guardado en: {salida}")
    messagebox.showinfo("Listo", f"Archivo combinado guardado en:\n{salida}")


if __name__ == "__main__":
    main()
