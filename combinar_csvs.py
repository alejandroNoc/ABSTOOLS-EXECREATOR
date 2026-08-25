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


class VentanaProgreso:
    """Ventanita pequeña con barra verde de progreso y porcentaje en medio."""

    def __init__(self, titulo="Procesando..."):
        self.root = tk.Tk()
        self.root.title(titulo)
        self.root.geometry("300x80")
        self.root.resizable(False, False)
        self.root.attributes("-topmost", True)

        # Centrar en pantalla
        self.root.update_idletasks()
        ancho, alto = 300, 80
        x = (self.root.winfo_screenwidth() // 2) - (ancho // 2)
        y = (self.root.winfo_screenheight() // 2) - (alto // 2)
        self.root.geometry(f"{ancho}x{alto}+{x}+{y}")

        self.canvas = tk.Canvas(self.root, width=260, height=30, bg="#e0e0e0",
                                 highlightthickness=1, highlightbackground="#aaaaaa")
        self.canvas.pack(pady=15)

        self.barra = self.canvas.create_rectangle(0, 0, 0, 30, fill="#2ecc71", width=0)
        self.texto = self.canvas.create_text(130, 15, text="0%",
                                              font=("Segoe UI", 10, "bold"), fill="#222222")

        self.root.update()

    def actualizar(self, porcentaje: float, mensaje: str = None):
        porcentaje = max(0, min(100, porcentaje))
        ancho_lleno = int(260 * porcentaje / 100)
        self.canvas.coords(self.barra, 0, 0, ancho_lleno, 30)
        texto = f"{int(porcentaje)}%"
        self.canvas.itemconfig(self.texto, text=texto)
        if mensaje:
            self.root.title(mensaje)
        self.root.update()

    def cerrar(self):
        self.root.destroy()


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

    progreso = VentanaProgreso("Leyendo archivos...")

    # Leer archivos: reservamos el 70% de la barra para esta fase
    todos_los_dfs = []
    n = len(archivos_csv)
    for idx, f in enumerate(archivos_csv, start=1):
        try:
            dfs = leer_csv_como_bloques(f)
            todos_los_dfs.extend(dfs)
            print(f"  {os.path.basename(f)}: {len(dfs)} tag(s) encontrados")
        except Exception as e:
            print(f"  Error leyendo {os.path.basename(f)}: {e}")

        porcentaje = (idx / n) * 70
        progreso.actualizar(porcentaje, f"Leyendo {os.path.basename(f)}...")

    if not todos_los_dfs:
        progreso.cerrar()
        messagebox.showwarning(
            "Sin datos", "No se encontraron datos válidos (Timestamp/Valor/Tag) en los CSV."
        )
        return

    # Combinar: el 30% restante de la barra
    progreso.actualizar(75, "Combinando datos...")
    resultado = todos_los_dfs[0]
    total_dfs = len(todos_los_dfs)
    for i, d in enumerate(todos_los_dfs[1:], start=2):
        resultado = resultado.merge(d, on="Timestamp", how="outer")
        porcentaje = 70 + (i / total_dfs) * 25
        progreso.actualizar(porcentaje, "Combinando datos...")

    progreso.actualizar(97, "Ordenando y rellenando huecos...")
    resultado = resultado.sort_values("Timestamp").reset_index(drop=True)
    resultado = resultado.ffill().bfill()

    progreso.actualizar(100, "Completado")
    progreso.cerrar()

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
