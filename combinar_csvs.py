"""
Combina varios archivos CSV (formato Timestamp / Valor / Tag) en una sola tabla,
alineados por Timestamp, rellenando los huecos con forward-fill y luego backward-fill.

PIEXTRACT - by Alejandro Burelo Sanchez

Requisitos:
    pip install pandas openpyxl

Dos modos de uso:

  1) MODO INTERACTIVO (doble clic o "python combinar_csvs.py" sin argumentos):
     se abren diálogos para elegir la carpeta y dónde guardar el resultado.

  2) MODO SILENCIOSO / DESATENDIDO (para llamarlo desde VBA, un .exe, etc.):
         python combinar_csvs.py "C:\\ruta\\a\\la\\carpeta\\con\\csv"
     - No pregunta carpeta ni nombre de salida.
     - Genera "COMBINADO.xlsx" dentro de esa misma carpeta (sobrescribiendo
       si ya existe).
     - Al terminar abre el archivo automáticamente (os.startfile, Windows).
     - Si algo falla, el error queda registrado en "COMBINADO_error.log"
       dentro de la misma carpeta.

En ambos modos se muestra una ventana con barra de progreso (verde) para
que se vea que el proceso está trabajando, ya que el .exe se compila con
--noconsole y de otra forma corre totalmente invisible.

Soporta dos formatos de CSV:
  1) Un solo tag por archivo: columnas Timestamp, Valor, Tag (en ese orden).
  2) Varios tags "en bloque" dentro del mismo archivo (como en el Excel
     original): grupos de 3 columnas (Timestamp, Valor, Tag) uno al lado
     del otro, separados por una columna vacía.
"""

import glob
import os
import sys
import time
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd
from openpyxl import Workbook

AUTOR = "PIEXTRACT BY: ALEJANDRO BURELO SANCHEZ"

# Excel (.xlsx) soporta como maximo 1,048,576 filas por hoja (incluyendo el
# encabezado). Dejamos un margen para el encabezado de cada hoja.
FILAS_MAX_POR_HOJA = 1_048_575

# Cada cuanto tiempo (segundos) se refresca la ventana mientras se escriben
# filas. No es un porcentaje fijo: el avance se ve libre/continuo, y esto
# solo evita redibujar la ventana mas seguido de lo que el ojo puede ver
# (lo cual ademas frenaria muchisimo la escritura con datasets grandes).
INTERVALO_REFRESCO_SEGUNDOS = 0.05


def guardar_excel_multi_hoja(df: pd.DataFrame, salida: str, progreso=None, paso_base: float = 0) -> int:
    """
    Guarda el DataFrame en un .xlsx escribiendo fila por fila con openpyxl
    (modo write_only, rapido incluso con millones de filas). Si supera el
    limite de una sola hoja de Excel, lo reparte automaticamente en
    "Datos1", "Datos2", etc. El progreso de cada hoja se reporta en la
    barra celeste de forma libre (no por saltos de 10%), refrescando la
    ventana varias veces por segundo mientras escribe. Devuelve la
    cantidad de hojas escritas.
    """
    total_filas_todas = len(df)

    if total_filas_todas <= FILAS_MAX_POR_HOJA:
        chunks = [df]
    else:
        n_hojas_calc = -(-total_filas_todas // FILAS_MAX_POR_HOJA)  # ceil
        chunks = [
            df.iloc[i * FILAS_MAX_POR_HOJA:(i + 1) * FILAS_MAX_POR_HOJA]
            for i in range(n_hojas_calc)
        ]

    n_hojas = len(chunks)
    wb = Workbook(write_only=True)
    filas_escritas_total = 0

    if progreso:
        progreso.mostrar_barra_hoja()

    for idx, chunk in enumerate(chunks, start=1):
        nombre_hoja = "Datos" if n_hojas == 1 else f"Datos{idx}"
        ws = wb.create_sheet(title=nombre_hoja)
        ws.append(list(chunk.columns))

        if progreso:
            valor_barra = paso_base + (filas_escritas_total / total_filas_todas)
            progreso.actualizar(valor_barra, f"Guardando hoja {idx}/{n_hojas}...")

        # NaN/NaT no son validos para openpyxl -> los pasamos a None (celda vacia)
        chunk_seguro = chunk.astype(object).where(pd.notnull(chunk), None)
        total_filas_hoja = len(chunk_seguro)
        ultimo_refresco = 0.0

        for i, fila in enumerate(chunk_seguro.itertuples(index=False, name=None), start=1):
            ws.append(fila)
            filas_escritas_total += 1

            if progreso:
                ahora = time.perf_counter()
                if ahora - ultimo_refresco >= INTERVALO_REFRESCO_SEGUNDOS or i == total_filas_hoja:
                    ultimo_refresco = ahora
                    valor_barra = paso_base + (filas_escritas_total / total_filas_todas)
                    progreso.actualizar(valor_barra)
                    progreso.actualizar_hoja(i, total_filas_hoja)

    if progreso:
        progreso.ocultar_barra_hoja()

    wb.save(salida)
    return n_hojas


class VentanaProgreso:
    """Ventana simple con barra de progreso verde y el credito del autor
    arriba. Se actualiza llamando a .actualizar(valor, mensaje) y se cierra
    con .cerrar(). No usa hilos: hay que llamar a .actualizar(...) seguido
    para que la ventana se refresque (Tkinter sin mainloop propio)."""

    def __init__(self, total: int):
        self.total = max(total, 1)
        self.root = tk.Tk()
        self.root.title("PIEXTRACT")
        self.root.geometry("460x220")
        self.root.resizable(False, False)
        self.root.attributes("-topmost", True)

        tk.Label(
            self.root,
            text=AUTOR,
            font=("Segoe UI", 9, "bold"),
        ).pack(pady=(12, 6))

        self.label_estado = tk.Label(
            self.root, text="Iniciando...", font=("Segoe UI", 9)
        )
        self.label_estado.pack(pady=(0, 4))

        estilo = ttk.Style()
        try:
            estilo.theme_use("default")
        except Exception:
            pass
        estilo.configure(
            "Verde.Horizontal.TProgressbar",
            troughcolor="#e0e0e0",
            background="#2ecc71",
            bordercolor="#2ecc71",
            lightcolor="#2ecc71",
            darkcolor="#2ecc71",
            thickness=18,
        )
        estilo.configure(
            "Celeste.Horizontal.TProgressbar",
            troughcolor="#e0e0e0",
            background="#5dade2",
            bordercolor="#5dade2",
            lightcolor="#5dade2",
            darkcolor="#5dade2",
            thickness=14,
        )

        self.barra = ttk.Progressbar(
            self.root,
            style="Verde.Horizontal.TProgressbar",
            orient="horizontal",
            length=400,
            mode="determinate",
            maximum=self.total,
        )
        self.barra.pack(pady=(2, 2))

        self.label_pct = tk.Label(self.root, text="0%", font=("Segoe UI", 8))
        self.label_pct.pack(pady=(0, 8))

        # Barra secundaria (celeste), para el llenado fila a fila de cada
        # hoja del Excel. Se muestra/oculta con mostrar_barra_hoja() /
        # ocultar_barra_hoja() -- solo esta visible mientras se estan
        # escribiendo filas al archivo.
        self.frame_hoja = tk.Frame(self.root)
        self.label_hoja = tk.Label(self.frame_hoja, text="", font=("Segoe UI", 8))
        self.label_hoja.pack()
        self.barra_hoja = ttk.Progressbar(
            self.frame_hoja,
            style="Celeste.Horizontal.TProgressbar",
            orient="horizontal",
            length=400,
            mode="determinate",
            maximum=100,
        )
        self.barra_hoja.pack(pady=(4, 4))
        # frame_hoja no se empaqueta todavia -> arranca oculta

        self._refrescar()

    def actualizar(self, valor: float, mensaje: str = ""):
        valor = min(valor, self.total)
        self.barra["value"] = valor
        pct = int(round((valor / self.total) * 100))
        self.label_pct.config(text=f"{pct}%")
        if mensaje:
            self.label_estado.config(text=mensaje)
        self._refrescar()

    def mostrar_barra_hoja(self):
        self.frame_hoja.pack(pady=(0, 4))
        self._refrescar()

    def ocultar_barra_hoja(self):
        self.frame_hoja.pack_forget()
        self._refrescar()

    def actualizar_hoja(self, fila_actual: int, total_filas: int):
        pct = int(round((fila_actual / total_filas) * 100)) if total_filas else 0
        self.barra_hoja["value"] = pct
        self.label_hoja.config(
            text=f"Guardando Filas {fila_actual:,}/{total_filas:,} ({pct}%)"
        )
        self._refrescar()

    def _refrescar(self):
        try:
            self.root.update_idletasks()
            self.root.update()
        except Exception:
            pass

    def cerrar(self):
        try:
            self.root.destroy()
        except Exception:
            pass


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
        initialfile="COMBINADO.xlsx",
        defaultextension=".xlsx",
        filetypes=[("Excel", "*.xlsx"), ("CSV", "*.csv")],
    )
    root.destroy()
    return salida


def mostrar_error(mensaje: str, carpeta: str | None = None, modo_silencioso: bool = False):
    """Muestra el error en un messagebox (modo interactivo) o lo escribe
    en un log junto a la carpeta procesada (modo silencioso / exe)."""
    if modo_silencioso:
        try:
            log_path = os.path.join(carpeta or os.getcwd(), "COMBINADO_error.log")
            with open(log_path, "w", encoding="utf-8") as f:
                f.write(mensaje)
        except Exception:
            pass
    else:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Error", mensaje)
        root.destroy()


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
    # --- Determinar modo: silencioso (con carpeta como argumento) o interactivo ---
    modo_silencioso = len(sys.argv) > 1
    carpeta = sys.argv[1] if modo_silencioso else elegir_carpeta()

    if not carpeta or not os.path.isdir(carpeta):
        msg = f"La carpeta no existe o no fue especificada: {carpeta!r}"
        print(msg)
        mostrar_error(msg, carpeta, modo_silencioso)
        os._exit(1)

    # Excluye el propio archivo de salida por si ya existe de una corrida anterior
    archivos_csv = sorted(
        f for f in glob.glob(os.path.join(carpeta, "*.csv"))
        if not os.path.basename(f).upper().startswith("COMBINADO")
    )
    if not archivos_csv:
        msg = "No se encontraron archivos .csv en esa carpeta."
        print(msg)
        mostrar_error(msg, carpeta, modo_silencioso)
        os._exit(1)

    print(f"Encontrados {len(archivos_csv)} archivos CSV:")
    for f in archivos_csv:
        print(f"  - {os.path.basename(f)}")

    # Total de pasos para la barra: 1 por archivo leido + 1 para combinar + 1 para guardar
    total_pasos = len(archivos_csv) + 2
    progreso = VentanaProgreso(total_pasos)

    todos_los_dfs = []
    for idx, f in enumerate(archivos_csv, start=1):
        nombre = os.path.basename(f)
        progreso.actualizar(idx - 1, f"Leyendo {nombre} ({idx}/{len(archivos_csv)})...")
        try:
            dfs = leer_csv_como_bloques(f)
            todos_los_dfs.extend(dfs)
            print(f"  {nombre}: {len(dfs)} tag(s) encontrados")
        except Exception as e:
            print(f"  Error leyendo {nombre}: {e}")
        progreso.actualizar(idx, f"Leido {nombre}")

    if not todos_los_dfs:
        progreso.cerrar()
        msg = "No se encontraron datos válidos (Timestamp/Valor/Tag) en los CSV."
        print(msg)
        mostrar_error(msg, carpeta, modo_silencioso)
        os._exit(1)

    progreso.actualizar(len(archivos_csv) + 1, "Combinando datos (merge por Timestamp)...")
    resultado = todos_los_dfs[0]
    for d in todos_los_dfs[1:]:
        resultado = resultado.merge(d, on="Timestamp", how="outer")

    resultado = resultado.sort_values("Timestamp").reset_index(drop=True)
    resultado = resultado.ffill().bfill()

    if modo_silencioso:
        salida = os.path.join(carpeta, "COMBINADO.xlsx")
    else:
        progreso.cerrar()
        salida = elegir_archivo_salida(carpeta)
        if not salida:
            print("No se guardó ningún archivo.")
            os._exit(0)
        progreso = VentanaProgreso(total_pasos)

    progreso.actualizar(total_pasos - 1, "Guardando archivo combinado...")

    if salida.lower().endswith(".xlsx"):
        n_hojas = guardar_excel_multi_hoja(
            resultado, salida, progreso=progreso, paso_base=total_pasos - 1
        )
        if n_hojas > 1:
            print(
                f"Datos repartidos en {n_hojas} hojas "
                f"(limite de Excel: {FILAS_MAX_POR_HOJA} filas por hoja)."
            )
    else:
        resultado.to_csv(salida, index=False)

    progreso.actualizar(total_pasos, "Listo.")
    print(f"Archivo combinado guardado en: {salida}")

    progreso.cerrar()

    # Abrir el archivo final automáticamente (Windows)
    try:
        os.startfile(salida)  # type: ignore[attr-defined]
    except Exception:
        pass  # si no es Windows o falla, simplemente no lo abre

    if not modo_silencioso:
        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo("Listo", f"Archivo combinado guardado en:\n{salida}")
        root.destroy()

    # Cierre forzado del proceso.
    # Motivo: pandas/numpy (via las librerias de calculo que usan internamente,
    # p.ej. OpenBLAS) a veces dejan hilos secundarios vivos que no terminan
    # solos dentro de un .exe empaquetado con PyInstaller. Eso hace que el
    # proceso quede "colgado" en el Administrador de Tareas aunque ya termino
    # su trabajo (el archivo ya se genero y se abrio). os._exit() termina el
    # proceso de inmediato sin esperar a esos hilos ni correr el cleanup
    # normal del interprete -- es seguro aca porque ya no queda nada
    # pendiente por guardar.
    sys.stdout.flush() if sys.stdout else None
    os._exit(0)


if __name__ == "__main__":
    try:
        main()
    except Exception:
        error_completo = traceback.format_exc()
        print(error_completo)
        carpeta_fallback = sys.argv[1] if len(sys.argv) > 1 else os.getcwd()
        mostrar_error(error_completo, carpeta_fallback, modo_silencioso=len(sys.argv) > 1)
        os._exit(1)
