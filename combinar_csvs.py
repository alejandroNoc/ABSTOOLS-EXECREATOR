"""
Combina varios archivos CSV (formato Timestamp / Valor / Tag) en una sola tabla,
alineados por Timestamp, rellenando los huecos con forward-fill y luego backward-fill.

Requisitos:
    pip install pandas openpyxl

Dos modos de uso:

  1) MODO INTERACTIVO (doble clic o "python combinar_csvs.py" sin argumentos):
     se abren diálogos para elegir la carpeta y dónde guardar el resultado.

  2) MODO SILENCIOSO / DESATENDIDO (para llamarlo desde VBA, un .exe, etc.):
         python combinar_csvs.py "C:\\ruta\\a\\la\\carpeta\\con\\csv"
     - No muestra ningún diálogo.
     - Genera "COMBINADO.xlsx" dentro de esa misma carpeta (sobrescribiendo
       si ya existe).
     - Al terminar abre el archivo automáticamente (os.startfile, Windows).
     - Si algo falla, el error queda registrado en "COMBINADO_error.log"
       dentro de la misma carpeta (útil porque un .exe compilado con
       --noconsole no muestra ninguna ventana de consola).

Soporta dos formatos de CSV:
  1) Un solo tag por archivo: columnas Timestamp, Valor, Tag (en ese orden).
  2) Varios tags "en bloque" dentro del mismo archivo (como en el Excel
     original): grupos de 3 columnas (Timestamp, Valor, Tag) uno al lado
     del otro, separados por una columna vacía.
"""

import glob
import os
import sys
import traceback
import pandas as pd


def elegir_carpeta() -> str:
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    carpeta = filedialog.askdirectory(title="Selecciona la carpeta con los CSV")
    root.destroy()
    return carpeta


def elegir_archivo_salida(carpeta_inicial: str) -> str:
    import tkinter as tk
    from tkinter import filedialog

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
        import tkinter as tk
        from tkinter import messagebox

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

    todos_los_dfs = []
    for f in archivos_csv:
        try:
            dfs = leer_csv_como_bloques(f)
            todos_los_dfs.extend(dfs)
            print(f"  {os.path.basename(f)}: {len(dfs)} tag(s) encontrados")
        except Exception as e:
            print(f"  Error leyendo {os.path.basename(f)}: {e}")

    if not todos_los_dfs:
        msg = "No se encontraron datos válidos (Timestamp/Valor/Tag) en los CSV."
        print(msg)
        mostrar_error(msg, carpeta, modo_silencioso)
        os._exit(1)

    resultado = todos_los_dfs[0]
    for d in todos_los_dfs[1:]:
        resultado = resultado.merge(d, on="Timestamp", how="outer")

    resultado = resultado.sort_values("Timestamp").reset_index(drop=True)
    resultado = resultado.ffill().bfill()

    if modo_silencioso:
        salida = os.path.join(carpeta, "COMBINADO.xlsx")
    else:
        salida = elegir_archivo_salida(carpeta)
        if not salida:
            print("No se guardó ningún archivo.")
            return

    if salida.lower().endswith(".xlsx"):
        resultado.to_excel(salida, index=False)
    else:
        resultado.to_csv(salida, index=False)

    print(f"Archivo combinado guardado en: {salida}")

    # Abrir el archivo final automáticamente (Windows)
    try:
        os.startfile(salida)  # type: ignore[attr-defined]
    except Exception:
        pass  # si no es Windows o falla, simplemente no lo abre

    if not modo_silencioso:
        import tkinter as tk
        from tkinter import messagebox

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
