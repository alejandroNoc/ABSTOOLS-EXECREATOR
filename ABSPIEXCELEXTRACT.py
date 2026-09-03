"""
ABSPIEXCELEXTRACT - Combinador de CSVs de PI

Este programa NO extrae datos de PI ni toca COM para nada -- esa
parte la hace la macro VBA "GenerarParametrosExtraccion" de forma
nativa (New/CreateObject de PISDK.PISDK y PITimeServer.PITime desde
VBA, con referencias tempranas), escribiendo un CSV individual por
tag (formato TAGn_nombre_conteo.csv, columnas Timestamp,Valor,Tag)
en una carpeta.

Historial: se probaron dos rutas para hacer la extraccion desde este
.exe (Python con pywin32, y PowerShell embebido) y ambas fallaron en
distintas maquinas por razones fuera de nuestro control:
  - pywin32 (gencache y Dispatch normal) no logra pasar objetos
    PITime como argumento de RecordedValues() de forma confiable
    dentro de un .exe compilado con PyInstaller.
  - PowerShell esta bloqueado por politica de grupo en algunas
    maquinas ("This program is blocked by group policy").
VBA, en cambio, usa COM nativo sin ninguna de esas capas y siempre
funciono bien -- por eso la extraccion se dejo ahi, y este .exe se
redujo a SOLO combinar (leer CSVs + pandas), que nunca tuvo problemas.

Este programa:
  1. Lee todos los CSV (formato Timestamp,Valor,Tag) de una carpeta.
  2. Combina todas las series en una sola tabla, con el enfoque
     "forward-fill + backfill" (pivot por Tag, sin inventar fechas,
     manteniendo el ultimo valor real conocido como hace PI Trend).
  3. Escribe un unico CSV combinado, listo para usar, y lo abre solo.

Si se abre con DOBLE CLIC (sin argumentos), pregunta la carpeta con
los CSV a combinar.

Muestra una ventanita de progreso (Tkinter) con el avance en vivo,
para poder compilarse con --noconsole y aun asi ver que esta pasando.

Requisitos para COMPILAR (no para el usuario final):
    pip install pandas pyinstaller
    pyinstaller --onefile --noconsole --name ABSPIEXCELEXTRACT ABSPIEXCELEXTRACT.py

Este .exe debe quedar instalado en: C:\\ABSTOOLS\\ABSPIEXCELEXTRACT.exe
(esa es la ruta que espera el modulo VBA "GenerarParametrosExtraccion"
si el usuario elige lanzarlo automaticamente al terminar de extraer).

Uso por linea de comandos (lanzado desde VBA, o manual):
    ABSPIEXCELEXTRACT.exe "C:\\ruta\\carpeta_con_csvs" [carpeta_salida]
"""
import sys
import os
import csv
import threading
from datetime import datetime

try:
    import pandas as pd
except ImportError:
    pd = None

import tkinter as tk
from tkinter import messagebox
from tkinter import ttk


class VentanaProgreso:
    """Ventanita de progreso: titulo, barra de estado (texto), barra
    de progreso visual (determinate mientras se conoce el total --
    archivo por archivo -- e indeterminate/animada durante el pivot,
    que es una operacion vectorizada sin pasos individuales), y una
    caja de texto tipo log. No usa mainloop() de forma bloqueante --
    se actualiza manualmente con update() desde el hilo de trabajo."""

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("ABSPIEXCELEXTRACT")
        self.root.geometry("640x460")
        self.root.resizable(True, True)

        titulo = tk.Label(self.root, text="ABSPIEXCELEXTRACT",
                           font=("Segoe UI", 14, "bold"))
        titulo.pack(pady=(12, 0))

        self.estado_var = tk.StringVar(value="Iniciando...")
        estado_lbl = tk.Label(self.root, textvariable=self.estado_var,
                               font=("Segoe UI", 10))
        estado_lbl.pack(pady=(4, 6))

        barra_frame = tk.Frame(self.root)
        barra_frame.pack(fill="x", padx=12, pady=(0, 4))

        self.progress = ttk.Progressbar(barra_frame, orient="horizontal",
                                         mode="determinate", maximum=100)
        self.progress.pack(fill="x", side="left", expand=True)

        self.progress_pct_var = tk.StringVar(value="")
        progress_pct_lbl = tk.Label(barra_frame, textvariable=self.progress_pct_var,
                                     font=("Segoe UI", 9), width=6)
        progress_pct_lbl.pack(side="left", padx=(8, 0))

        frame_log = tk.Frame(self.root)
        frame_log.pack(fill="both", expand=True, padx=12, pady=(8, 12))

        scrollbar = tk.Scrollbar(frame_log)
        scrollbar.pack(side="right", fill="y")

        self.texto = tk.Text(frame_log, wrap="word", font=("Consolas", 9),
                              yscrollcommand=scrollbar.set)
        self.texto.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=self.texto.yview)

        self.boton_cerrar = tk.Button(self.root, text="Cerrar",
                                       command=self.root.destroy, state="disabled")
        self.boton_cerrar.pack(pady=(0, 10))

        self.root.protocol("WM_DELETE_WINDOW", self._on_close_intentado)

    def _on_close_intentado(self):
        if self.boton_cerrar["state"] == "normal":
            self.root.destroy()

    def estado(self, texto):
        self.estado_var.set(texto)
        self.root.update_idletasks()
        self.root.update()

    def log(self, mensaje):
        self.texto.insert("end", str(mensaje) + "\n")
        self.texto.see("end")
        self.root.update_idletasks()
        self.root.update()

    def progreso(self, valor, maximo):
        if maximo <= 0:
            maximo = 1
        if self.progress["mode"] != "determinate":
            self.progress.stop()
            self.progress.config(mode="determinate")
        self.progress.config(maximum=maximo)
        self.progress["value"] = valor
        pct = int(round(100 * valor / maximo))
        self.progress_pct_var.set(f"{pct}%")
        self.root.update_idletasks()
        self.root.update()

    def progreso_indeterminado(self, activar=True):
        if activar:
            self.progress.config(mode="indeterminate")
            self.progress_pct_var.set("...")
            self.progress.start(12)
        else:
            self.progress.stop()
            self.progress.config(mode="determinate")
        self.root.update_idletasks()
        self.root.update()

    def terminar(self, exito=True):
        self.progreso_indeterminado(False)
        self.progress["value"] = self.progress["maximum"] if exito else 0
        self.progress_pct_var.set("100%" if exito else "")
        self.estado("Completado" if exito else "Termino con errores")
        self.boton_cerrar.config(state="normal")

    def iniciar_loop(self):
        self.root.mainloop()


def leer_csv_individual(ruta_csv):
    """Lee un CSV individual (formato Timestamp,Valor,Tag) generado
    por la macro VBA. Devuelve (tag, puntos), con puntos vacio si el
    archivo no tiene datos utilizables (tag no encontrado, error, o
    vacio). Acepta varios formatos de fecha comunes."""
    tag = os.path.splitext(os.path.basename(ruta_csv))[0]
    puntos = []
    tag_real = None

    formatos = ["%Y-%m-%d %H:%M:%S", "%m/%d/%Y %I:%M:%S %p", "%d/%m/%Y %H:%M:%S"]

    with open(ruta_csv, "r", encoding="utf-8-sig", errors="ignore") as f:
        reader = csv.reader(f)
        next(reader, None)  # encabezado
        for row in reader:
            if len(row) < 2:
                continue
            ts_raw, val_raw = row[0].strip(), row[1].strip()
            if not ts_raw or not val_raw:
                continue

            ts_local = None
            for fmt in formatos:
                try:
                    ts_local = datetime.strptime(ts_raw, fmt)
                    break
                except ValueError:
                    continue
            if ts_local is None and pd is not None:
                try:
                    ts_local = pd.to_datetime(ts_raw)
                except Exception:
                    continue
            if ts_local is None:
                continue

            try:
                val = float(val_raw)
            except ValueError:
                continue

            puntos.append((ts_local, val))
            if len(row) >= 3 and row[2].strip():
                tag_real = row[2].strip().strip('"')

    if tag_real:
        tag = tag_real
    return tag, puntos


def combinar_desde_carpeta(carpeta, ui):
    """Lee todos los CSV (formato TAGn_nombre_conteo.csv, o cualquier
    CSV con columnas Timestamp,Valor,Tag) que existan en una carpeta,
    y arma la lista 'series' que espera combinar_series()."""
    archivos = [f for f in os.listdir(carpeta)
                if f.lower().endswith(".csv") and not f.lower().startswith("combinado")]

    if not archivos:
        ui.log("No se encontraron archivos CSV en la carpeta.")
        return []

    ui.log(f"Encontrados {len(archivos)} archivos CSV en la carpeta.\n")

    series = []
    ui.progreso(0, len(archivos))
    for i, nombre in enumerate(sorted(archivos), start=1):
        ruta = os.path.join(carpeta, nombre)
        ui.estado(f"Leyendo {i}/{len(archivos)}: {nombre}")
        tag, puntos = leer_csv_individual(ruta)
        ui.log(f"[{i}/{len(archivos)}] {nombre} -> tag='{tag}', {len(puntos)} puntos")
        if puntos:
            series.append((tag, puntos))
        ui.progreso(i, len(archivos))

    return series


def combinar_series(series, ui):
    """Combina las series en una sola tabla, alineando por 'ultimo
    valor real conocido' (igual que PI Trend), con relleno hacia
    atras (backfill) al inicio de cada serie."""
    ui.estado("Combinando series...")
    ui.log("\n=== Combinando series ===")

    if not series:
        ui.log("No hay series validas para combinar.")
        return None

    total_puntos = sum(len(puntos) for _, puntos in series)
    filas_largo = []
    procesados = 0
    ui.estado("Combinando series: preparando datos...")
    ui.progreso(0, total_puntos)
    for tag, puntos in series:
        for ts, val in puntos:
            filas_largo.append({"Timestamp": ts, "Valor": val, "Tag": tag})
            procesados += 1
        ui.progreso(procesados, total_puntos)

    ui.estado("Combinando series: alineando por tiempo...")
    ui.progreso_indeterminado(True)

    df = pd.DataFrame(filas_largo)
    df = df.dropna(subset=["Valor"])
    df["Timestamp"] = pd.to_datetime(df["Timestamp"], errors="coerce")
    df = df.dropna(subset=["Timestamp"])

    pivot = df.pivot_table(index="Timestamp", columns="Tag", values="Valor", aggfunc="last")
    pivot = pivot.sort_index()
    pivot = pivot.ffill().bfill()

    resultado = pivot.reset_index()
    resultado.columns.name = None

    ui.progreso_indeterminado(False)

    ui.log(f"Filas combinadas: {len(resultado)}")
    ui.log(f"Columnas (tags): {len(resultado.columns) - 1}")

    return resultado


def trabajo_principal(ui, carpeta_entrada, carpeta_salida):
    """Corre en un hilo aparte para no congelar la ventana. No toca
    COM/PI para nada -- solo lee CSVs de disco y combina con pandas."""
    try:
        if pd is None:
            ui.log("ERROR: falta pandas en este build.")
            ui.terminar(exito=False)
            return

        series = combinar_desde_carpeta(carpeta_entrada, ui)
        resultado = combinar_series(series, ui)

        if resultado is None:
            ui.log("No se genero ningun archivo combinado (sin datos validos).")
            ui.terminar(exito=False)
            return

        nombre_salida = f"Combinado_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        ruta_salida = os.path.join(carpeta_salida, nombre_salida)
        resultado.to_csv(ruta_salida, index=False, encoding="utf-8")

        ui.log(f"\n=== LISTO ===")
        ui.log(f"Archivo combinado: {ruta_salida}")
        ui.terminar(exito=True)

        try:
            os.startfile(ruta_salida)
        except Exception:
            pass

    except Exception as e:
        ui.log(f"\nERROR INESPERADO: {e}")
        ui.terminar(exito=False)


def elegir_carpeta_con_dialogo():
    """Se muestra solo cuando el .exe se abre con doble clic (sin
    argumentos de linea de comandos): pide la carpeta con los CSV."""
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()
    carpeta = filedialog.askdirectory(title="Selecciona la carpeta con los CSV de PI a combinar")
    root.destroy()
    return carpeta if carpeta else None


def main():
    if len(sys.argv) >= 2:
        carpeta_entrada = sys.argv[1]
        if not os.path.isdir(carpeta_entrada):
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("ABSPIEXCELEXTRACT", f"No se encontro la carpeta:\n{carpeta_entrada}")
            sys.exit(1)
        carpeta_salida = sys.argv[2] if len(sys.argv) > 2 else carpeta_entrada
    else:
        carpeta_entrada = elegir_carpeta_con_dialogo()
        if not carpeta_entrada:
            sys.exit(0)
        carpeta_salida = carpeta_entrada

    os.makedirs(carpeta_salida, exist_ok=True)

    ui = VentanaProgreso()

    hilo = threading.Thread(target=trabajo_principal, args=(ui, carpeta_entrada, carpeta_salida), daemon=True)
    hilo.start()

    ui.iniciar_loop()


if __name__ == "__main__":
    main()
