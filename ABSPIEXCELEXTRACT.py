"""
ABSPIEXCELEXTRACT - Extractor + Combinador de datos PI

A partir de un archivo de parametros JSON generado por la macro VBA
"GenerarParametrosExtraccion", hace en un solo paso:
  1. Se conecta a PI-SDK y extrae cada tag con timestamps reales
     (RecordedValues, igual que hacia la macro VBA).
  2. Combina todas las series en una sola tabla, usando el mismo
     enfoque "forward-fill + backfill" que se valido en el flujo
     de Excel/Python (pivot por Tag, sin inventar fechas).
  3. Escribe un unico CSV combinado, listo para usar.

Muestra una ventanita de progreso (Tkinter) con el avance en vivo,
para poder compilarse con --noconsole y aun asi ver que esta pasando.

Pensado para compilarse como ABSPIEXCELEXTRACT.exe con PyInstaller,
de forma que el usuario final solo reciba el ejecutable y el JSON
de parametros, sin ver ni poder modificar la logica interna.

Requisitos para COMPILAR (no para el usuario final):
    pip install pywin32 pandas pyinstaller
    pyinstaller --onefile --noconsole --name ABSPIEXCELEXTRACT ABSPIEXCELEXTRACT.py

Este .exe debe quedar instalado en: C:\\ABSTOOLS\\ABSPIEXCELEXTRACT.exe
(esa es la ruta que espera el modulo VBA "GenerarParametrosExtraccion"
si el usuario elige lanzarlo automaticamente desde ProcessBook).

Uso manual (una vez compilado):
    ABSPIEXCELEXTRACT.exe parametros_extraccion_XXXX.json [carpeta_salida]
"""
import sys
import os
import json
import csv
import threading
from datetime import datetime

try:
    import win32com.client
    import pythoncom
except ImportError:
    win32com = None
    pythoncom = None

try:
    import pandas as pd
except ImportError:
    pd = None

import tkinter as tk
from tkinter import messagebox


class VentanaProgreso:
    """Ventanita de progreso simple: un titulo, una barra de estado
    y una caja de texto tipo log. No usa mainloop() de forma
    bloqueante -- se actualiza manualmente con update() desde el
    hilo de trabajo, igual que en el combinador original."""

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("ABSPIEXCELEXTRACT")
        self.root.geometry("640x420")
        self.root.resizable(True, True)

        titulo = tk.Label(self.root, text="ABSPIEXCELEXTRACT",
                           font=("Segoe UI", 14, "bold"))
        titulo.pack(pady=(12, 0))

        self.estado_var = tk.StringVar(value="Iniciando...")
        estado_lbl = tk.Label(self.root, textvariable=self.estado_var,
                               font=("Segoe UI", 10))
        estado_lbl.pack(pady=(4, 8))

        frame_log = tk.Frame(self.root)
        frame_log.pack(fill="both", expand=True, padx=12, pady=(0, 12))

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
        # Evita que cierren la ventana a medias mientras esta trabajando
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

    def terminar(self, exito=True):
        self.estado("Completado" if exito else "Termino con errores")
        self.boton_cerrar.config(state="normal")

    def iniciar_loop(self):
        self.root.mainloop()


def cargar_parametros(ruta_json):
    with open(ruta_json, "r", encoding="utf-8") as f:
        return json.load(f)


def limpiar_nombre_tag(tag):
    if "\\" in tag:
        return tag.split("\\")[-1]
    return tag


def extraer_series(params, ui):
    """Se conecta a PI y extrae cada tag. Devuelve una lista de
    (tag, [(datetime, valor), ...]) -- solo en memoria, sin CSV
    intermedios por tag."""
    servidor_nombre = params["server"]
    tags = params["tags"]
    utc_start = params["start_utc_seconds"]
    utc_end = params["end_utc_seconds"]

    ui.log(f"Servidor: {servidor_nombre}")
    ui.log(f"Tags a extraer: {len(tags)}")
    ui.estado("Conectando a PI...")

    try:
        pisdk = win32com.client.Dispatch("PISDK.PISDK")
    except Exception as e:
        ui.log(f"ERROR: no se pudo crear el objeto PISDK.PISDK. "
               f"Verifica que PI-SDK este instalado en esta maquina.\n{e}")
        return []

    try:
        server = pisdk.Servers(servidor_nombre)
    except Exception as e:
        ui.log(f"ERROR: no se pudo conectar al servidor '{servidor_nombre}'.\n{e}")
        return []

    ts_start = win32com.client.Dispatch("PITimeServer.PITime")
    ts_end = win32com.client.Dispatch("PITimeServer.PITime")
    ts_start.UTCSeconds = utc_start
    ts_end.UTCSeconds = utc_end

    series = []
    ui.log("\n=== Extrayendo trazas ===")
    for i, tag_raw in enumerate(tags, start=1):
        tag = limpiar_nombre_tag(tag_raw)
        ui.estado(f"Extrayendo {i}/{len(tags)}: {tag}")

        try:
            point = server.PIPoints(tag)
        except Exception:
            ui.log(f"[{i}/{len(tags)}] {tag}: TAG NO ENCONTRADO -- se omite")
            continue

        try:
            values = point.Data.RecordedValues(ts_start, ts_end)
        except Exception as e:
            ui.log(f"[{i}/{len(tags)}] {tag}: ERROR ({e}) -- se omite")
            continue

        puntos = []
        for v in values:
            try:
                ts_local = datetime.strptime(str(v.TimeStamp.LocalDate), "%m/%d/%Y %I:%M:%S %p")
            except ValueError:
                try:
                    ts_local = pd.to_datetime(str(v.TimeStamp.LocalDate))
                except Exception:
                    continue
            try:
                val = float(v.Value)
            except (TypeError, ValueError):
                continue
            puntos.append((ts_local, val))

        ui.log(f"[{i}/{len(tags)}] {tag}: {len(puntos)} puntos")
        if puntos:
            series.append((tag, puntos))

    return series


def combinar_series(series, ui):
    """Combina las series extraidas en una sola tabla, alineando
    por 'ultimo valor real conocido' (igual que PI Trend), con
    relleno hacia atras (backfill) al inicio de cada serie."""
    ui.estado("Combinando series...")
    ui.log("\n=== Combinando series ===")

    if not series:
        ui.log("No hay series validas para combinar.")
        return None

    filas_largo = []
    for tag, puntos in series:
        for ts, val in puntos:
            filas_largo.append({"Timestamp": ts, "Valor": val, "Tag": tag})

    df = pd.DataFrame(filas_largo)
    df = df.dropna(subset=["Valor"])
    df["Timestamp"] = pd.to_datetime(df["Timestamp"], errors="coerce")
    df = df.dropna(subset=["Timestamp"])

    pivot = df.pivot_table(index="Timestamp", columns="Tag", values="Valor", aggfunc="last")
    pivot = pivot.sort_index()
    pivot = pivot.ffill().bfill()  # PI-style: mantener ultimo valor real; backfill solo al inicio

    resultado = pivot.reset_index()
    resultado.columns.name = None

    ui.log(f"Filas combinadas: {len(resultado)}")
    ui.log(f"Columnas (tags): {len(resultado.columns) - 1}")

    return resultado


def trabajo_principal(ui, ruta_json, carpeta_salida):
    """Corre en un hilo aparte para no congelar la ventana.
    IMPORTANTE: los objetos COM (PISDK, PITimeServer) requieren que
    el hilo donde se usan tenga COM inicializado -- por default solo
    el hilo principal lo tiene. Como esta funcion corre en un
    threading.Thread aparte, hay que inicializar COM aqui mismo con
    pythoncom.CoInitialize() antes de crear cualquier objeto COM, y
    liberarlo con CoUninitialize() al terminar."""
    com_inicializado = False
    try:
        if win32com is None:
            ui.log("ERROR: falta pywin32 en este build.")
            ui.terminar(exito=False)
            return
        if pd is None:
            ui.log("ERROR: falta pandas en este build.")
            ui.terminar(exito=False)
            return

        pythoncom.CoInitialize()
        com_inicializado = True

        params = cargar_parametros(ruta_json)
        series = extraer_series(params, ui)
        resultado = combinar_series(series, ui)

        if resultado is None:
            ui.log("No se genero ningun archivo (sin datos validos).")
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
    finally:
        if com_inicializado:
            pythoncom.CoUninitialize()


def main():
    if len(sys.argv) < 2:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "ABSPIEXCELEXTRACT",
            "Uso: ABSPIEXCELEXTRACT.exe <archivo_parametros.json> [carpeta_salida]"
        )
        sys.exit(1)

    ruta_json = sys.argv[1]
    if not os.path.isfile(ruta_json):
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("ABSPIEXCELEXTRACT", f"No se encontro el archivo:\n{ruta_json}")
        sys.exit(1)

    carpeta_salida = sys.argv[2] if len(sys.argv) > 2 else os.path.dirname(os.path.abspath(ruta_json))
    os.makedirs(carpeta_salida, exist_ok=True)

    ui = VentanaProgreso()

    hilo = threading.Thread(target=trabajo_principal, args=(ui, ruta_json, carpeta_salida), daemon=True)
    hilo.start()

    ui.iniciar_loop()


if __name__ == "__main__":
    main()
