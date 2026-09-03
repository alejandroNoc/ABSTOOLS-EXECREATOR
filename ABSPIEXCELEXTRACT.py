"""
ABSPIEXCELEXTRACT - Extractor + Combinador de datos PI

A partir de un archivo de parametros JSON generado por la macro VBA
"GenerarParametrosExtraccion", hace en un solo paso:
  1. Se conecta a PI-SDK y extrae cada tag con timestamps reales
     (RecordedValues, igual que hacia la macro VBA).
  2. Escribe un CSV individual por tag (formato TAGn_nombre_conteo.csv),
     con las columnas Timestamp,Valor,Tag -- util para inspeccionar o
     auditar cada traza por separado.
  3. Combina todas las series en una sola tabla, usando el mismo
     enfoque "forward-fill + backfill" que se valido en el flujo
     de Excel/Python (pivot por Tag, sin inventar fechas).
  4. Escribe un unico CSV combinado, listo para usar.

Si se abre con DOBLE CLIC (sin argumentos de linea de comandos),
pregunta si se quiere:
  - Abrir un archivo de parametros .json (extrae de PI y combina), o
  - Abrir una carpeta con CSV ya existentes (solo combina, sin PI --
    util si ya se habian extraido los tags en otra corrida y solo
    hace falta rearmar el combinado).

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

Uso con doble clic: se abre el dialogo de seleccion descrito arriba.

Uso por linea de comandos (para el lanzamiento automatico desde VBA,
o manual): el primer argumento puede ser un archivo .json (modo
extraer+combinar) o una carpeta con CSV ya existentes (modo solo
combinar) -- se detecta automaticamente segun cual sea:
    ABSPIEXCELEXTRACT.exe parametros_extraccion_XXXX.json [carpeta_salida]
    ABSPIEXCELEXTRACT.exe "C:\\ruta\\carpeta_con_csvs" [carpeta_salida]
"""
import sys
import os
import json
import csv
import threading
from datetime import datetime

try:
    import win32com.client
    import win32com.client.gencache
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
from tkinter import ttk


def _preparar_cache_com():
    """Cuando este script corre como .exe compilado (PyInstaller), la
    cache de tipos que necesita gencache.EnsureDispatch puede intentar
    escribirse dentro del paquete empaquetado (de solo lectura) y fallar
    en silencio o de forma rara. Forzamos que use una carpeta temporal
    del usuario, que siempre es escribible."""
    if win32com is None:
        return None
    if getattr(sys, "frozen", False):
        import tempfile
        cache_dir = os.path.join(tempfile.gettempdir(), "abspiexcelextract_gen_py")
        os.makedirs(cache_dir, exist_ok=True)
        win32com.__gen_path__ = cache_dir
        return cache_dir
    return None


_cache_com_dir = None  # se completa en main() -> _preparar_cache_com()


def _limpiar_cache_com(cache_dir):
    """Borra la carpeta de cache de tipos COM (gen_py) Y limpia las
    referencias en memoria (sys.modules, registro interno de gencache).
    Solo borrar el archivo en disco NO alcanza: dentro del mismo proceso,
    Python ya dejo una referencia rota en sys.modules tras el primer
    intento fallido, y EnsureDispatch la sigue usando aunque el archivo
    ya no exista en disco -- por eso hay que limpiar tambien la memoria
    antes de reintentar."""
    import shutil
    import importlib

    # 1) Borrar del disco
    if cache_dir and os.path.isdir(cache_dir):
        try:
            shutil.rmtree(cache_dir, ignore_errors=True)
        except Exception:
            pass
        try:
            os.makedirs(cache_dir, exist_ok=True)
        except Exception:
            pass

    # 2) Quitar del cache de modulos ya importados en este proceso
    for nombre in list(sys.modules.keys()):
        if nombre.startswith("win32com.gen_py"):
            del sys.modules[nombre]

    # 3) Resetear el registro interno de gencache (que "recuerda" que
    # ya genero/importo ese tipo, aunque haya fallado). Se hace atributo
    # por atributo con try/except individual, por si alguno no existe
    # en la version de pywin32 instalada (no debe tumbar el reintento).
    for attr in ("dict", "dict_class_to_typelib", "clsidToPackageMap"):
        try:
            getattr(win32com.client.gencache, attr).clear()
        except Exception:
            pass

    # 4) Invalidar el cache de importacion de Python, para que vuelva
    # a mirar el disco (con la carpeta ya vacia/regenerada) en vez de
    # asumir que ya sabe que hay ahi
    importlib.invalidate_caches()


class VentanaProgreso:
    """Ventanita de progreso: titulo, barra de estado (texto), barra
    de progreso visual (determinate mientras se conoce el total --
    tag por tag, archivo por archivo -- e indeterminate/animada
    durante la combinacion, que es una operacion vectorizada sin
    pasos individuales que mostrar), y una caja de texto tipo log.
    No usa mainloop() de forma bloqueante -- se actualiza manualmente
    con update() desde el hilo de trabajo."""

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

        # --- Barra de progreso visual ---
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

    def progreso(self, valor, maximo):
        """Barra determinate: usar cuando se conoce el total de pasos
        (ej. extrayendo tag i de N, o leyendo archivo i de N)."""
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
        """Barra animada (va y viene): usar durante pasos sin un total
        conocido de antemano, como el pivot_table de pandas (es
        vectorizado y casi instantaneo, no tiene 'items' que contar)."""
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


def cargar_parametros(ruta_json):
    with open(ruta_json, "r", encoding="utf-8") as f:
        return json.load(f)


def limpiar_nombre_tag(tag):
    if "\\" in tag:
        return tag.split("\\")[-1]
    return tag


def limpiar_nombre_para_archivo(nombre_tag):
    """Reemplaza caracteres no permitidos en nombres de archivo de Windows
    (los tags de PI suelen traer ':' -- ej. 'TANK:15J03YB52COD.LV' --
    que NO es valido en un nombre de archivo)."""
    invalidos = [":", "\\", "/", "*", "?", '"', "<", ">", "|"]
    r = nombre_tag
    for ch in invalidos:
        r = r.replace(ch, "_")
    return r


def extraer_series(params, ui, carpeta_salida):
    """Se conecta a PI y extrae cada tag. Ademas de devolver todo en
    memoria para el combinado, escribe un CSV individual por tag
    (formato TAGn_nombre_conteo.csv), igual que hacia la macro VBA."""
    servidor_nombre = params["server"]
    tags = params["tags"]
    utc_start = params["start_utc_seconds"]
    utc_end = params["end_utc_seconds"]

    ui.log(f"Servidor: {servidor_nombre}")
    ui.log(f"Tags a extraer: {len(tags)}")
    ui.estado("Conectando a PI...")

    # IMPORTANTE: usamos gencache.EnsureDispatch (enlace temprano) en vez de
    # win32com.client.Dispatch (enlace tardio/generico) para PISDK y PITime.
    # Con Dispatch generico, pasar un objeto COM (como ts_start/ts_end) como
    # ARGUMENTO de otro metodo COM (RecordedValues) a veces falla con
    # "The Python instance can not be converted to a COM object", porque el
    # wrapper generico no conoce la firma exacta del metodo. EnsureDispatch
    # genera un wrapper con la definicion real del tipo (equivalente a usar
    # "Dim x As PISDK.PISDK" en VBA en vez de "CreateObject"), lo que
    # resuelve el problema de raiz.
    #
    # Si la cache de tipos (gen_py) quedo corrupta/desincronizada (ej. tras
    # actualizar el .exe), EnsureDispatch falla con "No module named
    # win32com.gen_py.<CLSID>...". En ese caso se borra la cache y se
    # reintenta UNA vez, regenerandola desde cero automaticamente, sin que
    # el usuario tenga que borrar nada a mano en %TEMP%.
    pisdk = None
    for intento in range(2):
        try:
            pisdk = win32com.client.gencache.EnsureDispatch("PISDK.PISDK")
            break
        except Exception as e:
            es_error_cache = "gen_py" in str(e) or "No module named" in str(e)
            if intento == 0 and es_error_cache:
                ui.log("Cache de COM desactualizada, regenerando automaticamente...")
                _limpiar_cache_com(_cache_com_dir)
                continue
            ui.log(f"ERROR: no se pudo crear el objeto PISDK.PISDK. "
                   f"Verifica que PI-SDK este instalado en esta maquina.\n{e}")
            return []
    if pisdk is None:
        return []

    try:
        server = pisdk.Servers(servidor_nombre)
    except Exception as e:
        ui.log(f"ERROR: no se pudo conectar al servidor '{servidor_nombre}'.\n{e}")
        return []

    ts_start = ts_end = None
    for intento in range(2):
        try:
            ts_start = win32com.client.gencache.EnsureDispatch("PITimeServer.PITime")
            ts_end = win32com.client.gencache.EnsureDispatch("PITimeServer.PITime")
            break
        except Exception as e:
            es_error_cache = "gen_py" in str(e) or "No module named" in str(e)
            if intento == 0 and es_error_cache:
                ui.log("Cache de COM desactualizada (PITime), regenerando automaticamente...")
                _limpiar_cache_com(_cache_com_dir)
                continue
            ui.log(f"ERROR: no se pudo crear el objeto PITimeServer.PITime.\n{e}")
            return []
    if ts_start is None or ts_end is None:
        return []

    ts_start.UTCSeconds = utc_start
    ts_end.UTCSeconds = utc_end

    series = []
    ui.log("\n=== Extrayendo trazas ===")
    ui.progreso(0, len(tags))
    for i, tag_raw in enumerate(tags, start=1):
        tag = limpiar_nombre_tag(tag_raw)
        ui.estado(f"Extrayendo {i}/{len(tags)}: {tag}")
        nombre_archivo_seguro = limpiar_nombre_para_archivo(tag)

        try:
            point = server.PIPoints(tag)
        except Exception:
            ui.log(f"[{i}/{len(tags)}] {tag}: TAG NO ENCONTRADO -- se omite")
            _escribir_csv_individual(carpeta_salida, i, nombre_archivo_seguro, 0,
                                      filas_error=[("", "", f"TAG NO ENCONTRADO: {tag}")])
            ui.progreso(i, len(tags))
            continue

        try:
            values = point.Data.RecordedValues(ts_start, ts_end)
        except Exception as e:
            ui.log(f"[{i}/{len(tags)}] {tag}: ERROR ({e}) -- se omite")
            _escribir_csv_individual(carpeta_salida, i, nombre_archivo_seguro, 0,
                                      filas_error=[("", "", f"ERROR: {e}")])
            ui.progreso(i, len(tags))
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

        # CSV individual de esta traza, con el conteo real ya en el nombre
        _escribir_csv_individual(carpeta_salida, i, nombre_archivo_seguro, len(puntos),
                                  puntos=puntos, tag_completo=tag)

        if puntos:
            series.append((tag, puntos))

        ui.progreso(i, len(tags))

    return series


def _escribir_csv_individual(carpeta_salida, indice, nombre_archivo_seguro, conteo,
                              puntos=None, tag_completo="", filas_error=None):
    """Escribe el CSV individual de una traza: TAGn_nombre_conteo.csv
    Mismo formato/contenido que generaba la macro VBA (columnas
    Timestamp,Valor,Tag), asi los archivos siguen siendo compatibles
    con cualquier otra herramienta que ya espere ese formato."""
    nombre = f"TAG{indice}_{nombre_archivo_seguro}_{conteo}.csv"
    ruta = os.path.join(carpeta_salida, nombre)
    with open(ruta, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["Timestamp", "Valor", "Tag"])
        if filas_error:
            for fila in filas_error:
                writer.writerow(fila)
        elif puntos:
            for ts, val in puntos:
                writer.writerow([ts.strftime("%Y-%m-%d %H:%M:%S"), val, tag_completo])


def leer_csv_individual(ruta_csv):
    """Lee un CSV individual (formato Timestamp,Valor,Tag) generado
    por este mismo programa o por la macro VBA. Devuelve (tag, puntos)
    o (None, []) si el archivo no tiene datos utilizables (tag no
    encontrado, error, o vacio)."""
    tag = os.path.splitext(os.path.basename(ruta_csv))[0]
    puntos = []
    tag_real = None

    formatos = ["%Y-%m-%d %H:%M:%S", "%m/%d/%Y %I:%M:%S %p", "%d/%m/%Y %H:%M:%S"]

    with open(ruta_csv, "r", encoding="utf-8-sig", errors="ignore") as f:
        reader = csv.reader(f)
        header = next(reader, None)
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
    """Modo 'solo combinar': lee todos los CSV individuales (formato
    TAGn_nombre_conteo.csv, o cualquier CSV con columnas
    Timestamp,Valor,Tag) que ya existan en una carpeta, y arma el
    combinado -- sin tocar PI para nada. Devuelve la lista 'series'
    en el mismo formato que usa combinar_series()."""
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
    """Combina las series extraidas en una sola tabla, alineando
    por 'ultimo valor real conocido' (igual que PI Trend), con
    relleno hacia atras (backfill) al inicio de cada serie."""
    ui.estado("Combinando series...")
    ui.log("\n=== Combinando series ===")

    if not series:
        ui.log("No hay series validas para combinar.")
        return None

    # --- Armar la tabla larga: progreso determinate (se conoce el total de puntos) ---
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

    # --- Pivot/alineacion: operacion vectorizada de pandas, sin pasos
    # individuales que mostrar -> barra animada (indeterminate) ---
    ui.estado("Combinando series: alineando por tiempo...")
    ui.progreso_indeterminado(True)

    df = pd.DataFrame(filas_largo)
    df = df.dropna(subset=["Valor"])
    df["Timestamp"] = pd.to_datetime(df["Timestamp"], errors="coerce")
    df = df.dropna(subset=["Timestamp"])

    pivot = df.pivot_table(index="Timestamp", columns="Tag", values="Valor", aggfunc="last")
    pivot = pivot.sort_index()
    pivot = pivot.ffill().bfill()  # PI-style: mantener ultimo valor real; backfill solo al inicio

    resultado = pivot.reset_index()
    resultado.columns.name = None

    ui.progreso_indeterminado(False)

    ui.log(f"Filas combinadas: {len(resultado)}")
    ui.log(f"Columnas (tags): {len(resultado.columns) - 1}")

    return resultado


def trabajo_principal(ui, modo, ruta_entrada, carpeta_salida):
    """Corre en un hilo aparte para no congelar la ventana.
    modo = 'extraer' (ruta_entrada es un JSON de parametros: se
           conecta a PI, extrae, y combina) o
           'combinar' (ruta_entrada es una carpeta con CSV ya
           existentes: solo los lee y combina, sin tocar PI).
    IMPORTANTE: los objetos COM (PISDK, PITimeServer) requieren que
    el hilo donde se usan tenga COM inicializado -- por default solo
    el hilo principal lo tiene. Como esta funcion corre en un
    threading.Thread aparte, hay que inicializar COM aqui mismo con
    pythoncom.CoInitialize() antes de crear cualquier objeto COM, y
    liberarlo con CoUninitialize() al terminar. En modo 'combinar'
    no hace falta COM en absoluto (no se toca PI), pero se inicializa
    igual por si acaso alguna libreria lo requiere de forma indirecta."""
    com_inicializado = False
    try:
        if pd is None:
            ui.log("ERROR: falta pandas en este build.")
            ui.terminar(exito=False)
            return

        if modo == "extraer":
            if win32com is None:
                ui.log("ERROR: falta pywin32 en este build.")
                ui.terminar(exito=False)
                return

            pythoncom.CoInitialize()
            com_inicializado = True

            params = cargar_parametros(ruta_entrada)
            series = extraer_series(params, ui, carpeta_salida)
        else:  # modo == "combinar"
            series = combinar_desde_carpeta(ruta_entrada, ui)

        resultado = combinar_series(series, ui)

        if resultado is None:
            ui.log("No se genero ningun archivo combinado (sin datos validos).")
            if modo == "extraer":
                ui.log(f"Revisa los CSV individuales por tag en: {carpeta_salida}")
            ui.terminar(exito=False)
            return

        nombre_salida = f"Combinado_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        ruta_salida = os.path.join(carpeta_salida, nombre_salida)
        resultado.to_csv(ruta_salida, index=False, encoding="utf-8")

        ui.log(f"\n=== LISTO ===")
        if modo == "extraer":
            ui.log(f"CSV individuales por tag: {carpeta_salida}")
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


def elegir_entrada_con_dialogo():
    """Se muestra solo cuando el .exe se abre con doble clic (sin
    argumentos de linea de comandos). Pregunta si el usuario quiere
    trabajar desde un archivo JSON de parametros (extrae de PI y
    combina) o desde una carpeta con CSV ya existentes (solo combina).
    Devuelve (modo, ruta_entrada, carpeta_salida) o (None, None, None)
    si el usuario cancela."""
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()

    opcion = messagebox.askyesnocancel(
        "ABSPIEXCELEXTRACT",
        "Que queres hacer?\n\n"
        "SI = Abrir un archivo de parametros (.json) -> extrae de PI y combina\n"
        "NO = Abrir una carpeta con CSV ya existentes -> solo combina\n"
        "CANCELAR = Salir"
    )

    if opcion is None:
        root.destroy()
        return None, None, None

    if opcion:  # Si -> JSON
        ruta_json = filedialog.askopenfilename(
            title="Selecciona el archivo de parametros (.json)",
            filetypes=[("Archivos JSON", "*.json"), ("Todos los archivos", "*.*")]
        )
        root.destroy()
        if not ruta_json:
            return None, None, None
        carpeta_salida = os.path.dirname(os.path.abspath(ruta_json))
        return "extraer", ruta_json, carpeta_salida
    else:  # No -> carpeta con CSV
        carpeta = filedialog.askdirectory(
            title="Selecciona la carpeta con los CSV ya extraidos"
        )
        root.destroy()
        if not carpeta:
            return None, None, None
        return "combinar", carpeta, carpeta


def main():
    global _cache_com_dir
    _cache_com_dir = _preparar_cache_com()

    if len(sys.argv) >= 2:
        # Modo linea de comandos (lanzado desde VBA, o manual con argumentos).
        # El primer argumento puede ser:
        #   - un archivo .json de parametros -> modo "extraer" (PI + combina)
        #   - una carpeta con CSV ya existentes -> modo "combinar" (sin PI)
        # Se detecta automaticamente segun que es (archivo vs carpeta).
        ruta_entrada = sys.argv[1]

        if os.path.isdir(ruta_entrada):
            modo = "combinar"
            carpeta_salida = sys.argv[2] if len(sys.argv) > 2 else ruta_entrada
        elif os.path.isfile(ruta_entrada):
            modo = "extraer"
            carpeta_salida = sys.argv[2] if len(sys.argv) > 2 else os.path.dirname(os.path.abspath(ruta_entrada))
        else:
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("ABSPIEXCELEXTRACT",
                                  f"No se encontro el archivo ni la carpeta:\n{ruta_entrada}")
            sys.exit(1)
    else:
        # Doble clic sin argumentos: preguntar que quiere hacer
        modo, ruta_entrada, carpeta_salida = elegir_entrada_con_dialogo()
        if modo is None:
            sys.exit(0)  # el usuario cancelo, salir sin error

    os.makedirs(carpeta_salida, exist_ok=True)

    ui = VentanaProgreso()

    hilo = threading.Thread(target=trabajo_principal, args=(ui, modo, ruta_entrada, carpeta_salida), daemon=True)
    hilo.start()

    ui.iniciar_loop()


if __name__ == "__main__":
    main()
