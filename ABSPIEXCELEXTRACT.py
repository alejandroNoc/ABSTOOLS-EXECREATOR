"""
ABSPIEXCELEXTRACT - Extractor + Combinador de datos PI

A partir de un archivo de parametros JSON generado por la macro VBA
"GenerarParametrosExtraccion", hace en un solo paso:
  1. Lanza un script PowerShell (generado internamente, embebido en
     este mismo archivo) que se conecta a PI-SDK via COM NATIVO de
     Windows y extrae cada tag con timestamps reales -- exactamente
     el mismo mecanismo que usaba la macro VBA (New-Object -ComObject,
     igual que "CreateObject"/"New" en VBA). Se uso PowerShell para la
     extraccion en vez de Python/pywin32 porque, tras varias pruebas,
     pywin32 (tanto con gencache/enlace temprano como con Dispatch/
     enlace tardio) no logra pasar objetos PITime como argumento de
     RecordedValues() de forma confiable dentro de un .exe compilado
     con PyInstaller -- falla con "The Python instance can not be
     converted to a COM object" o con errores de cache de tipos
     ("No module named win32com.gen_py..."). PowerShell usa COM nativo
     de Windows (igual que VBA), sin esa capa fragil.
  2. El script de PowerShell escribe un CSV individual por tag
     (formato TAGn_nombre_conteo.csv, columnas Timestamp,Valor,Tag).
  3. Python lee esos CSV y combina todo en una sola tabla, con el
     mismo enfoque "forward-fill + backfill" ya validado (pivot por
     Tag, sin inventar fechas) -- esta parte SI funciona perfecto en
     Python/pandas, no tiene relacion con COM.
  4. Escribe un unico CSV combinado, listo para usar.

Si se abre con DOBLE CLIC (sin argumentos de linea de comandos),
pregunta si se quiere:
  - Abrir un archivo de parametros .json (extrae de PI y combina), o
  - Abrir una carpeta con CSV ya existentes (solo combina, sin PI --
    util si ya se habian extraido los tags en otra corrida y solo
    hace falta rearmar el combinado).

Muestra una ventanita de progreso (Tkinter) con el avance en vivo,
para poder compilarse con --noconsole y aun asi ver que esta pasando.

Requisitos en la maquina que EJECUTA el .exe:
  - PI-SDK instalado (provee los objetos COM "PISDK.PISDK" y
    "PITimeServer.PITime" que usa el script de PowerShell).
  - PowerShell 5.1+ (viene por defecto en Windows 10/11 y Server
    2016+, asi que en la practica no requiere instalar nada extra).
  - pandas (para la combinacion; se compila dentro del .exe).

Requisitos para COMPILAR (no para el usuario final):
    pip install pandas pyinstaller
    pyinstaller --onefile --noconsole --name ABSPIEXCELEXTRACT ABSPIEXCELEXTRACT.py

Ya NO se requiere pywin32 para compilar (la extraccion via PI-SDK
ahora la hace PowerShell, no Python).

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
import subprocess
import tempfile
import re
from datetime import datetime

try:
    import pandas as pd
except ImportError:
    pd = None

import tkinter as tk
from tkinter import messagebox
from tkinter import ttk


# ============================================================
# Script de PowerShell embebido: hace la extraccion de PI usando
# COM nativo de Windows (New-Object -ComObject), igual que VBA.
# Se escribe a un archivo temporal en tiempo de ejecucion y se
# invoca con subprocess -- asi no depende de --add-data de
# PyInstaller ni de rutas relativas al .exe.
# ============================================================
_POWERSHELL_SCRIPT = r'''
param(
    [Parameter(Mandatory=$true)][string]$ParamsJson,
    [Parameter(Mandatory=$true)][string]$OutFolder
)

$ErrorActionPreference = "Stop"

function Limpiar-NombreTag($tag) {
    if ($tag -match '\\') {
        $partes = $tag -split '\\'
        return $partes[$partes.Count - 1]
    }
    return $tag
}

function Limpiar-NombreArchivo($nombre) {
    $invalidos = @(':','\','/','*','?','"','<','>','|')
    $r = $nombre
    foreach ($ch in $invalidos) { $r = $r.Replace($ch, '_') }
    return $r
}

Write-Output "LOG:Leyendo parametros..."
$paramsRaw = Get-Content -Raw -Path $ParamsJson
$params = $paramsRaw | ConvertFrom-Json

$servidor = $params.server
$tags = $params.tags
$utcStart = [double]$params.start_utc_seconds
$utcEnd = [double]$params.end_utc_seconds

Write-Output "LOG:Servidor: $servidor"
Write-Output "LOG:Tags a extraer: $($tags.Count)"

try {
    $piSDK = New-Object -ComObject "PISDK.PISDK"
} catch {
    Write-Output "LOG:ERROR: no se pudo crear el objeto PISDK.PISDK. Verifica que PI-SDK este instalado en esta maquina."
    Write-Output "LOG:$($_.Exception.Message)"
    exit 1
}

try {
    $piServer = $piSDK.Servers.Item($servidor)
} catch {
    Write-Output "LOG:ERROR: no se pudo conectar al servidor '$servidor'."
    Write-Output "LOG:$($_.Exception.Message)"
    exit 1
}

$tsStart = New-Object -ComObject "PITimeServer.PITime"
$tsEnd = New-Object -ComObject "PITimeServer.PITime"
$tsStart.UTCSeconds = $utcStart
$tsEnd.UTCSeconds = $utcEnd

if (!(Test-Path $OutFolder)) { New-Item -ItemType Directory -Path $OutFolder | Out-Null }

$i = 0
$total = $tags.Count
foreach ($tagRaw in $tags) {
    $i++
    $tag = Limpiar-NombreTag $tagRaw
    $nombreSeguro = Limpiar-NombreArchivo $tag
    Write-Output "PROGRESO:$i/$total"
    Write-Output "LOG:[$i/$total] Extrayendo: $tag"

    $piPoint = $null
    try {
        $piPoint = $piServer.PIPoints.Item($tag)
    } catch {
        $piPoint = $null
    }

    if ($null -eq $piPoint) {
        $ruta = Join-Path $OutFolder "TAG${i}_${nombreSeguro}_0.csv"
        $sw = New-Object System.IO.StreamWriter($ruta, $false, [System.Text.Encoding]::UTF8)
        $sw.WriteLine("Timestamp,Valor,Tag")
        $sw.WriteLine(",,`"TAG NO ENCONTRADO: $tag`"")
        $sw.Close()
        Write-Output "LOG:[$i/$total] $tag : TAG NO ENCONTRADO"
        continue
    }

    $piValues = $null
    try {
        $piValues = $piPoint.Data.RecordedValues($tsStart, $tsEnd)
    } catch {
        $ruta = Join-Path $OutFolder "TAG${i}_${nombreSeguro}_0.csv"
        $sw = New-Object System.IO.StreamWriter($ruta, $false, [System.Text.Encoding]::UTF8)
        $sw.WriteLine("Timestamp,Valor,Tag")
        $sw.WriteLine(",,`"ERROR: $($_.Exception.Message)`"")
        $sw.Close()
        Write-Output "LOG:[$i/$total] $tag : ERROR ($($_.Exception.Message))"
        continue
    }

    $conteo = $piValues.Count
    $ruta = Join-Path $OutFolder "TAG${i}_${nombreSeguro}_${conteo}.csv"

    $sw = New-Object System.IO.StreamWriter($ruta, $false, [System.Text.Encoding]::UTF8)
    $sw.WriteLine("Timestamp,Valor,Tag")
    foreach ($v in $piValues) {
        try {
            $ts = [datetime]$v.TimeStamp.LocalDate
            $tsStr = $ts.ToString("yyyy-MM-dd HH:mm:ss")
            $sw.WriteLine("$tsStr,$($v.Value),`"$tag`"")
        } catch {
            # fila individual con dato raro -- se salta, no se aborta todo el tag
            continue
        }
    }
    $sw.Close()

    Write-Output "LOG:[$i/$total] $tag : $conteo puntos exportados"
}

Write-Output "LOG:=== EXTRACCION COMPLETADA ==="
'''


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


def cargar_parametros(ruta_json):
    with open(ruta_json, "r", encoding="utf-8") as f:
        return json.load(f)


def limpiar_nombre_para_archivo(nombre_tag):
    invalidos = [":", "\\", "/", "*", "?", '"', "<", ">", "|"]
    r = nombre_tag
    for ch in invalidos:
        r = r.replace(ch, "_")
    return r


def extraer_via_powershell(ruta_json, carpeta_salida, ui):
    """Escribe el script de PowerShell embebido a un archivo temporal
    y lo ejecuta, leyendo su salida linea por linea para actualizar
    el log y la barra de progreso en vivo. La extraccion real de PI
    ocurre DENTRO del proceso de PowerShell (COM nativo de Windows,
    igual que hacia la macro VBA) -- Python solo orquesta y muestra
    el avance, sin tocar COM para nada."""
    fd, ruta_ps1 = tempfile.mkstemp(suffix=".ps1", prefix="abspiexcelextract_")
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            f.write(_POWERSHELL_SCRIPT)

        cmd = [
            "powershell.exe",
            "-NoProfile",
            "-ExecutionPolicy", "Bypass",
            "-File", ruta_ps1,
            "-ParamsJson", ruta_json,
            "-OutFolder", carpeta_salida,
        ]

        ui.estado("Conectando a PI (via PowerShell)...")

        proceso = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            bufsize=1,
            creationflags=subprocess.CREATE_NO_WINDOW if hasattr(subprocess, "CREATE_NO_WINDOW") else 0,
        )

        patron_progreso = re.compile(r"^PROGRESO:(\d+)/(\d+)$")

        for linea in proceso.stdout:
            linea = linea.rstrip("\n").rstrip("\r")
            if not linea:
                continue

            m = patron_progreso.match(linea)
            if m:
                actual, total = int(m.group(1)), int(m.group(2))
                ui.progreso(actual, total)
                continue

            if linea.startswith("LOG:"):
                mensaje = linea[4:]
                ui.log(mensaje)
                if mensaje.startswith("[") and "Extrayendo:" in mensaje:
                    ui.estado(mensaje.lstrip("LOG:"))
                continue

            # cualquier otra linea (por si PowerShell escribe algo inesperado)
            ui.log(linea)

        proceso.wait()

        if proceso.returncode != 0:
            ui.log(f"\nEl proceso de extraccion (PowerShell) termino con codigo {proceso.returncode}.")

    finally:
        try:
            os.remove(ruta_ps1)
        except Exception:
            pass


def leer_csv_individual(ruta_csv):
    """Lee un CSV individual (formato Timestamp,Valor,Tag) generado
    por el script de PowerShell o por la macro VBA. Devuelve (tag,
    puntos), con puntos vacio si el archivo no tiene datos utilizables
    (tag no encontrado, error, o vacio)."""
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
    """Modo 'solo combinar': lee todos los CSV individuales (formato
    TAGn_nombre_conteo.csv, o cualquier CSV con columnas
    Timestamp,Valor,Tag) que ya existan en una carpeta, y arma la
    lista 'series' que espera combinar_series(). Tambien se usa
    despues del modo 'extraer', para leer los CSV que escribio el
    script de PowerShell."""
    archivos = [f for f in os.listdir(carpeta)
                if f.lower().endswith(".csv") and not f.lower().startswith("combinado")]

    if not archivos:
        ui.log("No se encontraron archivos CSV en la carpeta.")
        return []

    ui.log(f"\nLeyendo {len(archivos)} archivos CSV para combinar...\n")

    series = []
    ui.progreso(0, len(archivos))
    for i, nombre in enumerate(sorted(archivos), start=1):
        ruta = os.path.join(carpeta, nombre)
        ui.estado(f"Leyendo {i}/{len(archivos)}: {nombre}")
        tag, puntos = leer_csv_individual(ruta)
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


def trabajo_principal(ui, modo, ruta_entrada, carpeta_salida):
    """Corre en un hilo aparte para no congelar la ventana.
    modo = 'extraer' (ruta_entrada es un JSON de parametros: lanza
           PowerShell para extraer de PI, y luego combina) o
           'combinar' (ruta_entrada es una carpeta con CSV ya
           existentes: solo los lee y combina, sin tocar PI)."""
    try:
        if pd is None:
            ui.log("ERROR: falta pandas en este build.")
            ui.terminar(exito=False)
            return

        if modo == "extraer":
            extraer_via_powershell(ruta_entrada, carpeta_salida, ui)
            series = combinar_desde_carpeta(carpeta_salida, ui)
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


def elegir_entrada_con_dialogo():
    """Se muestra solo cuando el .exe se abre con doble clic (sin
    argumentos de linea de comandos)."""
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
    if len(sys.argv) >= 2:
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
        modo, ruta_entrada, carpeta_salida = elegir_entrada_con_dialogo()
        if modo is None:
            sys.exit(0)

    os.makedirs(carpeta_salida, exist_ok=True)

    ui = VentanaProgreso()

    hilo = threading.Thread(target=trabajo_principal, args=(ui, modo, ruta_entrada, carpeta_salida), daemon=True)
    hilo.start()

    ui.iniciar_loop()


if __name__ == "__main__":
    main()
