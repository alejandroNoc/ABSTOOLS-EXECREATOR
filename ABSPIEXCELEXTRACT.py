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

Pensado para compilarse como ABSPIEXCELEXTRACT.exe con PyInstaller,
de forma que el usuario final solo reciba el ejecutable y el JSON
de parametros, sin ver ni poder modificar la logica interna.

Requisitos para COMPILAR (no para el usuario final):
    pip install pywin32 pandas pyinstaller
    pyinstaller --onefile --name ABSPIEXCELEXTRACT ABSPIEXCELEXTRACT.py

Este .exe debe quedar instalado en: C:\ABSTOOLS\ABSPIEXCELEXTRACT.exe
(esa es la ruta que espera el modulo VBA "GenerarParametrosExtraccion"
si el usuario elige lanzarlo automaticamente desde ProcessBook).

Uso manual (una vez compilado):
    ABSPIEXCELEXTRACT.exe parametros_extraccion_XXXX.json [carpeta_salida]
"""
import sys
import os
import json
import csv
from datetime import datetime

try:
    import win32com.client
except ImportError:
    print("ERROR: falta pywin32. Instalalo con: pip install pywin32")
    sys.exit(1)

try:
    import pandas as pd
except ImportError:
    print("ERROR: falta pandas. Instalalo con: pip install pandas")
    sys.exit(1)


def cargar_parametros(ruta_json):
    with open(ruta_json, "r", encoding="utf-8") as f:
        return json.load(f)


def limpiar_nombre_tag(tag):
    if "\\" in tag:
        return tag.split("\\")[-1]
    return tag


def extraer_series(params):
    """Se conecta a PI y extrae cada tag. Devuelve una lista de
    (tag, [(datetime, valor), ...]) -- solo en memoria, sin CSV
    intermedios por tag."""
    servidor_nombre = params["server"]
    tags = params["tags"]
    utc_start = params["start_utc_seconds"]
    utc_end = params["end_utc_seconds"]

    print("=== Conectando a PI ===")
    print(f"Servidor: {servidor_nombre}")
    print(f"Tags a extraer: {len(tags)}")

    try:
        pisdk = win32com.client.Dispatch("PISDK.PISDK")
    except Exception as e:
        print(f"ERROR: no se pudo crear el objeto PISDK.PISDK. "
              f"Verifica que PI-SDK este instalado en esta maquina.\n{e}")
        sys.exit(1)

    try:
        server = pisdk.Servers(servidor_nombre)
    except Exception as e:
        print(f"ERROR: no se pudo conectar al servidor '{servidor_nombre}'.\n{e}")
        sys.exit(1)

    ts_start = win32com.client.Dispatch("PITimeServer.PITime")
    ts_end = win32com.client.Dispatch("PITimeServer.PITime")
    ts_start.UTCSeconds = utc_start
    ts_end.UTCSeconds = utc_end

    series = []
    print("\n=== Extrayendo trazas ===")
    for i, tag_raw in enumerate(tags, start=1):
        tag = limpiar_nombre_tag(tag_raw)
        print(f"[{i}/{len(tags)}] {tag} ...", end=" ")

        try:
            point = server.PIPoints(tag)
        except Exception:
            print("TAG NO ENCONTRADO -- se omite")
            continue

        try:
            values = point.Data.RecordedValues(ts_start, ts_end)
        except Exception as e:
            print(f"ERROR ({e}) -- se omite")
            continue

        puntos = []
        for v in values:
            try:
                ts_local = datetime.strptime(str(v.TimeStamp.LocalDate), "%m/%d/%Y %I:%M:%S %p")
            except ValueError:
                # respaldo por si el formato de LocalDate varia segun locale/version
                try:
                    ts_local = pd.to_datetime(str(v.TimeStamp.LocalDate))
                except Exception:
                    continue
            try:
                val = float(v.Value)
            except (TypeError, ValueError):
                continue
            puntos.append((ts_local, val))

        print(f"{len(puntos)} puntos")
        if puntos:
            series.append((tag, puntos))

    return series


def combinar_series(series):
    """Combina las series extraidas en una sola tabla, alineando
    por 'ultimo valor real conocido' (igual que PI Trend), con
    relleno hacia atras (backfill) al inicio de cada serie."""
    print("\n=== Combinando series ===")

    if not series:
        print("No hay series validas para combinar.")
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

    print(f"Filas combinadas: {len(resultado)}")
    print(f"Columnas (tags): {len(resultado.columns) - 1}")

    return resultado


def main():
    if len(sys.argv) < 2:
        print("ABSPIEXCELEXTRACT")
        print("Uso: ABSPIEXCELEXTRACT.exe <archivo_parametros.json> [carpeta_salida]")
        input("\nPresiona Enter para salir...")
        sys.exit(1)

    ruta_json = sys.argv[1]
    if not os.path.isfile(ruta_json):
        print(f"ERROR: no se encontro el archivo {ruta_json}")
        input("\nPresiona Enter para salir...")
        sys.exit(1)

    carpeta_salida = sys.argv[2] if len(sys.argv) > 2 else os.path.dirname(os.path.abspath(ruta_json))
    os.makedirs(carpeta_salida, exist_ok=True)

    params = cargar_parametros(ruta_json)

    series = extraer_series(params)
    resultado = combinar_series(series)

    if resultado is None:
        print("No se genero ningun archivo (sin datos validos).")
        input("\nPresiona Enter para salir...")
        sys.exit(1)

    nombre_salida = f"Combinado_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
    ruta_salida = os.path.join(carpeta_salida, nombre_salida)
    resultado.to_csv(ruta_salida, index=False, encoding="utf-8")

    print(f"\n=== LISTO ===")
    print(f"Archivo combinado: {ruta_salida}")
    input("\nPresiona Enter para salir...")


if __name__ == "__main__":
    main()
