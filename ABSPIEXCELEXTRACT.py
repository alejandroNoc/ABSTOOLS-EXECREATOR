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
  3. Pregunta como se quiere el CSV combinado (dialogo con 3 botones):
       - Fragmentado: lo parte en varios CSV, cada uno con como maximo
         1,000,000 de filas, para poder abrirlos directo en Excel
         (que tiene un limite de ~1,048,576 filas por hoja).
       - Completo: un solo archivo, sin fragmentar, sin importar el
         limite de filas de Excel.
       - Ambos: escribe las dos versiones.
  4. Escribe el/los archivo(s) elegido(s) y, en el resumen final,
     pregunta si se quiere abrir el archivo (no lo abre solo).

Si se abre con DOBLE CLIC (sin argumentos), pregunta la carpeta con
los CSV a combinar.

Muestra una ventanita de progreso (Tkinter) con dos barras: una VERDE
de progreso GENERAL (avance de todo el trabajo, 0-100%) y una AZUL
CELESTE de progreso PARTICULAR (avance de la tarea puntual en curso:
archivo que se esta leyendo, o punto que se esta procesando). El
historial de texto detallado esta OCULTO por defecto -- se expande
con un enlace "Ver detalles" si hace falta revisarlo.

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
import math
import threading
from datetime import datetime

try:
    import pandas as pd
except ImportError:
    pd = None

import tkinter as tk
from tkinter import messagebox
from tkinter import ttk


# Limite de filas por archivo cuando se elige "fragmentado". Se usa un
# margen bajo el limite real de Excel (1,048,576 filas por hoja) para
# dejar espacio al encabezado y no quedar justo en el borde.
LIMITE_FILAS_EXCEL = 1000000

# Rango (inicio%, fin%) que ocupa cada fase dentro de la barra GENERAL.
# Son pesos aproximados, no una medicion exacta -- solo dan una nocion
# razonable de cuanto falta del trabajo total.
FASES_GENERAL = {
    "leer": (0, 35),        # leyendo los CSV de la carpeta
    "preparar": (35, 80),   # armando la tabla larga punto por punto
    "alinear": (80, 100),   # pivot_table + ffill/bfill (indeterminate)
}


class VentanaProgreso:
    """Ventanita de progreso con DOS barras:
      - General (verde): avance de todo el trabajo (0-100%), calculado
        combinando el peso de cada fase (FASES_GENERAL).
      - Particular (azul celeste): avance de la tarea puntual dentro de
        la fase actual (archivo i/N, punto i/N, o animada/indeterminate
        durante el pivot).
    Un historial de texto (log) queda oculto por defecto, expandible con
    un enlace "Ver detalles". No usa mainloop() de forma bloqueante --
    se actualiza manualmente con update() desde el hilo de trabajo."""

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("ABSPIEXCELEXTRACT")
        self.root.resizable(True, True)

        # El tema 'clam' es el que permite personalizar el color de
        # fondo de la barra en Windows de forma confiable (el tema
        # nativo por defecto suele ignorar 'background').
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("Green.Horizontal.TProgressbar", troughcolor="#e0e0e0",
                         background="#2e7d32", thickness=16)
        style.configure("Sky.Horizontal.TProgressbar", troughcolor="#e0e0e0",
                         background="#29b6f6", thickness=12)

        titulo = tk.Label(self.root, text="ABSPIEXCELEXTRACT",
                           font=("Segoe UI", 14, "bold"))
        titulo.pack(pady=(14, 2))

        self.estado_var = tk.StringVar(value="Iniciando...")
        estado_lbl = tk.Label(self.root, textvariable=self.estado_var,
                               font=("Segoe UI", 10))
        estado_lbl.pack(pady=(2, 8))

        # --- Barra GENERAL (verde) ---
        gen_frame = tk.Frame(self.root)
        gen_frame.pack(fill="x", padx=16, pady=(0, 2))
        tk.Label(gen_frame, text="General", font=("Segoe UI", 8), fg="#2e7d32",
                 width=8, anchor="w").pack(side="left")
        self.progress_general = ttk.Progressbar(gen_frame, orient="horizontal",
                                                  mode="determinate", maximum=100,
                                                  style="Green.Horizontal.TProgressbar")
        self.progress_general.pack(fill="x", side="left", expand=True)
        self.progress_general_pct_var = tk.StringVar(value="0%")
        tk.Label(gen_frame, textvariable=self.progress_general_pct_var,
                 font=("Segoe UI", 9), width=5).pack(side="left", padx=(6, 0))

        # --- Barra PARTICULAR (azul celeste) ---
        part_frame = tk.Frame(self.root)
        part_frame.pack(fill="x", padx=16, pady=(2, 6))
        tk.Label(part_frame, text="Actual", font=("Segoe UI", 8), fg="#0288d1",
                 width=8, anchor="w").pack(side="left")
        self.progress_particular = ttk.Progressbar(part_frame, orient="horizontal",
                                                     mode="determinate", maximum=100,
                                                     style="Sky.Horizontal.TProgressbar")
        self.progress_particular.pack(fill="x", side="left", expand=True)
        self.progress_particular_pct_var = tk.StringVar(value="")
        tk.Label(part_frame, textvariable=self.progress_particular_pct_var,
                 font=("Segoe UI", 9), width=5).pack(side="left", padx=(6, 0))

        # --- Enlace para expandir/colapsar el historial (oculto por defecto) ---
        self.log_visible = False
        self.toggle_btn = tk.Label(self.root, text="▼ Ver detalles",
                                    font=("Segoe UI", 8), fg="#555555", cursor="hand2")
        self.toggle_btn.pack(pady=(0, 4))
        self.toggle_btn.bind("<Button-1>", lambda e: self._toggle_log())

        self.frame_log = tk.Frame(self.root)
        scrollbar = tk.Scrollbar(self.frame_log)
        scrollbar.pack(side="right", fill="y")
        self.texto = tk.Text(self.frame_log, wrap="word", font=("Consolas", 9),
                              yscrollcommand=scrollbar.set)
        self.texto.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=self.texto.yview)

        self.boton_cerrar = tk.Button(self.root, text="Cerrar",
                                       command=self.root.destroy, state="disabled")
        self.boton_cerrar.pack(pady=(4, 12))

        self.root.protocol("WM_DELETE_WINDOW", self._on_close_intentado)

        self._fase_actual = None
        self._fase_ini, self._fase_fin = 0, 0

        self.root.update_idletasks()
        self.root.minsize(480, 0)
        self.root.geometry("")

    def _toggle_log(self):
        if self.log_visible:
            self.frame_log.pack_forget()
            self.toggle_btn.config(text="▼ Ver detalles")
            self.log_visible = False
        else:
            self.boton_cerrar.pack_forget()
            self.frame_log.pack(fill="both", expand=True, padx=16, pady=(0, 8))
            self.boton_cerrar.pack(pady=(4, 12))
            self.toggle_btn.config(text="▲ Ocultar detalles")
            self.log_visible = True
        self.root.update_idletasks()
        self.root.geometry("")

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

    def iniciar_fase(self, nombre):
        """Marca el inicio de una fase ('leer', 'preparar', 'alinear').
        Fija el rango que le corresponde a esa fase dentro de la barra
        GENERAL, y resetea la barra PARTICULAR a 0 para esta fase."""
        self._fase_actual = nombre
        self._fase_ini, self._fase_fin = FASES_GENERAL[nombre]
        self.progress_general["value"] = self._fase_ini
        self.progress_general_pct_var.set(f"{self._fase_ini}%")
        self.progress_particular.config(mode="determinate")
        self.progress_particular["value"] = 0
        self.progress_particular_pct_var.set("0%")
        self.root.update_idletasks()
        self.root.update()

    def progreso_particular(self, valor, maximo):
        """Actualiza la barra PARTICULAR (tarea puntual: archivo i/N,
        punto i/N) y, de forma proporcional, tambien la barra GENERAL
        dentro del rango de la fase actual."""
        if maximo <= 0:
            maximo = 1
        if self.progress_particular["mode"] != "determinate":
            self.progress_particular.stop()
            self.progress_particular.config(mode="determinate")
        self.progress_particular.config(maximum=maximo)
        self.progress_particular["value"] = valor

        fraccion = valor / maximo
        pct_particular = int(round(100 * fraccion))
        self.progress_particular_pct_var.set(f"{pct_particular}%")

        general_val = self._fase_ini + (self._fase_fin - self._fase_ini) * fraccion
        self.progress_general["value"] = general_val
        self.progress_general_pct_var.set(f"{int(round(general_val))}%")

        self.root.update_idletasks()
        self.root.update()

    def progreso_particular_indeterminado(self, activar=True):
        """Anima la barra PARTICULAR (va y viene) para pasos sin un
        total conocido de antemano (ej. el pivot_table de pandas)."""
        if activar:
            self.progress_particular.config(mode="indeterminate")
            self.progress_particular_pct_var.set("...")
            self.progress_particular.start(12)
        else:
            self.progress_particular.stop()
            self.progress_particular.config(mode="determinate")
        self.root.update_idletasks()
        self.root.update()

    def finalizar_fase(self):
        """Deja ambas barras en el 100% del tramo de la fase actual
        (util para fases indeterminadas, donde no hay forma de medir
        el avance intermedio -- se salta directo al final de la fase)."""
        self.progress_particular_indeterminado(False)
        self.progress_general["value"] = self._fase_fin
        self.progress_general_pct_var.set(f"{self._fase_fin}%")
        self.progress_particular["value"] = 100
        self.progress_particular_pct_var.set("100%")
        self.root.update_idletasks()
        self.root.update()

    def terminar(self, exito=True, resumen=None, archivo_a_abrir=None):
        """Marca el trabajo como terminado.
        Si exito=True y hay 'resumen': muestra un dialogo con el resumen
        del trabajo y PREGUNTA si se quiere abrir el archivo (Si/No). Al
        responder (cualquiera de las dos opciones), CIERRA toda la
        ventana -- no hace falta un boton 'Cerrar' aparte.
        Si hubo error: deja el log visible y habilita 'Cerrar' para que
        el usuario revise que paso antes de cerrar manualmente."""
        self.progress_general["value"] = 100 if exito else self.progress_general["value"]
        self.progress_general_pct_var.set("100%" if exito else self.progress_general_pct_var.get())
        self.progreso_particular_indeterminado(False)
        self.estado("Completado" if exito else "Termino con errores")

        if exito and resumen:
            pregunta = resumen + "\n\n¿Queres abrir el archivo combinado ahora?"
            abrir = messagebox.askyesno("ABSPIEXCELEXTRACT - Completado", pregunta, parent=self.root)
            if abrir and archivo_a_abrir:
                try:
                    os.startfile(archivo_a_abrir)
                except Exception:
                    pass
            self.root.destroy()
        else:
            self.boton_cerrar.config(state="normal")

    def preguntar_formato_salida(self):
        """Pregunta como quiere el usuario el CSV combinado:
        'fragmentado' (varios archivos, cada uno dentro del limite de
        filas de Excel, para abrirlos ahi directo), 'completo' (un solo
        archivo sin fragmentar), o 'ambos'. Dialogo modal con 3 botones.
        Devuelve None si se cierra sin elegir (cancelado)."""
        resultado = {"valor": None}

        top = tk.Toplevel(self.root)
        top.title("Formato de salida")
        top.transient(self.root)
        top.grab_set()
        top.resizable(False, False)

        tk.Label(top, text="Como queres el CSV combinado?",
                 font=("Segoe UI", 11, "bold")).pack(padx=20, pady=(16, 4))
        tk.Label(top,
                 text="Excel tiene un limite de ~1,048,576 filas por hoja.\n"
                      "Si el combinado supera eso, 'Fragmentado' lo parte\n"
                      "en varios CSV que si puedes abrir directo en Excel.",
                 font=("Segoe UI", 9), justify="left").pack(padx=20, pady=(0, 12))

        def elegir(valor):
            resultado["valor"] = valor
            top.destroy()

        frame_botones = tk.Frame(top)
        frame_botones.pack(pady=(0, 16), padx=20)

        tk.Button(frame_botones, text="Fragmentado\n(para Excel)", width=16,
                  command=lambda: elegir("fragmentado")).pack(side="left", padx=4)
        tk.Button(frame_botones, text="Completo\n(un solo archivo)", width=16,
                  command=lambda: elegir("completo")).pack(side="left", padx=4)
        tk.Button(frame_botones, text="Ambos", width=16,
                  command=lambda: elegir("ambos")).pack(side="left", padx=4)

        top.protocol("WM_DELETE_WINDOW", lambda: elegir(None))

        top.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - top.winfo_width()) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - top.winfo_height()) // 2
        top.geometry(f"+{x}+{y}")

        self.root.wait_window(top)
        return resultado["valor"]

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
    ui.iniciar_fase("leer")

    series = []
    ui.progreso_particular(0, len(archivos))
    for i, nombre in enumerate(sorted(archivos), start=1):
        ruta = os.path.join(carpeta, nombre)
        ui.estado(f"Leyendo {i}/{len(archivos)}: {nombre}")
        tag, puntos = leer_csv_individual(ruta)
        ui.log(f"[{i}/{len(archivos)}] {nombre} -> tag='{tag}', {len(puntos)} puntos")
        if puntos:
            series.append((tag, puntos))
        ui.progreso_particular(i, len(archivos))

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

    ui.iniciar_fase("preparar")
    total_puntos = sum(len(puntos) for _, puntos in series)
    filas_largo = []
    procesados = 0
    # Actualizar la barra cada cierto numero de puntos (no solo una vez
    # por tag) para que el avance se vea fluido incluso con tags que
    # tienen muchisimos puntos.
    intervalo_actualizacion = max(1, total_puntos // 200)
    ui.estado("Combinando series: preparando datos...")
    ui.progreso_particular(0, total_puntos)
    for tag, puntos in series:
        for ts, val in puntos:
            filas_largo.append({"Timestamp": ts, "Valor": val, "Tag": tag})
            procesados += 1
            if procesados % intervalo_actualizacion == 0:
                ui.progreso_particular(procesados, total_puntos)
    ui.progreso_particular(procesados, total_puntos)

    ui.iniciar_fase("alinear")
    ui.estado("Combinando series: alineando por tiempo...")
    ui.progreso_particular_indeterminado(True)

    df = pd.DataFrame(filas_largo)
    df = df.dropna(subset=["Valor"])
    df["Timestamp"] = pd.to_datetime(df["Timestamp"], errors="coerce")
    df = df.dropna(subset=["Timestamp"])

    pivot = df.pivot_table(index="Timestamp", columns="Tag", values="Valor", aggfunc="last")
    pivot = pivot.sort_index()
    pivot = pivot.ffill().bfill()

    resultado = pivot.reset_index()
    resultado.columns.name = None

    ui.finalizar_fase()

    ui.log(f"Filas combinadas: {len(resultado)}")
    ui.log(f"Columnas (tags): {len(resultado.columns) - 1}")

    return resultado


def escribir_csv_fragmentado(df, carpeta_salida, nombre_base, ui):
    """Parte el DataFrame combinado en varios CSV, cada uno con como
    maximo LIMITE_FILAS_EXCEL filas, para que se puedan abrir directo
    en Excel sin toparse con el limite de filas por hoja. Devuelve la
    lista de rutas escritas."""
    total_filas = len(df)
    n_partes = max(1, math.ceil(total_filas / LIMITE_FILAS_EXCEL))
    rutas = []

    for parte in range(n_partes):
        inicio = parte * LIMITE_FILAS_EXCEL
        fin = min(inicio + LIMITE_FILAS_EXCEL, total_filas)
        trozo = df.iloc[inicio:fin]
        nombre = f"{nombre_base}_parte{parte + 1}de{n_partes}.csv"
        ruta = os.path.join(carpeta_salida, nombre)
        trozo.to_csv(ruta, index=False, encoding="utf-8")
        ui.log(f"  Parte {parte + 1}/{n_partes}: {len(trozo)} filas -> {nombre}")
        rutas.append(ruta)

    return rutas


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

        nombre_base = f"Combinado_{datetime.now().strftime('%Y%m%d_%H%M%S')}"

        formato = ui.preguntar_formato_salida()
        if formato is None:
            ui.log("Operacion cancelada (no se eligio formato de salida).")
            ui.terminar(exito=False)
            return

        archivo_a_abrir = None
        archivos_generados = []

        if formato in ("completo", "ambos"):
            ruta_completo = os.path.join(carpeta_salida, f"{nombre_base}.csv")
            resultado.to_csv(ruta_completo, index=False, encoding="utf-8")
            ui.log(f"\nArchivo completo: {ruta_completo} ({len(resultado)} filas)")
            archivo_a_abrir = ruta_completo
            archivos_generados.append(os.path.basename(ruta_completo))

        if formato in ("fragmentado", "ambos"):
            ui.log(f"\nGenerando archivos fragmentados (limite {LIMITE_FILAS_EXCEL} filas c/u)...")
            rutas_frag = escribir_csv_fragmentado(resultado, carpeta_salida, nombre_base, ui)
            if archivo_a_abrir is None and rutas_frag:
                archivo_a_abrir = rutas_frag[0]
            archivos_generados.extend(os.path.basename(r) for r in rutas_frag)

        ui.log(f"\n=== LISTO ===")

        # --- Armar el resumen para el dialogo final (pregunta si abrir) ---
        nombres_formato = {"completo": "Completo (1 archivo)",
                            "fragmentado": "Fragmentado (para Excel)",
                            "ambos": "Ambos (completo + fragmentado)"}
        lista_archivos = "\n".join(f"  - {n}" for n in archivos_generados)
        resumen = (
            f"Combinacion completada.\n\n"
            f"Tags combinados: {len(series)}\n"
            f"Filas: {len(resultado)}\n"
            f"Columnas (tags): {len(resultado.columns) - 1}\n"
            f"Formato: {nombres_formato.get(formato, formato)}\n\n"
            f"Archivo(s) generado(s):\n{lista_archivos}\n\n"
            f"Carpeta: {carpeta_salida}"
        )

        ui.terminar(exito=True, resumen=resumen, archivo_a_abrir=archivo_a_abrir)

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
