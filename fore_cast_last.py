import os, sys, shutil, time, tempfile, threading, queue, ctypes
from pathlib import Path
from datetime import datetime, date
import tkinter as tk
from tkinter import filedialog
from tkinter.scrolledtext import ScrolledText

try:
    import win32com.client as win32
except Exception:
    win32 = None

import pythoncom  # COM en hilos

# ============== CONFIG ==============
CONFIG = {
    # Ya no usamos SUPPLYON_INPUT fijo: el usuario lo selecciona con un diálogo
    "FORECAST_FILE": r"C:\Forecast\Forecast Vitesco, Bosch and Conti Mechelen.xlsx",

    "SHEET_LAST_LAR": "Last LAR",
    "SHEET_OLD_LAR": "Old LAR CW-1",

    "SUPPLYON_DATE_COLS": ["AK", "BT"],
    "SUPPLYON_DATE_NAME_CANDIDATES": ["Fecha de entrega", "Fecha de entrega.1", "Delivery Date"],

    # Si True y el input es CSV, abrimos el **CSV original** con Excel (modo que funcionaba).
    # Si False, abrimos **copia temporal** del CSV.
    "OPEN_ORIGINAL_CSV": True,

    # Carpeta inicial del diálogo (ajústala si quieres)
    "INITIAL_DIR": r"C:\Forecast",
}
# ====================================

# ====== Cola de logs para tkinter ======
log_queue = queue.Queue()

def log(msg):
    line = f"[{datetime.now():%H:%M:%S}] {msg}"
    print(line)
    log_queue.put(line)

# ---------- Tkinter ventana de logs + controles ----------
class LogWindow:
    def __init__(self, root):
        self.root = root
        self.root.title("Proceso LAR - Logs")
        self.root.geometry("820x520+380+180")

        # ---- Controles superiores (selección de archivo e inicio) ----
        top = tk.Frame(root)
        top.pack(fill="x", padx=10, pady=10)

        self.btn_select = tk.Button(top, text="Seleccionar Delfor…", command=self.select_file)
        self.btn_select.pack(side="left")

        self.selected_path_var = tk.StringVar(value="(ningún archivo seleccionado)")
        self.path_label = tk.Label(top, textvariable=self.selected_path_var, anchor="w", justify="left")
        self.path_label.pack(side="left", padx=10, expand=True, fill="x")

        self.btn_run = tk.Button(top, text="Iniciar proceso", state="disabled", command=self.start_process)
        self.btn_run.pack(side="right")

        # ---- Log text ----
        self.text = ScrolledText(root, wrap="word", state="disabled")
        self.text.pack(expand=True, fill="both", padx=10, pady=(0,10))

        self.selected_file = None
        self.update_logs()

    def update_logs(self):
        while not log_queue.empty():
            line = log_queue.get_nowait()
            self.text.configure(state="normal")
            self.text.insert("end", line + "\n")
            self.text.see("end")
            self.text.configure(state="disabled")
        self.root.after(150, self.update_logs)

    def select_file(self):
        initialdir = CONFIG.get("INITIAL_DIR") or os.getcwd()
        filetypes = [
            ("Archivos Delfor", "*.csv *.xlsx"),
            ("CSV", "*.csv"),
            ("Excel XLSX", "*.xlsx"),
            ("Todos", "*.*"),
        ]
        path = filedialog.askopenfilename(
            title="Selecciona el Delfor (CSV o XLSX)",
            initialdir=initialdir,
            filetypes=filetypes
        )
        if not path:
            return
        self.selected_file = path
        self.selected_path_var.set(path)
        self.btn_run.config(state="normal")

    def start_process(self):
        if not self.selected_file:
            return
        # Deshabilitar botones mientras corre
        self.btn_run.config(state="disabled")
        self.btn_select.config(state="disabled")
        # Lanzar hilo de trabajo
        threading.Thread(target=lar_process, args=(self.selected_file, self.on_process_done), daemon=True).start()

    def on_process_done(self):
        # Rehabilitar selección e iniciar de nuevo si se desea
        self.btn_select.config(state="normal")
        self.btn_run.config(state="normal")

# ---------- utilidades ----------
def today_strings():
    t = date.today()
    # isocalendar() -> (ISO_year, ISO_week_number, ISO_weekday)
    return t.strftime("%Y%m%d"), f"CW{t.isocalendar()[1]:02d}", t.isocalendar()[0], t

def excel_doevents_safe(excel):
    try: excel.DoEvents(); return
    except: pass
    try: excel.Application.DoEvents(); return
    except: pass
    time.sleep(0.2)

# ---------- helpers para borrado robusto de CSV ----------
MOVEFILE_DELAY_UNTIL_REBOOT = 0x4

def _schedule_delete_on_reboot(path_str: str):
    """Marca el archivo para borrado en el próximo reinicio."""
    try:
        ctypes.windll.kernel32.MoveFileExW(path_str, None, MOVEFILE_DELAY_UNTIL_REBOOT)
        log(f"   ⚠️ Archivo programado para borrado al reiniciar: {path_str}")
    except Exception as e:
        log(f"   ⚠️ No se pudo programar borrado al reiniciar: {e}")

def _close_if_open_in_excel(file_path: str) -> bool:
    """Si el archivo está abierto en alguna instancia de Excel, intenta cerrarlo."""
    if win32 is None:
        return False
    try:
        app = win32.GetObject(None, "Excel.Application")  # instancia activa si existe
    except Exception:
        return False
    closed_any = False
    try:
        abs_target = os.path.abspath(file_path).lower()
        for wb in list(app.Workbooks):
            try:
                if os.path.abspath(wb.FullName).lower() == abs_target:
                    wb.Close(SaveChanges=0)
                    log(f"   Cerrado desde Excel (COM): {file_path}")
                    closed_any = True
            except Exception:
                pass
    except Exception:
        pass
    return closed_any

def _try_delete_or_schedule(p: Path, retries: int = 20, sleep_s: float = 0.5):
    """
    Borrar con reintentos “largos”.
    1) Intentar cerrar si lo tiene Excel.
    2) Reintentar varias veces.
    3) Si sigue, programar borrado al reiniciar.
    """
    for i in range(retries):
        try:
            p.unlink()
            log(f"   CSV original eliminado: {p.name}")
            return
        except PermissionError:
            if _close_if_open_in_excel(str(p)):
                time.sleep(0.3)
                continue
            time.sleep(sleep_s)
        except Exception as e:
            if i == retries - 1:
                log(f"   ⚠️ No se pudo borrar {p.name}: {e}")
            time.sleep(sleep_s)
    _schedule_delete_on_reboot(str(p))

# ---------- Normalizar fechas del Delfor (csv o xlsx) ----------
def normalize_supplyon_dates_with_excel(path_in, date_cols_letters, name_candidates=None):
    """
    Abre el archivo con Excel (COM) y reemplaza '.' por '/' en columnas indicadas.
    - Si input es CSV:
        * Si OPEN_ORIGINAL_CSV=True: abrir el **original** (modo probado).
        * Si OPEN_ORIGINAL_CSV=False: abrir **copia temporal** para no tocar el original.
      En ambos casos, guardar como YYYYMMDD.xlsx y borrar/planificar borrado del CSV.
    - Si input es XLSX: guardar como YYYYMMDD.xlsx (sin sufijos).
    Devuelve la nueva ruta .xlsx
    """
    if win32 is None:
        raise RuntimeError("pywin32 no disponible. Instala con: pip install pywin32")

    p = Path(path_in)
    if not p.exists():
        raise FileNotFoundError(f"No se encuentra el archivo SupplyOn: {path_in}")

    ymd, _, _, _ = today_strings()
    new_path = p.with_name(f"{ymd}.xlsx")

    log("➡️ Paso 1/5: Normalizando fechas del Delfor (COM)...")

    is_csv = (p.suffix.lower() == ".csv")
    open_original = CONFIG.get("OPEN_ORIGINAL_CSV", True) and is_csv

    # Si no vamos a abrir el original, crear copia temporal
    if is_csv and not open_original:
        work_dir = Path(tempfile.mkdtemp(prefix="lar_csv_"))
        path_to_open = work_dir / p.name
        shutil.copy2(str(p), str(path_to_open))
    else:
        work_dir = None
        path_to_open = p

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    wb = ws = None
    xlUp = -4162
    xlToLeft = -4159
    xlPart = 2
    xlByRows = 1
    try:
        # Abrimos el **path_to_open** (original o copia)
        wb = excel.Workbooks.Open(str(path_to_open))
        ws = wb.Worksheets(1)

        target_cols = set((date_cols_letters or []))

        # Detectar cabeceras por nombre
        headers = {}
        last_col = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        for c in range(1, last_col + 1):
            v = ws.Cells(1, c).Value
            headers[str(v) if v is not None else ""] = c

        for cand in (name_candidates or []):
            if cand in headers:
                col_idx = headers[cand]
                i = col_idx; col_letter = ""
                while i > 0:
                    i, r = divmod(i - 1, 26)
                    col_letter = chr(r + 65) + col_letter
                target_cols.add(col_letter)

        # Reemplazo en columnas objetivo
        for col in sorted(target_cols):
            last_row = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
            if last_row >= 1:
                ws.Range(f"{col}1:{col}{last_row}").Replace(
                    What=".", Replacement="/", LookAt=xlPart, SearchOrder=xlByRows, MatchCase=False
                )
                log(f"   Normalizado '.'→'/' en columna {col} (1:{last_row})")

        # Guardar SIEMPRE como XLSX con nombre YYYYMMDD.xlsx, evitando sufijos
        if Path(new_path).exists():
            try:
                Path(new_path).unlink()
            except Exception as e:
                log(f"   ⚠️ No se pudo eliminar {new_path}: {e}")
        wb.SaveAs(str(new_path), FileFormat=51)  # xlsx
        log(f"   Delfor guardado como {new_path}")

    finally:
        # Cerrar Excel y liberar COM
        try:
            if wb: wb.Close(SaveChanges=0)
        except Exception as e:
            log(f"   ⚠️ Cierre de libro: {e}")
        try:
            excel.Quit()
        except Exception as e:
            log(f"   ⚠️ Excel.Quit(): {e}")
        wb = None; ws = None; excel = None

        # Si trabajamos con copia temporal, limpiarla
        if is_csv and not open_original and 'work_dir' in locals() and work_dir:
            try:
                # borrar archivo de trabajo
                try:
                    (work_dir / p.name).unlink(missing_ok=True)
                except Exception:
                    pass
                # borrar carpeta temporal si vacía
                try:
                    work_dir.rmdir()
                except OSError:
                    pass
            except Exception as e:
                log(f"   ⚠️ Limpieza temporal falló: {e}")

        # Borrar CSV original si procede (con reintentos largos + cierre vía COM + borrar al reiniciar)
        if is_csv and p.exists():
            _try_delete_or_schedule(p)

    return str(new_path)

# ---------- Actualización con COM (Forecast) ----------
def update_forecast_with_com(forecast_path, supplyon_path, sheet_last, sheet_old):
    log("➡️ Paso 2/5: Actualizando Forecast con COM...")
    if win32 is None:
        raise RuntimeError("pywin32 no disponible. Instala con: pip install pywin32")

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False

    wb = tmpwb = ws_last = tmpws = None
    xlUp = -4162

    try:
        wb = excel.Workbooks.Open(os.path.abspath(forecast_path))
        ws_last = wb.Worksheets(sheet_last)

        last_lr = ws_last.Cells(ws_last.Rows.Count, "B").End(xlUp).Row
        log(f"   Última fila previa en Last LAR: {last_lr}")

        # LAST LAR
        if last_lr >= 2:
            log(f"   Limpiando Last LAR datos rango B2:BT{last_lr} (se conservan fórmulas en BU:CA)")
            ws_last.Range(f"B2:BT{last_lr}").Clear()

        tmpwb = excel.Workbooks.Open(os.path.abspath(supplyon_path))
        tmpws = tmpwb.Worksheets(1)
        delfor_lr = tmpws.Cells(tmpws.Rows.Count, "A").End(xlUp).Row
        log(f"   Última fila en Delfor: {delfor_lr}")
        if delfor_lr >= 2:
            log(f"   Pegando rango Delfor A2:BT{delfor_lr} → Last LAR B2")
            tmpws.Range(f"A2:BT{delfor_lr}").Copy()
            dest_last = ws_last.Range("B2")
            dest_last.PasteSpecial(-4163)  # valores
            dest_last.PasteSpecial(-4122)  # formatos
            excel_doevents_safe(excel)
            excel.CutCopyMode = False

        new_last = ws_last.Cells(ws_last.Rows.Count, "B").End(xlUp).Row
        log(f"   Última fila en Last LAR tras pegar Delfor: {new_last}")

        if new_last > last_lr:
            log(f"   Delfor más largo → arrastrando fórmulas A y BU:CA")
            ws_last.Range("A2").AutoFill(Destination=ws_last.Range(f"A2:A{new_last}"))
            ws_last.Range("BU2:CA2").AutoFill(Destination=ws_last.Range(f"BU2:CA{new_last}"))
        elif new_last < last_lr:
            log(f"   Delfor más corto → limpiando sobrante B{new_last+1}:CA{last_lr}")
            ws_last.Range(f"B{new_last+1}:CA{last_lr}").Clear()

        wb.Save()
        log("   Forecast actualizado y guardado.")
    finally:
        try:
            if tmpwb: tmpwb.Close(SaveChanges=0)
            if wb: wb.Close(SaveChanges=1)
        except Exception as e:
            log(f"   ⚠️ Cierre de libros: {e}")
        try:
            excel.ScreenUpdating = True
            excel.Quit()
        except Exception as e:
            log(f"   ⚠️ Excel.Quit(): {e}")

# ---------- Proceso LAR (hilo de trabajo) ----------
def lar_process(supplyon_input_path: str, on_done_callback=None):
    pythoncom.CoInitialize()  # COM en este hilo
    cfg = CONFIG
    try:
        # 1) Normalizar y convertir/guardar como YYYYMMDD.xlsx; borrar CSV si aplica
        supplyon_path = normalize_supplyon_dates_with_excel(
            supplyon_input_path,
            cfg.get("SUPPLYON_DATE_COLS"),
            cfg.get("SUPPLYON_DATE_NAME_CANDIDATES"),
        )

        # 2) Copia temporal del Forecast
        src = cfg["FORECAST_FILE"]
        tmp_dir = tempfile.mkdtemp(prefix="lar_")
        tmp_path = os.path.join(tmp_dir, "forecast_tmp.xlsx")
        shutil.copy2(src, tmp_path)
        log(f"➡️ Copia temporal: {tmp_path}")

        # 3) Actualizar Forecast
        update_forecast_with_com(tmp_path, supplyon_path, cfg["SHEET_LAST_LAR"], None)

        # 4) Sincronizar
        shutil.copy2(tmp_path, src)
        log(f"➡️ Sincronizado al archivo original: {src}")

        log("✅ Proceso LAR completado.")
    except Exception as e:
        log(f"❌ Error en proceso: {e}")
    finally:
        pythoncom.CoUninitialize()  # limpieza COM
        if on_done_callback:
            try:
                on_done_callback()
            except Exception:
                pass

# ---------- Arranque principal (tkinter en hilo principal) ----------
def main():
    root = tk.Tk()
    LogWindow(root)
    root.mainloop()

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        log(f"❌ Error: {e}")
        sys.exit(1)
