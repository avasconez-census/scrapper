# -*- coding: utf-8 -*-
"""
copiar_fotos.py
---------------
Interfaz gráfica para copiar fotos por matrícula desde un Excel (.xlsx) o PDF.

- Lee un listado de matrículas (Excel columna A o tablas de PDF).
- Busca sus fotos en la carpeta de origen.
- Copia (nunca mueve ni borra) las encontradas a la carpeta destino.
- Genera reportes de éxito/fallo junto al ejecutable.
"""

import os
import re
import sys
import shutil
import threading
import queue
from collections import defaultdict
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# =====================================================================
#  VALORES POR DEFECTO (se pueden cambiar desde la interfaz)
# =====================================================================
ORIGEN_DEFAULT = r"C:\Users\ROXANA\Desktop\CREDENCIALES\FOTOS TOTAL CREDENCIAL"
DESTINO_DEFAULT = "fotos-campeonato"
EXTENSIONES = None   # None = cualquier extensión; o ej. {".jpg", ".png"}

# Carpeta base: junto al .exe o junto al .py
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def ruta(p):
    return p if os.path.isabs(p) else os.path.join(BASE_DIR, p)


# =====================================================================
#  NORMALIZACIÓN
# =====================================================================

def clave(texto):
    """
    Normaliza un texto a su número puro.
    "7.799", "07799", "foto_7799", "7799_1"  →  todos "7799".
    """
    s = re.sub(r"[.,\s'´]", "", str(texto))
    runs = re.findall(r"\d+", s)
    if not runs:
        return None
    mas_largo = max(runs, key=len)
    return mas_largo.lstrip("0") or "0"


def filename_ambiguo(stem):
    """True si el nombre tiene 2+ bloques de dígitos de igual longitud máxima."""
    s = re.sub(r"[.,\s'´]", "", stem)
    runs = re.findall(r"\d+", s)
    if len(runs) < 2:
        return False
    maxlen = max(len(r) for r in runs)
    return sum(1 for r in runs if len(r) == maxlen) > 1


# =====================================================================
#  LECTURA DE LISTADOS
# =====================================================================

def _valores_columna_A_openpyxl(path_excel):
    import openpyxl
    wb = openpyxl.load_workbook(path_excel, data_only=True, read_only=True)
    ws = wb.active
    valores = [fila[0] for fila in ws.iter_rows(min_col=1, max_col=1, values_only=True)]
    wb.close()
    return valores


def _valores_columna_A_stdlib(path_excel):
    """Lector de .xlsx sin dependencias: usa zipfile + xml.etree."""
    import zipfile
    import xml.etree.ElementTree as ET

    def localname(tag):
        return tag.split("}")[-1]

    with zipfile.ZipFile(path_excel) as z:
        compartidas = []
        if "xl/sharedStrings.xml" in z.namelist():
            root = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in root:
                texto = "".join(t.text or "" for t in si.iter() if localname(t.tag) == "t")
                compartidas.append(texto)

        hojas = sorted(n for n in z.namelist()
                       if n.startswith("xl/worksheets/sheet") and n.endswith(".xml"))
        if not hojas:
            return []
        root = ET.fromstring(z.read(hojas[0]))

        por_fila = {}
        for c in root.iter():
            if localname(c.tag) != "c":
                continue
            ref = c.get("r", "")
            col = "".join(ch for ch in ref if ch.isalpha())
            if col != "A":
                continue
            fila_num = int("".join(ch for ch in ref if ch.isdigit()) or "0")
            t = c.get("t")
            valor = None
            if t == "inlineStr":
                valor = "".join(e.text or "" for e in c.iter() if localname(e.tag) == "t")
            else:
                v = next((e for e in c if localname(e.tag) == "v"), None)
                if v is not None and v.text is not None:
                    if t == "s":
                        idx = int(v.text)
                        valor = compartidas[idx] if 0 <= idx < len(compartidas) else None
                    else:
                        valor = v.text
            if valor is not None:
                por_fila[fila_num] = valor

    return [por_fila[k] for k in sorted(por_fila)]


def _leer_excel(path_excel):
    try:
        return _valores_columna_A_openpyxl(path_excel)
    except ImportError:
        return _valores_columna_A_stdlib(path_excel)
    except Exception:
        return _valores_columna_A_stdlib(path_excel)


def _leer_pdf(path_pdf):
    """
    Extrae matrículas de un PDF con múltiples tablas por página.
    Detecta la columna MATRICULA por su encabezado y lee solo esa columna,
    evitando ruido de fechas, nombres y otros campos.
    """
    import pdfplumber
    valores = []
    with pdfplumber.open(path_pdf) as pdf:
        for page in pdf.pages:
            for tabla in (page.extract_tables() or []):
                if not tabla:
                    continue
                # Detecta la columna MATRICULA por el encabezado de la primera fila
                encabezado = [str(c or "").strip().upper() for c in tabla[0]]
                try:
                    col_idx = encabezado.index("MATRICULA")
                except ValueError:
                    # Tabla sin encabezado reconocido: omitir
                    continue
                for fila in tabla[1:]:  # salta la fila de encabezado
                    if fila and col_idx < len(fila) and fila[col_idx]:
                        val = str(fila[col_idx]).strip()
                        if val:
                            valores.append(val)
    return valores


def leer_matriculas(path_archivo):
    """
    Lee un Excel o PDF y devuelve (lista_de_(clave, original), filas, duplicados).
    Detecta el tipo por extensión.
    """
    if not os.path.exists(path_archivo):
        raise FileNotFoundError(f"No encuentro el archivo: {path_archivo}")

    ext = os.path.splitext(path_archivo)[1].lower()
    if ext == ".pdf":
        valores_raw = _leer_pdf(path_archivo)
    else:
        valores_raw = _leer_excel(path_archivo)

    ordenadas = []
    vistas = set()
    filas = 0
    duplicados = 0
    for valor in valores_raw:
        if valor is None or str(valor).strip() == "":
            continue
        filas += 1
        if isinstance(valor, float) and valor.is_integer():
            valor = int(valor)
        k = clave(valor)
        if k is None:
            continue
        if k in vistas:
            duplicados += 1
            continue
        vistas.add(k)
        ordenadas.append((k, valor))
    return ordenadas, filas, duplicados


# =====================================================================
#  LÓGICA DE COPIA
# =====================================================================

def indexar_origen(path_origen):
    if not os.path.isdir(path_origen):
        raise NotADirectoryError(f"No encuentro la carpeta de origen:\n{path_origen}")

    indice = defaultdict(list)
    ambiguos = []
    total = 0
    for nombre in os.listdir(path_origen):
        full = os.path.join(path_origen, nombre)
        if not os.path.isfile(full):
            continue
        stem, ext = os.path.splitext(nombre)
        if EXTENSIONES is not None and ext.lower() not in EXTENSIONES:
            continue
        total += 1
        k = clave(stem)
        if k is None:
            continue
        indice[k].append(nombre)
        if filename_ambiguo(stem):
            ambiguos.append(nombre)
    return indice, total, ambiguos


def ejecutar(path_listado, path_origen, path_destino, log_fn):
    """
    Proceso principal. log_fn(msg) reporta progreso al llamador (hilo de UI).
    Retorna dict con contadores y rutas de reportes.
    """
    matriculas, filas_leidas, dups = leer_matriculas(path_listado)
    log_fn(f"Matrículas leídas: {filas_leidas}  (únicas: {len(matriculas)}, duplicadas: {dups})")

    indice, total_fotos, ambiguos = indexar_origen(path_origen)
    log_fn(f"Fotos en origen: {total_fotos}  (matrículas distintas: {len(indice)})")

    os.makedirs(path_destino, exist_ok=True)

    copiadas = 0
    encontradas = 0
    no_encontradas = []
    detalle_copias = []

    for k, original in matriculas:
        archivos = indice.get(k)
        if not archivos:
            no_encontradas.append(original)
            continue
        encontradas += 1
        for nombre in archivos:
            try:
                shutil.copy2(os.path.join(path_origen, nombre),
                             os.path.join(path_destino, nombre))
                copiadas += 1
                detalle_copias.append((original, nombre))
            except Exception as e:
                log_fn(f"[AVISO] No se pudo copiar {nombre}: {e}")

    # Reportes junto al ejecutable
    rep_no = ruta("no_encontradas.txt")
    with open(rep_no, "w", encoding="utf-8-sig") as f:
        f.write("Matrículas SIN foto encontrada en el origen\n")
        f.write("=" * 44 + "\n")
        for m in no_encontradas:
            f.write(f"{m}\n")

    rep_full = ruta("reporte.txt")
    with open(rep_full, "w", encoding="utf-8-sig") as f:
        f.write("REPORTE DE COPIADO\n")
        f.write("=" * 44 + "\n")
        f.write(f"Matrículas únicas          : {len(matriculas)}\n")
        f.write(f"Con foto encontrada         : {encontradas}\n")
        f.write(f"Sin foto encontrada         : {len(no_encontradas)}\n")
        f.write(f"Archivos copiados           : {len(detalle_copias)}\n\n")
        f.write("-- Copiadas (matrícula -> archivo) --\n")
        for m, nombre in detalle_copias:
            f.write(f"{m} -> {nombre}\n")
        if ambiguos:
            f.write("\n-- Nombres ambiguos (revisar a mano) --\n")
            for nombre in ambiguos:
                f.write(f"{nombre}\n")

    return {
        "total": len(matriculas),
        "encontradas": encontradas,
        "no_encontradas": len(no_encontradas),
        "copiadas": copiadas,
        "rep_no": rep_no,
        "rep_full": rep_full,
        "path_destino": path_destino,
    }


# =====================================================================
#  INTERFAZ GRÁFICA
# =====================================================================

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Copiar Fotos por Matrícula")
        self.resizable(False, False)
        self._q = queue.Queue()
        self._build_ui()

    def _build_ui(self):
        root = self
        pad = {"padx": 12, "pady": 3}

        # -- Listado --
        tk.Label(root, text="Archivo de listado (.xlsx o .pdf):", anchor="w").grid(
            row=0, column=0, columnspan=2, sticky="w", padx=12, pady=(12, 0))
        self.var_listado = tk.StringVar()
        tk.Entry(root, textvariable=self.var_listado, width=52).grid(
            row=1, column=0, padx=(12, 2), pady=2, sticky="ew")
        tk.Button(root, text="Examinar…", command=self._pick_listado).grid(
            row=1, column=1, padx=(2, 12), pady=2)

        # -- Carpeta origen --
        tk.Label(root, text="Carpeta de fotos (origen):", anchor="w").grid(
            row=2, column=0, columnspan=2, sticky="w", padx=12, pady=(8, 0))
        self.var_origen = tk.StringVar(value=ORIGEN_DEFAULT)
        tk.Entry(root, textvariable=self.var_origen, width=52).grid(
            row=3, column=0, padx=(12, 2), pady=2, sticky="ew")
        tk.Button(root, text="Examinar…", command=self._pick_origen).grid(
            row=3, column=1, padx=(2, 12), pady=2)

        # -- Carpeta destino --
        tk.Label(root, text="Carpeta destino:", anchor="w").grid(
            row=4, column=0, columnspan=2, sticky="w", padx=12, pady=(8, 0))
        self.var_destino = tk.StringVar(value=ruta(DESTINO_DEFAULT))
        tk.Entry(root, textvariable=self.var_destino, width=52).grid(
            row=5, column=0, padx=(12, 2), pady=2, sticky="ew")
        tk.Button(root, text="Examinar…", command=self._pick_destino).grid(
            row=5, column=1, padx=(2, 12), pady=2)

        ttk.Separator(root, orient="horizontal").grid(
            row=6, column=0, columnspan=2, sticky="ew", padx=12, pady=10)

        # -- Botón ejecutar --
        self.btn = tk.Button(root, text="Ejecutar", command=self._ejecutar,
                             bg="#2563eb", fg="white",
                             font=("", 11, "bold"), padx=24, pady=7,
                             relief="flat", cursor="hand2")
        self.btn.grid(row=7, column=0, columnspan=2, pady=4)

        # -- Estado --
        self.lbl_estado = tk.Label(root, text="", fg="#6b7280",
                                   wraplength=480, justify="left")
        self.lbl_estado.grid(row=8, column=0, columnspan=2,
                             padx=12, pady=(4, 0), sticky="w")

        # -- Panel de resultados --
        frame = tk.LabelFrame(root, text="Resultados", padx=12, pady=8)
        frame.grid(row=9, column=0, columnspan=2, padx=12, pady=10, sticky="ew")

        self.var_total  = tk.StringVar(value="—")
        self.var_enc    = tk.StringVar(value="—")
        self.var_noenc  = tk.StringVar(value="—")
        self.var_cop    = tk.StringVar(value="—")

        filas = [
            ("Total matrículas leídas:",  self.var_total),
            ("Encontradas con foto:",      self.var_enc),
            ("No encontradas:",            self.var_noenc),
            ("Archivos copiados:",         self.var_cop),
        ]
        for i, (txt, var) in enumerate(filas):
            tk.Label(frame, text=txt, anchor="w", width=26).grid(
                row=i, column=0, sticky="w")
            tk.Label(frame, textvariable=var, anchor="w",
                     font=("", 10, "bold")).grid(row=i, column=1, sticky="w")

        root.columnconfigure(0, weight=1)

    # -- Pickers --
    def _pick_listado(self):
        f = filedialog.askopenfilename(
            title="Seleccionar listado",
            filetypes=[
                ("Excel o PDF", "*.xlsx *.pdf"),
                ("Excel", "*.xlsx"),
                ("PDF", "*.pdf"),
                ("Todos", "*.*"),
            ],
        )
        if f:
            self.var_listado.set(f)

    def _pick_origen(self):
        d = filedialog.askdirectory(title="Carpeta de fotos (origen)")
        if d:
            self.var_origen.set(d)

    def _pick_destino(self):
        d = filedialog.askdirectory(title="Carpeta destino para las fotos copiadas")
        if d:
            self.var_destino.set(d)

    # -- Ejecución en hilo --
    def _log(self, msg):
        self._q.put(("log", msg))

    def _ejecutar(self):
        listado = self.var_listado.get().strip()
        origen  = self.var_origen.get().strip()
        destino = self.var_destino.get().strip()

        if not listado:
            messagebox.showwarning("Falta dato", "Selecciona el archivo de listado.")
            return
        if not origen:
            messagebox.showwarning("Falta dato", "Indica la carpeta de fotos (origen).")
            return
        if not destino:
            messagebox.showwarning("Falta dato", "Indica la carpeta destino.")
            return

        self.btn.config(state="disabled")
        self.lbl_estado.config(text="Procesando…", fg="#6b7280")
        for v in (self.var_total, self.var_enc, self.var_noenc, self.var_cop):
            v.set("…")

        threading.Thread(
            target=self._worker, args=(listado, origen, destino), daemon=True
        ).start()
        self.after(100, self._poll)

    def _worker(self, listado, origen, destino):
        try:
            resultado = ejecutar(listado, origen, destino, self._log)
            self._q.put(("ok", resultado))
        except Exception as e:
            self._q.put(("error", str(e)))

    def _poll(self):
        try:
            while True:
                tipo, data = self._q.get_nowait()
                if tipo == "log":
                    self.lbl_estado.config(text=data, fg="#6b7280")
                elif tipo == "ok":
                    self._mostrar(data)
                    return
                elif tipo == "error":
                    self.lbl_estado.config(text=f"Error: {data}", fg="#dc2626")
                    messagebox.showerror("Error", data)
                    self.btn.config(state="normal")
                    return
        except queue.Empty:
            pass
        self.after(100, self._poll)

    def _mostrar(self, r):
        self.var_total.set(str(r["total"]))
        self.var_enc.set(str(r["encontradas"]))
        self.var_noenc.set(str(r["no_encontradas"]))
        self.var_cop.set(str(r["copiadas"]))
        self.lbl_estado.config(
            text=f"Listo. Reportes guardados en: {BASE_DIR}", fg="#16a34a"
        )
        self.btn.config(state="normal")
        messagebox.showinfo(
            "Completado",
            f"Proceso completado.\n\n"
            f"Total: {r['total']}  |  Encontradas: {r['encontradas']}  |  "
            f"No encontradas: {r['no_encontradas']}\n\n"
            f"Fotos copiadas en:\n  {r['path_destino']}\n\n"
            f"Reportes:\n  {r['rep_no']}\n  {r['rep_full']}",
        )


# =====================================================================
#  ENTRADA
# =====================================================================

if __name__ == "__main__":
    app = App()
    app.mainloop()
