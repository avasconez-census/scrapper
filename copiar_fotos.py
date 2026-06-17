# -*- coding: utf-8 -*-
"""
copiar_fotos.py
---------------
Lee un Excel con un listado de matriculas (numeros enteros, una por fila en la
columna A, sin cabecera) y, por cada matricula, busca su(s) foto(s) dentro de
una carpeta de origen y las COPIA (nunca mueve ni borra) a una carpeta destino.

El emparejamiento es a prueba de formatos: no importa si la foto se llama
"7799", "7.799", "07799", "foto_7799", "7799_1", "7.799 (2)", etc. Todas esas
variantes se reconocen como la matricula 7799.

Solo LEE de la carpeta de origen y COPIA. El origen queda intacto.
"""

import os
import re
import sys
import shutil
from collections import defaultdict

# =====================================================================
#  CONFIGURACION  (edita estas 3 lineas si cambian las rutas)
# =====================================================================

# Carpeta donde estan TODAS las fotos (origen). Se lee, nunca se modifica.
ORIGEN = r"C:\Users\ROXANA\Desktop\CREDENCIALES"

# Excel con las matriculas (columna A, sin cabecera). Por defecto, junto a este script.
EXCEL = "matriculas.xlsx"

# Carpeta donde se copiaran las fotos encontradas. Se crea sola si no existe.
DESTINO = "fotos-campeonato"

# La extension NO importa: si el nombre (numero) coincide, se copia el archivo
# sea cual sea su extension (.jpg, .png, .heic, etc.). Por eso esta en None.
# Si algun dia quisieras limitarlo, pon un set, p.ej. {".jpg", ".png"}.
EXTENSIONES = None

# Si es True, NO copia nada: solo muestra y reporta lo que haria (prueba en seco).
DRY_RUN = False

# =====================================================================
#  A partir de aqui no hace falta tocar nada.
# =====================================================================

# Carpeta donde vive el programa. Las rutas relativas (Excel, destino) se
# resuelven SIEMPRE respecto a ella, corra como .py o empaquetado como .exe.
# Asi, pongas el ejecutable donde lo pongas, buscara el matriculas.xlsx que
# este a su lado y creara ahi mismo la carpeta fotos-campeonato.
if getattr(sys, "frozen", False):          # ejecutable .exe (PyInstaller)
    BASE_DIR = os.path.dirname(sys.executable)
else:                                       # script .py normal
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def ruta(p):
    """Convierte una ruta relativa en absoluta respecto a la carpeta del script."""
    return p if os.path.isabs(p) else os.path.join(BASE_DIR, p)


def clave(texto):
    """
    Normaliza un texto (nombre de archivo o matricula) a su numero puro.

    - Quita separadores de miles: puntos, comas, espacios y apostrofes.
    - Toma la secuencia de digitos mas larga (asi "7799_1" -> "7799").
    - Quita ceros a la izquierda ("07799" -> "7799").

    Devuelve la matricula como cadena de digitos, o None si no hay numeros.
    """
    s = re.sub(r"[.,\s'´]", "", str(texto))
    runs = re.findall(r"\d+", s)
    if not runs:
        return None
    mas_largo = max(runs, key=len)
    return mas_largo.lstrip("0") or "0"


def filename_ambiguo(stem):
    """True si el nombre tiene 2+ bloques de digitos de igual longitud maxima
    (p.ej. '7799-8596'), lo que hace dudoso a que matricula pertenece."""
    s = re.sub(r"[.,\s'´]", "", stem)
    runs = re.findall(r"\d+", s)
    if len(runs) < 2:
        return False
    maxlen = max(len(r) for r in runs)
    return sum(1 for r in runs if len(r) == maxlen) > 1


def _valores_columna_A_openpyxl(path_excel):
    """Lee la columna A con openpyxl (si esta instalado). Devuelve lista de valores."""
    import openpyxl
    wb = openpyxl.load_workbook(path_excel, data_only=True, read_only=True)
    ws = wb.active
    valores = [fila[0] for fila in ws.iter_rows(min_col=1, max_col=1, values_only=True)]
    wb.close()
    return valores


def _valores_columna_A_stdlib(path_excel):
    """Lee la columna A SIN dependencias externas: un .xlsx es un zip con XML.
    Funciona con cualquier Python instalado, sin pip ni internet."""
    import zipfile
    import xml.etree.ElementTree as ET

    def localname(tag):
        return tag.split("}")[-1]  # ignora el namespace

    with zipfile.ZipFile(path_excel) as z:
        # Cadenas compartidas (texto). Los numeros NO usan esto.
        compartidas = []
        if "xl/sharedStrings.xml" in z.namelist():
            root = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in root:
                texto = "".join(t.text or "" for t in si.iter() if localname(t.tag) == "t")
                compartidas.append(texto)

        # Primera hoja de calculo (sheet1, sheet2, ... -> la de numero mas bajo).
        hojas = sorted(n for n in z.namelist()
                       if n.startswith("xl/worksheets/sheet") and n.endswith(".xml"))
        if not hojas:
            return []
        root = ET.fromstring(z.read(hojas[0]))

        por_fila = {}
        for c in root.iter():
            if localname(c.tag) != "c":
                continue
            ref = c.get("r", "")               # p.ej. "A12"
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
                    if t == "s":              # indice a cadena compartida
                        idx = int(v.text)
                        valor = compartidas[idx] if 0 <= idx < len(compartidas) else None
                    else:                      # numero o texto directo
                        valor = v.text
            if valor is not None:
                por_fila[fila_num] = valor

    return [por_fila[k] for k in sorted(por_fila)]


def leer_matriculas(path_excel):
    """Lee la columna A del Excel (sin cabecera) y devuelve una lista de claves
    normalizadas, en orden, sin duplicados. Tambien devuelve cuantas filas se
    leyeron y cuantos duplicados se omitieron."""
    if not os.path.exists(path_excel):
        print(f"\n[ERROR] No encuentro el Excel: {path_excel}")
        print("        Debe llamarse 'matriculas.xlsx' y estar junto al programa.\n")
        sys.exit(1)

    # Intenta openpyxl (mas robusto); si no esta, usa el lector propio sin dependencias.
    try:
        valores = _valores_columna_A_openpyxl(path_excel)
    except ImportError:
        valores = _valores_columna_A_stdlib(path_excel)
    except Exception as e:
        print(f"   [AVISO] openpyxl fallo ({e}); uso el lector interno.")
        valores = _valores_columna_A_stdlib(path_excel)

    ordenadas = []
    vistas = set()
    filas = 0
    duplicados = 0
    for valor in valores:
        if valor is None or str(valor).strip() == "":
            continue
        filas += 1
        # Si Excel lo guardo como float (p.ej. 7799.0), pasarlo a entero limpio.
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


def indexar_origen(path_origen):
    """Recorre la carpeta de origen (sin subcarpetas) y construye un indice
    clave -> [archivos]. Devuelve tambien la lista de nombres ambiguos."""
    if not os.path.isdir(path_origen):
        print(f"\n[ERROR] No encuentro la carpeta de origen: {path_origen}")
        print("        Revisa la variable ORIGEN al inicio del script.\n")
        sys.exit(1)

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


def main():
    path_excel = ruta(EXCEL)
    path_origen = ORIGEN  # normalmente absoluta
    path_destino = ruta(DESTINO)

    print("=" * 64)
    print(" COPIADO SELECTIVO DE FOTOS POR MATRICULA")
    print("=" * 64)
    print(f" Excel    : {path_excel}")
    print(f" Origen   : {path_origen}")
    print(f" Destino  : {path_destino}")
    if DRY_RUN:
        print(" MODO     : PRUEBA EN SECO (no se copia nada)")
    print("-" * 64)

    matriculas, filas_leidas, dups = leer_matriculas(path_excel)
    print(f" Matriculas leidas: {filas_leidas}  (unicas: {len(matriculas)}, duplicadas omitidas: {dups})")

    indice, total_fotos, ambiguos = indexar_origen(path_origen)
    print(f" Fotos en el origen: {total_fotos}  (matriculas distintas detectadas: {len(indice)})")
    print("-" * 64)

    if not DRY_RUN:
        os.makedirs(path_destino, exist_ok=True)

    copiadas = 0
    encontradas = 0
    no_encontradas = []      # lista de matriculas (valor original) sin foto
    detalle_copias = []      # (matricula, archivo_copiado)

    for k, original in matriculas:
        archivos = indice.get(k)
        if not archivos:
            no_encontradas.append(original)
            continue
        encontradas += 1
        for nombre in archivos:
            origen_full = os.path.join(path_origen, nombre)
            destino_full = os.path.join(path_destino, nombre)
            if DRY_RUN:
                detalle_copias.append((original, nombre))
                continue
            try:
                shutil.copy2(origen_full, destino_full)
                copiadas += 1
                detalle_copias.append((original, nombre))
            except Exception as e:
                print(f"   [AVISO] No se pudo copiar {nombre}: {e}")

    # ---------------- Reportes ----------------
    rep_no = ruta("no_encontradas.txt")
    with open(rep_no, "w", encoding="utf-8-sig") as f:
        f.write("Matriculas SIN foto encontrada en el origen\n")
        f.write("=" * 44 + "\n")
        for m in no_encontradas:
            f.write(f"{m}\n")
    rep_full = ruta("reporte.txt")
    with open(rep_full, "w", encoding="utf-8-sig") as f:
        f.write("REPORTE DE COPIADO\n")
        f.write("=" * 44 + "\n")
        f.write(f"Matriculas unicas         : {len(matriculas)}\n")
        f.write(f"Con foto encontrada        : {encontradas}\n")
        f.write(f"Sin foto encontrada        : {len(no_encontradas)}\n")
        f.write(f"Archivos copiados          : {len(detalle_copias)}\n\n")
        f.write("-- Copiadas (matricula -> archivo) --\n")
        for m, nombre in detalle_copias:
            f.write(f"{m} -> {nombre}\n")
        if ambiguos:
            f.write("\n-- Nombres de archivo ambiguos (revisar a mano) --\n")
            for nombre in ambiguos:
                f.write(f"{nombre}\n")

    # ---------------- Resumen en pantalla ----------------
    print("-" * 64)
    accion = "Se copiarian" if DRY_RUN else "Copiadas"
    print(f" Matriculas con foto : {encontradas} de {len(matriculas)}")
    print(f" {accion} (archivos): {len(detalle_copias)}")
    print(f" Sin foto            : {len(no_encontradas)}  -> ver no_encontradas.txt")
    if ambiguos:
        print(f" Nombres ambiguos    : {len(ambiguos)}  -> ver reporte.txt (revisar a mano)")
    print("-" * 64)
    print(f" Reportes generados:\n   {rep_no}\n   {rep_full}")
    if not DRY_RUN:
        print(f" Fotos en: {path_destino}")
    print("=" * 64)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:  # noqa: BLE001
        print(f"\n[ERROR inesperado] {e}")
    finally:
        # Cuando se ejecuta como .exe (doble clic), la ventana se cerraria sola.
        # Esperamos un Enter para que se pueda leer el resumen o el error.
        if getattr(sys, "frozen", False):
            try:
                input("\nListo. Presiona Enter para cerrar esta ventana...")
            except EOFError:
                pass
