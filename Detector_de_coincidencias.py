"""
=============================================================================
 CONCILIACION DE VENTAS Y COBROS - MANUFACTURAS COUSINO LTDA.
 Cruza el Libro de Ventas (Sofnet/ERP) contra los cobros de las pasarelas
 de pago Transbank y Mercado Libre / Mercado Pago.
=============================================================================

FLUJO
  1. Busca los Excel de entrada en las subcarpetas de Sofnet, Transbank y
     Mercado Libre (recursivo).
  2. Detecta automaticamente en que fila empieza cada tabla (los exportes
     traen titulos/logos/parametros arriba) y normaliza cada fuente a las
     columnas estandar: FOLIO, FECHA, TOTAL, ORIGEN (+ ID de referencia).
  3. Corre el motor de conciliacion biyectiva 1 a 1 (reglas R1 -> R4).
  4. Agrupa el resultado por mes y escribe un Excel
     "Conciliacion_AAAA_MM.xlsx" por periodo en la carpeta Output.
  5. Agrega la pestaña "Informe CPE" (dentro del mismo Excel mensual) con el formato del formulario SII de
     Comprobantes de Pago Electronico (un registro por dia y canal) con los
     montos corregidos: solo comprobantes sin boleta/factura asociada.

USO
    python conciliacion_ventas.py
  (o bien ajustando --base si la carpeta raiz cambia)

REQUISITOS
    pip install pandas openpyxl
"""

from __future__ import annotations

import argparse
import calendar
import glob
import os
import re
import sys
from dataclasses import dataclass, field
from datetime import datetime, date
from typing import Optional

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# 1. CONFIGURACION DE RUTAS
# ---------------------------------------------------------------------------

DEFAULT_BASE_DIR = r"C:\Users\Usuario\OneDrive\Escritorio\Inputs\Ventas"

# Subcarpetas donde se buscan los archivos de entrada (relativas a BASE_DIR)
SOFNET_SUBDIR = os.path.join("Inputs", "Sofnet", "Libro de ventas", "Libro de ventas 2026")
TRANSBANK_SUBDIR = os.path.join("Inputs", "Transbank", "Ventas")
MELI_SUBDIR = os.path.join("Inputs", "Mercado libre")
OUTPUT_SUBDIR = os.path.join("Output", "Detector_de_coincidencias")

TOLERANCIA = 5  # $ de tolerancia para reglas R2 y R4

MESES_ES = {
    "enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6,
    "julio": 7, "agosto": 8, "septiembre": 9, "setiembre": 9, "octubre": 10,
    "noviembre": 11, "diciembre": 12,
}

# ---------------------------------------------------------------------------
# 2. UTILIDADES DE LIMPIEZA (montos y fechas)
# ---------------------------------------------------------------------------


def limpiar_monto(valor) -> Optional[float]:
    """Convierte '$1.234.567', '1234567,50', 1234567.0, etc. a float."""
    if valor is None:
        return None
    if isinstance(valor, (int, float)):
        if pd.isna(valor):
            return None
        return float(valor)

    texto = str(valor).strip()
    if texto == "" or texto == "-":
        return None

    texto = texto.replace("$", "").replace(" ", "")
    # Si tiene coma Y punto, asumimos punto=miles y coma=decimal (formato CL)
    if "," in texto and "." in texto:
        texto = texto.replace(".", "").replace(",", ".")
    elif "," in texto:
        # solo coma -> podria ser decimal o miles; si hay 3 digitos tras la
        # ultima coma la tratamos como miles, si no como decimal
        partes = texto.split(",")
        if len(partes[-1]) == 3 and len(partes) > 1:
            texto = texto.replace(",", "")
        else:
            texto = texto.replace(",", ".")
    else:
        # solo puntos: si el ultimo grupo tiene 3 digitos es separador de
        # miles (formato CLP habitual), si no es decimal
        partes = texto.split(".")
        if len(partes) > 1 and len(partes[-1]) == 3:
            texto = texto.replace(".", "")

    try:
        return float(texto)
    except ValueError:
        return None


def parsear_fecha(valor) -> Optional[date]:
    """Convierte fechas en distintos formatos (ISO, dd/mm/aaaa, texto en
    español tipo '30 abril 2026 12:00 AM', datetime de Excel) a date."""
    if valor is None:
        return None
    if isinstance(valor, datetime):
        return valor.date()
    if isinstance(valor, date):
        return valor

    texto = str(valor).strip()
    if texto == "" or texto == "-":
        return None

    # Formato con mes en español: "30 abril 2026 12:00 AM"
    m = re.match(r"^(\d{1,2})\s+([A-Za-zñÑ]+)\s+(\d{4})", texto)
    if m:
        dia, mes_txt, anio = m.groups()
        mes = MESES_ES.get(mes_txt.lower())
        if mes:
            return date(int(anio), mes, int(dia))

    # Formatos numericos comunes
    for fmt in (
        "%Y-%m-%d",
        "%Y-%m-%d %H:%M:%S",
        "%d/%m/%Y",
        "%d/%m/%Y %H:%M:%S",
        "%d-%m-%Y",
        "%m/%d/%Y",
    ):
        try:
            return datetime.strptime(texto, fmt).date()
        except ValueError:
            continue

    # Ultimo intento con pandas (dayfirst = formato chileno)
    try:
        ts = pd.to_datetime(texto, dayfirst=True, errors="raise")
        return ts.date()
    except Exception:
        return None


# ---------------------------------------------------------------------------
# 3. DETECCION DE ENCABEZADOS Y CARGA POR FUENTE
# ---------------------------------------------------------------------------


def _encontrar_fila_encabezado(path: str, columnas_clave: list[str], max_filas: int = 30) -> int:
    """Inspecciona las primeras filas del Excel y devuelve el indice (0-based)
    de la fila que contiene todas las columnas_clave (busqueda flexible por
    coincidencia parcial, insensible a mayusculas)."""
    preview = pd.read_excel(path, header=None, nrows=max_filas, dtype=str)
    claves_low = [c.lower() for c in columnas_clave]

    for i in range(len(preview)):
        fila_valores = [str(v).strip().lower() for v in preview.iloc[i].tolist() if pd.notna(v)]
        if not fila_valores:
            continue
        encontrados = 0
        for clave in claves_low:
            if any(clave in val for val in fila_valores):
                encontrados += 1
        if encontrados == len(claves_low):
            return i
    raise ValueError(
        f"No se encontro la fila de encabezado en '{os.path.basename(path)}' "
        f"buscando columnas {columnas_clave}"
    )


def cargar_sofnet(path: str) -> pd.DataFrame:
    fila_header = _encontrar_fila_encabezado(path, ["fecha", "cto", "total"])
    df = pd.read_excel(path, header=fila_header)
    df.columns = [str(c).strip() for c in df.columns]

    out = pd.DataFrame()
    out["FOLIO"] = df["Cto"].astype(str) + "-" + df.get("Tipo Doc", "").astype(str)
    out["FECHA"] = df["Fecha"].apply(parsear_fecha)
    out["TOTAL"] = df["Total"].apply(limpiar_monto)
    out["ORIGEN"] = "SOFNET"
    out["ARCHIVO"] = os.path.basename(path)
    out = out.dropna(subset=["FECHA", "TOTAL"])
    return out


def cargar_transbank(path: str) -> pd.DataFrame:
    fila_header = _encontrar_fila_encabezado(
        path, ["fecha de movimiento", "monto original de la venta"]
    )
    df = pd.read_excel(path, header=fila_header)
    df.columns = [str(c).strip() for c in df.columns]

    # Solo nos interesan las ventas efectivamente abonadas (se descartan
    # anulaciones, que no son cobros)
    if "Tipo de movimiento" in df.columns:
        df = df[df["Tipo de movimiento"].astype(str).str.strip().str.lower() == "venta"]

    col_fecha = "Fecha de movimiento"

    out = pd.DataFrame()
    out["FOLIO"] = df.get("Código de autorización de la venta", "").astype(str)
    out["FECHA"] = df[col_fecha].apply(parsear_fecha)
    out["TOTAL"] = df["Monto original de la venta"].apply(limpiar_monto)
    out["ORIGEN"] = "TRANSBANK"
    # Canal segun "N° de boleta" (col AB): 0000000000 / vacio / "-" -> Internet;
    # cualquier numero mayor que 0 (ej. 0000000617) -> Presencial
    col_boleta = next((c for c in df.columns if "de boleta" in c.lower()), None)
    out["CANAL"] = df[col_boleta].map(canal_transbank) if col_boleta else "Internet"
    out["ARCHIVO"] = os.path.basename(path)
    out = out.dropna(subset=["FECHA", "TOTAL"])
    return out


def cargar_meli(path: str) -> pd.DataFrame:
    fila_header = _encontrar_fila_encabezado(
        path, ["fecha de acreditación", "transaction_amount"]
    )
    df = pd.read_excel(path, header=fila_header)
    df.columns = [str(c).strip() for c in df.columns]

    col_monto = next((c for c in df.columns if "transaction_amount" in c.lower()), None)
    col_fecha = next((c for c in df.columns if "date_approved" in c.lower()), None)
    if col_fecha is None:
        col_fecha = next((c for c in df.columns if "date_created" in c.lower()), None)
    col_id = next((c for c in df.columns if "order_id" in c.lower()), None)
    if col_id is None:
        col_id = next((c for c in df.columns if "operation_id" in c.lower()), None)

    out = pd.DataFrame()
    out["FOLIO"] = df[col_id].astype(str) if col_id else ""
    out["FECHA"] = df[col_fecha].apply(parsear_fecha)
    out["TOTAL"] = df[col_monto].apply(limpiar_monto)
    out["ORIGEN"] = "MERCADO LIBRE"
    out["CANAL"] = "Internet"
    out["ARCHIVO"] = os.path.basename(path)
    out = out.dropna(subset=["FECHA", "TOTAL"])
    return out


def canal_transbank(n_boleta) -> str:
    """N° de boleta Transbank -> canal. Compara el valor numerico para que
    funcione igual si Excel lo guarda como texto '0000000617' o como 617."""
    digitos = re.sub(r"\D", "", str(n_boleta)) if pd.notna(n_boleta) else ""
    if isinstance(n_boleta, float) and not pd.isna(n_boleta):
        digitos = str(int(n_boleta))
    return "Presencial" if digitos and int(digitos) > 0 else "Internet"


def _buscar_archivos(carpeta: str) -> list[str]:
    if not os.path.isdir(carpeta):
        return []
    patrones = [os.path.join(carpeta, "**", "*.xlsx"), os.path.join(carpeta, "**", "*.xls")]
    archivos = []
    for patron in patrones:
        archivos.extend(glob.glob(patron, recursive=True))
    # descarta temporales de Excel (~$archivo.xlsx)
    return sorted(a for a in archivos if not os.path.basename(a).startswith("~$"))


def cargar_todas_las_fuentes(base_dir: str) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Devuelve (sofnet_df, cobros_df) ya normalizados."""
    sofnet_dir = os.path.join(base_dir, SOFNET_SUBDIR)
    transbank_dir = os.path.join(base_dir, TRANSBANK_SUBDIR)
    meli_dir = os.path.join(base_dir, MELI_SUBDIR)

    sofnet_frames, cobros_frames = [], []

    for path in _buscar_archivos(sofnet_dir):
        try:
            sofnet_frames.append(cargar_sofnet(path))
            print(f"  [Sofnet]        OK  {os.path.basename(path)}")
        except Exception as e:
            print(f"  [Sofnet]        ERROR en {os.path.basename(path)}: {e}")

    for path in _buscar_archivos(transbank_dir):
        try:
            cobros_frames.append(cargar_transbank(path))
            print(f"  [Transbank]     OK  {os.path.basename(path)}")
        except Exception as e:
            print(f"  [Transbank]     ERROR en {os.path.basename(path)}: {e}")

    for path in _buscar_archivos(meli_dir):
        try:
            cobros_frames.append(cargar_meli(path))
            print(f"  [Mercado Libre] OK  {os.path.basename(path)}")
        except Exception as e:
            print(f"  [Mercado Libre] ERROR en {os.path.basename(path)}: {e}")

    sofnet_df = (
        pd.concat(sofnet_frames, ignore_index=True) if sofnet_frames else pd.DataFrame(
            columns=["FOLIO", "FECHA", "TOTAL", "ORIGEN", "ARCHIVO"]
        )
    )
    cobros_df = (
        pd.concat(cobros_frames, ignore_index=True) if cobros_frames else pd.DataFrame(
            columns=["FOLIO", "FECHA", "TOTAL", "ORIGEN", "CANAL", "ARCHIVO"]
        )
    )
    return sofnet_df, cobros_df


# ---------------------------------------------------------------------------
# 4. MOTOR DE CONCILIACION BIYECTIVA (R1 -> R4)
# ---------------------------------------------------------------------------


@dataclass
class VentaSofnet:
    idx: int
    folio: str
    fecha: date
    total: float
    usada: bool = False


def conciliar(sofnet_df: pd.DataFrame, cobros_df: pd.DataFrame) -> pd.DataFrame:
    ventas = [
        VentaSofnet(idx=i, folio=r.FOLIO, fecha=r.FECHA, total=r.TOTAL)
        for i, r in sofnet_df.reset_index(drop=True).iterrows()
    ]

    resultados = []
    pendientes = list(cobros_df.reset_index(drop=True).itertuples())

    def disponibles():
        return [v for v in ventas if not v.usada]

    def buscar_y_marcar(cobro, condicion) -> Optional[VentaSofnet]:
        for v in disponibles():
            if condicion(v):
                v.usada = True
                return v
        return None

    reglas = [
        ("R1", lambda c, v: v.fecha == c.FECHA and v.total == c.TOTAL),
        ("R2", lambda c, v: v.fecha == c.FECHA and abs(v.total - c.TOTAL) <= TOLERANCIA),
        ("R3", lambda c, v: (v.fecha.year, v.fecha.month) == (c.FECHA.year, c.FECHA.month) and v.total == c.TOTAL),
        ("R4", lambda c, v: (v.fecha.year, v.fecha.month) == (c.FECHA.year, c.FECHA.month) and abs(v.total - c.TOTAL) <= TOLERANCIA),
    ]

    aun_pendientes = pendientes
    matches = {}  # index en pendientes -> (regla, VentaSofnet)

    for nombre_regla, condicion in reglas:
        siguiente_pendientes = []
        for cobro in aun_pendientes:
            match = buscar_y_marcar(cobro, lambda v: condicion(cobro, v))
            if match:
                matches[cobro.Index] = (nombre_regla, match)
            else:
                siguiente_pendientes.append(cobro)
        aun_pendientes = siguiente_pendientes

    for cobro in pendientes:
        regla, venta = matches.get(cobro.Index, (None, None))

        if venta and venta.folio:
            partes_folio = venta.folio.rsplit("-", 1)
            if len(partes_folio) == 2:
                folio_sofnet, tipo_doc_sofnet = partes_folio
            else:
                folio_sofnet, tipo_doc_sofnet = venta.folio, None
        else:
            folio_sofnet, tipo_doc_sofnet = None, None

        monto_neto = round(cobro.TOTAL / 1.19, 2)
        monto_iva = round(cobro.TOTAL - monto_neto, 2)

        resultados.append(
            {
                "ORIGEN_COBRO": cobro.ORIGEN,
                "CANAL": getattr(cobro, "CANAL", None) or "Internet",
                "ID_COBRO": cobro.FOLIO,
                "FECHA_COBRO": cobro.FECHA,
                "TOTAL_COBRO": cobro.TOTAL,
                "MONTO_NETO": monto_neto,
                "MONTO_IVA": monto_iva,
                "FOLIO_SOFNET": folio_sofnet,
                "TIPO_DOC_SOFNET": tipo_doc_sofnet,
                "FECHA_SOFNET": venta.fecha if venta else None,
                "TOTAL_SOFNET": venta.total if venta else None,
                "DIFERENCIA_MONTO": (cobro.TOTAL - venta.total) if venta else None,
                "DIFERENCIA_DIAS": (cobro.FECHA - venta.fecha).days if venta else None,
                "REGLA": regla if regla else "NO_CONCILIADO",
            }
        )

    resultado_df = pd.DataFrame(resultados)

    # Ventas de Sofnet que quedaron sin ningun cobro asociado (informativo)
    sofnet_sin_cruce = sofnet_df.reset_index(drop=True).loc[
        [v.idx for v in ventas if not v.usada]
    ].copy()

    return resultado_df, sofnet_sin_cruce


# ---------------------------------------------------------------------------
# 5. EXPORTACION A EXCEL, AGRUPADA POR MES
# ---------------------------------------------------------------------------

HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
HEADER_FONT = Font(name="Arial", bold=True, color="FFFFFF")
CELL_FONT = Font(name="Arial")
NO_CONCILIADO_FILL = PatternFill(start_color="FCE4E4", end_color="FCE4E4", fill_type="solid")


def _escribir_hoja(ws, df: pd.DataFrame):
    if df.empty:
        ws.append(["Sin registros"])
        return

    ws.append(list(df.columns))
    for cell in ws[1]:
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center")

    for _, row in df.iterrows():
        ws.append(list(row))

    # resaltar filas NO_CONCILIADO si existe esa columna
    if "REGLA" in df.columns:
        col_regla = list(df.columns).index("REGLA") + 1
        for r in range(2, ws.max_row + 1):
            if ws.cell(row=r, column=col_regla).value == "NO_CONCILIADO":
                for c in range(1, ws.max_column + 1):
                    ws.cell(row=r, column=c).fill = NO_CONCILIADO_FILL

    for r in range(2, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(row=r, column=c)
            if not cell.font.bold:
                cell.font = CELL_FONT

    for c, col_name in enumerate(df.columns, start=1):
        largo_datos = df[col_name].map(lambda x: len(str(x))).max() if not df.empty else 10
        largo = max(len(str(col_name)), int(largo_datos))
        ws.column_dimensions[get_column_letter(c)].width = min(largo + 3, 40)


# ---------------------------------------------------------------------------
# 5b. HOJA CPE (portado de escribirHojaCPE de Apps Script)
# ---------------------------------------------------------------------------

CPE_HEADERS = [
    "Venta del Día", "Canal", "Operador", "Tot. Doc.",
    "Monto Total", "Monto Neto", "Monto Iva", "Monto Exento", "Documentos",
    "folios", "bruto", "monto total", "neto", "iva",
]
CPE_HEADER_FILL = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
CPE_HEADER_FONT = Font(bold=True, color="FFFFFF")
# A-I: mismos anchos que en Apps Script; J-N: columnas de correccion
CPE_ANCHOS_PX = [110, 100, 120, 75, 120, 120, 110, 120, 400, 190, 110, 110, 110, 110]
CPE_OPERADOR = {"TRANSBANK": "TRANSBANK", "MERCADO LIBRE": "MERCADOPAGO"}


def _fmt_clp(monto) -> str:
    """1234567 -> '$1.234.567' (equivale a toLocaleString('es-CL'))."""
    return "$" + f"{monto:,.0f}".replace(",", ".")


def construir_filas_cpe(df: pd.DataFrame) -> list[list]:
    """Agrupa los cobros conciliados por dia y operador, con el formato de la
    hoja CPE: una fila por dia/operador y la lista 'folio ($monto)' de los
    documentos asociados ('?' si el cobro no se concilio)."""
    filas = []
    if df.empty:
        return filas
    df = df.assign(_ORDEN_CANAL=df["CANAL"].map({"Presencial": 0, "Internet": 1}))
    claves = ["FECHA_COBRO", "ORIGEN_COBRO", "_ORDEN_CANAL", "CANAL"]
    for (fecha, origen, _, canal), g in df.groupby(claves, sort=True):
        documentos = " - ".join(
            f"{folio if pd.notna(folio) and folio else '?'} ({_fmt_clp(monto)})"
            for folio, monto in zip(g["FOLIO_SOFNET"], g["TOTAL_COBRO"])
        )
        # Solo documentos con folio identificado
        encontrados = [
            (str(folio), monto)
            for folio, monto in zip(g["FOLIO_SOFNET"], g["TOTAL_COBRO"])
            if pd.notna(folio) and folio
        ]
        folios = "-".join(f for f, _ in encontrados) or None
        bruto = sum(m for _, m in encontrados) if encontrados else None
        filas.append([
            fecha,
            canal,
            CPE_OPERADOR.get(origen, origen),
            len(g),
            g["TOTAL_COBRO"].sum(),
            g["MONTO_NETO"].sum(),
            g["MONTO_IVA"].sum(),
            0,                                  # Monto Exento
            documentos,
            folios,                             # J: folios separados por guion
            bruto,                              # K: suma de los documentos con folio
            # L-N (monto total, neto, iva) se escriben como formulas
        ])
    return filas


def escribir_hoja_cpe(wb: Workbook, nombre: str, headers: list[str], filas: list[list]):
    """Equivalente a escribirHojaCPE: crea (o limpia) la hoja, escribe
    encabezado azul, datos, formatos de fecha/monto, anchos y congela fila 1."""
    if nombre in wb.sheetnames:
        hoja = wb[nombre]
        hoja.delete_rows(1, hoja.max_row)
    else:
        hoja = wb.create_sheet(nombre)

    # Encabezado
    hoja.append(headers)
    for cell in hoja[1]:
        cell.fill = CPE_HEADER_FILL
        cell.font = CPE_HEADER_FONT
        cell.alignment = Alignment(horizontal="center")

    # Datos
    for fila in filas:
        hoja.append(fila)

    for r in range(2, len(filas) + 2):
        # L: monto total sin documento = Monto Total CPE - bruto
        # M: neto = monto total / 1,19 (redondeado) | N: iva = monto total - neto
        hoja.cell(row=r, column=12).value = f"=E{r}-N(K{r})"
        hoja.cell(row=r, column=13).value = f"=ROUND(L{r}/1.19,0)"
        hoja.cell(row=r, column=14).value = f"=L{r}-M{r}"

        hoja.cell(row=r, column=1).number_format = "DD-MM-YYYY"           # col A fecha
        for c in list(range(5, 9)) + list(range(11, 15)):                   # cols E-H y K-N montos
            hoja.cell(row=r, column=c).number_format = "#,##0"
        hoja.cell(row=r, column=9).alignment = Alignment(horizontal="left")  # col I Documentos
        hoja.cell(row=r, column=10).alignment = Alignment(horizontal="right")  # col J folios

    # Filtros en el encabezado (como en la planilla de referencia)
    hoja.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{max(len(filas) + 1, 1)}"

    # Anchos de columna (px de Sheets -> ancho de caracteres de Excel, ~7 px/caracter)
    for i, px in enumerate(CPE_ANCHOS_PX, start=1):
        hoja.column_dimensions[get_column_letter(i)].width = round(px / 7, 1)

    hoja.freeze_panes = "A2"
    return hoja


# ---------------------------------------------------------------------------
# 5a. RESUMEN POR TIPO DE DOCUMENTO Y OPERADOR
# ---------------------------------------------------------------------------

TIPO_DOC_NO_CONCILIADO = 48   # codigo SII del Comprobante de Pago Electronico
ORDEN_OPERADOR = ["MERCADO LIBRE", "TRANSBANK", "SIN OPERADOR"]
RES_BORDE = Border(*(Side(style="thin", color="000000"),) * 4)
RES_HEADER_FONT = Font(name="Arial", bold=True, color="FFFFFF", size=12)
RES_FONT = Font(name="Arial", size=12)


def _tipo_doc(valor):
    """'33', '33.0', 33.0 -> 33; vacio -> None."""
    try:
        return int(float(str(valor).strip()))
    except (TypeError, ValueError):
        return None


def construir_resumen(conciliacion: pd.DataFrame, sin_cobro: pd.DataFrame) -> pd.DataFrame:
    """Filas (Tipo Doc, N_REGISTROS, Operador, MONTO_TOTAL (BRUTO)).
    - Cobros conciliados: tipo doc y monto del documento Sofnet.
    - Cobros no conciliados: tipo doc 48 y monto de la transaccion.
    - Ventas Sofnet sin cobro: operador 'SIN OPERADOR', tipo doc y monto Sofnet."""
    registros = []
    for r in conciliacion.itertuples():
        if r.REGLA == "NO_CONCILIADO":
            registros.append((TIPO_DOC_NO_CONCILIADO, r.ORIGEN_COBRO, r.TOTAL_COBRO))
        else:
            registros.append((_tipo_doc(r.TIPO_DOC_SOFNET), r.ORIGEN_COBRO, r.TOTAL_SOFNET))
    for r in sin_cobro.itertuples():
        tipo = str(r.FOLIO).rsplit("-", 1)[-1] if pd.notna(r.FOLIO) else None
        registros.append((_tipo_doc(tipo), "SIN OPERADOR", r.TOTAL))

    df = pd.DataFrame(registros, columns=["Tipo Doc", "Operador", "MONTO"])
    if df.empty:
        return pd.DataFrame(columns=["Tipo Doc", "N_REGISTROS", "Operador", "MONTO_TOTAL (BRUTO)"])
    df["_ORDEN"] = df["Operador"].map(lambda o: ORDEN_OPERADOR.index(o) if o in ORDEN_OPERADOR else len(ORDEN_OPERADOR))
    res = (
        df.groupby(["_ORDEN", "Operador", "Tipo Doc"], dropna=False)
        .agg(N_REGISTROS=("MONTO", "size"), **{"MONTO_TOTAL (BRUTO)": ("MONTO", "sum")})
        .reset_index()
        .sort_values(["_ORDEN", "Tipo Doc"])
    )
    return res[["Tipo Doc", "N_REGISTROS", "Operador", "MONTO_TOTAL (BRUTO)"]]


def escribir_hoja_resumen(wb: Workbook, conciliacion: pd.DataFrame, sin_cobro: pd.DataFrame):
    """Pestaña 'Resumen' como primera hoja: tabla por tipo doc y operador
    (con filtros) y, dos filas mas abajo, totales por tipo doc."""
    if "Resumen" in wb.sheetnames:
        del wb["Resumen"]
    ws = wb.create_sheet("Resumen", 0)
    wb.active = 0

    detalle = construir_resumen(conciliacion, sin_cobro)
    totales = (
        detalle.groupby("Tipo Doc", dropna=False)
        .agg(N_REGISTROS=("N_REGISTROS", "sum"), **{"MONTO_TOTAL (BRUTO)": ("MONTO_TOTAL (BRUTO)", "sum")})
        .reset_index()
        .sort_values("Tipo Doc")
    )

    def escribir_tabla(fila_ini: int, tabla: pd.DataFrame) -> int:
        for c, nombre in enumerate(tabla.columns, start=1):
            cell = ws.cell(row=fila_ini, column=c, value=nombre)
            cell.font, cell.fill, cell.border = RES_HEADER_FONT, HEADER_FILL, RES_BORDE
            cell.alignment = Alignment(horizontal="center", vertical="center")
        for i, fila in enumerate(tabla.itertuples(index=False), start=1):
            for c, valor in enumerate(fila, start=1):
                if isinstance(valor, float) and pd.isna(valor):
                    valor = None
                elif tabla.columns[c - 1] in ("Tipo Doc", "N_REGISTROS", "MONTO_TOTAL (BRUTO)") and valor is not None:
                    valor = int(round(valor))
                cell = ws.cell(row=fila_ini + i, column=c, value=valor)
                cell.font, cell.border = RES_FONT, RES_BORDE
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if tabla.columns[c - 1] == "MONTO_TOTAL (BRUTO)":
                    cell.number_format = "#,##0"
        return fila_ini + len(tabla)

    fin1 = escribir_tabla(1, detalle)
    ws.auto_filter.ref = f"A1:D{max(fin1, 1)}"
    escribir_tabla(fin1 + 3, totales)

    for letra, ancho in zip("ABCD", (16, 22, 22, 26)):
        ws.column_dimensions[letra].width = ancho
    return ws


def exportar_por_mes(resultado_df: pd.DataFrame, sofnet_sin_cruce: pd.DataFrame, output_dir: str):
    os.makedirs(output_dir, exist_ok=True)

    if resultado_df.empty:
        print("No hay resultados de conciliacion para exportar.")
        return []

    resultado_df = resultado_df.copy()
    resultado_df["PERIODO"] = resultado_df["FECHA_COBRO"].apply(
        lambda d: (d.year, d.month) if pd.notna(d) else (None, None)
    )

    generados = []
    for (anio, mes), grupo in resultado_df.groupby("PERIODO"):
        if anio is None:
            continue
        nombre = f"Conciliacion_{anio}_{mes:02d}.xlsx"
        ruta = os.path.join(output_dir, nombre)

        wb = Workbook()
        ws1 = wb.active
        ws1.title = "Conciliacion"
        grupo_ordenado = grupo.drop(columns=["PERIODO"]).sort_values(
            ["ORIGEN_COBRO", "FECHA_COBRO"]
        )
        _escribir_hoja(ws1, grupo_ordenado)

        # Ventas Sofnet del mismo periodo que quedaron sin cobro asociado
        ws3 = wb.create_sheet("Sofnet_sin_cobro")
        if not sofnet_sin_cruce.empty:
            sin_cruce_mes = sofnet_sin_cruce[
                sofnet_sin_cruce["FECHA"].apply(
                    lambda d: pd.notna(d) and (d.year, d.month) == (anio, mes)
                )
            ]
        else:
            sin_cruce_mes = sofnet_sin_cruce
        _escribir_hoja(ws3, sin_cruce_mes)

        # Resumen por tipo de documento y operador (primera pestaña)
        escribir_hoja_resumen(wb, grupo_ordenado, sin_cruce_mes)

        # Hoja CPE: cobros del mes agrupados por dia y operador
        escribir_hoja_cpe(wb, "CPE Detalle", CPE_HEADERS, construir_filas_cpe(grupo_ordenado))

        # Informe CPE (formato formulario SII), ultima pestaña
        escribir_hoja_informe(wb, grupo_ordenado, anio, mes)

        wb.save(ruta)
        generados.append(ruta)
        print(f"  Generado: {ruta}  ({len(grupo_ordenado)} cobros)")

    return generados


# ---------------------------------------------------------------------------
# 5c. INFORME CPE (formato formulario SII Comprobantes de Pago Electronico)
# ---------------------------------------------------------------------------

INF_HEADERS = ["Ventas del día", "Canal", "Tot.Documentos", "Monto Total",
               "Monto Neto", "Iva", "Monto Exento", "Motivo Cambio"]
INF_CANALES = ["Presencial", "Internet"]
INF_HEADER_FILL = PatternFill(start_color="EEEEEE", end_color="EEEEEE", fill_type="solid")
INF_HEADER_FONT = Font(name="Arial", bold=True, color="0B3B5C")
INF_ZEBRA_FILL = PatternFill(start_color="F9F9F9", end_color="F9F9F9", fill_type="solid")
INF_FONT = Font(name="Arial", color="555555")
INF_BORDE = Border(*(Side(style="thin", color="DDDDDD"),) * 4)
INF_ANCHOS = [14, 13, 16, 16, 16, 16, 16, 60]


def construir_filas_informe(df: pd.DataFrame, anio: int, mes: int) -> list[list]:
    """Una fila por dia del mes y canal (Presencial / Internet).
    Los montos corregidos consideran solo los comprobantes SIN boleta/factura
    asociada; los que si tienen documento se descuentan y se informan en
    'Motivo Cambio'."""
    filas = []
    ultimo = calendar.monthrange(anio, mes)[1]
    grupos = {k: g for k, g in df.groupby(["FECHA_COBRO", "CANAL"])} if not df.empty else {}

    for dia in range(1, ultimo + 1):
        fecha = date(anio, mes, dia)
        for canal in INF_CANALES:
            g = grupos.get((fecha, canal))
            if g is None or g.empty:
                filas.append([fecha, canal, 0, 0, 0, 0, 0, ""])
                continue

            con_doc = g[g["FOLIO_SOFNET"].notna() & (g["FOLIO_SOFNET"].astype(str) != "")]
            sin_doc = g.drop(con_doc.index)

            total = int(round(sin_doc["TOTAL_COBRO"].sum()))
            neto = int(round(total / 1.19))
            iva = total - neto

            motivo = ""
            if not con_doc.empty:
                folios = "-".join(str(f) for f in con_doc["FOLIO_SOFNET"])
                motivo = (f"Se descuentan {len(con_doc)} de {len(g)} comprobantes con "
                          f"boleta/factura emitida (folios {folios}; "
                          f"{_fmt_clp(con_doc['TOTAL_COBRO'].sum())})")

            filas.append([fecha, canal, len(sin_doc), total, neto, iva, 0, motivo])
    return filas


def escribir_hoja_informe(wb: Workbook, grupo: pd.DataFrame, anio: int, mes: int):
    """Agrega la pestaña 'Informe CPE' (formato formulario SII) al libro del mes."""
    if "Informe CPE" in wb.sheetnames:
        del wb["Informe CPE"]
    ws = wb.create_sheet("Informe CPE")

    ws.append(INF_HEADERS)
    for cell in ws[1]:
        cell.fill, cell.font, cell.border = INF_HEADER_FILL, INF_HEADER_FONT, INF_BORDE
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 32

    for i, fila in enumerate(construir_filas_informe(grupo, anio, mes)):
        ws.append(fila)
        r = ws.max_row
        for c in range(1, len(INF_HEADERS) + 1):
            cell = ws.cell(row=r, column=c)
            cell.font, cell.border = INF_FONT, INF_BORDE
            cell.alignment = Alignment(horizontal="right" if 3 <= c <= 7 else "left",
                                       vertical="center", wrap_text=(c == 8))
            if (i // 2) % 2 == 0:          # zebra por dia (2 filas por dia)
                cell.fill = INF_ZEBRA_FILL
        ws.cell(row=r, column=1).number_format = "DD-MM-YYYY"
        for c in range(3, 8):
            ws.cell(row=r, column=c).number_format = "0"

    for i, ancho in enumerate(INF_ANCHOS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = ancho
    ws.freeze_panes = "A2"

    return ws


# ---------------------------------------------------------------------------
# 6. MAIN
# ---------------------------------------------------------------------------


def main():
    parser = argparse.ArgumentParser(description="Conciliacion de ventas Sofnet vs Transbank/Mercado Libre")
    parser.add_argument("--base", default=DEFAULT_BASE_DIR, help="Carpeta raiz 'Ventas'")
    args = parser.parse_args()

    base_dir = args.base
    output_dir = os.path.join(base_dir, OUTPUT_SUBDIR)

    print(f"Carpeta base: {base_dir}\n")
    print("Cargando archivos...")
    sofnet_df, cobros_df = cargar_todas_las_fuentes(base_dir)
    print(f"\nVentas Sofnet cargadas: {len(sofnet_df)}")
    print(f"Cobros (Transbank + Mercado Libre) cargados: {len(cobros_df)}\n")

    if sofnet_df.empty or cobros_df.empty:
        print("No hay suficientes datos para conciliar. Revisa las rutas de entrada.")
        sys.exit(1)

    print("Ejecutando motor de conciliacion (R1 -> R4)...")
    resultado_df, sofnet_sin_cruce = conciliar(sofnet_df, cobros_df)

    total = len(resultado_df)
    conciliados = (resultado_df["REGLA"] != "NO_CONCILIADO").sum()
    print(f"  Conciliados: {conciliados}/{total}")
    print(f"  Sin conciliar: {total - conciliados}/{total}\n")

    print("Exportando reportes por mes...")
    exportar_por_mes(resultado_df, sofnet_sin_cruce, output_dir)
    print("\nProceso finalizado.")


if __name__ == "__main__":
    main()