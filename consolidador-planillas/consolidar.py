#!/usr/bin/env python3
"""
Consolidador y Validador de Planillas
Procesamiento por lotes de libros Excel, validación declarativa, 
segregación de inconsistencias y generación de reporte ejecutivo en 3 hojas.
"""

import argparse
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Set, Tuple

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import yaml


def cargar_configuracion(ruta_config: Path) -> Dict[str, Any]:
    if not ruta_config.exists():
        raise FileNotFoundError(f"No se encontró el archivo de reglas en: {ruta_config}")
    with open(ruta_config, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def normalizar_texto(val: Any) -> str:
    if val is None or pd.isna(val):
        return ""
    return str(val).strip()


def formatear_fecha(val: Any) -> str:
    if val is None or pd.isna(val) or val == "":
        return ""
    if isinstance(val, (datetime, pd.Timestamp)):
        return val.strftime("%Y-%m-%d")
    val_str = str(val).strip()
    try:
        dt = pd.to_datetime(val_str, errors="coerce")
        if pd.notna(dt):
            return dt.strftime("%Y-%m-%d")
    except Exception:
        pass
    return val_str


# Motor de validación declarativa fila por fila
def validar_fila(
    fila: Dict[str, Any],
    reglas: Dict[str, Any],
    claves_registradas: Dict[Tuple[str, ...], Dict[str, Any]],
    archivo: str,
    num_fila: int
) -> List[str]:
    errores = []

    # 1. Columnas obligatorias
    cols_req = reglas.get("columnas_requeridas", [])
    for col in cols_req:
        val = fila.get(col)
        if val is None or pd.isna(val) or str(val).strip() == "":
            errores.append(f"Campo obligatorio '{col}' ausente o vacío")

    # 2. Validación de turnos
    turnos_ok = reglas.get("validaciones", {}).get("turnos_permitidos", [])
    turno_val = normalizar_texto(fila.get("turno"))
    if turno_val and turnos_ok and turno_val not in turnos_ok:
        errores.append(f"Turno no permitido: '{turno_val}' (Permitidos: {', '.join(turnos_ok)})")

    # 3. Validación numérica: horas_trabajadas
    regla_ht = reglas.get("validaciones", {}).get("horas_trabajadas", {})
    val_ht = fila.get("horas_trabajadas")
    if val_ht is not None and not pd.isna(val_ht) and str(val_ht).strip() != "":
        try:
            num_ht = float(val_ht)
            min_ht = regla_ht.get("min", 0)
            max_ht = regla_ht.get("max", 24)
            if num_ht < min_ht or num_ht > max_ht:
                errores.append(f"Horas trabajadas ({num_ht}) fuera de rango permitido [{min_ht} - {max_ht}]")
        except (ValueError, TypeError):
            errores.append(f"Horas trabajadas no numéricas: '{val_ht}'")

    # 4. Validación numérica: horas_extra
    regla_he = reglas.get("validaciones", {}).get("horas_extra", {})
    val_he = fila.get("horas_extra")
    if val_he is not None and not pd.isna(val_he) and str(val_he).strip() != "":
        try:
            num_he = float(val_he)
            min_he = regla_he.get("min", 0)
            max_he = regla_he.get("max", 12)
            if num_he < min_he or num_he > max_he:
                errores.append(f"Horas extra ({num_he}) fuera de rango permitido [{min_he} - {max_he}]")
        except (ValueError, TypeError):
            errores.append(f"Horas extra no numéricas: '{val_he}'")

    # 5. Detección de duplicados por clave compuesta
    cols_clave = reglas.get("clave_unicidad", [])
    if cols_clave and not errores:
        valores_clave = tuple(normalizar_texto(fila.get(c)) for c in cols_clave)
        if any(v == "" for v in valores_clave):
            pass
        elif valores_clave in claves_registradas:
            origen_previo = claves_registradas[valores_clave]
            desc_clave = ", ".join(f"{c}={v}" for c, v in zip(cols_clave, valores_clave))
            errores.append(
                f"Clave duplicada ({desc_clave}). Registrado previamente en {origen_previo['archivo']} (Fila {origen_previo['fila']})"
            )
        else:
            claves_registradas[valores_clave] = {"archivo": archivo, "fila": num_fila}

    return errores


def procesar_archivos(
    carpeta_entrada: Path,
    reglas: Dict[str, Any]
) -> Tuple[pd.DataFrame, pd.DataFrame, Dict[str, Any]]:
    archivos = sorted([
        f for f in carpeta_entrada.glob("*.xlsx") 
        if not f.name.startswith("~$")
    ])

    if not archivos:
        raise ValueError(f"No se encontraron archivos .xlsx en {carpeta_entrada}")

    registros_validos = []
    registros_errores = []
    claves_registradas: Dict[Tuple[str, ...], Dict[str, Any]] = {}
    estadisticas_archivos = {}

    for ruta_archivo in archivos:
        nombre_archivo = ruta_archivo.name
        try:
            df = pd.read_excel(ruta_archivo)
        except Exception as e:
            print(f"[ERROR] No se pudo leer {nombre_archivo}: {e}")
            continue

        # Normalización de encabezados
        df.columns = [
            str(c).strip().lower().replace(" ", "_").replace("-", "_") 
            for c in df.columns
        ]

        total_filas_archivo = len(df)
        validas_archivo = 0
        errores_archivo = 0

        for idx, row in df.iterrows():
            num_fila_excel = idx + 2
            fila_dict = row.to_dict()

            # Sanitización de fechas
            if "fecha" in fila_dict:
                fila_dict["fecha"] = formatear_fecha(fila_dict["fecha"])

            errores = validar_fila(
                fila=fila_dict,
                reglas=reglas,
                claves_registradas=claves_registradas,
                archivo=nombre_archivo,
                num_fila=num_fila_excel
            )

            registro_con_metadata = dict(fila_dict)
            registro_con_metadata["archivo_origen"] = nombre_archivo
            registro_con_metadata["fila_origen"] = num_fila_excel

            if errores:
                errores_archivo += 1
                registro_con_metadata["motivo_rechazo"] = " | ".join(errores)
                registros_errores.append(registro_con_metadata)
            else:
                validas_archivo += 1
                registros_validos.append(registro_con_metadata)

        estadisticas_archivos[nombre_archivo] = {
            "total": total_filas_archivo,
            "validas": validas_archivo,
            "rechazadas": errores_archivo
        }

    df_validos = pd.DataFrame(registros_validos)
    df_errores = pd.DataFrame(registros_errores)

    resumen = {
        "archivos_totales": len(archivos),
        "estadisticas_archivos": estadisticas_archivos,
        "total_leidos": len(registros_validos) + len(registros_errores),
        "total_validos": len(registros_validos),
        "total_rechazados": len(registros_errores)
    }

    return df_validos, df_errores, resumen


# Estilización y formateo profesional con openpyxl
def formatear_hoja_resumen(ws: openpyxl.worksheet.worksheet.Worksheet, resumen: Dict[str, Any], df_validos: pd.DataFrame, df_errores: pd.DataFrame):
    ws.views.sheetView[0].showGridLines = True

    font_titulo = Font(name="Calibri", size=14, bold=True, color="1F4E79")
    font_sub = Font(name="Calibri", size=9, italic=True, color="595959")
    font_header_sec = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    font_bold = Font(name="Calibri", size=9, bold=True, color="000000")
    font_regular = Font(name="Calibri", size=9, color="000000")

    fill_header = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    fill_sub_header = PatternFill(start_color="2B3A4A", end_color="2B3A4A", fill_type="solid")
    fill_kpi_label = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9")
    )

    ws["A1"] = "INFORME EJECUTIVO DE CONSOLIDACIÓN Y CALIDAD DE DATOS"
    ws["A1"].font = font_titulo
    ws["A2"] = f"Generado el: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | Auditoría automática por lotes"
    ws["A2"].font = font_sub

    # Sección 1: KPIs Globales
    ws["A4"] = "INDICADORES GLOBALES DE PROCESAMIENTO"
    ws["A4"].font = Font(name="Calibri", size=10, bold=True, color="1F4E79")

    headers_kpi = ["Métrica de Control", "Valor"]
    for col_idx, h in enumerate(headers_kpi, start=1):
        c = ws.cell(row=5, column=col_idx, value=h)
        c.font = font_header_sec
        c.fill = fill_header
        c.alignment = Alignment(horizontal="left", vertical="center")

    tasa_rechazo = (
        (resumen["total_rechazados"] / resumen["total_leidos"] * 100)
        if resumen["total_leidos"] > 0 else 0.0
    )

    kpi_data = [
        ("Total Archivos Procesados", resumen["archivos_totales"]),
        ("Total Registros Leídos", resumen["total_leidos"]),
        ("Registros Válidos Consolidados", resumen["total_validos"]),
        ("Registros Rechazados (Inconsistencias)", resumen["total_rechazados"]),
        ("Tasa Global de Error / Rechazo", f"{tasa_rechazo:.2f}%")
    ]

    for r_idx, (k, v) in enumerate(kpi_data, start=6):
        c_k = ws.cell(row=r_idx, column=1, value=k)
        c_v = ws.cell(row=r_idx, column=2, value=v)
        c_k.font = font_bold
        c_k.fill = fill_kpi_label
        c_k.border = thin_border
        c_v.font = font_regular
        c_v.border = thin_border
        c_v.alignment = Alignment(horizontal="right")

    # Sección 2: Desglose por Sector
    row_sec2 = 13
    ws.cell(row=row_sec2, column=1, value="DESGLOSE OPERATIVO POR SECTOR").font = Font(name="Calibri", size=10, bold=True, color="1F4E79")

    headers_sector = ["Sector", "Registros Válidos", "Total Horas Normales", "Total Horas Extra"]
    for col_idx, h in enumerate(headers_sector, start=1):
        c = ws.cell(row=row_sec2 + 1, column=col_idx, value=h)
        c.font = font_header_sec
        c.fill = fill_sub_header
        c.alignment = Alignment(horizontal="center" if col_idx > 1 else "left", vertical="center")

    curr_row = row_sec2 + 2
    if not df_validos.empty and "sector" in df_validos.columns:
        res_sector = df_validos.groupby("sector", as_index=False).agg(
            registros=("legajo", "count"),
            horas_normales=("horas_trabajadas", "sum"),
            horas_extra=("horas_extra", "sum")
        )
        for _, s_row in res_sector.iterrows():
            c1 = ws.cell(row=curr_row, column=1, value=str(s_row["sector"]))
            c2 = ws.cell(row=curr_row, column=2, value=int(s_row["registros"]))
            c3 = ws.cell(row=curr_row, column=3, value=float(s_row["horas_normales"]))
            c4 = ws.cell(row=curr_row, column=4, value=float(s_row["horas_extra"]))

            for c in [c1, c2, c3, c4]:
                c.border = thin_border
                c.font = font_regular
            c2.alignment = Alignment(horizontal="right")
            c3.alignment = Alignment(horizontal="right")
            c4.alignment = Alignment(horizontal="right")
            curr_row += 1
    else:
        ws.cell(row=curr_row, column=1, value="Sin datos disponibles").font = font_regular
        curr_row += 1

    # Sección 3: Top Incidencias
    curr_row += 2
    ws.cell(row=curr_row, column=1, value="TOP INCIDENCIAS DE VALIDACIÓN DETECTADAS").font = Font(name="Calibri", size=10, bold=True, color="1F4E79")
    curr_row += 1

    headers_err = ["Causa de Inconsistencia", "Ocurrencias", "% Participación"]
    for col_idx, h in enumerate(headers_err, start=1):
        c = ws.cell(row=curr_row, column=col_idx, value=h)
        c.font = font_header_sec
        c.fill = fill_sub_header
        c.alignment = Alignment(horizontal="center" if col_idx > 1 else "left", vertical="center")
    curr_row += 1

    if not df_errores.empty and "motivo_rechazo" in df_errores.columns:
        motivos_expandidos = []
        for m in df_errores["motivo_rechazo"].dropna():
            for sub_m in str(m).split(" | "):
                motivos_expandidos.append(sub_m.split("(")[0].strip())
        
        serie_motivos = pd.Series(motivos_expandidos).value_counts()
        total_motivos = len(motivos_expandidos)
        for motivo, cant in serie_motivos.head(5).items():
            pct = (cant / total_motivos * 100) if total_motivos > 0 else 0
            c1 = ws.cell(row=curr_row, column=1, value=motivo)
            c2 = ws.cell(row=curr_row, column=2, value=int(cant))
            c3 = ws.cell(row=curr_row, column=3, value=f"{pct:.1f}%")
            for c in [c1, c2, c3]:
                c.border = thin_border
                c.font = font_regular
            c2.alignment = Alignment(horizontal="right")
            c3.alignment = Alignment(horizontal="right")
            curr_row += 1
    else:
        c = ws.cell(row=curr_row, column=1, value="No se detectaron inconsistencias.")
        c.font = font_regular
        c.border = thin_border

    ws.column_dimensions["A"].width = 46
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["C"].width = 24
    ws.column_dimensions["D"].width = 24


def formatear_hoja_datos(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    df: pd.DataFrame,
    es_errores: bool = False
):
    ws.views.sheetView[0].showGridLines = True

    header_color = "78281F" if es_errores else "1F4E79"
    fill_header = PatternFill(start_color=header_color, end_color=header_color, fill_type="solid")
    font_header = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
    font_cell = Font(name="Calibri", size=9, color="000000")
    
    fill_error_row = PatternFill(start_color="FDEDEC", end_color="FDEDEC", fill_type="solid") if es_errores else None

    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9")
    )

    if df.empty:
        ws.cell(row=1, column=1, value="Sin registros").font = font_cell
        return

    # Si es hoja de errores, ubicar motivo_rechazo y metadata al inicio
    cols = list(df.columns)
    if es_errores and "motivo_rechazo" in cols:
        cols.remove("motivo_rechazo")
        if "archivo_origen" in cols:
            cols.remove("archivo_origen")
        if "fila_origen" in cols:
            cols.remove("fila_origen")
        cols = ["archivo_origen", "fila_origen", "motivo_rechazo"] + cols

    # Encabezados
    for col_idx, col_name in enumerate(cols, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name.upper().replace("_", " "))
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=False)
        cell.border = thin_border

    # Filas
    for row_idx, (_, row_data) in enumerate(df[cols].iterrows(), start=2):
        for col_idx, col_name in enumerate(cols, start=1):
            val = row_data[col_name]
            cell = ws.cell(row=row_idx, column=col_idx)

            if pd.isna(val):
                cell.value = ""
            elif isinstance(val, (int, float)):
                cell.value = val
                if col_name in ["horas_trabajadas", "horas_extra"]:
                    cell.number_format = "0.0"
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else:
                cell.value = str(val)
                cell.alignment = Alignment(horizontal="left", vertical="center")

            cell.font = font_cell
            cell.border = thin_border
            if fill_error_row:
                cell.fill = fill_error_row

    # Autofiltro
    max_col_letter = get_column_letter(len(cols))
    ws.auto_filter.ref = f"A1:{max_col_letter}{len(df) + 1}"

    # Ancho de columnas adaptativo
    for col_idx, col_name in enumerate(cols, start=1):
        col_letter = get_column_letter(col_idx)
        max_len = max(len(col_name), 10)
        col_data = df[col_name].dropna().astype(str)
        if not col_data.empty:
            content_max = col_data.map(len).max()
            max_len = max(max_len, min(content_max + 4, 50))
        ws.column_dimensions[col_letter].width = max_len + 3


def exportar_libro_consolidado(
    df_validos: pd.DataFrame,
    df_errores: pd.DataFrame,
    resumen: Dict[str, Any],
    ruta_salida: Path
):
    wb = openpyxl.Workbook()

    # Hoja 1: Resumen
    ws_resumen = wb.active
    ws_resumen.title = "Resumen"
    formatear_hoja_resumen(ws_resumen, resumen, df_validos, df_errores)

    # Hoja 2: Consolidado
    ws_consolidado = wb.create_sheet(title="Consolidado")
    formatear_hoja_datos(ws_consolidado, df_validos, es_errores=False)

    # Hoja 3: Errores
    ws_errores = wb.create_sheet(title="Errores")
    formatear_hoja_datos(ws_errores, df_errores, es_errores=True)

    ruta_salida.parent.mkdir(parents=True, exist_ok=True)
    wb.save(ruta_salida)


def main():
    parser = argparse.ArgumentParser(
        description="Consolidador y Validador Automatizado de Planillas Operativas"
    )
    parser.add_argument(
        "-i", "--input",
        type=str,
        default="./planillas_ejemplo",
        help="Directorio que contiene las planillas Excel a consolidar (default: ./planillas_ejemplo)"
    )
    parser.add_argument(
        "-o", "--output",
        type=str,
        default="./reporte_consolidado.xlsx",
        help="Ruta del archivo Excel resultante (default: ./reporte_consolidado.xlsx)"
    )
    parser.add_argument(
        "-c", "--config",
        type=str,
        default="./config/reglas.yaml",
        help="Archivo YAML con las reglas declarativas de validación (default: ./config/reglas.yaml)"
    )

    args = parser.parse_args()

    carpeta_input = Path(args.input)
    ruta_output = Path(args.output)
    ruta_config = Path(args.config)

    print("==================================================================")
    print(" CONSOLIDADOR Y VALIDADOR DE PLANILLAS - AUDITORÍA DE DATOS")
    print("==================================================================")
    print(f"[INFO] Carpeta de entrada : {carpeta_input.resolve()}")
    print(f"[INFO] Archivo de reglas  : {ruta_config.resolve()}")
    print(f"[INFO] Archivo de salida  : {ruta_output.resolve()}")

    try:
        reglas = cargar_configuracion(ruta_config)
    except Exception as e:
        print(f"[FATAL] Error al cargar reglas de validación: {e}")
        sys.exit(1)

    try:
        df_validos, df_errores, resumen = procesar_archivos(carpeta_input, reglas)
    except Exception as e:
        print(f"[FATAL] Error en la consolidación: {e}")
        sys.exit(1)

    print("------------------------------------------------------------------")
    print(f"[OK] Archivos procesados   : {resumen['archivos_totales']}")
    print(f"[OK] Total filas leídas    : {resumen['total_leidos']}")
    print(f"[OK] Registros válidos     : {resumen['total_validos']}")
    print(f"[WARN] Registros rechazados : {resumen['total_rechazados']}")
    
    tasa = (resumen['total_rechazados'] / resumen['total_leidos'] * 100) if resumen['total_leidos'] > 0 else 0
    print(f"[MÉTRICA] Tasa de descarte : {tasa:.2f}%")

    try:
        exportar_libro_consolidado(df_validos, df_errores, resumen, ruta_output)
        print("------------------------------------------------------------------")
        print(f"[EXITO] Reporte generado satisfactoriamente:")
        print(f"        -> {ruta_output.resolve()}")
        print("        -> Hojas: ['Resumen', 'Consolidado', 'Errores']")
    except Exception as e:
        print(f"[FATAL] Error al exportar archivo Excel: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()

