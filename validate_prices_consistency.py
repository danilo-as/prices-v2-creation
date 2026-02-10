#!/usr/bin/env python3
"""
Script para validar que los precios del archivo Excel coincidan con los del CSV de la BD.
Compara campo por campo y reporta diferencias.
Para campos como Product e Incoterm, valida contra la BD usando name o chemical_code/code.
"""

import pandas as pd
import psycopg2
from dotenv import load_dotenv
import os
from typing import Dict, List, Optional, Tuple
from pathlib import Path

load_dotenv()

# Mapeo de columnas del Excel a columnas del CSV
# Formato: "Columna Excel" -> "columna_csv"
EXCEL_TO_CSV_MAPPING = {
    # Campos directos
    "Code": "internal_code",
    "Full Name": "full_name",
    "Short Name (40 char max)": "short_name",
    "Spec": "specification",
    "Grade": "grade",
    "Sustainable": "is_sustainable",
    "IOSCO Audit": "is_iosco_assured",
    "Is Public?": "is_public",
    "Is Active": "is_active",
    "isSpot": "is_spot",

    # Campos de entidades
    "Market": "market_name",
    "Price Category": "price_category_name",
    "Product": "product_name",
    "Sub-Product": "sub_product_name",
    "Capacity (Nominal/Anode/Cathode)": "capacity_name",
    "Cell Format": "cell_format_name",
    "Feedstock": "feedstock_name",
    "Purity": "purity_name",
    "Thickness": "thickness_name",
    "Mesh size": "mesh_size_name",
    "Service": "service_name",
    "Incoterm": "incoterm_name",
    "Region": "region_name",
    "Country(OnlyforDiffs)": "country_name",
    "Trade Type?": "trade_type_name",
    "Price Type": "price_type_name",
    "UOM": "unit_of_measure_name",
    "Frequency": "frequency_name",
    "Default Currency": "default_currency_id",
}

# Campos para los cuales se generará UPDATE SQL
UPDATE_FIELDS = {
    "Full Name": "full_name",
    "Short Name (40 char max)": "short_name",
    "Country(OnlyforDiffs)": "country_id",
    "Grade": "grade",
}

# Campos que requieren validación contra la BD (name o chemical_code/code/iso_code)
FIELDS_WITH_DB_LOOKUP = {
    "Product": {"schema": "prices_assessment", "table": "product", "columns": ["name", "chemical_code"]},
    "Sub-Product": {"schema": "prices_assessment", "table": "sub_product", "columns": ["name", "chemical_code"]},
    "Incoterm": {"schema": "prices_assessment", "table": "incoterm", "columns": ["name", "code"]},
    "Default Currency": {"schema": "primary_data", "table": "currency", "columns": ["name", "iso_code"]},
    "Country(OnlyforDiffs)": {"schema": "prices_assessment", "table": "country", "columns": ["name"]},
    "UOM": {"schema": "prices_assessment", "table": "unit_of_measure", "columns": ["name", "code"]},
}


def escape_sql_string(value: str) -> str:
    """Escapa comillas simples para SQL."""
    if value is None:
        return None
    return value.replace("'", "''")


def generate_update_sql(differences_records: List[Dict], db_lookups: Dict[str, Dict[str, str]]) -> str:
    """
    Genera SQL con UPDATEs para los campos especificados en UPDATE_FIELDS.
    """
    lines = ["-- UPDATE para tabla prices_assessment.grades"]
    lines.append("-- Solo actualiza: full_name, short_name, country_id, grade")
    lines.append(f"-- Total de registros con diferencias en estos campos\n")

    update_count = 0

    for record in differences_records:
        price_id = record["price_id"]
        code = record["code"]

        # Filtrar solo las diferencias de los campos que nos interesan
        updates = []
        for diff in record["differences"]:
            field = diff["field"]
            if field in UPDATE_FIELDS:
                db_col = UPDATE_FIELDS[field]
                excel_value = diff["excel_value"]

                if field == "Country(OnlyforDiffs)":
                    # Para country necesitamos el ID, no el nombre
                    if excel_value:
                        country_lookup = db_lookups.get("Country(OnlyforDiffs)", {})
                        country_id = country_lookup.get(excel_value.lower().strip())
                        if country_id:
                            updates.append(f"country_id = '{country_id}'")
                            updates.append(f"country_name = '{escape_sql_string(excel_value)}'")
                        else:
                            lines.append(f"-- ADVERTENCIA: País '{excel_value}' no encontrado en BD para {code}")
                    else:
                        updates.append("country_id = NULL")
                        updates.append("country_name = NULL")
                else:
                    # Campos de texto simples
                    if excel_value is None:
                        updates.append(f"{db_col} = NULL")
                    else:
                        updates.append(f"{db_col} = '{escape_sql_string(excel_value)}'")

        if updates:
            update_count += 1
            lines.append(f"-- Código: {code}")
            lines.append(f"UPDATE prices_assessment.grades SET")
            lines.append(f"    {', '.join(updates)}")
            lines.append(f"WHERE id = '{price_id}';")
            lines.append("")

    # Actualizar el encabezado con el conteo real
    lines[2] = f"-- Total de registros a actualizar: {update_count}\n"

    return "\n".join(lines)


def get_db_connection():
    """Establece conexión con la base de datos PostgreSQL."""
    return psycopg2.connect(
        host=os.getenv("DB_HOST"),
        port=os.getenv("DB_PORT"),
        dbname=os.getenv("DB_NAME"),
        user=os.getenv("DB_USER"),
        password=os.getenv("DB_PASSWORD"),
    )


def load_lookup_table(conn, schema: str, table: str, columns: list) -> Dict[str, str]:
    """
    Carga una tabla de lookup y retorna un diccionario de valor (lowercase) -> id.
    Esto permite buscar el mismo registro por diferentes columnas (name, chemical_code, code).
    También incluye el ID como clave para poder comparar contra UUIDs del CSV.
    """
    cur = conn.cursor()
    lookup = {}

    for column in columns:
        cur.execute(f"SELECT id, {column} FROM {schema}.{table} WHERE {column} IS NOT NULL")
        for row in cur.fetchall():
            uuid_val = str(row[0])
            value = row[1]
            if value:
                lookup[value.lower().strip()] = uuid_val
            # También agregar el UUID como clave para poder comparar contra IDs del CSV
            lookup[uuid_val.lower()] = uuid_val

    cur.close()
    return lookup


def load_db_lookups(conn) -> Dict[str, Dict[str, str]]:
    """Carga los lookups de la BD para los campos que lo requieren."""
    lookups = {}
    for field, config in FIELDS_WITH_DB_LOOKUP.items():
        lookups[field] = load_lookup_table(
            conn, config["schema"], config["table"], config["columns"]
        )
        print(f"   Cargado {field}: {len(lookups[field])} valores")
    return lookups


def get_cell_value(row, column: str) -> Optional[str]:
    """Obtiene el valor de una celda, retornando None si está vacío o NaN."""
    if column not in row.index:
        return None
    val = row[column]
    if pd.isna(val):
        return None
    return str(val).strip() if str(val).strip() else None


def normalize_value(value: Optional[str]) -> Optional[str]:
    """Normaliza un valor para comparación (lowercase, strip)."""
    if value is None:
        return None
    return value.lower().strip()


def normalize_bool(value) -> bool:
    """Normaliza un valor booleano."""
    if pd.isna(value):
        return False
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return value == 1 or value == 1.0
    val_str = str(value).lower().strip()
    return val_str in ("1", "1.0", "true", "yes", "si")


def values_match_with_db(excel_val: Optional[str], csv_val: Optional[str], db_lookup: Dict[str, str]) -> bool:
    """
    Compara valores del Excel con el CSV usando la BD como referencia.
    Ambos valores deben resolver al mismo ID en la BD.
    """
    if excel_val is None and csv_val is None:
        return True

    if excel_val is None or csv_val is None:
        return False

    excel_norm = normalize_value(excel_val)
    csv_norm = normalize_value(csv_val)

    # Si son iguales directamente
    if excel_norm == csv_norm:
        return True

    # Buscar ambos en la BD y comparar IDs
    excel_id = db_lookup.get(excel_norm)
    csv_id = db_lookup.get(csv_norm)

    if excel_id and csv_id and excel_id == csv_id:
        return True

    return False


def values_match(excel_val: Optional[str], csv_val: Optional[str], is_bool: bool = False) -> bool:
    """Compara valores del Excel con el CSV."""
    if is_bool:
        excel_bool = normalize_bool(excel_val)
        csv_bool = normalize_bool(csv_val)
        return excel_bool == csv_bool

    excel_norm = normalize_value(excel_val)
    csv_norm = normalize_value(csv_val)

    if excel_norm == csv_norm:
        return True

    if excel_norm is None and csv_norm is None:
        return True

    return False


def compare_records(excel_row, csv_row, db_lookups: Dict[str, Dict[str, str]]) -> List[Dict]:
    """Compara un registro del Excel con uno del CSV y retorna las diferencias."""
    differences = []

    bool_fields = ["Sustainable", "IOSCO Audit", "Is Public?", "Is Active", "isSpot"]

    for excel_col, csv_col in EXCEL_TO_CSV_MAPPING.items():
        if excel_col == "Code":  # Skip code, es la clave
            continue

        excel_val = get_cell_value(excel_row, excel_col)
        csv_val = get_cell_value(csv_row, csv_col) if csv_col in csv_row.index else None

        is_bool = excel_col in bool_fields

        # Campos que requieren validación contra la BD
        if excel_col in FIELDS_WITH_DB_LOOKUP:
            db_lookup = db_lookups.get(excel_col, {})
            if not values_match_with_db(excel_val, csv_val, db_lookup):
                differences.append({
                    "field": excel_col,
                    "excel_value": excel_val,
                    "csv_value": csv_val,
                    "note": "No resuelven al mismo ID en BD"
                })
        elif not values_match(excel_val, csv_val, is_bool):
            differences.append({
                "field": excel_col,
                "excel_value": excel_val,
                "csv_value": csv_val,
            })

    return differences


def main():
    # Rutas de archivos
    excel_path = "docs/Prices 2.0 Data Structure OLD.xlsx"
    csv_path = "docs/all_prices.csv"

    # Leer archivos
    print(f"Leyendo archivo Excel: {excel_path}")
    df_excel = pd.read_excel(excel_path, sheet_name="Sheet1")
    print(f"Total de filas en Excel: {len(df_excel)}\n")

    print(f"Leyendo archivo CSV: {csv_path}")
    df_csv = pd.read_csv(csv_path)
    print(f"Total de filas en CSV: {len(df_csv)}\n")

    # Conectar a la BD y cargar lookups
    print("Conectando a la base de datos...")
    conn = get_db_connection()
    print("Conexión exitosa\n")

    print("Cargando lookups de la BD...")
    db_lookups = load_db_lookups(conn)
    conn.close()
    print()

    # Crear índice por internal_code para el CSV
    csv_by_code = {}
    for idx, row in df_csv.iterrows():
        code = get_cell_value(row, "internal_code")
        if code:
            csv_by_code[code] = row

    print("=" * 80)
    print("VALIDANDO CONSISTENCIA DE PRECIOS")
    print("=" * 80)

    # Contadores
    total_excel = 0
    matched = 0
    not_in_csv = 0
    with_differences = 0
    all_differences = []
    codes_not_in_csv = []

    for idx, excel_row in df_excel.iterrows():
        excel_code = get_cell_value(excel_row, "Code")

        if not excel_code:
            continue

        total_excel += 1

        # Buscar en CSV
        if excel_code not in csv_by_code:
            not_in_csv += 1
            codes_not_in_csv.append(excel_code)
            continue

        csv_row = csv_by_code[excel_code]
        differences = compare_records(excel_row, csv_row, db_lookups)
        price_id = get_cell_value(csv_row, "id")

        if differences:
            with_differences += 1
            all_differences.append({
                "code": excel_code,
                "price_id": price_id,
                "row_num": idx + 2,  # +2 porque Excel empieza en 1 y tiene header
                "differences": differences
            })
        else:
            matched += 1

    # Resumen
    print(f"\nRESUMEN:")
    print(f"   Total registros en Excel (con código): {total_excel}")
    print(f"   Registros que coinciden exactamente: {matched}")
    print(f"   Registros NO encontrados en CSV: {not_in_csv}")
    print(f"   Registros con diferencias: {with_differences}")

    # Mostrar códigos no encontrados en CSV
    if codes_not_in_csv:
        print(f"\nCódigos del Excel NO encontrados en CSV ({len(codes_not_in_csv)}):")
        for code in codes_not_in_csv[:20]:
            print(f"   - {code}")
        if len(codes_not_in_csv) > 20:
            print(f"   ... y {len(codes_not_in_csv) - 20} más")

    # Mostrar diferencias encontradas
    if all_differences:
        print(f"\nDIFERENCIAS ENCONTRADAS ({len(all_differences)} registros):")
        print("-" * 80)

        # Agrupar diferencias por campo para resumen
        diff_by_field = {}
        for record in all_differences:
            for diff in record["differences"]:
                field = diff["field"]
                if field not in diff_by_field:
                    diff_by_field[field] = 0
                diff_by_field[field] += 1

        print("\nResumen de diferencias por campo:")
        for field, count in sorted(diff_by_field.items(), key=lambda x: -x[1]):
            print(f"   {field}: {count} diferencias")

        # Mostrar detalle de las primeras diferencias
        print("\nDetalle de las primeras 20 diferencias:")
        print("-" * 80)

        shown = 0
        for record in all_differences:
            if shown >= 20:
                break

            print(f"\nCódigo: {record['code']} | ID: {record['price_id']} (Fila Excel: {record['row_num']})")
            for diff in record["differences"]:
                note = f" [{diff.get('note')}]" if diff.get('note') else ""
                print(f"   {diff['field']}:{note}")
                print(f"      Excel: '{diff['excel_value']}'")
                print(f"      CSV:   '{diff['csv_value']}'")
            shown += 1

        if len(all_differences) > 20:
            print(f"\n... y {len(all_differences) - 20} registros más con diferencias")

    # Guardar reporte completo
    if all_differences or codes_not_in_csv:
        report_path = Path("reports")
        report_path.mkdir(exist_ok=True)

        report_file = report_path / "price_validation_report.txt"
        with open(report_file, "w") as f:
            f.write("REPORTE DE VALIDACIÓN DE PRECIOS\n")
            f.write("=" * 80 + "\n\n")
            f.write(f"Total registros en Excel (con código): {total_excel}\n")
            f.write(f"Registros que coinciden exactamente: {matched}\n")
            f.write(f"Registros NO encontrados en CSV: {not_in_csv}\n")
            f.write(f"Registros con diferencias: {with_differences}\n\n")

            if codes_not_in_csv:
                f.write(f"CÓDIGOS NO ENCONTRADOS EN CSV ({len(codes_not_in_csv)}):\n")
                f.write("-" * 80 + "\n")
                for code in codes_not_in_csv:
                    f.write(f"   {code}\n")
                f.write("\n")

            if all_differences:
                f.write(f"DIFERENCIAS DETALLADAS ({len(all_differences)} registros):\n")
                f.write("-" * 80 + "\n")
                for record in all_differences:
                    f.write(f"\nCódigo: {record['code']} | ID: {record['price_id']} (Fila Excel: {record['row_num']})\n")
                    for diff in record["differences"]:
                        note = f" [{diff.get('note')}]" if diff.get('note') else ""
                        f.write(f"   {diff['field']}:{note}\n")
                        f.write(f"      Excel: '{diff['excel_value']}'\n")
                        f.write(f"      CSV:   '{diff['csv_value']}'\n")

        print(f"\nReporte completo guardado en: {report_file}")

    # Generar archivo SQL con UPDATEs
    if all_differences:
        sql_path = Path("sqls")
        sql_path.mkdir(exist_ok=True)

        sql_content = generate_update_sql(all_differences, db_lookups)
        sql_file = sql_path / "update_grades.sql"
        sql_file.write_text(sql_content)
        print(f"Archivo SQL de updates generado: {sql_file}")

    return all_differences, codes_not_in_csv


if __name__ == "__main__":
    main()
