#!/usr/bin/env python3
"""
Script para crear registros de grades a partir del archivo Excel 'Prices 2.0 Data Structure.xlsx'.
Genera archivos SQL con INSERTs para la tabla grades.
"""

import pandas as pd
import psycopg2
from dotenv import load_dotenv
import os
from typing import Dict, Optional, Tuple
from pathlib import Path
import hashlib

load_dotenv()

SQLS_DIR = Path("sqls")

# Mapeo de columnas del Excel a tablas de lookup
# schema: esquema de la BD, table: nombre de la tabla, columns: columnas donde buscar (case insensitive)
LOOKUP_TABLES = {
    "Price Category": {"schema": "prices_assessment", "table": "price_category", "columns": ["name"]},
    "Market": {"schema": "primary_data", "table": "market", "columns": ["name"]},
    "Product": {"schema": "prices_assessment", "table": "product", "columns": ["name", "chemical_code"]},
    "Sub-Product": {"schema": "prices_assessment", "table": "sub_product", "columns": ["name", "chemical_code"]},
    "Capacity (Nominal/Anode/Cathode)": {"schema": "prices_assessment", "table": "capacity", "columns": ["name"]},
    "Cell Format": {"schema": "prices_assessment", "table": "cell_format", "columns": ["name"]},
    "Feedstock": {"schema": "prices_assessment", "table": "feedstock", "columns": ["name"]},
    "Purity": {"schema": "prices_assessment", "table": "purity", "columns": ["name"]},
    "Thickness": {"schema": "prices_assessment", "table": "thickness", "columns": ["name"]},
    "Mesh size": {"schema": "prices_assessment", "table": "mesh_size", "columns": ["size"]},
    "Service": {"schema": "prices_assessment", "table": "service", "columns": ["name"]},
    "Incoterm": {"schema": "prices_assessment", "table": "incoterm", "columns": ["name", "code"]},
    "Region": {"schema": "prices_assessment", "table": "region", "columns": ["name"]},
    "Country(OnlyforDiffs)": {"schema": "prices_assessment", "table": "country", "columns": ["name"]},
    "Trade Type?": {"schema": "prices_assessment", "table": "trade_type", "columns": ["name"]},
    "Price Type": {"schema": "prices_assessment", "table": "price_type", "columns": ["name"]},
    "UOM": {"schema": "prices_assessment", "table": "unit_of_measure", "columns": ["name", "code"]},
    "Frequency": {"schema": "calendar", "table": "frequencies", "columns": ["name", "short_name"]},
    "Default Currency": {"schema": "primary_data", "table": "currency", "columns": ["iso_code", "name"]},
}

# Mapeo de columnas del Excel a columnas de la tabla grades
EXCEL_TO_GRADES_MAPPING = {
    "Code": "internal_code",
    "Full Name": "full_name",
    "Short Name (40 char max)": "short_name",
    "Spec": "specification",
    "Grade": "grade",
    "Assessment Launched": "assessment_launched_at",
    "Last Assessed": "last_assessed_at",
    "Sustainable": "is_sustainable",
    "IOSCO Audit": "is_iosco_assured",
    "Is Public?": "is_public",
    "Is Active": "is_active",
    "isSpot": "is_spot",
}

# Mapeo de columnas FK del Excel a columnas de grades (id y name)
FK_MAPPING = {
    "Price Category": ("price_category_id", "price_category_name"),
    "Market": ("market_id", "market_name"),
    "Product": ("product_id", "product_name"),
    "Sub-Product": ("sub_product_id", "sub_product_name"),
    "Capacity (Nominal/Anode/Cathode)": ("capacity_id", "capacity_name"),
    "Cell Format": ("cell_format_id", "cell_format_name"),
    "Feedstock": ("feedstock_id", "feedstock_name"),
    "Purity": ("purity_id", "purity_name"),
    "Thickness": ("thickness_id", "thickness_name"),
    "Mesh size": ("mesh_size_id", "mesh_size_name"),
    "Service": ("service_id", "service_name"),
    "Incoterm": ("incoterm_id", "incoterm_name"),
    "Region": ("region_id", "region_name"),
    "Country(OnlyforDiffs)": ("country_id", "country_name"),
    "Trade Type?": ("trade_type_id", "trade_type_name"),
    "Price Type": ("price_type_id", "price_type_name"),
    "UOM": ("unit_of_measure_id", "unit_of_measure_name"),
    "Frequency": ("frequency_id", "frequency_name"),
    "Default Currency": ("default_currency_id", None),  # Solo ID, no tiene _name en grades
}


def get_db_connection():
    """Establece conexión con la base de datos PostgreSQL."""
    return psycopg2.connect(
        host=os.getenv("DB_HOST"),
        port=os.getenv("DB_PORT"),
        dbname=os.getenv("DB_NAME"),
        user=os.getenv("DB_USER"),
        password=os.getenv("DB_PASSWORD"),
    )


def load_lookup_table(conn, schema: str, table: str, columns: list) -> Dict[str, Tuple[str, str]]:
    """
    Carga una tabla de lookup y retorna un diccionario de valor (lowercase) -> (id, nombre_original).
    """
    cur = conn.cursor()
    lookup = {}

    for column in columns:
        cur.execute(f"SELECT id, {column} FROM {schema}.{table} WHERE {column} IS NOT NULL")
        for row in cur.fetchall():
            uuid_val = str(row[0])
            original_name = row[1]
            if original_name:
                lookup[original_name.lower()] = (uuid_val, original_name)

    cur.close()
    return lookup


def load_all_lookups(conn) -> Dict[str, Dict[str, Tuple[str, str]]]:
    """Carga todas las tablas de lookup."""
    lookups = {}
    for excel_col, config in LOOKUP_TABLES.items():
        lookups[excel_col] = load_lookup_table(
            conn, config["schema"], config["table"], config["columns"]
        )
        print(f"   Cargado {excel_col}: {len(lookups[excel_col])} valores")
    return lookups


def get_cell_value(row, column: str) -> Optional[str]:
    """Obtiene el valor de una celda, retornando None si está vacío o NaN."""
    if column not in row.index:
        return None
    val = row[column]
    if pd.isna(val):
        return None
    return str(val).strip() if str(val).strip() else None


def get_bool_value(row, column: str) -> bool:
    """Obtiene un valor booleano de una celda."""
    val = get_cell_value(row, column)
    if val is None:
        return False
    # Manejar valores como 1.0, 0.0, True, False, etc.
    if val.lower() in ("1", "1.0", "true", "yes", "si"):
        return True
    return False


def escape_sql_string(value: str) -> str:
    """Escapa comillas simples para SQL."""
    if value is None:
        return None
    return value.replace("'", "''")


def generate_unique_hash(row_data: dict) -> str:
    """Genera un hash SHA256 único para el registro."""
    # Usar campos clave para generar el hash
    key_fields = [
        row_data.get("internal_code", ""),
        row_data.get("market_id", ""),
        row_data.get("product_id", ""),
        row_data.get("price_category_id", ""),
    ]
    hash_input = "|".join(str(f) for f in key_fields)
    return hashlib.sha256(hash_input.encode()).hexdigest()


def process_row(row, lookups: Dict, row_num: int) -> Tuple[Optional[dict], list]:
    """
    Procesa una fila del Excel y retorna los datos para el INSERT.
    Retorna: (datos_del_registro, lista_de_errores)
    """
    errors = []
    data = {}

    # Campos directos
    for excel_col, grades_col in EXCEL_TO_GRADES_MAPPING.items():
        if grades_col in ("is_sustainable", "is_iosco_assured", "is_public", "is_active", "is_spot"):
            val = get_cell_value(row, excel_col)
            if val is None:
                # Valores por defecto cuando la celda está vacía
                if grades_col == "is_active":
                    data[grades_col] = True
                elif grades_col == "is_public":
                    data[grades_col] = False
                else:
                    data[grades_col] = False
            else:
                data[grades_col] = get_bool_value(row, excel_col)
        else:
            data[grades_col] = get_cell_value(row, excel_col)

    # Campos FK
    for excel_col, (id_col, name_col) in FK_MAPPING.items():
        excel_value = get_cell_value(row, excel_col)

        if excel_value is None:
            data[id_col] = None
            if name_col:
                data[name_col] = None
            continue

        lookup = lookups.get(excel_col, {})
        lookup_result = lookup.get(excel_value.lower())

        if lookup_result is None:
            # Verificar si es un campo requerido
            required_fks = ["Price Category", "Market", "Product", "UOM", "Frequency"]
            if excel_col in required_fks:
                errors.append(f"{excel_col}: '{excel_value}' no encontrado")
            data[id_col] = None
            if name_col:
                data[name_col] = None
        else:
            uuid_val, original_name = lookup_result
            data[id_col] = uuid_val
            if name_col:
                data[name_col] = original_name

    # Frequency por defecto: 'Monthly' si está vacío
    if data.get("frequency_id") is None:
        frequency_lookup = lookups.get("Frequency", {})
        monthly_result = frequency_lookup.get("monthly")
        if monthly_result:
            uuid_val, original_name = monthly_result
            data["frequency_id"] = uuid_val
            data["frequency_name"] = original_name

    # Valores por defecto
    data["has_only_price_mid"] = False
    data["is_price_grade"] = True
    data["order"] = row_num

    return data, errors


def generate_insert_sql(records: list) -> str:
    """Genera el SQL con INSERTs múltiples agrupados por mercado."""
    if not records:
        return ""

    # Agrupar registros por mercado
    records_by_market = {}
    for record in records:
        market_name = record.get("market_name") or "Sin Mercado"
        if market_name not in records_by_market:
            records_by_market[market_name] = []
        records_by_market[market_name].append(record)

    lines = ["-- INSERT para tabla prices_assessment.grades"]
    lines.append(f"-- Total de registros: {len(records)}")
    lines.append(f"-- Agrupados por {len(records_by_market)} mercados\n")

    # Determinar columnas (usar el primer registro como referencia)
    # "order" es keyword en PostgreSQL, debe ir entre comillas dobles
    first_record = records[0]
    columns = ["id"] + [f'"{col}"' if col == "order" else col for col in first_record.keys()]

    # Generar INSERTs por mercado
    for market_name in sorted(records_by_market.keys()):
        market_records = records_by_market[market_name]

        lines.append(f"-- ============================================================")
        lines.append(f"-- Mercado: {market_name} ({len(market_records)} registros)")
        lines.append(f"-- ============================================================")
        lines.append(f"INSERT INTO prices_assessment.grades ({', '.join(columns)})")
        lines.append("VALUES")

        value_lines = []
        for record in market_records:
            values = ["gen_random_uuid()"]

            for col in first_record.keys():
                val = record.get(col)
                if val is None:
                    values.append("NULL")
                elif isinstance(val, bool):
                    values.append("TRUE" if val else "FALSE")
                elif col.endswith("_at") and val:  # Fechas
                    values.append(f"'{val}'::date")
                else:
                    escaped = escape_sql_string(str(val))
                    values.append(f"'{escaped}'")

            value_lines.append(f"    ({', '.join(values)})")

        lines.append(",\n".join(value_lines) + ";")
        lines.append("")

    return "\n".join(lines)


def main():
    # Leer archivo Excel
    excel_path = "docs/Prices 2.0 Data Structure.xlsx"
    print(f"Leyendo archivo: {excel_path}")
    df = pd.read_excel(excel_path, sheet_name="Sheet1")
    print(f"Total de filas en Excel: {len(df)}\n")

    # Conectar a la BD
    print("Conectando a la base de datos...")
    conn = get_db_connection()
    print("Conexión exitosa\n")

    # Cargar tablas de lookup
    print("Cargando tablas de lookup...")
    lookups = load_all_lookups(conn)
    print()

    # Procesar filas
    print("=" * 80)
    print("PROCESANDO FILAS")
    print("=" * 80)

    valid_records = []
    rows_with_errors = []
    skipped_with_id = 0
    skipped_bmi_price = 0
    skipped_no_code = 0
    skipped_duplicate_code = 0
    seen_codes = set()
    duplicate_codes_detail = []

    for idx, row in df.iterrows():
        row_num = idx + 1

        # Ignorar filas que ya tienen ID (ya existen en la BD)
        existing_id = get_cell_value(row, "ID")
        if existing_id:
            skipped_with_id += 1
            continue

        # Ignorar filas con Price Category = "BMI Price" (se actualizan, no se crean)
        price_category = get_cell_value(row, "Price Category")
        if price_category and price_category.lower() == "bmi price":
            skipped_bmi_price += 1
            continue

        # Ignorar filas sin código interno
        internal_code = get_cell_value(row, "Code")
        if not internal_code:
            skipped_no_code += 1
            continue

        # Ignorar filas con código interno duplicado (mantener primera ocurrencia)
        if internal_code in seen_codes:
            skipped_duplicate_code += 1
            market = get_cell_value(row, "Market")
            product = get_cell_value(row, "Product")
            duplicate_codes_detail.append((row_num, internal_code, market, product))
            continue
        seen_codes.add(internal_code)

        data, errors = process_row(row, lookups, row_num)

        if errors:
            rows_with_errors.append((row_num, data.get("internal_code"), errors))
        else:
            valid_records.append(data)

    conn.close()

    # Resumen
    print(f"\n⏭️  Registros omitidos (ya tienen ID): {skipped_with_id}")
    print(f"⏭️  Registros omitidos (BMI Price - para UPDATE): {skipped_bmi_price}")
    print(f"⏭️  Registros omitidos (sin código interno): {skipped_no_code}")
    print(f"⏭️  Registros omitidos (código duplicado): {skipped_duplicate_code}")
    print(f"✅ Registros válidos para INSERT: {len(valid_records)}")
    print(f"❌ Registros con errores: {len(rows_with_errors)}")

    if duplicate_codes_detail:
        print("\nCódigos duplicados omitidos (se mantuvo primera ocurrencia):")
        for row_num, code, market, product in duplicate_codes_detail[:10]:
            print(f"   Fila {row_num}: {code} ({market} / {product})")
        if len(duplicate_codes_detail) > 10:
            print(f"   ... y {len(duplicate_codes_detail) - 10} más")

    if rows_with_errors:
        print("\nPrimeros 10 registros con errores:")
        for row_num, code, errors in rows_with_errors[:10]:
            print(f"   Fila {row_num} ({code}): {', '.join(errors)}")
        if len(rows_with_errors) > 10:
            print(f"   ... y {len(rows_with_errors) - 10} más")

    # Generar archivo SQL
    if valid_records:
        SQLS_DIR.mkdir(exist_ok=True)
        sql_content = generate_insert_sql(valid_records)
        file_path = SQLS_DIR / "insert_grades.sql"
        file_path.write_text(sql_content)
        print(f"\n📁 Archivo SQL generado: {file_path}")

    return valid_records, rows_with_errors


if __name__ == "__main__":
    main()
