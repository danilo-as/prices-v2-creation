#!/usr/bin/env python3
"""
Script para actualizar grades existentes con Price Category = "BMI Price".
Actualiza: internal_code, full_name, short_name
- Si tiene ID: usa el ID como filtro
- Si no tiene ID: usa las columnas con valor como filtro
"""

import pandas as pd
import psycopg2
from dotenv import load_dotenv
import os
from pathlib import Path

load_dotenv()

SQLS_DIR = Path("sqls")

# Mapeo de columnas del Excel a columnas de grades para el WHERE
FILTER_COLUMNS = {
    "Market": ("market_id", "primary_data.market", "name"),
    "Product": ("product_id", "prices_assessment.product", "name"),
    "Price Category": ("price_category_id", "prices_assessment.price_category", "name"),
    "Sub-Product": ("sub_product_id", "prices_assessment.sub_product", "name"),
    "Capacity (Nominal/Anode/Cathode)": ("capacity_id", "prices_assessment.capacity", "name"),
    "Cell Format": ("cell_format_id", "prices_assessment.cell_format", "name"),
    "Feedstock": ("feedstock_id", "prices_assessment.feedstock", "name"),
    "Purity": ("purity_id", "prices_assessment.purity", "name"),
    "Thickness": ("thickness_id", "prices_assessment.thickness", "name"),
    "Mesh size": ("mesh_size_id", "prices_assessment.mesh_size", "size"),
    "Service": ("service_id", "prices_assessment.service", "name"),
    "Incoterm": ("incoterm_id", "prices_assessment.incoterm", "name"),
    "Region": ("region_id", "prices_assessment.region", "name"),
    "Country(OnlyforDiffs)": ("country_id", "prices_assessment.country", "name"),
    "Trade Type?": ("trade_type_id", "prices_assessment.trade_type", "name"),
    "Price Type": ("price_type_id", "prices_assessment.price_type", "name"),
    "UOM": ("unit_of_measure_id", "prices_assessment.unit_of_measure", "name"),
    "Frequency": ("frequency_id", "calendar.frequencies", "name"),
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


def get_cell_value(row, column: str):
    """Obtiene el valor de una celda, retornando None si está vacío o NaN."""
    if column not in row.index:
        return None
    val = row[column]
    if pd.isna(val):
        return None
    return str(val).strip() if str(val).strip() else None


def escape_sql_string(value: str) -> str:
    """Escapa comillas simples para SQL."""
    if value is None:
        return None
    return value.replace("'", "''")


def normalize_for_comparison(value: str) -> str:
    """Normaliza un string para comparación."""
    import re
    normalized = value.lower().strip()
    normalized = re.sub(r'[-\s]+', ' ', normalized)
    return normalized


def load_lookup_table(conn, table: str, column: str) -> dict:
    """Carga una tabla de lookup y retorna un diccionario normalizado -> id."""
    cur = conn.cursor()
    cur.execute(f"SELECT id, {column} FROM {table} WHERE {column} IS NOT NULL")
    lookup = {}
    for row in cur.fetchall():
        uuid_val = str(row[0])
        name = row[1]
        if name:
            lookup[normalize_for_comparison(name)] = uuid_val
    cur.close()
    return lookup


def load_all_lookups(conn) -> dict:
    """Carga todas las tablas de lookup."""
    lookups = {}
    for excel_col, (grades_col, table, column) in FILTER_COLUMNS.items():
        lookups[excel_col] = load_lookup_table(conn, table, column)
    return lookups


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

    # Cargar lookups
    print("Cargando tablas de lookup...")
    lookups = load_all_lookups(conn)
    conn.close()
    print("Lookups cargados\n")

    # Filtrar registros con Price Category = "BMI Price"
    records_with_id = []
    records_without_id = []

    for idx, row in df.iterrows():
        price_category = get_cell_value(row, "Price Category")

        if not price_category or price_category.lower() != "bmi price":
            continue

        existing_id = get_cell_value(row, "ID")
        internal_code = get_cell_value(row, "Code")
        full_name = get_cell_value(row, "Full Name")
        short_name = get_cell_value(row, "Short Name (40 char max)")
        market = get_cell_value(row, "Market")

        if not internal_code:
            continue

        record = {
            "id": existing_id,
            "internal_code": internal_code,
            "full_name": full_name,
            "short_name": short_name,
            "market_name": market or "Sin Mercado",
            "filters": {}
        }

        # Construir filtros basados en columnas con valor
        if not existing_id:
            for excel_col, (grades_col, table, column) in FILTER_COLUMNS.items():
                excel_value = get_cell_value(row, excel_col)
                if excel_value:
                    lookup = lookups.get(excel_col, {})
                    normalized = normalize_for_comparison(excel_value)
                    lookup_id = lookup.get(normalized)
                    if lookup_id:
                        record["filters"][grades_col] = lookup_id
                elif grades_col == "sub_product_id":
                    # Si no hay sub_product, marcar como NULL para el WHERE
                    record["filters"][grades_col] = None

        if existing_id:
            records_with_id.append(record)
        else:
            records_without_id.append(record)

    total_records = len(records_with_id) + len(records_without_id)
    print(f"Registros con 'BMI Price' encontrados: {total_records}")
    print(f"  - Con ID (filtro por ID): {len(records_with_id)}")
    print(f"  - Sin ID (filtro por columnas): {len(records_without_id)}")

    if total_records == 0:
        print("No hay registros para actualizar.")
        return

    # Generar SQL
    lines = ["-- UPDATE grades with Price Category = 'BMI Price'"]
    lines.append(f"-- Total records: {total_records}")
    lines.append(f"-- Records with ID (filter by id): {len(records_with_id)}")
    lines.append(f"-- Records without ID (filter by columns): {len(records_without_id)}")
    lines.append("-- Updates: internal_code, full_name, short_name\n")

    # Agrupar por mercado
    all_records = records_with_id + records_without_id
    records_by_market = {}
    for record in all_records:
        market = record["market_name"]
        if market not in records_by_market:
            records_by_market[market] = []
        records_by_market[market].append(record)

    for market_name in sorted(records_by_market.keys()):
        market_records = records_by_market[market_name]

        lines.append(f"-- ============================================================")
        lines.append(f"-- Market: {market_name} ({len(market_records)} records)")
        lines.append(f"-- ============================================================")

        for record in market_records:
            # Construir SET clause
            set_parts = []

            escaped_code = escape_sql_string(record["internal_code"])
            set_parts.append(f"internal_code = '{escaped_code}'")

            if record["full_name"]:
                escaped_full_name = escape_sql_string(record["full_name"])
                set_parts.append(f"full_name = '{escaped_full_name}'")

            if record["short_name"]:
                escaped_short_name = escape_sql_string(record["short_name"])
                set_parts.append(f"short_name = '{escaped_short_name}'")

            set_clause = ", ".join(set_parts)

            # Construir WHERE clause
            if record["id"]:
                where_clause = f"id = '{record['id']}'"
            else:
                where_parts = []
                for col, val in record["filters"].items():
                    if val is None:
                        where_parts.append(f"{col} IS NULL")
                    else:
                        where_parts.append(f"{col} = '{val}'")

                if not where_parts:
                    lines.append(f"-- SKIPPED (no filters): {record['internal_code']}")
                    continue

                where_clause = " AND ".join(where_parts)

            lines.append(
                f"UPDATE prices_assessment.grades "
                f"SET {set_clause} "
                f"WHERE {where_clause};"
            )

        lines.append("")

    # Guardar archivo
    SQLS_DIR.mkdir(exist_ok=True)
    sql_content = "\n".join(lines)
    file_path = SQLS_DIR / "update_grades_internal_code.sql"
    file_path.write_text(sql_content)

    print(f"\n📁 SQL file generated: {file_path}")

    # Mostrar resumen por mercado
    print("\nSummary by market:")
    for market_name in sorted(records_by_market.keys()):
        print(f"  - {market_name}: {len(records_by_market[market_name])} records")


if __name__ == "__main__":
    main()
