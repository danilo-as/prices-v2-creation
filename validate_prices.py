#!/usr/bin/env python3
"""
Script para validar los valores del archivo Excel 'Prices 2.0 Data Structure.xlsx'
contra los valores existentes en la base de datos PostgreSQL.
Genera archivos SQL con INSERTs para valores faltantes.
"""

import pandas as pd
import psycopg2
from dotenv import load_dotenv
import os
from typing import Set, Dict, List, Tuple
from pathlib import Path

load_dotenv()

SQLS_DIR = Path("sqls")

# Mapeo de columnas del Excel a tablas de la BD
# "columns" es una lista de columnas donde buscar el valor (OR logic)
COLUMN_TO_TABLE_MAPPING = {
    "Price Category": {"table": "price_category", "columns": ["name"]},
    "Product": {"table": "product", "columns": ["name", "chemical_code"]},
    "Sub-Product": {"table": "sub_product", "columns": ["name", "chemical_code"]},
    "Capacity (Nominal/Anode/Cathode)": {"table": "capacity", "columns": ["name"]},
    "Cell Format": {"table": "cell_format", "columns": ["name"]},
    "Feedstock": {"table": "feedstock", "columns": ["name"]},
    "Purity": {"table": "purity", "columns": ["name"]},
    "Thickness": {"table": "thickness", "columns": ["name"]},
    "Mesh size": {"table": "mesh_size", "columns": ["size"]},
    "Service": {"table": "service", "columns": ["name"]},
    "Incoterm": {"table": "incoterm", "columns": ["name", "code"]},
    "Region": {"table": "region", "columns": ["name"]},
    "Country(OnlyforDiffs)": {"table": "country", "columns": ["name"]},
    "Trade Type?": {"table": "trade_type", "columns": ["name"]},
    "Price Type": {"table": "price_type", "columns": ["name"]},
    "UOM": {"table": "unit_of_measure", "columns": ["name", "code"]},
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


def get_db_values(conn, table: str, columns: list) -> Set[str]:
    """Obtiene todos los valores únicos de múltiples columnas en una tabla (case insensitive)."""
    cur = conn.cursor()
    values = set()
    for column in columns:
        cur.execute(f"SELECT DISTINCT LOWER({column}) FROM prices_assessment.{table} WHERE {column} IS NOT NULL")
        values.update(row[0] for row in cur.fetchall())
    cur.close()
    return values


def get_excel_values(df: pd.DataFrame, column: str) -> Tuple[Set[str], Dict[str, str]]:
    """
    Obtiene todos los valores únicos de una columna del Excel.
    Retorna: (valores_normalizados_lower, mapeo_lower_a_original)
    """
    if column not in df.columns:
        return set(), {}
    values = df[column].dropna().unique()
    original_values = {str(v).strip() for v in values if str(v).strip()}
    # Mapeo de valor en minúsculas -> valor original
    lower_to_original = {v.lower(): v for v in original_values}
    return set(lower_to_original.keys()), lower_to_original


def validate_column(
    excel_values: Set[str], db_values: Set[str]
) -> Tuple[Set[str], Set[str]]:
    """
    Valida los valores del Excel contra la BD.
    Retorna: (valores_no_encontrados_en_bd, valores_solo_en_bd)
    """
    not_in_db = excel_values - db_values
    only_in_db = db_values - excel_values
    return not_in_db, only_in_db


def escape_sql_string(value: str) -> str:
    """Escapa comillas simples para SQL."""
    return value.replace("'", "''")


def generate_insert_sql(table: str, column: str, values: Set[str]) -> str:
    """Genera SQL con INSERT múltiple para una tabla."""
    if not values:
        return ""

    lines = [f"-- INSERT para tabla prices_assessment.{table}"]
    lines.append(f"-- Total de registros: {len(values)}\n")
    lines.append(f"INSERT INTO prices_assessment.{table} (id, {column}, created_at)")
    lines.append("VALUES")

    value_lines = []
    for val in sorted(values):
        escaped_val = escape_sql_string(val)
        value_lines.append(f"    (gen_random_uuid(), '{escaped_val}', NOW())")

    lines.append(",\n".join(value_lines) + ";")

    return "\n".join(lines)


def save_sql_files(results: Dict[str, Dict]) -> None:
    """Guarda los archivos SQL con INSERTs en la carpeta sqls."""
    SQLS_DIR.mkdir(exist_ok=True)

    generated_files = []

    for excel_col, data in results.items():
        if not data["not_in_db"]:
            continue

        table = data["table"]
        # Usar la primera columna como destino del INSERT
        column = data["db_columns"][0]
        values = data["not_in_db"]

        sql_content = generate_insert_sql(table, column, values)

        file_path = SQLS_DIR / f"insert_{table}.sql"
        file_path.write_text(sql_content)
        generated_files.append(file_path)

    if generated_files:
        print(f"\n📁 Archivos SQL generados en '{SQLS_DIR}/':")
        for f in generated_files:
            print(f"   - {f.name}")


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

    # Resultados de validación
    results: Dict[str, Dict] = {}

    print("=" * 80)
    print("RESULTADOS DE VALIDACIÓN")
    print("=" * 80)

    for excel_col, db_info in COLUMN_TO_TABLE_MAPPING.items():
        table = db_info["table"]
        db_columns = db_info["columns"]

        # Obtener valores (normalizados a minúsculas para comparación)
        excel_values_lower, lower_to_original = get_excel_values(df, excel_col)
        db_values_lower = get_db_values(conn, table, db_columns)

        # Validar (comparación case insensitive)
        not_in_db_lower, only_in_db = validate_column(excel_values_lower, db_values_lower)

        # Convertir valores faltantes a sus versiones originales para los INSERTs
        not_in_db_original = {lower_to_original[v] for v in not_in_db_lower}

        results[excel_col] = {
            "table": table,
            "db_columns": db_columns,
            "excel_count": len(excel_values_lower),
            "db_count": len(db_values_lower),
            "not_in_db": not_in_db_original,
            "only_in_db": only_in_db,
        }

        # Mostrar resultados
        print(f"\n📋 {excel_col}")
        cols_str = ", ".join(db_columns)
        print(f"   Tabla BD: prices_assessment.{table} ({cols_str})")
        print(f"   Valores en Excel: {len(excel_values_lower)} | Valores en BD: {len(db_values_lower)}")

        if not_in_db_original:
            print(f"   ❌ Valores del Excel NO encontrados en BD ({len(not_in_db_original)}):")
            for val in sorted(not_in_db_original)[:10]:  # Mostrar máximo 10
                print(f"      - {val}")
            if len(not_in_db_original) > 10:
                print(f"      ... y {len(not_in_db_original) - 10} más")
        else:
            print(f"   ✅ Todos los valores del Excel existen en la BD")

    conn.close()

    # Resumen final
    print("\n" + "=" * 80)
    print("RESUMEN")
    print("=" * 80)

    columns_with_issues = [col for col, data in results.items() if data["not_in_db"]]
    columns_ok = [col for col, data in results.items() if not data["not_in_db"]]

    print(f"\n✅ Columnas válidas: {len(columns_ok)}")
    for col in columns_ok:
        print(f"   - {col}")

    if columns_with_issues:
        print(f"\n❌ Columnas con valores no encontrados en BD: {len(columns_with_issues)}")
        for col in columns_with_issues:
            count = len(results[col]["not_in_db"])
            print(f"   - {col}: {count} valor(es) no encontrado(s)")

    # Generar archivos SQL con INSERTs
    save_sql_files(results)

    return results


if __name__ == "__main__":
    main()