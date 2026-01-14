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
# Configuración:
#   - db: "db1" (principal) o "db2" (warehouse) - indica dónde validar
#   - schema: esquema de la tabla
#   - table: nombre de la tabla
#   - columns: columnas donde buscar el valor (OR logic, case insensitive)
#   - source_table: (opcional) tabla fuente en db2 para obtener IDs (solo si db="db2")
#   - target_schema: (opcional) esquema destino en db1 para INSERT (solo si db="db2")
#   - target_table: (opcional) tabla destino en db1 para INSERT (solo si db="db2")

COLUMN_TO_TABLE_MAPPING = {
    # Entidades validadas contra BD2 (warehouse) - usan IDs de BD2
    "Market": {
        "db": "db2",
        "schema": "primary_data",
        "table": "benchmark_markets",
        "columns": ["name", "code"],
        "target_schema": "primary_data",
        "target_table": "market",
        "include_slug_db2": True,  # Incluir slug en INSERT de DB2
    },
    "Product": {
        "db": "db2",
        "schema": "primary_data",
        "table": "products",
        "columns": ["name", "code"],
        "target_schema": "prices_assessment",
        "target_table": "product",
        "target_columns": ["name", "chemical_code"],  # Columnas para validar en DB1
        "code_column": "code",  # Columna fuente en DB2 para chemical_code
        "include_slug_db2": True,  # Incluir slug en INSERT de DB2
    },
    # Entidades validadas contra BD1 (principal)
    "Price Category": {"db": "db1", "schema": "prices_assessment", "table": "price_category", "columns": ["name"]},
    "Sub-Product": {"db": "db1", "schema": "prices_assessment", "table": "sub_product", "columns": ["name", "chemical_code"]},
    "Capacity (Nominal/Anode/Cathode)": {"db": "db1", "schema": "prices_assessment", "table": "capacity", "columns": ["name"]},
    "Cell Format": {"db": "db1", "schema": "prices_assessment", "table": "cell_format", "columns": ["name"]},
    "Feedstock": {"db": "db1", "schema": "prices_assessment", "table": "feedstock", "columns": ["name"]},
    "Purity": {"db": "db1", "schema": "prices_assessment", "table": "purity", "columns": ["name"]},
    "Thickness": {"db": "db1", "schema": "prices_assessment", "table": "thickness", "columns": ["name"]},
    "Mesh size": {"db": "db1", "schema": "prices_assessment", "table": "mesh_size", "columns": ["size"]},
    "Service": {"db": "db1", "schema": "prices_assessment", "table": "service", "columns": ["name"]},
    "Incoterm": {"db": "db1", "schema": "prices_assessment", "table": "incoterm", "columns": ["name", "code"]},
    "Region": {"db": "db1", "schema": "prices_assessment", "table": "region", "columns": ["name"]},
    "Country(OnlyforDiffs)": {"db": "db1", "schema": "prices_assessment", "table": "country", "columns": ["name"]},
    "Trade Type?": {"db": "db1", "schema": "prices_assessment", "table": "trade_type", "columns": ["name"]},
    "Price Type": {"db": "db1", "schema": "prices_assessment", "table": "price_type", "columns": ["name"]},
    "UOM": {"db": "db1", "schema": "prices_assessment", "table": "unit_of_measure", "columns": ["name", "code"]},
}


def get_db_connection():
    """Establece conexión con la base de datos PostgreSQL principal."""
    return psycopg2.connect(
        host=os.getenv("DB_HOST"),
        port=os.getenv("DB_PORT"),
        dbname=os.getenv("DB_NAME"),
        user=os.getenv("DB_USER"),
        password=os.getenv("DB_PASSWORD"),
    )


def get_db2_connection():
    """Establece conexión con la segunda base de datos (warehouse)."""
    return psycopg2.connect(
        host=os.getenv("DB2_HOST"),
        port=os.getenv("DB2_PORT"),
        dbname=os.getenv("DB2_NAME"),
        user=os.getenv("DB2_USER"),
        password=os.getenv("DB2_PASSWORD"),
    )


def get_db_values(conn, table: str, columns: list, schema: str = "prices_assessment") -> Set[str]:
    """Obtiene todos los valores únicos de múltiples columnas en una tabla (normalizados para comparación)."""
    cur = conn.cursor()
    values = set()
    for column in columns:
        cur.execute(f"SELECT DISTINCT {column} FROM {schema}.{table} WHERE {column} IS NOT NULL")
        values.update(normalize_for_comparison(row[0]) for row in cur.fetchall() if row[0])
    cur.close()
    return values


def get_db2_values_with_ids(conn, schema: str, table: str, columns: list, code_column: str = None) -> Tuple[Set[str], Dict[str, Tuple]]:
    """
    Obtiene valores de una tabla en BD2 con sus IDs.
    Si code_column está especificado, también captura ese valor.
    Retorna: (valores_normalizados, mapeo_normalizado_a_(id, nombre_original, code_o_None))
    """
    cur = conn.cursor()

    values_normalized = set()
    normalized_to_id_name = {}

    # Determinar si necesitamos el code
    if code_column:
        # Obtener name y code juntos
        cur.execute(f"SELECT id, name, {code_column} FROM {schema}.{table} WHERE name IS NOT NULL")
        for row in cur.fetchall():
            uuid_val = str(row[0])
            original_name = row[1]
            code_val = row[2]
            if original_name:
                val_normalized = normalize_for_comparison(original_name)
                values_normalized.add(val_normalized)
                if val_normalized not in normalized_to_id_name:
                    normalized_to_id_name[val_normalized] = (uuid_val, original_name, code_val)
        # También buscar por code
        cur.execute(f"SELECT id, name, {code_column} FROM {schema}.{table} WHERE {code_column} IS NOT NULL")
        for row in cur.fetchall():
            uuid_val = str(row[0])
            original_name = row[1]
            code_val = row[2]
            if code_val:
                val_normalized = normalize_for_comparison(code_val)
                values_normalized.add(val_normalized)
                if val_normalized not in normalized_to_id_name:
                    normalized_to_id_name[val_normalized] = (uuid_val, original_name, code_val)
    else:
        for column in columns:
            cur.execute(f"SELECT id, {column} FROM {schema}.{table} WHERE {column} IS NOT NULL")
            for row in cur.fetchall():
                uuid_val = str(row[0])
                original_value = row[1]
                if original_value:
                    val_normalized = normalize_for_comparison(original_value)
                    values_normalized.add(val_normalized)
                    if val_normalized not in normalized_to_id_name:
                        normalized_to_id_name[val_normalized] = (uuid_val, original_value, None)

    cur.close()
    return values_normalized, normalized_to_id_name


def get_excel_values(df: pd.DataFrame, column: str) -> Tuple[Set[str], Dict[str, str]]:
    """
    Obtiene todos los valores únicos de una columna del Excel.
    Retorna: (valores_normalizados, mapeo_normalizado_a_original)
    """
    if column not in df.columns:
        return set(), {}
    values = df[column].dropna().unique()
    original_values = {str(v).strip() for v in values if str(v).strip()}
    # Mapeo de valor normalizado -> valor original
    normalized_to_original = {normalize_for_comparison(v): v for v in original_values}
    return set(normalized_to_original.keys()), normalized_to_original


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


def slugify(value: str) -> str:
    """Convierte un string a slug (minúsculas, guiones, sin caracteres especiales)."""
    import re
    slug = value.lower().strip()
    slug = re.sub(r'[&]', '-and-', slug)
    slug = re.sub(r'[^a-z0-9\s-]', '', slug)
    slug = re.sub(r'[\s_]+', '-', slug)
    slug = re.sub(r'-+', '-', slug)
    slug = slug.strip('-')
    return slug


def normalize_for_comparison(value: str) -> str:
    """Normaliza un string para comparación (elimina guiones y espacios extras)."""
    import re
    normalized = value.lower().strip()
    normalized = re.sub(r'[-\s]+', ' ', normalized)  # Guiones y espacios -> espacio único
    return normalized


def generate_insert_sql(table: str, column: str, values: Set[str], schema: str = "prices_assessment") -> str:
    """Genera SQL con INSERT múltiple para una tabla."""
    if not values:
        return ""

    lines = [f"-- INSERT para tabla {schema}.{table}"]
    lines.append(f"-- Total de registros: {len(values)}\n")
    lines.append(f"INSERT INTO {schema}.{table} (id, {column}, created_at)")
    lines.append("VALUES")

    value_lines = []
    for val in sorted(values):
        escaped_val = escape_sql_string(val)
        value_lines.append(f"    (gen_random_uuid(), '{escaped_val}', NOW())")

    lines.append(",\n".join(value_lines) + ";")

    return "\n".join(lines)


def generate_db2_insert_sql(target_schema: str, target_table: str, source_table: str,
                            values_with_ids: Dict[str, Tuple], include_chemical_code: bool = False) -> str:
    """
    Genera SQL con INSERT múltiple usando IDs de BD2.
    values_with_ids: {valor_lower: (id_de_bd2, nombre_original, code_o_None)}
    include_chemical_code: si True, incluye columna chemical_code en el INSERT
    """
    if not values_with_ids:
        return ""

    lines = [f"-- INSERT para tabla {target_schema}.{target_table}"]
    lines.append(f"-- Total de registros: {len(values_with_ids)}")
    lines.append(f"-- IDs tomados de {source_table} (DB2)\n")

    if include_chemical_code:
        lines.append(f"INSERT INTO {target_schema}.{target_table} (id, name, chemical_code, created_at)")
    else:
        lines.append(f"INSERT INTO {target_schema}.{target_table} (id, name, created_at)")
    lines.append("VALUES")

    value_lines = []
    for val_lower in sorted(values_with_ids.keys()):
        uuid_val, original_name, code_val = values_with_ids[val_lower]
        escaped_name = escape_sql_string(original_name)
        if include_chemical_code:
            if code_val:
                escaped_code = escape_sql_string(code_val)
                value_lines.append(f"    ('{uuid_val}', '{escaped_name}', '{escaped_code}', NOW())")
            else:
                value_lines.append(f"    ('{uuid_val}', '{escaped_name}', NULL, NOW())")
        else:
            value_lines.append(f"    ('{uuid_val}', '{escaped_name}', NOW())")

    lines.append(",\n".join(value_lines) + ";")

    return "\n".join(lines)


def generate_new_inserts_both_dbs(source_schema: str, source_table: str,
                                   target_schema: str, target_table: str,
                                   values: Set[str], include_chemical_code: bool = False,
                                   include_slug_db2: bool = False) -> Tuple[str, str]:
    """
    Genera SQL con INSERT múltiple para AMBAS BDs cuando el valor no existe en ninguna.
    Usa el mismo UUID para ambas tablas.
    include_chemical_code: si True, incluye columna chemical_code (NULL) en DB1
    include_slug_db2: si True, incluye columna slug en INSERT de DB2
    Retorna: (sql_para_db2, sql_para_db1)
    """
    if not values:
        return "", ""

    import uuid

    lines_db2 = [f"-- INSERT para tabla {source_schema}.{source_table} (DB2 - Source)"]
    lines_db2.append(f"-- Total de registros: {len(values)}")
    lines_db2.append("-- Nuevos valores que no existen en DB2\n")
    if include_slug_db2:
        lines_db2.append(f"INSERT INTO {source_schema}.{source_table} (id, name, slug, created_at)")
    else:
        lines_db2.append(f"INSERT INTO {source_schema}.{source_table} (id, name, created_at)")
    lines_db2.append("VALUES")

    lines_db1 = [f"-- INSERT para tabla {target_schema}.{target_table} (DB1 - Target)"]
    lines_db1.append(f"-- Total de registros: {len(values)}")
    lines_db1.append("-- Nuevos valores (usar mismo ID que DB2)\n")
    if include_chemical_code:
        lines_db1.append(f"INSERT INTO {target_schema}.{target_table} (id, name, chemical_code, created_at)")
    else:
        lines_db1.append(f"INSERT INTO {target_schema}.{target_table} (id, name, created_at)")
    lines_db1.append("VALUES")

    value_lines_db2 = []
    value_lines_db1 = []

    for val in sorted(values):
        new_uuid = str(uuid.uuid4())
        escaped_val = escape_sql_string(val)
        if include_slug_db2:
            slug_val = slugify(val)
            value_lines_db2.append(f"    ('{new_uuid}', '{escaped_val}', '{slug_val}', NOW())")
        else:
            value_lines_db2.append(f"    ('{new_uuid}', '{escaped_val}', NOW())")
        if include_chemical_code:
            value_lines_db1.append(f"    ('{new_uuid}', '{escaped_val}', NULL, NOW())")
        else:
            value_lines_db1.append(f"    ('{new_uuid}', '{escaped_val}', NOW())")

    lines_db2.append(",\n".join(value_lines_db2) + ";")
    lines_db1.append(",\n".join(value_lines_db1) + ";")

    return "\n".join(lines_db2), "\n".join(lines_db1)


def save_sql_files(results: Dict[str, Dict], db2_data: Dict[str, Dict] = None) -> None:
    """
    Guarda los archivos SQL con INSERTs en la carpeta sqls.
    db2_data: {excel_col: {"config": {...}, "inserts": {...}, "not_found": set()}}
    """
    SQLS_DIR.mkdir(exist_ok=True)

    generated_files = []

    # Generar SQL para entidades de DB2
    if db2_data:
        for excel_col, data in db2_data.items():
            config = data["config"]
            target_schema = config["target_schema"]
            target_table = config["target_table"]
            source_schema = config["schema"]
            source_table = config["table"]
            # Determinar opciones de INSERT
            include_chemical_code = "code_column" in config
            include_slug_db2 = config.get("include_slug_db2", False)

            # 1. Valores que existen en DB2 → INSERT solo a DB1 con ID de DB2
            if data.get("inserts"):
                sql_content = generate_db2_insert_sql(
                    target_schema, target_table,
                    f"{source_schema}.{source_table}",
                    data["inserts"],
                    include_chemical_code=include_chemical_code
                )
                file_path = SQLS_DIR / f"insert_{target_table}_from_db2.sql"
                file_path.write_text(sql_content)
                generated_files.append(file_path)

            # 2. Valores que NO existen en DB2 → INSERT a AMBAS BDs
            if data.get("not_found"):
                sql_db2, sql_db1 = generate_new_inserts_both_dbs(
                    source_schema, source_table,
                    target_schema, target_table,
                    data["not_found"],
                    include_chemical_code=include_chemical_code,
                    include_slug_db2=include_slug_db2
                )
                if sql_db2:
                    file_path_db2 = SQLS_DIR / f"insert_{source_table}_new_db2.sql"
                    file_path_db2.write_text(sql_db2)
                    generated_files.append(file_path_db2)

                if sql_db1:
                    file_path_db1 = SQLS_DIR / f"insert_{target_table}_new_db1.sql"
                    file_path_db1.write_text(sql_db1)
                    generated_files.append(file_path_db1)

    # Generar SQL para entidades de DB1 (las que no usan DB2)
    for excel_col, data in results.items():
        if not data["not_in_db"]:
            continue

        table = data["table"]
        schema = data.get("schema", "prices_assessment")
        column = data["db_columns"][0]
        values = data["not_in_db"]

        sql_content = generate_insert_sql(table, column, values, schema)

        file_path = SQLS_DIR / f"insert_{table}.sql"
        file_path.write_text(sql_content)
        generated_files.append(file_path)

    if generated_files:
        print(f"\n📁 Archivos SQL generados en '{SQLS_DIR}/':")
        for f in sorted(generated_files, key=lambda x: x.name):
            print(f"   - {f.name}")


def main():
    # Leer archivo Excel
    excel_path = "docs/Prices 2.0 Data Structure.xlsx"
    print(f"Leyendo archivo: {excel_path}")
    df = pd.read_excel(excel_path, sheet_name="Sheet1")
    print(f"Total de filas en Excel: {len(df)}\n")

    # Conectar a la BD principal
    print("Conectando a la base de datos principal (DB1)...")
    conn1 = get_db_connection()
    print("Conexión exitosa")

    # Conectar a la BD secundaria (warehouse)
    print("Conectando a la base de datos secundaria (DB2 - warehouse)...")
    conn2 = get_db2_connection()
    print("Conexión exitosa\n")

    # Resultados de validación
    results_db1: Dict[str, Dict] = {}  # Entidades validadas contra DB1
    db2_inserts: Dict[str, Dict] = {}  # Entidades de DB2 con INSERTs pendientes
    db2_not_found: Dict[str, set] = {}  # Valores no encontrados en DB2

    print("=" * 80)
    print("RESULTADOS DE VALIDACIÓN")
    print("=" * 80)

    for excel_col, config in COLUMN_TO_TABLE_MAPPING.items():
        db = config["db"]
        schema = config["schema"]
        table = config["table"]
        columns = config["columns"]

        # Obtener valores del Excel
        excel_values_lower, lower_to_original = get_excel_values(df, excel_col)

        if db == "db2":
            # Validar contra DB2 (warehouse), generar INSERTs con IDs de DB2
            target_schema = config["target_schema"]
            target_table = config["target_table"]
            target_columns = config.get("target_columns", ["name"])
            code_column = config.get("code_column")

            # Obtener valores de DB2 con sus IDs (y code si aplica)
            db2_values_lower, db2_lower_to_id_name = get_db2_values_with_ids(
                conn2, schema, table, columns, code_column=code_column
            )

            # Obtener valores actuales en DB1 (tabla destino) - usar target_columns
            current_db1_values = get_db_values(conn1, target_table, target_columns, schema=target_schema)

            # Encontrar valores del Excel que no están en DB1
            not_in_db1_lower = excel_values_lower - current_db1_values

            # De esos, encontrar cuáles SÍ están en DB2
            can_insert_lower = not_in_db1_lower & db2_values_lower
            cannot_insert_lower = not_in_db1_lower - db2_values_lower

            # Preparar INSERTs con IDs de DB2 pero nombres del Excel
            inserts = {}
            for v in can_insert_lower:
                uuid_val, db2_name, code_val = db2_lower_to_id_name[v]
                excel_name = lower_to_original.get(v, db2_name)  # Usar nombre del Excel
                inserts[v] = (uuid_val, excel_name, code_val)
            not_found = {lower_to_original.get(v, v) for v in cannot_insert_lower}
            db2_inserts[excel_col] = {"config": config, "inserts": inserts, "not_found": not_found}

            # Mostrar resultados
            print(f"\n📋 {excel_col}")
            print(f"   Validación: DB2 ({schema}.{table}) -> DB1 ({target_schema}.{target_table})")
            print(f"   Valores en Excel: {len(excel_values_lower)} | En DB1: {len(current_db1_values)} | En DB2: {len(db2_values_lower)}")

            if can_insert_lower:
                print(f"   🔄 Para insertar (encontrados en DB2): {len(can_insert_lower)}")
                for val_lower in sorted(can_insert_lower)[:5]:
                    uuid_val, name, code = db2_lower_to_id_name[val_lower]
                    print(f"      - {name} (ID: {uuid_val[:8]}...)")
                if len(can_insert_lower) > 5:
                    print(f"      ... y {len(can_insert_lower) - 5} más")

            if cannot_insert_lower:
                print(f"   🆕 Nuevos (INSERT a DB2 y DB1): {len(cannot_insert_lower)}")
                for val_lower in sorted(cannot_insert_lower)[:5]:
                    print(f"      - {lower_to_original.get(val_lower, val_lower)}")
                if len(cannot_insert_lower) > 5:
                    print(f"      ... y {len(cannot_insert_lower) - 5} más")

            if not not_in_db1_lower:
                print(f"   ✅ Todos los valores del Excel existen en DB1")

        else:
            # Validar contra DB1 (principal)
            db_values_lower = get_db_values(conn1, table, columns, schema=schema)

            # Validar
            not_in_db_lower, only_in_db = validate_column(excel_values_lower, db_values_lower)
            not_in_db_original = {lower_to_original[v] for v in not_in_db_lower}

            results_db1[excel_col] = {
                "table": table,
                "schema": schema,
                "db_columns": columns,
                "excel_count": len(excel_values_lower),
                "db_count": len(db_values_lower),
                "not_in_db": not_in_db_original,
                "only_in_db": only_in_db,
            }

            # Mostrar resultados
            print(f"\n📋 {excel_col}")
            cols_str = ", ".join(columns)
            print(f"   Tabla DB1: {schema}.{table} ({cols_str})")
            print(f"   Valores en Excel: {len(excel_values_lower)} | Valores en DB1: {len(db_values_lower)}")

            if not_in_db_original:
                print(f"   ❌ NO encontrados en DB1 ({len(not_in_db_original)}):")
                for val in sorted(not_in_db_original)[:10]:
                    print(f"      - {val}")
                if len(not_in_db_original) > 10:
                    print(f"      ... y {len(not_in_db_original) - 10} más")
            else:
                print(f"   ✅ Todos los valores del Excel existen en DB1")

    conn1.close()
    conn2.close()

    # Resumen final
    print("\n" + "=" * 80)
    print("RESUMEN")
    print("=" * 80)

    # Resumen DB2
    db2_with_inserts = [col for col, data in db2_inserts.items() if data["inserts"]]
    db2_with_errors = [col for col, vals in db2_not_found.items() if vals]

    if db2_with_inserts:
        print(f"\n🔄 Entidades DB2 con valores para insertar:")
        for col in db2_with_inserts:
            count = len(db2_inserts[col]["inserts"])
            print(f"   - {col}: {count} valor(es)")

    if db2_with_errors:
        print(f"\n❌ Entidades DB2 con valores NO encontrados:")
        for col in db2_with_errors:
            count = len(db2_not_found[col])
            print(f"   - {col}: {count} valor(es)")

    # Resumen DB1
    db1_ok = [col for col, data in results_db1.items() if not data["not_in_db"]]
    db1_with_issues = [col for col, data in results_db1.items() if data["not_in_db"]]

    print(f"\n✅ Entidades DB1 válidas: {len(db1_ok)}")
    for col in db1_ok:
        print(f"   - {col}")

    if db1_with_issues:
        print(f"\n❌ Entidades DB1 con valores no encontrados:")
        for col in db1_with_issues:
            count = len(results_db1[col]["not_in_db"])
            print(f"   - {col}: {count} valor(es)")

    # Generar archivos SQL
    save_sql_files(results_db1, db2_inserts)

    return results_db1, db2_inserts


if __name__ == "__main__":
    main()