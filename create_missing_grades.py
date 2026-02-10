#!/usr/bin/env python3
"""
Script para crear registros de grades faltantes.
Compara los códigos del Excel 'Prices 2.0 Data Structure OLD.xlsx' contra los que ya existen
en el archivo 'ids_and_internal_codes_in_DB.csv' y genera INSERTs para los faltantes.
"""

import uuid
import json
import pandas as pd
import psycopg2
from dotenv import load_dotenv
import os
from typing import Dict, Optional, Tuple, Set
from pathlib import Path

load_dotenv()

SQLS_DIR = Path("sqls")

# Mapeo de columnas del Excel a tablas de lookup
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
    "Grade": "grade_name",
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
    "Default Currency": ("default_currency_id", None),
}


CONFIG_FK_MAPPING = {
    "product_id": ("productId", "productName", "product_name"),
    "sub_product_id": ("subProductId", "subProductName", "sub_product_name"),
    "capacity_id": ("capacityId", "capacityName", "capacity_name"),
    "cell_format_id": ("cellFormatId", "cellFormatName", "cell_format_name"),
    "feedstock_id": ("feedstockId", "feedstockName", "feedstock_name"),
    "purity_id": ("purityId", "purityName", "purity_name"),
    "thickness_id": ("thicknessId", "thicknessName", "thickness_name"),
    "mesh_size_id": ("meshSizeId", "meshSizeName", "mesh_size_name"),
    "service_id": ("serviceId", "serviceName", "service_name"),
    "incoterm_id": ("incotermId", "incotermName", "incoterm_name"),
    "region_id": ("regionId", "regionName", "region_name"),
    "country_id": ("countryId", "countryName", "country_name"),
    "trade_type_id": ("tradeTypeId", "tradeTypeName", "trade_type_name"),
    "price_type_id": ("priceTypeId", "priceTypeName", "price_type_name"),
    "unit_of_measure_id": ("unitOfMeasureId", "unitOfMeasureName", "unit_of_measure_name"),
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


def load_existing_codes_from_csv(csv_path: str) -> Set[str]:
    """Carga los códigos internos que ya existen en la BD desde el archivo CSV."""
    df = pd.read_csv(csv_path)
    existing_codes = set(df["internal_code"].dropna().str.strip())
    return existing_codes


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


def load_product_chemical_codes(conn) -> Dict[str, Optional[str]]:
    """Loads product_id -> chemical_code mapping."""
    cur = conn.cursor()
    cur.execute("SELECT id, chemical_code FROM prices_assessment.product WHERE chemical_code IS NOT NULL")
    mapping = {}
    for row in cur.fetchall():
        mapping[str(row[0])] = row[1]
    cur.close()
    return mapping


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
    if val.lower() in ("1", "1.0", "true", "yes", "si"):
        return True
    return False


def escape_sql_string(value: str) -> str:
    """Escapa comillas simples para SQL."""
    if value is None:
        return None
    return value.replace("'", "''")


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


def build_dw_configuration(record, product_chemical_codes):
    config = {}
    for record_key, (id_key, name_key, record_name_field) in CONFIG_FK_MAPPING.items():
        config[id_key] = record.get(record_key)
        config[name_key] = record.get(record_name_field)
    config["productAlias"] = None
    config["productChemicalCode"] = product_chemical_codes.get(record.get("product_id"))
    config["grade"] = record.get("grade_name")
    return config


def generate_insert_sql(records: list, uuids: list) -> str:
    """Genera el SQL con INSERTs múltiples agrupados por mercado."""
    if not records:
        return ""

    # Crear mapeo record -> uuid
    record_uuid_map = {}
    for i, record in enumerate(records):
        record_uuid_map[id(record)] = uuids[i]

    # Agrupar registros por mercado
    records_by_market = {}
    for record in records:
        market_name = record.get("market_name") or "Sin Mercado"
        if market_name not in records_by_market:
            records_by_market[market_name] = []
        records_by_market[market_name].append(record)

    lines = ["-- INSERT para tabla prices_assessment.price_definitions (GRADES FALTANTES)"]
    lines.append(f"-- Total de registros: {len(records)}")
    lines.append(f"-- Agrupados por {len(records_by_market)} mercados")
    lines.append("-- Estos registros existen en el Excel pero NO en la BD\n")

    # Determinar columnas
    first_record = records[0]
    columns = ["id"] + [f'"{col}"' if col == "order" else col for col in first_record.keys()]

    # Generar INSERTs por mercado
    for market_name in sorted(records_by_market.keys()):
        market_records = records_by_market[market_name]

        lines.append(f"-- ============================================================")
        lines.append(f"-- Mercado: {market_name} ({len(market_records)} registros)")
        lines.append(f"-- ============================================================")
        lines.append(f"INSERT INTO prices_assessment.price_definitions ({', '.join(columns)})")
        lines.append("VALUES")

        value_lines = []
        for record in market_records:
            record_uuid = record_uuid_map[id(record)]
            values = [f"'{record_uuid}'"]

            for col in first_record.keys():
                val = record.get(col)
                if val is None:
                    values.append("NULL")
                elif isinstance(val, bool):
                    values.append("TRUE" if val else "FALSE")
                elif col.endswith("_at") and val:
                    values.append(f"'{val}'::date")
                else:
                    escaped = escape_sql_string(str(val))
                    values.append(f"'{escaped}'")

            value_lines.append(f"    ({', '.join(values)})")

        lines.append(",\n".join(value_lines) + ";")
        lines.append("")

    return "\n".join(lines)


def generate_dw_insert_sql(records: list, uuids: list, product_chemical_codes: dict) -> str:
    """Genera el SQL con INSERTs para prices.grades (datawarehouse)."""
    if not records:
        return ""

    # Crear mapeo record -> uuid
    record_uuid_map = {}
    for i, record in enumerate(records):
        record_uuid_map[id(record)] = uuids[i]

    # Agrupar registros por mercado
    records_by_market = {}
    for record in records:
        market_name = record.get("market_name") or "Sin Mercado"
        if market_name not in records_by_market:
            records_by_market[market_name] = []
        records_by_market[market_name].append(record)

    lines = ["-- INSERT para tabla prices.grades (DATAWAREHOUSE)"]
    lines.append(f"-- Total de registros: {len(records)}")
    lines.append(f"-- Agrupados por {len(records_by_market)} mercados")
    lines.append("-- Estos registros existen en el Excel pero NO en la BD\n")

    dw_columns = [
        "id", "name", "market_id", "product_id", "configuration",
        "is_spot", "is_sustainable", "frequency_id", "frequency_name",
        "is_active", "is_iosco_assured", "is_public", "default_currency_id",
        "product_name", "price_category_id", "price_category_name",
        "internal_code", "short_name", "can_be_deleted"
    ]

    for market_name in sorted(records_by_market.keys()):
        market_records = records_by_market[market_name]

        lines.append(f"-- ============================================================")
        lines.append(f"-- Mercado: {market_name} ({len(market_records)} registros)")
        lines.append(f"-- ============================================================")
        lines.append(f"INSERT INTO prices.grades ({', '.join(dw_columns)})")
        lines.append("VALUES")

        value_lines = []
        for record in market_records:
            record_uuid = record_uuid_map[id(record)]
            config = build_dw_configuration(record, product_chemical_codes)
            config_json = json.dumps(config, ensure_ascii=False)
            config_json_escaped = config_json.replace("'", "''")

            values = []
            for col in dw_columns:
                if col == "id":
                    values.append(f"'{record_uuid}'")
                elif col == "name":
                    val = record.get("full_name")
                    values.append(f"'{escape_sql_string(str(val))}'" if val is not None else "NULL")
                elif col == "configuration":
                    values.append(f"'{config_json_escaped}'::jsonb")
                elif col == "can_be_deleted":
                    values.append("FALSE")
                elif col in ("is_spot", "is_sustainable", "is_active", "is_iosco_assured", "is_public"):
                    val = record.get(col)
                    if val is None:
                        values.append("FALSE")
                    else:
                        values.append("TRUE" if val else "FALSE")
                elif col in ("default_currency_id", "internal_code", "short_name"):
                    val = record.get(col)
                    if val is None:
                        values.append("NULL")
                    else:
                        values.append(f"'{escape_sql_string(str(val))}'")
                else:
                    val = record.get(col)
                    if val is None:
                        values.append("NULL")
                    else:
                        values.append(f"'{escape_sql_string(str(val))}'")

            value_lines.append(f"    ({', '.join(values)})")

        lines.append(",\n".join(value_lines) + ";")
        lines.append("")

    return "\n".join(lines)


def main():
    # Rutas de archivos
    excel_path = "docs/REPM grades.xlsx"
    csv_path = "docs/ids_and_internal_codes_in_DB.csv"

    # Cargar códigos existentes del CSV
    print(f"Leyendo códigos existentes de: {csv_path}")
    existing_codes = load_existing_codes_from_csv(csv_path)
    print(f"Total de códigos en BD (CSV): {len(existing_codes)}\n")

    # Leer archivo Excel
    print(f"Leyendo archivo Excel: {excel_path}")
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
    print("BUSCANDO GRADES FALTANTES")
    print("=" * 80)

    valid_records = []
    rows_with_errors = []
    skipped_exists_in_db = 0
    skipped_no_code = 0
    skipped_duplicate_code = 0
    seen_codes = set()
    duplicate_codes_detail = []

    for idx, row in df.iterrows():
        row_num = idx + 1

        # Obtener código interno
        internal_code = get_cell_value(row, "Code")

        # Ignorar filas sin código interno
        if not internal_code:
            skipped_no_code += 1
            continue

        # Ignorar si el código ya existe en la BD (está en el CSV)
        if internal_code in existing_codes:
            skipped_exists_in_db += 1
            continue

        # Ignorar filas con código interno duplicado en el Excel
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

    # Cargar chemical codes antes de cerrar la conexión
    product_chemical_codes = load_product_chemical_codes(conn)

    conn.close()

    # Resumen
    print(f"\nRESUMEN:")
    print(f"   Registros ya existentes en BD (omitidos): {skipped_exists_in_db}")
    print(f"   Registros sin código (omitidos): {skipped_no_code}")
    print(f"   Códigos duplicados en Excel (omitidos): {skipped_duplicate_code}")
    print(f"   Registros FALTANTES válidos para INSERT: {len(valid_records)}")
    print(f"   Registros con errores: {len(rows_with_errors)}")

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
        uuids = [str(uuid.uuid4()) for _ in valid_records]
        sql_primary = generate_insert_sql(valid_records, uuids)
        sql_dw = generate_dw_insert_sql(valid_records, uuids, product_chemical_codes)
        sql_content = sql_primary + "\n\n" + sql_dw
        file_path = SQLS_DIR / "insert_missing_grades.sql"
        file_path.write_text(sql_content)
        print(f"\nArchivo SQL generado: {file_path}")
    else:
        print("\nNo hay grades faltantes para insertar.")

    return valid_records, rows_with_errors


if __name__ == "__main__":
    main()
