#!/usr/bin/env python3
"""
Script para validar duplicados en el archivo Excel antes de generar los INSERTs.
Detecta:
1. Códigos internos duplicados (internal_code)
2. Registros que violarían la constraint grade_unique_pk (combinación de campos FK)
"""

import pandas as pd
import psycopg2
from dotenv import load_dotenv
import os
from pathlib import Path
from collections import defaultdict

load_dotenv()

SQLS_DIR = Path("sqls")

# Columnas que componen el unique_sha256 (grade_unique_pk)
UNIQUE_COLUMNS = [
    "Frequency",
    "Market",
    "Price Category",
    "Product",
    "Capacity (Nominal/Anode/Cathode)",
    "Thickness",
    "Cell Format",
    "Country(OnlyforDiffs)",
    "Mesh size",
    "Purity",
    "Incoterm",
    "Region",
    "Service",
    "Sub-Product",
    "Feedstock",
    "Trade Type?",
    "Grade",
    "UOM",
    "Price Type",
    "Sustainable",
]


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


def get_bool_value(row, column: str) -> bool:
    """Obtiene un valor booleano de una celda."""
    val = get_cell_value(row, column)
    if val is None:
        return False
    if val.lower() in ("1", "1.0", "true", "yes", "si"):
        return True
    return False


def build_unique_key(row) -> tuple:
    """Construye una clave única basada en las columnas del grade_unique_pk."""
    key_parts = []
    for col in UNIQUE_COLUMNS:
        if col == "Sustainable":
            val = str(get_bool_value(row, col))
        else:
            val = get_cell_value(row, col) or ""
        key_parts.append(val.lower())
    return tuple(key_parts)


def main():
    # Leer archivo Excel
    excel_path = "docs/Prices 2.0 Data Structure.xlsx"
    print(f"Reading file: {excel_path}")
    df = pd.read_excel(excel_path, sheet_name="Sheet1")
    print(f"Total rows in Excel: {len(df)}\n")

    # Filtrar registros que serían procesados para INSERT
    # (sin ID existente y sin Price Category = "BMI Price")
    records_for_insert = []

    for idx, row in df.iterrows():
        row_num = idx + 2  # +2 porque Excel empieza en 1 y tiene header

        # Ignorar filas que ya tienen ID
        existing_id = get_cell_value(row, "ID")
        if existing_id:
            continue

        # Ignorar filas con Price Category = "BMI Price"
        price_category = get_cell_value(row, "Price Category")
        if price_category and price_category.lower() == "bmi price":
            continue

        internal_code = get_cell_value(row, "Code")
        full_name = get_cell_value(row, "Full Name")
        market = get_cell_value(row, "Market")
        product = get_cell_value(row, "Product")
        unique_key = build_unique_key(row)

        records_for_insert.append({
            "row_num": row_num,
            "internal_code": internal_code,
            "full_name": full_name,
            "market": market,
            "product": product,
            "unique_key": unique_key,
        })

    print(f"Records to be inserted: {len(records_for_insert)}\n")

    # 1. Detectar códigos internos duplicados
    print("=" * 80)
    print("1. DUPLICATE INTERNAL CODES")
    print("=" * 80)

    code_to_records = defaultdict(list)
    for record in records_for_insert:
        code = record["internal_code"]
        if code:
            code_to_records[code].append(record)

    duplicate_codes = {k: v for k, v in code_to_records.items() if len(v) > 1}

    if duplicate_codes:
        print(f"\nFound {len(duplicate_codes)} duplicate internal codes:\n")
        for code, records in sorted(duplicate_codes.items()):
            print(f"  Code: {code} ({len(records)} occurrences)")
            for rec in records:
                print(f"    - Row {rec['row_num']}: {rec['market']} / {rec['product']}")
    else:
        print("\nNo duplicate internal codes found.")

    # 2. Detectar registros que violarían grade_unique_pk
    print("\n" + "=" * 80)
    print("2. DUPLICATE UNIQUE KEYS (grade_unique_pk violation)")
    print("=" * 80)

    key_to_records = defaultdict(list)
    for record in records_for_insert:
        key_to_records[record["unique_key"]].append(record)

    duplicate_keys = {k: v for k, v in key_to_records.items() if len(v) > 1}

    if duplicate_keys:
        print(f"\nFound {len(duplicate_keys)} duplicate unique key groups:\n")
        for i, (key, records) in enumerate(sorted(duplicate_keys.items(), key=lambda x: x[1][0]["row_num"]), 1):
            print(f"  Group {i}: {len(records)} records")
            print(f"    Market: {records[0]['market']}")
            print(f"    Product: {records[0]['product']}")
            print(f"    Codes: {', '.join(r['internal_code'] or 'NULL' for r in records)}")
            print(f"    Rows: {', '.join(str(r['row_num']) for r in records)}")
            print()
    else:
        print("\nNo duplicate unique keys found.")

    # 3. Registros sin código interno
    print("=" * 80)
    print("3. RECORDS WITHOUT INTERNAL CODE")
    print("=" * 80)

    no_code_records = [r for r in records_for_insert if not r["internal_code"]]

    if no_code_records:
        print(f"\nFound {len(no_code_records)} records without internal code:\n")
        for rec in no_code_records[:20]:
            print(f"  - Row {rec['row_num']}: {rec['market']} / {rec['product']} / {rec['full_name'][:50] if rec['full_name'] else 'NULL'}...")
        if len(no_code_records) > 20:
            print(f"  ... and {len(no_code_records) - 20} more")
    else:
        print("\nAll records have internal codes.")

    # Generar reporte
    SQLS_DIR.mkdir(exist_ok=True)
    report_path = SQLS_DIR / "duplicates_validation_report.txt"

    with open(report_path, "w") as f:
        f.write("=" * 80 + "\n")
        f.write("DUPLICATES VALIDATION REPORT\n")
        f.write("=" * 80 + "\n\n")

        f.write(f"Total records for INSERT: {len(records_for_insert)}\n\n")

        # Duplicate codes
        f.write("-" * 80 + "\n")
        f.write("1. DUPLICATE INTERNAL CODES\n")
        f.write("-" * 80 + "\n\n")

        if duplicate_codes:
            f.write(f"Found {len(duplicate_codes)} duplicate internal codes:\n\n")
            for code, records in sorted(duplicate_codes.items()):
                f.write(f"Code: {code} ({len(records)} occurrences)\n")
                for rec in records:
                    f.write(f"  - Row {rec['row_num']}: {rec['market']} / {rec['product']}\n")
                    f.write(f"    Full Name: {rec['full_name'][:80] if rec['full_name'] else 'NULL'}...\n")
                f.write("\n")
        else:
            f.write("No duplicate internal codes found.\n")

        # Duplicate unique keys
        f.write("\n" + "-" * 80 + "\n")
        f.write("2. DUPLICATE UNIQUE KEYS (grade_unique_pk violation)\n")
        f.write("-" * 80 + "\n\n")

        if duplicate_keys:
            f.write(f"Found {len(duplicate_keys)} duplicate unique key groups:\n\n")
            for i, (key, records) in enumerate(sorted(duplicate_keys.items(), key=lambda x: x[1][0]["row_num"]), 1):
                f.write(f"Group {i}: {len(records)} records\n")
                f.write(f"  Market: {records[0]['market']}\n")
                f.write(f"  Product: {records[0]['product']}\n")
                f.write(f"  Codes: {', '.join(r['internal_code'] or 'NULL' for r in records)}\n")
                f.write(f"  Rows: {', '.join(str(r['row_num']) for r in records)}\n")
                f.write(f"  Full Name: {records[0]['full_name'][:80] if records[0]['full_name'] else 'NULL'}...\n")
                f.write("\n")
        else:
            f.write("No duplicate unique keys found.\n")

        # No code records
        f.write("\n" + "-" * 80 + "\n")
        f.write("3. RECORDS WITHOUT INTERNAL CODE\n")
        f.write("-" * 80 + "\n\n")

        if no_code_records:
            f.write(f"Found {len(no_code_records)} records without internal code:\n\n")
            for rec in no_code_records:
                f.write(f"Row {rec['row_num']}: {rec['market']} / {rec['product']}\n")
                f.write(f"  Full Name: {rec['full_name'][:80] if rec['full_name'] else 'NULL'}...\n")
        else:
            f.write("All records have internal codes.\n")

        # Summary
        f.write("\n" + "=" * 80 + "\n")
        f.write("SUMMARY\n")
        f.write("=" * 80 + "\n\n")
        f.write(f"Total records for INSERT: {len(records_for_insert)}\n")
        f.write(f"Duplicate internal codes: {len(duplicate_codes)} codes affecting {sum(len(v) for v in duplicate_codes.values())} records\n")
        f.write(f"Duplicate unique keys: {len(duplicate_keys)} groups affecting {sum(len(v) for v in duplicate_keys.values())} records\n")
        f.write(f"Records without code: {len(no_code_records)}\n")

    print(f"\nReport saved to: {report_path}")

    # Return counts for programmatic use
    return {
        "total_records": len(records_for_insert),
        "duplicate_codes": len(duplicate_codes),
        "duplicate_keys": len(duplicate_keys),
        "no_code_records": len(no_code_records),
    }


if __name__ == "__main__":
    main()
