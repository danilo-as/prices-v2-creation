#!/usr/bin/env python3
"""
Script para validar la cantidad de grades por Market y Price Category
comparando el archivo Excel con la base de datos.
"""

import pandas as pd
import psycopg2
from dotenv import load_dotenv
import os
from collections import defaultdict
from pathlib import Path
from datetime import datetime

load_dotenv()

REPORTS_DIR = Path("reports")


def get_db_connection():
    """Establece conexión con la base de datos PostgreSQL."""
    return psycopg2.connect(
        host=os.getenv("DB_HOST"),
        port=os.getenv("DB_PORT"),
        dbname=os.getenv("DB_NAME"),
        user=os.getenv("DB_USER"),
        password=os.getenv("DB_PASSWORD"),
    )


def normalize_name(name: str) -> str:
    """Normaliza un nombre para comparación."""
    if not name or pd.isna(name):
        return ""
    return str(name).strip().lower()


def get_excel_counts(excel_path: str) -> dict:
    """Lee el Excel y cuenta grades por Market y Price Category."""
    df = pd.read_excel(excel_path, sheet_name="Sheet1")

    counts = defaultdict(lambda: defaultdict(int))

    for _, row in df.iterrows():
        market = row.get("Market")
        price_category = row.get("Price Category")

        if pd.isna(market) or pd.isna(price_category):
            continue

        market = str(market).strip()
        price_category = str(price_category).strip()

        counts[market][price_category] += 1

    return counts


def get_db_counts(conn) -> dict:
    """Obtiene conteo de grades activos por Market y Price Category desde la BD."""
    cur = conn.cursor()
    cur.execute('''
        SELECT
            m.name as market,
            pc.name as price_category,
            COUNT(g.id) as count
        FROM prices_assessment.grades g
        JOIN primary_data.market m ON g.market_id = m.id
        JOIN prices_assessment.price_category pc ON g.price_category_id = pc.id
        GROUP BY m.name, pc.name
        ORDER BY m.name, pc.name
    ''')

    counts = defaultdict(lambda: defaultdict(int))
    for row in cur.fetchall():
        market, price_category, count = row
        counts[market][price_category] = count

    cur.close()
    return counts


def main():
    excel_path = "docs/Prices 2.0 Data Structure.xlsx"

    # Lista para almacenar las líneas del reporte
    lines = []

    lines.append("=" * 80)
    lines.append("VALIDACIÓN DE CANTIDAD DE GRADES POR MARKET Y PRICE CATEGORY")
    lines.append(f"Fecha: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append("=" * 80)

    # Leer Excel
    print(f"Leyendo archivo: {excel_path}")
    excel_counts = get_excel_counts(excel_path)

    # Conectar a BD
    print("Conectando a la base de datos...")
    conn = get_db_connection()
    db_counts = get_db_counts(conn)
    conn.close()

    # Obtener todos los markets y categorías únicos
    all_markets = sorted(set(list(excel_counts.keys()) + list(db_counts.keys())))
    all_categories = set()
    for market in all_markets:
        all_categories.update(excel_counts[market].keys())
        all_categories.update(db_counts[market].keys())
    all_categories = sorted(all_categories)

    # Comparar y mostrar resultados
    lines.append("")
    lines.append(f"{'MARKET':<30} {'PRICE CATEGORY':<20} {'EXCEL':>8} {'BD':>8} {'DIFF':>8} {'STATUS':<10}")
    lines.append("=" * 80)

    total_excel = 0
    total_db = 0
    discrepancies = []

    for market in all_markets:
        for category in all_categories:
            excel_count = excel_counts[market].get(category, 0)
            db_count = db_counts[market].get(category, 0)

            if excel_count == 0 and db_count == 0:
                continue

            total_excel += excel_count
            total_db += db_count

            diff = excel_count - db_count

            if diff == 0:
                status = "OK"
            elif diff > 0:
                status = "FALTAN"
                discrepancies.append((market, category, excel_count, db_count, diff))
            else:
                status = "SOBRAN"
                discrepancies.append((market, category, excel_count, db_count, diff))

            lines.append(f"{market:<30} {category:<20} {excel_count:>8} {db_count:>8} {diff:>+8} {status:<10}")

    lines.append("=" * 80)
    lines.append(f"{'TOTAL':<30} {'':<20} {total_excel:>8} {total_db:>8} {total_excel - total_db:>+8}")
    lines.append("=" * 80)

    # Resumen de discrepancias
    if discrepancies:
        lines.append("")
        lines.append("=" * 80)
        lines.append("RESUMEN DE DISCREPANCIAS")
        lines.append("=" * 80)

        missing = [d for d in discrepancies if d[4] > 0]
        extra = [d for d in discrepancies if d[4] < 0]

        if missing:
            lines.append(f"\nFALTAN en BD ({len(missing)} combinaciones):")
            for market, category, excel_count, db_count, diff in missing:
                lines.append(f"  - {market} | {category}: {diff} grades faltantes (Excel: {excel_count}, BD: {db_count})")

        if extra:
            lines.append(f"\nSOBRAN en BD ({len(extra)} combinaciones):")
            for market, category, excel_count, db_count, diff in extra:
                lines.append(f"  - {market} | {category}: {abs(diff)} grades extra (Excel: {excel_count}, BD: {db_count})")
    else:
        lines.append("")
        lines.append("Todos los conteos coinciden entre Excel y BD")

    # Mostrar en consola
    report_content = "\n".join(lines)
    print(report_content)

    # Guardar archivo
    REPORTS_DIR.mkdir(exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    file_path = REPORTS_DIR / f"validate_grades_count_{timestamp}.txt"
    file_path.write_text(report_content)

    print(f"\nArchivo generado: {file_path}")


if __name__ == "__main__":
    main()
