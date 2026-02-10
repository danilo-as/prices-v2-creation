#!/usr/bin/env python3
"""
Script para generar sentencias SQL UPDATE que reordenan la columna "order"
en la tabla prices_assessment.price_definitions.

Grupos de ordenamiento (alfabético por internal_code dentro de cada grupo):
  1. Códigos que empiezan con BMIRE (pero NO BMIREPM)
  2. Códigos que empiezan con BMIREPM
  3. Todos los demás códigos
"""

import psycopg2
from dotenv import load_dotenv
import os
from pathlib import Path

load_dotenv()

SQLS_DIR = Path("sqls")


def get_db_connection():
    """Establece conexión con la base de datos PostgreSQL."""
    return psycopg2.connect(
        host=os.getenv("DB_HOST"),
        port=os.getenv("DB_PORT"),
        dbname=os.getenv("DB_NAME"),
        user=os.getenv("DB_USER"),
        password=os.getenv("DB_PASSWORD"),
    )


def fetch_price_definitions(conn):
    """Obtiene id e internal_code de todos los registros de price_definitions."""
    cur = conn.cursor()
    cur.execute(
        "SELECT id, internal_code FROM prices_assessment.price_definitions"
    )
    rows = cur.fetchall()
    cur.close()
    return rows


def classify_and_sort(rows):
    """
    Clasifica los registros en tres grupos y los ordena alfabéticamente
    por internal_code dentro de cada grupo.

    Grupo 1: BMIRE* (pero no BMIREPM*)
    Grupo 2: BMIREPM*
    Grupo 3: Todo lo demás
    """
    group_bmire = []
    group_bmirepm = []
    group_other = []

    for row_id, internal_code in rows:
        code = internal_code or ""
        if code.startswith("BMIREPM"):
            group_bmirepm.append((row_id, code))
        elif code.startswith("BMIRE"):
            group_bmire.append((row_id, code))
        else:
            group_other.append((row_id, code))

    # Ordenar cada grupo alfabéticamente por internal_code
    group_bmire.sort(key=lambda x: x[1])
    group_bmirepm.sort(key=lambda x: x[1])
    group_other.sort(key=lambda x: x[1])

    return group_bmire, group_bmirepm, group_other


def generate_update_sql(ordered_records):
    """Genera las sentencias SQL UPDATE para asignar el orden secuencial."""
    lines = ["-- UPDATE para reordenar columna \"order\" en prices_assessment.price_definitions"]
    lines.append(f"-- Total de registros: {len(ordered_records)}")
    lines.append("-- Orden: BMIRE (sin PM) -> BMIREPM -> otros (alfabético por internal_code)\n")

    for order_num, (record_id, internal_code) in enumerate(ordered_records, start=1):
        lines.append(
            f"UPDATE prices_assessment.price_definitions "
            f"SET \"order\" = {order_num} "
            f"WHERE id = '{record_id}'; "
            f"-- {internal_code}"
        )

    return "\n".join(lines)


def main():
    # Conectar a la BD
    print("Conectando a la base de datos...")
    conn = get_db_connection()
    print("Conexión exitosa\n")

    # Obtener registros
    print("Consultando registros de price_definitions...")
    rows = fetch_price_definitions(conn)
    conn.close()
    print(f"Total de registros obtenidos: {len(rows)}\n")

    if not rows:
        print("No se encontraron registros. Nada que hacer.")
        return

    # Clasificar y ordenar
    group_bmire, group_bmirepm, group_other = classify_and_sort(rows)

    # Concatenar en orden: Grupo 1 -> Grupo 2 -> Grupo 3
    ordered_records = group_bmire + group_bmirepm + group_other

    # Generar archivo SQL
    SQLS_DIR.mkdir(exist_ok=True)
    sql_content = generate_update_sql(ordered_records)
    file_path = SQLS_DIR / "update_order.sql"
    file_path.write_text(sql_content)

    # Resumen
    print("=" * 60)
    print("RESUMEN")
    print("=" * 60)
    print(f"   Total de registros:    {len(ordered_records)}")
    print(f"   Grupo BMIRE (sin PM):  {len(group_bmire)}")
    print(f"   Grupo BMIREPM:         {len(group_bmirepm)}")
    print(f"   Grupo otros:           {len(group_other)}")
    print(f"\n   Archivo generado: {file_path}")


if __name__ == "__main__":
    main()
