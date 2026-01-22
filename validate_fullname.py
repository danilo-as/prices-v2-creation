#!/usr/bin/env python3
"""
Script para validar que el full_name de cada registro en prices_updated.json
contenga los valores de sus propiedades correspondientes.
"""

import json
from pathlib import Path
from datetime import datetime

REPORTS_DIR = Path("reports")

# Campos a validar en el full_name
FIELDS_TO_VALIDATE = [
    ("market_name", "Market"),
    ("price_category_name", "Price Category"),
    ("product_name", "Product"),
    ("sub_product_name", "Sub-Product"),
    ("capacity_name", "Capacity"),
    ("purity_name", "Purity"),
    ("thickness_name", "Thickness"),
    ("cell_format_name", "Cell Format"),
    ("mesh_size_name", "Mesh Size"),
    ("feedstock_name", "Feedstock"),
    ("service_name", "Service"),
    ("region_name", "Region"),
    ("country_name", "Country"),
    ("trade_type_name", "Trade Type"),
    ("price_type_name", "Price Type"),
    ("unit_of_measure_name", "UOM"),
]

# Mapeo de incoterms: nombre completo -> código
INCOTERM_CODES = {
    "Delivered Duty Paid": "DDP",
    "Cost, Insurance and Freight": "CIF",
    "Free On Board": "FOB",
    "Ex Works": "EXW",
    "Delivered At Place": "DAP",
    "Free Carrier": "FCA",
    "Cost and Freight": "CFR",
    "Ex Works to Cost, Insurance and Freight": "EXW to CIF"
}

# Mapeo de sub_product: nombre completo -> código químico
SUB_PRODUCT_CODES = {
    "Cobalt Metal": "Co",
    "Cobalt Sulphate": "CoSO4",
    "Nickel Metal": "Ni",
    "Nickel Sulphate": "NiSO4",
    "Lithium Carbonate Equivalent": "LCE",
}

# Mapeo de regiones: nombre en BD -> variantes en full_name
REGION_ALIASES = {
    "N America": ["North America", "N America"],
    "S America": ["South America", "S America"],
    "Asia": ["Asia", "China"],  # China es parte de Asia
}

# Mapeo de UOM: nombre en BD -> variantes en full_name
UOM_ALIASES = {
    "Percentage": ["%", "Percentage"],
}


def normalize_for_search(text: str) -> str:
    """Normaliza texto para búsqueda (minúsculas, sin espacios extra)."""
    if not text:
        return ""
    return text.lower().strip()


def value_in_fullname(value: str, full_name: str, field_name: str = None) -> bool:
    """Verifica si un valor está presente en el full_name."""
    if not value:
        return True

    normalized_value = normalize_for_search(value)
    normalized_fullname = normalize_for_search(full_name)

    # Búsqueda directa
    if normalized_value in normalized_fullname:
        return True

    # Para incoterms, buscar también el código
    incoterm_code = INCOTERM_CODES.get(value)
    if incoterm_code:
        if normalize_for_search(incoterm_code) in normalized_fullname:
            return True

    # Para sub_product, buscar también el código químico
    if field_name == "sub_product_name":
        sub_product_code = SUB_PRODUCT_CODES.get(value)
        if sub_product_code:
            if normalize_for_search(sub_product_code) in normalized_fullname:
                return True

    # Para regiones, buscar alias
    if field_name == "region_name":
        aliases = REGION_ALIASES.get(value, [])
        for alias in aliases:
            if normalize_for_search(alias) in normalized_fullname:
                return True

    # Para UOM, buscar alias
    if field_name == "unit_of_measure_name":
        aliases = UOM_ALIASES.get(value, [])
        for alias in aliases:
            if normalize_for_search(alias) in normalized_fullname:
                return True

    return False


def validate_record(record: dict) -> list:
    """Valida un registro y retorna lista de errores encontrados."""
    errors = []
    full_name = record.get("full_name", "")

    if not full_name:
        errors.append("full_name está vacío")
        return errors

    for field_name, field_label in FIELDS_TO_VALIDATE:
        field_value = record.get(field_name)

        if field_value and not value_in_fullname(field_value, full_name, field_name):
            errors.append(f"{field_label}: '{field_value}' no encontrado en full_name")

    # Validar incoterm por separado (caso especial)
    incoterm_name = record.get("incoterm_name")
    if incoterm_name:
        incoterm_code = INCOTERM_CODES.get(incoterm_name, incoterm_name)
        if not value_in_fullname(incoterm_name, full_name) and not value_in_fullname(incoterm_code, full_name):
            errors.append(f"Incoterm: '{incoterm_name}' (código: {incoterm_code}) no encontrado en full_name")

    return errors


def main():
    json_path = Path("docs/prices_updated.json")

    print(f"Leyendo archivo: {json_path}")
    with open(json_path, "r", encoding="utf-8") as f:
        all_records = json.load(f)

    # Filtrar solo registros activos
    records = [r for r in all_records if r.get("is_active") is True]

    print(f"Total de registros: {len(all_records)}")
    print(f"Registros activos: {len(records)}")
    print(f"Registros inactivos (ignorados): {len(all_records) - len(records)}")

    # Validar cada registro activo
    records_with_errors = []

    for record in records:
        errors = validate_record(record)
        if errors:
            records_with_errors.append({
                "id": record.get("id"),
                "internal_code": record.get("internal_code"),
                "full_name": record.get("full_name"),
                "errors": errors
            })

    # Generar reporte
    lines = []
    lines.append("=" * 100)
    lines.append("VALIDACIÓN DE FULL_NAME EN PRICES_UPDATED.JSON")
    lines.append(f"Fecha: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append("=" * 100)
    lines.append("")
    lines.append(f"Total de registros en archivo: {len(all_records)}")
    lines.append(f"Registros activos analizados: {len(records)}")
    lines.append(f"Registros inactivos (ignorados): {len(all_records) - len(records)}")
    lines.append(f"Registros con errores: {len(records_with_errors)}")
    lines.append(f"Registros sin errores: {len(records) - len(records_with_errors)}")
    lines.append("")

    if records_with_errors:
        lines.append("=" * 100)
        lines.append("DETALLE DE ERRORES")
        lines.append("=" * 100)

        for i, record in enumerate(records_with_errors, 1):
            lines.append("")
            lines.append(f"--- Registro {i} ---")
            lines.append(f"ID: {record['id']}")
            lines.append(f"Internal Code: {record['internal_code']}")
            lines.append(f"Full Name: {record['full_name']}")
            lines.append("Errores:")
            for error in record["errors"]:
                lines.append(f"  - {error}")

        # Resumen por tipo de error
        lines.append("")
        lines.append("=" * 100)
        lines.append("RESUMEN POR TIPO DE ERROR")
        lines.append("=" * 100)

        error_counts = {}
        for record in records_with_errors:
            for error in record["errors"]:
                # Extraer el campo del error
                field = error.split(":")[0] if ":" in error else error
                error_counts[field] = error_counts.get(field, 0) + 1

        for field, count in sorted(error_counts.items(), key=lambda x: -x[1]):
            lines.append(f"  {field}: {count} errores")

    else:
        lines.append("No se encontraron errores. Todos los registros son válidos.")

    # Mostrar en consola
    report_content = "\n".join(lines)
    print(report_content)

    # Guardar archivo
    REPORTS_DIR.mkdir(exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    file_path = REPORTS_DIR / f"validate_fullname_{timestamp}.txt"
    file_path.write_text(report_content, encoding="utf-8")

    print(f"\nArchivo generado: {file_path}")


if __name__ == "__main__":
    main()
