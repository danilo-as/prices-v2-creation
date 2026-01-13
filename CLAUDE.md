# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Script para validar datos del archivo Excel "Prices 2.0 Data Structure.xlsx" contra una base de datos PostgreSQL.

## Development Setup

```bash
# Crear entorno virtual e instalar dependencias
python3 -m venv venv
source venv/bin/activate
pip install pandas openpyxl psycopg2-binary python-dotenv
```

Configurar variables de entorno en `.env`:
```
DB_HOST=localhost
DB_PORT=5432
DB_NAME=<database_name>
DB_USER=<user>
DB_PASSWORD=<password>
```

## Scripts

### 1. Validación de Entidades (`validate_prices.py`)
Valida que los valores del Excel existan en las tablas de lookup de la BD.
```bash
source venv/bin/activate
python validate_prices.py
```
Genera archivos SQL en `sqls/` con INSERTs para valores faltantes.

### 2. Creación de Grades (`create_grades.py`)
Crea registros en la tabla `grades` a partir de cada fila del Excel.
```bash
source venv/bin/activate
python create_grades.py
```
Genera `sqls/insert_grades.sql` con los INSERTs.

## Column to Database Table Mapping

| Columna Excel                      | Tabla BD (prices_assessment.) | Campo BD |
|------------------------------------|-------------------------------|----------|
| Price Category                     | price_category                | name     |
| Product                            | product                       | name     |
| Sub-Product                        | sub_product                   | name     |
| Capacity (Nominal/Anode/Cathode)   | capacity                      | name     |
| Cell Format                        | cell_format                   | name     |
| Feedstock                          | feedstock                     | name     |
| Purity                             | purity                        | name     |
| Thickness                          | thickness                     | name     |
| Mesh size                          | mesh_size                     | size     |
| Service                            | service                       | name     |
| Incoterm                           | incoterm                      | code     |
| Region                             | region                        | name     |
| Country(OnlyforDiffs)              | country                       | name     |
| Trade Type?                        | trade_type                    | name     |
| Price Type                         | price_type                    | name     |
| UOM                                | unit_of_measure               | name     |

## Database Schema

### Esquema `prices_assessment`
- `grades` - tabla principal con referencias FK a las tablas de lookup
- `prices` - precios asociados a grades
- `prices_currency` - valores de precios por moneda
- Tablas de lookup: `price_category`, `product`, `sub_product`, `capacity`, `cell_format`, `feedstock`, `purity`, `thickness`, `mesh_size`, `service`, `incoterm`, `region`, `country`, `trade_type`, `price_type`, `unit_of_measure`

### Otros esquemas
- `primary_data.market` - markets (Aluminium, Anode, Cathode, etc.)
- `primary_data.currency` - monedas (USD, CNY, EUR, etc.)
- `calendar.frequencies` - frecuencias (Daily, Monthly, 2 Weeks, etc.)