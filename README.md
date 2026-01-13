# Script de Validación de Precios

Script para validar los valores del archivo Excel "Prices 2.0 Data Structure.xlsx" contra la base de datos PostgreSQL y generar INSERTs SQL para valores faltantes.

## Requisitos

- Python 3.8+
- PostgreSQL

## Instalación

```bash
# Crear entorno virtual
python3 -m venv venv

# Activar entorno virtual
source venv/bin/activate

# Instalar dependencias
pip install pandas openpyxl psycopg2-binary python-dotenv
```

## Configuración

Crear archivo `.env` en la raíz del proyecto:

```
DB_HOST=localhost
DB_PORT=5432
DB_NAME=nombre_base_datos
DB_USER=usuario
DB_PASSWORD=contraseña
```

## Uso

```bash
source venv/bin/activate
python validate_prices.py
```

## Salida

El script:

1. Lee el archivo `docs/Prices 2.0 Data Structure.xlsx`
2. Compara cada columna del Excel contra su tabla correspondiente en PostgreSQL
3. Muestra un resumen de valores encontrados y faltantes
4. Genera archivos SQL con INSERTs en la carpeta `sqls/`

## Estructura de Archivos

```
.
├── docs/
│   └── Prices 2.0 Data Structure.xlsx
├── sqls/                    # Archivos SQL generados
│   ├── insert_product.sql
│   ├── insert_country.sql
│   └── ...
├── validate_prices.py
├── .env
└── README.md
```
