# Excel Generation Backend Module

Módulo Python para generar archivos Excel desde PostgreSQL y cargar datos desde Excel a PostgreSQL usando FastAPI, SQLAlchemy e IronXL.

## Características

- 🚀 **API REST con FastAPI** - Endpoints RESTful para generación y carga de Excel
- 🗄️ **Integración PostgreSQL** - Conexión asíncrona con SQLAlchemy
- 📊 **Generación Excel con IronXL** - Soporte para grandes volúmenes de datos
- 📤 **Carga desde Excel** - Importación de datos desde archivos Excel a BD
- 🎨 **Sistema de Templates** - Templates reutilizables para formateo
- ⚡ **Alto Rendimiento** - Manejo eficiente de hasta 200,000 filas y 50 columnas
- 🛡️ **Manejo de Errores** - Sistema robusto de validación y errores
- 🔄 **Validación de Columnas** - Mapeo automático entre Excel y BD
- 📝 **Logging Completo** - Sistema de logging configurable

## Instalación

1. **Clonar el repositorio:**

```bash
git clone https://github.com/Dazzlm/excel_generation.git
cd excel_generation
```

2. **Crear entorno virtual:**

```bash
python -m venv venv
source venv/bin/activate  # En Windows: venv\Scripts\activate
```

3. **Instalar dependencias:**

```bash
pip install -r requirements.txt
```

4. **Configurar variables de entorno:**

```bash
cp .env.example .env
# Editar .env con tu configuración de base de datos
```

## Configuración

### Variables de Entorno

Crear archivo `.env` basado en `.env.example`:


### Base de Datos

Asegúrate de tener PostgreSQL ejecutándose y una base de datos configurada con las tablas que deseas exportar/importar.

## Uso

### Iniciar el Servidor

```bash
python main.py
```

La API estará disponible en: `http://localhost:8000`

## Endpoints Disponibles

### 1. Generar Excel desde BD

**POST** `/api/v1/reports/excel`

#### Request Body:

```json
{
  "table": "nombre_tabla",
  "fields": ["campo1", "campo2", "campo3"],
  "filters": {
    "año": 2024,
    "estado": "activo"
  },
  "template": "default_template.xlsx"
}
```

#### Response:

Archivo Excel descargable con header `Content-Disposition: attachment`

### 2. Cargar Excel a BD

**POST** `/api/v1/upload/excel`

#### Form Data:

- `file`: Archivo Excel (.xlsx o .xls)
- `table`: Nombre de la tabla destino
- `update_existing`: (opcional) Si actualizar registros existentes
- `batch_size`: (opcional) Tamaño del lote (default: 1000)

#### Response:

```json
{
  "success": true,
  "message": "Procesadas 150 de 150 filas",
  "rows_processed": 150,
  "errors": null
}
```

### 3. Obtener Información de Tabla

**GET** `/api/v1/tables/{table_name}/columns`

#### Response:

```json
{
  "table": "usuarios",
  "columns": {
    "id": {
      "type": "integer",
      "nullable": false,
      "default": "nextval('usuarios_id_seq'::regclass)"
    },
    "nombre": {
      "type": "character varying",
      "nullable": false,
      "default": null
    }
  }
}
```

## Ejemplos de Uso

### Generar Excel con cURL

```bash
curl -X POST "http://localhost:8000/api/v1/reports/excel" \
  -H "Content-Type: application/json" \
  -d '{
    "table": "usuarios",
    "fields": ["id", "nombre", "email", "fecha_registro"],
    "filters": {"activo": true}
  }' \
  --output usuarios_export.xlsx
```

### Cargar Excel con cURL

```bash
curl -X POST "http://localhost:8000/api/v1/upload/excel" \
  -F "file=@datos.xlsx" \
  -F "table=usuarios" \
  -F "update_existing=false" \
  -F "batch_size=500"
```

### Ejemplo con Python

```python
import requests

# Generar Excel
response = requests.post(
    "http://localhost:8000/api/v1/reports/excel",
    json={
        "table": "productos",
        "fields": ["id", "nombre", "precio", "categoria"],
        "filters": {"categoria": "electronica"}
    }
)

with open("productos.xlsx", "wb") as f:
    f.write(response.content)

# Cargar Excel
with open("nuevos_productos.xlsx", "rb") as f:
    files = {"file": f}
    data = {
        "table": "productos",
        "update_existing": "false",
        "batch_size": "1000"
    }

    response = requests.post(
        "http://localhost:8000/api/v1/upload/excel",
        files=files,
        data=data
    )

    print(response.json())
```

## Validación de Columnas

### Mapeo Automático

El sistema automáticamente mapea columnas entre Excel y la base de datos:

- **Case-insensitive**: `Nombre` en Excel → `nombre` en BD
- **Validación de requeridas**: Detecta columnas obligatorias faltantes
- **Sugerencias**: Propone columnas similares para errores

### Tipos de Datos Soportados

- **Enteros**: `integer`, `bigint`, `smallint`
- **Decimales**: `numeric`, `decimal`, `real`, `double precision`
- **Booleanos**: `boolean` (acepta: true/false, 1/0, yes/no)
- **Fechas**: `date`, `timestamp`, `timestamptz`
- **Texto**: `text`, `varchar`, `character varying`

### Ejemplo de Validación

```json
{
  "success": false,
  "message": "Error de validación de columnas",
  "rows_processed": 0,
  "errors": [
    "Columna 'Email' del Excel no existe en tabla 'usuarios'",
    "Columna requerida 'nombre' no encontrada en Excel"
  ]
}
```

## Templates

### Ubicación

Los templates se almacenan en `excel_generation/templates/`

### Crear Template Personalizado

1. Crear archivo Excel con formato deseado
2. Guardar como `.xlsx` en la carpeta `templates/`
3. Usar el nombre del archivo en el parámetro `template`

### Template por Defecto

Se incluye `default_template.xlsx` con formato básico de encabezados.

## Estructura del Proyecto

```
excel_generation/
├── controllers/
│   └── excel_controller.py    # Endpoints REST API
├── services/
│   ├── generator_service.py   # Lógica de generación Excel
│   └── upload_service.py      # Lógica de carga desde Excel
├── templates/
│   └── default_template.xlsx  # Template por defecto
├── utils/
│   ├── db.py                 # Manejo de base de datos
│   ├── error_handling.py     # Manejo de errores
│   └── logger.py             # Sistema de logging
├── __init__.py
└── README.md
```

## Características Técnicas

### Rendimiento

- ✅ Soporte para 200,000 filas y 50 columnas
- ✅ Procesamiento por chunks para memoria eficiente
- ✅ Conexiones de DB asíncronas
- ✅ Stream processing para archivos grandes
- ✅ Procesamiento por lotes configurable

### Validaciones

- ✅ Validación de campos contra esquema de DB
- ✅ Validación de existencia de tablas
- ✅ Validación de parámetros de entrada
- ✅ Manejo de errores de conexión
- ✅ Mapeo automático de columnas case-insensitive
- ✅ Conversión automática de tipos de datos

### Seguridad

- ✅ Consultas parametrizadas (prevención SQL injection)
- ✅ Límites de filas configurable
- ✅ Validación de entrada
- ✅ Logging de actividad
- ✅ Validación de tipos de archivo

## API Documentation

Una vez iniciado el servidor, la documentación interactiva está disponible en:

- **Swagger UI**: `http://localhost:8000/docs`
- **ReDoc**: `http://localhost:8000/redoc`

## Troubleshooting

### Error de Conexión a DB

```
DatabaseError: Error accediendo a tabla
```

**Solución**: Verificar configuración de `.env` y que PostgreSQL esté ejecutándose.

### Error de Campos Inválidos

```
ExcelGenerationError: Campos inválidos para tabla
```

**Solución**: Verificar que todos los campos en `fields` existan en la tabla.

### Error de Columnas en Carga

```
ValidationError: Columna 'Email' del Excel no existe en tabla 'usuarios'
```

**Solución**: Revisar nombres de columnas en Excel vs BD, o usar endpoint `/tables/{table}/columns` para ver estructura.

### Error de Memoria

```
MemoryError: Unable to allocate array
```

**Solución**: Reducir el número de filas o usar `batch_size` menor.

## Contribución

1. Fork del repositorio
2. Crear branch para feature (`git checkout -b feature/nueva-funcionalidad`)
3. Commit cambios (`git commit -am 'Agregar nueva funcionalidad'`)
4. Push al branch (`git push origin feature/nueva-funcionalidad`)
5. Crear Pull Request

## Licencia

Este proyecto está bajo la Licencia MIT. Ver `LICENSE` para más detalles.

## Soporte

Para reportar bugs o solicitar features, crear un issue en:
https://github.com/Dazzlm/excel_generation/issues
