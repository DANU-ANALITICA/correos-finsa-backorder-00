# Proyecto Automatización de Envío Backorder a Proveedores Finsa

## Requisitos

### Ejecución local

Para ejecutar el proyecto de forma local, es necesario contar con:

* Un archivo `credenciales.json` correspondiente a una **Cuenta de Servicio de Google Cloud**.

    * La cuenta de servicio debe tener los permisos IAM necesarios para consultar y/o modificar información en BigQuery.

### Configuración de correo electrónico

Se requiere una **Contraseña de Aplicación de Google** para la cuenta desde la que se enviarán los correos.

#### 1. Habilitar la verificación en dos pasos

La cuenta de Google debe tener activada la verificación en dos pasos:

1. Acceder a **Google Account → Seguridad**.

2. Habilitar **Verificación en dos pasos**.

URL: https://myaccount.google.com

#### 2. Generar una contraseña de aplicación

1. Acceder a la sección **Contraseñas de aplicaciones**.

2. Crear una nueva contraseña utilizando un nombre identificable (por ejemplo: *"Cloud Run correos"*).

3. Guardar la contraseña generada en un lugar seguro.

> **Importante:** Google muestra esta contraseña una sola vez. Si se pierde, será necesario generar una nueva.

URL: https://myaccount.google.com/apppasswords

## Instalación

```bash
pip install -r requirements.txt
```

## Variables de entorno

Crear un archivo `.env`:

```env
GCP_PROJECT = "project_id"
GOOGLE_APPLICATION_CREDENTIALS = "/Ruta/a/Archivo/Json/credencials.json"  #solo si corres local
BUYERS_PASSWORD='{ id_buyer_int : "app_password"}'
```

## Ejecución

```bash
python main.py
```

## Descripción de main.py

Código creado para el envío por correo de Backorder atrasado a Proveedores para su control. 

Con la intención de correr dentro de GCP.

### Fuentes

Toma información de BigQuery

- `finsadashboard.raw_data.Proveedores`: Datos de proveedores
- `finsadashboard.mrts.mrts_backorder_MTY`: Datos de Compras | Backoder

Query Mails:

```sql
SELECT DISTINCT Proveedor, Nombre,
COALESCE(
  NULLIF(Email1, ''),
  NULLIF(Email2, ''),
  NULLIF(Email3, '')
) AS Email FROM `finsadashboard.raw_data.Proveedores`
WHERE COALESCE(
  NULLIF(Email1, ''),
  NULLIF(Email2, ''),
  NULLIF(Email3, '')
)  is not NULL;
```

### Librerías

| Librería | Descripción |
|-----------|-------------|
| `os` | Permite interactuar con el sistema operativo (archivos, variables de entorno, rutas, etc.). |
| `re` | Implementa expresiones regulares para búsqueda y manipulación de texto. |
| `smtplib` | Facilita el envío de correos electrónicos mediante SMTP. |
| `email.mime` | Creación y envío de correos electrónicos. |
| `python-dotenv` | Carga variables de entorno desde archivos `.env`. |
| `pandas` | Herramienta para análisis y transformación de datos mediante DataFrames. |
| `google-cloud-bigquery` | Cliente oficial para consultar y administrar datos en BigQuery desde Python. |

### Proceso

**1. Instalación de dependencias**

Instalar las librerías necesarias utilizando `requirements.txt` desde consola

**2. Importación de librerías**

Se cargan las librerías estándar de Python, dependencias externas y módulos de Google Cloud necesarios para el procesamiento, conexión a BigQuery y envío de correos.

**3. Extraer variables de entorno del .env**

Se cargan variables de entorno desde un archivo `.env`, incluyendo credenciales de Google Cloud y del correo emisor.

**4. Conexión a BQ**

Se inicializa un cliente de BigQuery utilizando las credenciales del proyecto para poder ejecutar consultas.

**5. Cargar información de destinatarios (proveedores)**

Se ejecuta una consulta para obtener los proveedores y sus correos asociados, los cuales serán usados como destinatarios del reporte.

**6. Validar correos destinatarios con regex**

Se validan los correos electrónicos utilizando expresiones regulares para asegurar que el formato sea correcto antes del envío.

**7. Función base para extracción de datos de compra - backorder**

Para cada proveedor se ejecuta una consulta parametrizada en BigQuery que obtiene información de backorders (productos pendientes o retrasados).

Toma como parámetro el Nombre del Proveedor y devuelve un DataFrame de pandas

**8. Función base para creación y envío de correo**

Se construye un correo en formato HTML y se envía mediante SMTP (Gmail), utilizando autenticación con contraseña de aplicación.

Toma como parámtros el DataFrame de pandas creado por la función anterior y el Mail del Proveedor. Regresa mensaje con resultado de la operación.

**9. Inicio de bucle proveedor por proveedor**

El script recorre cada proveedor, obtiene su información, genera el reporte y lo envía por correo. Si ocurre algún error en BigQuery o en el proceso, se captura y se muestra en consola.

