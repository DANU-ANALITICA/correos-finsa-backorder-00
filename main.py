# main.py

## librerías

import os
import smtplib

from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

from dotenv import load_dotenv

import pandas as pd
import re
import numpy as np
from google.cloud import bigquery
from email_validator import validate_email, EmailNotValidError

## Extrae variables del .env
load_dotenv()
PROJECT_ID = os.environ["GCP_PROJECT"]
EMISOR = os.environ.get("EMISOR", "")
PASSWORD = os.environ.get("PASSWORD", "")

##  En GCP se usa client = bigquery.Client(project=PROJECT_ID)
client = bigquery.Client(project=os.environ["GCP_PROJECT"])

# # query = '''SELECT DISTINCT Proveedor, Nombre,
# # COALESCE(
# #   NULLIF(Email1, ''),
# #   NULLIF(Email2, ''),
# #   NULLIF(Email3, '')
# # ) AS Email FROM `finsadashboard.raw_data.Proveedores`
# # WHERE 
# # COALESCE(
# #   NULLIF(Email1, ''),
# #   NULLIF(Email2, ''),
# #   NULLIF(Email3, '')
# # ) LIKE '%@%' 

# # --LIMIT 100
# # '''

# Query de prueba para envío a destinatario fijo
query = '''SELECT DISTINCT Proveedor, Nombre,
'leonardo.laureles@danuanalitica.com' AS Email 
--CASE WHEN Proveedor = 10096 THEN 'fmartinez@finsa.com.mx' ELSE 'ruben.garza@finsa.com.mx' END AS Email
FROM `finsadashboard.raw_data.Proveedores` 
--LIMIT 3
WHERE Proveedor IN (10096,10099)
'''

correos = client.query(query).to_dataframe()


query = """
        SELECT 
            EDP,
            FOLIO AS ORDEN_COMPRA,
            PARTIDA,
            FECHA_ALTA,
            ARTICULO,
            DESCRIPCION,
            CANTIDAD,
            CANTIDAD_RECIBIDA,
            BACKORDER,
            UNIDAD,
            COSTO,
            FECHA_EMBARQUE,
            DIAS_RETRASO_EMBARQUE,
            NOMBRE_SUCURSAL,
            NOMBRE_COMPRADOR,
            NOMBRE_PROVEEDOR,
            PROVEEDOR,
            COMPRADOR,
            SUCURSAL
        FROM `finsadashboard.mrts.mrts_backorder_MTY`
        WHERE BACKORDER > 0 
          AND DIAS_RETRASO_EMBARQUE > 0
          AND  PROVEEDOR IN (10096,10099)
        ORDER BY NOMBRE_PROVEEDOR, FECHA_ALTA
    """

backorder = client.query(query).to_dataframe()
backorder

loop_values = backorder[['NOMBRE_COMPRADOR','NOMBRE_PROVEEDOR', 'NOMBRE_SUCURSAL', 'NOMBRE_COMPRADOR', 'PROVEEDOR', 'COMPRADOR', 'SUCURSAL']].drop_duplicates()
loop_values = loop_values.merge(correos, left_on='PROVEEDOR', right_on='Proveedor', how='left')
# ─── DEF CORREOS_CLEAN ─────────────────────────────────────────────────────────────────

def clean_email_addresses(correos: pd.DataFrame) -> pd.DataFrame:
    extraido = correos['Email'].str.extract(r'\[(.*?)\]', expand=False)
    correos['Clean'] = extraido.fillna(correos['Email'])
    extraido = correos['Clean'].str.extract(r'<(.*?)>', expand=False)
    correos['Clean'] = (
    extraido
    .fillna(correos['Clean'])
    .str.replace(r'[\[\]<>]', '', regex=True)  # eliminar caracteres especiales
    .str.rsplit(' ', n=1).str[-1]              # quedarse con la última palabra
    .str.rsplit(':', n=1).str[-1]              # eliminar prefijo tipo "mailto:"
)
    patron = r'([a-zA-Z0-9Ññ._-]+@[a-zA-Z0-9Ññ_-]+(?:\.[a-zA-Z]{2,})+)'    
    correos['Clean'] = correos['Clean'].str.extract(patron)
    correos = correos.dropna(subset=['Clean'])
    
    return correos

# ─── DEF VALIDAR DOMINIO ─────────────────────────────────────────────────────────────────
def verificar_existencia_correo(correo):
    try:
        informacion = validate_email(correo, check_deliverability=True)
        return True
    except EmailNotValidError as e:
        return False


# ─── DEF DATA ─────────────────────────────────────────────────────────────────
def get_backorder(provider_name: str, branch_name: str, buyer_name: str, df: pd.DataFrame) -> pd.DataFrame:

    df = df[(df['NOMBRE_PROVEEDOR'] == provider_name) &
            (df['NOMBRE_SUCURSAL'] == branch_name) &
            (df['NOMBRE_COMPRADOR'] == buyer_name)]
    return df

# ─── CORREO ───────────────────────────────────────────────────────────────────
def send_email_backorder(df: pd.DataFrame, Proveedor: str):
    """
    Sends the backorder DataFrame as a styled HTML table via email.
    """
    if df.empty:
        print("El reporte está vacío, no se envía correo.")
        return

    html = f"""
    <html>
    <head>
    <style>
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            color: #333333;
            margin: 20px;
        }}
        h2 {{
            color: #0d2c54;
            border-bottom: 2px solid #0d2c54;
            padding-bottom: 8px;
        }}
        table {{
            border-collapse: collapse;
            width: 100%;
            margin-top: 15px;
            font-size: 13px;
        }}
        th {{
            background-color: #0d2c54;
            color: white;
            padding: 10px 12px;
            text-align: left;
            font-weight: 600;
        }}
        td {{
            padding: 10px 12px;
            border-bottom: 1px solid #dddddd;
        }}
        tr:nth-child(even) {{
            background-color: #f8f9fa;
        }}
        tr:hover {{
            background-color: #f1f3f5;
        }}
    </style>
    </head>
    <body>
    <p>Buen día,</p>
    <p>Te comparto el reporte de backorder de MTY para su revisión.</p>
    <h2>Reporte de Backorder - MTY</h2>
    {df.to_html(index=False)}
    <br>
    <p style="font-size: 11px; color: #777777;">Este es un correo automático generado por el sistema.</p>
    </body>
    </html>
    """

    msg = MIMEMultipart("alternative")
    msg["Subject"] = "Reporte de Backorder MTY"
    msg["From"]    = EMISOR
    msg["To"]      = Proveedor #", ".join([r.strip() for r in RECEPTOR.split(",")])
    msg.attach(MIMEText(html, "html"))

    if not EMISOR or not PASSWORD:
        print("⚠️ Advertencia: EMISOR o PASSWORD no están configurados en el archivo .env. No se puede enviar el correo.")
        return

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(EMISOR, PASSWORD)
            smtp.send_message(msg)
        print("✅ Correo enviado")
    except Exception as e:
        print("❌ Error al enviar el correo:", e)

## PROVEEDOR X PROVEEDOR

if __name__ == "__main__":
    try:
      loop_values = clean_email_addresses(loop_values)
      resultados = loop_values['Email'].apply(verificar_existencia_correo)
      loop_values['es_valido'] = resultados

      for row in loop_values.itertuples():
        print(row.Nombre, row.Email)
        print("Obteniendo datos de backorder",row.Nombre)
        df = get_backorder(row.NOMBRE_PROVEEDOR, row.NOMBRE_SUCURSAL, row.NOMBRE_COMPRADOR, backorder)
        print(df)
        
        print("\nEnviando reporte por correo...")
        send_email_backorder(df, row.Email)
    except Exception as e:
        print("Ocurrió un error al consultar BigQuery:", e)