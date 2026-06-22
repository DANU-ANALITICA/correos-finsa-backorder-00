# main.py
import asyncio
import io
import os
import re
import smtplib
import time
import unicodedata
import zipfile
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from pathlib import Path

from dotenv import load_dotenv


import numpy as np
import pandas as pd
from google.cloud import bigquery
from playwright.async_api import async_playwright

#----ESTO SALE DEL .env
load_dotenv()
PROJECT_ID = os.environ["GCP_PROJECT"]
EMISOR = os.environ.get("EMISOR", "")
PASSWORD = os.environ.get("PASSWORD", "")
###RECEPTOR = os.environ.get("RECEPTOR", "")

client = bigquery.Client(project=PROJECT_ID)
# query = '''SELECT DISTINCT Proveedor, Nombre,
# COALESCE(
#   NULLIF(Email1, ''),
#   NULLIF(Email2, ''),
#   NULLIF(Email3, '')
# ) AS Email FROM `finsadashboard.raw_data.Proveedores`
# WHERE COALESCE(
#   NULLIF(Email1, ''),
#   NULLIF(Email2, ''),
#   NULLIF(Email3, '')
# )  is not NULL'''

query = '''SELECT DISTINCT Proveedor, Nombre,
'daniel.perez@danuanalitica.com' AS Email 
FROM `finsadashboard.raw_data.Proveedores` LIMIT 3
'''

correos = client.query(query).to_dataframe()
correos

# ─── BIGQUERY ─────────────────────────────────────────────────────────────────
def get_backorder_mty(provider_name: str ) -> pd.DataFrame:
    """
    Retrieves backorder records from BigQuery for a specific provider.
    """
    #client = bigquery.Client(project=PROJECT_ID)
    
    query = """
        SELECT 
            NOMBRE_PROVEEDOR,
            FECHA_ALTA,
            SUCURSAL,
            ALMACEN,
            ARTICULO,
            CANTIDAD,
            CANTIDAD_RECIBIDA,
            BACKORDER,
            UNIDAD,
            COSTO,
            FECHA_EMBARQUE,
            DIAS_RETRASO_EMBARQUE
        FROM `finsadashboard.mrts.mrts_backorder_MTY`
        WHERE BACKORDER > 0 
          AND NOMBRE_PROVEEDOR = @provider_name
        ORDER BY NOMBRE_PROVEEDOR, FECHA_ALTA DESC
    """
    
    job_config = bigquery.QueryJobConfig(
        query_parameters=[
            bigquery.ScalarQueryParameter("provider_name", "STRING", provider_name)
        ]
    )
    
    df = client.query(query, job_config=job_config).to_dataframe()
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

if __name__ == "__main__":
    try:
      for row in correos.itertuples():
        print(row.Nombre, row.Email)
        print("Obteniendo datos de backorder",row.Nombre)
        df = get_backorder_mty(row.Nombre)
        print(df)
        
        print("\nEnviando reporte por correo...")
        send_email_backorder(df, row.Email)
    except Exception as e:
        print("Ocurrió un error al consultar BigQuery:", e)
