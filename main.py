# main.py

## librerías

import os
import smtplib
from email.message import EmailMessage
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.base import MIMEBase
from email import encoders
import imaplib
import time 
from dotenv import load_dotenv

import pandas as pd
import re
import numpy as np
import ast
import openpyxl
from google.cloud import bigquery
from email_validator import validate_email, EmailNotValidError

from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
from openpyxl.utils import get_column_letter


EJECUTION_MODE = ""

## Extrae variables del .env
load_dotenv()
PROJECT_ID = os.environ["GCP_PROJECT"]
BUYERS_PASSWORD = os.environ.get("BUYERS_PASSWORD", "")
BUYERS_PASSWORD = ast.literal_eval(BUYERS_PASSWORD)
buyer_passwords = pd.DataFrame(list(BUYERS_PASSWORD.items()), columns=['Usuario_ID', 'Password'])
GOOGLE_CREDENTIALS = os.environ["GOOGLE_APPLICATION_CREDENTIALS"]

client = bigquery.Client(project=PROJECT_ID)

DATASET_Y_TABLA = "raw_data.historico_correos_backorder"  
TABLA_COMPLETA_ID = f"{PROJECT_ID}.{DATASET_Y_TABLA}"

job_config = bigquery.LoadJobConfig(
    write_disposition="WRITE_APPEND" 
)

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
# # ) LIKE '%@%'  AND Proveedor IN (95)

# # --LIMIT 100
# # '''

# Query de prueba para envío a destinatario fijo
query = '''SELECT DISTINCT Proveedor, Nombre,
--'leonardo.laureles@danuanalitica.com' AS Email 
--CASE WHEN Proveedor = 10096 THEN 'fmartinez@finsa.com.mx' ELSE 'ruben.garza@finsa.com.mx' END AS Email
    CASE 
        WHEN Proveedor = 95 THEN 'keyla.islas@danuanalitica.com'
        WHEN Proveedor = 3 THEN 'fmartinez@finsa.com.mx'
        WHEN Proveedor = 56 THEN 'keyla.islas@danuanalitica.com'
        ELSE 'lucia.balli@danuanalitica.com' 
    END AS Email
FROM `finsadashboard.raw_data.Proveedores` 
--LIMIT 3
WHERE Proveedor IN (95)
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
            NOMBRE_ALMACEN,
            NOMBRE_COMPRADOR,
            NOMBRE_PROVEEDOR,
            PROVEEDOR,
            1 AS COMPRADOR,
            "pruebas@finsa.com.mx" AS Email_COMPRADOR,
            SUCURSAL, PEDIDO
        FROM `finsadashboard.mrts.mrts_backorder_MTY`
        WHERE BACKORDER > 0 
          AND DIAS_RETRASO_EMBARQUE > 0
          AND  PROVEEDOR IN  (95)
        ORDER BY NOMBRE_PROVEEDOR, FECHA_ALTA
    """

backorder = client.query(query).to_dataframe()


loop_values = backorder[['NOMBRE_COMPRADOR','NOMBRE_PROVEEDOR', 'NOMBRE_SUCURSAL', 'NOMBRE_ALMACEN', 'NOMBRE_COMPRADOR', 'PROVEEDOR', 'COMPRADOR', 'SUCURSAL', 'Email_COMPRADOR']].drop_duplicates()
loop_values = loop_values.merge(correos, left_on='PROVEEDOR', right_on='Proveedor', how='left')
loop_values = loop_values.merge(buyer_passwords, left_on='COMPRADOR', right_on='Usuario_ID', how='left')
loop_values = loop_values[['NOMBRE_PROVEEDOR','NOMBRE_COMPRADOR', 'NOMBRE_SUCURSAL', 'NOMBRE_ALMACEN','Email_COMPRADOR', 'Email', 'Password']]


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
        validate_email(correo, check_deliverability=True)
        return True
    except EmailNotValidError as e:
        return False


# ─── DEF DATA ─────────────────────────────────────────────────────────────────
def get_backorder(provider_name: str, branch_name: str, buyer_name: str, df: pd.DataFrame) -> pd.DataFrame:

    df = df[(df['NOMBRE_PROVEEDOR'] == provider_name) &
            (df['NOMBRE_SUCURSAL'] == branch_name) &
            (df['NOMBRE_COMPRADOR'] == buyer_name)]
    df = df[['FECHA_ALTA','ORDEN_COMPRA','PEDIDO','ARTICULO','DESCRIPCION','EDP','PARTIDA','CANTIDAD','CANTIDAD_RECIBIDA','BACKORDER','UNIDAD','FECHA_EMBARQUE','DIAS_RETRASO_EMBARQUE']]
    df.columns = ['FECHA ALTA','ORDEN DE COMPRA','PEDIDO','ARTICULO','DESCRIPCIÓN','EDP','PARTIDA','CANTIDAD','CANTIDAD RECIBIDA','BACKORDER','UNIDAD','FECHA DE EMBARQUE','DIAS DE RETRASOEMBARQUE']
    return df

def limpiar_cadena(cadena: str) -> str:
    if pd.isna(cadena):
        return ""
    # Eliminar caracteres especiales y acentos
    cadena = re.sub(r'[^a-zA-Z0-9]', '', cadena)
    return cadena.strip()

# ─── CORREO ───────────────────────────────────────────────────────────────────
def send_email_backorder(df: pd.DataFrame, Proveedor : str, Email_proveedor: str, Email_comprador: str, Comprador: str,  Sucursal: str, Password: str, Almacen: str) -> None:
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
    <p><b>Buen día estimado proveedor espero y se encuentre bien</b></p>
    <p><b>Anexo al presente un archivo en Excel que contiene todas las líneas abiertas pendientes de entrega (BackOrder), por favor le solicitamos enviar las fechas de entrega de cada una de las partidas en la mayor brevedad posible, sabiendo que las fechas indicadas son de arribo en almacén de FINSA.</b></p>
    <p><b>Para Finsa y proveedores es importante mantener al cliente final informado sobre la entrega de sus productos</b></p>
    <p><b>Agradecemos su puntal apoyo</b></p>
    <p><b>Saludos</b></p>
    <h2>Reporte de Backorder - {Sucursal}</h2>
    {df.to_html(index=False)}
    <br>
    <p style="font-size: 11px; color: #777777;">Este es un correo automático generado por el sistema.</p>
    </body>
    </html>
    """

    prov = limpiar_cadena(Proveedor)
    suc = limpiar_cadena(Sucursal)

    excel_file = f"Backorder_{prov}_{suc}.xlsx"

    df1 = df.copy()
    df1['NUEVA FECHA COMPROMISO'] = ""

    with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
        df1.to_excel(writer, index=False, sheet_name='Datos')
        
        ws = writer.sheets['Datos']
        
        header_fill = PatternFill("solid", fgColor="4472C4")
        for cell in ws[1]:  # fila 1 = encabezados
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal='center')
        
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            ws.column_dimensions[get_column_letter(col[0].column)].width = max_len + 4

    msg = MIMEMultipart("mixed") #alternative
    msg["Subject"] = f"Reporte de Backorder - {Proveedor}"
    msg["From"]    = Email_comprador
    msg["To"]      = Email_proveedor #", ".join([r.strip() for r in RECEPTOR.split(",")])
    msg["Cc"]      = "leonardo.laureles@danuanalitica.com"
    msg['User-Agent'] = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Python-SMTPLIB"
    msg['X-Mailer'] = "Python-SMTP-Client"
    msg.attach(MIMEText(html, "html"))
    # with open(excel_file, "rb") as f:
    #     part = MIMEApplication(f.read(), Name=excel_file)
    # part['Content-Disposition'] = f'attachment; filename={excel_file}'
    # msg.attach(part)



    try:
        with open(excel_file, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
        
        encoders.encode_base64(part)
        
        part.add_header(
            "Content-Disposition",
            f"attachment; filename={excel_file}"
        )
        msg.attach(part)
        print("📦 Archivo Excel empaquetado correctamente")

    except FileNotFoundError:
        print(f"❌ Error: No se encontró el archivo {excel_file}")
        return





    if not Email_comprador or not Password:
        print("⚠️ Advertencia: Email_comprador o Password no están configurados en el archivo .env. No se puede enviar el correo.")
        return
    
    HOST =  "finsa--com--mx.criticalmail.net" #"smtp.gmail.com" #

    try:
        with smtplib.SMTP_SSL(HOST, 465, timeout=30) as smtp:
            time.sleep(1)
            smtp.login(Email_comprador, Password)
            smtp.send_message(msg)
            print("✅ Correo enviado")

            if EJECUTION_MODE != "PRUEBA":

                try:
                    print(f"Subiendo datos a {TABLA_COMPLETA_ID}...")
                    df = df.copy()
                    df['FECHA_ENVIO'] = pd.Timestamp.now()
                    df['Email_proveedor'] = Email_proveedor
                    df['Email_comprador'] = Email_comprador
                    df['Comprador'] = Comprador
                    df['Sucursal'] = Sucursal
                    df['Proveedor'] = Proveedor
                    df['Almacen'] = Almacen   
                    job = client.load_table_from_dataframe(df, TABLA_COMPLETA_ID, job_config=job_config)
                    job.result()
                        
                    print(f"¡Tabla subida exitosamente! Se cargaron {job.output_rows} filas.")

                except Exception as e:
                    print(f"Error al subir los datos a BigQuery: {e}")
        
    except Exception as e:
        print("❌ Error al enviar el correo:", e)

    try:
        with imaplib.IMAP4_SSL(HOST, 993, timeout=30) as imap:
            imap.login(Email_comprador, Password)
            carpeta_enviados = "Sent" 
            
            imap.append(
                carpeta_enviados, 
                r'\Seen', 
                imaplib.Time2Internaldate(time.time()), 
                msg.as_bytes()
            )
            print("Copia guardada con éxito en la carpeta de Enviados (IMAP).")
    except Exception as e:
        print(f"No se pudo guardar la copia en Enviados: {e}")
        

## PROVEEDOR X PROVEEDOR

if __name__ == "__main__":
    try:
      loop_values = clean_email_addresses(loop_values)
      resultados = loop_values['Email'].apply(verificar_existencia_correo)
      loop_values['es_valido'] = resultados
      loop_values = loop_values[loop_values['es_valido'] == True]
      print(loop_values)

      for row in loop_values.itertuples():
        print(row.NOMBRE_PROVEEDOR, row.Email)
        print("Obteniendo datos de backorder",row.NOMBRE_PROVEEDOR)
        df = get_backorder(row.NOMBRE_PROVEEDOR, row.NOMBRE_SUCURSAL, row.NOMBRE_COMPRADOR, backorder)
        print(df)
        
        print("\nEnviando reporte por correo...")
        send_email_backorder(df, row.NOMBRE_PROVEEDOR, row.Email, row.Email_COMPRADOR, row.NOMBRE_COMPRADOR, row.NOMBRE_SUCURSAL, row.Password, row.NOMBRE_ALMACEN)

    except Exception as e:
        print("Ocurrió un error al consultar BigQuery:", e)