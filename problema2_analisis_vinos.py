# 1. Configuración y Carga de Datos
import pandas as pd
import sqlite3 # Necesario para exportar a SQLite
import json # Necesario para exportar a MongoDB/JSON
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

# URLs de los archivos CSV
URL_VINOS = "/workspaces/PythonPC5/PC5/data/winemag-data-130k-v2.csv"
URL_PAISES = "https://gist.githubusercontent.com/kintero/7d1db891401f56256c79/raw/a61f6d0dda82c3f04d2e6e76c3870552ef6cf0c6/paises.csv" 

# 1. Cargar los DataFrames
try:
    df_vinos = pd.read_csv(URL_VINOS, index_col=0)
    df_paises = pd.read_csv(URL_PAISES)
except Exception as e:
    print(f"Error al cargar archivos: {e}")
    exit()

print("Datos cargados exitosamente. Iniciando análisis...")

# 2. Exploración y Transformación
# a. Renombrar columnas para claridad
df_vinos.rename(columns={
    'points': 'Puntuacion',
    'price': 'Precio',
    'country': 'Pais',
    'taster_name': 'Catador'
}, inplace=True)


# b. Creación de columnas
# b.1. Etiquetar Países según Continente

# Renombrar las columnas de df_paises para coincidir con el esquema esperado
# 'nombre' (país) -> 'Pais'
# 'continente' (continente) -> 'Continente'
df_paises.rename(columns={
    'nombre': 'Pais',
    'continente': 'Continente'
}, inplace=True)

# Realizar el merge limpio. Como las columnas de fusión ahora se llaman igual ('Pais'), se usa 'on'.
df_vinos = pd.merge(
    df_vinos, 
    df_paises, 
    on='Pais', 
    how='left'
)

# Limpiar nulos para el análisis de continente/país
df_vinos.dropna(subset=['Pais', 'Continente'], inplace=True) 

# b.2. Crear columna: Indicador de Calidad
df_vinos['Alta_Calidad'] = df_vinos['Puntuacion'] > 90

# b.3. Crear columna: Puntos por Euro (Densidad de Precio)
df_vinos['Puntos_por_Euro'] = df_vinos['Puntuacion'] / df_vinos['Precio']
df_vinos.dropna(subset=['Puntos_por_Euro'], inplace=True) # Eliminar filas donde el precio era nulo


# 3. Generación de 4 Reportes Distintos
# R1: Top 10 de vinos con la mejor relación Puntuación/Precio
reporte_1 = df_vinos.sort_values(
    by=['Puntuacion', 'Puntos_por_Euro'],
    ascending=[False, False]
).head(10)[['title', 'Puntuacion', 'Precio', 'Pais', 'Puntos_por_Euro']]

# R2: Resumen de Puntuación y Precio por País
reporte_2 = df_vinos.groupby('Pais').agg(
    Puntuacion_Promedio=('Puntuacion', 'mean'),
    Precio_Promedio=('Precio', 'mean'),
    Total_Vinos=('title', 'count')
).sort_values(by='Puntuacion_Promedio', ascending=False).head(15)

# R3: Distribución de vinos de Alta Calidad (>90 puntos) por Continente
reporte_3 = df_vinos[df_vinos['Alta_Calidad'] == True].groupby('Continente').agg(
    Vinos_Alta_Calidad=('Alta_Calidad', 'count'),
    Puntuacion_Maxima=('Puntuacion', 'max')
).sort_values(by='Vinos_Alta_Calidad', ascending=False)

# R4: Top 10 Variedades de uva más caras (promedio de precio)
reporte_4 = df_vinos.groupby('variety').agg(
    Precio_Promedio=('Precio', 'mean'),
    Total_Vinos=('title', 'count')
).sort_values(by='Precio_Promedio', ascending=False).head(10)


# 4. Exportación de los 4 Reportes (4 Formatos Diversos)
print("\nExportando reportes...")

# Reporte 1 a CSV
reporte_1.to_csv("reporte_1_top_vinos.csv", index=False)
print("- Reporte 1 guardado como reporte_1_top_vinos.csv (CSV)")

# Reporte 2 a Excel
reporte_2.to_excel("reporte_2_pais_resumen.xlsx")
print("- Reporte 2 guardado como reporte_2_pais_resumen.xlsx (Excel)")

# Reporte 3 a SQLite
conn = sqlite3.connect('reportes_vinos.db')
reporte_3.to_sql('alta_calidad_continente', conn, if_exists='replace')
conn.close()
print("- Reporte 3 guardado en base de datos SQLite (reportes_vinos.db)")

# Reporte 4 a JSON
reporte_4.reset_index().to_json("reporte_4_variedades_caras.json", orient='records', indent=4)
print("- Reporte 4 guardado como reporte_4_variedades_caras.json (JSON)")


# 5. Envío de Correo

def enviar_correo(archivo_adjunto):
    GMAIL_USER = 'ravargasm@unac.edu.pe'
    
    # Leer la Contraseña de Aplicación de la variable de entorno
    try:
        # La variable de entorno SE LLAMA GMAIL_PASS en mi sistema
        GMAIL_PASS = os.environ['GMAIL_PASS']
    except KeyError:
        print("\nERROR DE SEGURIDAD: La variable de entorno 'GMAIL_PASS' no está configurada.")
        print("Antes de ejecutar, configura la variable en tu terminal usando 'export GMAIL_PASS=\"[Contraseña de 16 caracteres]\"'")
        return 

    remitente = GMAIL_USER
    destinatario = 'sanchescinthya21@gmail.com' # Correo al que se envía la prueba

    # 1. Crear el mensaje
    msg = MIMEMultipart()
    msg['From'] = remitente
    msg['To'] = destinatario
    msg['Subject'] = "Reporte de Vinos - Problema 2"
    
    # 2. Adjuntar archivo
    part = MIMEBase('application', 'octet-stream')
    # Manejo de errores si el archivo no existe
    try:
        part.set_payload(open(archivo_adjunto, 'rb').read())
    except FileNotFoundError:
        print(f"\nERROR: El archivo adjunto '{archivo_adjunto}' no fue encontrado.")
        return 
        
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename="{archivo_adjunto}"')
    msg.attach(part)

    # 3. Enviar correo
    try:
        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server.login(GMAIL_USER, GMAIL_PASS)
        server.sendmail(remitente, destinatario, msg.as_string())
        server.quit()
        print(f"\nCorreo enviado exitosamente con el archivo {archivo_adjunto}")
    except Exception as e:
        print(f"\nError al enviar correo. Asegúrate de que la Contraseña de Aplicación sea correcta y que el remitente sea válido: {e}")

# Ejecutar el envío del correo con el primer reporte
enviar_correo("reporte_1_top_vinos.csv")
