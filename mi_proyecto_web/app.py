from flask import Flask, render_template, request, redirect, url_for, flash, send_file
from sqlalchemy import create_engine, text
import pandas as pd
import os
import re
import logging
import pyodbc
import getpass  # Importar el módulo para obtener el nombre del usuario
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill 
from apscheduler.schedulers.background import BackgroundScheduler
import smtplib
from email.mime.text import MIMEText
from db_utils import obtener_datos_kpi, obtener_datos_clicks
from excel_utils import generar_graficos_excel
import io
import base64
from matplotlib.figure import Figure
import plotly.graph_objs as go
import subprocess
from matplotlib.backends.backend_pdf import PdfPages
import matplotlib
matplotlib.use('Agg')  # Establece el backend 'Agg' para evitar el error de Tcl/Tk
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches

app = Flask(__name__)
app.secret_key = "your_secret_key"

# Configurar logging
logging.basicConfig(level=logging.INFO)

@app.route('/')
def inicio():
    return render_template('inicio.html')

@app.route('/index')
def index():
    return render_template('index.html')

@app.route('/main_app')
def main_app():
    return redirect(url_for('index'))  # Redirige a la función de carga de archivos


# Ruta para la login opción

@app.route('/login')
def login():
    return render_template('login.html')



def get_connection(db_name):
    try:
        connection = pyodbc.connect(
            f"DRIVER={{ODBC Driver 17 for SQL Server}};SERVER=192.168.201.12;DATABASE={db_name};Trusted_Connection=yes;"
        )
        logging.info("Conexión establecida correctamente con la base de datos.")
        return connection
    except pyodbc.Error as e:
        logging.error(f"Error al conectar con la base de datos: {e}")
        raise Exception(f"Error al conectar con la base de datos: {e}")

def ejecutar_consulta(connection, query):
    try:
        cursor = connection.cursor()
        cursor.execute(query)
        rows = cursor.fetchall()
        columns = [column[0] for column in cursor.description]
        logging.info(f"Consulta ejecutada correctamente: {query}")
        logging.info(f"Filas devueltas: {rows}")
        return [dict(zip(columns, row)) for row in rows]
    except pyodbc.Error as e:
        logging.error(f"Error al ejecutar la consulta: {e}")
        return f"Error al ejecutar la consulta: {e}"



def generar_graficos(resultados, db_name):
    pdf_buffer = io.BytesIO()
    imagenes_base64 = []

    logging.info(f"Generando gráficos para la base de datos: {db_name}")

    # Crear gráficos
    if resultados and resultados[0] and "TotalRecords" in resultados[0][0]:
        fig1, ax1 = plt.subplots()
        labels = ['Total Records', 'Complete Records', 'Correction Records']
        sizes = [
            resultados[0][0].get('TotalRecords', 0),
            resultados[0][0].get('CompleteRecords', 0),
            resultados[0][0].get('CorrectionRecords', 0)
        ]
        ax1.pie(sizes, labels=labels, autopct='%1.1f%%', startangle=90)
        ax1.axis('equal')
        plt.title('Distribución de Registros con CNAE')
        
        # Añadir leyenda
        total_registros = resultados[0][0].get('TotalRecords', 0)
        registros_con_cnae = resultados[0][0].get('CompleteRecords', 0)
        registros_sin_cnae = total_registros - registros_con_cnae
        patches = [
            mpatches.Patch(color='blue', label=f'Total de registros: {total_registros}'),
            mpatches.Patch(color='orange', label=f'Total de Registros con CNAE: {registros_con_cnae}'),
            mpatches.Patch(color='green', label=f'Total de Registros sin CNAE: {registros_sin_cnae}')
        ]
        ax1.legend(handles=patches, loc="best")
        
        buffer1 = io.BytesIO()
        fig1.savefig(buffer1, format='png')
        buffer1.seek(0)
        imagenes_base64.append(base64.b64encode(buffer1.read()).decode('utf-8'))
        plt.close(fig1)
        logging.info("Gráfico 1 generado con éxito.")
    else:
        imagenes_base64.append(None)
        logging.warning("No hay datos para el gráfico 1.")

    # Modificación del segundo gráfico:
    if resultados and resultados[1] and "ClickDate" in resultados[1][0]:
        fig2, ax2 = plt.subplots()
        fechas = [row['ClickDate'] for row in resultados[1]]
        clicks = [row['ClickCount'] for row in resultados[1]]
        
        # Añadir distinción visual
        opciones_linkedin = [clicks[i] if i % 2 == 0 else 0 for i in range(len(clicks))]
        telefono_alternativo = [clicks[i] if i % 2 != 0 else 0 for i in range(len(clicks))]

        ax2.bar(fechas, opciones_linkedin, color='blue', label='LinkedIn')
        ax2.bar(fechas, telefono_alternativo, color='green', bottom=opciones_linkedin, label='Teléfono Alternativo')

        plt.title('Clicks por Fecha')
        plt.xlabel('Fecha')
        plt.ylabel('Cantidad de Clicks')
        plt.xticks(rotation=45, ha='right')  # Rotar etiquetas de fechas para que se vean mejor
        plt.legend()
        
        buffer2 = io.BytesIO()
        fig2.savefig(buffer2, format='png')
        buffer2.seek(0)
        imagenes_base64.append(base64.b64encode(buffer2.read()).decode('utf-8'))
        plt.close(fig2)
        logging.info("Gráfico 2 generado con éxito.")
    else:
        imagenes_base64.append(None)
        logging.warning("No hay datos para el gráfico 2.")

    if resultados and resultados[2] and "Total" in resultados[2][0]:
        fig3, ax3 = plt.subplots()
        labels = ['Total', 'Contactadas']
        values = [
            resultados[2][0].get('Total', 0),
            resultados[2][0].get('Contactadas', 0)
        ]
        ax3.bar(labels, values, color=['green', 'orange'])
        plt.title('Contactadas vs Total')
        buffer3 = io.BytesIO()
        fig3.savefig(buffer3, format='png')
        buffer3.seek(0)
        imagenes_base64.append(base64.b64encode(buffer3.read()).decode('utf-8'))
        plt.close(fig3)
        logging.info("Gráfico 3 generado con éxito.")
    else:
        imagenes_base64.append(None)
        logging.warning("No hay datos para el gráfico 3.")

    # Guardar gráficos en PDF
    with PdfPages(pdf_buffer) as pdf:
        for i, img in enumerate(imagenes_base64):
            if img:
                fig, ax = plt.subplots()
                ax.imshow(plt.imread(io.BytesIO(base64.b64decode(img))))
                ax.axis('off')
                pdf.savefig(fig)
                plt.close(fig)

    pdf_buffer.seek(0)
    return imagenes_base64, pdf_buffer

@app.route('/graficos', methods=['GET', 'POST'])
def graficos():
    if request.method == 'POST':
        db_name = request.form.get('db_name')

        if not db_name:
            return "Por favor, introduce una base de datos.", 400

        try:
            connection = get_connection(db_name)

            query1 = "SELECT TotalRecords, CompleteRecords, CorrectionRecords FROM CNAE_KPI_Audit"
            query2 = "SELECT CAST(ClickDate AS DATE) AS ClickDate, COUNT(*) AS ClickCount FROM LinkedinClickLog GROUP BY CAST(ClickDate AS DATE)"
            query3 = "SELECT COUNT(*) AS Total, SUM(CASE WHEN PERSONA_CONTACTADA IS NOT NULL THEN 1 ELSE 0 END) AS Contactadas FROM Empresas"

            resultados = [
                ejecutar_consulta(connection, query1),
                ejecutar_consulta(connection, query2),
                ejecutar_consulta(connection, query3)
            ]

            # Log de los resultados de las consultas
            logging.info(f"Resultados de la consulta 1: {resultados[0]}")
            logging.info(f"Resultados de la consulta 2: {resultados[1]}")
            logging.info(f"Resultados de la consulta 3: {resultados[2]}")

            imagenes_base64, pdf_buffer = generar_graficos(resultados, db_name)

            return render_template('graficos.html', imagenes_base64=imagenes_base64, pdf_link=url_for('descargar_pdf_file', db_name=db_name))
        except Exception as e:
            logging.error(f"Ocurrió un error: {e}")
            return f"Ocurrió un error: {e}", 500

    return render_template('graficos.html')

@app.route('/descargar_pdf/<db_name>', methods=['GET'])
def descargar_pdf_file(db_name):
    try:
        connection = get_connection(db_name)

        query1 = "SELECT TotalRecords, CompleteRecords, CorrectionRecords FROM CNAE_KPI_Audit"
        query2 = "SELECT CAST(ClickDate AS DATE) AS ClickDate, COUNT(*) AS ClickCount FROM LinkedinClickLog GROUP BY CAST(ClickDate AS DATE)"
        query3 = "SELECT COUNT(*) AS Total, SUM(CASE WHEN PERSONA_CONTACTADA IS NOT NULL THEN 1 ELSE 0 END) AS Contactadas FROM Empresas"

        resultados = [
            ejecutar_consulta(connection, query1),
            ejecutar_consulta(connection, query2),
            ejecutar_consulta(connection, query3)
        ]

        _, pdf_buffer = generar_graficos(resultados, db_name)

        return send_file(pdf_buffer, as_attachment=True, download_name=f"graficos_{db_name}_{datetime.now().strftime('%Y%m%d')}.pdf", mimetype='application/pdf')
    except Exception as e:
        logging.error(f"Ocurrió un error: {e}")
        return f"Ocurrió un error: {e}", 500


@app.route('/descargar-excel')
def descargar_excel():
    excel_path = os.path.join(os.getcwd(), "env", "mi_proyecto_web", "mi_proyecto_web", "kpi_clicks_report.xlsx")
    return send_file(excel_path, as_attachment=True)

# Ruta para la tercera opción, actualmente pendiente de definir
@app.route('/pendiente')
def pendiente():
    return "<h1>Esta función está pendiente de implementación</h1>"


# Configuración de la base de datos
DATABASE_URI = "mssql+pyodbc://sa:infinity@192.168.201.12/HERMESV5_HISTORICO?driver=ODBC+Driver+17+for+SQL+Server"
engine = create_engine(DATABASE_URI)

# Configurar el logging
LOG_FOLDER = r'C:\env\mi_proyecto_web\logs'  # Cambiar a ruta absoluta
if not os.path.exists(LOG_FOLDER):
    os.makedirs(LOG_FOLDER)

# Configurar el logging para crear un archivo diario
def configurar_logging():
    log_filename = os.path.join(LOG_FOLDER, f"{datetime.now().strftime('%Y-%m')}.log")
    logging.basicConfig(
        filename=log_filename,
        level=logging.INFO,
        format='%(asctime)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    print(f"Logging configurado correctamente en: {log_filename}")

configurar_logging()

# Crear la carpeta 'uploads' si no existe
UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)


# Función para enviar correos
def enviar_correo(asunto, cuerpo):
    destinatario = "fdargallo@telcohumanmedia.com"
    remitente = "fdargallo@telcohumanmedia.com"
    password = "Frasco401*"

    msg = MIMEText(cuerpo)
    msg['Subject'] = asunto
    msg['From'] = remitente
    msg['To'] = destinatario

    try:
        with smtplib.SMTP('smtp.office365.com', 587) as server:
            server.starttls()
            server.login(remitente, password)
            server.send_message(msg)
        print("Correo enviado correctamente.")
    except Exception as e:
        print(f"Error al enviar el correo: {e}")

# Variable para rastrear si el informe ya fue enviado este mes
informe_enviado = False

def enviar_informe_mensual():
    global informe_enviado
    if informe_enviado:
        return  # Si ya se envió, no hacer nada

    log_filename = os.path.join(LOG_FOLDER, f"{(datetime.now() - pd.DateOffset(months=1)).strftime('%Y-%m')}.log")
    if os.path.exists(log_filename):
        enviar_correo("Informe mensual", f"Adjunto el informe del mes anterior: {log_filename}")
    else:
        enviar_correo("Informe mensual", "No había fichero disponible para el mes anterior.")
    
    informe_enviado = True  # Marcar como enviado

# Crear el scheduler
scheduler = BackgroundScheduler()

# Programar la tarea para enviar el informe mensual
scheduler.add_job(enviar_informe_mensual, 'cron', day=1, hour=0)  # Ejecutar el día 1 de cada mes a la medianoche


# Función para registrar en el log los datos específicos
def log_usuario(nombre_usuario, nombre_archivo, registros_originales, registros_limpios):
    logging.info(
        f"Usuario: {nombre_usuario}, Archivo subido: {nombre_archivo}, "
        f"Registros originales: {registros_originales}, Teléfonos limpios: {registros_limpios}"
    )
    print("Log de usuario registrado.")

def limpiar_telefono(telefono):
    if pd.isna(telefono):
        return None
    telefono = str(telefono)
    telefono = re.sub(r'[ .\-\/]', '', telefono)
    telefono = re.sub(r'^(\+34|0034)', '', telefono)
    return telefono if len(telefono) == 9 and telefono.isdigit() else None


@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        flash('No file part')
        return redirect(url_for('index'))

    file = request.files['file']
    if file.filename == '':
        flash('No selected file')
        return redirect(url_for('index'))

    if file:
        filepath = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(filepath)

        # Procesar el archivo Excel
        df = pd.read_excel(filepath)

        # Loguear los datos iniciales
        registros_originales = len(df)
        nombre_usuario = getpass.getuser()  # Obtener el nombre del usuario logueado

        # Renombrar las columnas a un formato estándar
        df.rename(columns=lambda x: x.strip().lower(), inplace=True)
        df.rename(columns={
            'telefono': 'telefono',
            'teléfono': 'telefono',
            'telf1': 'telefono',
            'tel1': 'telefono',
            'telefono1': 'telefono',
            'TELEFONO1': 'telefono',
            'TELEFONO': 'telefono',
            'phone': 'telefono'
        }, inplace=True)

        # Verificar si 'telefono' existe
        if 'telefono' not in df.columns:
            raise KeyError("No se encontró la columna 'telefono' en el DataFrame.")

        # Limpiar los teléfonos
        df['telefono_limpio'] = df['telefono'].apply(limpiar_telefono)
        registros_limpios = df['telefono_limpio'].notnull().sum()

        # Registrar la actividad en el log
        log_usuario(nombre_usuario, file.filename, registros_originales, registros_limpios)

        # Conectar a la base de datos y ejecutar la lógica de scoring
        with engine.connect() as conn:
            for index, row in df.iterrows():
                telefono = str(row['telefono'])

                # Consulta de scoring
                query_scoring = text("""
                    SELECT CASE 
                        WHEN COUNT(*) >= 20 THEN (CAST(SUM(CASE WHEN CallStatusNum < 11 THEN 1 ELSE 0 END) AS FLOAT) / COUNT(*)) * 100 
                        ELSE (SUM(CASE WHEN CallStatusNum < 11 THEN 1 ELSE 0 END) / 20.0) * 100 
                    END AS conteo 
                    FROM ODCalls WHERE ANI = :telefono
                """)
                result_scoring = conn.execute(query_scoring, {'telefono': telefono}).fetchone()
                conteo = result_scoring[0] if result_scoring else 0
                df.at[index, 'scoring'] = conteo

                # Consulta de media de intentos para contacto
                query_media_intentos = text("""
                    SELECT CASE 
                        WHEN COUNT(*) >= 20 THEN (COUNT(*) / CAST(NULLIF(SUM(CASE WHEN CallStatusNum < 11 THEN 1 ELSE 0 END), 0) AS FLOAT)) 
                        ELSE (20.0 / NULLIF(SUM(CASE WHEN CallStatusNum < 11 THEN 1 ELSE 0 END), 0))  
                    END AS conteo 
                    FROM ODCalls 
                    WHERE ANI = :telefono AND Duration > 1
                """)

                result_media_intentos = conn.execute(query_media_intentos, {'telefono': telefono}).fetchone()
                media_intentos = result_media_intentos[0] if result_media_intentos else None
                df.at[index, 'media_intentos_para_contacto'] = media_intentos

                # Consulta de media de intentos para contacto positivo
                query_media_intentos_positivo = text("""
                    SELECT 
                        CASE 
                            WHEN COUNT(CASE WHEN Duration > 1 AND CallStatusNum BETWEEN 1 AND 10 THEN 1 END) = 0 THEN 0
                            ELSE COUNT(*) * 1.0 / COUNT(CASE WHEN Duration > 1 AND CallStatusNum BETWEEN 1 AND 10 THEN 1 END)
                        END AS media_llamadas_para_contacto_positivo
                    FROM ODCalls
                    WHERE ANI = :telefono AND (Duration > 1 OR CallStatusNum BETWEEN 1 AND 10)
                """)

                result_media_intentos_positivo = conn.execute(query_media_intentos_positivo, {'telefono': telefono}).fetchone()
                
                
                
                # Agregar un log para depuración
                if result_media_intentos_positivo is None:
                    logging.info(f'No se encontraron resultados para el teléfono: {telefono}')
                #else:
                    #logging.info(f'Resultado para el teléfono {telefono}: {result_media_intentos_positivo[0]}')

                media_intentos_positivo = result_media_intentos_positivo[0] if result_media_intentos_positivo else None
                df.at[index, 'media_intentos_para_contacto_positivo'] = media_intentos_positivo

        # Ordenar el DataFrame de mayor a menor según el scoring
        df.sort_values(by='scoring', ascending=False, inplace=True)

        # Convertir los valores numéricos pequeños a un formato más legible con 2 decimales
        df['media_intentos_para_contacto_positivo'] = df['media_intentos_para_contacto_positivo'].apply(lambda x: '{:.2f}'.format(x) if pd.notna(x) else x)

        # Guardar el archivo Excel con el scoring, media de intentos y todos los datos
        output_filepath = os.path.join(UPLOAD_FOLDER, 'resultados_Scoring.xlsx')
        df.to_excel(output_filepath, index=False)

        # Aplicar estilo de colores en la columna 'scoring' usando openpyxl
        wb = load_workbook(output_filepath)
        ws = wb.active
        scoring_column_index = ws.max_column - 2  # Columna 'scoring'
        media_intentos_column_index = ws.max_column - 1  # Columna 'media_intentos_para_contacto'
        media_intentos_positivo_column_index = ws.max_column  # Última columna es 'media_intentos_para_contacto_positivo'

        for row in range(2, ws.max_row + 1):  # Empezamos en la fila 2 para evitar el encabezado
            score = ws.cell(row=row, column=scoring_column_index).value
            if score is not None:
                # Definir el color basado en el score (de verde a rojo)
                red = int(min(255, (255 - int(2.55 * score))))
                green = int(min(255, int(2.55 * score)))
                fill = PatternFill(start_color=f"{red:02X}{green:02X}00", end_color=f"{red:02X}{green:02X}00", fill_type="solid")
                
                ws.cell(row=row, column=scoring_column_index).fill = fill

        wb.save(output_filepath)
        wb.close()

        # Generar la tabla HTML para mostrar en el navegador
        resultados_limpieza = df.to_html(classes=['table', 'scoring'], justify='center')

        return render_template('resultados.html', tabla=resultados_limpieza)

@app.route('/download')
def download_file():
    # Descargar el archivo de resultados
    output_filepath = os.path.join(UPLOAD_FOLDER, 'resultados_Scoring.xlsx')
    return send_file(output_filepath, as_attachment=True)


if __name__ == '__main__':
    configurar_logging()
    # Prueba de envío de correo después de procesar el archivo
    enviar_correo("Prueba de correo", "Este es un mensaje de prueba para verificar el envío de correos.")
    if datetime.now().day == 1:
        enviar_informe_mensual()
    if not scheduler.running:
        scheduler.start()  # Inicia el scheduler solo si no está en ejecución
    app.run(debug=True, host='127.0.0.1', port=5000)
