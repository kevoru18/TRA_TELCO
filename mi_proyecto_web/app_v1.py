from flask import Flask, render_template, request, redirect, url_for, flash, send_file
from sqlalchemy import create_engine
import pandas as pd
import os
import re
import logging
import datetime

app = Flask(__name__)
app.secret_key = "your_secret_key"

# Configuración de la base de datos
DATABASE_URI = "mssql+pyodbc://{usuario}:{contraseña}@{servidor}/{base_de_datos}?driver=ODBC+Driver+17+for+SQL+Server"
engine = create_engine(DATABASE_URI.format(
    usuario="your_username",
    contraseña="your_password",
    servidor="your_server_name",
    base_de_datos="your_database_name"
))

# Configurar el logging
logging.basicConfig(filename='activity.log', level=logging.INFO)

# Crear la carpeta 'uploads' si no existe
UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

def limpiar_telefono(telefono):
    if pd.isna(telefono):
        return None
    telefono = str(telefono)
    telefono = re.sub(r'[ .\-\/]', '', telefono)
    telefono = re.sub(r'^(\+34|0034)', '', telefono)
    if len(telefono) == 9 and telefono.isdigit():
        return telefono
    else:
        return None

@app.route('/')
def index():
    return render_template('index.html')

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

        # Renombrar columnas relevantes a 'telefono'
        columnas_telefono = ['Telefono', 'telefono', 'Teléfono', 'teléfono', 'TELF1', 'tel1', 'phone', 'Phone']
        for columna in columnas_telefono:
            if columna in df.columns:
                df.rename(columns={columna: 'telefono'}, inplace=True)
                break

        if 'telefono' not in df.columns:
            flash('No se encontró una columna de teléfono válida')
            return redirect(url_for('index'))

        # Limpiar los números de teléfono
        df['telefono_limpio'] = df['telefono'].apply(limpiar_telefono)

        # Supongamos que hay una columna 'llamadas' en el DataFrame
        df['llamadas'] = df['llamadas'].fillna(0).astype(int)

        # Agrupar por 'telefono_limpio' y contar llamadas
        resultados = df.groupby('telefono_limpio')['llamadas'].sum().reset_index()
        resultados.rename(columns={'llamadas': 'total_llamadas'}, inplace=True)

        # Guardar los resultados en un archivo Excel
        resultados.to_excel('resultados.xlsx', index=False)

        # Mostrar los resultados de la limpieza en la web
        resultados_html = resultados.to_html()
        
        # Registrar la actividad
        logging.info(f'Archivo {file.filename} cargado y procesado por {request.remote_addr} el {datetime.datetime.now()}')

        return render_template('resultados.html', tabla=resultados_html)

@app.route('/download')
def download_file():
    # Descargar el archivo de resultados
    return send_file('resultados.xlsx', as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
