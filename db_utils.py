# db_utils.py

import pyodbc
import pandas as pd

def obtener_datos_kpi(db_name):
    conexion = pyodbc.connect(
        "DRIVER={SQL Server};"
        "SERVER=192.168.201.12;"
        f"DATABASE={db_name};"
        "UID=sa;"
        "PWD=infinity;"
    )
    cursor = conexion.cursor()
    
    query_kpi = "SELECT TotalRecords, CompleteRecords, CorrectionRecords FROM CNAE_KPI_Audit"
    cursor.execute(query_kpi)
    
    kpi_data = cursor.fetchone()
    
    conexion.close()
    return {
        "TotalRecords": kpi_data[0],
        "CompleteRecords": kpi_data[1],
        "CorrectionRecords": kpi_data[2]
    }

def obtener_datos_clicks(db_name):
    conexion = pyodbc.connect(
        "DRIVER={SQL Server};"
        "SERVER=192.168.201.12;"
        f"DATABASE={db_name};"
        "UID=sa;"
        "PWD=infinity;"
    )
    query_clicks = """
        SELECT CAST(ClickDate AS DATE) AS ClickDate, ButtonName, COUNT(*) AS ClickCount
        FROM LinkedinClickLog
        GROUP BY CAST(ClickDate AS DATE), ButtonName
        ORDER BY ClickDate
    """
    clicks_data = pd.read_sql(query_clicks, conexion)
    conexion.close()
    return clicks_data
