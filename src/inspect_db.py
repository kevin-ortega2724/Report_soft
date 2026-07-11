import sqlite3
import pandas as pd
from pathlib import Path

db_path = "C:/Users/USUARIO/Desktop/ReportSoft/consolidados/datame/gruplac_957.db"

try:
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    print("--- TABLAS EN LA BASE DE DATOS ---")
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
    tables = cursor.fetchall()
    for table in tables:
        print(table[0])

    # Intentar buscar información de áreas y categorías si existen tablas comunes
    potential_tables = ['grupos', 'categorias', 'areas', 'indicadores', 'resultados']
    for table in potential_tables:
        try:
            cursor.execute(f"SELECT name FROM sqlite_master WHERE type='table' AND name='{table}'")
            if cursor.fetchone():
                print(f"\n--- MUESTRA DE LA TABLA {table} ---")
                df = pd.read_sql_query(f"SELECT * FROM {table} LIMIT 5", conn)
                print(df)
        except Exception as e:
            pass

    conn.close()
except Exception as e:
    print(f"Error al acceder a la base de datos: {e}")
