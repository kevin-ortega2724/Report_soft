import sqlite3
from pathlib import Path

db_path = "C:/Users/USUARIO/Desktop/ReportSoft/consolidados/datame/gruplac_957.db"

try:
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    print("--- TABLAS EN LA BASE DE DATOS ---")
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
    tables = cursor.fetchall()
    for table in tables:
        print(f"Tabla encontrada: {table[0]}")

    # Intentar buscar la tabla de grupos o categorías para ver la estructura
    for table_name in [t[0] for t in tables]:
        print(f"\n--- Estructura de la tabla: {table_name} ---")
        cursor.execute(f"PRAGMA table_info({table_name})")
        columns = cursor.fetchall()
        for col in columns:
            print(col)

        # Tomar una muestra de 3 filas para entender el contenido
        print(f"--- Muestra de datos de {table_name} ---")
        try:
            cursor.execute(f"SELECT * FROM {table_name} LIMIT 3")
            rows = cursor.fetchall()
            for row in rows:
                print(row)
        except Exception as e:
            print(f"No se pudo leer la tabla {table_name}: {e}")

    conn.close()
except Exception as e:
    print(f"Error al acceder a la base de datos: {e}")
