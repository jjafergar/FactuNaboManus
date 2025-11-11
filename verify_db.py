import sqlite3
import os

DB_PATH = os.path.join(os.path.dirname(__file__), "factunabo_history.db")

if not os.path.exists(DB_PATH):
    print("ERROR: La base de datos 'factunabo_history.db' no existe.")
else:
    try:
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        print("Contenido de la tabla 'envios':")
        cursor.execute("SELECT * FROM envios;")
        rows = cursor.fetchall()

        if not rows:
            print("La tabla 'envios' está vacía.")
        else:
            # Obtener los nombres de las columnas para una mejor visualización
            names = [description[0] for description in cursor.description]
            print(names)
            for row in rows:
                print(row)

        conn.close()
    except Exception as e:
        print(f"Error al leer la base de datos: {e}")
