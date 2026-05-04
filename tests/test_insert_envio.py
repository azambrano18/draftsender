import os
from dotenv import load_dotenv
import psycopg2
from datetime import datetime
from pytz import timezone

load_dotenv()

DATABASE_URL = os.getenv("DATABASE_URL")

if not DATABASE_URL:
    raise Exception("DATABASE_URL no está definido")

try:
    with psycopg2.connect(DATABASE_URL) as conn:
        with conn.cursor() as cur:
            # Mostrar información de conexión
            cur.execute("SELECT current_database(), current_user, inet_server_addr()")
            db_info = cur.fetchone()
            print("Conectado a base:", db_info[0])
            print("Usuario:", db_info[1])
            print("Servidor:", db_info[2])

            # Insert de prueba
            cur.execute("""
                INSERT INTO envios_clicks (
                    remitente,
                    destinatario,
                    metodo_envio,
                    asunto,
                    cuerpo,
                    fecha_envio,
                    token,
                    url_destino
                )
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
                RETURNING id
            """, (
                "prueba@dominio.com",
                "destino@cliente.com",
                "Envio1",
                "Test Insert Definitivo",
                "<p>Contenido prueba definitiva</p>",
                datetime.now(timezone("America/Santiago")),
                "tokendefinitivo123",
                "https://example.com"
            ))
            new_id = cur.fetchone()[0]
            conn.commit()
            print(f"[✓] Registro insertado con id: {new_id}")

except Exception as e:
    import traceback
    traceback.print_exc()
    print("[X] Error al insertar:", e)