from dotenv import load_dotenv
load_dotenv()

import os
import psycopg2
from datetime import datetime
from pytz import timezone

DATABASE_URL = os.getenv("DATABASE_URL")

if not DATABASE_URL:
    raise ValueError("DATABASE_URL no está definido como variable de entorno")

# Cambia esto por el token que quieres probar
TOKEN_PRUEBA = "tokendefinitivo123"

fecha_click = datetime.now(timezone("America/Santiago"))

try:
    with psycopg2.connect(DATABASE_URL) as conn:
        with conn.cursor() as cur:
            # Verifica si existe el envío
            cur.execute("""
                SELECT id, clicks_count, url_destino
                FROM envios_clicks
                WHERE token = %s
            """, (TOKEN_PRUEBA,))
            fila = cur.fetchone()

            if not fila:
                print(f"[X] No se encontró ningún registro con token {TOKEN_PRUEBA}")
            else:
                id_envio, clicks_anteriores, url_destino = fila
                print(f"[✓] Registro encontrado con ID {id_envio}")
                print(f"  URL destino: {url_destino}")
                print(f"  Clicks anteriores: {clicks_anteriores}")

                # Actualizar clicks_count y last_click_at
                cur.execute("""
                    UPDATE envios_clicks
                    SET
                        clicks_count = COALESCE(clicks_count,0) + 1,
                        last_click_at = %s
                    WHERE id = %s
                """, (
                    fecha_click,
                    id_envio
                ))

                print(f"[✓] Click registrado correctamente. Ahora clicks_count es {clicks_anteriores + 1}.")

except Exception as e:
    print("[X] Error al registrar click:", e)
