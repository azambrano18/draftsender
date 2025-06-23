# test_registro_envio.py

import os
import psycopg2
import dns.resolver
from datetime import datetime
from pytz import timezone

def registrar_envio(remitente: str, destinatario: str, url_destino: str, tipo_envio: str, token: str):
    dominio = destinatario.split("@")[-1].lower()
    try:
        servidor = sorted(dns.resolver.resolve(dominio, 'MX'), key=lambda r: r.preference)[0].exchange.to_text().lower()
    except Exception:
        servidor = "desconocido"

    fecha_envio = datetime.now(timezone("America/Santiago"))
    DATABASE_URL = os.getenv("DATABASE_URL")

    try:
        with psycopg2.connect(DATABASE_URL) as conn:
            with conn.cursor() as cur:
                cur.execute("""
                    SELECT id FROM clicks
                    WHERE remitente = %s AND destinatario = %s AND token = %s
                """, (remitente, destinatario, token))
                existe = cur.fetchone()

                if existe:
                    cur.execute("""
                        UPDATE clicks SET
                            fecha_envio = %s,
                            tipo_envio = %s,
                            dominio_destinatario = %s,
                            servidor_correo = %s,
                            url_destino = %s
                        WHERE id = %s
                    """, (
                        fecha_envio, tipo_envio, dominio, servidor,
                        url_destino, existe[0]
                    ))
                    print(f"[✓] Registro actualizado para: {remitente} → {destinatario}")
                else:
                    cur.execute("""
                        INSERT INTO clicks (
                            remitente, fecha_envio, tipo_envio, destinatario,
                            click_apertura, cantidad_click, dominio_destinatario,
                            servidor_correo, url_destino, navegador, so, pais,
                            ip_public, token, user_agent
                        ) VALUES (%s, %s, %s, %s, NULL, 0, %s, %s, %s, NULL, NULL, NULL, NULL, %s, NULL)
                    """, (
                        remitente, fecha_envio, tipo_envio, destinatario,
                        dominio, servidor, url_destino, token
                    ))
                    print(f"[✓] Nuevo envío registrado: {remitente} → {destinatario}")
    except Exception as e:
        print(f"[X] ERROR: {e}")

# ==== Datos simulados para prueba ====

remitente = "mlobos@dhumanlab.com"
destinatario = "aazs90@gmail.com"
url_destino = "https://tracking-mr-XXXX.onrender.com/click?from=mlobos@dhumanlab.com&to=aazs90@gmail.com&url=https://ejemplo.com&token=abc123test"
token = "abc123test"
tipo_envio = "envio1"

registrar_envio(remitente, destinatario, url_destino, tipo_envio, token)
