# -*- coding: utf-8 -*-
"""
Módulo para gestionar la cola de envíos offline.
"""
import os
import sqlite3
import json
from datetime import datetime
from typing import List, Dict, Optional

DB_PATH = os.path.join(os.path.dirname(__file__), "factunabo_history.db")

def add_to_queue(xml_content: bytes, num_factura: str, empresa: str, ejercicio: str,
                 cliente_doc: str, api_key: str) -> int:
    """Añade un envío a la cola offline."""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    fecha_creacion = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cursor.execute("""
        INSERT INTO offline_queue 
        (xml_content, num_factura, empresa, ejercicio, cliente_doc, api_key, fecha_creacion, estado)
        VALUES (?, ?, ?, ?, ?, ?, ?, 'PENDIENTE')
    """, (xml_content, num_factura, empresa, ejercicio, cliente_doc, api_key, fecha_creacion))
    
    queue_id = cursor.lastrowid
    conn.commit()
    conn.close()
    return queue_id

def get_pending_items(limit: int = 50) -> List[Dict]:
    """Obtiene items pendientes de la cola."""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("""
        SELECT id, xml_content, num_factura, empresa, ejercicio, cliente_doc, api_key, intentos
        FROM offline_queue
        WHERE estado = 'PENDIENTE'
        ORDER BY fecha_creacion ASC
        LIMIT ?
    """, (limit,))
    
    items = []
    for row in cursor.fetchall():
        items.append({
            "id": row[0],
            "xml_content": row[1],
            "num_factura": row[2],
            "empresa": row[3],
            "ejercicio": row[4],
            "cliente_doc": row[5],
            "api_key": row[6],
            "intentos": row[7]
        })
    
    conn.close()
    return items

def mark_as_sent(queue_id: int):
    """Marca un item como enviado exitosamente."""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("""
        UPDATE offline_queue
        SET estado = 'ENVIADO', ultimo_intento = ?
        WHERE id = ?
    """, (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), queue_id))
    conn.commit()
    conn.close()

def mark_as_failed(queue_id: int, error_msg: str, max_retries: int = 3):
    """Marca un item como fallido o incrementa intentos."""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    cursor.execute("SELECT intentos FROM offline_queue WHERE id = ?", (queue_id,))
    row = cursor.fetchone()
    if row:
        intentos = row[0] + 1
        if intentos >= max_retries:
            estado = "FALLIDO"
        else:
            estado = "PENDIENTE"
        
        cursor.execute("""
            UPDATE offline_queue
            SET estado = ?, intentos = ?, ultimo_intento = ?
            WHERE id = ?
        """, (estado, intentos, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), queue_id))
    
    conn.commit()
    conn.close()

def clear_sent_items():
    """Elimina items enviados de la cola (opcional, para limpieza)."""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM offline_queue WHERE estado = 'ENVIADO'")
    conn.commit()
    conn.close()

def get_queue_stats() -> Dict:
    """Obtiene estadísticas de la cola."""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    cursor.execute("SELECT estado, COUNT(*) FROM offline_queue GROUP BY estado")
    stats = {}
    for row in cursor.fetchall():
        stats[row[0]] = row[1]
    
    conn.close()
    return stats

