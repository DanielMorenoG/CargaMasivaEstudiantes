import mysql.connector
import pandas as pd

def conectar():
    return mysql.connector.connect(
        host="localhost",
        user="root",
        password="",
        database="notas2026"
    )

# ── OBTENER UN USUARIO ────────────────────────────────────────────────────────
def obtenerusuarios(username):
    conn = conectar()
    cursor = conn.cursor(dictionary=True)
    cursor.execute("SELECT * FROM usuarios WHERE username=%s", (username,))
    usuario = cursor.fetchone()
    cursor.close()
    conn.close()
    return usuario

# ── OBTENER ESTUDIANTES ───────────────────────────────────────────────────────
def obtenerestudiantes():
    conn = conectar()
    cursor = conn.cursor(dictionary=True)
    cursor.execute("SELECT * FROM estudiantes")
    rows = cursor.fetchall()
    cursor.close()
    conn.close()
    return pd.DataFrame(rows)

# ── REGISTRAR ESTUDIANTE ──────────────────────────────────────────────────────
def insertar_estudiante(nombre, edad, carrera, nota1, nota2, nota3, promedio, desempenio):
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute(
        """INSERT INTO estudiantes
           (Nombre, Edad, Carrera, nota1, nota2, nota3, Promedio, `Desempeño`)
           VALUES (%s, %s, %s, %s, %s, %s, %s, %s)""",
        (nombre, int(edad), carrera,
         float(nota1), float(nota2), float(nota3),
         float(promedio), desempenio)
    )
    conn.commit()
    cursor.close()
    conn.close()

# ── OBTENER CLAVES COMPUESTAS EXISTENTES ──────────────────────────────────────
def obtener_claves_existentes():
    """
    Devuelve un set de tuplas (nombre_lower, carrera_lower).

    Clave única: Nombre + Carrera (según requisito del taller).
    Bloquea que el mismo nombre se registre dos veces en la misma carrera,
    independientemente de la edad.
    """
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("SELECT Nombre, Carrera FROM estudiantes")
    rows = cursor.fetchall()
    cursor.close()
    conn.close()
    return {
        (str(r[0]).strip().lower(), str(r[1]).strip().lower())
        for r in rows
    }

def existe_estudiante(nombre, edad, carrera):
    """Comprueba si ya existe un registro con mismo Nombre + Carrera."""
    claves = obtener_claves_existentes()
    return (str(nombre).strip().lower(), str(carrera).strip().lower()) in claves

# ── CARGA MASIVA ──────────────────────────────────────────────────────────────
def insertar_masivo(filas):
    """
    filas: lista de tuplas
    (nombre, edad, carrera, nota1, nota2, nota3, promedio, desempenio)

    Filtra duplicados usando la clave compuesta Nombre+Edad+Carrera tanto
    contra la BD como dentro del propio lote.

    Retorna:
        {"insertados": int, "duplicados": list[str]}
    """
    if not filas:
        return {"insertados": 0, "duplicados": []}

    existentes = obtener_claves_existentes()
    a_insertar = []
    duplicados = []

    for fila in filas:
        nombre, edad, carrera = fila[0], int(fila[1]), fila[2]
        clave = (str(nombre).strip().lower(), str(carrera).strip().lower())

        if clave in existentes:
            duplicados.append(f"{nombre} ({carrera})")
        else:
            a_insertar.append(fila)
            existentes.add(clave)       # captura duplicados dentro del mismo lote

    if a_insertar:
        conn = conectar()
        cursor = conn.cursor()
        cursor.executemany(
            """INSERT INTO estudiantes
               (Nombre, Edad, Carrera, nota1, nota2, nota3, Promedio, `Desempeño`)
               VALUES (%s, %s, %s, %s, %s, %s, %s, %s)""",
            a_insertar,
        )
        conn.commit()
        cursor.close()
        conn.close()

    return {"insertados": len(a_insertar), "duplicados": duplicados}


if __name__ == "__main__":
    try:
        conn = conectar()
        print("✓ Conexión exitosa")
        conn.close()
        insertar_estudiante("TEST_DEBUG", 20, "Fisica", 3.0, 4.0, 5.0, 4.0, "Bueno")
        print("✓ INSERT OK")
    except Exception as e:
        print(f"✗ Error: {e}")