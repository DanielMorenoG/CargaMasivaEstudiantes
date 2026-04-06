import mysql.connector

conn = mysql.connector.connect(
    host="localhost",
    user="root",
    password="",
    database="notas2026"
)
cursor = conn.cursor()

# Ver cuántos estudiantes hay ANTES
cursor.execute("SELECT COUNT(*) FROM estudiantes")
print("Estudiantes ANTES:", cursor.fetchone()[0])

# Intentar insertar
try:
    cursor.execute(
        "INSERT INTO estudiantes (Nombre, Edad, Carrera, nota1, nota2, nota3, Promedio, `Desempeño`) VALUES (%s,%s,%s,%s,%s,%s,%s,%s)",
        ("Lobo2_TEST", 20, "Fisica", 3.0, 4.0, 5.0, 4.0, "Bueno")
    )
    conn.commit()
    print("INSERT ejecutado y commit hecho")
except Exception as e:
    print("ERROR en INSERT:", e)

# Ver cuántos hay DESPUÉS
cursor.execute("SELECT COUNT(*) FROM estudiantes")
print("Estudiantes DESPUÉS:", cursor.fetchone()[0])

# Ver último registro
cursor.execute("SELECT * FROM estudiantes ORDER BY id DESC LIMIT 3")
for row in cursor.fetchall():
    print("Último registro:", row)

cursor.close()
conn.close()