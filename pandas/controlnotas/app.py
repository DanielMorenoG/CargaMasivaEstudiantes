from flask import Flask, render_template, request, redirect, session, make_response
from functools import wraps
from database import obtenerusuarios
from dashprincipal import creartablero

app = Flask(__name__)

# CLAVE NECESARIA PARA USAR SESSION
app.secret_key = "40414732"

# crear dashboard
creartablero(app)


# ── Protección global: intercepta CUALQUIER ruta /dashprincipal/... ───────────
# Dash registra sus rutas internas en /dashprincipal/*, Flask no las ve con @route.
# before_request corre antes de CADA petición, así que sí las atrapa.
@app.before_request
def verificar_sesion_global():
    ruta = request.path
    if ruta.startswith("/dashprincipal"):
        if "username" not in session:
            return redirect("/")


# ── Headers anti-caché para TODAS las respuestas ──────────────────────────────
@app.after_request
def agregar_headers_cache(response):
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    response.headers["Pragma"]        = "no-cache"
    response.headers["Expires"]       = "0"
    return response


# ── LOGIN ─────────────────────────────────────────────────────────────────────
@app.route("/", methods=["GET", "POST"])
def login():
    # Si ya tiene sesión activa, redirigir al dashboard
    if "username" in session:
        return redirect("/dashprincipal")

    error = None

    if request.method == "POST":
        username = request.form["username"]
        password = request.form["password"]

        usuario = obtenerusuarios(username)

        if usuario:
            if usuario["password"] == password:
                session["username"] = usuario["username"]
                session["rol"]      = usuario["rol"]
                return redirect("/dashprincipal")
            else:
                error = "Contraseña incorrecta. Inténtalo de nuevo."
        else:
            error = "El usuario no existe."

    return render_template("login.html", error=error)


# ── DASHBOARD (wrapper Flask que sirve el HTML con el iframe) ─────────────────
@app.route("/dashprincipal")
def dashprinci():
    if "username" not in session:
        return redirect("/")
    return render_template("dashprinci.html", usuario=session["username"])


# ── CERRAR SESIÓN ─────────────────────────────────────────────────────────────
@app.route("/logout")
def logout():
    session.clear()
    resp = make_response(redirect("/"))
    resp.delete_cookie("session")
    return resp


if __name__ == "__main__":
    app.run(debug=True)