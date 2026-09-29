from flask import Flask, jsonify, render_template, request
import pandas as pd
from datetime import datetime
import os
import tempfile

from backend.data_loader import load_data


app = Flask(__name__)


# ========================================================
# DATOS CARGADOS POR EL USUARIO
# ========================================================

df = None
df_normalizado = None
df_subsistema_fmr = None
df_fmr_po = None
df_fmr_detalle = None
df_subsistema_fmr_po = None
df_3week_subsistema = None
datos_cargados = False


def cargar_datos(excel_file=None):

    global df
    global df_normalizado
    global df_subsistema_fmr
    global df_fmr_po
    global df_fmr_detalle
    global df_subsistema_fmr_po
    global df_3week_subsistema
    global datos_cargados

    (
        df,
        df_normalizado,
        df_subsistema_fmr,
        df_fmr_po,
        df_fmr_detalle,
        df_subsistema_fmr_po,
        df_3week_subsistema
    ) = load_data() if excel_file is None else load_data(excel_file)

    datos_cargados = True


def exigir_datos_cargados():

    if not datos_cargados:

        return jsonify({
            "error": "Debe cargar el archivo Excel antes de consultar los datos."
        }), 409

    return None


# ========================================================
# FUNCIONES AUXILIARES
# ========================================================

def limpiar_texto(valor):

    if pd.isna(valor):
        return ""

    return str(valor).strip()


def tiene_dato(valor):

    if pd.isna(valor):
        return False

    return str(valor).strip() != ""


# ========================================================
# CONVERTIR ETA A FECHAS
# ========================================================

def obtener_fechas_eta(valor):

    if pd.isna(valor):
        return []

    if isinstance(valor, pd.Timestamp):

        return [
            valor.date()
        ]

    if isinstance(valor, datetime):

        return [
            valor.date()
        ]

    texto = str(valor).strip()

    if not texto:
        return []

    texto = texto.replace(
        "\n",
        ","
    )

    texto = texto.replace(
        ";",
        ","
    )

    partes = [
        parte.strip()
        for parte in texto.split(",")
        if parte.strip()
    ]

    fechas = []

    for parte in partes:

        fecha = pd.to_datetime(
            parte,
            errors="coerce",
            dayfirst=True
        )

        if pd.notna(fecha):

            fechas.append(
                fecha.date()
            )

    return fechas


# ========================================================
# ESTADO DE UNA PO
#
# REGLAS:
#
# 1. ETA futura              -> AMARILLO
# 2. Sin ETA                 -> ROJO
# 3. ETA pasada + ON SITE    -> VERDE
# 4. ETA pasada + sin ON SITE-> ROJO
#
# Si se proporciona SubSistema:
#     filtra por SubSistema + FMR + PO
#
# Si NO se proporciona:
#     busca FMR + PO en todos los SubSistemas
# ========================================================

def calcular_estado_po(
    fmr,
    po,
    subsistema=None
):

    datos = df_fmr_detalle[
        (
            df_fmr_detalle["FMR"].astype(str)
            == str(fmr)
        )
        &
        (
            df_fmr_detalle["PO"].astype(str)
            == str(po)
        )
    ].copy()


    if subsistema:

        datos = datos[
            datos["SubSistema"].astype(str)
            == str(subsistema)
        ]


    # ----------------------------------------------------
    # SIN DATOS DE PROCUREMENT
    # ----------------------------------------------------

    if datos.empty:

        return {
            "color": "red",
            "estado": "sin_eta"
        }


    hoy = datetime.now().date()

    tiene_eta = False
    tiene_eta_futura = False
    tiene_on_site = False


    for _, fila in datos.iterrows():

        # -----------------------------------------------
        # ETA
        # -----------------------------------------------

        fechas = obtener_fechas_eta(
            fila["ETA"]
        )


        if fechas:

            tiene_eta = True

            if any(
                fecha > hoy
                for fecha in fechas
            ):

                tiene_eta_futura = True


        # -----------------------------------------------
        # ON SITE
        # -----------------------------------------------

        if tiene_dato(
            fila["ON SITE"]
        ):

            tiene_on_site = True


    # ====================================================
    # 1. ETA FUTURA
    # ====================================================

    if tiene_eta_futura:

        return {
            "color": "yellow",
            "estado": "futura"
        }


    # ====================================================
    # 2. SIN ETA
    # ====================================================

    if not tiene_eta:

        return {
            "color": "red",
            "estado": "sin_eta"
        }


    # ====================================================
    # 3. ETA PASADA + ON SITE
    # ====================================================

    if tiene_on_site:

        return {
            "color": "green",
            "estado": "pasada_on_site"
        }


    # ====================================================
    # 4. ETA  + SIN ON SITE
    # ====================================================

    return {
        "color": "yellow",
        "estado": "eta_sin_on_site"
    }


# ========================================================
# ESTADO DE FMR
# ========================================================

def calcular_estado_fmr(
    fmr,
    subsistema=None
):

    datos = df_subsistema_fmr_po[
        df_subsistema_fmr_po["FMR"].astype(str)
        == str(fmr)
    ].copy()


    if subsistema:

        datos = datos[
            datos["SubSistema"].astype(str)
            == str(subsistema)
        ]


    if datos.empty:

        return {
            "color": "red",
            "estado": "sin_eta"
        }


    estados = []


    for _, fila in datos.iterrows():

        estado = calcular_estado_po(
            fmr=fmr,
            po=fila["PO"],
            subsistema=fila["SubSistema"]
        )

        estados.append(
            estado
        )


    # ----------------------------------------------------
    # SI ALGUNA PO ES FUTURA
    # ----------------------------------------------------

    # ----------------------------------------------------
    # COMBINACIÓN DE COLORES
    # ----------------------------------------------------

    colores = {
        estado["color"]
        for estado in estados
    }


    if len(colores) >= 2:

        return {
            "color": "orange",
            "estado": "combinada"
        }


    # ----------------------------------------------------
    # TODAS LAS PO SON AMARILLAS
    # ----------------------------------------------------

    if "yellow" in colores:

        return {
            "color": "yellow",
            "estado": "futura"
        }


    # ----------------------------------------------------
    # TODAS LAS PO SON VERDES
    # ----------------------------------------------------

    if "green" in colores:

        return {
            "color": "green",
            "estado": "pasada_on_site"
        }

    # ----------------------------------------------------
    # RESTO
    # ----------------------------------------------------

    return {
        "color": "red",
        "estado": "sin_eta"
    }


# ========================================================
# FORMATEAR ETA
# ========================================================

def formatear_eta(valor):

    if pd.isna(valor):

        return ""


    fechas = obtener_fechas_eta(
        valor
    )


    if not fechas:

        return str(valor)


    return ", ".join(
        fecha.strftime("%d-%b-%Y")
        for fecha in fechas
    )


# ========================================================
# PÁGINA PRINCIPAL
# ========================================================

@app.route("/")
def home():

    return render_template(
        "index.html"
    )


@app.route("/dashboard")
def dashboard():

    return render_template(
        "dashboard.html"
    )


@app.route("/api/estado")
def api_estado():

    return jsonify({
        "datos_cargados": datos_cargados,
        "archivo": None
    })


@app.route("/api/cargar-datos", methods=["POST"])
def api_cargar_datos():

    archivo = request.files.get("archivo")

    if archivo is None or not archivo.filename:

        return jsonify({
            "error": "Seleccione un archivo Excel para continuar."
        }), 400

    extension = os.path.splitext(archivo.filename)[1].lower()

    if extension not in {".xlsx", ".xlsm"}:

        return jsonify({
            "error": "El archivo debe tener formato .xlsx o .xlsm."
        }), 400

    archivo_temporal = tempfile.NamedTemporaryFile(
        suffix=extension,
        delete=False
    )

    archivo_temporal.close()

    try:

        archivo.save(archivo_temporal.name)
        cargar_datos(archivo_temporal.name)

        return jsonify({
            "datos_cargados": True,
            "archivo": archivo.filename
        })

    except (FileNotFoundError, ValueError, KeyError, OSError) as error:

        return jsonify({
            "error": f"No se pudo cargar el archivo Excel: {error}"
        }), 500

    finally:

        if os.path.exists(archivo_temporal.name):
            os.remove(archivo_temporal.name)


# ========================================================
# API 3WEEK
#
# SIN FECHA:
# devuelve las fechas disponibles
#
# CON FECHA:
# devuelve los SubSistemas de esa fecha
# ========================================================

@app.route("/api/3week")
def api_3week():

    respuesta = exigir_datos_cargados()

    if respuesta is not None:

        return respuesta

    fecha = (
        request.args
        .get("fecha", "")
        .strip()
    )


    # ====================================================
    # CASO 1:
    # LISTAR FECHAS
    # ====================================================

    if not fecha:

        fechas = (
            df_3week_subsistema["3week"]
            .dropna()
            .drop_duplicates()
            .sort_values()
            .dt.strftime("%d/%m/%Y")
            .tolist()
        )


        print(
            "3WEEK DISPONIBLES:",
            fechas
        )


        return jsonify(
            fechas
        )


    # ====================================================
    # CASO 2:
    # SUBSISTEMAS SEGÚN FECHA
    # ====================================================

    print(
        "3WEEK SELECCIONADA:",
        fecha
    )


    datos = (
        df_3week_subsistema
        .copy()
    )


    datos["fecha_texto"] = (
        datos["3week"]
        .dt.strftime("%d/%m/%Y")
    )


    datos = datos[
        datos["fecha_texto"]
        == fecha
    ]


    subsistemas = (
        datos["SubSistema"]
        .dropna()
        .astype(str)
        .str.strip()
        .drop_duplicates()
        .sort_values()
        .tolist()
    )


    print(
        "SUBSISTEMAS ENCONTRADOS:",
        subsistemas
    )


    return jsonify({

        "fecha": fecha,

        "subsistemas":
            subsistemas,

        "total_subsistemas":
            len(subsistemas)

    })


# ========================================================
# API TODOS LOS SUBSISTEMAS
# ========================================================

@app.route("/api/subsistemas")
def api_subsistemas():

    respuesta = exigir_datos_cargados()

    if respuesta is not None:

        return respuesta

    subsistemas = (

        df_subsistema_fmr[
            "SubSistema"
        ]

        .dropna()

        .astype(str)

        .str.strip()

        .drop_duplicates()

        .sort_values()

        .tolist()
    )


    return jsonify(
        subsistemas
    )


# ========================================================
# API FMR DE UN SUBSISTEMA
# ========================================================

@app.route(
    "/api/subsistema/<path:subsistema>"
)
def api_fmr_por_subsistema(
    subsistema
):

    respuesta = exigir_datos_cargados()

    if respuesta is not None:

        return respuesta

    print(
        "SUBSISTEMA SELECCIONADO:",
        subsistema
    )


    datos = df_subsistema_fmr[
        df_subsistema_fmr[
            "SubSistema"
        ].astype(str)
        == str(subsistema)
    ].copy()


    fmrs = (
        datos["FMR"]
        .dropna()
        .astype(str)
        .str.strip()
        .drop_duplicates()
        .sort_values()
        .tolist()
    )


    resultado = []


    for fmr in fmrs:

        estado = calcular_estado_fmr(
            fmr=fmr,
            subsistema=subsistema
        )


        resultado.append({

            "fmr":
                fmr,

            "color":
                estado["color"],

            "estado":
                estado["estado"]

        })


    return jsonify({

        "subsistema":
            subsistema,

        "fmr":
            fmrs,

        "fmrs":
            resultado,

        "total_fmr":
            len(fmrs)

    })


# ========================================================
# API TODOS LOS FMR
# ========================================================

@app.route("/api/fmrs")
def api_fmrs():

    respuesta = exigir_datos_cargados()

    if respuesta is not None:

        return respuesta

    fmrs = (

        df_subsistema_fmr[
            "FMR"
        ]

        .dropna()

        .astype(str)

        .str.strip()

        .drop_duplicates()

        .sort_values()

        .tolist()
    )


    return jsonify(
        fmrs
    )


# ========================================================
# API PRINCIPAL FMR
#
# ESTA ES LA RUTA QUE RECUPERAMOS
# DEL FLUJO QUE YA FUNCIONABA
#
# /api/fmr/747667
#
# /api/fmr/747667?subsistema=36045-ROT-005
# ========================================================

@app.route(
    "/api/fmr/<path:fmr>"
)
def api_detalle_fmr(
    fmr
):

    respuesta = exigir_datos_cargados()

    if respuesta is not None:

        return respuesta

    fmr = str(fmr).strip()


    subsistema = (
        request.args
        .get(
            "subsistema",
            ""
        )
        .strip()
    )


    print(
        "FMR SOLICITADO:",
        fmr
    )


    print(
        "SUBSISTEMA:",
        subsistema
    )


    # ====================================================
    # VALIDAR FMR
    # ====================================================

    datos_fmr = df_subsistema_fmr[
        df_subsistema_fmr[
            "FMR"
        ].astype(str)
        == fmr
    ].copy()


    if datos_fmr.empty:

        return jsonify({

            "encontrado":
                False,

            "fmr":
                fmr,

            "po":
                [],

            "pos":
                [],

            "procurement":
                [],

            "subsistemas":
                []

        }), 404


    # ====================================================
    # SUBSISTEMAS DEL FMR
    # ====================================================

    subsistemas = (
        datos_fmr["SubSistema"]
        .dropna()
        .astype(str)
        .str.strip()
        .drop_duplicates()
        .sort_values()
        .tolist()
    )


    # ====================================================
    # PO
    # ====================================================

    datos_po = df_subsistema_fmr_po[
        df_subsistema_fmr_po[
            "FMR"
        ].astype(str)
        == fmr
    ].copy()


    if subsistema:

        datos_po = datos_po[
            datos_po[
                "SubSistema"
            ].astype(str)
            == subsistema
        ]


    pos = (
        datos_po["PO"]
        .dropna()
        .astype(str)
        .str.strip()
        .drop_duplicates()
        .sort_values()
        .tolist()
    )


    # ====================================================
    # ESTADO DE CADA PO
    # ====================================================

    po_detalle = []


    for po in pos:

        estado = calcular_estado_po(
            fmr=fmr,
            po=po,
            subsistema=subsistema
            if subsistema
            else None
        )


        po_detalle.append({

            "po":
                po,

            "color":
                estado["color"],

            "estado":
                estado["estado"]

        })


    # ====================================================
    # PROCUREMENT
    # ====================================================

    datos_procurement = df_fmr_detalle[
        df_fmr_detalle[
            "FMR"
        ].astype(str)
        == fmr
    ].copy()


    if subsistema:

        datos_procurement = datos_procurement[
            datos_procurement[
                "SubSistema"
            ].astype(str)
            == subsistema
        ]


    procurement = []


    for _, fila in datos_procurement.iterrows():

        procurement.append({

            "po":
                limpiar_texto(
                    fila["PO"]
                ),

            "status":
                limpiar_texto(
                    fila["STATUS"]
                ),

            "buyer":
                limpiar_texto(
                    fila["BUYER"]
                ),

            "expeditor":
                limpiar_texto(
                    fila["EXPEDITOR"]
                ),

            "eta":
                formatear_eta(
                    fila["ETA"]
                ),

            "on_site":
                limpiar_texto(
                    fila["ON SITE"]
                ),

            "entregado":
                limpiar_texto(
                    fila["ENTREGADO"]
                ),

            "descripcion":
                limpiar_texto(
                    fila["DESCRIPCION"]
                )

        })


    # ====================================================
    # DESCRIPCIÓN
    # ====================================================

    descripcion = ""


    for registro in procurement:

        if registro["descripcion"]:

            descripcion = (
                registro["descripcion"]
            )

            break


    # ====================================================
    # ESTADO FMR
    # ====================================================

    estado_fmr = calcular_estado_fmr(
        fmr=fmr,
        subsistema=subsistema
        if subsistema
        else None
    )


    # ====================================================
    # RESPUESTA
    # ====================================================

    return jsonify({

        "encontrado":
            True,

        "fmr":
            fmr,

        "subsistema":
            subsistema,

        "subsistemas":
            subsistemas,

        "po":
            pos,

        "pos":
            po_detalle,

        "total_po":
            len(pos),

        "descripcion":
            descripcion,

        "procurement":
            procurement,

        "total_procurement":
            len(procurement),

        "color":
            estado_fmr["color"],

        "estado":
            estado_fmr["estado"]

    })


# ========================================================
# EJECUTAR
# ========================================================

if __name__ == "__main__":

    app.run(
        debug=True
    )