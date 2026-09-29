import pandas as pd
from pathlib import Path


# ========================================================
# CONFIGURACIÓN
# ========================================================

BASE_DIR = Path(__file__).resolve().parent.parent

EXCEL_FILE = (
    BASE_DIR
    / "data"
    / "FMRs_Excepciones_27.08.26_Oeste.xlsx"
)


# ========================================================
# CARGAR DATOS
# ========================================================

def load_data(excel_file=EXCEL_FILE):

    # ----------------------------------------------------
    # HOJA PRINCIPAL
    # ----------------------------------------------------

    df = pd.read_excel(
        excel_file,
        sheet_name=0
    )

    # ----------------------------------------------------
    # HOJA FMR_NORMALIZADO
    # ----------------------------------------------------

    df_normalizado = pd.read_excel(
        excel_file,
        sheet_name="FMR_Normalizado"
    )

    # ----------------------------------------------------
    # LIMPIAR NOMBRES DE COLUMNAS
    # ----------------------------------------------------

    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
    )

    df_normalizado.columns = (
        df_normalizado.columns
        .astype(str)
        .str.strip()
    )


    # ====================================================
    # TABLA SUBSISTEMA ↔ FMR
    # ====================================================

    # ============================================
    # 3WEEK ↔ SUBSISTEMA
    # ============================================

    df_3week_subsistema = df_normalizado[
        ["3week", "SubSistema"]
    ].copy()

    # Convertir 3week a fecha
    df_3week_subsistema["3week"] = pd.to_datetime(
        df_3week_subsistema["3week"],
        errors="coerce",
        dayfirst=True
    )

    # Limpiar SubSistema
    df_3week_subsistema["SubSistema"] = (
        df_3week_subsistema["SubSistema"]
        .astype("string")
        .str.strip()
    )

    # Eliminar filas sin fecha o sin SubSistema
    df_3week_subsistema = df_3week_subsistema.dropna(
        subset=["3week", "SubSistema"]
    )

    # Eliminar duplicados
    df_3week_subsistema = (
        df_3week_subsistema
        .drop_duplicates()
        .reset_index(drop=True)
    )
    
    df_subsistema_fmr = df_normalizado[
        [
            "SubSistema",
            "FMR"
        ]
    ].copy()

    df_subsistema_fmr["SubSistema"] = (
        df_subsistema_fmr["SubSistema"]
        .astype("string")
        .str.strip()
    )

    df_subsistema_fmr["FMR"] = (
        df_subsistema_fmr["FMR"]
        .astype("string")
        .str.strip()
    )

    df_subsistema_fmr = (
        df_subsistema_fmr
        .dropna(
            subset=[
                "SubSistema",
                "FMR"
            ]
        )
    )

    df_subsistema_fmr = df_subsistema_fmr[
        ~df_subsistema_fmr["FMR"].str.contains(
            "…",
            regex=False,
            na=False
        )
    ]

    df_subsistema_fmr = (
        df_subsistema_fmr
        .drop_duplicates()
        .reset_index(drop=True)
    )


    # ====================================================
    # TABLA SUBSISTEMA ↔ FMR ↔ PO
    # ====================================================
    #
    # ESTA TABLA ES MUY IMPORTANTE.
    #
    # Las PO se determinan independientemente de:
    #
    # - DESCRIPCION
    # - STATUS
    # - BUYER
    # - EXPEDITOR
    # - ETA
    # - ON SITE
    # - ENTREGADO
    #
    # Por lo tanto, una PO seguirá existiendo aunque
    # todos esos campos estén vacíos.
    #
    # ====================================================

    df_subsistema_fmr_po = df_normalizado[
        [
            "SubSistema",
            "FMR",
            "PO"
        ]
    ].copy()


    # ----------------------------------------------------
    # Limpiar SubSistema
    # ----------------------------------------------------

    df_subsistema_fmr_po["SubSistema"] = (
        df_subsistema_fmr_po["SubSistema"]
        .astype("string")
        .str.strip()
    )


    # ----------------------------------------------------
    # Limpiar FMR
    # ----------------------------------------------------

    df_subsistema_fmr_po["FMR"] = (
        df_subsistema_fmr_po["FMR"]
        .astype("string")
        .str.strip()
    )


    # ----------------------------------------------------
    # Limpiar PO
    # ----------------------------------------------------

    df_subsistema_fmr_po["PO"] = (
        df_subsistema_fmr_po["PO"]
        .astype("string")
        .str.strip()
    )


    # ----------------------------------------------------
    # Eliminar FMR inválido
    # ----------------------------------------------------

    df_subsistema_fmr_po = (
        df_subsistema_fmr_po
        .dropna(
            subset=[
                "SubSistema",
                "FMR"
            ]
        )
    )

    df_subsistema_fmr_po = (
        df_subsistema_fmr_po[
            ~df_subsistema_fmr_po["FMR"].str.contains(
                "…",
                regex=False,
                na=False
            )
        ]
    )


    # ----------------------------------------------------
    # Eliminar PO vacías
    #
    # IMPORTANTE:
    # Una PO vacía no es una PO.
    #
    # Pero una PO existente con DESCRIPCION vacía sí
    # debe mantenerse.
    # ----------------------------------------------------

    df_subsistema_fmr_po = (
        df_subsistema_fmr_po
        .dropna(
            subset=["PO"]
        )
    )


    # ----------------------------------------------------
    # Separar múltiples PO de una misma celda
    # ----------------------------------------------------

    df_subsistema_fmr_po["PO"] = (
        df_subsistema_fmr_po["PO"]
        .str.split("\n")
    )

    df_subsistema_fmr_po = (
        df_subsistema_fmr_po
        .explode("PO")
        .reset_index(drop=True)
    )


    # ----------------------------------------------------
    # Limpiar PO nuevamente
    # ----------------------------------------------------

    df_subsistema_fmr_po["PO"] = (
        df_subsistema_fmr_po["PO"]
        .astype("string")
        .str.strip()
    )


    # ----------------------------------------------------
    # Eliminar PO inválidas
    # ----------------------------------------------------

    df_subsistema_fmr_po = (
        df_subsistema_fmr_po[
            ~df_subsistema_fmr_po["PO"].str.contains(
                "…",
                regex=False,
                na=False
            )
        ]
    )


    # ----------------------------------------------------
    # Eliminar duplicados
    # ----------------------------------------------------

    df_subsistema_fmr_po = (
        df_subsistema_fmr_po
        .drop_duplicates()
        .reset_index(drop=True)
    )


    # ====================================================
    # TABLA FMR ↔ PO GENERAL
    # ====================================================

    df_fmr_po = (
        df_subsistema_fmr_po[
            [
                "FMR",
                "PO"
            ]
        ]
        .drop_duplicates()
        .reset_index(drop=True)
    )


    # ====================================================
    # DETALLE DE PROCUREMENT
    # ====================================================

    df_fmr_detalle = df_normalizado[
        [
            "SubSistema",
            "FMR",
            "PO",
            "DESCRIPCION",
            "STATUS",
            "BUYER",
            "EXPEDITOR",
            "ETA",
            "ON SITE",
            "ENTREGADO"
        ]
    ].copy()


    # ----------------------------------------------------
    # LIMPIAR CAMPOS
    # ----------------------------------------------------

    df_fmr_detalle["SubSistema"] = (
        df_fmr_detalle["SubSistema"]
        .astype("string")
        .str.strip()
    )

    df_fmr_detalle["FMR"] = (
        df_fmr_detalle["FMR"]
        .astype("string")
        .str.strip()
    )

    df_fmr_detalle["PO"] = (
        df_fmr_detalle["PO"]
        .astype("string")
        .str.strip()
    )


    # ----------------------------------------------------
    # FMR válido
    # ----------------------------------------------------

    df_fmr_detalle = (
        df_fmr_detalle
        .dropna(
            subset=["FMR"]
        )
    )

    df_fmr_detalle = (
        df_fmr_detalle[
            ~df_fmr_detalle["FMR"].str.contains(
                "…",
                regex=False,
                na=False
            )
        ]
    )


    # ====================================================
    # NORMALIZAR PO
    # ====================================================

    df_fmr_detalle["PO"] = (
        df_fmr_detalle["PO"]
        .str.split("\n")
    )

    df_fmr_detalle = (
        df_fmr_detalle
        .explode("PO")
        .reset_index(drop=True)
    )

    df_fmr_detalle["PO"] = (
        df_fmr_detalle["PO"]
        .astype("string")
        .str.strip()
    )


    # ====================================================
    # ELIMINAR PO INVÁLIDAS
    # ====================================================

    df_fmr_detalle = (
        df_fmr_detalle[
            ~df_fmr_detalle["PO"].str.contains(
                "…",
                regex=False,
                na=False
            )
        ]
    )


    # ====================================================
    # STATUS
    # ====================================================

    status_orden = {

        "a. FMR Received": 1,

        "b. RFQ Issued": 2,

        "c. TEA Requested": 3,

        "f. Bid Tab Issued to MYSRL": 4,

        "h. FPO to Supplier Signature": 5,

        "k. Awarded": 6
    }

    df_fmr_detalle["_status_orden"] = (
        df_fmr_detalle["STATUS"]
        .map(status_orden)
        .fillna(0)
    )


    # ====================================================
    # IMPORTANTE
    # ====================================================
    #
    # YA NO FILTRAMOS POR DESCRIPCION.
    #
    # Una fila puede tener:
    #
    # PO       = 783376
    # STATUS   = k. Awarded
    # BUYER    = Pavel Ponce
    # DESCRIPCION = NaN
    #
    # Y DEBE SEGUIR EXISTIENDO.
    #
    # ====================================================

    columnas_procurement = [

        "STATUS",

        "BUYER",

        "EXPEDITOR",

        "ETA",

        "ON SITE",

        "ENTREGADO",

        "DESCRIPCION"
    ]


    # ----------------------------------------------------
    # Solo eliminamos filas que no tengan absolutamente
    # ninguna información de Procurement.
    #
    # Pero NO eliminamos una PO de la tabla de PO.
    # ----------------------------------------------------

    df_fmr_detalle = df_fmr_detalle[
        df_fmr_detalle[
            columnas_procurement
        ]
        .notna()
        .any(axis=1)
        |
        df_fmr_detalle["PO"].notna()
    ].copy()


    # ====================================================
    # DUPLICADOS
    # ====================================================

    df_fmr_detalle = (
        df_fmr_detalle
        .drop_duplicates(
            subset=[
                "SubSistema",
                "FMR",
                "PO",
                "DESCRIPCION",
                "STATUS",
                "BUYER",
                "EXPEDITOR",
                "ETA",
                "ON SITE",
                "ENTREGADO"
            ]
        )
        .reset_index(drop=True)
    )


    # ====================================================
    # RETORNAR
    # ====================================================

    return (
        df,
        df_normalizado,
        df_subsistema_fmr,
        df_fmr_po,
        df_fmr_detalle,
        df_subsistema_fmr_po,
        df_3week_subsistema
    )


# ========================================================
# PRUEBA
# ========================================================

if __name__ == "__main__":

    (
        df,
        df_normalizado,
        df_subsistema_fmr,
        df_fmr_po,
        df_fmr_detalle,
        df_subsistema_fmr_po,
        df_3week_subsistema
    ) = load_data()


    print("\n========================================")
    print("DATOS CARGADOS")
    print("========================================")

    print(
        "Tabla principal:",
        len(df),
        "filas"
    )

    print(
        "FMR_Normalizado:",
        len(df_normalizado),
        "filas"
    )

    print(
        "SubSistema ↔ FMR:",
        len(df_subsistema_fmr),
        "filas"
    )

    print(
        "SubSistema ↔ FMR ↔ PO:",
        len(df_subsistema_fmr_po),
        "filas"
    )

    print(
        "FMR ↔ PO:",
        len(df_fmr_po),
        "filas"
    )

    print(
        "Detalle Procurement:",
        len(df_fmr_detalle),
        "filas"
    )


    # ====================================================
    # PRUEBA ESPECÍFICA
    # ====================================================

    print("\n========================================")
    print("PRUEBA")
    print("36045-ROT-005 → 747667")
    print("========================================")


    prueba_po = df_subsistema_fmr_po[
        (
            df_subsistema_fmr_po["SubSistema"]
            == "36045-ROT-005"
        )
        &
        (
            df_subsistema_fmr_po["FMR"]
            == "747667"
        )
    ]


    print("\nPO ENCONTRADAS:")

    print(
        prueba_po.to_string(
            index=False
        )
    )


    prueba_procurement = df_fmr_detalle[
        (
            df_fmr_detalle["SubSistema"]
            == "36045-ROT-005"
        )
        &
        (
            df_fmr_detalle["FMR"]
            == "747667"
        )
    ]


    print("\nDETALLE PROCUREMENT:")

    print(
        prueba_procurement.to_string(
            index=False
        )
    )

    print("3week ↔ SubSistema:", len(df_3week_subsistema), "filas")

    print("\n========================================")
    print("FECHAS 3WEEK DISPONIBLES")
    print("========================================")

    print(
    df_3week_subsistema["3week"]
    .drop_duplicates()
    .sort_values()
    .dt.strftime("%d/%m/%Y")
    .tolist()
    )