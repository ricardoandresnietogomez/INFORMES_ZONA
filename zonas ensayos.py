import os
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

try:
    import win32com.client as win32
except ImportError:
    win32 = None

ZONAS = ["METROPOLITANA", "NORTE", "ORIENTE", "SUR"]
FECHA_CORTE = pd.Timestamp("2026-04-15")


def safe_numeric(series, fill_value=0):
    return pd.to_numeric(series, errors="coerce").fillna(fill_value)


def ensure_directories(*paths):
    for path in paths:
        os.makedirs(path, exist_ok=True)


def format_bar_label(value, decimals=2):
    if pd.isna(value) or value <= 0:
        return ""
    return f"{value:.{decimals}f}"


def infer_group_colors(columns, grupos_camioneta, grupos_moto):
    colores = []
    for grupo in columns:
        if grupo in grupos_camioneta:
            colores.append("#4472C4")
        elif grupo in grupos_moto:
            colores.append("#ED7D31")
        else:
            colores.append("gray")
    return colores


def save_zone_bar_chart(
    df,
    zona,
    value_col,
    title,
    y_label,
    output_dir,
    file_name,
    columns_name,
    grupos_camioneta=None,
    grupos_moto=None,
    decimals=2,
    width=0.78
):
    datos_zona = df[df["ZONA"] == zona].copy()

    if datos_zona.empty:
        return None

    tabla = datos_zona.pivot_table(
        index="ITEM_SERVI",
        columns=columns_name,
        values=value_col,
        aggfunc="sum",
        fill_value=0
    )

    if tabla.empty:
        return None

    tabla["TOTAL"] = tabla.sum(axis=1)
    tabla = tabla.sort_values("TOTAL", ascending=False).drop(columns="TOTAL")

    colors = None
    if grupos_camioneta is not None and grupos_moto is not None:
        colors = infer_group_colors(tabla.columns, grupos_camioneta, grupos_moto)

    ax = tabla.plot(
        kind="bar",
        figsize=(ANCHO_GRAFICA, ALTO_GRAFICA),
        color=colors,
        width=width
    )

    aplicar_formato_grafica(ax, title, "Actividad", y_label)

    for container in ax.containers:
        labels = [format_bar_label(barra.get_height(), decimals) for barra in container]
        ax.bar_label(container, labels=labels, fontsize=FONT_LABEL, padding=2)

    plt.legend(
        title="GRUPOS",
        fontsize=FONT_LEYENDA,
        title_fontsize=FONT_LEYENDA,
        bbox_to_anchor=(1.02, 1),
        loc="upper left"
    )

    plt.tight_layout()

    ruta = os.path.join(output_dir, file_name.format(zona=zona))
    plt.savefig(ruta, dpi=DPI_GRAFICA, bbox_inches="tight")
    plt.close()
    return ruta


def crear_mail_territorio(territorio, correos_territorio):
    if win32 is None:
        print(
            f"Outlook no está disponible para {territorio}. "
            "Se omite el envío del correo y continúa la generación del informe."
        )
        return None

    try:
        outlook = win32.Dispatch("Outlook.Application")
        return outlook.CreateItem(0)
    except Exception as exc:
        print(f"No se pudo abrir Outlook para {territorio}: {exc}")
        return None


# ==========================================================
# RUTA DEL ARCHIVO DE ORIGEN
# ==========================================================

archivo_origen = (
    r"D:\OneDrive - Grupo EPM"
    r"\PYTHON PODA ARCHIVOS DE INFORMES"
    r"\Archivo mes a mes_AJUSTADO.xlsx"
)
    
# ==========================================================
# ESTRUCTURA DE CARPETAS
# ==========================================================

carpeta_resultados = (
    r"D:\OneDrive - Grupo EPM"
    r"\PYTHON PODA ARCHIVOS DE INFORMES"
    r"\Datos_Globales"
)




# ==========================================================
# ESTANDAR GRAFICAS CORREO OUTLOOK
# ==========================================================

ANCHO_GRAFICA = 40
ALTO_GRAFICA = 20

DPI_GRAFICA = 300

FONT_TITULO = 36
FONT_EJE = 28
FONT_TICKS = 22
FONT_LABEL = 22
FONT_LEYENDA = 24



FONT_LEYENDA = 30
ROTACION_X = 45


def aplicar_formato_grafica(ax, titulo, xlabel, ylabel):

    ax.tick_params(
        axis="x",
        labelsize=FONT_TICKS
    )

    ax.tick_params(
        axis="y",
        labelsize=FONT_TICKS
    )

    plt.title(
        titulo,
        fontsize=FONT_TITULO,
        fontweight="bold"
    )

    plt.xlabel(
        xlabel,
        fontsize=FONT_EJE
    )

    plt.ylabel(
        ylabel,
        fontsize=FONT_EJE
    )

    plt.xticks(
        rotation=ROTACION_X,
        ha="right"
    )

    if ax.get_legend() is not None:

        plt.legend(
            fontsize=FONT_LEYENDA,
            title_fontsize=FONT_LEYENDA
        )

    plt.tight_layout()

carpeta_reportes = os.path.join(
    carpeta_resultados,
    "00_Reportes"
)

carpeta_territorios = os.path.join(
    carpeta_resultados,
    "01_Territorios"
)

carpeta_horas = os.path.join(
    carpeta_resultados,
    "02_Horas_Territorio"
)
carpeta_km = os.path.join(
    carpeta_resultados,
    "03_Kilometros"
)


ruta_archivo = os.path.join(
    carpeta_km,
    "Kilometros_Item_Territorio.png"
)




# ==========================================================
# CREAR TODAS LAS CARPETAS
# ==========================================================

ensure_directories(
    carpeta_resultados,
    carpeta_reportes,
    carpeta_territorios,
    carpeta_horas,
    carpeta_km
)

# ==========================================================
# ARCHIVO PRINCIPAL DE SALIDA
# ==========================================================

archivo_salida = os.path.join(
    carpeta_reportes,
    "Informe Final.xlsx"
)

print(
    "\nEstructura de carpetas creada correctamente."
)








# ==========================================================
# BUSCAR AUTOMATICAMENTE LA HOJA ENERO
# ==========================================================

print("Buscando hoja Enero...")

excel = pd.ExcelFile(
    archivo_origen,
    engine="openpyxl"
)

hoja_enero = None

for hoja in excel.sheet_names:

    if "ENERO" in hoja.upper():
        hoja_enero = hoja
        break

if hoja_enero is None:

    print("\nHojas encontradas:")

    for hoja in excel.sheet_names:
        print("-", hoja)

    raise Exception(
        "No se encontró una hoja que contenga la palabra ENERO"
    )

print(f"Hoja encontrada: {hoja_enero}")

# ==========================================================
# LEER HOJA
# ==========================================================

df = pd.read_excel(
    archivo_origen,
    sheet_name=hoja_enero,
    engine="openpyxl"
)

print(f"Registros leídos: {len(df):,}")

# ==========================================================
# ELIMINAR COLUMNAS V:AB
# ==========================================================

if len(df.columns) >= 28:

    print("Eliminando columnas V:AB...")

    df.drop(
        columns=df.columns[21:28],
        inplace=True
    )

# ==========================================================
# IDENTIFICAR COLUMNAS
# ==========================================================

col_usuario = None
col_fecha = None
col_circuito = None

for columna in df.columns:

    nombre = str(columna).strip().upper()

    if nombre == "USUARIO":
        col_usuario = columna

    elif nombre == "FECHA":
        col_fecha = columna

    elif "CIRCUITO" in nombre:
        col_circuito = columna

if col_usuario is None:
    raise Exception("No existe la columna USUARIO")

if col_fecha is None:
    raise Exception("No existe la columna FECHA")

if col_circuito is None:
    raise Exception(
        "No se encontró la columna CIRCUITO"
    )

print("USUARIO :", col_usuario)
print("FECHA   :", col_fecha)
print("CIRCUITO:", col_circuito)

# ==========================================================
# EXTRAER CODIGO USUARIO
# ==========================================================

print("Extrayendo códigos de usuario...")

df[col_usuario] = (
    df[col_usuario]
    .astype(str)
    .str.upper()
    .str.extract(
        r'([A-Z]{3}\d{2})',
        expand=False
    )
)

# ==========================================================
# CONVERTIR FECHA
# ==========================================================

df[col_fecha] = pd.to_datetime(
    df[col_fecha],
    errors="coerce",
    dayfirst=True
)

# ==========================================================
# HOMOLOGAR ITEMS DESDE EL 15/04/2026
# ==========================================================

if "ITEM_SERVI" in df.columns:

    print(
        "Homologando nombres de ITEM_SERVI "
        "desde el 15/04/2026..."
    )

    fecha_homologacion = pd.Timestamp(
        "2026-04-15"
    )

    # NORMALIZAR TEXTO

    df["ITEM_SERVI"] = (
        df["ITEM_SERVI"]
        .astype(str)
        .str.strip()
    )

    cambios_items = {

        # ==================================================
        # INTERVENCION FORESTAL
        # ==================================================

        "Intervencion Forestal tipo cerco vivo y/o roceria":
            "Intervencion forestal en cerco vivo y/o roceria",

        # ==================================================
        # PODA URBANA
        # ==================================================

        "Poda Urbana":
            "Poda en redes Urbanas",

        "Poda urbana":
            "Poda en redes Urbanas",

        "Poda en redes urbanas":
            "Poda en redes Urbanas",

        "Poda en Redes Urbanas":
            "Poda en redes Urbanas",

        "PODA EN REDES URBANAS":
            "Poda en redes Urbanas",

        "Poda en redes Urbanas":
            "Poda en redes Urbanas",

        # ==================================================
        # PODA RURAL TIPO A
        # ==================================================

        "Poda en redes Rurales sin interrupcion del servicio Tipo A":
            "Poda en redes Rurales Tipo A",

        "Poda en redes Rurales sin interrupción del servicio Tipo A":
            "Poda en redes Rurales Tipo A",

        "Poda en redes Rurales Tipo A":
            "Poda en redes Rurales Tipo A",

        # ==================================================
        # PODA RURAL TIPO B
        # ==================================================

        "Poda en redes Rurales sin interrupcion del servicio Tipo B":
            "Poda en redes Rurales Tipo B",

        "Poda en redes Rurales sin interrupción del servicio Tipo B":
            "Poda en redes Rurales Tipo B",

        "Poda en redes Rurales Tipo B":
            "Poda en redes Rurales Tipo B",

        # ==================================================
        # PODA CON INTERRUPCION
        # ==================================================

        "Poda en redes Rurales con interrupcion del servicio":
            "Poda en redes Rurales con interrupcion del servicio",

        "Poda en redes Rurales con interrupción del servicio":
            "Poda en redes Rurales con interrupcion del servicio",

        "Poda en redes Rurales con interrupción del servicio de energía":
            "Poda en redes Rurales con interrupcion del servicio",

        # ==================================================
        # TALA PORTE ALTO
        # ==================================================

        "Tala de arboles clasificacion porte alto > 15":
            "Tala de arboles clasificacion porte alto DAP >40 CM",

        "Tala de árboles clasificación porte alto >15":
            "Tala de arboles clasificacion porte alto DAP >40 CM",

        # ==================================================
        # TALA PORTE MEDIO
        # ==================================================

        "Tala de arboles clasificacion porte medio":
            "Tala de arboles clasificacion porte medio DAP <= 40 CM",

        "Tala de arboles clasificacion porte medio <= 15":
            "Tala de arboles clasificacion porte medio DAP <= 40 CM",

        "Tala de árboles clasificación porte medio <=15":
            "Tala de arboles clasificacion porte medio DAP <= 40 CM",

        # ==================================================
        # GUADUAS
        # ==================================================

        "Tala de guaduas T1":
            "Tala de guaduas",

        "Tala de guaduas T2":
            "Tala de guaduas",

        "Tala de guaduas T3":
            "Tala de guaduas"
    }

    mascara_fecha = (
        df[col_fecha] >= fecha_homologacion
    )

    registros_modificados = 0

    for item_original, item_nuevo in cambios_items.items():

        mascara_item = (
            df["ITEM_SERVI"]
            .astype(str)
            .str.strip()
            .str.upper()
            ==
            item_original.upper()
        )

        cantidad = (
            mascara_fecha & mascara_item
        ).sum()

        registros_modificados += cantidad

        df.loc[
            mascara_fecha & mascara_item,
            "ITEM_SERVI"
        ] = item_nuevo

    print(
        f"Registros homologados: "
        f"{registros_modificados:,}"
    )

    print(
        "\nVALIDACION PODAS URBANAS\n"
    )

    print(
        df[
            df[col_fecha] >= fecha_homologacion
        ][
            "ITEM_SERVI"
        ]
        .value_counts()
        .loc[
            lambda x:
            x.index.str.contains(
                "urban",
                case=False,
                na=False
            )
        ]
    )

else:

    print(
        "No existe la columna ITEM_SERVI"
    )

# ==========================================================
# AUDITORIA Y ELIMINACION DE DATOS ATIPICOS
# DESDE EL 15/04/2026
# ==========================================================

print("\nValidando datos atípicos...")

fecha_filtro = pd.Timestamp("2026-04-15")

df["CANTIDAD"] = safe_numeric(df["CANTIDAD"])

mascara_fecha = (
    df[col_fecha] >= fecha_filtro
)

mascara_atipicos = (

    (
        (df["ITEM_SERVI"].astype(str).str.strip()
         == "Poda en redes Rurales con interrupcion del servicio")
        & (df["CANTIDAD"] > 30)
    )

    |

    (
        (df["ITEM_SERVI"].astype(str).str.strip()
         == "Poda en redes Rurales Tipo A")
        & (df["CANTIDAD"] > 30)
    )

    |

    (
        (df["ITEM_SERVI"].astype(str).str.strip()
         == "Poda en redes Rurales Tipo B")
        & (df["CANTIDAD"] > 30)
    )

    |

    (
        (df["ITEM_SERVI"].astype(str).str.strip()
         == "Poda en redes Urbanas")
        & (df["CANTIDAD"] > 20)
    )

    |

    (
        (df["ITEM_SERVI"].astype(str).str.strip()
         == "Tala de arboles clasificacion porte medio DAP <= 40 CM")
        & (df["CANTIDAD"] > 40)
    )

    |

    (
        (df["ITEM_SERVI"].astype(str).str.strip()
         == "Tala de arboles clasificacion porte alto DAP >40 CM")
        & (df["CANTIDAD"] > 40)
    )

    |

    (
        (df["ITEM_SERVI"].astype(str).str.strip()
         == "Tala de guaduas")
        & (df["CANTIDAD"] > 150)
    )
)

mascara_atipicos = (
    mascara_fecha
    & mascara_atipicos
)

registros_atipicos = df[
    mascara_atipicos
].copy()

print(
    f"Registros atípicos detectados: "
    f"{len(registros_atipicos):,}"
)

df = df[
    ~mascara_atipicos
].copy()

print(
    f"Registros después de depuración: "
    f"{len(df):,}"
)



# ==========================================================
# ELIMINAR FILAS CON ITEM_SERVI NUMERICO
# ==========================================================

if "ITEM_SERVI" in df.columns:

    print("Eliminando registros con ITEM_SERVI numérico...")

    mascara_numerica = pd.to_numeric(
        df["ITEM_SERVI"],
        errors="coerce"
    ).notna()

    registros_eliminados = mascara_numerica.sum()

    df = df[~mascara_numerica].copy()

    print(
        f"Registros eliminados: {registros_eliminados}"
    )

else:

    print(
        "No existe la columna ITEM_SERVI"
    )

# ==========================================================
# ELIMINAR FILAS CON LETRAS EN DURA_ACT
# ==========================================================

if "DURA_ACT" in df.columns:

    print("Eliminando registros con letras en DURA_ACT...")

    df["DURA_ACT"] = pd.to_numeric(
        df["DURA_ACT"],
        errors="coerce"
    )

    registros_eliminados = (
        df["DURA_ACT"].isna()
    ).sum()

    df = df[
        df["DURA_ACT"].notna()
    ].copy()

    print(
        f"Registros eliminados en DURA_ACT: {registros_eliminados}"
    )

else:

    print(
        "No existe la columna DURA_ACT"
    )

# ==========================================================
# ELIMINAR FILAS CON "-"
# EN CLASIFICACION_MERCADO_TRM_VEGETACION
# ==========================================================

if "CLASIFICACION_MERCADO_TRM_VEGETACION" in df.columns:

    print(
        "Eliminando registros con '-' en "
        "CLASIFICACION_MERCADO_TRM_VEGETACION..."
    )

    registros_eliminados = (
        df["CLASIFICACION_MERCADO_TRM_VEGETACION"]
        .astype(str)
        .str.strip()
        .eq("-")
        .sum()
    )

    df = df[
        df["CLASIFICACION_MERCADO_TRM_VEGETACION"]
        .astype(str)
        .str.strip()
        != "-"
    ].copy()

    print(
        f"Registros eliminados: {registros_eliminados}"
    )

else:

    print(
        "No existe la columna "
        "CLASIFICACION_MERCADO_TRM_VEGETACION"
    )

# ==========================================================
# CREAR COLUMNAS AÑO MES DIA
# ==========================================================

print("Generando columnas AÑO, MES y DIA...")

df["AÑO"] = df[col_fecha].dt.year
df["MES"] = df[col_fecha].dt.month
df["DIA"] = df[col_fecha].dt.day

# ==========================================================
# CREAR ZONA
# ==========================================================

print("Generando columna ZONA...")

df["ZONA"] = (
    df[col_usuario]
    .astype(str)
    .str[:3]
    .map({
        "DMC": "METROPOLITANA",
        "DME": "ORIENTE",
        "DMD": "SUR",
        "DMN": "NORTE",
        "DMW": "NORTE"
    })
)


# ==========================================================
# CREAR TIPO DE RECURSO
# ==========================================================

print("Generando TIPO DE RECURSO...")

df["TIPO DE RECURSO"] = pd.NA

fecha_corte_recurso = pd.Timestamp("2026-04-15")

# ==========================================================
# ANTES DEL NUEVO ESQUEMA OPERATIVO
# ==========================================================

df.loc[
    df[col_fecha] < fecha_corte_recurso,
    "TIPO DE RECURSO"
] = "CAMIONETA"

# ==========================================================
# GRUPOS CAMIONETA
# ==========================================================

grupos_camioneta = [

    # METROPOLITANA
    "DMC80", "DMC81", "DMC84",

    # ORIENTE
    "DME80", "DME81", "DME84","DME85",

    # NORTE
    "DMN80", "DMN81","DMW80",

    # SUR
    "DMD80", "DMD81",

    # FENOMENO DEL NIÑO
    #"DMW80"
]

# ==========================================================
# GRUPOS MOTO
# ==========================================================

grupos_moto = [

    # METROPOLITANA
    "DMC82", "DMC83",

    # ORIENTE
    "DME82", "DME83",

    # NORTE
    "DMN82", "DMN83",

    # SUR
    "DMD82", "DMD83"
]

# ==========================================================
# ASIGNACION DESDE EL 15/04/2026
# ==========================================================

df.loc[
    (
        df[col_fecha] >= fecha_corte_recurso
    )
    &
    (
        df[col_usuario].isin(grupos_camioneta)
    ),
    "TIPO DE RECURSO"
] = "CAMIONETA"

df.loc[
    (
        df[col_fecha] >= fecha_corte_recurso
    )
    &
    (
        df[col_usuario].isin(grupos_moto)
    ),
    "TIPO DE RECURSO"
] = "MOTO"

# ==========================================================
# VALIDACION
# ==========================================================

print("\nDistribución por tipo de recurso:")

print(
    df["TIPO DE RECURSO"]
    .value_counts(dropna=False)
)

registros_sin_recurso = int(
    df["TIPO DE RECURSO"]
    .isna()
    .sum()
)

print(
    f"Registros sin tipo de recurso: "
    f"{registros_sin_recurso:,}"
)

# ==========================================================
# MOSTRAR GRUPOS SIN CLASIFICAR
# ==========================================================

if registros_sin_recurso > 0:

    print(
        "\nGRUPOS SIN RECURSO ASIGNADO:"
    )

    print(
        sorted(
            df.loc[
                df["TIPO DE RECURSO"].isna(),
                col_usuario
            ]
            .dropna()
            .unique()
        )
    )


# ==========================================================
# REORDENAR COLUMNAS
# ==========================================================

columnas = list(df.columns)

for col in ["ZONA", "TIPO DE RECURSO", "AÑO", "MES", "DIA"]:
    columnas.remove(col)

pos_usuario = columnas.index(col_usuario)

columnas.insert(pos_usuario + 1, "ZONA")
columnas.insert(pos_usuario + 2, "TIPO DE RECURSO")
columnas.insert(pos_usuario + 3, "AÑO")
columnas.insert(pos_usuario + 4, "MES")
columnas.insert(pos_usuario + 5, "DIA")

df = df[columnas]

# ==========================================================
# FILTRAR CIRCUITOS
# ==========================================================

print("Separando circuitos...")

circuitos = (
    df[col_circuito]
    .astype(str)
    .str.upper()
)

df_informe = df[
    circuitos.str.contains("23L", na=False)
].copy()

df_33 = df[
    circuitos.str.contains("30L", na=False)
].copy()

df_115 = df[
    circuitos.str.contains("40L", na=False)
].copy()

print(f"Informe : {len(df_informe):,}")
print(f"33KV    : {len(df_33):,}")
print(f"115 KV  : {len(df_115):,}")


# ==========================================================
# PROMEDIOS DESDE EL 15/04/2026
# ==========================================================

print("Calculando promedios...")

fecha_promedios = pd.Timestamp("2026-04-15")

df_promedios = df_informe[
    df_informe[col_fecha] >= fecha_promedios
].copy()

df_promedios["CANTIDAD"] = safe_numeric(df_promedios["CANTIDAD"])
df_promedios["DURA_ACT"] = safe_numeric(df_promedios["DURA_ACT"])
df_promedios["LONGITUD_KM"] = safe_numeric(df_promedios["LONGITUD_KM"])

print("\nVALIDACION DUPLICADOS\n")

duplicados = df_promedios[
    (df_promedios["ZONA"] == "NORTE")
    &
    (
        df_promedios["ITEM_SERVI"]
        == "Poda en redes Rurales Tipo A"
    )
]

print(
    "Registros totales:",
    len(duplicados)
)

print(
    "Duplicados exactos:",
    duplicados.duplicated().sum()
)



# ==========================================================
# PROMEDIOS POR USUARIO
# ==========================================================

resumen = (
    df_promedios.groupby(
        [
            col_usuario,
            "ZONA",
            "TIPO DE RECURSO",
            "ITEM_SERVI"
        ],
        dropna=False
    )
    .agg(
        DIAS_TRABAJADOS=(
            col_fecha,
            lambda x: x.dt.date.nunique()
        ),
        CANTIDAD_TOTAL=("CANTIDAD", "sum"),
        DURACION_TOTAL=("DURA_ACT", "sum"),
        LONGITUD_TOTAL_KM=("LONGITUD_KM", "sum")
    )
    .reset_index()
)

resumen["PROM_CANTIDAD_DIA"] = (
    resumen["CANTIDAD_TOTAL"]
    / resumen["DIAS_TRABAJADOS"]
).round(2)

resumen["PROM_DURACION_DIA"] = (
    resumen["DURACION_TOTAL"]
    / resumen["DIAS_TRABAJADOS"]
).round(2)

resumen["PROM_LONGITUD_DIA_KM"] = (
    resumen["LONGITUD_TOTAL_KM"]
    / resumen["DIAS_TRABAJADOS"]
).round(2)

resumen["HORAS_POR_ITEM"] = np.where(
    resumen["CANTIDAD_TOTAL"] > 0,
    resumen["DURACION_TOTAL"]
    / resumen["CANTIDAD_TOTAL"],
    np.nan
)

resumen["MINUTOS_POR_ITEM"] = (
    resumen["HORAS_POR_ITEM"]
    * 60
)

resumen["HORAS_POR_ITEM"] = (
    resumen["HORAS_POR_ITEM"]
    .round(4)
)

resumen["MINUTOS_POR_ITEM"] = (
    resumen["MINUTOS_POR_ITEM"]
    .round(2)
)

# ==========================================================
# PROMEDIOS KM/DIA POR GRUPO
# DIAS REALES SIN DUPLICAR FECHAS
# ==========================================================

promedio_grupos_km = (
    df_promedios.groupby(
        [
            "ZONA",
            col_usuario,
            "TIPO DE RECURSO"
        ],
        dropna=False
    )
    .agg(
        DIAS_TRABAJADOS=(
            col_fecha,
            lambda x: x.dt.date.nunique()
        ),
        LONGITUD_TOTAL_KM=(
            "LONGITUD_KM",
            "sum"
        )
    )
    .reset_index()
)

promedio_grupos_km["KM_POR_DIA"] = np.where(
    promedio_grupos_km["DIAS_TRABAJADOS"] > 0,
    promedio_grupos_km["LONGITUD_TOTAL_KM"]
    /
    promedio_grupos_km["DIAS_TRABAJADOS"],
    np.nan
)

promedio_grupos_km["KM_POR_DIA"] = (
    promedio_grupos_km["KM_POR_DIA"]
    .round(2)
)

print("\nVALIDACION DIAS REALES POR GRUPO\n")

print(
    promedio_grupos_km[
        [
            col_usuario,
            "ZONA",
            "TIPO DE RECURSO",
            "DIAS_TRABAJADOS",
            "LONGITUD_TOTAL_KM",
            "KM_POR_DIA"
        ]
    ]
    .sort_values(
        ["ZONA", col_usuario]
    )
)


# ==========================================================
# IDENTIFICAR GRUPOS FENOMENO DEL NIÑO
# ==========================================================

grupos_fenomeno_nino = [
    "DMC84",
    "DME84",
    "DME85",
    "DMW80"
]

promedio_grupos_km["FENOMENO_NINO"] = np.where(
    promedio_grupos_km[col_usuario].isin(grupos_fenomeno_nino),
    "SI",
    "NO"
)

# ==========================================================
# TABLA CONSOLIDADA DE RENDIMIENTO POR GRUPO
# ==========================================================

# ==========================================================
# ORDEN PERSONALIZADO DE TERRITORIOS
# ==========================================================

orden_territorios = {
    "METROPOLITANA": 1,
    "NORTE": 2,
    "ORIENTE": 3,
    "SUR": 4
}

tabla_rendimiento_grupos = (
    promedio_grupos_km[
        [
            col_usuario,
            "ZONA",
            "TIPO DE RECURSO",
            "FENOMENO_NINO",
            "DIAS_TRABAJADOS",
            "LONGITUD_TOTAL_KM",
            "KM_POR_DIA"
        ]
    ]
    .copy()
)

tabla_rendimiento_grupos["ORDEN_ZONA"] = (
    tabla_rendimiento_grupos["ZONA"]
    .map(orden_territorios)
)

tabla_rendimiento_grupos = (
    tabla_rendimiento_grupos
    .sort_values(
        ["ORDEN_ZONA", "KM_POR_DIA"],
        ascending=[True, False]
    )
    .drop(columns="ORDEN_ZONA")
)

# ==========================================================
# ORDEN FIJO DE TERRITORIOS
# ==========================================================

orden_territorios = {
    "METROPOLITANA": 1,
    "NORTE": 2,
    "ORIENTE": 3,
    "SUR": 4
}

tabla_rendimiento_grupos["ORDEN_ZONA"] = (
    tabla_rendimiento_grupos["ZONA"]
    .map(orden_territorios)
)

tabla_rendimiento_grupos = (
    tabla_rendimiento_grupos
    .sort_values(
        ["ORDEN_ZONA", "KM_POR_DIA"],
        ascending=[True, False]
    )
    .drop(columns="ORDEN_ZONA")
    .reset_index(drop=True)
)


print("\nRENDIMIENTO OPERATIVO POR GRUPO\n")
print(tabla_rendimiento_grupos)

# ==========================================================
# CONCLUSION AUTOMATICA RENDIMIENTO POR GRUPO
# ==========================================================

ranking_grupos = (
    tabla_rendimiento_grupos
    .sort_values(
        "KM_POR_DIA",
        ascending=False
    )
    .reset_index(drop=True)
)

grupo_top = ranking_grupos.iloc[0][col_usuario]
territorio_top = ranking_grupos.iloc[0]["ZONA"]
km_top = ranking_grupos.iloc[0]["KM_POR_DIA"]

grupo_bajo = ranking_grupos.iloc[-1][col_usuario]
km_bajo = ranking_grupos.iloc[-1]["KM_POR_DIA"]

km_promedio_grupos = (
    tabla_rendimiento_grupos["KM_POR_DIA"]
    .mean()
)

superior_promedio = (
    (km_top / km_promedio_grupos) - 1
) * 100

html_conclusion_grupos = f"""
<p style="font-family:Arial;font-size:11pt;text-align:justify;">

<b>Conclusión Ejecutiva:</b>

<br><br>

El rendimiento operativo promedio alcanzó
<b>{km_promedio_grupos:.2f} km por día</b>,
destacándose
<b>{grupo_top}</b>
del territorio
<b>{territorio_top}</b>
como el grupo con mayor capacidad de cobertura,
registrando
<b>{km_top:.2f} km por día</b>,
equivalente a un desempeño
<b>{superior_promedio:.1f}% superior</b>
al promedio general.

<br><br>

En contraste,
<b>{grupo_bajo}</b>
registró el menor rendimiento con
<b>{km_bajo:.2f} km por día</b>,
reflejando oportunidades de fortalecimiento operativo frente a los grupos de mayor desempeño.

</p>

<br><br>
"""

# ==========================================================
# RESUMEN POR TIPO DE RECURSO
# CAMIONETAS / FENOMENO DEL NIÑO / MOTOS
# ==========================================================

tabla_resumen_recursos = (
    tabla_rendimiento_grupos.copy()
)

tabla_resumen_recursos["CATEGORIA"] = np.select(
    [
        tabla_resumen_recursos["FENOMENO_NINO"] == "SI",
        tabla_resumen_recursos["TIPO DE RECURSO"] == "MOTO"
    ],
    [
        "CAMIONETAS FENOMENO DEL NIÑO",
        "MOTOS"
    ],
    default="CAMIONETAS"
)

tabla_resumen_recursos = (
    tabla_resumen_recursos
    .groupby(
        ["ZONA", "CATEGORIA"],
        dropna=False
    )
    .agg(
        GRUPOS=(col_usuario, "nunique"),
        KM_EJECUTADOS=("LONGITUD_TOTAL_KM", "sum"),
        DIAS_TRABAJADOS=("DIAS_TRABAJADOS", "sum"),
        KM_POR_DIA=("KM_POR_DIA", "mean")
    )
    .reset_index()
)

tabla_resumen_recursos["KM_EJECUTADOS"] = (
    tabla_resumen_recursos["KM_EJECUTADOS"]
    .round(2)
)

tabla_resumen_recursos["KM_POR_DIA"] = (
    tabla_resumen_recursos["KM_POR_DIA"]
    .round(2)
)

print("\nRESUMEN RECURSOS\n")
print(tabla_resumen_recursos)


# ==========================================================
# KM MES A MES POR TERRITORIO
# ==========================================================

for zona in ZONAS:

    datos_zona = df_informe[
        df_informe["ZONA"] == zona
    ].copy()

    if datos_zona.empty:
        continue

    km_mes_zona = (
        datos_zona.groupby(
            pd.Grouper(
                key=col_fecha,
                freq="M"
            )
        )
        .agg(
            KM=("LONGITUD_KM", "sum")
        )
        .reset_index()
    )

    km_mes_zona["MES"] = (
        km_mes_zona[col_fecha]
        .dt.strftime("%Y-%m")
    )

    fecha_cambio = pd.Timestamp("2026-04-15")

    promedio_antes = (
        km_mes_zona.loc[
            km_mes_zona[col_fecha] < fecha_cambio,
            "KM"
        ]
        .mean()
    )

    promedio_despues = (
        km_mes_zona.loc[
            km_mes_zona[col_fecha] >= fecha_cambio,
            "KM"
        ]
        .mean()
    )

    plt.figure(figsize=(18,10))

    barras = plt.bar(
        km_mes_zona["MES"],
        km_mes_zona["KM"],
        color="#0E5A8A"
    )

    for barra in barras:

        valor = barra.get_height()

        plt.text(
            barra.get_x() + barra.get_width()/2,
            valor + 2,
            f"{valor:.2f}".replace(".", ","),
            ha="center",
            fontsize=8,
            fontweight="bold"
        )

    plt.axhline(
        promedio_antes,
        color="green",
        linewidth=1.5,
        label=f"Promedio Antes ({promedio_antes:.2f} km)"
    )

    plt.axhline(
        promedio_despues,
        color="orange",
        linewidth=1.5,
        label=f"Promedio Después ({promedio_despues:.2f} km)"
    )

    indice_cambio = km_mes_zona[
        km_mes_zona[col_fecha] >= fecha_cambio
    ].index.min()

    plt.axvline(
        indice_cambio - 0.5,
        color="red",
        linestyle="--"
    )

    plt.text(
        indice_cambio + 0.4,
        km_mes_zona["KM"].max() * 0.93,
        "Abr-2026\n2 Camionetas\n+ 2 Motos",
        color="red",
        fontsize=10,
        fontweight="bold",
        bbox=dict(
            facecolor="white",
            edgecolor="red"
        )
    )

    plt.title(
        f"Kilómetros Ejecutados Mes a Mes - {zona}",
        fontsize=16,
        fontweight="bold"
    )

    plt.xlabel("Mes")
    plt.ylabel("Kilómetros Ejecutados")

    plt.xticks(rotation=45)

    plt.legend()

    plt.grid(alpha=0.20)

    plt.tight_layout()

    ruta_zona_mes = os.path.join(
        carpeta_km,
        f"{zona}_KM_MES.png"
    )

    plt.savefig(
        ruta_zona_mes,
        dpi=300,
        bbox_inches="tight"
    )

    plt.close()

    print(f"Gráfica creada: {ruta_zona_mes}")


# ==========================================================
# HTML RESUMEN DE RECURSOS POR TERRITORIO
# ==========================================================
html_resumen_recursos = ""

for territorio in ["METROPOLITANA", "NORTE", "ORIENTE", "SUR"]:

    datos_territorio = (
        tabla_resumen_recursos[
            tabla_resumen_recursos["ZONA"] == territorio
        ]
        .copy()
    )

    if datos_territorio.empty:
        continue

    html_resumen_recursos += f"""
    <h4 style="color:#0066CC;">
    {territorio}
    </h4>

    <table border="1"
    style="border-collapse:collapse;font-family:Arial;font-size:10pt;">
    <tr style="background:#D9EAD3;">
        <th>Categoría</th>
        <th>Grupos</th>
        <th>KM Ejecutados</th>
        <th>Días Trabajados</th>
        <th>KM por Día Promedio</th>
    </tr>
    """

    for _, fila in datos_territorio.iterrows():

        html_resumen_recursos += f"""
        <tr>
            <td><b>{fila['CATEGORIA']}</b></td>
            <td align="center">{fila['GRUPOS']}</td>
            <td align="center">{fila['KM_EJECUTADOS']:.2f}</td>
            <td align="center">{fila['DIAS_TRABAJADOS']:.0f}</td>
            <td align="center">
                <b>{fila['KM_POR_DIA']:.2f}</b>
            </td>
        </tr>
        """

    html_resumen_recursos += """
    </table>
    <br><br>
    """


print("\nGRUPOS FENOMENO DEL NIÑO\n")
print(
    tabla_rendimiento_grupos[
        tabla_rendimiento_grupos["FENOMENO_NINO"] == "SI"
    ]
    .sort_values(
        "KM_POR_DIA",
        ascending=False
    )
)

print("\n" + "="*120)
print("RENDIMIENTO OPERATIVO DE TODOS LOS GRUPOS")
print("="*120)

print(
    tabla_rendimiento_grupos[
        [
            col_usuario,
            "ZONA",
            "TIPO DE RECURSO",
            "FENOMENO_NINO",
            "DIAS_TRABAJADOS",
            "LONGITUD_TOTAL_KM",
            "KM_POR_DIA"
        ]
    ]
    .sort_values(
        ["FENOMENO_NINO", "KM_POR_DIA"],
        ascending=[False, False]
    )
    .to_string(index=False)
)


print("\n")
print("="*120)
print("RESUMEN FENOMENO DEL NIÑO VS OPERACION NORMAL")
print("="*120)

print(
    tabla_rendimiento_grupos.groupby(
        "FENOMENO_NINO",
        dropna=False
    )
    .agg(
        GRUPOS=(col_usuario, "count"),
        DIAS_TRABAJADOS=("DIAS_TRABAJADOS", "sum"),
        KM_TOTALES=("LONGITUD_TOTAL_KM", "sum"),
        KM_PROMEDIO_DIA=("KM_POR_DIA", "mean")
    )
    .round(2)
)



print("\nPROMEDIOS KM/DIA POR GRUPO - METROPOLITANA\n")

print(
    promedio_grupos_km[
        (promedio_grupos_km["ZONA"] == "METROPOLITANA")
        &
        (promedio_grupos_km["TIPO DE RECURSO"] == "CAMIONETA")
    ][[
        col_usuario,
        "DIAS_TRABAJADOS",
        "LONGITUD_TOTAL_KM",
        "KM_POR_DIA"
    ]]
    .sort_values(
        "KM_POR_DIA",
        ascending=False
    )
)


# ==========================================================
# ELIMINAR ARCHIVO EXISTENTE
# ==========================================================

try:

    if os.path.exists(archivo_salida):
        os.remove(archivo_salida)

except:
    pass


# ==========================================================
# PROMEDIO POR TERRITORIO PONDERADO
# ==========================================================

promedio_territorio = (
    resumen.groupby(
        [
            "ZONA",
            "ITEM_SERVI"
        ],
        dropna=False
    )
    .agg(
        CANTIDAD_TOTAL=(
            "CANTIDAD_TOTAL",
            "sum"
        ),
        DURACION_TOTAL=(
            "DURACION_TOTAL",
            "sum"
        )
    )
    .reset_index()
)

print("\nVALIDACION NORTE TIPO A\n")

print(
    promedio_territorio[
        (promedio_territorio["ZONA"] == "NORTE")
        &
        (
            promedio_territorio["ITEM_SERVI"]
            == "Poda en redes Rurales Tipo A"
        )
    ]
)

promedio_territorio["HORAS_POR_ITEM_PROM"] = (
    promedio_territorio["DURACION_TOTAL"]
    /
    promedio_territorio["CANTIDAD_TOTAL"]
)

promedio_territorio["MINUTOS_POR_ITEM_PROM"] = (
    promedio_territorio["HORAS_POR_ITEM_PROM"]
    * 60
)

promedio_territorio["HORAS_POR_ITEM_PROM"] = (
    promedio_territorio["HORAS_POR_ITEM_PROM"]
    .round(4)
)

promedio_territorio["MINUTOS_POR_ITEM_PROM"] = (
    promedio_territorio["MINUTOS_POR_ITEM_PROM"]
    .round(2)
)

promedio_territorio["HORAS_POR_ITEM_PROM"] = (
    promedio_territorio["HORAS_POR_ITEM_PROM"]
    .round(4)
)

promedio_territorio["MINUTOS_POR_ITEM_PROM"] = (
    promedio_territorio["MINUTOS_POR_ITEM_PROM"]
    .round(2)
)


print("\nREGISTROS EN DF_PROMEDIOS")
print(len(df_promedios))

print(
    df_promedios[
        (df_promedios["ZONA"] == "NORTE")
        &
        (
            df_promedios["ITEM_SERVI"]
            == "Poda en redes Rurales Tipo A"
        )
    ]["CANTIDAD"].sum()
)

# ==========================================================
# PROMEDIO POR TIPO DE RECURSO
# ==========================================================

promedio_tipo_recurso = (
    resumen.groupby(
        [
            "TIPO DE RECURSO",
            "ITEM_SERVI"
        ],
        dropna=False
    )
    .agg(
        HORAS_POR_ITEM_PROM=(
            "HORAS_POR_ITEM",
            "mean"
        ),
        MINUTOS_POR_ITEM_PROM=(
            "MINUTOS_POR_ITEM",
            "mean"
        )
    )
    .reset_index()
)

promedio_tipo_recurso["HORAS_POR_ITEM_PROM"] = (
    promedio_tipo_recurso["HORAS_POR_ITEM_PROM"]
    .round(4)
)

promedio_tipo_recurso["MINUTOS_POR_ITEM_PROM"] = (
    promedio_tipo_recurso["MINUTOS_POR_ITEM_PROM"]
    .round(2)
)


# ==========================================================
# PROMEDIO GLOBAL
# ==========================================================

promedio_global = (
    resumen.groupby(
        [
            "ITEM_SERVI"
        ],
        dropna=False
    )
    .agg(
        HORAS_POR_ITEM_PROM=(
            "HORAS_POR_ITEM",
            "mean"
        ),
        MINUTOS_POR_ITEM_PROM=(
            "MINUTOS_POR_ITEM",
            "mean"
        )
    )
    .reset_index()
)

promedio_global["HORAS_POR_ITEM_PROM"] = (
    promedio_global["HORAS_POR_ITEM_PROM"]
    .round(4)
)

promedio_global["MINUTOS_POR_ITEM_PROM"] = (
    promedio_global["MINUTOS_POR_ITEM_PROM"]
    .round(2)
)

# ==========================================================
# PROMEDIO GRUPO USUARIO
# ==========================================================

promedio_grupo_usuario = (
    resumen.groupby(
        [
            col_usuario,
            "ZONA",
            "TIPO DE RECURSO",
            "ITEM_SERVI"
        ],
        dropna=False
    )
    .agg(
        HORAS_POR_ITEM_PROM=(
            "HORAS_POR_ITEM",
            "mean"
        ),
        MINUTOS_POR_ITEM_PROM=(
            "MINUTOS_POR_ITEM",
            "mean"
        )
    )
    .reset_index()
)

promedio_grupo_usuario["HORAS_POR_ITEM_PROM"] = (
    promedio_grupo_usuario["HORAS_POR_ITEM_PROM"]
    .round(4)
)

promedio_grupo_usuario["MINUTOS_POR_ITEM_PROM"] = (
    promedio_grupo_usuario["MINUTOS_POR_ITEM_PROM"]
    .round(2)
)

promedio_grupo_usuario = promedio_grupo_usuario[
    [
        col_usuario,
        "ZONA",
        "TIPO DE RECURSO",
        "ITEM_SERVI",
        "HORAS_POR_ITEM_PROM",
        "MINUTOS_POR_ITEM_PROM"
    ]
]


# ==========================================================
# TABLA CANTIDAD POR TERRITORIO
# ==========================================================

tabla_cantidad_territorio = (
    resumen.groupby(
        [
            "ZONA",
            "ITEM_SERVI"
        ],
        dropna=False
    )
    .agg(
        CANTIDAD_TOTAL=(
            "CANTIDAD_TOTAL",
            "sum"
        )
    )
    .reset_index()
)

tabla_cantidad_territorio = (
    tabla_cantidad_territorio
    .sort_values(
        [
            "ZONA",
            "CANTIDAD_TOTAL"
        ],
        ascending=[
            True,
            False
        ]
    )
)

tabla_cantidad_territorio = tabla_cantidad_territorio[
    [
        "ZONA",
        "ITEM_SERVI",
        "CANTIDAD_TOTAL"
    ]
]



# ==========================================================
# CANTIDAD GLOBAL
# ==========================================================

cantidad_global = (
    resumen.groupby(
        [
            "ITEM_SERVI"
        ],
        dropna=False
    )
    .agg(
        CANTIDAD_TOTAL=(
            "CANTIDAD_TOTAL",
            "sum"
        )
    )
    .reset_index()
)

cantidad_global = (
    cantidad_global
    .sort_values(
        "CANTIDAD_TOTAL",
        ascending=False
    )
)

cantidad_global = cantidad_global[
    [
        "ITEM_SERVI",
        "CANTIDAD_TOTAL"
    ]
]

print(
    f"Registros cantidad global: {len(cantidad_global):,}"
)


# ==========================================================
# PARETO DE CANTIDADES
# ==========================================================

pareto_cantidades = cantidad_global.copy()

pareto_cantidades = (
    pareto_cantidades
    .sort_values(
        "CANTIDAD_TOTAL",
        ascending=False
    )
    .reset_index(drop=True)
)

pareto_cantidades["PORCENTAJE"] = (
    pareto_cantidades["CANTIDAD_TOTAL"]
    /
    pareto_cantidades["CANTIDAD_TOTAL"].sum()
    * 100
)

pareto_cantidades["PORCENTAJE_ACUMULADO"] = (
    pareto_cantidades["PORCENTAJE"]
    .cumsum()
)

pareto_cantidades["PORCENTAJE"] = (
    pareto_cantidades["PORCENTAJE"]
    .round(2)
)

pareto_cantidades["PORCENTAJE_ACUMULADO"] = (
    pareto_cantidades["PORCENTAJE_ACUMULADO"]
    .round(2)
)


# ==========================================================
# COMPARATIVO TERRITORIOS POR ITEM
# ==========================================================


carpeta_graficas = carpeta_territorios


os.makedirs(
    carpeta_graficas,
    exist_ok=True
)

cantidad_zona_recurso_item = (
    resumen.groupby(
        [
            "ZONA",
            "TIPO DE RECURSO",
            "ITEM_SERVI"
        ],
        dropna=False
    )
    .agg(
        CANTIDAD_TOTAL=(
            "CANTIDAD_TOTAL",
            "sum"
        )
    )
    .reset_index()
)


grafica_territorio = cantidad_zona_recurso_item.pivot_table(
    index="ITEM_SERVI",
    columns=["ZONA", "TIPO DE RECURSO"],
    values="CANTIDAD_TOTAL",
    aggfunc="sum",
    fill_value=0
)

# QUITAR ROCERIA
grafica_territorio = grafica_territorio[
    grafica_territorio.index !=
    "Intervencion forestal en cerco vivo y/o roceria"
]

grafica_territorio["TOTAL"] = (
    grafica_territorio.sum(axis=1)
)
grafica_territorio = (
    grafica_territorio
    .sort_values(
        "TOTAL",
        ascending=False
    )
)

grafica_territorio.drop(
    columns="TOTAL",
    inplace=True
)



ax = grafica_territorio.plot(
    kind="bar",
    figsize=(ANCHO_GRAFICA, ALTO_GRAFICA),
    width=0.7
)

aplicar_formato_grafica(
    ax,
    "Cantidad Ejecutada por Item y Territorio",
    "ITEM_SERVI",
    "Cantidad Total Ejecutada"
)

for container in ax.containers:

    ax.bar_label(
        container,
        fmt="%.0f",
        padding=1,
        fontsize=FONT_LABEL
    )

plt.tight_layout()

ruta_grafica = os.path.join(
    carpeta_graficas,
    "Comparativo_Territorios_Por_Item.png"
)

plt.savefig(
    ruta_grafica,
    dpi=DPI_GRAFICA,
    bbox_inches="tight"
)



plt.close()

print(
    f"Gráfica creada: {ruta_grafica}"
)

# ==========================================================
# CANTIDAD POR ZONA, RECURSO E ITEM
# ==========================================================

cantidad_zona_recurso_item = (
    resumen.groupby(
        [
            "ZONA",
            "TIPO DE RECURSO",
            "ITEM_SERVI"
        ],
        dropna=False
    )
    .agg(
        CANTIDAD_TOTAL=(
            "CANTIDAD_TOTAL",
            "sum"
        )
    )
    .reset_index()
)




cantidad_zona_recurso_item = (
    cantidad_zona_recurso_item
    .sort_values(
        [
            "ZONA",
            "TIPO DE RECURSO",
            "CANTIDAD_TOTAL"
        ],
        ascending=[
            True,
            True,
            False
        ]
    )
)

print(
    f"Registros cantidad zona recurso item: {len(cantidad_zona_recurso_item):,}"
)


# ==========================================================
# RESUMEN DESDE EL 15 DE ABRIL DE 2026
# ==========================================================

fecha_corte = pd.Timestamp("2026-04-15")

df_post_abril_2026 = df_promedios[
    df_promedios[col_fecha] >= fecha_corte
].copy()

resumen_post_abril_2026 = (
    df_post_abril_2026.groupby(
        [
            "ZONA",
            "ITEM_SERVI"
        ],
        dropna=False
    )
    .agg(
        CANTIDAD_TOTAL=(
            "CANTIDAD",
            "sum"
        ),
        LONGITUD_TOTAL_KM=(
            "LONGITUD_KM",
            "sum"
        )
    )
    .reset_index()
)

print(
    f"Registros resumen post abril 2026: {len(resumen_post_abril_2026)}"
)

# ==========================================================
# PRODUCTIVIDAD POR KILOMETRO
# ==========================================================

productividad_km = (
    resumen_post_abril_2026.copy()
)

productividad_km["UNIDADES_POR_KM"] = (
    productividad_km["CANTIDAD_TOTAL"]
    /
    productividad_km["LONGITUD_TOTAL_KM"]
)

productividad_km["UNIDADES_POR_KM"] = (
    productividad_km["UNIDADES_POR_KM"]
    .replace(
        [float("inf"), -float("inf")],
        0
    )
    .fillna(0)
    .round(2)
)

productividad_km = productividad_km.sort_values(
    "UNIDADES_POR_KM",
    ascending=False
)

print(
    f"Registros productividad por km: {len(productividad_km):,}"
)


# ==========================================================
# HORAS PROMEDIO POR ITEM Y GRUPO
# ==========================================================

carpeta_graficas = carpeta_horas

os.makedirs(
    carpeta_graficas,
    exist_ok=True
)

promedio_grupo = (
    resumen.groupby(
        [
            "ZONA",
            col_usuario,
            "TIPO DE RECURSO",
            "ITEM_SERVI"
        ],
        dropna=False
    )
    .agg(
        DURACION_TOTAL=("DURACION_TOTAL", "sum"),
        CANTIDAD_TOTAL=("CANTIDAD_TOTAL", "sum")
    )
    .reset_index()
)

promedio_grupo["HORAS_POR_ITEM"] = (
    promedio_grupo["DURACION_TOTAL"]
    /
    promedio_grupo["CANTIDAD_TOTAL"]
)

for zona in ZONAS:
    ruta_archivo = save_zone_bar_chart(
        promedio_grupo,
        zona,
        "HORAS_POR_ITEM",
        f"{zona}\nHoras Promedio por Item y Grupo",
        "Horas por Item",
        carpeta_horas,
        "{zona}_Horas_Promedio.png",
        col_usuario,
        grupos_camioneta,
        grupos_moto,
        decimals=2,
        width=0.78
    )

    if ruta_archivo is not None:
        print(f"Gráfica creada: {ruta_archivo}")

# ==========================================================
# COMPARATIVO ZONA VS ITEM VS DURACION
# ==========================================================



carpeta_graficas = carpeta_horas

os.makedirs(
    carpeta_graficas,
    exist_ok=True
)

print("\nVALIDACION GRAFICA")

print(
    promedio_territorio[
        promedio_territorio["ITEM_SERVI"]
        ==
        "Poda en redes Rurales Tipo B"
    ]
)


promedio_territorio_recurso = (
    resumen.groupby(
        [
            "ZONA",
            "TIPO DE RECURSO",
            "ITEM_SERVI"
        ],
        dropna=False
    )
    .agg(
        CANTIDAD_TOTAL=("CANTIDAD_TOTAL", "sum"),
        DURACION_TOTAL=("DURACION_TOTAL", "sum")
    )
    .reset_index()
)

promedio_territorio_recurso["HORAS_POR_ITEM_PROM"] = (
    promedio_territorio_recurso["DURACION_TOTAL"]
    /
    promedio_territorio_recurso["CANTIDAD_TOTAL"]
)

grafica = promedio_territorio_recurso.pivot_table(
    index="ITEM_SERVI",
    columns=["ZONA", "TIPO DE RECURSO"],
    values="HORAS_POR_ITEM_PROM",
    aggfunc="first",
    fill_value=0
)


# Ordenar por duración promedio total

grafica["TOTAL"] = grafica.max(axis=1)

grafica = grafica.sort_values(
    "TOTAL",
    ascending=False
)

grafica.drop(
    columns="TOTAL",
    inplace=True
)

# Crear gráfica

ax = grafica.plot(
    kind="bar",
    figsize=(ANCHO_GRAFICA, ALTO_GRAFICA)
)

aplicar_formato_grafica(
    ax,
    "Comparativo de Horas Promedio por Item y Territorio",
    "ITEM_SERVI",
    "Horas Promedio por Item"
)




# Etiquetas

for container in ax.containers:

    ax.bar_label(
        container,
        fmt="%.2f",
        padding=2,
        fontsize=FONT_LABEL
    )

plt.tight_layout()

ruta_archivo = os.path.join(
    carpeta_graficas,
    "Comparativo_Zonas_Duracion.png"
)

plt.savefig(
    ruta_archivo,
    dpi=DPI_GRAFICA,
    bbox_inches="tight"
)

plt.close()



plt.close()

print(
    f"Gráfica creada: {ruta_archivo}"
)

fecha_corte = pd.to_datetime("2026-04-15")

df_post_abril_2026 = df_promedios[
    df_promedios[col_fecha] >= fecha_corte
].copy()

resumen_post_abril_2026 = (
    df_post_abril_2026.groupby(
        [
            "ZONA",
            "ITEM_SERVI"
        ],
        dropna=False
    )
    .agg(
        CANTIDAD_TOTAL=("CANTIDAD", "sum"),
        LONGITUD_TOTAL_KM=("LONGITUD_KM", "sum")
    )
    .reset_index()
)

# ==========================================================
# COMPARATIVO KM POR ITEM Y ZONA
# ==========================================================


carpeta_graficas = carpeta_km

os.makedirs(
    carpeta_graficas,
    exist_ok=True
)

km_zona_recurso_item = (
    df_post_abril_2026.groupby(
        [
            "ZONA",
            "TIPO DE RECURSO",
            "ITEM_SERVI"
        ],
        dropna=False
    )
    .agg(
        LONGITUD_TOTAL_KM=(
            "LONGITUD_KM",
            "sum"
        )
    )
    .reset_index()
)

grafica_km_zona = km_zona_recurso_item.pivot_table(
    index="ITEM_SERVI",
    columns=["ZONA", "TIPO DE RECURSO"],
    values="LONGITUD_TOTAL_KM",
    aggfunc="sum",
    fill_value=0
)


grafica_km_zona["TOTAL"] = (
    grafica_km_zona.sum(axis=1)
)

grafica_km_zona = (
    grafica_km_zona
    .sort_values(
        "TOTAL",
        ascending=False
    )
)

grafica_km_zona.drop(
    columns="TOTAL",
    inplace=True
)

ax = grafica_km_zona.plot(
    kind="bar",
    figsize=(ANCHO_GRAFICA, ALTO_GRAFICA)
)

aplicar_formato_grafica(
    ax,
    "Kilómetros Intervenidos por Item, Territorio y Recurso",
    "ITEM_SERVI",
    "Kilómetros Intervenidos (KM)"
)

for container in ax.containers:

    ax.bar_label(
        container,
        fmt="%.2f",
        padding=2,
        fontsize=FONT_LABEL
    )

plt.tight_layout()

ruta_archivo = os.path.join(
    carpeta_horas,
    "Kilometros_Item_Territorio.png"
)

plt.savefig(
    ruta_archivo,
    dpi=DPI_GRAFICA,
    bbox_inches="tight"
)

plt.close()

print(
    f"Gráfica creada: {ruta_archivo}"
)

plt.close()

# ==========================================================
# KILOMETROS EJECUTADOS MES A MES
# ==========================================================

km_mes = (
    df_informe.groupby(
        pd.Grouper(
            key=col_fecha,
            freq="M"
        )
    )
    .agg(
        KM=("LONGITUD_KM", "sum")
    )
    .reset_index()
)

km_mes = km_mes[
    km_mes["FECHA"].notna()
].copy()

km_mes["MES"] = (
    km_mes["FECHA"]
    .dt.strftime("%Y-%m")
)

fecha_cambio = pd.Timestamp("2026-04-15")

promedio_antes = (
    km_mes.loc[
        km_mes["FECHA"] < fecha_cambio,
        "KM"
    ]
    .mean()
)

promedio_despues = (
    km_mes.loc[
        km_mes["FECHA"] >= fecha_cambio,
        "KM"
    ]
    .mean()
)

plt.figure(
    figsize=(16,8)
)

barras = plt.bar(
    km_mes["MES"],
    km_mes["KM"],
    color="#0E5A8A",
    width=0.40
)

for barra in barras:

    valor = barra.get_height()

    plt.text(
        barra.get_x() + barra.get_width()/2,
        valor + 2,
        f"{valor:.2f}".replace(".", ","),
        ha="center",
        fontsize=8,
        fontweight="bold"
    )

plt.axhline(
    promedio_antes,
    color="green",
    linestyle="--",
    linewidth=1.5,
    label=f"Promedio Antes: {promedio_antes:.2f} km"
)

plt.axhline(
    promedio_despues,
    color="red",
    linestyle="--",
    linewidth=1.5,
    label=f"Promedio Después: {promedio_despues:.2f} km"
)

indice_cambio = km_mes[
    km_mes["FECHA"] >= fecha_cambio
].index.min()

plt.axvline(
    indice_cambio - 0.5,
    color="red",
    linestyle="--",
    linewidth=1
)

plt.text(
    indice_cambio,
    km_mes["KM"].max() * 0.88,
    "Abr-2026\nInicio esquema\n2 Camionetas + 2 Motos",
    color="red",
    fontsize=9,
    fontweight="bold",
    bbox=dict(
        facecolor="white",
        edgecolor="red"
    )
)

plt.title(
    "Kilómetros Ejecutados Mes a Mes",
    fontsize=14,
    fontweight="bold"
)

plt.xlabel("Mes")
plt.ylabel("Kilómetros Ejecutados")

plt.xticks(
    rotation=45
)

plt.legend()

plt.grid(
    alpha=0.20
)

ruta_km_mes = os.path.join(
    carpeta_km,
    "KM_Mes_a_Mes.png"
)

plt.tight_layout()

plt.savefig(
    ruta_km_mes,
    dpi=300,
    bbox_inches="tight"
)

plt.close()

print(
    f"Gráfica creada: {ruta_km_mes}"
)



# ==========================================================
# RESUMEN DESDE 15-ABR-2026
# ==========================================================


resumen_post_abril_2026 = (
    df_post_abril_2026.groupby(
        [
            "ZONA",
            "ITEM_SERVI"
        ],
        dropna=False
    )
    .agg(
        CANTIDAD_TOTAL=(
            "CANTIDAD",
            "sum"
        ),
        LONGITUD_TOTAL_KM=(
            "LONGITUD_KM",
            "sum"
        )
    )
    .reset_index()
)

resumen_post_abril_2026["LONGITUD_TOTAL_KM"] = (
    resumen_post_abril_2026["LONGITUD_TOTAL_KM"]
    .round(3)
)

resumen_post_abril_2026 = (
    resumen_post_abril_2026
    .sort_values(
        [
            "ZONA",
            "CANTIDAD_TOTAL"
        ],
        ascending=[
            True,
            False
        ]
    )
)

print(
    f"Registros posteriores al 15-04-2026: {len(resumen_post_abril_2026):,}"
)


# ==========================================================
# KM POR ITEM Y GRUPO
# ==========================================================

km_grupo = (
    df_post_abril_2026.groupby(
        [
            "ZONA",
            col_usuario,
            "TIPO DE RECURSO",
            "ITEM_SERVI"
        ],
        dropna=False
    )
    .agg(
        LONGITUD_TOTAL_KM=(
            "LONGITUD_KM",
            "sum"
        )
    )
    .reset_index()
)

for zona in ["METROPOLITANA","NORTE","ORIENTE","SUR"]:

    datos_zona_resumen = km_grupo[
        km_grupo["ZONA"] == zona
    ].copy()

    if datos_zona_resumen.empty:
        continue

    tabla = datos_zona_resumen.pivot_table(
        index="ITEM_SERVI",
        columns=col_usuario,
        values="LONGITUD_TOTAL_KM",
        aggfunc="sum",
        fill_value=0
    )

    tabla["TOTAL"] = tabla.sum(axis=1)

    tabla = tabla.sort_values(
        "TOTAL",
        ascending=False
    )

    tabla.drop(
        columns="TOTAL",
        inplace=True
    )

    colores = []

    for grupo in tabla.columns:

        if grupo in grupos_camioneta:
            colores.append("#4472C4")

        elif grupo in grupos_moto:
            colores.append("#ED7D31")

        else:
            colores.append("gray")

    ax = tabla.plot(
        kind="bar",
        figsize=(ANCHO_GRAFICA, ALTO_GRAFICA),
        color=colores,
        width=0.78
    )

    aplicar_formato_grafica(
        ax,
        f"{zona}\nKilómetros Intervenidos por Item y Grupo",
        "Longitud Total (KM)",
        "ITEM_SERVI"
)

    for container in ax.containers:

        etiquetas = []

        for barra in container:

            valor = barra.get_height()

            if valor > 0:
                etiquetas.append(f"{valor:.1f}")
            else:
                etiquetas.append("")

        ax.bar_label(
            container,
            labels=etiquetas,
            fontsize=FONT_LABEL,
            padding=2
        )

    plt.legend(
     title="GRUPOS",
     fontsize=FONT_LEYENDA,
     title_fontsize=FONT_LEYENDA,
     bbox_to_anchor=(1.02,1),
     loc="upper left"
    )

    plt.tight_layout()

    ruta_archivo = os.path.join(
        carpeta_km,
        f"{zona}_KM_Comparativo_Grupos.png"
    )

    plt.savefig(
        ruta_archivo,
        dpi=DPI_GRAFICA,

        bbox_inches="tight"
    )

    plt.close()

    print(
        f"Gráfica creada: {ruta_archivo}"
    )

# ==========================================================
# PRODUCTIVIDAD POR KM Y ZONA
# ==========================================================

productividad_km_zona = (
    resumen_post_abril_2026.groupby(
        "ZONA",
        dropna=False
    )
    .agg(
        CANTIDAD_TOTAL=(
            "CANTIDAD_TOTAL",
            "sum"
        ),
        LONGITUD_TOTAL_KM=(
            "LONGITUD_TOTAL_KM",
            "sum"
        )
    )
    .reset_index()
)

productividad_km_zona["UNIDADES_POR_KM"] = (
    productividad_km_zona["CANTIDAD_TOTAL"]
    /
    productividad_km_zona["LONGITUD_TOTAL_KM"]
)

productividad_km_zona["UNIDADES_POR_KM"] = (
    productividad_km_zona["UNIDADES_POR_KM"]
    .replace(
        [float("inf"), -float("inf")],
        0
    )
    .fillna(0)
    .round(2)
)

productividad_km_zona = (
    productividad_km_zona
    .sort_values(
        "UNIDADES_POR_KM",
        ascending=False
    )
)




promedio_territorio_recurso = (
    resumen.groupby(
        [
            "ZONA",
            "TIPO DE RECURSO",
            "ITEM_SERVI"
        ],
        dropna=False
    )
    .agg(
        CANTIDAD_TOTAL=("CANTIDAD_TOTAL","sum"),
        DURACION_TOTAL=("DURACION_TOTAL","sum")
    )
    .reset_index()
)

promedio_territorio_recurso["HORAS_POR_ITEM_PROM"] = (
    promedio_territorio_recurso["DURACION_TOTAL"]
    /
    promedio_territorio_recurso["CANTIDAD_TOTAL"]
)




# ==========================================================
# PRODUCTIVIDAD POR KM Y ACTIVIDAD
# ==========================================================

productividad_km = (
    resumen_post_abril_2026.copy()
)

productividad_km["UNIDADES_POR_KM"] = (
    productividad_km["CANTIDAD_TOTAL"]
    /
    productividad_km["LONGITUD_TOTAL_KM"]
)

productividad_km["UNIDADES_POR_KM"] = (
    productividad_km["UNIDADES_POR_KM"]
    .replace(
        [float("inf"), -float("inf")],
        0
    )
    .fillna(0)
    .round(2)
)

productividad_km = productividad_km.sort_values(
    [
        "ZONA",
        "UNIDADES_POR_KM"
    ],
    ascending=[
        True,
        False
    ]
)





# ==========================================================
# COMPARATIVO DE GRUPOS POR TERRITORIO
# ==========================================================

cantidad_grupos = (
    resumen.groupby(
        [
            "ZONA",
            col_usuario,
            "TIPO DE RECURSO",
            "ITEM_SERVI"
        ],
        dropna=False
    )
    .agg(
        CANTIDAD_TOTAL=(
            "CANTIDAD_TOTAL",
            "sum"
        )
    )
    .reset_index()
)

for zona in ZONAS:
    ruta_archivo = save_zone_bar_chart(
        cantidad_grupos,
        zona,
        "CANTIDAD_TOTAL",
        f"{zona}\nCantidad Ejecutada por Item y Grupo",
        "Cantidad Ejecutada",
        carpeta_territorios,
        "{zona}_Comparativo_Grupos.png",
        col_usuario,
        grupos_camioneta,
        grupos_moto,
        decimals=0,
        width=0.78
    )

    if ruta_archivo is not None:
        print(f"Gráfica creada: {ruta_archivo}")


# ==========================================================
# DURACION POR UNIDAD EJECUTADA
# ZONA, RECURSO E ITEM
# CAMIONETA VS MOTO
# ==========================================================

print(
    "\nGenerando análisis de duración por unidad ejecutada..."
)

# ==========================================================
# CONFIGURACION
# ==========================================================

zonas_duracion = [
    "METROPOLITANA",
    "NORTE",
    "ORIENTE",
    "SUR"
]

recursos_duracion = [
    "CAMIONETA",
    "MOTO"
]

# ==========================================================
# PREPARAR INFORMACION
# ==========================================================

df_duracion_actividad = df_promedios.copy()

df_duracion_actividad["ZONA"] = (
    df_duracion_actividad["ZONA"]
    .astype("string")
    .str.strip()
    .str.upper()
)

df_duracion_actividad["TIPO DE RECURSO"] = (
    df_duracion_actividad["TIPO DE RECURSO"]
    .astype("string")
    .str.strip()
    .str.upper()
)

df_duracion_actividad["ITEM_SERVI"] = (
    df_duracion_actividad["ITEM_SERVI"]
    .astype("string")
    .str.strip()
)

df_duracion_actividad["CANTIDAD"] = pd.to_numeric(
    df_duracion_actividad["CANTIDAD"],
    errors="coerce"
)

df_duracion_actividad["DURA_ACT"] = pd.to_numeric(
    df_duracion_actividad["DURA_ACT"],
    errors="coerce"
)

df_duracion_actividad["LONGITUD_KM"] = pd.to_numeric(
    df_duracion_actividad["LONGITUD_KM"],
    errors="coerce"
)

# ==========================================================
# FILTRAR REGISTROS VALIDOS
# ==========================================================

df_duracion_actividad = df_duracion_actividad[
    df_duracion_actividad["ZONA"].isin(
        zonas_duracion
    )
    &
    df_duracion_actividad["TIPO DE RECURSO"].isin(
        recursos_duracion
    )
    &
    df_duracion_actividad["ITEM_SERVI"].notna()
    &
    df_duracion_actividad["CANTIDAD"].notna()
    &
    df_duracion_actividad["DURA_ACT"].notna()
    &
    (df_duracion_actividad["CANTIDAD"] > 0)
    &
    (df_duracion_actividad["DURA_ACT"] > 0)
].copy()

print(
    f"Registros válidos antes de separar rocería: "
    f"{len(df_duracion_actividad):,}"
)

if df_duracion_actividad.empty:

    raise Exception(
        "No existen registros válidos para generar el análisis "
        "de duración por unidad ejecutada."
    )

# ==========================================================
# SEPARAR ACTIVIDADES CONTROLADAS POR KILOMETROS
# ==========================================================
#
# La intervención forestal en cerco vivo y/o rocería no debe
# compararse en minutos por unidad porque se controla por KM.
# Se excluye de las gráficas de unidades y se conserva en una
# tabla independiente para análisis por kilómetro.
# ==========================================================

mascara_roceria = (
    df_duracion_actividad["ITEM_SERVI"]
    .str.contains(
        r"cerco vivo|roceria|rocería",
        case=False,
        na=False,
        regex=True
    )
)

df_duracion_roceria = (
    df_duracion_actividad[
        mascara_roceria
    ]
    .copy()
)

df_duracion_unidades = (
    df_duracion_actividad[
        ~mascara_roceria
    ]
    .copy()
)

print(
    f"Registros analizados por unidad: "
    f"{len(df_duracion_unidades):,}"
)

print(
    f"Registros separados para análisis por KM: "
    f"{len(df_duracion_roceria):,}"
)

if df_duracion_unidades.empty:

    raise Exception(
        "No quedaron registros de actividades medidas por unidad "
        "después de excluir la rocería."
    )

# ==========================================================
# CONSOLIDAR DURACION POR UNIDAD
# ==========================================================
#
# MINUTOS POR UNIDAD =
# SUMA DE DURA_ACT * 60 / SUMA DE CANTIDAD
#
# Ejemplo:
# 4 horas / 20 guaduas = 0.20 horas por guadua
# 0.20 * 60 = 12 minutos por guadua
# ==========================================================

duracion_actividad_zona_recurso = (
    df_duracion_unidades.groupby(
        [
            "ZONA",
            "TIPO DE RECURSO",
            "ITEM_SERVI"
        ],
        dropna=False
    )
    .agg(
        REGISTROS=(
            "DURA_ACT",
            "size"
        ),
        DIAS_TRABAJADOS=(
            col_fecha,
            lambda x: x.dt.date.nunique()
        ),
        CANTIDAD_TOTAL=(
            "CANTIDAD",
            "sum"
        ),
        DURACION_TOTAL_HORAS=(
            "DURA_ACT",
            "sum"
        )
    )
    .reset_index()
)

duracion_actividad_zona_recurso[
    "HORAS_POR_UNIDAD"
] = np.where(
    duracion_actividad_zona_recurso[
        "CANTIDAD_TOTAL"
    ] > 0,

    duracion_actividad_zona_recurso[
        "DURACION_TOTAL_HORAS"
    ]
    /
    duracion_actividad_zona_recurso[
        "CANTIDAD_TOTAL"
    ],

    np.nan
)

duracion_actividad_zona_recurso[
    "MINUTOS_POR_UNIDAD"
] = (
    duracion_actividad_zona_recurso[
        "HORAS_POR_UNIDAD"
    ]
    * 60
)

duracion_actividad_zona_recurso[
    "CANTIDAD_PROMEDIO_POR_REGISTRO"
] = np.where(
    duracion_actividad_zona_recurso[
        "REGISTROS"
    ] > 0,

    duracion_actividad_zona_recurso[
        "CANTIDAD_TOTAL"
    ]
    /
    duracion_actividad_zona_recurso[
        "REGISTROS"
    ],

    np.nan
)

duracion_actividad_zona_recurso[
    "DURACION_PROMEDIO_REGISTRO_HORAS"
] = np.where(
    duracion_actividad_zona_recurso[
        "REGISTROS"
    ] > 0,

    duracion_actividad_zona_recurso[
        "DURACION_TOTAL_HORAS"
    ]
    /
    duracion_actividad_zona_recurso[
        "REGISTROS"
    ],

    np.nan
)

# ==========================================================
# REDONDEAR RESULTADOS
# ==========================================================

columnas_redondear_unidad = [
    "CANTIDAD_TOTAL",
    "DURACION_TOTAL_HORAS",
    "HORAS_POR_UNIDAD",
    "MINUTOS_POR_UNIDAD",
    "CANTIDAD_PROMEDIO_POR_REGISTRO",
    "DURACION_PROMEDIO_REGISTRO_HORAS"
]

decimales_unidad = {
    "CANTIDAD_TOTAL": 2,
    "DURACION_TOTAL_HORAS": 2,
    "HORAS_POR_UNIDAD": 4,
    "MINUTOS_POR_UNIDAD": 2,
    "CANTIDAD_PROMEDIO_POR_REGISTRO": 2,
    "DURACION_PROMEDIO_REGISTRO_HORAS": 2
}

for columna in columnas_redondear_unidad:

    duracion_actividad_zona_recurso[columna] = (
        duracion_actividad_zona_recurso[columna]
        .round(
            decimales_unidad[columna]
        )
    )

duracion_actividad_zona_recurso = (
    duracion_actividad_zona_recurso
    .sort_values(
        [
            "ZONA",
            "MINUTOS_POR_UNIDAD"
        ],
        ascending=[
            True,
            False
        ]
    )
    .reset_index(
        drop=True
    )
)

print(
    f"Resultados consolidados por unidad: "
    f"{len(duracion_actividad_zona_recurso):,}"
)

# ==========================================================
# ANALISIS SEPARADO DE ROCERIA POR KILOMETRO
# ==========================================================

if not df_duracion_roceria.empty:

    duracion_roceria_km = (
        df_duracion_roceria.groupby(
            [
                "ZONA",
                "TIPO DE RECURSO",
                "ITEM_SERVI"
            ],
            dropna=False
        )
        .agg(
            REGISTROS=(
                "DURA_ACT",
                "size"
            ),
            DIAS_TRABAJADOS=(
                col_fecha,
                lambda x: x.dt.date.nunique()
            ),
            LONGITUD_TOTAL_KM=(
                "LONGITUD_KM",
                "sum"
            ),
            DURACION_TOTAL_HORAS=(
                "DURA_ACT",
                "sum"
            )
        )
        .reset_index()
    )

    duracion_roceria_km[
        "HORAS_POR_KM"
    ] = np.where(
        duracion_roceria_km[
            "LONGITUD_TOTAL_KM"
        ] > 0,

        duracion_roceria_km[
            "DURACION_TOTAL_HORAS"
        ]
        /
        duracion_roceria_km[
            "LONGITUD_TOTAL_KM"
        ],

        np.nan
    )

    duracion_roceria_km[
        "MINUTOS_POR_KM"
    ] = (
        duracion_roceria_km[
            "HORAS_POR_KM"
        ]
        * 60
    )

    duracion_roceria_km[
        "LONGITUD_TOTAL_KM"
    ] = (
        duracion_roceria_km[
            "LONGITUD_TOTAL_KM"
        ]
        .round(3)
    )

    duracion_roceria_km[
        "DURACION_TOTAL_HORAS"
    ] = (
        duracion_roceria_km[
            "DURACION_TOTAL_HORAS"
        ]
        .round(2)
    )

    duracion_roceria_km[
        "HORAS_POR_KM"
    ] = (
        duracion_roceria_km[
            "HORAS_POR_KM"
        ]
        .round(4)
    )

    duracion_roceria_km[
        "MINUTOS_POR_KM"
    ] = (
        duracion_roceria_km[
            "MINUTOS_POR_KM"
        ]
        .round(2)
    )

else:

    duracion_roceria_km = pd.DataFrame(
        columns=[
            "ZONA",
            "TIPO DE RECURSO",
            "ITEM_SERVI",
            "REGISTROS",
            "DIAS_TRABAJADOS",
            "LONGITUD_TOTAL_KM",
            "DURACION_TOTAL_HORAS",
            "HORAS_POR_KM",
            "MINUTOS_POR_KM"
        ]
    )

# ==========================================================
# TABLA COMPARATIVA CAMIONETA VS MOTO
# MINUTOS POR UNIDAD
# ==========================================================

tabla_comparativa_duracion = (
    duracion_actividad_zona_recurso
    .pivot_table(
        index=[
            "ZONA",
            "ITEM_SERVI"
        ],
        columns="TIPO DE RECURSO",
        values="MINUTOS_POR_UNIDAD",
        aggfunc="first"
    )
    .reset_index()
)

tabla_comparativa_duracion.columns.name = None

if "CAMIONETA" not in tabla_comparativa_duracion.columns:

    tabla_comparativa_duracion[
        "CAMIONETA"
    ] = np.nan

if "MOTO" not in tabla_comparativa_duracion.columns:

    tabla_comparativa_duracion[
        "MOTO"
    ] = np.nan

# Una diferencia positiva significa que la moto tardó más.
# Una diferencia negativa significa que la moto tardó menos.

tabla_comparativa_duracion[
    "DIFERENCIA_MINUTOS_MOTO_MENOS_CAMIONETA"
] = (
    tabla_comparativa_duracion["MOTO"]
    -
    tabla_comparativa_duracion["CAMIONETA"]
)

tabla_comparativa_duracion[
    "DIFERENCIA_ABSOLUTA_MINUTOS"
] = (
    tabla_comparativa_duracion[
        "DIFERENCIA_MINUTOS_MOTO_MENOS_CAMIONETA"
    ]
    .abs()
)

tabla_comparativa_duracion[
    "DIFERENCIA_PORCENTUAL"
] = np.where(
    tabla_comparativa_duracion[
        "CAMIONETA"
    ] > 0,

    (
        tabla_comparativa_duracion["MOTO"]
        -
        tabla_comparativa_duracion["CAMIONETA"]
    )
    /
    tabla_comparativa_duracion["CAMIONETA"]
    * 100,

    np.nan
)

# ==========================================================
# IDENTIFICAR RECURSO CON MENOR DURACION
# ==========================================================

ambos_recursos = (
    tabla_comparativa_duracion["CAMIONETA"].notna()
    &
    tabla_comparativa_duracion["MOTO"].notna()
)

tiempo_similar = (
    ambos_recursos
    &
    np.isclose(
        tabla_comparativa_duracion["CAMIONETA"],
        tabla_comparativa_duracion["MOTO"],
        atol=0.10
    )
)

condiciones_duracion = [
    tiempo_similar,

    (
        ambos_recursos
        &
        ~tiempo_similar
        &
        (
            tabla_comparativa_duracion["CAMIONETA"]
            <
            tabla_comparativa_duracion["MOTO"]
        )
    ),

    (
        ambos_recursos
        &
        ~tiempo_similar
        &
        (
            tabla_comparativa_duracion["MOTO"]
            <
            tabla_comparativa_duracion["CAMIONETA"]
        )
    )
]

resultados_duracion = [
    "TIEMPO SIMILAR",
    "CAMIONETA",
    "MOTO"
]



tabla_comparativa_duracion[
    "RECURSO_MENOR_DURACION"
] = np.select(
    condiciones_duracion,
    resultados_duracion,
    default="SIN COMPARACION"
)

columnas_comparativas_redondear = [
    "CAMIONETA",
    "MOTO",
    "DIFERENCIA_MINUTOS_MOTO_MENOS_CAMIONETA",
    "DIFERENCIA_ABSOLUTA_MINUTOS",
    "DIFERENCIA_PORCENTUAL"
]

for columna in columnas_comparativas_redondear:

    tabla_comparativa_duracion[columna] = (
        tabla_comparativa_duracion[columna]
        .round(2)
    )

tabla_comparativa_duracion = (
    tabla_comparativa_duracion
    .sort_values(
        [
            "ZONA",
            "DIFERENCIA_ABSOLUTA_MINUTOS"
        ],
        ascending=[
            True,
            False
        ]
    )
    .reset_index(
        drop=True
    )
)

# ==========================================================
# CREAR CUATRO GRAFICAS
# UNA POR CADA ZONA
# ==========================================================

graficas_duracion_generadas = []

for zona in zonas_duracion:

    datos_zona_resumen = (
        duracion_actividad_zona_recurso[
            duracion_actividad_zona_recurso["ZONA"] == zona
        ]
        .copy()
    )

    if datos_zona_resumen.empty:

        print(
            f"No existen datos de duración por unidad para {zona}."
        )

        continue

    tabla_grafica_duracion = datos_zona_resumen.pivot_table(
        index="ITEM_SERVI",
        columns="TIPO DE RECURSO",
        values="MINUTOS_POR_UNIDAD",
        aggfunc="first"
    )

    for recurso in recursos_duracion:

        if recurso not in tabla_grafica_duracion.columns:

            tabla_grafica_duracion[
                recurso
            ] = np.nan

    tabla_grafica_duracion = tabla_grafica_duracion[
        [
            "CAMIONETA",
            "MOTO"
        ]
    ]

    tabla_grafica_duracion["ORDEN"] = (
        tabla_grafica_duracion[
            [
                "CAMIONETA",
                "MOTO"
            ]
        ]
        .max(axis=1)
    )

    tabla_grafica_duracion = (
        tabla_grafica_duracion
        .sort_values(
            "ORDEN",
            ascending=False
        )
    )

    tabla_grafica_duracion.drop(
        columns="ORDEN",
        inplace=True
    )

    cantidad_items = len(
        tabla_grafica_duracion
    )

    altura_grafica = max(
        8,
        cantidad_items * 0.70
    )

    ax = tabla_grafica_duracion.plot(
        kind="bar",
        figsize=(
        ANCHO_GRAFICA,
        altura_grafica
        ),
        color=[
        "#4472C4",
        "#ED7D31"
        ],
        width=0.78
    )

    plt.xticks(
    rotation=45,
    ha="right"
)

    plt.title(
        f"{zona}\n"
        "Tiempo promedio por unidad ejecutada\n"
        "Camioneta vs Moto",
        fontsize=15,
        fontweight="bold"
    )

    plt.xlabel(
        "Minutos promedio por unidad ejecutada"
    )

    plt.ylabel(
        "Actividad"
    )

    plt.legend(
        title="Tipo de recurso"
    )

    plt.grid(
        axis="x",
        linestyle="--",
        alpha=0.30
    )

    for contenedor in ax.containers:

        etiquetas = []

        for barra in contenedor:

            valor_minutos = barra.get_width()

            if (
                pd.isna(valor_minutos)
                or valor_minutos <= 0
            ):

                etiquetas.append("")

            else:

                etiquetas.append(
                    f"{valor_minutos:.1f} min"
                )

        ax.bar_label(
            contenedor,
            labels=etiquetas,
            padding=3,
            fontsize=FONT_LABEL
        )

    plt.tight_layout()

    nombre_grafica_duracion = (
        f"{zona}_Minutos_Por_Unidad_Camioneta_vs_Moto.png"
    )

    ruta_grafica_duracion = os.path.join(
    carpeta_horas,
    nombre_grafica_duracion
    )

    plt.savefig(
        ruta_grafica_duracion,
        dpi=DPI_GRAFICA,

        bbox_inches="tight"
    )

    plt.close()

    graficas_duracion_generadas.append(
        nombre_grafica_duracion
    )

    print(
        f"Gráfica de minutos por unidad creada: "
        f"{ruta_grafica_duracion}"
    )

# ==========================================================
# CONCLUSIONES AUTOMATICAS
# ==========================================================

texto_conclusiones_duracion = (
    "ANALISIS DE TIEMPO POR UNIDAD EJECUTADA\n"
    "COMPARATIVO CAMIONETA VS MOTO POR ZONA\n"
    + "=" * 90
    + "\n\n"
)

texto_conclusiones_duracion += (
    "OBJETIVO\n\n"
    "Comparar el tiempo promedio requerido para ejecutar una "
    "unidad de cada actividad, según el territorio y el tipo "
    "de recurso utilizado.\n\n"
)

texto_conclusiones_duracion += (
    "METODO DE CALCULO\n\n"
    "El tiempo por unidad se calculó dividiendo la suma de "
    "DURA_ACT entre la suma de CANTIDAD para cada combinación "
    "de zona, actividad y tipo de recurso.\n\n"
    "MINUTOS POR UNIDAD = "
    "DURACION TOTAL EN HORAS * 60 / CANTIDAD TOTAL\n\n"
    "Para tala de guaduas, el resultado corresponde al tiempo "
    "estimado por guadua y no a la duración completa del "
    "registro o de la jornada.\n\n"
    "La intervención forestal en cerco vivo y/o rocería fue "
    "excluida de este comparativo por unidades. Esta actividad "
    "se analiza de forma independiente mediante horas y minutos "
    "por kilómetro.\n\n"
)

for zona in zonas_duracion:

    comparacion_zona = (
        tabla_comparativa_duracion[
            tabla_comparativa_duracion["ZONA"] == zona
        ]
        .copy()
    )

    texto_conclusiones_duracion += (
        "\n"
        + "=" * 90
        + f"\nTERRITORIO: {zona}\n"
        + "=" * 90
        + "\n\n"
    )

    if comparacion_zona.empty:

        texto_conclusiones_duracion += (
            "No se encontraron registros válidos para este "
            "territorio.\n\n"
        )

        continue

    actividades_comparables = comparacion_zona[
        comparacion_zona["CAMIONETA"].notna()
        &
        comparacion_zona["MOTO"].notna()
    ].copy()

    if actividades_comparables.empty:

        texto_conclusiones_duracion += (
            "No existen actividades con información simultánea "
            "de CAMIONETA y MOTO. No es posible establecer una "
            "comparación directa entre los recursos.\n\n"
        )

    else:

        camioneta_menor = int(
            (
                actividades_comparables[
                    "RECURSO_MENOR_DURACION"
                ] == "CAMIONETA"
            )
            .sum()
        )

        moto_menor = int(
            (
                actividades_comparables[
                    "RECURSO_MENOR_DURACION"
                ] == "MOTO"
            )
            .sum()
        )

        tiempos_similares = int(
            (
                actividades_comparables[
                    "RECURSO_MENOR_DURACION"
                ] == "TIEMPO SIMILAR"
            )
            .sum()
        )

        texto_conclusiones_duracion += (
            "CONCLUSION GENERAL DEL TERRITORIO\n\n"
            f"Se compararon {len(actividades_comparables)} "
            f"actividades con información disponible para ambos "
            f"recursos.\n\n"
            f"- CAMIONETA presentó menor tiempo por unidad en "
            f"{camioneta_menor} actividades.\n"
            f"- MOTO presentó menor tiempo por unidad en "
            f"{moto_menor} actividades.\n"
            f"- Se observaron tiempos similares en "
            f"{tiempos_similares} actividades.\n\n"
        )

        actividades_comparables = (
            actividades_comparables
            .sort_values(
                "DIFERENCIA_ABSOLUTA_MINUTOS",
                ascending=False
            )
        )

        mayor_diferencia = (
            actividades_comparables
            .iloc[0]
        )

        texto_conclusiones_duracion += (
            "MAYOR DIFERENCIA ENTRE RECURSOS\n\n"
            f"Actividad: "
            f"{mayor_diferencia['ITEM_SERVI']}\n"
            f"- CAMIONETA: "
            f"{mayor_diferencia['CAMIONETA']:.1f} "
            f"minutos por unidad.\n"
            f"- MOTO: "
            f"{mayor_diferencia['MOTO']:.1f} "
            f"minutos por unidad.\n"
            f"- Diferencia absoluta: "
            f"{mayor_diferencia['DIFERENCIA_ABSOLUTA_MINUTOS']:.1f} "
            f"minutos por unidad.\n"
            f"- Recurso con menor tiempo: "
            f"{mayor_diferencia['RECURSO_MENOR_DURACION']}.\n\n"
        )

        texto_conclusiones_duracion += (
            "DETALLE POR ACTIVIDAD\n\n"
        )

        for _, actividad in actividades_comparables.iterrows():

            diferencia = actividad[
                "DIFERENCIA_MINUTOS_MOTO_MENOS_CAMIONETA"
            ]

            if diferencia > 0:

                interpretacion = (
                    f"La CAMIONETA empleó "
                    f"{abs(diferencia):.1f} minutos menos "
                    f"por unidad que la MOTO."
                )

            elif diferencia < 0:

                interpretacion = (
                    f"La MOTO empleó "
                    f"{abs(diferencia):.1f} minutos menos "
                    f"por unidad que la CAMIONETA."
                )

            else:

                interpretacion = (
                    "Los dos recursos presentaron un tiempo similar."
                )

            texto_conclusiones_duracion += (
                f"Actividad: {actividad['ITEM_SERVI']}\n"
                f"- CAMIONETA: "
                f"{actividad['CAMIONETA']:.1f} minutos por unidad.\n"
                f"- MOTO: "
                f"{actividad['MOTO']:.1f} minutos por unidad.\n"
                f"- Resultado: {interpretacion}\n\n"
            )

    actividades_sin_comparacion = comparacion_zona[
        comparacion_zona["CAMIONETA"].isna()
        |
        comparacion_zona["MOTO"].isna()
    ].copy()

    if not actividades_sin_comparacion.empty:

        texto_conclusiones_duracion += (
            "ACTIVIDADES SIN COMPARACION DIRECTA\n\n"
        )

        for _, actividad in actividades_sin_comparacion.iterrows():

            if pd.notna(
                actividad["CAMIONETA"]
            ):

                recurso_disponible = "CAMIONETA"

            elif pd.notna(
                actividad["MOTO"]
            ):

                recurso_disponible = "MOTO"

            else:

                recurso_disponible = "SIN INFORMACION"

            texto_conclusiones_duracion += (
                f"- {actividad['ITEM_SERVI']}: solamente tiene "
                f"información para {recurso_disponible}.\n"
            )

        texto_conclusiones_duracion += "\n"

texto_conclusiones_duracion += (
    "\n"
    + "=" * 90
    + "\nNOTA TECNICA PARA INTERVENTORIA\n"
    + "=" * 90
    + "\n\n"
    "1. El indicador representa minutos promedio por unidad "
    "ejecutada y no la duración completa de una jornada.\n\n"
    "2. Una menor duración no debe interpretarse automáticamente "
    "como mayor productividad o mejor cumplimiento contractual.\n\n"
    "3. Las comparaciones deben realizarse para la misma actividad, "
    "zona y período operativo.\n\n"
    "4. Los resultados deben revisarse junto con el número de "
    "registros, cantidad ejecutada, días trabajados y condiciones "
    "operativas.\n\n"
    "5. La rocería se analiza separadamente por kilómetros y no "
    "por unidades ejecutadas.\n\n"
    "6. Los resultados se actualizan automáticamente cuando se "
    "incorporan nuevos registros.\n"
)

# ==========================================================
# GUARDAR CONCLUSIONES
# ==========================================================

ruta_conclusiones_duracion = os.path.join(
    carpeta_reportes,
    "Conclusiones_Minutos_Por_Unidad_Zona_Recurso.txt"
)

with open(
    ruta_conclusiones_duracion,
    "w",
    encoding="utf-8"
) as archivo:

    archivo.write(
        texto_conclusiones_duracion
    )

# ==========================================================
# GUARDAR EXCEL
# ==========================================================

ruta_excel_duracion = os.path.join(
    carpeta_reportes,
    "Analisis_Minutos_Por_Unidad_Zona_Recurso.xlsx"
)

with pd.ExcelWriter(
    ruta_excel_duracion,
    engine="openpyxl"
) as writer:

    duracion_actividad_zona_recurso.to_excel(
        writer,
        sheet_name="Minutos_Por_Unidad",
        index=False
    )

    tabla_comparativa_duracion.to_excel(
        writer,
        sheet_name="Comparativo_Recursos",
        index=False
    )

    duracion_roceria_km.to_excel(
        writer,
        sheet_name="Roceria_Por_KM",
        index=False
    )

    df_duracion_unidades[
        [
            col_fecha,
            col_usuario,
            "ZONA",
            "TIPO DE RECURSO",
            "ITEM_SERVI",
            "CANTIDAD",
            "DURA_ACT"
        ]
    ].to_excel(
        writer,
        sheet_name="Datos_Unidades",
        index=False
    )

print(
    f"Conclusiones guardadas en: "
    f"{ruta_conclusiones_duracion}"
)



print(
    f"Excel guardado en: "
    f"{ruta_excel_duracion}"
)

# ==========================================================
# RESUMEN EJECUTIVO GENERAL
# ==========================================================



resumen_ejecutivo = ""

# ==========================================================
# CONCLUSIONES EJECUTIVAS AUTOMATICAS
# ==========================================================

total_general = int(
    cantidad_global["CANTIDAD_TOTAL"].sum()
)

actividad_top = (
    cantidad_global
    .sort_values(
        "CANTIDAD_TOTAL",
        ascending=False
    )
    .iloc[0]
)

item_top = actividad_top["ITEM_SERVI"]

cantidad_top = int(
    actividad_top["CANTIDAD_TOTAL"]
)

territorio_top = (
    tabla_cantidad_territorio
    .groupby("ZONA")["CANTIDAD_TOTAL"]
    .sum()
    .sort_values(
        ascending=False
    )
)

zona_lider = territorio_top.index[0]

cantidad_zona = int(
    territorio_top.iloc[0]
)

descripcion_grafica = f"""
DESCRIPCION DE LA GRAFICA

La gráfica presenta la cantidad total ejecutada por actividad y territorio para el período analizado.

Se comparan los volúmenes de ejecución de cada ITEM_SERVI en los territorios Metropolitana,
Norte, Oriente y Sur, permitiendo identificar la distribución operativa de las actividades
forestales y de mantenimiento de vegetación asociadas a la infraestructura eléctrica.

Los resultados corresponden a la información consolidada disponible al momento de generación
del informe y variarán automáticamente de acuerdo con los datos procesados en cada ejecución.
"""

conclusion_grafica = f"""
CONCLUSION EJECUTIVA

Durante el período analizado se registraron {total_general:,.0f} actividades ejecutadas.

La actividad con mayor participación fue:

{item_top}

con un total de {cantidad_top:,.0f} registros.

El territorio con mayor volumen de ejecución fue:

{zona_lider}

con {cantidad_zona:,.0f} actividades ejecutadas.

Estos resultados cambian automáticamente en cada ejecución del informe,
debido a que se calculan directamente a partir de los datos más recientes disponibles.
"""

nota_tecnica = """
NOTA TECNICA

La actividad 'Intervencion forestal en cerco vivo y/o roceria'
se remunera y controla con base en kilometros lineales intervenidos.

Por esta razon, los valores reportados para esta actividad no deben
compararse directamente con actividades medidas por unidades ejecutadas
como podas, talas o guaduas.

Para un analisis integral se recomienda revisar conjuntamente los
indicadores de longitud intervenida y productividad por kilometro.
"""

resumen_ejecutivo += f"""

============================================================
CARPETA:
02_Horas_Territorio

ARCHIVOS:
METROPOLITANA_Horas_Promedio.png
NORTE_Horas_Promedio.png
ORIENTE_Horas_Promedio.png
SUR_Horas_Promedio.png
Comparativo_Zonas_Duracion.png
============================================================

TITULO:
Cantidad Ejecutada por Item y Territorio

{descripcion_grafica}

{conclusion_grafica}

{nota_tecnica}

"""

# ==========================================================
# DATAFRAME DE CONCLUSIONES
# ==========================================================

df_conclusiones = pd.DataFrame({

    "SECCION": [

        "DESCRIPCION",

        "CONCLUSION EJECUTIVA",

        "NOTA TECNICA"

    ],

    "TEXTO": [

        descripcion_grafica,

        conclusion_grafica,

        nota_tecnica

    ]

})

print("Generando archivo final...")

with pd.ExcelWriter(
    archivo_salida,
    engine="openpyxl"
) as writer:

    # ======================================================
    # HOJA CONCLUSIONES
    # ======================================================

    df_conclusiones.to_excel(
        writer,
        sheet_name="Resumen_Ejecutivo",
        index=False
    )

    # ======================================================
    # HOJAS DEL INFORME
    # ======================================================

    df_informe.to_excel(
        writer,
        sheet_name="Informe",
        index=False
    )

    df_33.to_excel(
        writer,
        sheet_name="33KV",
        index=False
    )

    df_115.to_excel(
        writer,
        sheet_name="115 KV",
        index=False
    )

    # AQUÍ DEBEN CONTINUAR TODAS LAS DEMÁS HOJAS

    resumen.to_excel(
        writer,
        sheet_name="Promedios",
        index=False
    )

    promedio_territorio.to_excel(
        writer,
        sheet_name="Promedio_Territorio",
        index=False
    )

    promedio_tipo_recurso.to_excel(
        writer,
        sheet_name="Promedio_Tipo_Recurso",
        index=False
    )

    promedio_global.to_excel(
        writer,
        sheet_name="Promedio_Global",
        index=False
    )


# ==========================================================
# CONCLUSIONES COMPARATIVO HORAS POR ITEM Y TERRITORIO
# ==========================================================

actividad_mayor = (
    promedio_territorio
    .sort_values(
        "HORAS_POR_ITEM_PROM",
        ascending=False
    )
    .iloc[0]
)

actividad_menor = (
    promedio_territorio
    .sort_values(
        "HORAS_POR_ITEM_PROM",
        ascending=True
    )
    .iloc[0]
)

# ==========================================================
# PROMEDIO TERRITORIO RECURSO
# ==========================================================

promedio_territorio_recurso = (
    resumen.groupby(
        [
            "ZONA",
            "TIPO DE RECURSO",
            "ITEM_SERVI"
        ],
        dropna=False
    )
    .agg(
        CANTIDAD_TOTAL=(
            "CANTIDAD_TOTAL",
            "sum"
        ),
        DURACION_TOTAL=(
            "DURACION_TOTAL",
            "sum"
        )
    )
    .reset_index()
)

promedio_territorio_recurso["HORAS_POR_ITEM_PROM"] = (
    promedio_territorio_recurso["DURACION_TOTAL"]
    /
    promedio_territorio_recurso["CANTIDAD_TOTAL"]
)

promedio_territorio_recurso["MINUTOS_POR_ITEM_PROM"] = (
    promedio_territorio_recurso["HORAS_POR_ITEM_PROM"]
    * 60
)



descripcion_horas = """
DESCRIPCION

La gráfica presenta el comportamiento de las horas promedio requeridas para la ejecución de cada actividad en los diferentes territorios operativos.

El análisis permite comparar los tiempos de ejecución entre Metropolitana, Norte, Oriente y Sur, identificando diferencias operativas y variaciones en el desempeño de las labores forestales ejecutadas.
"""

conclusion_horas = f"""
CONCLUSION EJECUTIVA

La actividad que registró el mayor tiempo promedio de ejecución fue:

{actividad_mayor['ITEM_SERVI']}

en el territorio {actividad_mayor['ZONA']},

con un promedio de {actividad_mayor['HORAS_POR_ITEM_PROM']:.2f} horas por ítem

Por su parte, la actividad con menor tiempo promedio de ejecución fue:

{actividad_menor['ITEM_SERVI']}

en el territorio {actividad_menor['ZONA']},

con un promedio de {actividad_menor['HORAS_POR_ITEM_PROM']:.2f} horas por ítem.

Los resultados reflejan diferencias operativas entre actividades y territorios, asociadas a factores como complejidad técnica, condiciones de acceso, características de la vegetación y necesidades específicas de intervención.

Estos valores se recalculan automáticamente en cada ejecución del informe y pueden variar de acuerdo con la información procesada.
"""

nota_horas = """
NOTA TECNICA

La actividad 'Intervencion forestal en cerco vivo y/o roceria'
se remunera con base en kilometros lineales intervenidos.

Por esta razon, los tiempos promedio de esta actividad no deben compararse directamente con actividades medidas por unidades ejecutadas, como podas, talas o guaduas.

Para un analisis integral se recomienda revisar simultaneamente los indicadores de longitud intervenida y productividad por kilometro.
"""

resumen_ejecutivo += f"""

============================================================
GRAFICAS:
METROPOLITANA_Horas_Promedio.png
NORTE_Horas_Promedio.png
ORIENTE_Horas_Promedio.png
SUR_Horas_Promedio.png
============================================================

TITULO:
Horas Promedio por Item y Territorio

{descripcion_horas}

{conclusion_horas}

{nota_horas}

"""

ruta_conclusion_horas = os.path.join(
    carpeta_reportes,
    "Conclusiones_Horas_Por_Item.txt"
)

with open(
    ruta_conclusion_horas,
    "w",
    encoding="utf-8"
) as archivo:

    archivo.write(descripcion_horas)
    archivo.write("\n\n")
    archivo.write(conclusion_horas)
    archivo.write("\n\n")
    archivo.write(nota_horas)

print(
    f"Conclusiones guardadas: {ruta_conclusion_horas}"
)


# ==========================================================
# CONCLUSIONES KM POR ITEM Y TERRITORIO
# ==========================================================

km_mayor = (
    resumen_post_abril_2026
    .sort_values(
        "LONGITUD_TOTAL_KM",
        ascending=False
    )
    .iloc[0]
)

km_menor = (
    resumen_post_abril_2026[
        resumen_post_abril_2026["LONGITUD_TOTAL_KM"] > 0
    ]
    .sort_values(
        "LONGITUD_TOTAL_KM",
        ascending=True
    )
    .iloc[0]
)

descripcion_km = """
DESCRIPCION

La gráfica presenta la longitud total intervenida en kilómetros para cada actividad y territorio.

El análisis permite comparar la participación de los diferentes territorios en términos de longitud intervenida, identificando las actividades que concentran la mayor cobertura operativa dentro del período evaluado.

Los resultados reflejan la distribución territorial de las actividades forestales y de mantenimiento de vegetación registradas en el sistema.
"""

conclusion_km = f"""
CONCLUSION EJECUTIVA

La actividad con mayor longitud intervenida correspondió a:

{km_mayor['ITEM_SERVI']}

en el territorio {km_mayor['ZONA']},

registrando un total de {km_mayor['LONGITUD_TOTAL_KM']:.2f} kilómetros intervenidos.

Por otra parte, la menor longitud intervenida se presentó en:

{km_menor['ITEM_SERVI']}

en el territorio {km_menor['ZONA']},

con un total de {km_menor['LONGITUD_TOTAL_KM']:.2f} kilómetros.

Los resultados evidencian diferencias en la distribución territorial de las actividades ejecutadas, reflejando las necesidades operativas, condiciones de vegetación y requerimientos de mantenimiento presentes en cada zona.

Estos valores se recalculan automáticamente cada vez que se ejecuta el informe y pueden variar conforme se incorporen nuevos registros operativos.
"""

nota_km = """
NOTA TECNICA

La actividad 'Intervencion forestal en cerco vivo y/o roceria'
se mide y remunera mediante kilometros lineales intervenidos.

Por esta razon, su participacion dentro de los indicadores de longitud
constituye una medida directa de ejecucion y cobertura operativa.

Las longitudes reportadas corresponden a la informacion disponible
para el periodo analizado y pueden variar en futuras actualizaciones
del informe.
"""

resumen_ejecutivo += f"""

============================================================
============================================================
CARPETA:
03_Kilometros

ARCHIVOS:
KM_Item_Zona.png
KM_Total_Item.png
============================================================
============================================================

TITULO:
Kilómetros Intervenidos por Item - {km_mayor['ZONA']}

{descripcion_km}

{conclusion_km}

{nota_km}

"""


# ==========================================================
# GUARDAR ARCHIVO
# ==========================================================

ruta_conclusion_km = os.path.join(
    carpeta_reportes,
    "Conclusiones_KM_Item_Territorio.txt"
)

with open(
    ruta_conclusion_km,
    "w",
    encoding="utf-8"
) as archivo:

    archivo.write(descripcion_km)
    archivo.write("\n\n")
    archivo.write(conclusion_km)
    archivo.write("\n\n")
    archivo.write(nota_km)

print(
    f"Conclusiones guardadas: {ruta_conclusion_km}"
)


# ==========================================================
# CONCLUSIONES KM POR ZONA
# ==========================================================

ruta_conclusiones_zonas = os.path.join(
    carpeta_reportes,
    "Conclusiones_KM_Por_Zona.txt"
)

with open(
    ruta_conclusiones_zonas,
    "w",
    encoding="utf-8"
) as archivo:

    archivo.write(
        "CONCLUSIONES DE KILOMETROS INTERVENIDOS POR TERRITORIO\n\n"
    )

    zonas = sorted(
        resumen_post_abril_2026["ZONA"]
        .dropna()
        .unique()
    )

    for zona in zonas:

        datos_zona_resumen = (
            resumen_post_abril_2026[
                resumen_post_abril_2026["ZONA"] == zona
            ]
            .copy()
        )

        if len(datos_zona_resumen) == 0:
            continue

        actividad_mayor = (
            datos_zona_resumen
            .sort_values(
                "LONGITUD_TOTAL_KM",
                ascending=False
            )
            .iloc[0]
        )

        actividad_menor = (
            datos_zona_resumen[
                datos_zona_resumen["LONGITUD_TOTAL_KM"] > 0
            ]
            .sort_values(
                "LONGITUD_TOTAL_KM",
                ascending=True
            )
            .iloc[0]
        )

        total_km = (
            datos_zona_resumen["LONGITUD_TOTAL_KM"]
            .sum()
        )

        conclusion = f"""
=============================================================
TERRITORIO: {zona}
=============================================================

Durante el período analizado se intervinieron
{total_km:.1f} kilómetros de vegetación.

La actividad con mayor participación fue:

{actividad_mayor['ITEM_SERVI']}

con una longitud total intervenida de
{actividad_mayor['LONGITUD_TOTAL_KM']:.1f} km.

Por su parte, la actividad con menor participación fue:

{actividad_menor['ITEM_SERVI']}

con una longitud total intervenida de
{actividad_menor['LONGITUD_TOTAL_KM']:.1f} km.

Los resultados reflejan la distribución de los esfuerzos
operativos en el territorio y permiten identificar las
actividades que concentraron la mayor cobertura de intervención
durante el período evaluado.

Estos valores se recalculan automáticamente cada vez que se
ejecuta el informe y pueden variar según la programación de
mantenimiento, las condiciones de vegetación y los registros
incorporados al sistema.

"""

        archivo.write(conclusion)

    archivo.write(
        "\nNOTA TECNICA:\n"
    )

    archivo.write(
        "\nLa actividad 'Intervencion forestal en cerco vivo y/o roceria'"
        "\nse mide y remunera mediante kilometros lineales intervenidos."
        "\nPor esta razon debe analizarse principalmente mediante"
        "\nindicadores de cobertura y productividad por kilometro.\n"
    )

print(
    f"Conclusiones guardadas: {ruta_conclusiones_zonas}"
)







# ==========================================================
# GUARDAR RESUMEN EJECUTIVO COMPLETO
# ==========================================================

ruta_resumen = os.path.join(
    carpeta_reportes,
    "Resumen_Ejecutivo_Completo.txt"
)

with open(
    ruta_resumen,
    "w",
    encoding="utf-8"
) as archivo:

    archivo.write(
        "RESUMEN EJECUTIVO AUTOMATICO\n"
    )

    archivo.write(
        "=" * 100
    )

    archivo.write("\n\n")

    archivo.write(
        resumen_ejecutivo
    )

print(
    f"Resumen ejecutivo generado: {ruta_resumen}"
)

# ==========================================================
# DATAFRAME PARA CONCLUSIONES EJECUTIVAS DEL CORREO
# ==========================================================

items_excluir = [
    "Intervencion forestal en cerco vivo y/o roceria"
]

df_correo = df_promedios[
    ~df_promedios["ITEM_SERVI"].isin(items_excluir)
].copy()

resumen_zona = (
    df_correo.groupby(
        [
            "ZONA",
            "ITEM_SERVI"
        ]
    )
    .agg(
        CANTIDAD_TOTAL=("CANTIDAD","sum"),
        KM_TOTAL=("LONGITUD_KM","sum"),
        HORAS_TOTAL=("DURA_ACT","sum")
    )
    .reset_index()
)

resumen_zona["HORAS_POR_ITEM"] = np.where(
    resumen_zona["CANTIDAD_TOTAL"] > 0,
    resumen_zona["HORAS_TOTAL"]
    /
    resumen_zona["CANTIDAD_TOTAL"],
    np.nan
)

resumen_recurso = (
    df_correo.groupby(
        [
            "ZONA",
            "TIPO DE RECURSO",
            "ITEM_SERVI"
        ]
    )
    .agg(
        CANTIDAD_TOTAL=("CANTIDAD","sum"),
        KM_TOTAL=("LONGITUD_KM","sum"),
        HORAS_TOTAL=("DURA_ACT","sum")
    )
    .reset_index()
)

resumen_recurso["HORAS_POR_ITEM"] = np.where(
    resumen_recurso["CANTIDAD_TOTAL"] > 0,
    resumen_recurso["HORAS_TOTAL"]
    /
    resumen_recurso["CANTIDAD_TOTAL"],
    np.nan
)


print("================================")
print("INICIANDO GENERACION DE CORREOS")
print("================================")


# ==========================================================
# EXCLUIR ROCERIA DEL ANALISIS EJECUTIVO
# ==========================================================

items_excluir = [
    "Intervencion forestal en cerco vivo y/o roceria"
]

df_correo = df_post_abril_2026[
    ~df_post_abril_2026["ITEM_SERVI"].isin(items_excluir)
].copy()

# ==========================================================
# RESUMEN POR ZONA
# ==========================================================

resumen_zona = (
    df_correo.groupby(
        ["ZONA","ITEM_SERVI"],
        dropna=False
    )
    .agg(
        CANTIDAD_TOTAL=("CANTIDAD","sum"),
        KM_TOTAL=("LONGITUD_KM","sum"),
        HORAS_TOTAL=("DURA_ACT","sum")
    )
    .reset_index()
)

resumen_zona["HORAS_POR_ITEM"] = np.where(
    resumen_zona["CANTIDAD_TOTAL"] > 0,
    resumen_zona["HORAS_TOTAL"] /
    resumen_zona["CANTIDAD_TOTAL"],
    np.nan
)

# ==========================================================
# RESUMEN POR RECURSO
# ==========================================================

resumen_recurso = (
    df_correo.groupby(
        [
            "ZONA",
            "TIPO DE RECURSO",
            "ITEM_SERVI"
        ]
    )
    .agg(
        CANTIDAD_TOTAL=("CANTIDAD","sum"),
        KM_TOTAL=("LONGITUD_KM","sum"),
        HORAS_TOTAL=("DURA_ACT","sum")
    )
    .reset_index()
)

resumen_recurso["HORAS_POR_ITEM"] = np.where(
    resumen_recurso["CANTIDAD_TOTAL"] > 0,
    resumen_recurso["HORAS_TOTAL"] /
    resumen_recurso["CANTIDAD_TOTAL"],
    np.nan
)

# ==========================================================
# EFICIENCIA POR ACTIVIDAD
# EXCLUYENDO ROCERIA
# ==========================================================

datos_eficiencia = resumen_recurso[
    ~resumen_recurso["ITEM_SERVI"]
    .str.contains(
        r"cerco vivo|roceria|rocería",
        case=False,
        na=False
    )
].copy()

FACTOR_KM = 100

datos_eficiencia["INDICE_EFICIENCIA"] = (
    (
        datos_eficiencia["CANTIDAD_TOTAL"]
        +
        (
            datos_eficiencia["KM_TOTAL"]
            * FACTOR_KM
        )
    )
    /
    datos_eficiencia["HORAS_TOTAL"]
)

tabla_eficiencia_item = (
    datos_eficiencia
    .pivot_table(
        index=[
            "ZONA",
            "ITEM_SERVI"
        ],
        columns="TIPO DE RECURSO",
        values="INDICE_EFICIENCIA",
        aggfunc="sum"
    )
    .reset_index()
)

tabla_eficiencia_item.columns.name = None


tabla_eficiencia_item["DIFERENCIA"] = (
    tabla_eficiencia_item["MOTO"]
    -
    tabla_eficiencia_item["CAMIONETA"]
)

tabla_eficiencia_item["VENTAJA_%"] = np.where(
    tabla_eficiencia_item["CAMIONETA"] > 0,
    (
        (
            tabla_eficiencia_item["MOTO"]
            /
            tabla_eficiencia_item["CAMIONETA"]
        ) - 1
    ) * 100,
    np.nan
)

tabla_eficiencia_item["RECURSO_MAS_EFICIENTE"] = np.where(
    tabla_eficiencia_item["MOTO"]
    >
    tabla_eficiencia_item["CAMIONETA"],
    "MOTO",
    "CAMIONETA"
)



# ==========================================================
# EXCLUIR ROCERIA DEL INDICE DE EFICIENCIA
# ==========================================================

datos_eficiencia = resumen_recurso[
    ~resumen_recurso["ITEM_SERVI"]
    .str.contains(
        r"cerco vivo|roceria|rocería",
        case=False,
        na=False
    )
].copy()

# ==========================================================
# INDICE DE EFICIENCIA OPERATIVA
# ==========================================================

eficiencia_grupos = (
    datos_eficiencia
    .groupby(
        [
            "ZONA",
            "TIPO DE RECURSO"
        ],
        dropna=False
    )
    .agg(
        CANTIDAD_TOTAL=("CANTIDAD_TOTAL", "sum"),
        KM_TOTAL=("KM_TOTAL", "sum"),
        HORAS_TOTAL=("HORAS_TOTAL", "sum")
    )
    .reset_index()
)

FACTOR_KM = 100

eficiencia_grupos["INDICE_EFICIENCIA"] = (
    (
        eficiencia_grupos["CANTIDAD_TOTAL"]
        +
        (
            eficiencia_grupos["KM_TOTAL"]
            * FACTOR_KM
        )
    )
    /
    eficiencia_grupos["HORAS_TOTAL"]
)

eficiencia_grupos["INDICE_EFICIENCIA"] = (
    eficiencia_grupos["INDICE_EFICIENCIA"]
    .round(2)
)

print("\nINDICE DE EFICIENCIA SIN ROCERIA")
print(
    eficiencia_grupos.sort_values(
        "INDICE_EFICIENCIA",
        ascending=False
    )
)



print("================================")
print("INICIANDO GENERACION DE CORREOS")
print("================================")

# ======================================================
# DESTINATARIOS
# ======================================================

correos_territorio = {

    "METROPOLITANA": {
        "to": "GONZALO.ALZATE@chec.com.co;Julian.Rios.Garcia@chec.com.co;CHRISTIAN.SANCHEZ@chec.com.co;CLAUDIA.JANETH.DIAZ@chec.com.co",
        "cc": "CESAR.AMAYA@contratista.chec.com.co;CESAR.AUGUSTO.HERNANDEZ@chec.com.co;CLAUDIA.BUITRAGO@chec.com.co;EXCENOVER.BARRERA@chec.com.co;JAVIER.AMBUILA@chec.com.co;JHON.CEBALLOS@chec.com.co;JOSE.DUQUE.GARCIA@chec.com.co;Juan.Alvarez.Barreto@chec.com.co;LUIS.EDUARDO.GARCIA@chec.com.co;Maria.Bernal.Herrera@chec.com.co;SERGIO.ANDRES.SANCHEZ@chec.com.co;JUAN.PALACIO@chec.com.co;HECTOR.CASTANEDA@chec.com.co;CRISTIAN.SERNA@chec.com.co;JOSE.ALBER.ALVAREZ@chec.com.co;LUIS.EDUARDO.MEJIA@chec.com.co;Luis.Giraldo.Cano@chec.com.co;CESAR.AMAYA@contratista.chec.com.co;JHON.WALTER.RIVERA@chec.com.co;DAVID.SANCHEZ@chec.com.co"
    },

    "NORTE": {
        "to": "PAOLA.CARDONA@chec.com.co;Julian.Manzur@chec.com.co;OSCAR.ARMANDO.ARIAS@chec.com.co",
        "cc": "JACOBO.GOMEZ@contratista.chec.com.co;CESAR.AUGUSTO.HERNANDEZ@chec.com.co;CLAUDIA.BUITRAGO@chec.com.co;EXCENOVER.BARRERA@chec.com.co;JAVIER.AMBUILA@chec.com.co;JHON.CEBALLOS@chec.com.co;JOSE.DUQUE.GARCIA@chec.com.co;Juan.Alvarez.Barreto@chec.com.co;LUIS.EDUARDO.GARCIA@chec.com.co;Maria.Bernal.Herrera@chec.com.co;JUAN.PALACIO@chec.com.co;SERGIO.ANDRES.SANCHEZ@chec.com.co;HECTOR.CASTANEDA@chec.com.co;CRISTIAN.SERNA@chec.com.co;JOSE.ALBER.ALVAREZ@chec.com.co;LUIS.EDUARDO.MEJIA@chec.com.co;Luis.Giraldo.Cano@chec.com.co;CESAR.AMAYA@contratista.chec.com.co"
    },

    "ORIENTE": {
        "to": "CARLOS.AGUILAR@chec.com.co;DAVID.SANCHEZ@chec.com.co",
        "cc": "CESAR.AMAYA@contratista.chec.com.co;ALEXANDRA.RODRIGUEZ@chec.com.co;CESAR.AUGUSTO.HERNANDEZ@chec.com.co;CHRISTIAN.SANCHEZ@chec.com.co;CLAUDIA.JANETH.DIAZ@chec.com.co;CLAUDIA.BUITRAGO@chec.com.co;EXCENOVER.BARRERA@chec.com.co;JAVIER.AMBUILA@chec.com.co;JHON.CEBALLOS@chec.com.co;JOSE.DUQUE.GARCIA@chec.com.co;Juan.Alvarez.Barreto@chec.com.co;LUIS.EDUARDO.GARCIA@chec.com.co;Maria.Bernal.Herrera@chec.com.co;SERGIO.ANDRES.SANCHEZ@chec.com.co;JUAN.PALACIO@chec.com.co"
    },

    "SUR": {
        "to": "JORGE.GUTIERREZ@chec.com.co;Javier.Pinilla@chec.com.co;PAOLA.CARDONA@chec.com.co",
        "cc": "JACOBO.GOMEZ@contratista.chec.com.co;ALEXANDRA.RODRIGUEZ@chec.com.co;CESAR.AUGUSTO.HERNANDEZ@chec.com.co;CHRISTIAN.SANCHEZ@chec.com.co;CLAUDIA.JANETH.DIAZ@chec.com.co;CLAUDIA.BUITRAGO@chec.com.co;EXCENOVER.BARRERA@chec.com.co;JAVIER.AMBUILA@chec.com.co;JHON.CEBALLOS@chec.com.co;JOSE.DUQUE.GARCIA@chec.com.co;Juan.Alvarez.Barreto@chec.com.co;LUIS.EDUARDO.GARCIA@chec.com.co;Maria.Bernal.Herrera@chec.com.co;SERGIO.ANDRES.SANCHEZ@chec.com.co;JUAN.PALACIO@chec.com.co;LUIS.JAVIER.HERRERA@chec.com.co;EDISON.SANCHEZ@chec.com.co;JULIAN.DAVID.ZULUAGA@chec.com.co;VICTOR.SEGURA@chec.com.co;EFREN.SUAREZ@chec.com.co"
    }

}

# ==========================================================
# TABLA CANTIDADES IGUAL A LA GRAFICA
# ==========================================================
tabla_global_cantidades = (
    cantidad_zona_recurso_item
    .pivot_table(
        index="ITEM_SERVI",
        columns=["ZONA", "TIPO DE RECURSO"],
        values="CANTIDAD_TOTAL",
        aggfunc="sum"
    )
)


# ==========================================================
# PRODUCTIVIDAD
# CANTIDADES POR HORA
# ==========================================================
productividad_recurso = (
    resumen_recurso.copy()
)

productividad_recurso["CANTIDADES_POR_HORA"] = np.where(
    productividad_recurso["HORAS_TOTAL"] > 0,
    productividad_recurso["CANTIDAD_TOTAL"]
    /
    productividad_recurso["HORAS_TOTAL"],
    np.nan
)

tabla_productividad = (
    productividad_recurso
    .pivot_table(
        index="ITEM_SERVI",
        columns=["ZONA", "TIPO DE RECURSO"],
        values="CANTIDADES_POR_HORA",
        aggfunc="sum"
    )
)

# ==========================================================
# PROMEDIO PRODUCTIVIDAD POR TIPO DE RECURSO
# ==========================================================
tabla_resumen_productividad = []

for actividad in tabla_productividad.index:

    fila = tabla_productividad.loc[actividad]

    valores_camioneta = []
    valores_moto = []

    for columna in tabla_productividad.columns:

        zona, recurso = columna

        valor = fila.get(columna, np.nan)

        if pd.isna(valor):
            continue

        if recurso == "CAMIONETA":
            valores_camioneta.append(valor)

        elif recurso == "MOTO":
            valores_moto.append(valor)

    prom_camioneta = (
        np.mean(valores_camioneta)
        if len(valores_camioneta) > 0
        else np.nan
    )

    prom_moto = (
        np.mean(valores_moto)
        if len(valores_moto) > 0
        else np.nan
    )

    prom_general = np.nanmean(
        [prom_camioneta, prom_moto]
    )

    tabla_resumen_productividad.append({
        "ACTIVIDAD": actividad,
        "PROM_CAMIONETAS": round(prom_camioneta, 2),
        "PROM_MOTOS": round(prom_moto, 2),
        "PROM_GENERAL": round(prom_general, 2)
    })

tabla_resumen_productividad = pd.DataFrame(
    tabla_resumen_productividad
)

tabla_resumen_productividad = (
    tabla_resumen_productividad
    .sort_values(
        "PROM_GENERAL",
        ascending=False
    )
)

# ==========================================================
# EFICIENCIA POR PODADOR
# CAMIONETA = 3 PODADORES
# MOTO = 2 PODADORES
# ==========================================================

tabla_eficiencia_podador = (
    tabla_resumen_productividad.copy()
)

tabla_eficiencia_podador["CAMIONETA_PODADOR"] = (
    tabla_eficiencia_podador["PROM_CAMIONETAS"] / 3
)

tabla_eficiencia_podador["MOTO_PODADOR"] = (
    tabla_eficiencia_podador["PROM_MOTOS"] / 2
)

tabla_eficiencia_podador["OBSERVACION"] = np.where(
    tabla_eficiencia_podador["CAMIONETA_PODADOR"]
    >
    tabla_eficiencia_podador["MOTO_PODADOR"],
    "CAMIONETA MÁS EFICIENTE",
    "MOTO MÁS EFICIENTE"
)

tabla_eficiencia_podador["VENTAJA_%"] = np.where(
    tabla_eficiencia_podador["CAMIONETA_PODADOR"]
    >
    tabla_eficiencia_podador["MOTO_PODADOR"],
    (
        tabla_eficiencia_podador["CAMIONETA_PODADOR"]
        /
        tabla_eficiencia_podador["MOTO_PODADOR"]
        - 1
    ) * 100,
    (
        tabla_eficiencia_podador["MOTO_PODADOR"]
        /
        tabla_eficiencia_podador["CAMIONETA_PODADOR"]
        - 1
    ) * 100
)

tabla_eficiencia_podador["VENTAJA_%"] = (
    tabla_eficiencia_podador["VENTAJA_%"]
    .round(1)
)



# ==========================================================
# TIEMPO PROMEDIO DE EJECUCION POR ACTIVIDAD
# ==========================================================
tabla_resumen_tiempos = []


# ==========================================================
# CONCLUSION AUTOMATICA TIEMPOS Y PRODUCTIVIDAD
# ==========================================================






for actividad in promedio_territorio_recurso["ITEM_SERVI"].unique():

    datos_act = promedio_territorio_recurso[
        promedio_territorio_recurso["ITEM_SERVI"] == actividad
    ].copy()

    prom_cam = (
        datos_act.loc[
            datos_act["TIPO DE RECURSO"] == "CAMIONETA",
            "HORAS_POR_ITEM_PROM"
        ]
        .mean()
    )

    prom_moto = (
        datos_act.loc[
            datos_act["TIPO DE RECURSO"] == "MOTO",
            "HORAS_POR_ITEM_PROM"
        ]
        .mean()
    )

    prom_general_horas = np.nanmean(
        [prom_cam, prom_moto]
    )

    prom_general_min = (
        prom_general_horas * 60
    )

    tabla_resumen_tiempos.append({
        "ACTIVIDAD": actividad,
        "PROM_CAM_MIN": round(prom_cam * 60, 1),
        "PROM_MOTO_MIN": round(prom_moto * 60, 1),
        "PROM_GENERAL_MIN": round(prom_general_min, 1),
        "PROM_GENERAL_HORAS": round(prom_general_horas, 2)
    })

tabla_resumen_tiempos = pd.DataFrame(
    tabla_resumen_tiempos
)

tabla_resumen_tiempos = (
    tabla_resumen_tiempos
    .sort_values(
        "PROM_GENERAL_HORAS",
        ascending=False
    )
)

actividad_mas_lenta = (
    tabla_resumen_tiempos
    .sort_values(
        "PROM_GENERAL_HORAS",
        ascending=False
    )
    .iloc[0]
)

actividad_mas_rapida = (
    tabla_resumen_tiempos
    .sort_values(
        "PROM_GENERAL_HORAS",
        ascending=True
    )
    .iloc[0]
)

actividad_mas_productiva = (
    tabla_resumen_productividad
    .sort_values(
        "PROM_GENERAL",
        ascending=False
    )
    .iloc[0]
)

actividad_menos_productiva = (
    tabla_resumen_productividad
    .sort_values(
        "PROM_GENERAL",
        ascending=True
    )
    .iloc[0]
)



html_comparativo_cantidades = """
<table border="1"
style="border-collapse:collapse;font-family:Arial;font-size:10pt;">
<tr style="background:#D9EAD3;">
<th>Actividad</th>
<th>Metrop. Cam</th>
<th>Metrop. Moto</th>
<th>Norte Cam</th>
<th>Norte Moto</th>
<th>Oriente Cam</th>
<th>Oriente Moto</th>
<th>Sur Cam</th>
<th>Sur Moto</th>
</tr>
"""

for actividad in tabla_global_cantidades.index:

    fila = tabla_global_cantidades.loc[actividad]

    valores = pd.to_numeric(
        fila,
        errors="coerce"
    )

    valor_max = valores.max()
    valor_min = valores.min()

    html_comparativo_cantidades += f"""
    <tr>
        <td>{actividad}</td>
    """

    for columna in tabla_global_cantidades.columns:

        valor = fila.get(columna, np.nan)

        if pd.isna(valor):

            html_comparativo_cantidades += """
            <td align="center">-</td>
            """

        else:

            if abs(valor - valor_max) < 0.0001:

                estilo = """
                background-color:#C6EFCE;
                color:#006100;
                font-weight:bold;
                """

            elif abs(valor - valor_min) < 0.0001:

                estilo = """
                background-color:#FFC7CE;
                color:#9C0006;
                font-weight:bold;
                """

            else:

                estilo = """
                background-color:white;
                color:black;
                """

            html_comparativo_cantidades += f"""
            <td align="center" style="{estilo}">
                {valor:,.0f}
            </td>
            """

    html_comparativo_cantidades += "</tr>"


html_comparativo_cantidades += "</table>"



# ==========================================================
# TABLA PRODUCTIVIDAD (CANTIDADES POR HORA)
# ==========================================================
html_productividad = """
<table border="1"
style="border-collapse:collapse;font-family:Arial;font-size:10pt;">
<tr style="background:#D9EAD3;">
<th>Actividad</th>
<th>Metrop. Cam</th>
<th>Metrop. Moto</th>
<th>Norte Cam</th>
<th>Norte Moto</th>
<th>Oriente Cam</th>
<th>Oriente Moto</th>
<th>Sur Cam</th>
<th>Sur Moto</th>
</tr>
"""

for actividad in tabla_productividad.index:

    fila = tabla_productividad.loc[actividad]

    valores = pd.to_numeric(
        fila,
        errors="coerce"
    )

    valor_max = valores.max()
    valor_min = valores.min()

    html_productividad += f"""
    <tr>
        <td>{actividad}</td>
    """

    for columna in tabla_productividad.columns:

        valor = fila.get(columna, np.nan)

        if pd.isna(valor):

            html_productividad += """
            <td align="center">-</td>
            """

        else:

            if abs(valor - valor_max) < 0.0001:

                estilo = """
                background-color:#C6EFCE;
                color:#006100;
                font-weight:bold;
                """

            elif abs(valor - valor_min) < 0.0001:

                estilo = """
                background-color:#FFC7CE;
                color:#9C0006;
                font-weight:bold;
                """

            else:

                estilo = """
                background-color:white;
                color:black;
                """

            html_productividad += f"""
            <td align="center" style="{estilo}">
                {valor:.2f}
            </td>
            """

    html_productividad += "</tr>"




html_productividad += """
</table>

"""
# ==========================================================
# CONCLUSION AUTOMATICA PRODUCTIVIDAD OPERATIVA
# EXCLUYENDO ROCERIA
# ==========================================================

datos_conclusion = cantidad_global.copy()

datos_conclusion = datos_conclusion[
    datos_conclusion["ITEM_SERVI"] !=
    "Intervencion forestal en cerco vivo y/o roceria"
].copy()

total_actividades = datos_conclusion["CANTIDAD_TOTAL"].sum()

datos_conclusion["PORCENTAJE"] = (
    datos_conclusion["CANTIDAD_TOTAL"]
    / total_actividades * 100
)

datos_conclusion = (
    datos_conclusion
    .sort_values(
        "CANTIDAD_TOTAL",
        ascending=False
    )
    .reset_index(drop=True)
)

top1 = datos_conclusion.iloc[0]
top2 = datos_conclusion.iloc[1]
top3 = datos_conclusion.iloc[2]

porcentaje_top3 = (
    top1["PORCENTAJE"]
    + top2["PORCENTAJE"]
    + top3["PORCENTAJE"]
)

# ==========================================================
# PRODUCTIVIDAD GLOBAL
# ==========================================================

datos_prod = resumen_recurso[
    ~resumen_recurso["ITEM_SERVI"]
    .str.contains(
        r"cerco vivo|roceria|rocería",
        case=False,
        na=False
    )
].copy()

cantidad_total_global = (
    datos_prod["CANTIDAD_TOTAL"].sum()
)

horas_total_global = (
    datos_prod["HORAS_TOTAL"].sum()
)

productividad_global = (
    cantidad_total_global
    / horas_total_global
    if horas_total_global > 0
    else 0
)

# ==========================================================
# DOS ACTIVIDADES DE MENOR PARTICIPACION
# ==========================================================

actividades_menor = (
    datos_conclusion
    .sort_values(
        "PORCENTAJE",
        ascending=True
    )
    .reset_index(drop=True)
)

menor1 = actividades_menor.iloc[0]
menor2 = actividades_menor.iloc[1]

# ==========================================================
# HTML CONCLUSION
# ==========================================================

html_conclusion_productividad = f"""
<p style="font-family:Arial;font-size:11pt;text-align:justify;">

<b>Conclusión cantidades:</b>

<br><br>

Excluyendo la actividad de intervención forestal en cerco vivo y/o rocería,
se ejecutaron <b>{total_actividades:,.0f}</b> actividades.

La operación se concentró principalmente en
<b>{top1['ITEM_SERVI']}</b> ({top1['PORCENTAJE']:.1f}%),
<b>{top2['ITEM_SERVI']}</b> ({top2['PORCENTAJE']:.1f}%)
y
<b>{top3['ITEM_SERVI']}</b> ({top3['PORCENTAJE']:.1f}%),

actividades que en conjunto representaron
<b>{porcentaje_top3:.1f}%</b>
de la ejecución total.

<br><br>

Considerando el volumen ejecutado de cada actividad,
la operación registró un rendimiento promedio de
<b>{productividad_global:.2f}</b>
actividades por hora trabajada.

En contraste,
<b>{menor1['ITEM_SERVI']}</b>
registró la menor participación
(<b>{menor1['PORCENTAJE']:.1f}%</b>),
seguida de
<b>{menor2['ITEM_SERVI']}</b>
(<b>{menor2['PORCENTAJE']:.1f}%</b>),
evidenciando una incidencia operativa significativamente menor dentro del período analizado.

</p>

<br><br>
"""

# ==========================================================
# TABLA RESUMEN PRODUCTIVIDAD
# ==========================================================
html_resumen_productividad = """
<table border="1"
style="border-collapse:collapse;font-family:Arial;font-size:10pt;">

<tr style="background:#D9EAD3;">
<th>Actividad</th>
<th>Prom. Camionetas</th>
<th>Prom. Motos</th>
<th>Prom. General</th>
</tr>
"""

for _, fila in tabla_resumen_productividad.iterrows():

    html_resumen_productividad += f"""
    <tr>
        <td>{fila['ACTIVIDAD']}</td>
        <td align="center">{fila['PROM_CAMIONETAS']:.2f}</td>
        <td align="center">{fila['PROM_MOTOS']:.2f}</td>
        <td align="center"><b>{fila['PROM_GENERAL']:.2f}</b></td>
    </tr>
    """

html_resumen_productividad += """
</table>

"""

# ==========================================================
# HTML EFICIENCIA POR PODADOR
# ==========================================================

html_eficiencia_podador = """
<table border="1"
style="border-collapse:collapse;font-family:Arial;font-size:10pt;">
<tr style="background:#D9EAD3;">
<th>Actividad</th>
<th>Cantidades/Hora por Podador (Camioneta)</th>
<th>Cantidades/Hora por Podador (Moto)</th>
<th>Ventaja (%)</th>
<th>Observación</th>
</tr>
"""

for _, fila in tabla_eficiencia_podador.iterrows():

    html_eficiencia_podador += f"""
    <tr>
        <td>{fila['ACTIVIDAD']}</td>

        <td align="center">
            {fila['CAMIONETA_PODADOR']:.2f}
        </td>

        <td align="center">
            {fila['MOTO_PODADOR']:.2f}
        </td>

        <td align="center">
            {fila['VENTAJA_%']:.1f}%
        </td>

        <td align="center">
            <b>{fila['OBSERVACION']}</b>
        </td>
    </tr>
    """

html_eficiencia_podador += """
</table>

<br>

<p style="font-family:Arial;font-size:10pt;text-align:justify;">

<b>Metodología del cálculo:</b>

<br><br>

La productividad individual se calculó a partir de la capacidad promedio de producción por hora de cada recurso, considerando la conformación de las cuadrillas operativas:

<ul>
<li><b>Camioneta = 3 podadores</b></li>
<li><b>Moto = 2 podadores</b></li>
</ul>

Por lo tanto:

<br>

<b>
Cantidades/Hora por Podador (Camioneta)
=
Cantidades/Hora de la Camioneta ÷ 3
</b>

<br><br>

<b>
Cantidades/Hora por Podador (Moto)
=
Cantidades/Hora de la Moto ÷ 2
</b>

<br><br>

La ventaja porcentual corresponde a la diferencia relativa de productividad entre ambos recursos para una misma actividad y se calcula mediante:

<br><br>

<b>
((Productividad Mayor − Productividad Menor)
÷ Productividad Menor)
× 100
</b>

<br><br>

La columna <b>Observación</b> identifica automáticamente el recurso que presenta la mayor productividad individual por podador para cada actividad.

</p>

<br><br>
"""



# ==========================================================
# HTML TIEMPO PROMEDIO DE EJECUCION
# ==========================================================
html_resumen_tiempos = """
<table border="1"
style="border-collapse:collapse;font-family:Arial;font-size:10pt;">

<tr style="background:#D9EAD3;">
<th>Actividad</th>
<th>Prom. Camionetas (min)</th>
<th>Prom. Motos (min)</th>
<th>Prom. General (min)</th>
<th>Prom. General (horas)</th>
</tr>
"""

for _, fila in tabla_resumen_tiempos.iterrows():

    html_resumen_tiempos += f"""
    <tr>
        <td>{fila['ACTIVIDAD']}</td>
        <td align="center">{fila['PROM_CAM_MIN']:.1f}</td>
        <td align="center">{fila['PROM_MOTO_MIN']:.1f}</td>
        <td align="center"><b>{fila['PROM_GENERAL_MIN']:.1f}</b></td>
        <td align="center">{fila['PROM_GENERAL_HORAS']:.2f}</td>
    </tr>
    """

html_resumen_tiempos += """
</table>
"""

# ==========================================================
# HTML RENDIMIENTO OPERATIVO POR GRUPO
# ==========================================================

html_rendimiento_grupos = """
<table border="1"
style="border-collapse:collapse;font-family:Arial;font-size:10pt;">
<tr style="background:#D9EAD3;">
<th>Grupo</th>
<th>Territorio</th>
<th>Tipo Recurso</th>
<th>Fenómeno del Niño</th>
<th>Días Trabajados</th>
<th>KM Ejecutados</th>
<th>KM por Día</th>
</tr>
"""

for _, fila in tabla_rendimiento_grupos.iterrows():

    if fila["FENOMENO_NINO"] == "SI":
        color_fenomeno = (
            "background-color:#FFF2CC;"
            "color:#7F6000;"
            "font-weight:bold;"
        )
    else:
        color_fenomeno = (
            "background-color:#E2F0D9;"
            "color:#006100;"
        )

    html_rendimiento_grupos += f"""
    <tr>
        <td><b>{fila[col_usuario]}</b></td>
        <td>{fila['ZONA']}</td>
        <td>{fila['TIPO DE RECURSO']}</td>
        <td style="{color_fenomeno}">
            {fila['FENOMENO_NINO']}
        </td>
        <td align="center">
            {fila['DIAS_TRABAJADOS']:,.0f}
        </td>
        <td align="center">
            {fila['LONGITUD_TOTAL_KM']:.2f}
        </td>
        <td align="center">
            <b>{fila['KM_POR_DIA']:.2f}</b>
        </td>
    </tr>
    """

texto_fenomeno_nino = ", ".join(
    sorted(grupos_fenomeno_nino)
)

html_rendimiento_grupos += f"""
</table>

<br>

<p style="font-family:Arial;font-size:10pt;">
<b>Nota:</b> Se identificaron
<b>{len(grupos_fenomeno_nino)}</b> grupos asociados al
<b>Fenómeno del Niño</b>:
({texto_fenomeno_nino}).


</p>
"""






# ==========================================================
# HTML TIEMPO PROMEDIO DE EJECUCION
# ==========================================================

html_resumen_tiempos = """
<table border="1" style="border-collapse:collapse;font-family:Arial;font-size:10pt;">

<tr style="background:#D9EAD3;">
<th>Actividad</th>
<th>Prom. Camionetas (min)</th>
<th>Prom. Motos (min)</th>
<th>Prom. General (min)</th>
<th>Prom. General (horas)</th>
</tr>
"""

for _, fila in tabla_resumen_tiempos.iterrows():

    html_resumen_tiempos += f"""
    <tr>
        <td>{fila['ACTIVIDAD']}</td>
        <td align="center">{fila['PROM_CAM_MIN']:.1f}</td>
        <td align="center">{fila['PROM_MOTO_MIN']:.1f}</td>
        <td align="center"><b>{fila['PROM_GENERAL_MIN']:.1f}</b></td>
        <td align="center">{fila['PROM_GENERAL_HORAS']:.2f}</td>
    </tr>
    """

html_resumen_tiempos += """
</table>
"""


# ==========================================================
# # TABLA KM PROMEDIO IGUAL A LA GRAFICA
# ==========================================================
tabla_global_km = (
    km_zona_recurso_item
    .pivot_table(
        index="ITEM_SERVI",
        columns=["ZONA", "TIPO DE RECURSO"],
        values="LONGITUD_TOTAL_KM",
        aggfunc="sum"
    )
)


# ==========================================================
# KM POR DIA REAL
# SIN DIAS DUPLICADOS
# ==========================================================

km_dia_recurso = (
    df_promedios.groupby(
        [
            "ZONA",
            "TIPO DE RECURSO",
            "ITEM_SERVI"
        ],
        dropna=False
    )
    .agg(
        LONGITUD_TOTAL_KM=(
            "LONGITUD_KM",
            "sum"
        ),
        DIAS_UNICOS=(
            col_fecha,
            lambda x: x.dt.date.nunique()
        )
    )
    .reset_index()
)

km_dia_recurso["KM_POR_DIA"] = np.where(
    km_dia_recurso["DIAS_UNICOS"] > 0,
    km_dia_recurso["LONGITUD_TOTAL_KM"]
    /
    km_dia_recurso["DIAS_UNICOS"],
    np.nan
)

km_dia_recurso["KM_POR_DIA"] = (
    km_dia_recurso["KM_POR_DIA"]
    .round(2)
)

tabla_km_dia = (
    km_dia_recurso
    .pivot_table(
        index="ITEM_SERVI",
        columns=["ZONA", "TIPO DE RECURSO"],
        values="KM_POR_DIA",
        aggfunc="mean"
    )
)

# ==========================================================
# RESUMEN KM POR DIA
# ==========================================================

tabla_resumen_km_dia = []

for actividad in tabla_km_dia.index:

    fila = tabla_km_dia.loc[actividad]

    valores_cam = []
    valores_moto = []

    for columna in tabla_km_dia.columns:

        zona, recurso = columna

        valor = fila.get(columna, np.nan)

        if pd.isna(valor):
            continue

        if recurso == "CAMIONETA":
            valores_cam.append(valor)

        elif recurso == "MOTO":
            valores_moto.append(valor)

    prom_cam = (
        np.mean(valores_cam)
        if len(valores_cam) > 0
        else np.nan
    )

    prom_moto = (
        np.mean(valores_moto)
        if len(valores_moto) > 0
        else np.nan
    )

    prom_general = np.nanmean(
        [prom_cam, prom_moto]
    )

    tabla_resumen_km_dia.append({
        "ACTIVIDAD": actividad,
        "PROM_CAMIONETAS": round(prom_cam, 2),
        "PROM_MOTOS": round(prom_moto, 2),
        "PROM_GENERAL": round(prom_general, 2)
    })

tabla_resumen_km_dia = pd.DataFrame(
    tabla_resumen_km_dia
)

tabla_resumen_km_dia = (
    tabla_resumen_km_dia
    .sort_values(
        "PROM_GENERAL",
        ascending=False
    )
)



# ==========================================================
# PRODUCTIVIDAD KM POR HORA
# ==========================================================
productividad_km_recurso = (
    resumen_recurso.copy()
)

productividad_km_recurso["KM_POR_HORA"] = np.where(
    productividad_km_recurso["HORAS_TOTAL"] > 0,
    productividad_km_recurso["KM_TOTAL"]
    /
    productividad_km_recurso["HORAS_TOTAL"],
    np.nan
)

tabla_productividad_km = (
    productividad_km_recurso
    .pivot_table(
        index="ITEM_SERVI",
        columns=["ZONA", "TIPO DE RECURSO"],
        values="KM_POR_HORA",
        aggfunc="sum"
    )
)


# ==========================================================
# CAPACIDAD OPERATIVA REAL KM POR DIA
# ==========================================================

capacidad_km_dia_zona = (
    resumen.groupby(
        ["ZONA", "TIPO DE RECURSO"],
        dropna=False
    )
    .agg(
        KM_TOTAL=("LONGITUD_TOTAL_KM", "sum"),
        DIAS_TRABAJADOS=("DIAS_TRABAJADOS", "sum")
    )
    .reset_index()
)

capacidad_km_dia_zona["KM_POR_DIA"] = np.where(
    capacidad_km_dia_zona["DIAS_TRABAJADOS"] > 0,
    capacidad_km_dia_zona["KM_TOTAL"]
    /
    capacidad_km_dia_zona["DIAS_TRABAJADOS"],
    np.nan
)

capacidad_km_dia_zona["KM_POR_DIA"] = (
    capacidad_km_dia_zona["KM_POR_DIA"]
    .round(2)
)

# ==========================================================
# TABLA CAPACIDAD KM DIA
# ==========================================================

tabla_capacidad_km_dia = (
    capacidad_km_dia_zona
    .pivot_table(
        index="TIPO DE RECURSO",
        columns="ZONA",
        values="KM_POR_DIA",
        aggfunc="first"
    )
)

# ==========================================================
# CAPACIDAD OPERATIVA REAL POR TERRITORIO
# PROMEDIO DE LOS GRUPOS
# ==========================================================

capacidad_grupo = (
    df_promedios.groupby(
        [
            "ZONA",
            col_usuario,
            "TIPO DE RECURSO"
        ],
        dropna=False
    )
    .agg(
        KM_TOTAL=("LONGITUD_KM", "sum"),
        DIAS_TRABAJADOS=(
            col_fecha,
            lambda x: x.dt.date.nunique()
        )
    )
    .reset_index()
)

capacidad_grupo["KM_POR_DIA"] = np.where(
    capacidad_grupo["DIAS_TRABAJADOS"] > 0,
    capacidad_grupo["KM_TOTAL"]
    /
    capacidad_grupo["DIAS_TRABAJADOS"],
    np.nan
)

capacidad_grupo["KM_POR_DIA"] = (
    capacidad_grupo["KM_POR_DIA"]
    .round(2)
)

print("\nVALIDACION CAPACIDAD POR GRUPO\n")
print(capacidad_grupo)

capacidad_km_dia_zona = (
    capacidad_grupo.groupby(
        ["ZONA", "TIPO DE RECURSO"],
        dropna=False
    )
    .agg(
        KM_POR_DIA=("KM_POR_DIA", "mean")
    )
    .reset_index()
)

capacidad_km_dia_zona["KM_POR_DIA"] = (
    capacidad_km_dia_zona["KM_POR_DIA"]
    .round(2)
)

tabla_capacidad_km_dia = (
    capacidad_km_dia_zona
    .pivot_table(
        index="TIPO DE RECURSO",
        columns="ZONA",
        values="KM_POR_DIA",
        aggfunc="first"
    )
)

print("\nVALIDACION CAPACIDAD TERRITORIAL REAL\n")
print(tabla_capacidad_km_dia)


# ==========================================================
# HTML CAPACIDAD REAL KM POR DIA
# ==========================================================

html_capacidad_km_dia = """
<table border="1"
style="border-collapse:collapse;font-family:Arial;font-size:10pt;">
<tr style="background:#D9EAD3;">
<th>Recurso</th>
<th>Metropolitana</th>
<th>Norte</th>
<th>Oriente</th>
<th>Sur</th>
</tr>
"""

for recurso in tabla_capacidad_km_dia.index:

    fila = tabla_capacidad_km_dia.loc[recurso]

    html_capacidad_km_dia += f"""
    <tr>
        <td><b>{recurso}</b></td>
        <td align="center">{fila.get('METROPOLITANA',0):.2f}</td>
        <td align="center">{fila.get('NORTE',0):.2f}</td>
        <td align="center">{fila.get('ORIENTE',0):.2f}</td>
        <td align="center">{fila.get('SUR',0):.2f}</td>
    </tr>
    """

html_capacidad_km_dia += """
</table>
"""


# ==========================================================
# CAPACIDAD OPERATIVA GLOBAL CORREGIDA
# ==========================================================

dias_periodo = (
    df_promedios[col_fecha]
    .dt.date
    .nunique()
)

capacidad_global_km_dia = (
    df_promedios.groupby(
        "TIPO DE RECURSO",
        dropna=False
    )
    .agg(
        KM_TOTAL=("LONGITUD_KM", "sum")
    )
    .reset_index()
)

capacidad_global_km_dia["KM_POR_DIA"] = (
    capacidad_global_km_dia["KM_TOTAL"]
    /
    dias_periodo
)

capacidad_global_km_dia["KM_POR_DIA"] = (
    capacidad_global_km_dia["KM_POR_DIA"]
    .round(2)
)

print("\nVALIDACION CAPACIDAD GLOBAL\n")
print(capacidad_global_km_dia)
print("\nVALIDACION CAPACIDAD GLOBAL\n")

print(capacidad_global_km_dia)

print("\nTOTAL DIAS UNICOS DEL PERIODO\n")

print(
    df_promedios[col_fecha]
    .dt.date
    .nunique()
)

print("\nKM TOTALES POR RECURSO\n")

print(
    df_promedios.groupby(
        "TIPO DE RECURSO"
    )["LONGITUD_KM"]
    .sum()
)

print(capacidad_global_km_dia)

html_capacidad_global_km_dia = """
<table border="1"
style="border-collapse:collapse;font-family:Arial;font-size:10pt;">
<tr style="background:#D9EAD3;">
<th>Recurso</th>
<th>KM por Día Promedio</th>
</tr>
"""

# ==========================================================
# CAPACIDAD OPERATIVA GLOBAL
# DESDE TABLA TERRITORIAL YA CORREGIDA
# ==========================================================

capacidad_global_km_dia = pd.DataFrame({
    "TIPO DE RECURSO": tabla_capacidad_km_dia.index,
    "KM_POR_DIA": tabla_capacidad_km_dia.sum(axis=1).values
})

print("\nVALIDACION NUEVA CAPACIDAD GLOBAL\n")
print(capacidad_global_km_dia)

# ==========================================================
# CAPACIDAD OPERATIVA GLOBAL PROMEDIO
# ==========================================================

capacidad_global_km_dia = pd.DataFrame({
    "TIPO DE RECURSO": tabla_capacidad_km_dia.index,
    "KM_POR_DIA": tabla_capacidad_km_dia.mean(axis=1, skipna=True).values
})

capacidad_global_km_dia["KM_POR_DIA"] = (
    capacidad_global_km_dia["KM_POR_DIA"]
    .round(2)
)

print("\nVALIDACION CAPACIDAD GLOBAL PROMEDIO\n")
print(capacidad_global_km_dia)

capacidad_global_km_dia["KM_POR_DIA"] = (
    capacidad_global_km_dia["KM_POR_DIA"]
    .round(2)
)

print("\nVALIDACION CAPACIDAD GLOBAL PROMEDIO\n")
print(capacidad_global_km_dia)

html_capacidad_global_km_dia = """
<table border="1"
style="border-collapse:collapse;font-family:Arial;font-size:10pt;">
<tr style="background:#D9EAD3;">
<th>Recurso</th>
<th>KM por Día Promedio</th>
</tr>
"""

for _, fila in capacidad_global_km_dia.iterrows():

    html_capacidad_global_km_dia += f"""
    <tr>
        <td><b>{fila['TIPO DE RECURSO']}</b></td>
        <td align="center">{fila['KM_POR_DIA']:.2f}</td>
    </tr>
    """

html_capacidad_global_km_dia += """
</table>
"""


# ==========================================================
# TABLA KM ACUMULADOS
# ==========================================================
html_comparativo_km = """
<table border="1"
style="border-collapse:collapse;font-family:Arial;font-size:10pt;">
<tr style="background:#D9EAD3;">
<th>Actividad</th>
<th>Metrop. Cam</th>
<th>Metrop. Moto</th>
<th>Norte Cam</th>
<th>Norte Moto</th>
<th>Oriente Cam</th>
<th>Oriente Moto</th>
<th>Sur Cam</th>
<th>u Moto</th>
</tr>
"""


for actividad in tabla_global_km.index:

    fila = tabla_global_km.loc[actividad]

    valores = pd.to_numeric(
        fila,
        errors="coerce"
    )

    valor_max = valores.max()
    valor_min = valores.min()

    html_comparativo_km += f"""
    <tr>
        <td>{actividad}</td>
    """

    for columna in tabla_global_km.columns:

        valor = fila.get(columna, np.nan)

        if pd.isna(valor):

            html_comparativo_km += """
            <td align="center">-</td>
            """

        else:

            if abs(valor - valor_max) < 0.0001:

                estilo = """
                background-color:#FFC7CE;
                color:#9C0006;
                font-weight:bold;
                """

            elif abs(valor - valor_min) < 0.0001:

                estilo = """
                background-color:#C6EFCE;
                color:#006100;
                font-weight:bold;
                """

            else:

                estilo = """
                background-color:white;
                color:black;
                """

            html_comparativo_km += f"""
            <td align="center" style="{estilo}">
                {valor:.2f}
            </td>
            """

    html_comparativo_km += "</tr>"

html_comparativo_km += "</table>"

# ==========================================================
# CONCLUSION AUTOMATICA KM INTERVENIDOS
# ==========================================================

km_totales_actividad = (
    km_zona_recurso_item
    .groupby("ITEM_SERVI", dropna=False)
    .agg(
        KM_TOTAL=("LONGITUD_TOTAL_KM", "sum")
    )
    .reset_index()
    .sort_values(
        "KM_TOTAL",
        ascending=False
    )
)

actividad_km_1 = km_totales_actividad.iloc[0]["ITEM_SERVI"]
actividad_km_2 = km_totales_actividad.iloc[1]["ITEM_SERVI"]
actividad_km_3 = km_totales_actividad.iloc[2]["ITEM_SERVI"]

html_conclusion_km = f"""
<p style="font-family:Arial;font-size:11pt;text-align:justify;">

<b>Conclusión:</b>

<br><br>

La cobertura de intervención se concentró principalmente en
<b>{actividad_km_1}</b>,
<b>{actividad_km_2}</b> y
<b>{actividad_km_3}</b>,
actividades que representaron la mayor longitud intervenida durante el período analizado.

Asimismo, los mayores niveles de cobertura se registraron en los recursos y territorios con mayor acumulado de kilómetros intervenidos, evidenciando una mayor participación operativa en estas actividades.

En contraste, las actividades con menor longitud intervenida presentaron una participación reducida dentro de la cobertura total ejecutada, reflejando una menor demanda operativa frente a las actividades de mayor intervención.

</p>

<br><br>
"""


# ==========================================================
# TABLA KM PROMEDIO POR DIA
# ==========================================================

html_km_dia = """
<table border="1"
style="border-collapse:collapse;font-family:Arial;font-size:10pt;">
<tr style="background:#D9EAD3;">
<th>Actividad</th>
<th>Metrop. Cam</th>
<th>Metrop. Moto</th>
<th>Norte Cam</th>
<th>Norte Moto</th>
<th>Oriente Cam</th>
<th>Oriente Moto</th>
<th>Sur Cam</th>
<th>Sur Moto</th>
</tr>
"""

for actividad in tabla_km_dia.index:

    fila = tabla_km_dia.loc[actividad]

    valores = pd.to_numeric(
        fila,
        errors="coerce"
    )

    valor_max = valores.max()
    valor_min = valores.min()

    html_km_dia += f"""
    <tr>
        <td>{actividad}</td>
    """

    for columna in tabla_km_dia.columns:

        valor = fila.get(columna, np.nan)

        if pd.isna(valor):

            html_km_dia += """
            <td align="center">-</td>
            """

        else:

            if abs(valor - valor_max) < 0.0001:

                estilo = """
                background-color:#FFC7CE;
                color:#9C0006;
                font-weight:bold;
                """

            elif abs(valor - valor_min) < 0.0001:

                estilo = """
                background-color:#C6EFCE;
                color:#006100;
                font-weight:bold;
                """

            else:

                estilo = """
                background-color:white;
                color:black;
                """

            html_km_dia += f"""
            <td align="center" style="{estilo}">
                {valor:.2f}
            </td>
            """

    html_km_dia += "</tr>"

html_km_dia += "</table>"

# ==========================================================
# HTML RESUMEN KM POR DIA
# ==========================================================

html_resumen_km_dia = """
<table border="1"
style="border-collapse:collapse;font-family:Arial;font-size:10pt;">
<tr style="background:#D9EAD3;">
<th>Actividad</th>
<th>Prom. Camionetas</th>
<th>Prom. Motos</th>
<th>Prom. General</th>
</tr>
"""

for _, fila in tabla_resumen_km_dia.iterrows():

    html_resumen_km_dia += f"""
    <tr>
        <td>{fila['ACTIVIDAD']}</td>
        <td align="center">{fila['PROM_CAMIONETAS']:.2f}</td>
        <td align="center">{fila['PROM_MOTOS']:.2f}</td>
        <td align="center"><b>{fila['PROM_GENERAL']:.2f}</b></td>
    </tr>
    """

html_resumen_km_dia += """
</table>
"""

# ==========================================================
# PRODUCTIVIDAD KM POR HORA
# ==========================================================
productividad_km_recurso = (
    resumen_recurso.copy()
)

productividad_km_recurso["KM_POR_HORA"] = np.where(
    productividad_km_recurso["HORAS_TOTAL"] > 0,
    productividad_km_recurso["KM_TOTAL"]
    /
    productividad_km_recurso["HORAS_TOTAL"],
    np.nan
)

tabla_productividad_km = (
    productividad_km_recurso
    .pivot_table(
        index="ITEM_SERVI",
        columns=["ZONA", "TIPO DE RECURSO"],
        values="KM_POR_HORA",
        aggfunc="sum"
    )
)

# ==========================================================
# TABLA PRODUCTIVIDAD KM POR HORA
# ==========================================================
html_productividad_km = """
<table border="1"
style="border-collapse:collapse;font-family:Arial;font-size:10pt;">
<tr style="background:#D9EAD3;">
<th>Actividad</th>
<th>Metrop. Cam</th>
<th>Metrop. Moto</th>
<th>Norte Cam</th>
<th>Norte Moto</th>
<th>Oriente Cam</th>
<th>Oriente Moto</th>
<th>Sur Cam</th>
<th>Sur Moto</th>
</tr>
"""


for actividad in tabla_productividad_km.index:

    fila = tabla_productividad_km.loc[actividad]

    valores = pd.to_numeric(
        fila,
        errors="coerce"
    )

    valor_max = valores.max()
    valor_min = valores.min()

    html_productividad_km += f"""
    <tr>
        <td>{actividad}</td>
    """

    for columna in tabla_productividad_km.columns:

        valor = fila.get(columna, np.nan)

        if pd.isna(valor):

            html_productividad_km += """
            <td align="center">-</td>
            """

        else:

            if abs(valor - valor_max) < 0.0001:

                estilo = """
                background-color:#C6EFCE;
                color:#006100;
                font-weight:bold;
                """

            elif abs(valor - valor_min) < 0.0001:

                estilo = """
                background-color:#FFC7CE;
                color:#9C0006;
                font-weight:bold;
                """

            else:

                estilo = """
                background-color:white;
                color:black;
                """

            html_productividad_km += f"""
            <td align="center" style="{estilo}">
                {valor:.2f}
            </td>
            """

    html_productividad_km += "</tr>"

html_productividad_km += """
</table>
<br><br>
"""


# ==========================================================
# CONCLUSION AUTOMATICA KM
# ==========================================================

texto_conclusion_km = ""

for territorio in ["METROPOLITANA", "NORTE", "ORIENTE", "SUR"]:

    datos = (
        km_zona_recurso_item[
            km_zona_recurso_item["ZONA"] == territorio
        ]
        .copy()
    )

    if datos.empty:
        continue

    actividad_principal = (
        datos
        .sort_values(
            "LONGITUD_TOTAL_KM",
            ascending=False
        )
        .iloc[0]
    )

    texto_conclusion_km += f"""
    <p style="font-family:Arial;font-size:11pt;">
    <b>{territorio}</b><br>
    La actividad con mayor kilometraje promedio corresponde a
    <b>{actividad_principal['ITEM_SERVI']}</b>,
    con un promedio de
    <b>{actividad_principal['LONGITUD_TOTAL_KM']:.2f} km</b>
    ejecutados por registro.
    </p>
    """


# ==========================================================
# TABLA HORAS IGUAL A LA GRAFICA
# ==========================================================

tabla_global_horas = (
    promedio_territorio_recurso
    .pivot_table(
        index="ITEM_SERVI",
        columns=["ZONA", "TIPO DE RECURSO"],
        values="HORAS_POR_ITEM_PROM",
        aggfunc="first"
    )
)

html_comparativo_territorios = """


<table border="1"
style="border-collapse:collapse;font-family:Arial;font-size:10pt;">

<tr style="background:#D9EAD3;">
<th>Actividad</th>
<th>Metrop. Cam</th>
<th>Metrop. Moto</th>
<th>Norte Cam</th>
<th>Norte Moto</th>
<th>Oriente Cam</th>
<th>Oriente Moto</th>
<th>Sur Cam</th>
<th>Sur Moto</th>
</tr>
"""

for actividad in tabla_global_horas.index:

    fila = tabla_global_horas.loc[actividad]

    valores = pd.to_numeric(
        fila,
        errors="coerce"
    )

    valor_max = valores.max()
    valor_min = valores.min()

    html_comparativo_territorios += f"""
    <tr>
        <td>{actividad}</td>
    """

    for columna in tabla_global_horas.columns:

        valor = fila.get(columna, np.nan)

        if pd.isna(valor):

            html_comparativo_territorios += """
            <td align="center">-</td>
            """

        else:

            if abs(valor - valor_max) < 0.0001:

                estilo = """
                background-color:#FFC7CE;
                color:#9C0006;
                font-weight:bold;
                """

            elif abs(valor - valor_min) < 0.0001:

                estilo = """
                background-color:#C6EFCE;
                color:#006100;
                font-weight:bold;
                """

            else:

                estilo = """
                background-color:white;
                color:black;
                """

            html_comparativo_territorios += f"""
            <td align="center" style="{estilo}">
                {valor:.2f}
            </td>
            """

    html_comparativo_territorios += "</tr>"

html_comparativo_territorios += "</table>"

# ==========================================================
# CONCLUSIONES AUTOMATICAS POR ACTIVIDAD
# ==========================================================
html_comparativo_territorios += """
<br>

<p style="
font-family:Arial;
font-size:10pt;
">

<b>Interpretación de colores:</b>

<br><br>

<span style="
background-color:#FFC7CE;
color:#9C0006;
padding:3px 8px;
font-weight:bold;
">
Valores Altos
</span>

&nbsp;&nbsp;

Representan los mayores tiempos promedio de ejecución para la actividad evaluada.

<br><br>

<span style="
background-color:#C6EFCE;
color:#006100;
padding:3px 8px;
font-weight:bold;
">
Valores Bajos
</span>

&nbsp;&nbsp;

Representan los menores tiempos promedio de ejecución para la actividad evaluada.

<br><br>

Los valores sin resaltado corresponden a desempeños intermedios entre los recursos y territorios comparados.

</p>
"""

"""

# for actividad in tabla_global_horas.index:

#     datos_cam = []

#     for zona in ["METROPOLITANA", "NORTE", "ORIENTE", "SUR"]:

#         valor = tabla_global_horas.loc[actividad].get(
#             (zona, "CAMIONETA"),
#             np.nan
#         )

#         if pd.notna(valor):
#             datos_cam.append((zona, valor))

#     if len(datos_cam) > 1:

#         mejor_cam = min(datos_cam, key=lambda x: x[1])
#         peor_cam = max(datos_cam, key=lambda x: x[1])

#         texto_conclusiones_horas += f"""
#         <p style="font-family:Arial;font-size:11pt;">

#         <b>{actividad}</b><br><br>

#         🚚 <b>Camionetas:</b><br><br>

#         El menor tiempo promedio se registró en
#         <b>{mejor_cam[0]}</b>
#         con
#         <b>{mejor_cam[1]:.2f}</b>
#         s por ítem.

#         <br><br>

#         Por su parte,
#         <b>{peor_cam[0]}</b>
#         presentó el mayor tiempo promedio,
#         con
#         <b>{peor_cam[1]:.2f}</b>
#         s por ítem.

#         </p>
#         """

#     datos_moto = []

#     for zona in ["METROPOLITANA", "NORTE", "ORIENTE", "SUR"]:

#         valor = tabla_global_horas.loc[actividad].get(
#             (zona, "MOTO"),
#             np.nan
#         )

#         if pd.notna(valor):
#             datos_moto.append((zona, valor))

#     if len(datos_moto) > 1:

#         mejor_moto = min(datos_moto, key=lambda x: x[1])
#         peor_moto = max(datos_moto, key=lambda x: x[1])

#         texto_conclusiones_horas += f"""
#         <p style="font-family:Arial;font-size:11pt;">

#         🏍️ <b>Motos:</b><br><br>

#         El menor tiempo promedio se registró en
#         <b>{mejor_moto[0]}</b>
#         con
#         <b>{mejor_moto[1]:.2f}</b>
#         horas por ítem.

#         <br><br>

#         Mientras que
#         <b>{peor_moto[0]}</b>
#         presentó el mayor tiempo promedio,
#         con
#         <b>{peor_moto[1]:.2f}</b>
#         horas por ítem.

#         </p>

#         <hr>
#         """


for territorio in ["METROPOLITANA", "NORTE", "ORIENTE", "SUR"]:

    datos_zona_resumen = resumen_post_abril_2026[
        resumen_post_abril_2026["ZONA"] == territorio
    ].copy()

    if len(datos_zona_resumen) == 0:
        continue

    fila_mayor = (
        datos_zona_resumen
        .sort_values(
            "LONGITUD_TOTAL_KM",
            ascending=False
        )
        .iloc[0]
    )

    fila_menor = (
        datos_zona_resumen[
            datos_zona_resumen["LONGITUD_TOTAL_KM"] > 0
        ]
        .sort_values(
            "LONGITUD_TOTAL_KM",
            ascending=True
        )
        .iloc[0]
    )

    eficiencia_territorio = (
        tabla_eficiencia_item[
            tabla_eficiencia_item["ZONA"] == territorio
        ]
        .copy()
    )


    gana_moto = (
        eficiencia_territorio[
            eficiencia_territorio[
                "RECURSO_MAS_EFICIENTE"
            ] == "MOTO"
        ]
    )

    gana_camioneta = (
        eficiencia_territorio[
            eficiencia_territorio[
                "RECURSO_MAS_EFICIENTE"
            ] == "CAMIONETA"
        ]
    )



    detalle_actividades = ""

    for _, fila in eficiencia_territorio.iterrows():

        if (
            pd.isna(fila["CAMIONETA"])
            or
            pd.isna(fila["MOTO"])
        ):
            continue

        diferencia = abs(
            fila["VENTAJA_%"]
        )

        if fila["RECURSO_MAS_EFICIENTE"] == "MOTO":

            formula_porcentaje = (
                f"(({fila['MOTO']:.2f} / "
                f"{fila['CAMIONETA']:.2f}) - 1) × 100"
            )

            texto_interpretacion = (
                f"La MOTO presentó una eficiencia "
                f"{diferencia:.1f}% superior "
                f"respecto a la CAMIONETA."
            )

        else:

            formula_porcentaje = (
                f"(({fila['CAMIONETA']:.2f} / "
                f"{fila['MOTO']:.2f}) - 1) × 100"
            )

            texto_interpretacion = (
                f"La CAMIONETA presentó una eficiencia "
                f"{diferencia:.1f}% superior "
                f"respecto a la MOTO."
        )        








        detalle_actividades += f"""
        <br>

        <b>{fila['ITEM_SERVI']}</b>

        <br>

        Camioneta:
        <b>{fila['CAMIONETA']:.2f}</b>

        <br>

        Moto:
        <b>{fila['MOTO']:.2f}</b>

        <br>

                
        Recurso más eficiente:
        <b>{fila['RECURSO_MAS_EFICIENTE']}</b>

        <br>

        Ventaja relativa:
        <b>{diferencia:.1f}%</b>

        <br>

        Fórmula aplicada:

        <br>

        <b>{formula_porcentaje}</b>

        <br>

        Interpretación:

        <br>

        {texto_interpretacion}





        <br><br>
        """


    



    # ======================================================
    # TABLA EFICIENCIA HTML
    # ======================================================

    tabla_eficiencia_html = """
    <table border="1" cellpadding="5" cellspacing="0"
    style="border-collapse:collapse;font-family:Arial;font-size:11pt;">
    <tr style="background-color:#D9EAD3;">
    <th>Actividad</th>
    <th>Camioneta</th>
    <th>Moto</th>
    <th>Más Eficiente</th>
    <th>Diferencia</th>
    </tr>
    """

    for _, fila in eficiencia_territorio.iterrows():

        if pd.isna(fila["CAMIONETA"]) or pd.isna(fila["MOTO"]):
            continue

        tabla_eficiencia_html += f"""
        <tr>
        <td>{fila['ITEM_SERVI']}</td>
        <td align="center">{fila['CAMIONETA']:.2f}</td>
        <td align="center">{fila['MOTO']:.2f}</td>
        <td align="center"><b>{fila['RECURSO_MAS_EFICIENTE']}</b></td>
        <td align="center">{abs(fila['VENTAJA_%']):.1f}%</td>
        </tr>
        """



  
    print("="*80)
    print("TERRITORIO:", territorio)
    print("COLUMNAS:")
    print(datos_zona_resumen.columns.tolist())
    print("SHAPE:")
    print(datos_zona_resumen.shape)
    print("="*80)

    datos_zona_resumen_resumen = resumen_zona[
    resumen_zona["ZONA"] == territorio
    ].copy()

    km_total = datos_zona_resumen_resumen["KM_TOTAL"].sum()

    print(f"Generando correo para {territorio}")


    actividad_mayor_zona = fila_mayor["ITEM_SERVI"]
    actividad_menor_zona = fila_menor["ITEM_SERVI"]

    km_mayor_zona = fila_mayor["LONGITUD_TOTAL_KM"]
    km_menor_zona = fila_menor["LONGITUD_TOTAL_KM"]


    print("="*50)
    print("TERRITORIO ACTUAL:", territorio)
    print("="*50)
       
    print(type(km_total))
    print(type(km_mayor_zona))
    print(type(km_menor_zona))  


    print("km_total")
    print(km_total)
    print(type(km_total))

    print("km_mayor_zona")
    print(km_mayor_zona)
    print(type(km_mayor_zona))

    print("km_menor_zona")
    print(km_menor_zona)
    print(type(km_menor_zona))

    print("actividad_mayor_zona")
    print(type(actividad_mayor_zona))

    print("actividad_menor_zona")
    print(type(actividad_menor_zona))
    print("------------------------")
    print("territorio:", territorio)

    print("km_total =", km_total)
    print(type(km_total))

    print("km_mayor_zona =", km_mayor_zona)
    print(type(km_mayor_zona))

    print("km_menor_zona =", km_menor_zona)
    print(type(km_menor_zona))

    print("actividad_mayor_zona =", actividad_mayor_zona)
    print(type(actividad_mayor_zona))

    print("actividad_menor_zona =", actividad_menor_zona)
    print(type(actividad_menor_zona))
    print("------------------------")



    actividad_top_cantidad = (
        datos_zona_resumen_resumen
        .sort_values(
            "CANTIDAD_TOTAL",
            ascending=False
        )
        .iloc[0]
    )

    top_km = (
        datos_zona_resumen_resumen
        .sort_values(
            "KM_TOTAL",
            ascending=False
        )
        .head(3)
    )

    mejores_tiempos = (
        datos_zona_resumen_resumen[
            datos_zona_resumen_resumen["CANTIDAD_TOTAL"] >= 20
        ]
        .sort_values(
            "HORAS_POR_ITEM",
            ascending=True
        )
        .head(3)
    )



    texto_km = ""

    for _, fila in top_km.iterrows():

        texto_km += (
            f"• {fila['ITEM_SERVI']}: "
            f"{fila['KM_TOTAL']:.1f} km<br>"
        )

    texto_tiempos = ""

    for _, fila in mejores_tiempos.iterrows():

        texto_tiempos += (
            f"• {fila['ITEM_SERVI']}: "
            f"{fila['HORAS_POR_ITEM']:.2f} horas por ítem<br>"
        )

    comportamiento_operativo = f"""
    <p style="font-family:Arial;font-size:11pt;">

    <b>Comportamiento operativo</b><br><br>

    La combinación de altas cantidades ejecutadas y coberturas relevantes
    evidencia una importante concentración de recursos en actividades de
    manejo de vegetación dentro del territorio <b>{territorio}</b>.

    <br><br>

    La actividad con mayor volumen de ejecución fue:

    <b>{actividad_top_cantidad['ITEM_SERVI']}</b>,

    con un total de
    <b>{actividad_top_cantidad['CANTIDAD_TOTAL']:,.0f}</b>
    unidades ejecutadas.

    <br><br>

    Las actividades con mayor cobertura intervenida fueron:

    <br><br>

    {texto_km}

    <br>

    Desde la perspectiva de tiempos de ejecución,
    <b>{territorio}</b> presenta algunos de los mejores desempeños operativos:

    <br><br>

    {texto_tiempos}

    </p>
    """


    if len(datos_zona_resumen_resumen) == 0:
        continue

    cantidad_total_zona = (
        datos_zona_resumen_resumen["CANTIDAD_TOTAL"].sum()
    )

    km_total_zona = (
        datos_zona_resumen_resumen["KM_TOTAL"].sum()
    )

    top_cantidad = (
        datos_zona_resumen_resumen
        .sort_values(
            "CANTIDAD_TOTAL",
            ascending=False
        )
        .head(3)
        .copy()
    )

    top_km = (
        datos_zona_resumen_resumen
        .sort_values(
            "KM_TOTAL",
            ascending=False
        )
        .head(3)
        .copy()
    )

    top_horas = (
        datos_zona_resumen_resumen[
            datos_zona_resumen_resumen["CANTIDAD_TOTAL"] >= 20
        ]
        .sort_values(
            "HORAS_POR_ITEM",
            ascending=True
        )
        .head(3)
        .copy()
    )

    top_cantidad["PORC"] = (
        top_cantidad["CANTIDAD_TOTAL"]
        /
        cantidad_total_zona
        * 100
    )

    top_km["PORC"] = (
        top_km["KM_TOTAL"]
        /
        km_total_zona
        * 100
    )


    # ==========================================================
    # CONCLUSION EJECUTIVA AUTOMATICA
    # ==========================================================

    actividad_cantidad_1 = top_cantidad.iloc[0]
    actividad_cantidad_2 = top_cantidad.iloc[1]

    actividad_km_1 = top_km.iloc[0]
    actividad_km_2 = top_km.iloc[1]

    actividad_tiempo_1 = top_horas.iloc[0]
    actividad_tiempo_2 = top_horas.iloc[1]
    actividad_tiempo_3 = top_horas.iloc[2]

    texto_conclusion_territorio = f"""
    <p style="font-family:Arial;font-size:11pt;">

    La operación del territorio <b>{territorio}</b> se concentra principalmente en la actividad
    <b>{actividad_cantidad_1['ITEM_SERVI']}</b>, que representa aproximadamente
    <b>{actividad_cantidad_1['PORC']:.1f}%</b> de las cantidades ejecutadas.
    Asimismo, <b>{actividad_cantidad_2['ITEM_SERVI']}</b> mantiene una participación relevante dentro de la gestión operativa de la zona.

    <br><br>

    En términos de cobertura, <b>{actividad_km_1['ITEM_SERVI']}</b> aporta cerca de
    <b>{actividad_km_1['PORC']:.1f}%</b> de los kilómetros intervenidos, mientras que
    <b>{actividad_km_2['ITEM_SERVI']}</b> contribuye con aproximadamente
    <b>{actividad_km_2['PORC']:.1f}%</b>, concentrando una parte importante del esfuerzo operativo del territorio.

    <br><br>

    Desde la perspectiva de productividad, los menores tiempos promedio de ejecución corresponden a:

    <br><br>

    • {actividad_tiempo_1['ITEM_SERVI']}: {actividad_tiempo_1['HORAS_POR_ITEM']:.2f} horas por ítem<br>
    • {actividad_tiempo_2['ITEM_SERVI']}: {actividad_tiempo_2['HORAS_POR_ITEM']:.2f} horas por ítem<br>
    • {actividad_tiempo_3['ITEM_SERVI']}: {actividad_tiempo_3['HORAS_POR_ITEM']:.2f} horas por ítem<br>

    <br>

    En conjunto, los resultados evidencian un adecuado equilibrio entre cantidades ejecutadas, cobertura intervenida y eficiencia operativa durante el período analizado.

    </p>
    """





    camioneta = resumen_recurso[
        (resumen_recurso["ZONA"] == territorio)
        &
        (resumen_recurso["TIPO DE RECURSO"] == "CAMIONETA")
    ]

    moto = resumen_recurso[
        (resumen_recurso["ZONA"] == territorio)
        &
        (resumen_recurso["TIPO DE RECURSO"] == "MOTO")
    ]

    cant_camioneta = (
        camioneta["CANTIDAD_TOTAL"].sum()
    )

    cant_moto = (
        moto["CANTIDAD_TOTAL"].sum()
    )

    km_camioneta = (
        camioneta["KM_TOTAL"].sum()
    )

    km_moto = (
        moto["KM_TOTAL"].sum()
    )

    participacion_camioneta = (
        cant_camioneta / cantidad_total_zona * 100
        if cantidad_total_zona > 0 else 0
    )

    participacion_moto = (
        cant_moto / cantidad_total_zona * 100
        if cantidad_total_zona > 0 else 0
    )


    participacion_moto = (
    cant_moto
    /
    cantidad_total_zona
    * 100
    )
   
    # ==========================================================
    # RESUMEN DEL TERRITORIO
    # ==========================================================

    html_resumen_territorio = ""



    # ==========================================================
    # ANALISIS EJECUTIVO POR TIPO DE RECURSO
    # ==========================================================

    datos_recurso = resumen_recurso[
        resumen_recurso["ZONA"] == territorio
    ].copy()



    # ==========================================================
    # PRODUCTIVIDAD OPERATIVA POR GRUPO
    # SIN ROCERIA
    # ==========================================================

    datos_grupo_productividad = resumen[
        (resumen["ZONA"] == territorio)
        &
        (
            ~resumen["ITEM_SERVI"]
            .astype(str)
            .str.contains(
                r"cerco vivo|roceria|rocería",
                case=False,
                na=False
            )
        )
    ].copy()

    tabla_productividad_grupo = (
        datos_grupo_productividad.groupby(
            [
                col_usuario,
                "TIPO DE RECURSO"
            ],
            dropna=False
        )
        .agg(
            DIAS_TRABAJADOS=("DIAS_TRABAJADOS", "max"),
            CANTIDADES_EJECUTADAS=("CANTIDAD_TOTAL", "sum"),
            HORAS_TOTALES=("DURACION_TOTAL", "sum"),
            KM_EJECUTADOS=("LONGITUD_TOTAL_KM", "sum")
        )
        .reset_index()
    )

    total_cantidades = (
        tabla_productividad_grupo["CANTIDADES_EJECUTADAS"]
        .sum()
    )

    tabla_productividad_grupo["PARTICIPACION"] = np.where(
        total_cantidades > 0,
        tabla_productividad_grupo["CANTIDADES_EJECUTADAS"]
        / total_cantidades * 100,
        0
    )

    tabla_productividad_grupo["CANTIDADES_DIA"] = np.where(
        tabla_productividad_grupo["DIAS_TRABAJADOS"] > 0,
        tabla_productividad_grupo["CANTIDADES_EJECUTADAS"]
        /
        tabla_productividad_grupo["DIAS_TRABAJADOS"],
        np.nan
    )

    tabla_productividad_grupo["CANTIDADES_HORA"] = np.where(
        tabla_productividad_grupo["HORAS_TOTALES"] > 0,
        tabla_productividad_grupo["CANTIDADES_EJECUTADAS"]
        /
        tabla_productividad_grupo["HORAS_TOTALES"],
        np.nan
    )

    km_grupo = (
        tabla_rendimiento_grupos[
            tabla_rendimiento_grupos["ZONA"] == territorio
        ][[
            col_usuario,
            "KM_POR_DIA"
        ]]
    )

    tabla_productividad_grupo = (
        tabla_productividad_grupo
        .merge(
            km_grupo,
            on=col_usuario,
            how="left"
        )
    )

    tabla_productividad_grupo["PARTICIPACION"] = (
        tabla_productividad_grupo["PARTICIPACION"]
        .round(1)
    )

    tabla_productividad_grupo["CANTIDADES_DIA"] = (
        tabla_productividad_grupo["CANTIDADES_DIA"]
        .round(2)
    )

    tabla_productividad_grupo["CANTIDADES_HORA"] = (
        tabla_productividad_grupo["CANTIDADES_HORA"]
        .round(2)
    )

    tabla_productividad_grupo = (
        tabla_productividad_grupo
        .sort_values(
            "CANTIDADES_EJECUTADAS",
            ascending=False
        )
    )


    # ==========================================================
    # ACTIVIDAD PRINCIPAL DEL GRUPO
    # ==========================================================

    actividad_principal = (
        datos_grupo_productividad
        .groupby(
            [col_usuario, "ITEM_SERVI"],
            dropna=False
        )
        .agg(
            CANTIDADES_EJECUTADAS=("CANTIDAD_TOTAL", "sum")
        )
        .reset_index()
    )

    actividad_principal = (
        actividad_principal
        .sort_values(
            [col_usuario, "CANTIDADES_EJECUTADAS"],
            ascending=[True, False]
        )
        .drop_duplicates(
            subset=[col_usuario],
            keep="first"
        )
    )

    actividad_principal = actividad_principal[
        [col_usuario, "ITEM_SERVI"]
    ].rename(
        columns={
            "ITEM_SERVI": "ACTIVIDAD_PRINCIPAL"
        }
    )

    tabla_productividad_grupo = (
        tabla_productividad_grupo
        .merge(
            actividad_principal,
            on=col_usuario,
        how="left"
    )
)



    # ==========================================================
    # HTML PRODUCTIVIDAD POR GRUPO
    # ==========================================================

    html_productividad_grupo = """
    <table border="1"
    style="border-collapse:collapse;font-family:Arial;font-size:10pt;">
    <tr style="background:#D9EAD3;">
    <th>Grupo</th>
    <th>Tipo Recurso</th>
    <th>Días</th>
    <th>Cantidades</th>
    <th>% Participación</th>
    <th>Cant/Día</th>
    <th>Cant/Hora</th>
    <th>KM/Día</th>
    <th>Actividad Principal</th>
    </tr>
    """

    for _, fila in tabla_productividad_grupo.iterrows():

        html_productividad_grupo += f"""
        <tr>
            <td><b>{fila[col_usuario]}</b></td>
            <td>{fila['TIPO DE RECURSO']}</td>
            <td align="center">{fila['DIAS_TRABAJADOS']:.0f}</td>
            <td align="center">{fila['CANTIDADES_EJECUTADAS']:,.0f}</td>
            <td align="center">{fila['PARTICIPACION']:.1f}%</td>
            <td align="center">{fila['CANTIDADES_DIA']:.2f}</td>
            <td align="center">{fila['CANTIDADES_HORA']:.2f}</td>
            <td align="center">{fila['KM_POR_DIA']:.2f}</td>
            <td>{fila['ACTIVIDAD_PRINCIPAL']}</td>
        </tr>
        """

    html_productividad_grupo += """
    </table>
    """


    # ----------------------------------------------------------
    # CANTIDADES
    # ----------------------------------------------------------

    tabla_cantidad_recurso = (
        datos_recurso.pivot_table(
            index="ITEM_SERVI",
            columns="TIPO DE RECURSO",
            values="CANTIDAD_TOTAL",
            aggfunc="sum",
            fill_value=0
        )
        .reset_index()
    )

    tabla_cantidad_recurso["TOTAL"] = (
        tabla_cantidad_recurso.get("CAMIONETA",0)
        +
        tabla_cantidad_recurso.get("MOTO",0)
    )

    tabla_cantidad_recurso["PARTICIPACION_CAMIONETA"] = (
        tabla_cantidad_recurso.get("CAMIONETA",0)
        /
        tabla_cantidad_recurso["TOTAL"]
        * 100
    ).round(1)

    tabla_cantidad_recurso["PARTICIPACION_MOTO"] = (
        tabla_cantidad_recurso.get("MOTO",0)
        /
        tabla_cantidad_recurso["TOTAL"]
        * 100
    ).round(1)

    tabla_cantidad_recurso["RECURSO_LIDER"] = np.where(
        tabla_cantidad_recurso.get("CAMIONETA",0)
        >
        tabla_cantidad_recurso.get("MOTO",0),
        "CAMIONETA",
        "MOTO"
    )

    tabla_cantidad_recurso = (
        tabla_cantidad_recurso
        .sort_values("TOTAL", ascending=False)
        .head(5)
    )



    tabla_km_recurso = (
        datos_recurso.pivot_table(
            index="ITEM_SERVI",
            columns="TIPO DE RECURSO",
            values="KM_TOTAL",
            aggfunc="sum",
            fill_value=0
        )
        .reset_index()
    )

    tabla_km_recurso["TOTAL"] = (
        tabla_km_recurso.get("CAMIONETA",0)
        +
        tabla_km_recurso.get("MOTO",0)
    )

    tabla_km_recurso["RECURSO_LIDER"] = np.where(
        tabla_km_recurso.get("CAMIONETA",0)
        >
        tabla_km_recurso.get("MOTO",0),
        "CAMIONETA",
        "MOTO"
    )

    tabla_km_recurso = (
        tabla_km_recurso
        .sort_values("TOTAL", ascending=False)
        .head(5)
    )




    tabla_horas_recurso = (
        datos_recurso.pivot_table(
            index="ITEM_SERVI",
            columns="TIPO DE RECURSO",
            values="HORAS_POR_ITEM",
            aggfunc="mean"
        )
        .reset_index()
    )

    tabla_horas_recurso = (
        tabla_horas_recurso
        .sort_values(
            "CAMIONETA",
            ascending=True,
            na_position="last"
        )
    )


    # ==========================================================
    # HTML TABLA CANTIDADES
    # ==========================================================

    html_tabla_cantidades = """
    <table border="1"
    style="border-collapse:collapse;font-family:Arial;font-size:11pt;">
    <tr style="background-color:#D9EAD3;">
    <th>Actividad</th>
    <th>Camionetas</th>
    <th>Motos</th>
    <th>Participación Principal</th>
    </tr>
    """

    for _, fila in tabla_cantidad_recurso.iterrows():

        html_tabla_cantidades += f"""
        <tr>
        <td>{fila['ITEM_SERVI']}</td>
        <td align="center">{fila.get('CAMIONETA',0):,.0f}</td>
        <td align="center">{fila.get('MOTO',0):,.0f}</td>
        <td align="center"><b>{fila['RECURSO_LIDER']}</b></td>
        </tr>
        """

    html_tabla_cantidades += "</table>"


    # ==========================================================
    # HTML TABLA KM
    # ==========================================================

    html_tabla_km = """
    <table border="1"
    style="border-collapse:collapse;font-family:Arial;font-size:11pt;">
    <tr style="background-color:#D9EAD3;">
    <th>Actividad</th>
    <th>Camionetas (KM)</th>
    <th>Motos (KM)</th>
    <th>Participación Principal</th>
    </tr>
    """

    for _, fila in tabla_km_recurso.iterrows():

        html_tabla_km += f"""
        <tr>
        <td>{fila['ITEM_SERVI']}</td>
        <td align="center">{fila.get('CAMIONETA',0):.1f}</td>
        <td align="center">{fila.get('MOTO',0):.1f}</td>
        <td align="center"><b>{fila['RECURSO_LIDER']}</b></td>
        </tr>
        """

    html_tabla_km += "</table>"


    # ==========================================================
    # HTML TABLA HORAS
    # ==========================================================

    html_tabla_horas = """
    <table border="1"
    style="border-collapse:collapse;font-family:Arial;font-size:11pt;">
    <tr style="background-color:#D9EAD3;">
    <th>Actividad</th>
    <th>Camionetas (Horas/Item)</th>
    <th>Motos (Horas/Item)</th>
    </tr>
    """


    for _, fila in tabla_horas_recurso.iterrows():

        html_tabla_horas += f"""
        <tr>
        <td>{fila['ITEM_SERVI']}</td>
        <td align="center">{fila.get('CAMIONETA',0):.2f}</td>
        <td align="center">{fila.get('MOTO',0):.2f}</td>
        </tr>
        """

    html_tabla_horas += "</table>"


    # ==========================================================
    # HALLAZGOS AUTOMATICOS
    # ==========================================================

    actividad_camioneta = (
        tabla_cantidad_recurso
        .sort_values(
            "PARTICIPACION_CAMIONETA",
            ascending=False
        )
        .iloc[0]
    )

    actividad_moto = (
        tabla_cantidad_recurso
        .sort_values(
            "PARTICIPACION_MOTO",
            ascending=False
        )
        .iloc[0]
    )

    hallazgos_recursos = f"""
    <p style="font-family:Arial;font-size:11pt;">

    <b>Hallazgos Principales</b>

    <br><br>

    ✅ <b>Camionetas</b>

    <br><br>

    La mayor concentración de actividades corresponde a
    <b>{actividad_camioneta['ITEM_SERVI']}</b>,
    donde las camionetas representan aproximadamente
    <b>{actividad_camioneta['PARTICIPACION_CAMIONETA']:.1f}%</b>
    de la ejecución total de esta actividad.

    <br><br>

    ✅ <b>Motos</b>

    <br><br>

    La actividad con mayor participación de motos corresponde a
    <b>{actividad_moto['ITEM_SERVI']}</b>,
    con una participación aproximada de
    <b>{actividad_moto['PARTICIPACION_MOTO']:.1f}%</b>.

    <br><br>

    Los resultados se recalculan automáticamente para cada territorio
    y permiten identificar los focos operativos de cada tipo de recurso.

    </p>
    """




    # ==========================================================
    # GRAFICA INDICE DE EFICIENCIA
    # ==========================================================

    eficiencia_zona = eficiencia_grupos[
        eficiencia_grupos["ZONA"] == territorio
    ].copy()

    # ----------------------------------------------------------
    # PEGAR AQUI EL BLOQUE
    # INTERPRETACION DEL INDICE
    # ----------------------------------------------------------

    if len(eficiencia_zona) >= 2:

        eficiencia_ordenada = (
            eficiencia_zona
            .sort_values(
                "INDICE_EFICIENCIA",
                ascending=False
            )
            .reset_index(drop=True)
        )

        mejor_recurso = eficiencia_ordenada.iloc[0]
        segundo_recurso = eficiencia_ordenada.iloc[1]

        diferencia_abs = (
            mejor_recurso["INDICE_EFICIENCIA"]
            -
            segundo_recurso["INDICE_EFICIENCIA"]
        )

        diferencia_pct = (
            (
                mejor_recurso["INDICE_EFICIENCIA"]
                /
                segundo_recurso["INDICE_EFICIENCIA"]
            ) - 1
        ) * 100

        texto_indice_eficiencia = f"""
        <h3 style="color:#0066CC;">
        Interpretación del Índice de Eficiencia Operativa
        </h3>

        <p style="font-family:Arial;font-size:11pt;">

        <b>Metodología del cálculo</b>
        <br><br>

        <b>Nota metodológica:</b>

        La actividad Intervención Forestal en Cerco Vivo y/o Rocería fue excluida de los comparativos de cantidades, productividad y tiempos de ejecución debido a que su unidad principal de control corresponde a kilómetros lineales intervenidos y no a unidades ejecutadas. Por esta razón se analiza de forma independiente para evitar distorsiones en los indicadores de desempeño.

        <br><br>

        El índice integra tres variables operativas fundamentales:

        <br><br>

        • Cantidades ejecutadas.<br>
        • Kilómetros intervenidos.<br>
        • Horas consumidas.<br>

        <br>

        La fórmula utilizada es:

        <br><br>

        <b>
        (Cantidad Total + (Kilómetros Totales × 100))
        ÷ Horas Totales
        </b>

        <br><br>

        El resultado permite comparar qué recurso genera una mayor producción
        operativa respecto al esfuerzo invertido.

        <br><br>

        <b>Resultado obtenido para {territorio}</b>

        <br><br>

        El recurso con mayor eficiencia fue
        <b>{mejor_recurso['TIPO DE RECURSO']}</b>
        con un índice de
        <b>{mejor_recurso['INDICE_EFICIENCIA']:.2f}</b>.

        <br><br>

        El segundo recurso evaluado fue
        <b>{segundo_recurso['TIPO DE RECURSO']}</b>
        con un índice de
        <b>{segundo_recurso['INDICE_EFICIENCIA']:.2f}</b>.

        <br><br>

        La diferencia observada fue de
        <b>{diferencia_abs:.2f}</b> puntos,
        equivalente aproximadamente a
        <b>{diferencia_pct:.1f}%</b>.

        <br><br>

        <b>Conclusión ejecutiva</b>

        <br><br>

        Los resultados indican que el recurso
        <b>{mejor_recurso['TIPO DE RECURSO']}</b>
        logró una mejor relación entre producción,
        cobertura intervenida y horas utilizadas,
        evidenciando un mayor aprovechamiento de la capacidad operativa disponible.

        <br><br>

        Este hallazgo no implica necesariamente una mejor calidad técnica,
        ya que el resultado puede estar influenciado por:

        <br><br>

        • Tipo de actividad ejecutada.<br>
        • Complejidad de los trabajos.<br>
        • Condiciones de acceso.<br>
        • Distancias de desplazamiento.<br>
        • Distribución geográfica de las órdenes.<br>

        <br>

        Por esta razón se recomienda complementar el análisis con los
        indicadores de cantidades, kilómetros y horas presentados anteriormente.

        </p>
        """

    else:

        texto_indice_eficiencia = ""

    # ==========================================================
    # CONTINUA AQUI EL CODIGO DE LA GRAFICA
    # ==========================================================

    if not eficiencia_zona.empty:

        eficiencia_zona = (
            eficiencia_zona
            .sort_values(
                "INDICE_EFICIENCIA",
                ascending=False
            )
        )

        plt.figure(
            figsize=(ANCHO_GRAFICA, 10)
        )


    print("VOY A CREAR CORREO:", territorio)
    # ==========================================================
    # CREAR CORREO OUTLOOK
    # ==========================================================

    mail = crear_mail_territorio(territorio, correos_territorio)
    if mail is None:
        continue

    firma = mail.HTMLBody







    # ==========================================================
    # GRAFICA PRINCIPAL
    # ==========================================================

    ruta_grafica = (
        r"D:\OneDrive - Grupo EPM\PYTHON PODA ARCHIVOS DE INFORMES"
        r"\Datos_Globales\01_Territorios\Comparativo_Territorios_Por_Item.png"
    )
    print("CANTIDAD =", ruta_grafica)
    adjunto = mail.Attachments.Add(ruta_grafica)

    adjunto.PropertyAccessor.SetProperty(
        "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
        "grafica_cantidad"
    )

    # ==========================================================
    # GRAFICA HORAS PROMEDIO
    # ==========================================================

    ruta_grafica_horas = (
        r"D:\OneDrive - Grupo EPM\PYTHON PODA ARCHIVOS DE INFORMES"
        r"\Datos_Globales\02_Horas_Territorio\Comparativo_Zonas_Duracion.png"
    )
    print("HORAS =", ruta_grafica_horas)

    adjunto_horas = mail.Attachments.Add(
        ruta_grafica_horas
    )

    adjunto_horas.PropertyAccessor.SetProperty(
        "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
        "grafica_horas"
    )



    ruta_km = (
        r"D:\OneDrive - Grupo EPM\PYTHON PODA ARCHIVOS DE INFORMES"
        r"\Datos_Globales\02_Horas_Territorio\Kilometros_Item_Territorio.png"
    )

    adjunto_km = mail.Attachments.Add(
        ruta_km
    )

    adjunto_km.PropertyAccessor.SetProperty(
        "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
        "grafica_km"
    )

    ruta_km_mes = os.path.join(
        carpeta_km,
        "KM_Mes_a_Mes.png"
    )

    adjunto_km_mes = mail.Attachments.Add(
        ruta_km_mes
    )

    adjunto_km_mes.PropertyAccessor.SetProperty(
        "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
        "grafica_km_mes"
    )

    ruta_grupo = os.path.join(
        carpeta_territorios,
        f"{territorio}_Comparativo_Grupos.png"
    )

    adjunto_grupo = mail.Attachments.Add(
        ruta_grupo
    )

    adjunto_grupo.PropertyAccessor.SetProperty(
        "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
        "grafica_grupo"
    )

    ruta_km_mes_zona = os.path.join(
        carpeta_km,
        f"{territorio}_KM_MES.png"
    )

    adjunto_km_mes_zona = mail.Attachments.Add(
        ruta_km_mes_zona
    )

    adjunto_km_mes_zona.PropertyAccessor.SetProperty(
        "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
        "grafica_km_mes_zona"
    )


    ruta_km_grupo = os.path.join(
        carpeta_km,
        f"{territorio}_KM_Comparativo_Grupos.png"
    )

    adjunto_km_grupo = mail.Attachments.Add(
        ruta_km_grupo
    )

    adjunto_km_grupo.PropertyAccessor.SetProperty(
        "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
        "grafica_km_grupo"
    )



    ruta_horas_grupo = os.path.join(
        carpeta_horas,
        f"{territorio}_Horas_Promedio.png"
    )

    adjunto_horas_grupo = mail.Attachments.Add(
        ruta_horas_grupo
    )

    adjunto_horas_grupo.PropertyAccessor.SetProperty(
        "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
        "grafica_horas_grupo"
    )




    print("================================")
    print("CONCLUSION TERRITORIO")
    print("================================")
    print(texto_conclusion_territorio)

    print(texto_conclusion_territorio)





    # ==========================================================
    # TABLA RENDIMIENTO SOLO DEL TERRITORIO DEL CORREO
    # ==========================================================

    datos_rendimiento_territorio = (
        tabla_rendimiento_grupos[
            tabla_rendimiento_grupos["ZONA"] == territorio
        ]
        .copy()
    )






    html_rendimiento_territorio = """
    <table border="1"
    style="border-collapse:collapse;font-family:Arial;font-size:10pt;">
    <tr style="background:#D9EAD3;">
    <th>Grupo</th>
    <th>Tipo Recurso</th>
    <th>Fenómeno del Niño</th>
    <th>Días Trabajados</th>
    <th>KM Ejecutados</th>
    <th>KM por Día</th>
    </tr>
    """
    ranking_km = (
        datos_rendimiento_territorio
        .sort_values("KM_POR_DIA", ascending=False)
        .reset_index(drop=True)
    )

    lider = ranking_km.iloc[0]
    segundo = ranking_km.iloc[1]
    ultimo = ranking_km.iloc[-1]

    html_conclusion_rendimiento = f"""
    <p style="font-family:Arial;font-size:11pt;text-align:justify;">

    <b>Conclusión Ejecutiva:</b>

    <br><br>

    La mayor capacidad de cobertura correspondió al grupo
    <b>{lider[col_usuario]}</b>,
    con un rendimiento promedio de
    <b>{lider['KM_POR_DIA']:.2f} km por día</b>,
    seguido por
    <b>{segundo[col_usuario]}</b>
    con
    <b>{segundo['KM_POR_DIA']:.2f} km por día</b>.

    <br><br>

    Por su parte,
    <b>{ultimo[col_usuario]}</b>
    registró el menor rendimiento operativo,
    con una cobertura promedio de
    <b>{ultimo['KM_POR_DIA']:.2f} km por día</b>.

    <br><br>

    La comparación se realiza sobre la capacidad real de cobertura de cada grupo mediante el indicador
    <b>KM por Día</b>, calculado a partir de los kilómetros ejecutados y los días efectivamente trabajados.

    <br><br>

    Es importante considerar que las camionetas operan con
    <b>3 podadores que ejecutan simultáneamente la misma actividad</b>,
    mientras que cada motocicleta dispone de
    <b>1 trabajador</b>, aunque normalmente los grupos de motocicleta desarrollan sus actividades de forma coordinada en parejas.

    <br><br>

    Por esta razón, el indicador de cobertura diaria constituye la referencia principal para evaluar el desempeño operativo de los grupos dentro del territorio.

    </p>
    <br><br>
    """



    for _, fila in datos_rendimiento_territorio.iterrows():

        color_fenomeno = (
            "background-color:#FFF2CC;color:#7F6000;font-weight:bold;"
            if fila["FENOMENO_NINO"] == "SI"
            else
            "background-color:#E2F0D9;color:#006100;"
        )

        html_rendimiento_territorio += f"""
        <tr>
            <td><b>{fila[col_usuario]}</b></td>
            <td>{fila['TIPO DE RECURSO']}</td>
            <td style="{color_fenomeno}">
                {fila['FENOMENO_NINO']}
            </td>
            <td align="center">
                {fila['DIAS_TRABAJADOS']:,.0f}
            </td>
            <td align="center">
                {fila['LONGITUD_TOTAL_KM']:.2f}
            </td>
            <td align="center">
                <b>{fila['KM_POR_DIA']:.2f}</b>
            </td>
        </tr>
        """

    html_rendimiento_territorio += "</table>"


    # ==========================================================
    # CONCLUSION EJECUTIVA DEL TERRITORIO
    # ==========================================================

    # RECURSO CON MAYOR COBERTURA KM/DIA
    ranking_km = (
        datos_rendimiento_territorio
        .sort_values("KM_POR_DIA", ascending=False)
        .reset_index(drop=True)
    )

    lider = ranking_km.iloc[0]
    segundo = ranking_km.iloc[1]

    # RECURSO QUE GENERA MAS CANTIDADES
    resumen_recurso_territorio = (
        tabla_productividad_grupo
        .groupby("TIPO DE RECURSO", dropna=False)
        .agg(
            CANTIDADES=("CANTIDADES_EJECUTADAS", "sum"),
            KM_DIA=("KM_POR_DIA", "mean")
        )
        .reset_index()
    )

    recurso_cantidades = (
        resumen_recurso_territorio
        .sort_values("CANTIDADES", ascending=False)
        .iloc[0]["TIPO DE RECURSO"]
    )

    recurso_km = (
        resumen_recurso_territorio
        .sort_values("KM_DIA", ascending=False)
        .iloc[0]["TIPO DE RECURSO"]
    )

    actividad_principal = (
        tabla_productividad_grupo
        .groupby("ACTIVIDAD_PRINCIPAL")
        ["CANTIDADES_EJECUTADAS"]
        .sum()
        .sort_values(ascending=False)
        .index[0]
    )





    # ==========================================================
    # RECURSO RECOMENDADO PARA CANTIDADES Y KM
    # ==========================================================

    resumen_recursos_territorio = (
        tabla_productividad_grupo
        .groupby("TIPO DE RECURSO", dropna=False)
        .agg(
            CANTIDADES=("CANTIDADES_EJECUTADAS", "sum"),
            KM_DIA=("KM_POR_DIA", "mean")
        )
        .reset_index()
    )

    recurso_cantidades = (
        resumen_recursos_territorio
        .sort_values("CANTIDADES", ascending=False)
        .iloc[0]["TIPO DE RECURSO"]
    )

    recurso_km = (
        resumen_recursos_territorio
        .sort_values("KM_DIA", ascending=False)
        .iloc[0]["TIPO DE RECURSO"]
    )

    actividad_principal_territorio = (
        tabla_productividad_grupo
        .groupby("ACTIVIDAD_PRINCIPAL")
        ["CANTIDADES_EJECUTADAS"]
        .sum()
        .sort_values(ascending=False)
        .index[0]
    )










    html_conclusion_rendimiento = f"""
    <p style="font-family:Arial;font-size:11pt;text-align:justify;">

    <b>Conclusión Ejecutiva:</b>

    <br><br>

    La mayor capacidad de cobertura correspondió al grupo
    <b>{lider[col_usuario]}</b>,
    con un rendimiento promedio de
    <b>{lider['KM_POR_DIA']:.2f} km por día</b>,
    seguido por
    <b>{segundo[col_usuario]}</b>
    con
    <b>{segundo['KM_POR_DIA']:.2f} km por día</b>.

    <br><br>

    La actividad predominante del territorio fue
    <b>{actividad_principal_territorio}</b>,
    concentrando la mayor participación dentro de la operación ejecutada.

    <br><br>

    Los resultados evidencian que para actividades medidas por cantidades ejecutadas (podas, talas y guaduas), el recurso con mejor desempeño operativo fue
    <b>{recurso_cantidades}</b>,
    al concentrar el mayor volumen de actividades desarrolladas durante el período analizado.

    <br><br>

    Por otra parte, cuando el objetivo principal es maximizar la cobertura territorial y los kilómetros intervenidos por jornada, el recurso más favorable fue
    <b>{recurso_km}</b>,
    al registrar los mayores indicadores de KM por día.

    <br><br>

    Es importante considerar que las camionetas operan con
    <b>3 podadores ejecutando simultáneamente la misma actividad</b>,
    mientras que cada motocicleta dispone de
    <b>1 trabajador</b>, aunque normalmente los grupos de motocicleta desarrollan sus actividades de forma coordinada en parejas.

    <br><br>

    En consecuencia, para actividades orientadas a maximizar cantidades ejecutadas se recomienda priorizar el recurso
    <b>{recurso_cantidades}</b>,
    mientras que para actividades cuyo objetivo principal sea aumentar la cobertura territorial en kilómetros se recomienda priorizar el recurso
    <b>{recurso_km}</b>.

    </p>
    <br><br>
    """



    # ==========================================================
    # RESUMEN RECURSOS SOLO DEL TERRITORIO
    # ==========================================================
    datos_resumen_territorio = (
        tabla_resumen_recursos[
            tabla_resumen_recursos["ZONA"] == territorio
        ]
        .copy()
    )

    html_resumen_recursos_territorio = """
    <table border="1"
    style="border-collapse:collapse;font-family:Arial;font-size:10pt;">
    <tr style="background:#D9EAD3;">
    <th>Categoría</th>
    <th>Grupos</th>
    <th>KM Ejecutados</th>
    <th>Días Trabajados</th>
    <th>KM por Día Promedio</th>
    </tr>
    """

    for _, fila in datos_resumen_territorio.iterrows():

        html_resumen_recursos_territorio += f"""
        <tr>
            <td><b>{fila['CATEGORIA']}</b></td>
            <td align="center">{fila['GRUPOS']}</td>
            <td align="center">{fila['KM_EJECUTADOS']:.2f}</td>
            <td align="center">{fila['DIAS_TRABAJADOS']:.0f}</td>
            <td align="center"><b>{fila['KM_POR_DIA']:.2f}</b></td>
        </tr>
        """

    html_resumen_recursos_territorio += """
    </table>
    """



    nota_metodologica = """
    <p style="font-family:Arial;font-size:10pt;text-align:justify;">
    <b>Nota:</b> Los resultados corresponden al desempeño del recurso operativo completo (cuadrilla). La conformación de cada recurso puede variar de acuerdo con la programación operativa vigente y las necesidades del servicio. Por lo tanto, los indicadores reflejan la capacidad operativa de cobertura y ejecución del recurso asignado, y no la productividad individual del personal.
    </p>
    <br><br>
    """



    # ==========================================================
    # HTML DEL CORREO
    # ==========================================================

    html_body = f"""
    <html>

    <body style="font-family:Arial;font-size:11pt;">

    <p><b>Cordial saludo</b></p>

    <p style="font-family:Arial;font-size:11pt;">
    
    Se presenta el informe de los trabajos realizados en cada territorio desde abril de 2026, cuando entró en operación el nuevo esquema de trabajo con camionetas y motos.

    </p>

    <br>

    <h3 style="color:#0066CC;">
    Cantidad ejecutada por items
    </h3>

    
    <img src="cid:grafica_cantidad" width="700">
    <br><br>

        
    <h3 style="color:#0066CC;">
    Comparativo de Cantidades por Actividad y Recurso
    </h3>
    
    Los valores resaltados en rojo corresponden a las menores ejecuciones.
    
    <br><br>
    {html_comparativo_cantidades}

    <br><br>


    <h3 style="color:#0066CC;">
    Productividad Operativa (Cantidades por Hora)
    </h3>

    Los valores resaltados en rojo corresponden a los menores niveles de productividad por hora

    <br><br>
    
    {html_productividad}



    {html_conclusion_productividad}

    <h3 style="color:#0066CC;">
    Capacidad Promedio de Producción por Hora
    </h3>



    Este indicador refleja la producción promedio por hora de cada actividad, calculada a partir de la información consolidada de los territorios Metropolitana, Norte, Oriente y Sur. Valores más altos indican una mayor capacidad de ejecución por hora.

    <br><br>
    
    {html_resumen_productividad}

    <h3 style="color:#0066CC;">
    Eficiencia Individual por Podador
    </h3>

    {html_eficiencia_podador}  
    
   <br><br>
  
  <h3 style="color:#0066CC;">
    Kilómetros Intervenidos por Item y Territorio
    </h3>

   <img src="cid:grafica_km" width="700">


    <br><br>

    <h3 style="color:#0066CC;">
    Comparativo de Kilómetros Intervenidos por Actividad y Recurso
    </h3>

    
    En color rojo se resaltan los mayores kilómetros intervenidos acumulados desde abril de 2026, identificando los recursos y territorios con la mayor cobertura de intervención para cada actividad

    <br><br>

    {html_comparativo_km}

    {html_conclusion_km}

    <h3 style="color:#0066CC;font-family:Arial;">
    Kilómetros Ejecutados Mes a Mes
    </h3>

    <p style="font-family:Arial;font-size:11pt;">
    La gráfica presenta la evolución mensual de los kilómetros ejecutados, identificando el comportamiento operativo antes y después de la implementación del nuevo esquema operativo en abril de 2026 (2 camionetas y 2 motos).
    </p>

    <img src="cid:grafica_km_mes" width="700">
    

    <br><br>

    <h3 style="color:#0066CC;">
    Capacidad Operativa Real (KM por Día territorio)
    </h3>

    <p style="font-family:Arial;font-size:11pt;">
    La tabla presenta los kilómetros promedio intervenidos por día trabajado para cada tipo de recurso en cada territorio, considerando la información acumulada desde abril de 2026. Valores más altos indican una mayor capacidad operativa diaria y una mayor cobertura promedio de intervención por jornada de trabajo.
    </p>
    <br><br>
    
    {html_capacidad_km_dia}

    
    <br><br>

    <h3 style="color:#0066CC;">
    Capacidad Operativa Global por Recurso
    </h3>

    La tabla presenta la capacidad operativa promedio de cada tipo de recurso, expresada en kilómetros intervenidos por día trabajado, consolidando la información de todos los territorios desde abril de 2026. Valores más altos indican una mayor cobertura promedio diaria de intervención.
 <br><br>   
    {html_capacidad_global_km_dia}
    
  
    
    <br><br>


    <h3 style="color:#0066CC;">
    Rendimiento Operativo por Grupo
    </h3>

    La tabla presenta el rendimiento operativo de cada grupo desde abril de 2026, mostrando los días efectivamente trabajados, los kilómetros ejecutados y el promedio de kilómetros intervenidos por día. Este indicador permite comparar la capacidad de cobertura diaria entre grupos, independientemente de la fecha de ingreso o del número total de días trabajados.
    <br><br>
    {html_rendimiento_grupos}

    {html_conclusion_grupos}
    
        

    <h3 style="color:#0066CC;">
    Rendimiento Operativo del Territorio {territorio}
    </h3>

    La tabla presenta el rendimiento operativo de los grupos asignados al territorio desde abril de 2026. El indicador KM por Día representa la cobertura promedio diaria alcanzada por cada grupo y se calcula dividiendo los kilómetros ejecutados entre los días efectivamente trabajados.

    <br><br>

    {html_rendimiento_territorio}
    



    <h3 style="color:#0066CC;">
 Resumen Operativo por Tipo de Recurso - {territorio}
    </h3>

    La tabla consolida el desempeño operativo de los recursos del territorio desde abril de 2026, mostrando el número de grupos, los kilómetros ejecutados, los días efectivamente trabajados y la capacidad promedio de intervención diaria (KM por Día).

    <br><br>

    {html_resumen_recursos_territorio}

   {nota_metodologica}

    <h3 style="color:#0066CC;">
    Desempeño Operativo por Grupo (Sin Rocería) - {territorio}
    </h3>

    La tabla consolida las actividades de poda, tala y guaduas ejecutadas por cada grupo desde abril de 2026, excluyendo la rocería. Incluye la participación operativa, productividad diaria, productividad por hora y cobertura promedio diaria de intervención.
 <br><br>
    {html_productividad_grupo}

    

    <h3 style="color:#0066CC;">
    Kilómetros Ejecutados Mes a Mes - {territorio}
    </h3>

    <p style="font-family:Arial;font-size:11pt;">
    La gráfica presenta el comportamiento histórico mensual de los kilómetros ejecutados en el territorio {territorio}, permitiendo comparar el desempeño anterior y posterior a la implementación del esquema operativo iniciado en abril de 2026.
    </p>

    <img src="cid:grafica_km_mes_zona" width="700">


    {html_conclusion_rendimiento} 
    <br><br>
   
    </body>

    </html>
    """

    

    


    print("Ruta horas:")
    print(ruta_grafica_horas)

    print("Existe:")
    print(os.path.exists(ruta_grafica_horas))


# ======================================================
# DESTINATARIOS DEL TERRITORIO
# ======================================================


    if territorio in correos_territorio:

        mail.To = correos_territorio[territorio]["to"]
        mail.CC = correos_territorio[territorio]["cc"]

        mail.Subject = (
            f"Informe Ejecutivo de Vegetación - {territorio}"
        )

        mail.HTMLBody = (
            html_body
            + "<br><br>"
            + firma
        )

        # Guarda el correo como borrador con los datos generados
        # y luego lo abre para revisión final si Outlook está disponible.
        mail.Save()
        mail.Display()

    #<img src="cid:grafica_cantidad" width="5 {html_resumen_territorio}