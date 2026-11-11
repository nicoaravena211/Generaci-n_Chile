"""
Consulta las APIs del Coordinador Eléctrico Nacional y deja los datos listos
para la página web (docs/index.html) en docs/data/.

Genera:
  docs/data/datos.json            -> mensual, diario, budgets, carrusel
  docs/data/horario_YYYYMM.json   -> series horarias (para el Excel PRMTE)

Uso:
  python scripts/actualizar_datos.py            # uso diario
  python scripts/actualizar_datos.py --forzar   # re-consulta todo el histórico horario

Variables de entorno requeridas: API_KEY_PRMTE y API_KEY_CMG.
"""
import argparse
import calendar
import json
import os
import sys
import time
import unicodedata
from datetime import datetime
from pathlib import Path

import pandas as pd
import requests

# =========================================================
# CONFIGURACION
# =========================================================
RAIZ = Path(__file__).resolve().parent.parent
CARPETA_DATOS = RAIZ / "docs" / "data"
ARCHIVO_EXCEL = RAIZ / "budgets.xlsx"

API_KEY_PRMTE = os.environ.get("API_KEY_PRMTE")
API_KEY_CMG = os.environ.get("API_KEY_CMG")


def ahora():
    """Hora actual de Chile, sin zona horaria (los runners de GitHub usan UTC)."""
    return pd.Timestamp.now(tz="America/Santiago").tz_localize(None)


ANIO_OPERACIONAL = 2026
FECHA_INICIO = f"{ANIO_OPERACIONAL}-01"
_hoy = ahora()
if _hoy.year == ANIO_OPERACIONAL:
    FECHA_FIN = _hoy.strftime("%Y-%m")
elif _hoy.year > ANIO_OPERACIONAL:
    FECHA_FIN = f"{ANIO_OPERACIONAL}-12"
else:
    FECHA_FIN = f"{ANIO_OPERACIONAL}-01"

PRECIO_PPA = 83.49
TIPO_CMG = "PRELIMINAR"
BARRA_CMG_BESS = "M.ELENA_______220"
DIAS_CARRUSEL = 30
INCLUIR_DIA_ACTUAL_CARRUSEL = True

HOJA_BUDGETS_BESS = "Budgets"
HOJA_CONSOLIDADO_BESS = "Consolidado"
HOJA_BUDGETS_PROYECTOS = "Budgets_Proyectos"
HOJA_CONSOLIDADO_PROYECTOS = "Consolidado_Proyectos"
URL_PRMTE = "https://medidas.api.coordinador.cl/medidas-v2/measurement"
URL_CMG = "https://sipub.api.coordinador.cl/costo-marginal-real/v4/findByDate"

PROYECTOS = {
    "MARIA ELENA PFV": [{"mpid": "MARELENA_220_JT1_GSS", "canal": 3}],
    "SAN PEDRO III": [{"mpid": "SOLRJAMA_220_JT1_RUC", "canal": 3}, {"mpid": "SOLRJAMA_220_JT2_RUC", "canal": 3}],
    "DOÑA CARMEN": [{"mpid": "CDNCARMN_220_J1_ECM", "canal": 3}],
    "LA QUINTA": [{"mpid": "CABRERO_023_PMGD7_QTC", "canal": 3}],
    "LA PERLA": [{"mpid": "LSANGLES_013_PMGD7_LAP", "canal": 3}],
    "LA HUERTA": [{"mpid": "PARRONAL_013_PMGD2_HUE", "canal": 3}],
    "CHACAICO": [{"mpid": "LSANGLES_013_PMGD8_CCC", "canal": 3}],
    "SANCLEMENTE": [{"mpid": "SCLMENTE_013_PMGD4_CFL", "canal": 3}],
    "TRILALEO": [{"mpid": "CHOLGUAN_015_PMGD6_YTL", "canal": 3}],
    "COLLANCO": [{"mpid": "CNSTUCON_023_PMGD4_OAO", "canal": 3}],
}
BESS_MPIDS = ["MARELENA_023_E1_GSS", "MARELENA_023_E7_GSS", "MARELENA_023_E8_GSS", "MARELENA_023_E11_GSS"]
PROYECTOS_PMGD = ["LA QUINTA", "LA PERLA", "LA HUERTA", "CHACAICO", "SANCLEMENTE", "TRILALEO", "COLLANCO"]
ORDEN_GRAFICOS = ["MARIA ELENA PFV", "SAN PEDRO III", "DOÑA CARMEN", "CONSOLIDADO PMGDS", "LA QUINTA", "LA PERLA",
                  "LA HUERTA", "CHACAICO", "SANCLEMENTE", "TRILALEO", "COLLANCO"]
ORDEN_CARRUSEL = ["MARIA ELENA BESS"] + list(PROYECTOS.keys())
ORDEN_DESCARGABLE_MENSUAL = ["LA QUINTA", "LA PERLA", "LA HUERTA", "CHACAICO", "SANCLEMENTE", "TRILALEO", "COLLANCO",
                             "DOÑA CARMEN", "MARIA ELENA PFV", "SAN PEDRO III"]
ORDEN_ANUAL = ["MARIA ELENA BESS", "MARIA ELENA PFV", "SAN PEDRO III", "DOÑA CARMEN", "LA QUINTA", "LA PERLA",
               "LA HUERTA", "CHACAICO", "SANCLEMENTE", "TRILALEO", "COLLANCO"]

# Barra de costo marginal de cada instalación (para el CMG horario del PRMTE y el reporte semanal)
BARRAS_CMG = {
    "MARIA ELENA PFV": "M.ELENA_______220",
    "BESS MARIA ELENA": "M.ELENA_______220",
    "SAN PEDRO III": "SOLARJAMA_____220",
    "DOÑA CARMEN": "DONA_CARMEN___220",
}
TIPOS_CMG = {"Preliminar": "PRELIMINAR", "Definitivo": "DEFINITIVO"}
# Versión del formato de los archivos horarios (2 = incluye CMG e ingresos)
VERSION_HORARIO = 2

# Carga horaria del BESS menor a este valor (kWh) se considera offset y pasa a 0
UMBRAL_OFFSET_CARGA_BESS_KWH = 450

# =========================================================
# UTILIDADES
# =========================================================
def quitar_tildes(texto):
    return "".join(c for c in unicodedata.normalize("NFD", str(texto)) if unicodedata.category(c) != "Mn")


def normalizar_nombre_columna(nombre):
    nombre = quitar_tildes(nombre).strip().replace(" ", "_").replace("-", "_")
    while "__" in nombre:
        nombre = nombre.replace("__", "_")
    return nombre.lower()


def normalizar_columnas_excel(df):
    eq = {"ano": "Anio", "anio": "Anio", "mes": "Mes", "proyecto": "Proyecto",
          "budget_generacion_mwh": "Budget_Generacion_MWh", "budget_ppa_mwh": "Budget_PPA_MWh",
          "energia_ppa_mwh": "Energia_PPA_MWh", "energia_spot_mwh": "Energia_Spot_MWh",
          "ingreso_ppa_usd": "Ingreso_PPA_USD", "ingreso_spot_usd": "Ingreso_Spot_USD",
          "generacion_mwh": "Generacion_MWh"}
    return df.rename(columns={c: eq.get(normalizar_nombre_columna(c), str(c).strip()) for c in df.columns})


def convertir_numerico(serie):
    if pd.api.types.is_numeric_dtype(serie):
        return pd.to_numeric(serie, errors="coerce")
    return pd.to_numeric(serie.astype(str).str.strip().str.replace(" ", "", regex=False)
                         .str.replace(",", ".", regex=False), errors="coerce")


def parse_fecha_safe(serie):
    return pd.to_datetime(serie.astype(str).str[:19], errors="coerce")


def generar_periodos(fecha_inicio, fecha_fin):
    inicio, fin = datetime.strptime(fecha_inicio, "%Y-%m"), datetime.strptime(fecha_fin, "%Y-%m")
    periodos, actual = [], inicio
    while actual <= fin:
        periodos.append(actual.strftime("%Y%m"))
        actual = actual.replace(year=actual.year + 1, month=1) if actual.month == 12 else actual.replace(month=actual.month + 1)
    return periodos


def periodo_a_texto(periodo):
    meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre",
             "Noviembre", "Diciembre"]
    return f"{meses[int(periodo[4:6]) - 1]} {int(periodo[:4])}"


def a_registros(df, columnas=None):
    """DataFrame -> lista de dicts lista para JSON (NaN -> null, fechas -> texto)."""
    if df is None or df.empty:
        return []
    d = df[columnas].copy() if columnas else df.copy()
    for c in d.columns:
        if pd.api.types.is_datetime64_any_dtype(d[c]):
            d[c] = d[c].dt.strftime("%Y-%m-%d")
    return json.loads(d.to_json(orient="records", force_ascii=False, double_precision=4))


def completar_con_previo(df_nuevo, registros_previos, claves, filtro=None):
    """Si hoy una consulta falla, conserva lo que ya estaba guardado de la corrida anterior."""
    if not registros_previos:
        return df_nuevo
    previo = pd.DataFrame(registros_previos)
    if "Fecha_Dia" in previo.columns:
        previo["Fecha_Dia"] = pd.to_datetime(previo["Fecha_Dia"], errors="coerce")
    if filtro is not None:
        previo = previo[filtro(previo)]
    if previo.empty:
        return df_nuevo
    if df_nuevo is None or df_nuevo.empty:
        return previo
    return pd.concat([df_nuevo, previo], ignore_index=True).drop_duplicates(claves, keep="first")


def cargar_previo():
    archivo = CARPETA_DATOS / "datos.json"
    if archivo.exists():
        try:
            return json.loads(archivo.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            pass
    return {}


# =========================================================
# LECTURA EXCEL (budgets.xlsx = base de datos de meses cerrados)
# =========================================================
def leer_hoja_excel(hoja, requeridas):
    if not ARCHIVO_EXCEL.exists():
        raise FileNotFoundError(f"No se encontró {ARCHIVO_EXCEL.name} en la raíz del repositorio.")
    df = normalizar_columnas_excel(pd.read_excel(ARCHIVO_EXCEL, sheet_name=hoja, engine="openpyxl"))
    faltantes = [c for c in requeridas if c not in df.columns]
    if faltantes:
        raise ValueError(f"Faltan columnas en {hoja}: " + ", ".join(faltantes))
    return df[requeridas].copy()


def cargar_budgets_bess():
    req = ["Anio", "Mes", "Budget_Generacion_MWh", "Budget_PPA_MWh"]
    df = leer_hoja_excel(HOJA_BUDGETS_BESS, req)
    for c in req:
        df[c] = convertir_numerico(df[c])
    df = df.dropna(subset=req)
    df[["Anio", "Mes"]] = df[["Anio", "Mes"]].astype(int)
    if df.empty or not df["Mes"].between(1, 12).all() or df.duplicated(["Anio", "Mes"]).any():
        raise ValueError("La hoja Budgets no contiene información válida o tiene duplicados.")
    return df


def cargar_consolidado_bess():
    req = ["Anio", "Mes", "Energia_PPA_MWh", "Energia_Spot_MWh", "Ingreso_PPA_USD", "Ingreso_Spot_USD"]
    df = leer_hoja_excel(HOJA_CONSOLIDADO_BESS, req)
    for c in req:
        df[c] = convertir_numerico(df[c])
    df = df.dropna(subset=req)
    df[["Anio", "Mes"]] = df[["Anio", "Mes"]].astype(int)
    if df.duplicated(["Anio", "Mes"]).any():
        raise ValueError("Hay meses duplicados en Consolidado.")
    df["Energia_Total_MWh"] = df["Energia_PPA_MWh"] + df["Energia_Spot_MWh"]
    df["Ingreso_Total_USD"] = df["Ingreso_PPA_USD"] + df["Ingreso_Spot_USD"]
    df["Fuente"] = "Excel consolidado"
    return df


def cargar_budgets_proyectos():
    req = ["Anio", "Mes", "Proyecto", "Budget_Generacion_MWh"]
    df = leer_hoja_excel(HOJA_BUDGETS_PROYECTOS, req)
    for c in ["Anio", "Mes", "Budget_Generacion_MWh"]:
        df[c] = convertir_numerico(df[c])
    df["Proyecto"] = df["Proyecto"].astype(str).str.strip().str.upper()
    df = df.dropna(subset=req)
    df[["Anio", "Mes"]] = df[["Anio", "Mes"]].astype(int)
    desconocidos = sorted(set(df["Proyecto"]) - set(PROYECTOS))
    if desconocidos or df.duplicated(["Anio", "Mes", "Proyecto"]).any():
        raise ValueError(f"Budgets_Proyectos contiene proyectos no reconocidos ({desconocidos}) o duplicados.")
    return df


def cargar_consolidado_proyectos():
    req = ["Anio", "Mes", "Proyecto", "Generacion_MWh"]
    df = leer_hoja_excel(HOJA_CONSOLIDADO_PROYECTOS, req)
    for c in ["Anio", "Mes", "Generacion_MWh"]:
        df[c] = convertir_numerico(df[c])
    df["Proyecto"] = df["Proyecto"].astype(str).str.strip().str.upper()
    df = df.dropna(subset=req)
    df[["Anio", "Mes"]] = df[["Anio", "Mes"]].astype(int)
    desconocidos = sorted(set(df["Proyecto"]) - set(PROYECTOS))
    if desconocidos or df.duplicated(["Anio", "Mes", "Proyecto"]).any() or (df["Generacion_MWh"] < 0).any():
        raise ValueError(f"Consolidado_Proyectos contiene registros inválidos ({desconocidos}).")
    df["Fuente"] = "Excel consolidado"
    return df


# =========================================================
# API
# =========================================================
_CACHE_PRMTE = {}  # evita repetir la misma consulta dentro de una corrida


def obtener_datos_prmte(periodo, mpid, canal):
    clave = (periodo, mpid, canal)
    if clave in _CACHE_PRMTE:
        return _CACHE_PRMTE[clave], None
    params = {"channelId": canal, "measurePointId": mpid, "period": periodo, "user_key": API_KEY_PRMTE}
    ultimo_error = None
    for intento in range(3):
        try:
            r = requests.get(URL_PRMTE, params=params, timeout=60)
            if r.status_code == 200:
                data = r.json()
                if isinstance(data, list) and data:
                    _CACHE_PRMTE[clave] = data
                    return data, None
                ultimo_error = "respuesta sin registros"
            elif r.status_code == 429:
                ultimo_error = "límite temporal de solicitudes (HTTP 429)"
            elif r.status_code >= 500:
                ultimo_error = f"error temporal del servidor (HTTP {r.status_code})"
            else:
                ultimo_error = f"respuesta HTTP {r.status_code}"
        except requests.Timeout:
            ultimo_error = "tiempo de espera agotado"
        except requests.ConnectionError:
            ultimo_error = "error de conexión"
        except requests.RequestException as error:
            ultimo_error = f"error de comunicación ({type(error).__name__})"
        except ValueError:
            ultimo_error = "respuesta JSON inválida"
        if intento < 2:
            time.sleep(3 * (intento + 1))
    return None, ultimo_error


def extraer_mediciones(datos, canal):
    registros, campo = [], f"channel{canal}"
    if datos is None:
        return registros
    for bloque in datos:
        for m in bloque.get("measurement", []):
            registros.append({"Fecha": m.get("dateRange"), "Energia_kWh": m.get(campo)})
    return registros


_CACHE_CMG = {}


def descargar_cmg_mes(periodo, tipo=TIPO_CMG, barra=BARRA_CMG_BESS):
    """CMG (cada 15 min) de una barra para un mes. tipo: PRELIMINAR o DEFINITIVO."""
    clave = (periodo, tipo, barra)
    if clave in _CACHE_CMG:
        return _CACHE_CMG[clave]
    anio, mes = int(periodo[:4]), int(periodo[4:6])
    inicio = pd.Timestamp(anio, mes, 1)
    fin = pd.Timestamp(anio, mes, calendar.monthrange(anio, mes)[1])
    hoy_n = ahora().normalize()
    if anio == hoy_n.year and mes == hoy_n.month:
        fin = min(fin, hoy_n)
    registros, pagina = [], 0
    while True:
        params = {"startDate": inicio.strftime("%Y-%m-%d"), "endDate": fin.strftime("%Y-%m-%d"), "page": pagina,
                  "limit": 5000, "type": tipo, "bar_transf": barra, "user_key": API_KEY_CMG}
        try:
            r = requests.get(URL_CMG, params=params, timeout=60)
            if r.status_code != 200:
                return None
            datos = r.json().get("data", [])
        except (requests.RequestException, ValueError):
            return None
        if not datos:
            break
        registros.extend(datos)
        if len(datos) < 5000:
            break
        pagina += 1
    if not registros:
        return None
    df = pd.DataFrame(registros)
    req = ["fecha", "hra", "min", "cmg_usd_mwh_"]
    if not all(c in df.columns for c in req):
        return None
    df["hra"] = convertir_numerico(df["hra"]).fillna(0)
    df["min"] = convertir_numerico(df["min"]).fillna(0)
    df["CMG_USD_MWh"] = convertir_numerico(df["cmg_usd_mwh_"]).fillna(0)
    df["Fecha"] = pd.to_datetime(df["fecha"], errors="coerce") + pd.to_timedelta(df["hra"], unit="h") + pd.to_timedelta(df["min"], unit="m")
    resultado = df[["Fecha", "CMG_USD_MWh"]].dropna(subset=["Fecha"]).groupby("Fecha", as_index=False).agg(CMG_USD_MWh=("CMG_USD_MWh", "mean"))
    _CACHE_CMG[clave] = resultado
    return resultado


def obtener_pendientes_proyectos(df_consolidado, periodos):
    existentes = {(str(f.Proyecto), f"{int(f.Anio):04d}{int(f.Mes):02d}") for f in df_consolidado.itertuples()}
    return {p: [per for per in periodos if (p, per) not in existentes] for p in PROYECTOS}


def descargar_proyectos_pendientes(pendientes):
    resultados, advertencias = [], []
    for proyecto, periodos in pendientes.items():
        for periodo in periodos:
            registros, correctos = [], 0
            for conf in PROYECTOS[proyecto]:
                datos, error = obtener_datos_prmte(periodo + "012345", conf["mpid"], conf["canal"])
                regs = extraer_mediciones(datos, conf["canal"])
                if datos is None or not regs:
                    advertencias.append(f"{proyecto}, {periodo_a_texto(periodo)}, {conf['mpid']}: {error or 'sin mediciones'}")
                    continue
                correctos += 1
                registros.extend(regs)
            if correctos != len(PROYECTOS[proyecto]) or not registros:
                continue
            df = pd.DataFrame(registros)
            df["Fecha"] = parse_fecha_safe(df["Fecha"])
            df["Energia_kWh"] = convertir_numerico(df["Energia_kWh"]).fillna(0)
            df = df.dropna(subset=["Fecha"]).groupby("Fecha", as_index=False).agg(Energia_kWh=("Energia_kWh", "sum"))
            resultados.append({"Anio": int(periodo[:4]), "Mes": int(periodo[4:6]), "Proyecto": proyecto,
                               "Generacion_MWh": df["Energia_kWh"].abs().sum() / 1000, "Fuente": "API"})
    return pd.DataFrame(resultados, columns=["Anio", "Mes", "Proyecto", "Generacion_MWh", "Fuente"]), sorted(set(advertencias))


def procesar_bess_api(periodos):
    mensuales, diarios, advertencias = [], [], []
    for periodo in periodos:
        registros, correctos = [], 0
        for mpid in BESS_MPIDS:
            datos, error = obtener_datos_prmte(periodo + "012345", mpid, 3)
            regs = extraer_mediciones(datos, 3)
            if datos is None or not regs:
                advertencias.append(f"BESS, {periodo_a_texto(periodo)}, {mpid}: {error or 'sin mediciones'}")
                continue
            correctos += 1
            registros.extend(regs)
        if correctos != len(BESS_MPIDS) or not registros:
            continue
        df = pd.DataFrame(registros)
        df["Fecha"] = parse_fecha_safe(df["Fecha"])
        df["Energia_kWh"] = convertir_numerico(df["Energia_kWh"]).fillna(0)
        df = df.dropna(subset=["Fecha"]).groupby("Fecha", as_index=False).agg(Energia_kWh=("Energia_kWh", "sum"))
        df["Energia_MWh"] = df["Energia_kWh"].abs() / 1000
        df["Tipo"] = "PPA"
        df.loc[(df["Fecha"].dt.hour >= 6) & (df["Fecha"].dt.hour < 21), "Tipo"] = "SPOT"
        df["Fecha_Dia"] = df["Fecha"].dt.normalize()
        diario = (df.groupby(["Fecha_Dia", "Tipo"], as_index=False).agg(Energia_MWh=("Energia_MWh", "sum"))
                  .pivot_table(index="Fecha_Dia", columns="Tipo", values="Energia_MWh", fill_value=0).reset_index())
        diario.columns.name = None
        for c in ["PPA", "SPOT"]:
            if c not in diario.columns:
                diario[c] = 0
        diario = diario.rename(columns={"PPA": "Energia_PPA_MWh", "SPOT": "Energia_Spot_MWh"})
        diario["Energia_Total_MWh"] = diario["Energia_PPA_MWh"] + diario["Energia_Spot_MWh"]
        diario["Anio"] = diario["Fecha_Dia"].dt.year
        diario["Mes"] = diario["Fecha_Dia"].dt.month
        diarios.append(diario)
        cmg = descargar_cmg_mes(periodo)
        if cmg is None or cmg.empty:
            advertencias.append(f"BESS, {periodo_a_texto(periodo)}: CMG no disponible; energía disponible, ingresos Spot no valorizados")
            df["CMG_USD_MWh"] = 0.0
        else:
            df = df.merge(cmg, on="Fecha", how="left")
            df["CMG_USD_MWh"] = df["CMG_USD_MWh"].fillna(0)
        df["Precio"] = PRECIO_PPA
        df.loc[df["Tipo"] == "SPOT", "Precio"] = df["CMG_USD_MWh"]
        df["Ingreso_USD"] = df["Energia_MWh"] * df["Precio"]
        eppa = df.loc[df["Tipo"] == "PPA", "Energia_MWh"].sum()
        espot = df.loc[df["Tipo"] == "SPOT", "Energia_MWh"].sum()
        ippa = df.loc[df["Tipo"] == "PPA", "Ingreso_USD"].sum()
        ispot = df.loc[df["Tipo"] == "SPOT", "Ingreso_USD"].sum()
        mensuales.append({"Anio": int(periodo[:4]), "Mes": int(periodo[4:6]), "Energia_PPA_MWh": eppa,
                          "Energia_Spot_MWh": espot, "Energia_Total_MWh": eppa + espot, "Ingreso_PPA_USD": ippa,
                          "Ingreso_Spot_USD": ispot, "Ingreso_Total_USD": ippa + ispot, "Fuente": "API"})
    cols_m = ["Anio", "Mes", "Energia_PPA_MWh", "Energia_Spot_MWh", "Energia_Total_MWh", "Ingreso_PPA_USD",
              "Ingreso_Spot_USD", "Ingreso_Total_USD", "Fuente"]
    cols_d = ["Fecha_Dia", "Energia_PPA_MWh", "Energia_Spot_MWh", "Energia_Total_MWh", "Anio", "Mes"]
    return (pd.DataFrame(mensuales, columns=cols_m),
            (pd.concat(diarios, ignore_index=True) if diarios else pd.DataFrame(columns=cols_d)),
            sorted(set(advertencias)))


def procesar_registros_diarios(registros, proyecto):
    df = pd.DataFrame(registros)
    df["Fecha"] = parse_fecha_safe(df["Fecha"])
    df["Energia_kWh"] = convertir_numerico(df["Energia_kWh"]).fillna(0)
    df = df.dropna(subset=["Fecha"])
    if df.empty:
        return pd.DataFrame()
    df["Fecha_Dia"] = df["Fecha"].dt.normalize()
    diario = df.groupby("Fecha_Dia", as_index=False).agg(Energia_MWh=("Energia_kWh", lambda s: s.abs().sum() / 1000))
    diario["Proyecto"] = proyecto
    diario["Anio"] = diario["Fecha_Dia"].dt.year
    diario["Mes"] = diario["Fecha_Dia"].dt.month
    diario["Fuente"] = "API"
    return diario[["Fecha_Dia", "Anio", "Mes", "Proyecto", "Energia_MWh", "Fuente"]]


def descargar_diarios_carrusel(periodos, fecha_inicio, fecha_fin):
    resultados, advertencias = [], []
    configuraciones = {"MARIA ELENA BESS": [{"mpid": m, "canal": 3} for m in BESS_MPIDS]}
    configuraciones.update(PROYECTOS)
    for proyecto in ORDEN_CARRUSEL:
        diarios = []
        for periodo in periodos:
            registros, correctos = [], 0
            for conf in configuraciones[proyecto]:
                datos, error = obtener_datos_prmte(periodo + "012345", conf["mpid"], conf["canal"])
                regs = extraer_mediciones(datos, conf["canal"])
                if datos is None or not regs:
                    advertencias.append(f"Carrusel {proyecto}, {periodo_a_texto(periodo)}, {conf['mpid']}: {error or 'sin mediciones'}")
                    continue
                correctos += 1
                registros.extend(regs)
            if correctos == len(configuraciones[proyecto]) and registros:
                d = procesar_registros_diarios(registros, proyecto)
                if not d.empty:
                    diarios.append(d)
        if not diarios:
            advertencias.append(f"Carrusel {proyecto}: sin datos diarios entre {fecha_inicio:%d-%m-%Y} y {fecha_fin:%d-%m-%Y}")
            continue
        d = pd.concat(diarios, ignore_index=True)
        d = d[(d["Fecha_Dia"] >= fecha_inicio) & (d["Fecha_Dia"] <= fecha_fin)].copy()
        if d.empty:
            advertencias.append(f"Carrusel {proyecto}: sin mediciones en los últimos {DIAS_CARRUSEL} días")
            continue
        resultados.append(d.groupby(["Fecha_Dia", "Anio", "Mes", "Proyecto", "Fuente"], as_index=False).agg(Energia_MWh=("Energia_MWh", "sum")))
    cols = ["Fecha_Dia", "Anio", "Mes", "Proyecto", "Energia_MWh", "Fuente"]
    return (pd.concat(resultados, ignore_index=True).sort_values(["Proyecto", "Fecha_Dia"]) if resultados
            else pd.DataFrame(columns=cols)), sorted(set(advertencias))


# =========================================================
# SERIES HORARIAS (para el Excel PRMTE del navegador)
# =========================================================
def obtener_serie_prmte_rango(configuraciones, canal, nombre_columna, fecha_inicio, fecha_fin):
    """Descarga y suma los MPID de una instalación para un canal y un rango (datos cada 15 min)."""
    registros = []
    periodos = [p.strftime("%Y%m") for p in pd.period_range(pd.Timestamp(fecha_inicio), pd.Timestamp(fecha_fin), freq="M")]
    for conf in configuraciones:
        canal_consulta = conf.get("canal", canal)
        for periodo in periodos:
            datos, _ = obtener_datos_prmte(periodo + "012345", conf["mpid"], canal_consulta)
            mediciones = extraer_mediciones(datos, canal_consulta)
            if datos is None or not mediciones:
                continue
            registros.extend(mediciones)
    if not registros:
        return pd.DataFrame(columns=["Fecha", nombre_columna])
    df = pd.DataFrame(registros)
    df["Fecha"] = parse_fecha_safe(df["Fecha"])
    df[nombre_columna] = convertir_numerico(df["Energia_kWh"]).fillna(0)
    df = df.dropna(subset=["Fecha"])
    inicio = pd.Timestamp(fecha_inicio).normalize()
    fin_exclusivo = pd.Timestamp(fecha_fin).normalize() + pd.Timedelta(days=1)
    df = df[(df["Fecha"] >= inicio) & (df["Fecha"] < fin_exclusivo)].copy()
    if df.empty:
        return pd.DataFrame(columns=["Fecha", nombre_columna])
    return df.groupby("Fecha", as_index=False)[nombre_columna].sum().sort_values("Fecha")


def a_horario(df, columna, indice_horario):
    """Agrega la serie de 15 min a una hora y la alinea al calendario completo."""
    if df.empty:
        return pd.Series(float("nan"), index=indice_horario, name=columna)
    return (df.set_index("Fecha").resample("h")[columna].sum(min_count=1)
            .reindex(indice_horario).rename(columna))


def lista_json(serie):
    return [None if pd.isna(v) else round(float(v), 3) for v in serie]


def cmg_a_horario(df_cmg, indice):
    """CMG de 15 min -> promedio horario."""
    if df_cmg is None or df_cmg.empty:
        return pd.Series(float("nan"), index=indice)
    return df_cmg.set_index("Fecha")["CMG_USD_MWh"].resample("h").mean().reindex(indice)


def ingreso_a_horario(df_energia, columna, df_cmg, indice):
    """Ingreso spot horario = suma de (energía 15 min x CMG 15 min)."""
    vacio = pd.Series(float("nan"), index=indice)
    if df_energia is None or df_energia.empty or df_cmg is None or df_cmg.empty:
        return vacio
    m = df_energia[["Fecha", columna]].merge(df_cmg, on="Fecha", how="inner")
    if m.empty:
        return vacio
    m["Ingreso"] = m[columna] / 1000 * m["CMG_USD_MWh"]
    return m.set_index("Fecha")["Ingreso"].resample("h").sum(min_count=1).reindex(indice)


def agregar_cmg(series_parque, nombre, periodo, indice, df_energia=None, columna=None):
    """Agrega CMG horario (preliminar y definitivo) y, si hay energía, el ingreso spot."""
    barra = BARRAS_CMG[nombre]
    cmgs = {tipo: descargar_cmg_mes(periodo, codigo, barra) for tipo, codigo in TIPOS_CMG.items()}
    for tipo, df_cmg in cmgs.items():
        series_parque[f"CMG {tipo} USD/MWh"] = lista_json(cmg_a_horario(df_cmg, indice))
    if df_energia is not None:
        for tipo, df_cmg in cmgs.items():
            series_parque[f"Ingreso {tipo} USD"] = lista_json(ingreso_a_horario(df_energia, columna, df_cmg, indice))


def generar_horario_mes(periodo, fecha_tope):
    anio, mes = int(periodo[:4]), int(periodo[4:6])
    inicio = pd.Timestamp(anio, mes, 1)
    fin = min(pd.Timestamp(anio, mes, calendar.monthrange(anio, mes)[1]), fecha_tope)
    indice = pd.date_range(inicio, fin + pd.Timedelta(hours=23), freq="h")

    bess_iny_15 = obtener_serie_prmte_rango([{"mpid": m, "canal": 3} for m in BESS_MPIDS], 3, "Inyección kWh", inicio, fin)
    bess_car_15 = obtener_serie_prmte_rango([{"mpid": m, "canal": 1} for m in BESS_MPIDS], 1, "Carga BESS kWh", inicio, fin)

    series = {}
    for proyecto, conf in PROYECTOS.items():
        df_15 = obtener_serie_prmte_rango(conf, 3, "Inyección kWh", inicio, fin)
        if proyecto == "MARIA ELENA PFV":
            base = df_15.merge(bess_iny_15.rename(columns={"Inyección kWh": "BESS Raw kWh"}), on="Fecha", how="left")
            base["BESS Raw kWh"] = convertir_numerico(base["BESS Raw kWh"]).fillna(0)
            base["Neto PV Red kWh"] = base["Inyección kWh"]
            base.loc[base["BESS Raw kWh"] > 0, "Neto PV Red kWh"] = 0
            series[proyecto] = {
                "Inyección kWh": lista_json(a_horario(base[["Fecha", "Inyección kWh"]], "Inyección kWh", indice)),
                "Neto PV Red kWh": lista_json(a_horario(base[["Fecha", "Neto PV Red kWh"]], "Neto PV Red kWh", indice)),
            }
            agregar_cmg(series[proyecto], proyecto, periodo, indice, base[["Fecha", "Neto PV Red kWh"]], "Neto PV Red kWh")
        else:
            series[proyecto] = {"Inyección kWh": lista_json(a_horario(df_15, "Inyección kWh", indice))}
            if proyecto in BARRAS_CMG:  # San Pedro III y Doña Carmen
                agregar_cmg(series[proyecto], proyecto, periodo, indice, df_15, "Inyección kWh")

    carga = pd.to_numeric(a_horario(bess_car_15, "Carga BESS kWh", indice), errors="coerce")
    carga[carga < UMBRAL_OFFSET_CARGA_BESS_KWH] = 0  # offset: se reemplaza por cero
    series["BESS MARIA ELENA"] = {
        "Inyección kWh": lista_json(a_horario(bess_iny_15, "Inyección kWh", indice)),
        "Carga BESS kWh": lista_json(carga),
    }
    agregar_cmg(series["BESS MARIA ELENA"], "BESS MARIA ELENA", periodo, indice)
    return {"version": VERSION_HORARIO, "periodo": periodo, "fechas": [f.strftime("%Y-%m-%d %H:%M") for f in indice], "series": series}


def horario_actualizado(archivo):
    """True si el archivo ya tiene el formato actual (con CMG)."""
    try:
        return json.loads(archivo.read_text(encoding="utf-8")).get("version", 1) >= VERSION_HORARIO
    except (OSError, ValueError):
        return False


def conservar_previo(nuevo, archivo):
    """Si hoy una serie quedó vacía (consulta fallida) pero antes había datos, se conservan los anteriores."""
    if not archivo.exists():
        return nuevo
    try:
        viejo = json.loads(archivo.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return nuevo
    n = len(nuevo["fechas"])
    for nombre, cols in nuevo["series"].items():
        for col, vals in cols.items():
            if any(v is not None for v in vals):
                continue
            antes = viejo.get("series", {}).get(nombre, {}).get(col)
            if antes and any(v is not None for v in antes):
                cols[col] = (antes + [None] * n)[:n]
    return nuevo


def tiene_datos(horario):
    return any(v is not None for s in horario["series"].values() for col in s.values() for v in col)


# =========================================================
# PRINCIPAL
# =========================================================
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--forzar", action="store_true", help="Re-consulta todo el histórico horario")
    args = ap.parse_args()

    if not API_KEY_PRMTE or not API_KEY_CMG:
        sys.exit("Faltan las variables de entorno API_KEY_PRMTE y/o API_KEY_CMG (secrets del repositorio).")

    CARPETA_DATOS.mkdir(parents=True, exist_ok=True)
    previo = cargar_previo()
    ahora_cl = ahora()

    periodos_disponibles = generar_periodos(FECHA_INICIO, FECHA_FIN)
    fin_car = ahora_cl.normalize() if INCLUIR_DIA_ACTUAL_CARRUSEL else ahora_cl.normalize() - pd.Timedelta(days=1)
    ini_car = fin_car - pd.Timedelta(days=DIAS_CARRUSEL - 1)
    periodos_carrusel = [p.strftime("%Y%m") for p in pd.period_range(ini_car, fin_car, freq="M")]

    try:
        df_budgets_bess = cargar_budgets_bess()
        df_cons_bess = cargar_consolidado_bess()
        df_budgets_proy = cargar_budgets_proyectos()
        df_cons_proy = cargar_consolidado_proyectos()
    except Exception as error:
        sys.exit(f"No fue posible validar budgets.xlsx: {error}")

    # ---- Proyectos solares: meses que no están en el Excel ----
    print("Consultando meses pendientes de proyectos...")
    df_proy_api, adv_proy = descargar_proyectos_pendientes(obtener_pendientes_proyectos(df_cons_proy, periodos_disponibles))
    df_proy_api = completar_con_previo(df_proy_api, previo.get("proyectos_mensual"), ["Anio", "Mes", "Proyecto"],
                                       filtro=lambda d: d["Fuente"] == "API")
    fuentes = [df_cons_proy] + ([df_proy_api] if not df_proy_api.empty else [])
    df_proy = pd.concat(fuentes, ignore_index=True)
    df_proy["Prioridad"] = df_proy["Fuente"].map({"API": 1, "Excel consolidado": 2}).fillna(2)
    df_proy = (df_proy.sort_values("Prioridad").drop_duplicates(["Anio", "Mes", "Proyecto"], keep="last")
               .drop(columns="Prioridad").sort_values(["Proyecto", "Anio", "Mes"]))

    # ---- BESS ----
    periodos_excel = {f"{int(f.Anio):04d}{int(f.Mes):02d}" for f in df_cons_bess.itertuples()}
    pendientes_bess = [p for p in periodos_disponibles if p not in periodos_excel]
    print("Consultando meses pendientes del BESS...")
    df_bess_api, df_bess_diario, adv_bess = procesar_bess_api(pendientes_bess)
    df_bess_api = completar_con_previo(df_bess_api, previo.get("bess_mensual"), ["Anio", "Mes"],
                                       filtro=lambda d: d["Fuente"] == "API")
    df_bess_diario = completar_con_previo(df_bess_diario, previo.get("bess_diario"), ["Fecha_Dia"],
                                          filtro=lambda d: ~(d["Anio"].astype(int) * 100 + d["Mes"].astype(int)).astype(str).isin(periodos_excel))
    fuentes = [df_cons_bess] + ([df_bess_api] if not df_bess_api.empty else [])
    df_bess = pd.concat(fuentes, ignore_index=True)
    df_bess["Prioridad"] = df_bess["Fuente"].map({"API": 1, "Excel consolidado": 2}).fillna(2)
    df_bess = (df_bess.sort_values("Prioridad").drop_duplicates(["Anio", "Mes"], keep="last")
               .drop(columns="Prioridad").sort_values(["Anio", "Mes"]))

    # ---- Carrusel (últimos 30 días) ----
    print("Preparando energía diaria del carrusel...")
    df_carrusel, adv_car = descargar_diarios_carrusel(periodos_carrusel, ini_car, fin_car)
    df_carrusel = completar_con_previo(df_carrusel, previo.get("carrusel"), ["Proyecto", "Fecha_Dia"],
                                       filtro=lambda d: (d["Fecha_Dia"] >= ini_car) & (d["Fecha_Dia"] <= fin_car))
    if not df_carrusel.empty:
        df_carrusel = df_carrusel.sort_values(["Proyecto", "Fecha_Dia"])

    # ---- Series horarias por mes ----
    print("Actualizando series horarias...")
    mes_anterior = (ahora_cl.normalize().replace(day=1) - pd.Timedelta(days=1)).strftime("%Y%m")
    periodos_horario = []
    for p in periodos_disponibles:
        archivo = CARPETA_DATOS / f"horario_{p}.json"
        if archivo.exists() and p < mes_anterior and not args.forzar and horario_actualizado(archivo):
            periodos_horario.append(p)
            continue
        horario = conservar_previo(generar_horario_mes(p, ahora_cl.normalize()), archivo)
        if tiene_datos(horario):
            archivo.write_text(json.dumps(horario, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
            periodos_horario.append(p)
        elif archivo.exists():
            periodos_horario.append(p)  # no se pisa un archivo bueno con uno vacío
            print(f"  {p}: sin datos nuevos, se conserva el archivo anterior")
        else:
            print(f"  {p}: sin datos horarios disponibles")

    # ---- datos.json ----
    advertencias = sorted(set(adv_proy + adv_bess + adv_car))
    datos = {
        "meta": {
            "anio": ANIO_OPERACIONAL,
            "fecha_inicio": FECHA_INICIO,
            "fecha_fin": FECHA_FIN,
            "mes_actual": int(FECHA_FIN[5:7]),
            "precio_ppa": PRECIO_PPA,
            "generado": ahora_cl.strftime("%d-%m-%Y %H:%M") + " (hora de Chile)",
            "fecha_datos": ahora_cl.strftime("%Y-%m-%d"),
            "dias_carrusel": DIAS_CARRUSEL,
            "carrusel_inicio": ini_car.strftime("%Y-%m-%d"),
            "carrusel_fin": fin_car.strftime("%Y-%m-%d"),
            "proyectos": list(PROYECTOS.keys()),
            "proyectos_pmgd": PROYECTOS_PMGD,
            "orden_graficos": ORDEN_GRAFICOS,
            "orden_carrusel": ORDEN_CARRUSEL,
            "orden_mensual": ORDEN_DESCARGABLE_MENSUAL,
            "orden_anual": ORDEN_ANUAL,
            "periodos_horario": periodos_horario,
        },
        "bess_mensual": a_registros(df_bess, ["Anio", "Mes", "Energia_PPA_MWh", "Energia_Spot_MWh", "Energia_Total_MWh",
                                              "Ingreso_PPA_USD", "Ingreso_Spot_USD", "Ingreso_Total_USD", "Fuente"]),
        "bess_diario": a_registros(df_bess_diario.sort_values("Fecha_Dia") if not df_bess_diario.empty else df_bess_diario,
                                   ["Fecha_Dia", "Anio", "Mes", "Energia_PPA_MWh", "Energia_Spot_MWh", "Energia_Total_MWh"]),
        "proyectos_mensual": a_registros(df_proy, ["Anio", "Mes", "Proyecto", "Generacion_MWh", "Fuente"]),
        "budgets_bess": a_registros(df_budgets_bess),
        "budgets_proyectos": a_registros(df_budgets_proy),
        "carrusel": a_registros(df_carrusel, ["Fecha_Dia", "Proyecto", "Energia_MWh"]),
        "advertencias": advertencias,
    }
    (CARPETA_DATOS / "datos.json").write_text(json.dumps(datos, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    print(f"Listo. {len(advertencias)} advertencias.")
    for a in advertencias:
        print(" -", a)


if __name__ == "__main__":
    main()
