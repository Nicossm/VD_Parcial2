"""
limpieza_pideya.py
------------------
Limpia las tres tablas del caso PideYa (usuarios, restaurantes, pedidos)
y deja los archivos listos para cargarlos en Power BI.

Uso (desde la carpeta raíz del proyecto):
    python src/limpieza_pideya.py

Estructura esperada:
    data/raw/usuarios.csv
    data/raw/restaurantes.csv
    data/raw/pedidos.csv

Salidas:
    data/processed/usuarios_limpio.csv
    data/processed/restaurantes_limpio.csv
    data/processed/pedidos_limpio.csv
    data/processed/pideya_limpio.xlsx      (las 3 tablas en un solo Excel)
    data/processed/log_calidad_datos.csv   (qué se corrigió y cuántas filas)

Requisitos: pandas, openpyxl  ->  pip install pandas openpyxl
"""

from pathlib import Path
import unicodedata

import numpy as np
import pandas as pd

# ---------------------------------------------------------------------------
# Rutas (relativas a la raíz del proyecto, no a donde ejecutes el script)
# ---------------------------------------------------------------------------
# Busca los CSV en este orden:
#   1) <carpeta del script>/data/raw          (script suelto en la raíz)
#   2) <carpeta padre>/data/raw               (script dentro de src/)
#   3) <carpeta del script>                   (CSV al lado del script)
AQUI = Path(__file__).resolve().parent
for raw_candidata in [AQUI / "data" / "raw", AQUI.parent / "data" / "raw", AQUI]:
    if (raw_candidata / "usuarios.csv").exists():
        RAW = raw_candidata
        break
else:
    raise FileNotFoundError(
        "No encontré usuarios.csv. Ponlo en 'data/raw/' junto al script "
        f"o en la misma carpeta del script ({AQUI})."
    )
OUT = (RAW.parent / "processed") if RAW.name == "raw" else (AQUI / "processed")
print(f"Leyendo datos desde: {RAW}")
OUT.mkdir(parents=True, exist_ok=True)

# Fecha de corte: un pedido o registro posterior a esta fecha es inválido
FECHA_CORTE = pd.Timestamp("2026-10-31")

# Rango aceptado para el tiempo de entrega (minutos)
TIEMPO_MIN, TIEMPO_MAX = 1, 120

# Bitácora de calidad: cada corrección queda registrada aquí
log = []


def registrar(tabla, problema, accion, filas):
    """Agrega una línea a la bitácora y la muestra por consola."""
    log.append({"tabla": tabla, "problema": problema, "accion": accion, "filas_afectadas": int(filas)})
    print(f"  [{tabla}] {problema}: {filas} fila(s) -> {accion}")


# ---------------------------------------------------------------------------
# Funciones auxiliares
# ---------------------------------------------------------------------------
def sin_tildes(texto: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFD", texto) if unicodedata.category(c) != "Mn")


def limpiar_texto(serie: pd.Series) -> pd.Series:
    """Quita espacios sobrantes y convierte '' en NaN."""
    s = serie.astype("string").str.strip().str.replace(r"\s+", " ", regex=True)
    return s.replace("", pd.NA)


def normalizar_por_catalogo(serie: pd.Series, catalogo: list[str]) -> pd.Series:
    """
    Lleva cada valor a su forma oficial del catálogo, ignorando
    mayúsculas, tildes y espacios. Ej: 'NUNOA', 'ñuñoa ' -> 'Ñuñoa'.
    """
    mapa = {sin_tildes(v).lower(): v for v in catalogo}
    clave = limpiar_texto(serie).map(lambda x: sin_tildes(x).lower() if pd.notna(x) else x)
    return clave.map(mapa)


def parsear_fechas(serie: pd.Series) -> pd.Series:
    """
    Acepta los tres formatos que vienen mezclados:
      2026-06-27  |  27/06/2026  |  27-06-2026   (siempre día antes que mes)
    """
    s = limpiar_texto(serie)
    iso = pd.to_datetime(s, format="%Y-%m-%d", errors="coerce")
    barra = pd.to_datetime(s, format="%d/%m/%Y", errors="coerce")
    guion = pd.to_datetime(s, format="%d-%m-%Y", errors="coerce")
    return iso.fillna(barra).fillna(guion)


def parsear_monto(serie: pd.Series) -> pd.Series:
    """'$11.200' -> 11200 | '12600 CLP' -> 12600 | '8700' -> 8700."""
    s = limpiar_texto(serie).str.upper()
    s = s.str.replace("CLP", "", regex=False).str.replace("$", "", regex=False)
    s = s.str.replace(".", "", regex=False).str.replace(",", "", regex=False).str.strip()
    return pd.to_numeric(s, errors="coerce")


def quitar_duplicados(df, tabla, clave):
    n = df.duplicated().sum()
    if n:
        df = df.drop_duplicates()
        registrar(tabla, "Filas duplicadas exactas", "Eliminadas", n)
    n = df[clave].duplicated().sum()
    if n:
        df = df.drop_duplicates(subset=clave, keep="first")
        registrar(tabla, f"{clave} repetido con datos distintos", "Se conserva la primera aparición", n)
    return df


# Catálogos oficiales
COMUNAS = ["Santiago", "Providencia", "Las Condes", "Ñuñoa", "Maipú", "La Florida"]
CATEGORIAS = ["Comida rápida", "Saludable", "Asiática", "Italiana", "Postres y café"]
METODOS_PAGO = ["Tarjeta", "Efectivo", "Transferencia"]
ESTADOS = ["Entregado", "Cancelado"]
TIPOS_CUENTA = ["Estándar", "Premium"]


# ===========================================================================
# 1. USUARIOS
# ===========================================================================
print("\n== USUARIOS ==")
usuarios = pd.read_csv(RAW / "usuarios.csv", dtype=str)
filas_ini_u = len(usuarios)
usuarios = usuarios.apply(limpiar_texto)

usuarios = quitar_duplicados(usuarios, "usuarios", "id_usuario")
usuarios["id_usuario"] = usuarios["id_usuario"].astype(int)

# Fechas en 3 formatos distintos
formatos_distintos = (~usuarios["fecha_registro"].str.match(r"^\d{4}-\d{2}-\d{2}$")).sum()
usuarios["fecha_registro"] = parsear_fechas(usuarios["fecha_registro"])
registrar("usuarios", "fecha_registro en formato dd/mm/aaaa o dd-mm-aaaa", "Convertida a fecha (aaaa-mm-dd)", formatos_distintos)
n = usuarios["fecha_registro"].isna().sum()
if n:
    registrar("usuarios", "fecha_registro no interpretable", "Queda vacía", n)

# Comuna: mayúsculas, espacios y 'Nunoa' sin ñ
original = usuarios["comuna"].copy()
usuarios["comuna"] = normalizar_por_catalogo(usuarios["comuna"], COMUNAS)
registrar("usuarios", "comuna con mayúsculas/espacios/sin tilde inconsistentes", "Normalizada al nombre oficial",
          (original != usuarios["comuna"]).sum())

# Tipo de cuenta vacío
usuarios["tipo_cuenta"] = normalizar_por_catalogo(usuarios["tipo_cuenta"], TIPOS_CUENTA)
n = usuarios["tipo_cuenta"].isna().sum()
usuarios["tipo_cuenta"] = usuarios["tipo_cuenta"].fillna("Sin información")
registrar("usuarios", "tipo_cuenta vacío", "Rellenado con 'Sin información'", n)

usuarios = usuarios.sort_values("id_usuario").reset_index(drop=True)


# ===========================================================================
# 2. RESTAURANTES
# ===========================================================================
print("\n== RESTAURANTES ==")
restaurantes = pd.read_csv(RAW / "restaurantes.csv", dtype=str)
filas_ini_r = len(restaurantes)

original = restaurantes["nombre_restaurante"].copy()
restaurantes = restaurantes.apply(limpiar_texto)
registrar("restaurantes", "nombre_restaurante con espacios al inicio/final", "Espacios eliminados",
          (original != restaurantes["nombre_restaurante"]).sum())

restaurantes = quitar_duplicados(restaurantes, "restaurantes", "id_restaurante")
restaurantes["id_restaurante"] = restaurantes["id_restaurante"].astype(int)

original = restaurantes["categoria"].copy()
restaurantes["categoria"] = normalizar_por_catalogo(restaurantes["categoria"], CATEGORIAS)
registrar("restaurantes", "categoria con mayúsculas/espacios/sin tilde", "Normalizada",
          (original != restaurantes["categoria"]).sum())

restaurantes["comuna"] = normalizar_por_catalogo(restaurantes["comuna"], COMUNAS)
n = restaurantes["comuna"].isna().sum()
restaurantes["comuna"] = restaurantes["comuna"].fillna("Sin información")
registrar("restaurantes", "comuna vacía", "Rellenada con 'Sin información'", n)

restaurantes = restaurantes.sort_values("id_restaurante").reset_index(drop=True)


# ===========================================================================
# 3. PEDIDOS
# ===========================================================================
print("\n== PEDIDOS ==")
pedidos = pd.read_csv(RAW / "pedidos.csv", dtype=str)
filas_ini_p = len(pedidos)
pedidos = pedidos.apply(limpiar_texto)

pedidos = quitar_duplicados(pedidos, "pedidos", "id_pedido")
for c in ["id_pedido", "id_usuario", "id_restaurante"]:
    pedidos[c] = pedidos[c].astype(int)

# --- Integridad referencial (claves foráneas que no existen) ---
huerf_u = ~pedidos["id_usuario"].isin(usuarios["id_usuario"])
huerf_r = ~pedidos["id_restaurante"].isin(restaurantes["id_restaurante"])
registrar("pedidos", "id_usuario inexistente en Usuarios (ej. 9999)", "Pedido eliminado", huerf_u.sum())
registrar("pedidos", "id_restaurante inexistente en Restaurantes (ej. 999)", "Pedido eliminado", (huerf_r & ~huerf_u).sum())
pedidos = pedidos[~(huerf_u | huerf_r)].copy()

# --- Fechas ---
formatos_distintos = (~pedidos["fecha_pedido"].str.match(r"^\d{4}-\d{2}-\d{2}$")).sum()
pedidos["fecha_pedido"] = parsear_fechas(pedidos["fecha_pedido"])
registrar("pedidos", "fecha_pedido en formato dd/mm/aaaa o dd-mm-aaaa", "Convertida a fecha", formatos_distintos)
fuera = pedidos["fecha_pedido"].isna() | (pedidos["fecha_pedido"] > FECHA_CORTE)
if fuera.sum():
    registrar("pedidos", "fecha_pedido vacía o futura", "Pedido eliminado", fuera.sum())
    pedidos = pedidos[~fuera].copy()

# --- Monto ---
texto_monto = ~pedidos["monto_clp"].str.fullmatch(r"\d+")
pedidos["monto_clp"] = parsear_monto(pedidos["monto_clp"])
registrar("pedidos", "monto_clp con texto ('$11.200', '12600 CLP')", "Convertido a número entero", texto_monto.sum())
n = (pedidos["monto_clp"].isna() | (pedidos["monto_clp"] <= 0)).sum()
if n:
    pedidos.loc[pedidos["monto_clp"] <= 0, "monto_clp"] = np.nan
    registrar("pedidos", "monto_clp vacío o <= 0", "Queda vacío", n)
pedidos["monto_clp"] = pedidos["monto_clp"].astype("Int64")

# --- Tiempo de entrega ---
pedidos["tiempo_entrega_min"] = pd.to_numeric(pedidos["tiempo_entrega_min"], errors="coerce")
vacio_orig = pedidos["tiempo_entrega_min"].isna()
fuera_rango = pedidos["tiempo_entrega_min"].notna() & ~pedidos["tiempo_entrega_min"].between(TIEMPO_MIN, TIEMPO_MAX)
registrar("pedidos", f"tiempo_entrega_min fuera de rango (<{TIEMPO_MIN} o >{TIEMPO_MAX}: -5, 0, 500, 999)",
          "Reemplazado por vacío (no se imputa)", fuera_rango.sum())
pedidos.loc[fuera_rango, "tiempo_entrega_min"] = np.nan
registrar("pedidos", "tiempo_entrega_min vacío en origen", "Se mantiene vacío", vacio_orig.sum())
pedidos["tiempo_entrega_min"] = pedidos["tiempo_entrega_min"].astype("Int64")

# --- Método de pago (mayúsculas, espacios y el error de tipeo 'Tajeta') ---
original = pedidos["metodo_pago"].copy()
pedidos["metodo_pago"] = pedidos["metodo_pago"].replace({"Tajeta": "Tarjeta"})
pedidos["metodo_pago"] = normalizar_por_catalogo(pedidos["metodo_pago"], METODOS_PAGO)
registrar("pedidos", "metodo_pago inconsistente (TARJETA, tarjeta, 'Tajeta', espacios)", "Normalizado",
          (original != pedidos["metodo_pago"]).sum())

# --- Estado ---
original = pedidos["estado"].copy()
pedidos["estado"] = normalizar_por_catalogo(pedidos["estado"], ESTADOS)
registrar("pedidos", "estado con mayúsculas/espacios inconsistentes", "Normalizado",
          (original != pedidos["estado"]).sum())

# --- Calificación ---
pedidos["calificacion"] = pd.to_numeric(pedidos["calificacion"], errors="coerce")
fuera = pedidos["calificacion"].notna() & ~pedidos["calificacion"].between(1, 5)
if fuera.sum():
    pedidos.loc[fuera, "calificacion"] = np.nan
    registrar("pedidos", "calificacion fuera de 1-5", "Reemplazada por vacío", fuera.sum())
cancel = pedidos["estado"] == "Cancelado"
registrar("pedidos", "calificacion vacía en pedidos Cancelados", "Correcto: un pedido cancelado no se califica",
          (pedidos["calificacion"].isna() & cancel).sum())
registrar("pedidos", "calificacion vacía en pedidos Entregados", "Se mantiene vacía (cliente no calificó)",
          (pedidos["calificacion"].isna() & ~cancel).sum())
pedidos["calificacion"] = pedidos["calificacion"].astype("Int64")

# Nota: los pedidos Cancelados traen tiempo_entrega_min. No se borra el dato,
# pero en Power BI los KPIs de tiempo y calificación deben filtrar estado = "Entregado".

# --- Coherencia temporal: pedido anterior al registro del usuario ---
fecha_reg = pedidos["id_usuario"].map(usuarios.set_index("id_usuario")["fecha_registro"])
pedidos["pedido_antes_registro"] = (pedidos["fecha_pedido"] < fecha_reg).astype(int)
registrar("pedidos", "fecha_pedido anterior a la fecha_registro del usuario",
          "Se conserva y se marca en columna pedido_antes_registro = 1", pedidos["pedido_antes_registro"].sum())

# --- Columnas de apoyo para Power BI ---
dias = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
pedidos["dia_semana"] = pedidos["fecha_pedido"].dt.dayofweek.map(dict(enumerate(dias)))
pedidos["dia_semana_num"] = pedidos["fecha_pedido"].dt.dayofweek + 1  # para ordenar en Power BI
pedidos["es_entregado"] = (pedidos["estado"] == "Entregado").astype(int)

pedidos = pedidos.sort_values(["fecha_pedido", "id_pedido"]).reset_index(drop=True)


# ===========================================================================
# 4. VALIDACIONES FINALES
# ===========================================================================
print("\n== VALIDACIONES ==")
assert usuarios["id_usuario"].is_unique
assert restaurantes["id_restaurante"].is_unique
assert pedidos["id_pedido"].is_unique
assert pedidos["id_usuario"].isin(usuarios["id_usuario"]).all()
assert pedidos["id_restaurante"].isin(restaurantes["id_restaurante"]).all()
assert set(pedidos["metodo_pago"].dropna()) <= set(METODOS_PAGO)
assert set(pedidos["estado"].dropna()) <= set(ESTADOS)
print("  Claves únicas, relaciones y catálogos OK")

resumen = pd.DataFrame({
    "tabla": ["usuarios", "restaurantes", "pedidos"],
    "filas_originales": [filas_ini_u, filas_ini_r, filas_ini_p],
    "filas_limpias": [len(usuarios), len(restaurantes), len(pedidos)],
})
print(resumen.to_string(index=False))


# ===========================================================================
# 5. EXPORTAR
# ===========================================================================
# CSV en UTF-8 con BOM para que Power BI/Excel lean bien las tildes y la ñ
for df, nombre in [(usuarios, "usuarios"), (restaurantes, "restaurantes"), (pedidos, "pedidos")]:
    df.to_csv(OUT / f"{nombre}_limpio.csv", index=False, encoding="utf-8-sig", date_format="%Y-%m-%d")

log_df = pd.DataFrame(log)
log_df.to_csv(OUT / "log_calidad_datos.csv", index=False, encoding="utf-8-sig")

# Excel solo con las 3 tablas del dashboard (fechas como fecha pura, sin hora)
with pd.ExcelWriter(OUT / "pideya_limpio.xlsx", engine="openpyxl",
                    date_format="YYYY-MM-DD", datetime_format="YYYY-MM-DD") as xw:
    usuarios.assign(fecha_registro=usuarios["fecha_registro"].dt.date).to_excel(xw, sheet_name="Usuarios", index=False)
    restaurantes.to_excel(xw, sheet_name="Restaurantes", index=False)
    pedidos.assign(fecha_pedido=pedidos["fecha_pedido"].dt.date).to_excel(xw, sheet_name="Pedidos", index=False)

# Log y resumen van aparte (para el informe, no para Power BI)
with pd.ExcelWriter(OUT / "log_calidad_datos.xlsx", engine="openpyxl") as xw:
    log_df.to_excel(xw, sheet_name="Log_calidad", index=False)
    resumen.to_excel(xw, sheet_name="Resumen", index=False)

print(f"\nListo. Archivos en: {OUT}")
