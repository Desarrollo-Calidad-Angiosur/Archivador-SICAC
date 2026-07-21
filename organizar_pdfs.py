import re
import shutil
import unicodedata
from pathlib import Path

import pandas as pd

# =========================================================
# CONFIGURACIÓN
# =========================================================
# RUTA_PDFS = Path(r"C:\Users\pcalidad\Documents\SISCAC")
RUTA_PDFS = Path(r"C:\Users\pcalidad\Downloads\Historias clínicas")
RUTA_EXCEL = Path(r"C:\Users\pcalidad\Documents\SISCAC\maestro.xlsx")
HOJA_EXCEL = 0
CARPETA_SALIDA = RUTA_PDFS / "ENCARPETADO"

# True = mueve los archivos
# False = copia los archivos y deja intactos los originales
MOVER_ARCHIVOS = False
SOLO_DIRECTORIO_PRINCIPAL = True

# =========================================================
# FUNCIONES AUXILIARES
# =========================================================
def quitar_tildes(texto: str) -> str:
    texto = str(texto)
    return "".join(
        c for c in unicodedata.normalize("NFD", texto)
        if unicodedata.category(c) != "Mn"
    )


def limpiar_para_nombre(texto: str) -> str:
    """
    Deja solo letras, números y guion bajo.
    Reemplaza espacios por underscore.
    Quita tildes.
    Convierte a mayúsculas.
    """
    texto = quitar_tildes(texto).upper().strip()
    texto = re.sub(r"\s+", "_", texto)
    texto = re.sub(r"[^A-Z0-9_]", "", texto)
    texto = re.sub(r"_+", "_", texto).strip("_")
    return texto


def limpiar_documento(doc) -> str:
    """Deja solo números del documento."""
    if pd.isna(doc):
        return ""
    return re.sub(r"\D", "", str(doc))


def formatear_fecha(valor_fecha) -> str:
    """
    Convierte la fecha a formato AAAAMMDD.
    """
    if pd.isna(valor_fecha):
        return "SINFECHA"

    fecha = pd.to_datetime(valor_fecha, errors="coerce", dayfirst=True)
    if pd.isna(fecha):
        return "SINFECHA"

    return fecha.strftime("%Y%m%d")


def ruta_disponible(destino: Path) -> Path:
    """
    Si ya existe el archivo, agrega _2, _3, etc.
    """
    if not destino.exists():
        return destino

    base = destino.stem
    ext = destino.suffix
    carpeta = destino.parent
    i = 2

    while True:
        nueva = carpeta / f"{base}_{i}{ext}"
        if not nueva.exists():
            return nueva
        i += 1


def extraer_partes_pdf(nombre_archivo: str):
    """
    Espera nombres tipo:
    CC_98590824_ACE1752523_HISTORIA_CLÍNICA.pdf
    """
    patron = re.compile(
        r"^(?P<tipo>[A-Za-z]+)_(?P<documento>\d+)_(?P<admision>[A-Za-z0-9]+)_(?P<resto>.+)\.pdf$",
        re.IGNORECASE
    )

    m = patron.match(nombre_archivo)
    if not m:
        return None

    return {
        "tipo_doc": limpiar_para_nombre(m.group("tipo")),
        "documento": limpiar_documento(m.group("documento")),
        "admision": limpiar_para_nombre(m.group("admision")),
        "resto": limpiar_para_nombre(m.group("resto")),
    }


# =========================================================
# LEER Y PREPARAR EXCEL
# =========================================================
def preparar_maestro_excel(ruta_excel: Path, hoja=0) -> pd.DataFrame:
    df = pd.read_excel(ruta_excel, sheet_name=hoja, dtype=str)

    columnas_originales = list(df.columns)
    columnas_norm = {
        c: limpiar_para_nombre(c).replace("_", "")
        for c in columnas_originales
    }

    col_admision = None
    col_documento = None
    col_fecha = None
    col_tipo_proc = None

    for col, norm in columnas_norm.items():
        if norm == "ADMISION":
            col_admision = col
        elif norm == "DOCUMENTO":
            col_documento = col
        elif norm in ("FECHA", "FECHAATENCION", "FECHADEATENCION"):
            col_fecha = col
        elif norm in ("TIPOPROCEDIMIENTO", "TIPO_PROCEDIMIENTO", "PROCEDIMIENTO", "TIPOSOPORTE"):
            col_tipo_proc = col

    faltantes = []
    if not col_admision:
        faltantes.append("Admision")
    if not col_documento:
        faltantes.append("Documento")
    if not col_fecha:
        faltantes.append("Fecha")
    if not col_tipo_proc:
        faltantes.append("Tipo procedimiento")

    if faltantes:
        raise ValueError(
            f"No encontré estas columnas en el Excel: {', '.join(faltantes)}"
        )

    df = df[[col_admision, col_documento, col_fecha, col_tipo_proc]].copy()
    df.columns = ["ADMISION", "DOCUMENTO", "FECHA", "TIPO_PROCEDIMIENTO"]

    df["ADMISION"] = df["ADMISION"].astype(str).map(limpiar_para_nombre)
    df["DOCUMENTO"] = df["DOCUMENTO"].map(limpiar_documento)
    df["TIPO_PROCEDIMIENTO"] = df["TIPO_PROCEDIMIENTO"].astype(str).map(limpiar_para_nombre)
    df["FECHA_FMT"] = df["FECHA"].map(formatear_fecha)

    df = df.dropna(subset=["ADMISION"])
    df = df.drop_duplicates(subset=["ADMISION"], keep="first")

    return df.set_index("ADMISION")


# =========================================================
# PROCESO PRINCIPAL
# =========================================================
def main():
    CARPETA_SALIDA.mkdir(parents=True, exist_ok=True)

    maestro = preparar_maestro_excel(RUTA_EXCEL, HOJA_EXCEL)

    if SOLO_DIRECTORIO_PRINCIPAL:
        pdfs = [p for p in RUTA_PDFS.iterdir() if p.is_file() and p.suffix.lower() == ".pdf"]
    else:
        pdfs = list(RUTA_PDFS.rglob("*.pdf"))

    log = []

    for pdf in pdfs:
        if pdf.parent == CARPETA_SALIDA:
            continue

        partes = extraer_partes_pdf(pdf.name)
        if not partes:
            log.append({
                "archivo_original": pdf.name,
                "estado": "NO_PROCESADO",
                "motivo": "Nombre no cumple patrón esperado",
                "carpeta_destino": "",
                "archivo_nuevo": ""
            })
            continue

        tipo_doc = partes["tipo_doc"]
        documento = partes["documento"]
        admision = partes["admision"]

        carpeta_paciente = f"{tipo_doc}{documento}"
        ruta_carpeta_paciente = CARPETA_SALIDA / carpeta_paciente
        ruta_carpeta_paciente.mkdir(parents=True, exist_ok=True)

        if admision not in maestro.index:
            log.append({
                "archivo_original": pdf.name,
                "estado": "NO_PROCESADO",
                "motivo": f"No se encontró la admisión {admision} en el Excel",
                "carpeta_destino": str(ruta_carpeta_paciente),
                "archivo_nuevo": ""
            })
            continue

        fila = maestro.loc[admision]

        documento_excel = limpiar_documento(fila["DOCUMENTO"])
        fecha_fmt = fila["FECHA_FMT"]
        tipo_procedimiento = fila["TIPO_PROCEDIMIENTO"]

        if documento_excel and documento_excel != documento:
            log.append({
                "archivo_original": pdf.name,
                "estado": "REVISAR",
                "motivo": f"Documento del Excel ({documento_excel}) no coincide con PDF ({documento})",
                "carpeta_destino": str(ruta_carpeta_paciente),
                "archivo_nuevo": ""
            })


        nombre_nuevo = f"{carpeta_paciente}_{tipo_procedimiento}_{fecha_fmt}.pdf"
        nombre_nuevo = limpiar_para_nombre(nombre_nuevo[:-4]) + ".pdf"

        destino_final = ruta_disponible(ruta_carpeta_paciente / nombre_nuevo)

        try:
            if MOVER_ARCHIVOS:
                shutil.move(str(pdf), str(destino_final))
                accion = "MOVIDO"
            else:
                shutil.copy2(str(pdf), str(destino_final))
                accion = "COPIADO"

            log.append({
                "archivo_original": pdf.name,
                "estado": accion,
                "motivo": "",
                "carpeta_destino": str(ruta_carpeta_paciente),
                "archivo_nuevo": destino_final.name
            })

        except Exception as e:
            log.append({
                "archivo_original": pdf.name,
                "estado": "ERROR",
                "motivo": str(e),
                "carpeta_destino": str(ruta_carpeta_paciente),
                "archivo_nuevo": destino_final.name
            })

    # Guardar log
    df_log = pd.DataFrame(log)
    ruta_log = CARPETA_SALIDA / "log_procesamiento.xlsx"
    df_log.to_excel(ruta_log, index=False)

    print("Proceso terminado.")
    print(f"Log generado en: {ruta_log}")


if __name__ == "__main__":
    main()
    input("Presione Enter para salir...")