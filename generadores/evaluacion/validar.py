"""Comprueba que el corpus de evaluación y sus oráculos sean coherentes. No regenera nada.

Los errores son códigos con el identificador afectado, sin valores ni texto de documentos.
"""

from __future__ import annotations

import argparse
import csv
from collections import Counter
from pathlib import Path

from .comun import (
    CAMPOS,
    PATRON_ID,
    PISTAS_PROHIBIDAS,
    cuit_valido,
    leer_json,
    normalizar,
    parrafos_docx,
    representacion,
)

FAMILIAS = {"BASE_A", "BASE_B", "MED_C", "MED_D", "ARG_A"}
ATRIBUCION_ARG_A = (
    "CC BY-NC-SA 4.0",
    "https://creativecommons.org/licenses/by-nc-sa/4.0/",
    "https://repositorio.cpau.org/items/show/823",
    "Consejo Profesional de Arquitectura y Urbanismo",
)


def leer(root: Path) -> dict:
    carpeta = root / "oraculos"
    with (carpeta / "inventario.csv").open(newline="", encoding="utf-8") as stream:
        inventario = list(csv.DictReader(stream))
    datos = {n: leer_json(carpeta / f"{n}.json") for n in ("registros", "campos", "fragmentos", "consultas")}
    datos["inventario"] = inventario
    return datos


def validar_estructura(datos: dict) -> list[str]:
    errores = []
    registros = datos["registros"]["registros"]
    byid = {r["doc_id"]: r for r in registros}
    if len(byid) != len(registros):
        errores.append("ID_DUPLICADO")
    if Counter(r["grupo"] for r in registros) != Counter({"medicion": 50, "ajuste": 15}):
        errores.append("PARTICIONES")
    for tipo in CAMPOS:
        for grupo, cantidad in (("medicion", 10), ("ajuste", 3)):
            if sum(r["tipo"] == tipo and r["grupo"] == grupo for r in registros) != cantidad:
                errores.append(f"RECUENTO_TIPO:{tipo}_{grupo}")
    adversariales = [r for r in registros if r["prompt_injection"]]
    if len(adversariales) != 3 or any(r["grupo"] != "medicion" for r in adversariales):
        errores.append("ADVERSARIALES")

    sujetos = {"medicion": set(), "ajuste": set()}
    for r in registros:
        ident = r["doc_id"]
        familia = r["plantilla_id"].split("-", 1)[-1]
        if familia not in FAMILIAS or (r["grupo"] == "ajuste" and familia in {"MED_C", "MED_D", "ARG_A"}):
            errores.append(f"PLANTILLA:{ident}")
        presentes, ausentes, na = set(r["valores"]), set(r["ausentes"]), set(r["no_aplicables"])
        if presentes & ausentes or presentes | ausentes | na != set(CAMPOS[r["tipo"]]):
            errores.append(f"ESQUEMA_CAMPOS:{ident}")
        for campo, valor in r["valores"].items():
            if campo.startswith("cuit") and not cuit_valido(valor):
                errores.append(f"CUIT_DIGITO:{ident}")
            if campo.startswith("cuit") or campo in {"contraparte", "proveedor", "emisor"}:
                sujetos[r["grupo"]].add(valor)
        if r["prompt_injection"]:
            adv = r.get("adversarial") or {}
            esperado = dict(r["valores"], clasificacion=r["clasificacion_esperada"])
            objetivos = adv.get("objetivos") or {}
            # El ataque pide algo que contradice el oráculo, sin pistas que lo neutralicen.
            if not objetivos or any(esperado.get(k) == v for k, v in objetivos.items()):
                errores.append(f"ADVERSARIAL_NO_CONTRASTABLE:{ident}")
            if any(p in adv.get("texto", "") for p in PISTAS_PROHIBIDAS):
                errores.append(f"ADVERSARIAL_NEUTRALIZADO:{ident}")
        if "ambiguo_entre_categorias" in r.get("casos_dificiles", []) and not (
            r.get("ambiguedad") or {}
        ).get("justificacion"):
            errores.append(f"AMBIGUO_SIN_JUSTIFICACION:{ident}")
    if sujetos["medicion"] & sujetos["ajuste"]:
        errores.append("SUJETOS_COMPARTIDOS_ENTRE_PARTICIONES")
    errores.extend(validar_tramites(datos["registros"]["contextos"], byid))

    for fila in datos["inventario"]:
        r = byid.get(fila["doc_id"])
        if (
            r is None
            or any(fila[k] != r[k] for k in ("grupo", "tipo", "formato", "plantilla_id"))
            or (fila["prompt_injection"] == "si") != r["prompt_injection"]
        ):
            errores.append(f"INVENTARIO_DISCREPANTE:{fila['doc_id']}")
        if PATRON_ID.search(fila["ruta"]):
            errores.append(f"ID_EN_NOMBRE_ARCHIVO:{fila['doc_id']}")
    if {f["doc_id"] for f in datos["inventario"]} != set(byid):
        errores.append("INVENTARIO_INCOMPLETO")

    vistos = set()
    for fila in datos["campos"]["campos"]:
        ident, campo = fila["doc_id"], fila["campo"]
        r = byid.get(ident)
        vistos.add((ident, campo))
        if r is None or campo not in CAMPOS[r["tipo"]]:
            errores.append(f"CAMPO_DESCONOCIDO:{ident}")
            continue
        presente = campo in r["valores"]
        if fila["presente"] != presente or fila["valor_esperado"] != r["valores"].get(campo):
            errores.append(f"CAMPO_DISCREPANTE:{ident}")
        familia = r["plantilla_id"].split("-", 1)[-1]
        if presente and (
            not fila["localizador"]
            or fila["representacion"]
            != representacion(campo, r["valores"][campo], familia, r["valores"].get("moneda"))
        ):
            errores.append(f"CAMPO_SIN_LOCALIZADOR:{ident}")
    if vistos != {(r["doc_id"], c) for r in registros for c in CAMPOS[r["tipo"]]}:
        errores.append("CAMPOS_INCOMPLETOS")

    fragmentos = {f["fragmento_id"]: f for f in datos["fragmentos"]["fragmentos"]}
    if len(fragmentos) != len(datos["fragmentos"]["fragmentos"]):
        errores.append("FRAGMENTO_DUPLICADO")
    for f in fragmentos.values():
        if f["doc_id"] not in byid or not f["texto"] or not f["localizador"]:
            errores.append(f"FRAGMENTO_INVALIDO:{f['fragmento_id']}")

    consultas = datos["consultas"]["consultas"]
    if sorted(q["pregunta_id"] for q in consultas) != [f"Q{i:02d}" for i in range(1, 31)]:
        errores.append("CONSULTAS_RECUENTO")
    if Counter(q["categoria"] for q in consultas) != Counter(
        {"directa": 10, "condiciones": 8, "multidocumento": 6, "sin_evidencia": 6}
    ):
        errores.append("CONSULTAS_DISTRIBUCION")
    for q in consultas:
        ident, contexto, refs = q["pregunta_id"], q["documentos_autorizados"], q["fragmentos_relevantes"]
        if not contexto or any(byid.get(d, {}).get("grupo") != "medicion" for d in contexto):
            errores.append(f"CONSULTA_CONTEXTO:{ident}")
        if PATRON_ID.search(q["pregunta"]):
            errores.append(f"CONSULTA_CON_ID:{ident}")
        if q["categoria"] == "sin_evidencia":
            if refs or q["respuesta_referencia"] is not None or not q["motivo_abstencion"]:
                errores.append(f"ABSTENCION_INVALIDA:{ident}")
            continue
        if not refs or not q["afirmaciones_esperadas"]:
            errores.append(f"RESPUESTA_SIN_FUENTE:{ident}")
        citados = {fragmentos[x]["doc_id"] for x in refs if x in fragmentos}
        if any(x not in fragmentos or fragmentos[x]["doc_id"] not in contexto for x in refs):
            errores.append(f"CITA_FUERA_DE_CONTEXTO:{ident}")
        if q["categoria"] == "multidocumento" and len(citados) < 2:
            errores.append(f"MULTIDOCUMENTO_INCOMPLETO:{ident}")
    return errores


def validar_tramites(contextos: dict, byid: dict) -> list[str]:
    """Cada trámite: mismo proveedor y moneda, e importes que concilian entre documentos."""
    errores = []
    for nombre, ctx in contextos.items():
        docs = [byid.get(d) for d in ctx["docs"]]
        if len(docs) < 3 or any(r is None or r["grupo"] != "medicion" for r in docs):
            errores.append(f"TRAMITE_INCOHERENTE:{nombre}")
            continue
        cuits = {v for r in docs for k, v in r["valores"].items() if k.startswith("cuit")}
        monedas = {r["valores"]["moneda"] for r in docs}
        importes = {
            r["tipo"]: r["valores"].get("importe_total", r["valores"].get("importe"))
            for r in docs
            if r["tipo"] != "contrato"
        }
        orden = importes.get("orden_compra")
        factura = next(r for r in docs if r["tipo"] == "factura")
        if (
            len(cuits) != 1
            or len(monedas) != 1
            or ("presupuesto" in importes and importes["presupuesto"] != orden)
            or (factura["discrepancia_documentada"] is None and importes["factura"] != orden)
            or ("comprobante" in importes and importes["comprobante"] != importes["factura"])
        ):
            errores.append(f"TRAMITE_INCOHERENTE:{nombre}")
    return errores


def validar_archivos(root: Path, datos: dict) -> list[str]:
    """Archivos presentes, marca de ficticio, sin ID ni pistas, y localizadores que coinciden."""
    from pypdf import PdfReader

    errores = []
    byid = {r["doc_id"]: r for r in datos["registros"]["registros"]}
    textos: dict[str, list[str]] = {}
    for fila in datos["inventario"]:
        ident, ruta = fila["doc_id"], root / fila["ruta"]
        if not ruta.is_file():
            errores.append(f"ARCHIVO_FALTANTE:{ident}")
            continue
        if fila["formato"] == "docx":
            partes = parrafos_docx(ruta)
        else:
            partes = [p.extract_text() or "" for p in PdfReader(ruta).pages]
            if byid[ident]["captura"] and any(p.strip() for p in partes):
                errores.append(f"IMAGEN_CON_CAPA_DE_TEXTO:{ident}")
        textos[ident] = partes
        contenido = " ".join(partes)
        if fila["formato"] == "docx" and "DOCUMENTO FICTICIO" not in contenido:
            errores.append(f"MARCA_FALTANTE:{ident}")
        if PATRON_ID.search(contenido) or any(p in contenido for p in PISTAS_PROHIBIDAS):
            errores.append(f"PISTA_EN_CONTENIDO:{ident}")
        if byid[ident]["plantilla_id"].endswith("ARG_A") and any(
            normalizar(a) not in normalizar(contenido) for a in ATRIBUCION_ARG_A
        ):
            errores.append(f"ATRIBUCION_ARG_A:{ident}")
    for fila in datos["campos"]["campos"] + datos["fragmentos"]["fragmentos"]:
        ident = fila["doc_id"]
        if fila["localizador"] is None or byid[ident]["captura"] or ident not in textos:
            continue  # Las capturas simuladas no tienen texto que cotejar.
        if not localizador_coincide(fila, textos[ident]):
            errores.append(f"LOCALIZADOR_SIN_RESPALDO:{fila.get('fragmento_id', ident)}")
    return errores


def localizador_coincide(fila: dict, partes: list[str]) -> bool:
    loc = fila["localizador"]
    indice = loc["indice"] if loc["tipo"] == "parrafo" else loc["pagina"]
    if not 1 <= indice <= len(partes):
        return False
    texto = normalizar(partes[indice - 1])
    if "campo" in fila:
        objetivo = normalizar(fila["representacion"])
        if loc["tipo"] == "pagina":
            objetivo = normalizar(fila["etiqueta"]) + objetivo
    else:
        # Un fragmento extenso puede seguir en la página siguiente: se busca su inicio.
        objetivo = normalizar(fila["texto"])[:60]
    return objetivo in texto


def validar(root: Path, archivos: bool = True) -> list[str]:
    datos = leer(root)
    errores = validar_estructura(datos)
    if archivos and not errores:
        errores.extend(validar_archivos(root, datos))
    return errores


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--raiz", type=Path, default=Path(__file__).resolve().parents[2])
    parser.add_argument("--solo-estructura", action="store_true")
    args = parser.parse_args()
    errores = validar(args.raiz.resolve(), not args.solo_estructura)
    print(f"errores: {len(errores)}")
    for error in errores:
        print(error)
    return bool(errores)


if __name__ == "__main__":
    raise SystemExit(main())
