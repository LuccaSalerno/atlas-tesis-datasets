"""Constantes y utilidades compartidas por la generación y la validación."""

from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any

MARCA = "DOCUMENTO FICTICIO - SIN VALIDEZ COMERCIAL NI FISCAL"
# REV-E90-08: el ID de inventario no puede aparecer en el contenido ni en metadatos.
PATRON_ID = re.compile(r"E90-[AM]-(?:CTR|PRE|OC|FAC|COM)-\d{2}")
# REV-E90-02/10: pistas de construcción retiradas en v2.
PISTAS_PROHIBIDAS = (
    "Presenta una oferta económica",
    "Instruye una adquisición ficticia",
    "Representa un cargo por servicios",
    "Deja constancia de recepción de prestaciones",
    "Escenario académico candidato",
    "no a una instrucción autorizada",
    "Factura sintética no fiscal",
)
# P4: perfil de formato por familia de plantilla.
PERFILES = {
    "BASE_A": {
        "fecha": "dd/mm/aaaa",
        "importe": "plano",
        "moneda": "codigo",
        "cuit": "guiones",
    },
    "BASE_B": {
        "fecha": "dd-mm-aaaa",
        "importe": "simbolo",
        "moneda": "simbolo",
        "cuit": "sin_guiones",
    },
    "MED_C": {
        "fecha": "dd/mm/aa",
        "importe": "plano",
        "moneda": "codigo",
        "cuit": "guiones",
    },
    "MED_D": {
        "fecha": "letras",
        "importe": "plano",
        "moneda": "nombre",
        "cuit": "sin_guiones",
    },
    "ARG_A": {
        "fecha": "dd/mm/aaaa",
        "importe": "plano",
        "moneda": "nombre",
        "cuit": "guiones",
    },
}
SIMBOLOS = {"ARS": "$", "USD": "U$S", "EUR": "€"}
NOMBRES_MONEDA = {
    "ARS": "pesos argentinos",
    "USD": "dólares estadounidenses",
    "EUR": "euros",
}
MESES = (
    "enero",
    "febrero",
    "marzo",
    "abril",
    "mayo",
    "junio",
    "julio",
    "agosto",
    "septiembre",
    "octubre",
    "noviembre",
    "diciembre",
)
CAMPOS = {
    "contrato": (
        "contraparte",
        "cuit_contraparte",
        "fecha_celebracion",
        "vigencia_desde",
        "vigencia_hasta",
        "monto_total",
        "moneda",
    ),
    "presupuesto": (
        "numero",
        "fecha",
        "proveedor",
        "cuit_proveedor",
        "importe_total",
        "moneda",
    ),
    "orden_compra": (
        "numero",
        "fecha",
        "proveedor",
        "cuit_proveedor",
        "importe_total",
        "moneda",
    ),
    "factura": (
        "numero",
        "fecha_emision",
        "emisor",
        "cuit_emisor",
        "importe_total",
        "moneda",
    ),
    "comprobante": (
        "numero",
        "fecha",
        "contraparte",
        "cuit_contraparte",
        "importe",
        "moneda",
    ),
}
ETIQUETAS = {
    "numero": "Número documental",
    "contraparte": "Contraparte",
    "proveedor": "Proveedor",
    "emisor": "Emisor",
    "cuit_contraparte": "CUIT de la contraparte",
    "cuit_proveedor": "CUIT del proveedor",
    "cuit_emisor": "CUIT del emisor",
    "fecha": "Fecha",
    "fecha_emision": "Fecha de emisión",
    "fecha_celebracion": "Fecha de celebración",
    "vigencia_desde": "Vigencia desde",
    "vigencia_hasta": "Vigencia hasta",
    "monto_total": "Monto total",
    "importe_total": "Importe total",
    "importe": "Importe",
    "moneda": "Moneda",
}


def leer_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def escribir_json(path: Path, data: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(data, ensure_ascii=False, indent=2, sort_keys=True) + "\n",
        encoding="utf-8",
    )


def digito_cuit(base: str) -> int | None:
    if not re.fullmatch(r"\d{10}", base):
        return None
    residue = (
        11
        - sum(
            int(c) * k
            for c, k in zip(base, (5, 4, 3, 2, 7, 6, 5, 4, 3, 2), strict=True)
        )
        % 11
    )
    return 0 if residue == 11 else None if residue == 10 else residue


def cuit_valido(value: str) -> bool:
    if not re.fullmatch(r"30-\d{8}-\d", value):
        return False
    digits = value.replace("-", "")
    return digito_cuit(digits[:10]) == int(digits[-1])


def normalizar(value: str) -> str:
    return re.sub(r"\s+", "", value).casefold().replace("–", "-").replace("—", "-")


def importe_es(value: str) -> str:
    from decimal import Decimal

    return (
        f"{Decimal(value):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    )


def representacion(
    campo: str, value: str, familia: str = "BASE_A", moneda: str | None = None
) -> str:
    """P4: texto impreso del valor normalizado según el perfil de la familia."""
    perfil = PERFILES[familia]
    if campo.startswith(("fecha", "vigencia")):
        year, month, day = value.split("-")
        if perfil["fecha"] == "dd-mm-aaaa":
            return f"{day}-{month}-{year}"
        if perfil["fecha"] == "dd/mm/aa":
            return f"{day}/{month}/{year[2:]}"
        if perfil["fecha"] == "letras":
            return f"{int(day)} de {MESES[int(month) - 1]} de {year}"
        return f"{day}/{month}/{year}"
    if campo in {"importe", "importe_total", "monto_total"}:
        if perfil["importe"] == "simbolo":
            return SIMBOLOS[moneda or "ARS"] + " " + importe_es(value)
        return importe_es(value)
    if campo == "moneda":
        if perfil["moneda"] == "simbolo":
            return SIMBOLOS[value]
        if perfil["moneda"] == "nombre":
            return NOMBRES_MONEDA[value]
        return value
    if campo.startswith("cuit") and perfil["cuit"] == "sin_guiones":
        return value.replace("-", "")
    return value


_UNIDADES = (
    "",
    "uno",
    "dos",
    "tres",
    "cuatro",
    "cinco",
    "seis",
    "siete",
    "ocho",
    "nueve",
    "diez",
    "once",
    "doce",
    "trece",
    "catorce",
    "quince",
    "dieciséis",
    "diecisiete",
    "dieciocho",
    "diecinueve",
    "veinte",
    "veintiuno",
    "veintidós",
    "veintitrés",
    "veinticuatro",
    "veinticinco",
    "veintiséis",
    "veintisiete",
    "veintiocho",
    "veintinueve",
)
_DECENAS = (
    "",
    "",
    "",
    "treinta",
    "cuarenta",
    "cincuenta",
    "sesenta",
    "setenta",
    "ochenta",
    "noventa",
)
_CENTENAS = (
    "",
    "ciento",
    "doscientos",
    "trescientos",
    "cuatrocientos",
    "quinientos",
    "seiscientos",
    "setecientos",
    "ochocientos",
    "novecientos",
)


def _hasta_999(n: int) -> str:
    if n == 100:
        return "cien"
    c, resto = divmod(n, 100)
    partes = [_CENTENAS[c]] if c else []
    if resto < 30:
        partes.append(_UNIDADES[resto])
    else:
        d, u = divmod(resto, 10)
        partes.append(_DECENAS[d] + (" y " + _UNIDADES[u] if u else ""))
    return " ".join(p for p in partes if p)


def _apocope(texto: str) -> str:
    if texto.endswith("veintiuno"):
        return texto[:-9] + "veintiún"
    return texto[:-3] + "un" if texto.endswith("uno") else texto


def entero_en_letras(n: int) -> str:
    if n == 0:
        return "cero"
    millones, resto = divmod(n, 1_000_000)
    miles, unidades = divmod(resto, 1000)
    partes = []
    if millones:
        partes.append(
            "un millón"
            if millones == 1
            else _apocope(_hasta_999(millones)) + " millones"
        )
    if miles:
        partes.append("mil" if miles == 1 else _apocope(_hasta_999(miles)) + " mil")
    if unidades:
        partes.append(_hasta_999(unidades))
    return " ".join(partes)


def importe_en_letras(value: str) -> str:
    """P4: complemento en letras de un importe, al estilo «… con 05/100»."""
    from decimal import Decimal

    total = Decimal(value)
    entero = int(total)
    centavos = int((total - entero) * 100)
    return f"{entero_en_letras(entero)} con {centavos:02d}/100"


def parrafos_docx(path: Path) -> list[str]:
    """RF-11/IA-M04: índice 1-based de todos los párrafos de word/document.xml."""
    from xml.etree import ElementTree as ET
    from zipfile import ZipFile

    namespace = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    with ZipFile(path) as archive:
        root = ET.fromstring(archive.read("word/document.xml"))
        return [
            "".join(n.text or "" for n in p.iter(namespace + "t"))
            for p in root.iter(namespace + "p")
        ]
