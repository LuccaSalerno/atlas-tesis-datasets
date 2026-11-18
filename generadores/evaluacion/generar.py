"""Paso 2: documentos, inventario, campos y fragmentos a partir de `oraculos/registros.json`.

Arma un DOCX por registro, lo convierte a PDF con LibreOffice y, para los escaneados y
fotografiados, simula la captura con pdftoppm y Pillow. Los DOCX y PDF intermedios quedan
en una carpeta temporal.
"""

from __future__ import annotations

import argparse
import csv
import os
import random
import shutil
import subprocess
import tempfile
import time
from datetime import UTC, datetime
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

from .comun import (
    CAMPOS,
    MARCA,
    NOMBRES_MONEDA,
    escribir_json,
    importe_en_letras,
    importe_es,
    leer_json,
    normalizar,
    parrafos_docx,
    representacion,
)

PIE = "Documento ficticio de evaluación académica; sujetos, importes y operaciones inventados."
ESTILOS = {
    "BASE_A": ("Liberation Sans", 11),
    "BASE_B": ("Liberation Serif", 11.5),
    "MED_C": ("DejaVu Sans", 9.5),
    "MED_D": ("Liberation Serif", 11),
    "ARG_A": ("Liberation Serif", 11),
}
TITULOS = {
    "contrato": {
        "BASE_A": "Contrato de prestación de servicios",
        "BASE_B": "Acuerdo de servicios profesionales",
        "MED_C": "Contrato de servicios técnicos",
        "MED_D": "Contrato de mantenimiento",
        "ARG_A": "Contrato de proyecto y dirección de obra",
    },
    "presupuesto": {
        "BASE_A": "Presupuesto",
        "BASE_B": "Cotización",
        "MED_C": "Presupuesto",
        "MED_D": "Presupuesto",
    },
    "orden_compra": {
        "BASE_A": "Orden de compra",
        "BASE_B": "Orden de compra",
        "MED_C": "Pedido de compra",
        "MED_D": "Orden de compra",
    },
    "factura": {
        "BASE_A": "Factura",
        "BASE_B": "Factura de servicios",
        "MED_C": "Factura",
        "MED_D": "Factura",
    },
    "comprobante": {
        "BASE_A": "Constancia de recepción",
        "BASE_B": "Comprobante de recepción de servicios",
        "MED_C": "Acta de recepción",
        "MED_D": "Constancia de conformidad",
    },
}
TITULOS_AMBIGUOS = {
    "E90-M-PRE-07": "Propuesta comercial",
    "E90-M-OC-07": "Pedido de servicios",
    "E90-M-FAC-08": "Liquidación de servicios",
    "E90-M-COM-08": "Constancia de conformidad",
}
ETIQUETAS_A = {
    "numero": "Número",
    "fecha": "Fecha",
    "fecha_emision": "Fecha de emisión",
    "proveedor": "Proveedor",
    "emisor": "Emisor",
    "contraparte": "Prestador",
    "cuit_proveedor": "CUIT",
    "cuit_emisor": "CUIT",
    "cuit_contraparte": "CUIT del prestador",
    "importe_total": "Importe total",
    "importe": "Importe recibido",
    "monto_total": "Precio total",
    "moneda": "Moneda",
    "fecha_celebracion": "Fecha de celebración",
    "vigencia_desde": "Inicio de vigencia",
    "vigencia_hasta": "Fin de vigencia",
}
ETIQUETAS_B = {
    "numero": "N.º:",
    "fecha": "Emitido el:",
    "fecha_emision": "Fecha de emisión:",
    "proveedor": "Razón social del proveedor:",
    "emisor": "Razón social:",
    "contraparte": "Proveedor:",
    "cuit_proveedor": "C.U.I.T. N.º:",
    "cuit_emisor": "C.U.I.T. N.º:",
    "cuit_contraparte": "C.U.I.T. N.º:",
    "fecha_celebracion": "Celebrado el:",
    "vigencia_desde": "Inicio de vigencia:",
    "vigencia_hasta": "Finalización:",
}
PARTE_C = {
    "presupuesto": "Cotiza:",
    "orden_compra": "Proveedor adjudicado:",
    "factura": "Emitido por:",
    "comprobante": "Prestador:",
}
EQUIPOS = {
    "calibración de balanzas analíticas y de precisión": "balanza",
    "mantenimiento preventivo de cabinas de seguridad biológica": "cabina",
    "verificación de termómetros y registradores de temperatura": "registrador",
    "mantenimiento de equipos de cromatografía líquida": "cromatógrafo",
    "soporte técnico de espectrofotómetros UV-visible": "espectrofotómetro",
    "validación térmica de autoclaves y estufas": "autoclave",
    "calibración de micropipetas y material volumétrico": "micropipeta",
    "revisión de instalaciones de gases de laboratorio": "línea de gas",
}
TAREAS = (
    "Relevamiento inicial e identificación de {e} N.º {k}",
    "Verificación funcional de {e} N.º {k} según procedimiento interno",
    "Ajuste, limpieza y control de componentes de {e} N.º {k}",
    "Calibración con patrón trazable de {e} N.º {k}",
    "Emisión de informe técnico de {e} N.º {k}",
    "Reemplazo de consumibles menores de {e} N.º {k}",
)


def normalizar_pdf(source: Path, dest: Path, titulo: str) -> None:
    from pypdf import PdfReader, PdfWriter

    writer = PdfWriter()
    writer.clone_document_from_reader(PdfReader(source))
    writer.add_metadata(
        {
            "/Title": titulo,
            "/Author": "ATLAS corpus sintético",
            "/CreationDate": "D:20261022000000Z",
            "/ModDate": "D:20261022000000Z",
            "/Producer": "ATLAS",
        }
    )
    dest.parent.mkdir(parents=True, exist_ok=True)
    with dest.open("wb") as stream:
        writer.write(stream)


def docx_determinista(path: Path) -> None:
    """Fijar timestamps ZIP; la reproducibilidad binaria depende también del renderer registrado."""
    with ZipFile(path) as z:
        parts = [(n, z.read(n)) for n in z.namelist()]
    with ZipFile(path, "w", compression=ZIP_DEFLATED) as z:
        for name, data in sorted(parts):
            info = ZipInfo(name, date_time=(2026, 10, 22, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            z.writestr(info, data)


class Constructor:
    """Escribe el DOCX y registra cómo localizar cada campo y fragmento."""

    def __init__(self, registro: dict, fuentes: dict):
        from docx import Document
        from docx.shared import Cm, Pt, RGBColor

        self.r = registro
        self.fuentes = fuentes
        self.familia = registro["plantilla_id"].split("-", 1)[1]
        self.v = registro["valores"]
        self.rng = random.Random(registro["semilla"] ^ 0x5EED)
        self.campos: dict[str, dict] = {}
        self.fragmentos: list[tuple[str, str]] = []
        self.Pt = Pt
        font, size = ESTILOS[self.familia]
        doc = Document()
        sec = doc.sections[0]
        sec.page_width, sec.page_height = Cm(21), Cm(29.7)
        sec.top_margin, sec.bottom_margin = Cm(1.8), Cm(1.9)
        sec.left_margin = sec.right_margin = Cm(2.2 if self.familia == "MED_D" else 2)
        for name in ("Normal", "Title", "Heading 1", "Heading 2"):
            style = doc.styles[name]
            style.font.name = font
            style.font.color.rgb = RGBColor(0, 0, 0)
            style.font.size = Pt(
                {"Normal": size, "Title": 18, "Heading 1": 12.5, "Heading 2": 11}[name]
            )
            style.paragraph_format.space_after = Pt(5)
            style.paragraph_format.line_spacing = 1.08
        cp = doc.core_properties
        cp.author = "ATLAS corpus sintético"
        cp.title = registro["archivo_id"]
        cp.created = datetime(2026, 10, 22, tzinfo=UTC)
        cp.modified = cp.created
        self.doc = doc

    # --- utilidades -------------------------------------------------------
    def rep(self, campo: str) -> str:
        return representacion(campo, self.v[campo], self.familia, self.v.get("moneda"))

    def fecha(self, iso: str) -> str:
        return representacion("fecha", iso, self.familia)

    def p(
        self,
        texto: str,
        clave: str | None = None,
        size: float | None = None,
        bold: bool = False,
    ):
        par = self.doc.add_paragraph()
        run = par.add_run(texto)
        run.bold = bold
        if size:
            run.font.size = self.Pt(size)
        if clave:
            self.fragmentos.append((clave, texto))
        return par

    def h(self, texto: str, level: int = 1) -> None:
        self.doc.add_heading(texto, level=level)

    def registrar(self, campo: str, etiqueta: str, modo: str = "contiguo") -> None:
        self.campos[campo] = {
            "etiqueta": etiqueta,
            "modo": modo,
            "representacion": self.rep(campo),
        }

    def linea(self, campo: str, etiqueta: str, sep: str = " ") -> None:
        """Párrafo «etiqueta valor»; se omite si el campo está ausente."""
        if campo not in self.v:
            return
        texto = etiqueta + sep + self.rep(campo)
        self.p(texto, "CAMPO-" + campo)
        self.registrar(campo, etiqueta)

    def tabla(
        self, filas: list[tuple[str, str]], encabezado: tuple[str, str] | None = None
    ) -> None:
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn

        filas = [(c, e) for c, e in filas if c in self.v]
        table = self.doc.add_table(rows=0, cols=2)
        table.style = "Table Grid"
        if encabezado:
            cells = table.add_row().cells
            for cell, texto in zip(cells, encabezado, strict=True):
                cell.text = texto
                cell.paragraphs[0].runs[0].bold = True
        for campo, etiqueta in filas:
            cells = table.add_row().cells
            cells[0].text = etiqueta
            cells[1].text = self.rep(campo)
            self.registrar(campo, etiqueta, "tabla")
            self.fragmentos.append(("CAMPO-" + campo, self.rep(campo)))
        borders = OxmlElement("w:tblBorders")
        for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
            item = OxmlElement("w:" + edge)
            for key, val in {"val": "single", "sz": "4", "color": "BFBFBF"}.items():
                item.set(qn("w:" + key), val)
            borders.append(item)
        table._tbl.tblPr.append(borders)

    def salto(self) -> None:
        self.doc.add_page_break()

    def importe_campo(self) -> str:
        return next(
            c for c in ("monto_total", "importe_total", "importe") if c in self.v
        )

    # --- secciones compartidas -------------------------------------------
    def encabezado(self) -> None:
        self.p(MARCA, size=7.5)
        ident = self.r["doc_id"]
        titulo = TITULOS_AMBIGUOS.get(ident) or TITULOS[self.r["tipo"]][self.familia]
        if self.familia != "MED_D" or self.r["tipo"] == "contrato":
            self.doc.add_paragraph(titulo, style="Title")
        self.titulo = titulo

    def condiciones(self) -> None:
        c = self.r["condiciones"]
        tipo = self.r["tipo"]
        if tipo == "presupuesto":
            self.p(
                f"Validez de la oferta: {c['validez_dias']} días corridos.",
                "COND-validez",
            )
            self.p(f"Forma de pago: {c['forma_pago']}.", "COND-forma_pago")
            self.p(
                f"Plazo de ejecución: {c['plazo_ejecucion_dias_habiles']} días hábiles desde la aceptación.",
                "COND-plazo",
            )
        elif tipo == "orden_compra":
            self.p(f"Lugar de entrega: {c['lugar_entrega']}.", "COND-lugar")
            self.p(
                f"Fecha de entrega comprometida: {self.fecha(c['fecha_entrega'])}.",
                "COND-fecha_entrega",
            )
            self.p(f"Condición de pago: {c['condicion_pago']}.", "COND-pago")
        elif tipo == "factura":
            self.p(
                f"Vencimiento del pago: {self.fecha(c['vencimiento'])}.",
                "COND-vencimiento",
            )
            self.p(f"Forma de pago: {c['forma_pago']}.", "COND-forma_pago")
            self.p(f"Período facturado: {c['periodo']}.", "COND-periodo")
        elif tipo == "comprobante":
            self.p(
                f"Los servicios se recibieron {c['estado_recepcion']}.", "COND-estado"
            )
            self.p(f"Registró la recepción: {c['responsable']}.", "COND-responsable")

    def referencias(self) -> None:
        for ref in self.r["referencias"]:
            if ref["tipo"] == "contrato":
                texto = (
                    "Emitida en el marco del contrato de servicios celebrado el "
                    + self.fecha(ref["fecha_celebracion"])
                    + "."
                )
            else:
                nombre = {
                    "presupuesto": "presupuesto",
                    "orden_compra": "orden de compra",
                    "factura": "factura",
                }[ref["tipo"]]
                texto = (
                    f"Corresponde a la {nombre} {ref['numero']}."
                    if ref["tipo"] != "presupuesto"
                    else f"Según presupuesto {ref['numero']}."
                )
            self.p(texto, "REF")

    def adversarial(self, size: float | None = None) -> None:
        adv = self.r["adversarial"]
        if adv:
            self.p(adv["texto"], "ADVERSARIAL", size=size)

    def pie(self) -> None:
        self.p(PIE, size=7.5)

    def items(self, neto: Decimal, cantidad: int) -> list[tuple[str, Decimal]]:
        equipo = EQUIPOS[self.r["servicio"]]
        pesos = [self.rng.uniform(0.6, 1.6) for _ in range(cantidad)]
        result = []
        acumulado = Decimal(0)
        for k, w in enumerate(pesos, 1):
            texto = self.rng.choice(TAREAS).format(e=equipo, k=k)
            if k < cantidad:
                monto = (neto * Decimal(w / sum(pesos))).quantize(Decimal(".01"))
                acumulado += monto
            else:
                monto = neto - acumulado
            result.append((texto, monto))
        return result

    def montos(self) -> tuple[Decimal, Decimal, Decimal]:
        """REV-E90-16: el IVA es el 21 % del neto redondeado a centavos; si neto + IVA no
        alcanza el total de referencia, la diferencia se imprime como ajuste por redondeo."""
        total = Decimal(self.v[self.importe_campo()])
        if self.r["tipo"] == "comprobante":
            return total, Decimal(0), total
        neto = (total / Decimal("1.21")).quantize(Decimal(".01"), ROUND_HALF_UP)
        iva = (neto * Decimal("0.21")).quantize(Decimal(".01"), ROUND_HALF_UP)
        return neto, iva, total

    # --- familias no contractuales ---------------------------------------
    def base_a(self) -> None:
        tipo = self.r["tipo"]
        parte = next(
            c for c in ("proveedor", "emisor", "contraparte") if c in CAMPOS[tipo]
        )
        cuit = "cuit_" + parte
        self.p(
            f"Destinatario: {self.r['comprador']}."
            if tipo in {"presupuesto", "factura"}
            else f"Emisor: {self.r['comprador']}."
        )
        filas = [("numero", ETIQUETAS_A["numero"])]
        filas.append(
            (
                "fecha_emision" if tipo == "factura" else "fecha",
                ETIQUETAS_A["fecha_emision" if tipo == "factura" else "fecha"],
            )
        )
        filas += [(parte, ETIQUETAS_A[parte]), (cuit, ETIQUETAS_A[cuit])]
        filas += [
            (self.importe_campo(), ETIQUETAS_A[self.importe_campo()]),
            ("moneda", ETIQUETAS_A["moneda"]),
        ]
        self.tabla(filas, ("Dato", "Detalle"))
        self.h("Descripción")
        self.p("Servicio: " + self.r["servicio"] + ".", "OBJETO")
        self.referencias()
        if self.r["multipagina"]:
            self.salto()
        self.h("Condiciones")
        self.condiciones()
        self.adversarial()
        self.pie()

    def base_b(self) -> None:
        tipo = self.r["tipo"]
        parte = next(
            c for c in ("proveedor", "emisor", "contraparte") if c in CAMPOS[tipo]
        )
        self.linea("numero", ETIQUETAS_B["numero"])
        campo_fecha = "fecha_emision" if tipo == "factura" else "fecha"
        self.linea(campo_fecha, ETIQUETAS_B[campo_fecha])
        self.h("Datos de la operación")
        self.linea(parte, ETIQUETAS_B[parte])
        self.linea("cuit_" + parte, ETIQUETAS_B["cuit_" + parte])
        self.p(
            ("Cliente: " if tipo in {"presupuesto", "factura"} else "Emisor: ")
            + self.r["comprador"]
            + "."
        )
        self.h("Detalle del servicio")
        self.p(
            "Se contempla " + self.r["servicio"] + " en las instalaciones del cliente.",
            "OBJETO",
        )
        campo = self.importe_campo()
        texto = "Total: " + self.rep(campo)
        self.p(texto, "CAMPO-" + campo)
        self.registrar(campo, "Total:")
        self.registrar("moneda", "Total:")
        self.referencias()
        if self.r["multipagina"]:
            self.salto()
        self.h("Condiciones comerciales")
        self.condiciones()
        self.adversarial()
        self.pie()

    def med_c(self) -> None:
        tipo = self.r["tipo"]
        parte = next(
            c for c in ("proveedor", "emisor", "contraparte") if c in CAMPOS[tipo]
        )
        self.linea(parte, PARTE_C[tipo])
        self.linea("cuit_" + parte, "CUIT")
        self.p(
            ("Cliente: " if tipo in {"presupuesto", "factura"} else "Emite: ")
            + self.r["comprador"]
        )
        self.linea("numero", self.titulo + " Nro.")
        self.linea("fecha_emision" if tipo == "factura" else "fecha", "Fecha")
        self.p("Concepto: " + self.r["servicio"] + ".", "OBJETO")
        self.referencias()
        neto, iva, _total = self.montos()
        cantidad = 30 if self.r["multipagina"] else self.rng.randint(4, 7)
        from docx.shared import Cm

        table = self.doc.add_table(rows=1, cols=2)
        table.style = "Table Grid"
        table.rows[0].cells[0].text = "Detalle"
        table.rows[0].cells[1].text = "Importe"
        for texto, monto in self.items(neto, cantidad):
            cells = table.add_row().cells
            cells[0].text = texto
            cells[1].text = importe_es(str(monto))
        for row in table.rows:
            row.cells[0].width, row.cells[1].width = Cm(13), Cm(3.5)
        moneda = self.v["moneda"]
        campo = self.importe_campo()
        if tipo == "comprobante":
            etiqueta = "Total recibido"
        else:
            self.p(f"Neto gravado {moneda} {importe_es(str(neto))}")
            self.p(f"IVA 21 % {moneda} {importe_es(str(iva))}")
            ajuste = _total - neto - iva
            if ajuste:
                signo = "+" if ajuste > 0 else "−"
                self.p(f"Ajuste por redondeo {signo} {moneda} {importe_es(str(abs(ajuste)))}")
            etiqueta = "Total"
        self.p(f"{etiqueta} {moneda} {self.rep(campo)}", "CAMPO-" + campo, bold=True)
        self.campos[campo] = {
            "etiqueta": f"{etiqueta} {moneda}",
            "modo": "contiguo",
            "representacion": self.rep(campo),
        }
        self.campos["moneda"] = {
            "etiqueta": etiqueta,
            "modo": "contiguo",
            "representacion": self.rep("moneda"),
        }
        self.adversarial(size=6.5)
        self.h("Condiciones", level=2)
        self.condiciones()
        self.pie()

    def med_d(self) -> None:
        tipo = self.r["tipo"]
        ident = self.r["doc_id"]
        parte = next(
            c for c in ("proveedor", "emisor", "contraparte") if c in CAMPOS[tipo]
        )
        emite_proveedor = tipo in {"presupuesto", "factura"}
        campo_fecha = "fecha_emision" if tipo == "factura" else "fecha"
        if campo_fecha in self.v:
            self.linea(campo_fecha, "Buenos Aires,")
        else:
            self.p("Buenos Aires")
        self.linea("numero", "Ref.: " + self.titulo + " N.º")
        if emite_proveedor:
            self.p("Señores " + self.r["comprador"])
        else:
            self.linea(parte, "Señores")
            self.linea("cuit_" + parte, "CUIT")
        self.p("De nuestra consideración:")
        servicio = self.r["servicio"]
        cuerpo = {
            "presupuesto": f"Nos dirigimos a ustedes para presentar nuestra propuesta por {servicio}.",
            "orden_compra": f"Por la presente les encomendamos {servicio}, en las condiciones que se detallan.",
            "factura": f"Por los servicios de {servicio} prestados durante el período indicado, facturamos según el detalle siguiente.",
            "comprobante": f"Por la presente dejamos constancia de haber recibido los servicios de {servicio}.",
        }[tipo]
        if ident == "E90-M-OC-07":
            cuerpo = (
                f"Por la presente les solicitamos la provisión de {servicio}, conforme a su cotización, "
                "que aceptamos. Reproducimos a continuación los ítems y precios cotizados."
            )
        elif ident == "E90-M-FAC-08":
            cuerpo = (
                f"Habiéndose completado y recibido de conformidad la prestación de {servicio}, "
                "liquidamos los servicios según el detalle siguiente."
            )
        elif ident == "E90-M-COM-08":
            ref = self.r["referencias"][0]
            cuerpo = (
                f"Por la presente dejamos constancia de haber recibido de conformidad los servicios de {servicio}, "
                f"correspondientes a su factura {ref['numero']}."
            )
        self.p(cuerpo, "OBJETO")
        if ident != "E90-M-COM-08":
            self.referencias()
        neto, _, _ = self.montos()
        if ident == "E90-M-OC-07":
            for texto, monto in self.items(neto, 3):
                self.p(f"— {texto}: {importe_es(str(monto))} más IVA.")
            self.p("Validez de la cotización aceptada: 15 días corridos.")
        if self.r["multipagina"]:
            self.h("Detalle de los trabajos", level=2)
            for texto, _ in self.items(neto, 26):
                self.p("— " + texto + ".")
        campo = self.importe_campo()
        rep = self.rep(campo)
        moneda = NOMBRES_MONEDA[self.v["moneda"]]
        verbo = {
            "presupuesto": "El importe total de la propuesta es de",
            "orden_compra": "El importe total de esta orden es de",
            "factura": "El importe total a abonar es de",
            "comprobante": "Lo recibido corresponde a un importe total de",
        }[tipo]
        if ident == "E90-M-FAC-08":
            verbo = "El importe total liquidado es de"
        elif ident == "E90-M-OC-07":
            verbo = "El importe total del pedido es de"
        texto = f"{verbo} {rep} {moneda} ({importe_en_letras(self.v[campo])}), IVA incluido."
        if tipo == "comprobante":
            texto = f"{verbo} {rep} {moneda} ({importe_en_letras(self.v[campo])})."
        self.p(texto, "CAMPO-" + campo)
        self.campos[campo] = {
            "etiqueta": verbo,
            "modo": "contiguo",
            "representacion": rep,
        }
        self.campos["moneda"] = {
            "etiqueta": f"{verbo} {rep}",
            "modo": "contiguo",
            "representacion": self.rep("moneda"),
        }
        self.condiciones()
        if ident == "E90-M-PRE-07":
            self.h("Aceptación del cliente", level=2)
            self.p(
                "Aceptamos la presente propuesta en todos sus términos. Al recibir la aceptación firmada, "
                "el proveedor programará los trabajos.",
                "ACEPTACION",
            )
            self.p(
                "Firma: ____________________   Aclaración: ____________________   Fecha: __________"
            )
        if self.r["adversarial"]:
            self.p("Observaciones: " + self.r["adversarial"]["texto"], "ADVERSARIAL")
        self.p("Atentamente,")
        if emite_proveedor:
            self.linea(parte, "Por")
            self.linea("cuit_" + parte, "CUIT")
        else:
            self.p(
                "Por "
                + self.r["comprador"]
                + " — "
                + self.r["condiciones"].get("responsable", "Área de Compras")
            )
        self.pie()

    # --- contratos --------------------------------------------------------
    def contrato(self) -> None:
        fam = self.familia
        fuente = self.fuentes[fam]
        if fam == "ARG_A":
            self.contrato_arg_a(fuente)
            return
        comprador = self.r["comprador"]
        if fam == "BASE_A":
            self.p(
                f"Entre {comprador}, en adelante el Cliente, y el Prestador identificado en las condiciones "
                "particulares, se celebra el presente contrato.",
            )
            self.h("Condiciones particulares")
            self.tabla(
                [
                    ("contraparte", ETIQUETAS_A["contraparte"]),
                    ("cuit_contraparte", ETIQUETAS_A["cuit_contraparte"]),
                    ("fecha_celebracion", ETIQUETAS_A["fecha_celebracion"]),
                    ("vigencia_desde", ETIQUETAS_A["vigencia_desde"]),
                    ("vigencia_hasta", ETIQUETAS_A["vigencia_hasta"]),
                    ("monto_total", ETIQUETAS_A["monto_total"]),
                    ("moneda", ETIQUETAS_A["moneda"]),
                ]
            )
        elif fam == "BASE_B":
            self.p(f"Cliente: {comprador}.")
            for campo in (
                "contraparte",
                "cuit_contraparte",
                "fecha_celebracion",
                "vigencia_desde",
                "vigencia_hasta",
            ):
                self.linea(campo, ETIQUETAS_B[campo])
            texto = "Honorarios totales: " + self.rep("monto_total")
            self.p(texto, "CAMPO-monto_total")
            self.registrar("monto_total", "Honorarios totales:")
            self.registrar("moneda", "Honorarios totales:")
        elif fam == "MED_C":
            self.h("Declaración de trabajo")
            self.p(f"Cliente: {comprador}.")
            self.linea("contraparte", "Prestador:")
            self.linea("cuit_contraparte", "CUIT")
            self.linea("fecha_celebracion", "Fecha de firma:")
            texto = f"Vigencia: desde {self.rep('vigencia_desde')} hasta {self.rep('vigencia_hasta')}."
            self.p(texto, "CAMPO-vigencia")
            self.registrar("vigencia_desde", "Vigencia: desde")
            self.registrar("vigencia_hasta", "hasta")
            moneda = self.v["moneda"]
            self.p(f"Precio: {moneda} {self.rep('monto_total')}", "CAMPO-monto_total")
            self.registrar("monto_total", f"Precio: {moneda}")
            self.registrar("moneda", "Precio:")
        else:  # MED_D: datos dentro de la prosa
            partes = []
            if "fecha_celebracion" in self.v:
                partes.append(f"En Buenos Aires, el {self.rep('fecha_celebracion')}")
                self.registrar("fecha_celebracion", "En Buenos Aires, el")
            else:
                partes.append("En Buenos Aires")
            partes.append(
                f", entre {comprador}, en adelante el Comprador, y {self.rep('contraparte')}, "
                f"CUIT {self.rep('cuit_contraparte')}, en adelante el Proveedor, se celebra el presente "
                f"contrato de mantenimiento referido a {self.r['servicio']}."
            )
            texto = "".join(partes)
            self.p(texto, "CAMPO-partes")
            self.registrar("contraparte", "en adelante el Comprador, y")
            self.registrar("cuit_contraparte", "CUIT")
            texto = f"El contrato rige desde el {self.rep('vigencia_desde')} hasta el {self.rep('vigencia_hasta')}."
            self.p(texto, "CAMPO-vigencia")
            self.registrar("vigencia_desde", "rige desde el")
            self.registrar("vigencia_hasta", "hasta el")
            rep = self.rep("monto_total")
            moneda = NOMBRES_MONEDA[self.v["moneda"]]
            texto = f"El precio total es de {rep} {moneda} ({importe_en_letras(self.v['monto_total'])})."
            self.p(texto, "CAMPO-monto_total")
            self.registrar("monto_total", "El precio total es de")
            self.registrar("moneda", f"El precio total es de {rep}")
        if fam != "MED_D":
            self.p("Objeto: " + self.r["servicio"] + ".", "OBJETO")
        self.p(
            "Traducción y adaptación de "
            + fuente["titulo"]
            + " "
            + fuente["version"]
            + ". "
            + fuente["atribucion"]
            + " Licencia: "
            + fuente["licencia"]
            + ". Fuente: "
            + fuente["url_publicacion"],
            size=8.5,
        )
        extenso = "contrato_extenso" in self.r["casos_dificiles"]
        for n, clause in enumerate(fuente["clausulas"], 1):
            if extenso:
                self.salto()
            titulo = {
                "BASE_A": f"Cláusula {n}. {clause['titulo']}",
                "BASE_B": f"{n}. {clause['titulo']}",
                "MED_C": f"Sección {n} — {clause['titulo']}",
                "MED_D": f"{n} {clause['titulo']}",
            }[fam]
            self.h(titulo)
            self.p(clause["adaptacion_es"], f"CLAUSULA-{n}")
            if self.r["adversarial"] and n == 3:
                self.p(self.r["adversarial"]["texto"], "ADVERSARIAL")
            if extenso:
                self.h("Correspondencia con el modelo público", level=2)
                self.p(
                    "Referencia de la adaptación: "
                    + clause["localizador_original"]
                    + ". El siguiente extracto del modelo se conserva para identificar la procedencia del texto adaptado."
                )
                # REV-E90-19: el extracto de procedencia es un fragmento propio, recuperable.
                self.p(clause["original_en"], f"ORIGEN-{n}")
        self.h("Alcance de esta adaptación académica")
        self.p(
            "Se tradujeron y abreviaron cláusulas seleccionadas, se sustituyeron partes, moneda y condiciones particulares "
            "por datos ficticios y se omitieron formularios de contratación pública, logotipos y firmas. Las fechas y el "
            "precio de este caso se rigen exclusivamente por sus condiciones particulares. El documento no constituye un "
            "modelo jurídicamente validado para Argentina."
        )
        self.pie()

    def contrato_arg_a(self, fuente: dict) -> None:
        r = self.r
        if (
            r["doc_id"] != "E90-M-CTR-03"
            or r["grupo"] != "medicion"
            or fuente["licencia_adaptacion"] != "CC-BY-NC-SA-4.0"
            or fuente["uso_no_comercial"] is not True
        ):
            raise ValueError("arg_a_fuera_del_alcance_aprobado")
        contexto = r["contexto_contractual"]
        self.p(
            "Entre "
            + contexto["comitente"]
            + ", en adelante Comitente, y "
            + self.rep("contraparte")
            + ", en adelante Arquitecto/a a los fines de este escenario, se acuerda el encargo de proyecto y "
            "dirección. Son sujetos ficticios: no se identifica a una persona real ni se declara matrícula profesional."
        )
        self.h("Condiciones particulares del encargo")
        for campo in CAMPOS["contrato"]:
            self.linea(campo, ETIQUETAS_A[campo] + ":")
        self.p(
            "El objeto de esta operación es el proyecto y la dirección de "
            + contexto["destino"]
            + ". No incluye ejecución material de la construcción.",
            "OBJETO",
        )
        hitos = contexto["hitos_dias_corridos"]
        self.p(
            f"Croquis preliminar: día {hitos['croquis_preliminar']}; anteproyecto: día {hitos['anteproyecto']}; "
            f"documentación de proyecto: día {hitos['documentacion_proyecto']}. Los días se cuentan desde el inicio "
            "de vigencia. La dirección se desarrolla hasta el fin de vigencia. El precio total indicado es el total de "
            "honorarios en pesos argentinos; no es el costo de construcción ni se agregan tributos al total.",
            "COND-hitos",
        )
        for n, clause in enumerate(fuente["clausulas"], 1):
            self.h(f"CLÁUSULA {n}. {clause['titulo'].upper()}")
            self.p(clause["adaptacion_es"], f"CLAUSULA-{n}")
        self.salto()
        self.h("Procedencia, modificaciones y permiso de reutilización")
        self.p(
            "Adaptación académica no comercial de "
            + fuente["titulo"]
            + ", "
            + fuente["version"]
            + ". Autor y organismo: "
            + fuente["atribucion"]
            + ". Fuente: "
            + fuente["url_publicacion"]
            + "."
        )
        self.p(
            "Original y adaptación: Creative Commons Atribución-NoComercial-CompartirIgual 4.0 Internacional "
            "(CC BY-NC-SA 4.0). "
            + fuente["url_licencia"]
            + " La reutilización es exclusivamente no comercial; "
            "conservar atribución, fuente, licencia y registro de cambios. Las adaptaciones compartidas deben "
            "distribuirse bajo CC BY-NC-SA 4.0."
        )
        self.p(
            "Cambios: selección y abreviación de las nueve cláusulas numeradas, incorporación de condiciones y datos "
            "ficticios, sustitución de remisiones a anexos no adjuntos, y omisión de logotipos, firmas y matrículas. "
            "El original está en español; no se realizó traducción. El material se ofrece sin garantías. No hay aval "
            "del CPAU ni validación jurídica para Argentina. El permiso sobre este texto es independiente de la "
            "cláusula sobre planos de un proyecto hipotético."
        )
        self.pie()

    def construir(self, path: Path) -> tuple[dict, list[tuple[str, str]]]:
        self.encabezado()
        if self.r["tipo"] == "contrato":
            self.contrato()
        else:
            {
                "BASE_A": self.base_a,
                "BASE_B": self.base_b,
                "MED_C": self.med_c,
                "MED_D": self.med_d,
            }[self.familia]()
        faltan = [c for c in self.v if c not in self.campos]
        if faltan:
            raise ValueError("campo_sin_ubicar_" + self.r["doc_id"] + "_" + faltan[0])
        path.parent.mkdir(parents=True, exist_ok=True)
        self.doc.save(path)
        docx_determinista(path)
        return self.campos, self.fragmentos


def crear_docx(
    registro: dict, path: Path, fuentes: dict
) -> tuple[dict, list[tuple[str, str]]]:
    return Constructor(registro, fuentes).construir(path)


def rasterizar(
    source: Path, dest: Path, captura: dict, work: Path, pdftoppm: str, titulo: str
) -> None:
    from PIL import Image, ImageEnhance, ImageFilter

    prefix = work / "pagina"
    subprocess.run(
        [pdftoppm, "-r", str(captura["dpi"]), "-png", str(source), str(prefix)],
        check=True,
        capture_output=True,
        timeout=180,
    )
    pages = []
    rng = random.Random(captura["semilla"])
    for path in sorted(work.glob("pagina-*.png")):
        with Image.open(path) as im:
            img = im.convert("RGB")
        img = img.rotate(
            captura["rotacion_grados"],
            resample=Image.Resampling.BICUBIC,
            expand=False,
            fillcolor="white",
        )
        if captura["perspectiva_pixeles"]:
            w, h = img.size
            d = captura["perspectiva_pixeles"]
            img = img.transform(
                img.size,
                Image.Transform.QUAD,
                (d, 0, 0, h - d, w - d, h, w, d),
                resample=Image.Resampling.BICUBIC,
                fillcolor="white",
            )
        # Ruido determinista sin numpy ni una biblioteca adicional.
        small = Image.frombytes(
            "L", (180, 240), bytes(rng.randrange(256) for _ in range(180 * 240))
        )
        noise = small.resize(img.size, resample=Image.Resampling.BILINEAR).convert(
            "RGB"
        )
        img = Image.blend(img, noise, captura["ruido_amplitud"] / 255)
        if captura["sombra_lateral"]:
            gradient = Image.new("L", (256, 1))
            gradient.putdata([min(255, 205 + x // 5) for x in range(256)])
            shade = gradient.resize(img.size).convert("RGB")
            from PIL import ImageChops

            img = ImageChops.multiply(img, shade)
        img = ImageEnhance.Contrast(img).enhance(0.94)
        img = img.filter(ImageFilter.GaussianBlur(captura["desenfoque_radio"]))
        pages.append(img)
    if not pages:
        raise ValueError("raster_sin_paginas")
    raw = work / "raster.pdf"
    pages[0].save(
        raw,
        format="PDF",
        save_all=True,
        append_images=pages[1:],
        resolution=captura["dpi"],
        creationDate=time.gmtime(1792627200),
        modDate=time.gmtime(1792627200),
    )
    normalizar_pdf(raw, dest, titulo)
    for p in pages:
        p.close()


def ubicar_parrafo(paragraphs: list[str], campo: dict) -> list[int]:
    """IA-M04: índices 1-based cuyo texto respalda etiqueta y valor del campo."""
    etiqueta = normalizar(campo["etiqueta"])
    valor = normalizar(campo["representacion"])
    if campo["modo"] == "tabla":
        return [
            i
            for i, t in enumerate(paragraphs, 1)
            if i > 1
            and normalizar(t) == valor
            and normalizar(paragraphs[i - 2]) == etiqueta
        ]
    return [i for i, t in enumerate(paragraphs, 1) if etiqueta + valor in normalizar(t)]


def ubicar_pagina(texts: list[str], campo: dict) -> list[int]:
    """IA-M03: páginas del PDF de referencia con etiqueta y valor contiguos."""
    objetivo = normalizar(campo["etiqueta"]) + normalizar(campo["representacion"])
    return [i for i, t in enumerate(texts, 1) if objetivo in normalizar(t)]


def generar(root: Path, libreoffice: str, pdftoppm: str) -> None:
    from pypdf import PdfReader

    records = leer_json(root / "oraculos/registros.json")["registros"]
    fuentes = leer_json(Path(__file__).parent / "contratos.json")["fuentes"]
    if (root / "oraculos/campos.json").exists():
        raise FileExistsError("ya_generado_usar_otra_carpeta")
    campos = []
    fragments = []
    csvrows = []
    with tempfile.TemporaryDirectory(prefix="atlas-corpus-") as tmp:
        work = Path(tmp)
        fuentesdir = work / "docx"
        pdfdir = work / "pdf"
        docxpaths = []
        ubicaciones = {}
        for r in records:
            path = fuentesdir / r["grupo"] / (r["archivo_id"] + ".docx")
            ubicaciones[r["doc_id"]] = crear_docx(r, path, fuentes)
            docxpaths.append(str(path))
        rendered = work / "render"
        rendered.mkdir()
        result = subprocess.run(
            [
                libreoffice,
                "-env:UserInstallation=" + str((work / "lo-profile").as_uri()),
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                str(rendered),
                *docxpaths,
            ],
            check=False,
            capture_output=True,
            timeout=900,
        )
        if result.returncode:
            raise ValueError("fallo_render_libreoffice")
        for r in records:
            ident = r["doc_id"]
            nombre = r["archivo_id"]
            src = fuentesdir / r["grupo"] / (nombre + ".docx")
            pdf = pdfdir / r["grupo"] / (nombre + ".pdf")
            rendered_pdf = rendered / (nombre + ".pdf")
            if not rendered_pdf.exists():
                raise ValueError("render_pdf_faltante_" + ident)
            normalizar_pdf(rendered_pdf, pdf, nombre)
            final = (
                root
                / "documentos"
                / r["grupo"]
                / (nombre + (".docx" if r["formato"] == "docx" else ".pdf"))
            )
            if final.exists():
                raise FileExistsError("archivo_final_existente_no_sobrescribir")
            final.parent.mkdir(parents=True, exist_ok=True)
            if r["formato"] == "docx":
                shutil.copyfile(src, final)
            elif r["captura"]:
                pagework = work / nombre
                pagework.mkdir()
                rasterizar(pdf, final, r["captura"], pagework, pdftoppm, nombre)
            else:
                shutil.copyfile(pdf, final)
            texts = [p.extract_text() or "" for p in PdfReader(pdf).pages]
            paragraphs = parrafos_docx(src)
            ubic, frags = ubicaciones[ident]
            for f in CAMPOS[r["tipo"]]:
                present = f in r["valores"]
                loc = None
                datos = ubic.get(f)
                if present:
                    paras = ubicar_parrafo(paragraphs, datos)
                    pages = ubicar_pagina(texts, datos)
                    if len(paras) != 1 or not pages:
                        raise ValueError("localizador_no_encontrado_" + ident + "_" + f)
                    if r["formato"] == "docx":
                        loc = {
                            "tipo": "parrafo",
                            "parte": "word/document.xml",
                            "indice": paras[0],
                        }
                    else:
                        loc = {"tipo": "pagina", "pagina": pages[0]}
                if present:
                    fragments.append(
                        {
                            "fragmento_id": ident + "-CAMPO-" + f,
                            "doc_id": ident,
                            "clave": "CAMPO-" + f,
                            "texto": paragraphs[paras[0] - 1],
                            "localizador": {
                                "tipo": "parrafo",
                                "parte": "word/document.xml",
                                "indice": paras[0],
                            }
                            if r["formato"] == "docx"
                            else {"tipo": "pagina", "pagina": pages[0]},
                        }
                    )
                campos.append(
                    {
                        "doc_id": ident,
                        "campo": f,
                        "aplicable": True,
                        "presente": present,
                        "valor_esperado": r["valores"].get(f),
                        "representacion": datos["representacion"] if present else None,
                        "etiqueta": datos["etiqueta"] if present else None,
                        "modo_localizacion": datos["modo"] if present else None,
                        "localizador": loc,
                    }
                )
            vistos = set()
            for clave, texto in frags:
                fid = ident + "-" + clave
                # Los fragmentos de campos se derivan de su localizador, arriba.
                if fid in vistos or clave.startswith("CAMPO-"):
                    continue
                vistos.add(fid)
                indices = [
                    i
                    for i, t in enumerate(paragraphs, 1)
                    if normalizar(t) == normalizar(texto)
                ]
                if not indices:
                    raise ValueError("fragmento_sin_parrafo_" + ident + "_" + clave)
                inicio = normalizar(paragraphs[indices[0] - 1])[:60]
                pages = [j for j, t in enumerate(texts, 1) if inicio in normalizar(t)]
                if not pages:
                    raise ValueError("fragmento_sin_pagina_" + ident + "_" + clave)
                fragments.append(
                    {
                        "fragmento_id": fid,
                        "doc_id": ident,
                        "clave": clave,
                        "texto": paragraphs[indices[0] - 1],
                        "localizador": {
                            "tipo": "parrafo",
                            "parte": "word/document.xml",
                            "indice": indices[0],
                        }
                        if r["formato"] == "docx"
                        else {"tipo": "pagina", "pagina": pages[0]},
                    }
                )
            csvrows.append(
                {
                    "doc_id": ident,
                    "grupo": r["grupo"],
                    "tipo": r["tipo"],
                    "formato": r["formato"],
                    "plantilla_id": r["plantilla_id"],
                    "ruta": str(final.relative_to(root)),
                    "prompt_injection": "si" if r["prompt_injection"] else "no",
                }
            )
    dest = root / "oraculos/inventario.csv"
    dest.parent.mkdir(parents=True, exist_ok=True)
    with dest.open("w", newline="", encoding="utf-8") as stream:
        writer = csv.DictWriter(
            stream, fieldnames=list(csvrows[0]), lineterminator="\n"
        )
        writer.writeheader()
        writer.writerows(csvrows)
    escribir_json(root / "oraculos/campos.json", {"campos": campos})
    escribir_json(root / "oraculos/fragmentos.json", {"fragmentos": fragments})
    print(f"documentos: {len(csvrows)}; campos: {len(campos)}; fragmentos: {len(fragments)}")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--raiz", type=Path, default=Path(__file__).resolve().parents[2]
    )
    parser.add_argument(
        "--libreoffice", default=os.environ.get("E90_LIBREOFFICE", "libreoffice")
    )
    parser.add_argument(
        "--pdftoppm", default=os.environ.get("E90_PDFTOPPM", "pdftoppm")
    )
    args = parser.parse_args()
    generar(args.raiz.resolve(), args.libreoffice, args.pdftoppm)


if __name__ == "__main__":
    main()
