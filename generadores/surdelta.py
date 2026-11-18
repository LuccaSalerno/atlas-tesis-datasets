#!/usr/bin/env python3
"""Genera los documentos del escenario SurDelta (videos y defensa) y su oráculo.

Todos los sujetos son ficticios. Los CUIT son el de la cadena de evidencia y los de la
partición de ajuste del corpus de evaluación (nunca los de medición); en ARCA: «No existe».

    python generadores/surdelta.py [--libreoffice soffice]

Escribe en `demo/surdelta_videos/`. No sobrescribe: para regenerar, borrar antes los PDF y
`oraculo.json`.
"""

from __future__ import annotations

import argparse
import json
import subprocess
import tempfile
from dataclasses import dataclass
from decimal import Decimal
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt, RGBColor
from pypdf import PdfReader, PdfWriter

RAIZ = Path(__file__).resolve().parent.parent
SALIDA = RAIZ / "demo" / "surdelta_videos" / "documentos"
ORACULO = RAIZ / "demo" / "surdelta_videos" / "oraculo.json"
PIE = "Documento ficticio para demostración académica · sin validez comercial ni fiscal"

COMPRADOR = "SurDelta Componentes S.R.L."
# Proveedores ficticios. CUIT con consulta ARCA informada por el equipo («No existe»).
ANDES = ("Andes Sistemas S.A.", "30-96472518-3")  # PD-07, consulta del 23/10/2026
EQUIPAR = ("Distribuidora Equipar S.A.", "30-98641015-1")  # E90 ajuste, 23/10/2026
HORIZONTE = ("Horizonte Servicios S.R.L.", "30-98641014-3")  # E90 ajuste, 23/10/2026
NUBE = ("Nube Austral Tecnología S.R.L.", "30-98641017-8")  # E90 ajuste, 23/10/2026
ACEROS = ("Aceros Cuatro Vientos S.A.", "30-98641018-6")  # E90 ajuste, 23/10/2026


def pesos(valor: str) -> str:
    """'9200000.00' → '9.200.000,00'."""
    entero, decimales = f"{Decimal(valor):.2f}".split(".")
    return f"{int(entero):,}".replace(",", ".") + "," + decimales


def fecha(iso: str) -> str:
    """'2026-09-25' → '25/09/2026'."""
    anio, mes, dia = iso.split("-")
    return f"{dia}/{mes}/{anio}"


class Redactor:
    """Documento DOCX con el formato común del escenario."""

    def __init__(self, membrete: str) -> None:
        self.d = Document()
        seccion = self.d.sections[0]
        seccion.top_margin, seccion.bottom_margin = Cm(2.2), Cm(2.2)
        seccion.left_margin, seccion.right_margin = Cm(2.4), Cm(2.4)
        normal = self.d.styles["Normal"]
        normal.font.name, normal.font.size = "Liberation Sans", Pt(10.5)
        normal.paragraph_format.space_after = Pt(6)
        for estilo, tamano in (("Title", 17), ("Heading 1", 12.5), ("Heading 2", 11)):
            s = self.d.styles[estilo]
            s.font.name, s.font.size, s.font.bold = "Liberation Sans", Pt(tamano), True
            s.font.color.rgb = RGBColor(28, 52, 84)
        encabezado = seccion.header.paragraphs[0]
        encabezado.text = membrete
        encabezado.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        encabezado.runs[0].font.size = Pt(8)
        pie = seccion.footer.paragraphs[0]
        pie.alignment = WD_ALIGN_PARAGRAPH.CENTER
        texto = pie.add_run(PIE)
        texto.font.size, texto.font.color.rgb = Pt(7), RGBColor(120, 120, 120)

    def titulo(self, texto: str) -> None:
        self.d.add_heading(texto, 0)

    def h(self, texto: str) -> None:
        self.d.add_heading(texto, level=1)

    def p(self, texto: str) -> None:
        self.d.add_paragraph(texto)

    def item(self, texto: str) -> None:
        self.d.add_paragraph(texto, style="List Bullet")

    def tabla(self, filas: list[tuple[str, ...]], encabezado: bool = False) -> None:
        tabla = self.d.add_table(rows=0, cols=len(filas[0]))
        tabla.style, tabla.alignment = "Table Grid", WD_TABLE_ALIGNMENT.CENTER
        for i, fila in enumerate(filas):
            celdas = tabla.add_row().cells
            for celda, valor in zip(celdas, fila, strict=True):
                celda.text = valor
                if encabezado and i == 0:
                    celda.paragraphs[0].runs[0].font.bold = True
        self.d.add_paragraph()

    def pagina(self) -> None:
        self.d.add_page_break()

    def guardar(self, ruta: Path) -> None:
        self.d.save(ruta)


@dataclass(frozen=True)
class Pieza:
    codigo: str
    archivo: str
    tipo: str
    campos: dict[str, str]
    redactor: Redactor


# --- Contratos ----------------------------------------------------------------------------


def contrato_mantenimiento() -> Pieza:
    """Contrato largo (10 páginas): la vigencia, el precio y los trabajos adicionales quedan
    lejos de la carátula, para mostrar la fuente de cada dato en un documento real."""
    proveedor, cuit = ANDES
    r = Redactor("SurDelta Componentes S.R.L. · Contrato C-2026-014")
    normal = r.d.styles["Normal"].paragraph_format
    normal.line_spacing, normal.space_after = 1.4, Pt(9)
    r.titulo("Contrato de servicio de mantenimiento integral de planta")
    r.p(
        "En la localidad de Avellaneda, provincia de Buenos Aires, a los 15 días del mes de"
        " septiembre de 2026, se celebra el presente contrato de servicios entre las partes"
        " que se identifican a continuación, que en adelante se denominarán «la comitente» y"
        " «la prestadora»."
    )
    r.tabla(
        [
            ("Contrato N.º", "C-2026-014"),
            ("Comitente", COMPRADOR),
            ("Prestadora / contraparte", proveedor),
            ("CUIT de la prestadora", cuit),
            ("Fecha de celebración", "15/09/2026"),
        ]
    )
    r.d.add_heading("Antecedentes", level=1)
    for texto in (
        (
            "La comitente opera una planta de fabricación de componentes industriales con una línea"
            " de montaje automatizada, un depósito de insumos y oficinas administrativas, todos"
            " ellos dependientes de una red de datos y de servidores locales."
        ),
        (
            "Las interrupciones de esa infraestructura detienen la línea de montaje y demoran los"
            " despachos, por lo que la comitente necesita un servicio de mantenimiento con tiempos"
            " de respuesta acotados, registro de cada intervención y un costo previsible."
        ),
        (
            "La prestadora declara contar con personal técnico propio, herramientas y repuestos"
            " suficientes para prestar el servicio en las condiciones que se pactan, y haber"
            " relevado la planta antes de presentar su propuesta."
        ),
        (
            "En atención a lo expuesto, las partes acuerdan celebrar el presente contrato, que se"
            " regirá por las cláusulas siguientes y por los anexos que lo integran."
        ),
    ):
        r.p(texto)
    r.pagina()

    r.h("1. Objeto")
    r.p(
        "1.1. La prestadora realizará el mantenimiento preventivo y correctivo de la"
        " infraestructura tecnológica de la planta de la comitente: equipos de red, servidores"
        " de respaldo, estaciones de control de la línea de montaje y su cableado estructurado,"
        " según el inventario del Anexo A."
    )
    r.p("1.2. El servicio comprende, como abono mensual:")
    for texto in (
        "Una visita preventiva mensual con informe técnico.",
        "Atención de incidentes críticos en un plazo máximo de 4 horas hábiles.",
        "Hasta 20 horas mensuales de soporte correctivo incluidas en el abono.",
        "Informe trimestral de estado de la infraestructura con recomendaciones.",
    ):
        r.item(texto)
    r.p(
        "1.3. Quedan fuera del objeto la provisión de equipos nuevos, el desarrollo de"
        " programas, la reparación de maquinaria de producción y cualquier tarea sobre"
        " instalaciones eléctricas que no sean las del cableado de datos."
    )
    r.p(
        "1.4. Los niveles de servicio de cada tipo de incidente se detallan en el Anexo B y"
        " forman parte de las obligaciones de la prestadora."
    )
    r.h("2. Obligaciones de la prestadora")
    r.p(
        "2.1. Designar un responsable técnico, que será el interlocutor de la comitente para"
        " toda cuestión operativa, y un reemplazante para sus ausencias."
    )
    r.p(
        "2.2. Registrar cada intervención con su fecha, alcance, horas insumidas y resultado,"
        " y entregar ese registro junto con cada factura mensual."
    )
    r.pagina()
    r.p(
        "2.3. Prestar el servicio con personal propio, debidamente capacitado y registrado,"
        " del que será la única empleadora a todos los efectos legales."
    )
    r.p(
        "2.4. Utilizar repuestos nuevos u originales del fabricante, salvo acuerdo escrito en"
        " contrario, y conservar los componentes reemplazados a disposición de la comitente."
    )
    r.p(
        "2.5. Cumplir las normas internas de la planta sobre accesos, horarios, circulación y"
        " uso de elementos de protección personal."
    )
    r.p(
        "2.6. No modificar órdenes de compra ni aprobar gastos en nombre de la comitente, ni"
        " comprometerla frente a terceros."
    )
    r.d.add_heading("Obligaciones de la comitente", level=2)
    r.p(
        "2.7. La comitente facilitará el acceso a la planta en los horarios convenidos,"
        " designará un referente interno para recibir los informes y emitirá cada mes la"
        " orden de compra correspondiente al abono."
    )
    r.p(
        "2.8. La comitente comunicará con anticipación las paradas programadas de la línea de"
        " montaje, para que las tareas preventivas se hagan sin afectar la producción."
    )
    r.p(
        "2.9. La comitente informará cualquier cambio en el inventario del Anexo A, que podrá"
        " actualizarse por acuerdo escrito sin modificar el resto del contrato."
    )
    r.pagina()

    r.h("3. Vigencia")
    r.p("3.1. El contrato rige desde el 01/10/2026 hasta el 30/09/2027, ambas fechas inclusive.")
    r.tabla([("Inicio de vigencia", "01/10/2026"), ("Fin de vigencia", "30/09/2027")])
    r.p(
        "3.2. Su renovación requiere un nuevo instrumento firmado por ambas partes. No habrá"
        " renovación tácita ni automática, aunque la prestadora continúe prestando el servicio."
    )
    r.p(
        "3.3. Durante los treinta días anteriores al fin de la vigencia, la prestadora entregará"
        " un informe de cierre con el estado de cada equipo del inventario, las intervenciones"
        " pendientes y las recomendaciones para el período siguiente."
    )
    r.p(
        "3.4. Si al vencimiento el contrato no se renueva, la prestadora colaborará durante un"
        " período de transición razonable con quien la reemplace, entregando la documentación"
        " técnica, las credenciales de administración y los registros de intervenciones."
    )
    r.pagina()

    r.h("4. Precio y forma de pago")
    r.p(
        "4.1. El precio total del contrato asciende a PESOS NOVENTA Y SEIS MILLONES"
        " (ARS 96.000.000,00), IVA incluido, pagadero en doce abonos mensuales de"
        " ARS 8.000.000,00."
    )
    r.tabla(
        [
            ("Monto total del contrato", "ARS 96.000.000,00"),
            ("Moneda", "Pesos argentinos (ARS)"),
            ("Abono mensual", "ARS 8.000.000,00"),
        ]
    )
    r.p(
        "4.2. Cada abono se factura a mes vencido contra la orden de compra mensual emitida por"
        " la comitente. El precio es fijo durante toda la vigencia y no está sujeto a ajustes."
    )
    r.p(
        "4.3. La comitente pagará cada factura conformada dentro de los treinta días corridos"
        " de recibida, mediante transferencia bancaria a la cuenta que la prestadora informe"
        " por escrito."
    )
    r.p(
        "4.4. La comitente practicará las retenciones impositivas que correspondan según la"
        " normativa vigente y entregará los comprobantes respectivos."
    )
    r.pagina()

    r.h("5. Trabajos adicionales")
    r.p(
        "5.1. Las horas de soporte que excedan las incluidas en el abono, y cualquier trabajo"
        " fuera del objeto, se consideran trabajos adicionales."
    )
    r.p(
        "5.2. Los trabajos adicionales requieren autorización previa y por escrito de la"
        " comitente, mediante una orden de compra específica. Sin esa autorización no serán"
        " reconocidos ni podrán facturarse."
    )
    r.p(
        "5.3. Para solicitar la autorización, la prestadora presentará un presupuesto con la"
        " descripción del trabajo, las horas estimadas y su costo. La comitente responderá"
        " dentro de los dos días hábiles."
    )
    r.p(
        "5.4. Ante un incidente crítico que exceda las horas del abono, la prestadora podrá"
        " iniciar la atención y deberá pedir la autorización en el mismo día. Las horas que la"
        " comitente no autorice correrán por cuenta de la prestadora."
    )
    r.h("6. Facturación y conformidad")
    r.p(
        "6.1. La prestadora emitirá una factura por cada abono mensual, con referencia a la"
        " orden de compra correspondiente y al registro de intervenciones del mes."
    )
    r.p(
        "6.2. La comitente revisará cada factura contra el contrato y la orden de compra antes"
        " de prestar conformidad. Una factura con importes no autorizados será observada y"
        " deberá ser reemplazada, previa nota de crédito por la factura observada."
    )
    r.pagina()

    r.h("7. Confidencialidad")
    r.p(
        "7.1. Las partes mantendrán reserva sobre la información técnica y comercial a la que"
        " accedan con motivo del contrato y la utilizarán solo para su ejecución."
    )
    r.p(
        "7.2. La obligación de reserva se extiende al personal de cada parte y subsiste"
        " durante dos años después de la finalización del contrato."
    )
    r.h("8. Rescisión")
    r.p(
        "8.1. Cualquiera de las partes podrá rescindir el contrato con un preaviso de 60 días."
        " La rescisión no afecta las obligaciones devengadas hasta esa fecha."
    )
    r.p(
        "8.2. La comitente podrá rescindir sin preaviso si la prestadora incumple de manera"
        " reiterada los niveles de servicio del Anexo B, previa intimación fehaciente."
    )
    r.h("9. Responsabilidad y seguros")
    r.p(
        "9.1. La prestadora responderá por los daños que su personal cause a los bienes de la"
        " comitente o de terceros durante la prestación del servicio."
    )
    r.p(
        "9.2. La prestadora mantendrá vigentes los seguros de riesgos del trabajo de su"
        " personal y de responsabilidad civil, y entregará sus constancias cuando la comitente"
        " lo solicite."
    )
    r.pagina()

    r.h("10. Seguridad e higiene en planta")
    r.p(
        "10.1. El personal de la prestadora recibirá la inducción de seguridad de la planta"
        " antes de su primera visita y respetará las indicaciones del referente de seguridad."
    )
    r.p(
        "10.2. Ninguna intervención sobre las estaciones de control de la línea de montaje se"
        " hará con la línea en marcha, salvo autorización del jefe de producción."
    )
    r.h("11. Penalidades por incumplimiento de los niveles de servicio")
    r.p(
        "11.1. Cada incidente crítico atendido fuera del plazo del Anexo B dará lugar a una"
        " bonificación del 2 % del abono del mes, hasta un máximo del 10 % mensual."
    )
    r.h("12. Cesión y subcontratación")
    r.p(
        "12.1. La prestadora no podrá ceder el contrato ni subcontratar el servicio, total o"
        " parcialmente, sin consentimiento previo y por escrito de la comitente."
    )
    r.h("13. Comunicaciones")
    r.p(
        "13.1. Las comunicaciones entre las partes se harán por escrito, al domicilio o a la"
        " dirección de correo electrónico que cada una informe, y se tendrán por recibidas el"
        " día hábil siguiente a su envío."
    )
    r.h("14. Jurisdicción")
    r.p(
        "14.1. Para cualquier controversia, las partes se someten a los tribunales ordinarios"
        " de la ciudad de Avellaneda, con renuncia a cualquier otro fuero."
    )
    r.pagina()

    r.h("Anexo A. Inventario de equipos comprendidos")
    r.tabla(
        [
            ("Sector", "Equipo", "Cantidad"),
            ("Centro de datos", "Servidores de respaldo", "2"),
            ("Centro de datos", "Conmutadores de red principales", "3"),
            ("Línea de montaje", "Estaciones de control", "6"),
            ("Línea de montaje", "Puntos de acceso inalámbrico industriales", "4"),
            ("Depósito", "Terminales de lectura de códigos", "5"),
            ("Oficinas", "Puestos de trabajo y periféricos", "28"),
            ("Toda la planta", "Cableado estructurado y racks", "1 instalación"),
        ],
        encabezado=True,
    )
    r.h("Anexo B. Niveles de servicio")
    r.tabla(
        [
            ("Tipo de incidente", "Ejemplo", "Plazo de atención"),
            ("Crítico", "Línea de montaje detenida", "4 horas hábiles"),
            ("Alto", "Servidor de respaldo fuera de servicio", "1 día hábil"),
            ("Medio", "Falla de un puesto o periférico", "2 días hábiles"),
            ("Bajo", "Consulta o ajuste menor", "5 días hábiles"),
        ],
        encabezado=True,
    )
    r.pagina()

    r.h("Firmas")
    r.p(
        "En prueba de conformidad, se firman dos ejemplares de un mismo tenor y a un solo"
        " efecto, en el lugar y la fecha indicados en el encabezamiento."
    )
    r.tabla(
        [
            ("Por la comitente", "Por la prestadora"),
            (COMPRADOR, proveedor),
            ("Firma: ____________________", "Firma: ____________________"),
            ("Aclaración: ________________", "Aclaración: ________________"),
            ("Cargo: Gerencia de Administración", "Cargo: Apoderado/a"),
        ],
        encabezado=True,
    )
    return Pieza(
        "SD-C1",
        "SD-C1_Contrato_Mantenimiento_Planta.pdf",
        "contrato",
        {
            "contraparte": proveedor,
            "cuit_contraparte": cuit,
            "fecha_celebracion": "2026-09-15",
            "vigencia_desde": "2026-10-01",
            "vigencia_hasta": "2027-09-30",
            "monto_total": "96000000.00",
            "moneda": "ARS",
        },
        r,
    )


def contrato_soporte() -> Pieza:
    proveedor, cuit = NUBE
    r = Redactor("SurDelta Componentes S.R.L. · Contrato C-2026-021")
    r.titulo("Contrato de soporte informático")
    r.p(
        "En la localidad de Avellaneda, provincia de Buenos Aires, a los 30 días del mes de"
        " octubre de 2026, se celebra el presente contrato de soporte informático."
    )
    r.tabla(
        [
            ("Contrato N.º", "C-2026-021"),
            ("Comitente", COMPRADOR),
            ("Prestadora / contraparte", proveedor),
            ("CUIT de la prestadora", cuit),
            ("Fecha de celebración", "30/10/2026"),
        ]
    )
    r.h("1. Objeto")
    r.p(
        "Mesa de ayuda para las áreas administrativa y comercial, administración de cuentas"
        " de correo y licencias, y mantenimiento de 40 puestos de trabajo."
    )
    r.h("2. Vigencia")
    r.p("El contrato rige desde el 01/12/2026 hasta el 30/11/2027, ambas fechas inclusive.")
    r.h("3. Precio")
    r.p(
        "El precio total asciende a ARS 18.000.000,00, IVA incluido, en pesos argentinos,"
        " pagadero en doce cuotas mensuales iguales."
    )
    r.tabla(
        [
            ("Monto total del contrato", "ARS 18.000.000,00"),
            ("Moneda", "Pesos argentinos (ARS)"),
            ("Inicio de vigencia", "01/12/2026"),
            ("Fin de vigencia", "30/11/2027"),
        ]
    )
    return Pieza(
        "SD-C3",
        "SD-C3_Contrato_Soporte_Informatico.pdf",
        "contrato",
        {
            "contraparte": proveedor,
            "cuit_contraparte": cuit,
            "fecha_celebracion": "2026-10-30",
            "vigencia_desde": "2026-12-01",
            "vigencia_hasta": "2027-11-30",
            "monto_total": "18000000.00",
            "moneda": "ARS",
        },
        r,
    )


def contrato_suministro() -> Pieza:
    """Provisión de la materia prima de la planta: un trámite sin conflictos, cerrado."""
    proveedor, cuit = ACEROS
    r = Redactor("SurDelta Componentes S.R.L. · Contrato C-2026-009")
    r.titulo("Contrato de provisión de acero inoxidable")
    r.p(
        "En la localidad de Avellaneda, provincia de Buenos Aires, a los 20 días del mes de"
        " julio de 2026, se celebra el presente contrato de provisión entre las partes que se"
        " identifican a continuación."
    )
    r.tabla(
        [
            ("Contrato N.º", "C-2026-009"),
            ("Comitente", COMPRADOR),
            ("Proveedora / contraparte", proveedor),
            ("CUIT de la proveedora", cuit),
            ("Fecha de celebración", "20/07/2026"),
        ]
    )
    r.h("1. Objeto")
    r.p(
        "La proveedora entregará a la comitente barras y chapas de acero inoxidable AISI 316,"
        " destinadas a la fabricación de cuerpos y vástagos de la línea de válvulas"
        " industriales de la planta."
    )
    for texto in (
        "Una entrega mensual de 1.800 kg, en la planta de la comitente.",
        "Certificado de colada y de composición química con cada entrega.",
        "Reposición sin cargo del material que no cumpla la especificación.",
    ):
        r.item(texto)
    r.h("2. Calidad y recepción")
    r.p(
        "La comitente inspeccionará cada entrega dentro de los cinco días hábiles de recibida."
        " El material rechazado se retirará a cargo de la proveedora, que deberá reponerlo"
        " antes de la entrega siguiente."
    )
    r.pagina()
    r.h("3. Vigencia")
    r.p("El contrato rige desde el 01/08/2026 hasta el 31/07/2027, ambas fechas inclusive.")
    r.tabla([("Inicio de vigencia", "01/08/2026"), ("Fin de vigencia", "31/07/2027")])
    r.h("4. Precio y forma de pago")
    r.p(
        "El precio total del contrato asciende a PESOS CINCUENTA Y CUATRO MILLONES"
        " (ARS 54.000.000,00), IVA incluido, en doce entregas mensuales de ARS 4.500.000,00."
        " Cada entrega se factura contra la orden de compra mensual emitida por la comitente."
    )
    r.tabla(
        [
            ("Monto total del contrato", "ARS 54.000.000,00"),
            ("Moneda", "Pesos argentinos (ARS)"),
            ("Entrega mensual", "ARS 4.500.000,00"),
        ]
    )
    r.h("5. Facturación")
    r.p(
        "La proveedora emitirá una factura por cada entrega, con referencia a la orden de"
        " compra correspondiente y al remito firmado en la recepción."
    )
    return Pieza(
        "SD-C4",
        "SD-C4_Contrato_Provision_Acero.pdf",
        "contrato",
        {
            "contraparte": proveedor,
            "cuit_contraparte": cuit,
            "fecha_celebracion": "2026-07-20",
            "vigencia_desde": "2026-08-01",
            "vigencia_hasta": "2027-07-31",
            "monto_total": "54000000.00",
            "moneda": "ARS",
        },
        r,
    )


# --- Órdenes, presupuestos y facturas -----------------------------------------------------


def orden(
    codigo: str,
    numero: str,
    iso: str,
    sujeto: tuple[str, str],
    total: str,
    items: list[tuple[str, str]],
    referencia: str,
) -> Pieza:
    proveedor, cuit = sujeto
    r = Redactor(f"SurDelta Componentes S.R.L. · Orden de compra {numero}")
    r.titulo("Orden de compra")
    r.tabla(
        [
            ("Orden de compra N.º", numero),
            ("Fecha", fecha(iso)),
            ("Emitida por", COMPRADOR),
            ("Proveedor", proveedor),
            ("CUIT del proveedor", cuit),
            ("Referencia", referencia),
        ]
    )
    r.h("Detalle")
    r.tabla(
        [("Concepto", "Importe (IVA incluido)"), *[(c, f"ARS {pesos(i)}") for c, i in items]],
        encabezado=True,
    )
    r.tabla([("Importe total", f"ARS {pesos(total)}"), ("Moneda", "Pesos argentinos (ARS)")])
    r.p("Autoriza: Gerencia de Administración. Facturar a mes vencido con referencia a esta orden.")
    return Pieza(
        codigo,
        f"{codigo}_Orden_Compra_{numero}.pdf",
        "orden_compra",
        {
            "numero": numero,
            "fecha": iso,
            "proveedor": proveedor,
            "cuit_proveedor": cuit,
            "importe_total": total,
            "moneda": "ARS",
        },
        r,
    )


def presupuesto(
    codigo: str,
    numero: str,
    iso: str,
    sujeto: tuple[str, str],
    total: str,
    items: list[tuple[str, str]],
) -> Pieza:
    proveedor, cuit = sujeto
    r = Redactor(f"{proveedor} · Presupuesto {numero}")
    r.titulo("Presupuesto")
    r.tabla(
        [
            ("Presupuesto N.º", numero),
            ("Fecha", fecha(iso)),
            ("Proveedor", proveedor),
            ("CUIT del proveedor", cuit),
            ("Cliente", COMPRADOR),
            ("Validez de la oferta", "30 días"),
        ]
    )
    r.h("Detalle")
    r.tabla(
        [("Concepto", "Importe (IVA incluido)"), *[(c, f"ARS {pesos(i)}") for c, i in items]],
        encabezado=True,
    )
    r.tabla([("Importe total", f"ARS {pesos(total)}"), ("Moneda", "Pesos argentinos (ARS)")])
    return Pieza(
        codigo,
        f"{codigo}_Presupuesto_{numero}.pdf",
        "presupuesto",
        {
            "numero": numero,
            "fecha": iso,
            "proveedor": proveedor,
            "cuit_proveedor": cuit,
            "importe_total": total,
            "moneda": "ARS",
        },
        r,
    )


def factura(
    codigo: str,
    numero: str,
    iso: str,
    sujeto: tuple[str, str],
    total: str,
    items: list[tuple[str, str]],
    referencia: str,
    nota: str | None = None,
) -> Pieza:
    emisor, cuit = sujeto
    r = Redactor(f"{emisor} · Factura A {numero}")
    r.titulo("Factura A")
    r.tabla(
        [
            ("Factura N.º", numero),
            ("Fecha de emisión", fecha(iso)),
            ("Emisor", emisor),
            ("CUIT del emisor", cuit),
            ("Cliente", COMPRADOR),
            ("Referencia", referencia),
        ]
    )
    r.h("Detalle")
    r.tabla(
        [("Concepto", "Importe (IVA incluido)"), *[(c, f"ARS {pesos(i)}") for c, i in items]],
        encabezado=True,
    )
    neto = (Decimal(total) / Decimal("1.21")).quantize(Decimal("0.01"))
    r.tabla(
        [
            ("Subtotal neto", f"ARS {pesos(str(neto))}"),
            ("IVA 21 %", f"ARS {pesos(str(Decimal(total) - neto))}"),
            ("Importe total", f"ARS {pesos(total)}"),
            ("Moneda", "Pesos argentinos (ARS)"),
        ]
    )
    if nota:
        r.p(nota)
    r.p("Comprobante ficticio: sin CAE ni código QR fiscal.")
    return Pieza(
        codigo,
        f"{codigo}_Factura_{numero}.pdf",
        "factura",
        {
            "numero": numero,
            "fecha_emision": iso,
            "emisor": emisor,
            "cuit_emisor": cuit,
            "importe_total": total,
            "moneda": "ARS",
        },
        r,
    )


def piezas() -> list[Pieza]:
    abono = [("Abono mensual de mantenimiento integral, octubre de 2026", "8000000.00")]
    return [
        # Historia principal: R4 en la factura v1, R3 aceptada en la OC.
        contrato_mantenimiento(),
        orden(
            "SD-O1", "OC-2026-118", "2026-09-25", ANDES, "8000000.00", abono, "Contrato C-2026-014"
        ),
        factura(
            "SD-F1",
            "0004-00001248",
            "2026-10-31",
            ANDES,
            "9200000.00",
            [*abono, ("Horas adicionales de soporte correctivo (40 h)", "1200000.00")],
            "Orden de compra OC-2026-118",
        ),
        factura(
            "SD-F2",
            "0004-00001263",
            "2026-11-05",
            ANDES,
            "8000000.00",
            abono,
            "Orden de compra OC-2026-118",
            "Reemplaza a la factura 0004-00001248, anulada mediante la nota de crédito"
            " 0004-00000087.",
        ),
        # Compra de equipos: la factura llega con el CUIT de otra empresa (R1).
        presupuesto(
            "SD-P2",
            "PRE-0457",
            "2026-10-12",
            EQUIPAR,
            "14500000.00",
            [
                ("6 notebooks de 14 pulgadas", "10800000.00"),
                ("6 monitores de 24 pulgadas", "3700000.00"),
            ],
        ),
        orden(
            "SD-O2",
            "OC-2026-131",
            "2026-10-20",
            EQUIPAR,
            "14500000.00",
            [("6 notebooks y 6 monitores según presupuesto PRE-0457", "14500000.00")],
            "Presupuesto PRE-0457",
        ),
        factura(
            "SD-F3",
            "0002-00000519",
            "2026-10-28",
            HORIZONTE,
            "14500000.00",
            [("6 notebooks y 6 monitores", "14500000.00")],
            "Orden de compra OC-2026-131",
        ),
        # Renovación de soporte: otro tipo de gestión, sin advertencias.
        contrato_soporte(),
        presupuesto(
            "SD-P3",
            "PRE-NA-0112",
            "2026-10-20",
            NUBE,
            "18000000.00",
            [("Soporte informático anual, 40 puestos", "18000000.00")],
        ),
        # Provisión de acero: el caso habitual, sin advertencias, cerrado.
        contrato_suministro(),
        orden(
            "SD-O4",
            "OC-2026-112",
            "2026-10-02",
            ACEROS,
            "4500000.00",
            [("Entrega de octubre: 1.800 kg de barras y chapas AISI 316", "4500000.00")],
            "Contrato C-2026-009",
        ),
        factura(
            "SD-F4",
            "0003-00000871",
            "2026-10-16",
            ACEROS,
            "4500000.00",
            [("Entrega de octubre: 1.800 kg de barras y chapas AISI 316", "4500000.00")],
            "Orden de compra OC-2026-112",
        ),
    ]


def a_pdf(libreoffice: str, docx: Path, destino: Path, titulo: str) -> None:
    with tempfile.TemporaryDirectory() as tmp:
        subprocess.run(
            [
                libreoffice,
                f"-env:UserInstallation={Path(tmp, 'perfil').as_uri()}",
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                tmp,
                str(docx),
            ],
            check=True,
            capture_output=True,
            timeout=180,
        )
        escritor = PdfWriter()
        escritor.clone_document_from_reader(PdfReader(Path(tmp, docx.stem + ".pdf")))
        escritor.add_metadata(
            {
                "/Title": titulo,
                "/Author": "ATLAS · escenario de demostración SurDelta",
                "/CreationDate": "D:20261109000000Z",
                "/ModDate": "D:20261109000000Z",
            }
        )
        with destino.open("wb") as salida:
            escritor.write(salida)


def main(argumentos: list[str] | None = None) -> int:
    analizador = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    analizador.add_argument("--libreoffice", default="soffice")
    opciones = analizador.parse_args(argumentos)
    if ORACULO.exists() or (SALIDA.exists() and any(SALIDA.iterdir())):
        print("El corpus SurDelta ya existe: no se sobrescribe", flush=True)
        return 1
    SALIDA.mkdir(parents=True, exist_ok=True)
    documentos = {}
    with tempfile.TemporaryDirectory() as tmp:
        for pieza in piezas():
            docx = Path(tmp, pieza.archivo.removesuffix(".pdf") + ".docx")
            pieza.redactor.guardar(docx)
            destino = SALIDA / pieza.archivo
            a_pdf(opciones.libreoffice, docx, destino, pieza.archivo.removesuffix(".pdf"))
            documentos[pieza.codigo] = {
                "archivo": pieza.archivo,
                "tipo_documental": pieza.tipo,
                "paginas": len(PdfReader(destino).pages),
                "campos": pieza.campos,
            }
    oraculo = {
        "id": "surdelta",
        "uso": "Demostración (videos y defensa). No acredita EVD-01 ni métricas de §6.3.",
        "documentos": documentos,
    }
    ORACULO.write_text(json.dumps(oraculo, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    print(
        f"{len(documentos)} PDF en {SALIDA.relative_to(RAIZ)}; oráculo en {ORACULO.relative_to(RAIZ)}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
