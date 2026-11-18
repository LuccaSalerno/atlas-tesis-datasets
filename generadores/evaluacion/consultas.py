"""Paso 3: las 30 preguntas de RAG, escritas sobre los documentos ya generados.

Preguntas en lenguaje natural y sin IDs: directas, sobre condiciones contractuales,
multidocumento y sin evidencia (el sistema debe abstenerse).
"""

from __future__ import annotations

import argparse
from datetime import date
from decimal import Decimal
from pathlib import Path

from .comun import (
    NOMBRES_MONEDA,
    escribir_json,
    importe_es,
    leer_json,
)


# REV-E90-19 (DO-2026-10-23-CIERRE D5): función experimental del extracto inglés de
# procedencia en los contratos extensos. Es parte del contexto disponible y del conjunto
# relevante de recuperación; no prevalece sobre la adaptación ni agrega afirmaciones
# exigidas.
EXTRACTO_DE_PROCEDENCIA = (
    "Se aceptan como correctos los plazos y supuestos más específicos del extracto en inglés "
    "del modelo público (fragmento ORIGEN), con cita a ese extracto, siempre que no "
    "contradigan la adaptación española; no se exigen para una respuesta completa."
)

def fecha_es(iso: str) -> str:
    y, m, d = iso.split("-")
    return f"{d}/{m}/{y}"


def preparar_consultas(root: Path) -> None:
    destino = root / "oraculos/consultas.json"
    if destino.exists():
        raise FileExistsError("consultas_ya_preparadas_no_sobrescribir")
    datos = leer_json(root / "oraculos/registros.json")
    registros = {r["doc_id"]: r for r in datos["registros"]}
    contextos = {k: v["docs"] for k, v in datos["contextos"].items()}
    frags = {
        f["fragmento_id"]
        for f in leer_json(root / "oraculos/fragmentos.json")[
            "fragmentos"
        ]
    }
    result = []

    def v(ident: str, campo: str) -> str:
        return registros[ident]["valores"][campo]

    def moneda(ident: str) -> str:
        code = v(ident, "moneda")
        return f"{NOMBRES_MONEDA[code]} ({code})"

    def monto(ident: str) -> str:
        r = registros[ident]
        campo = next(
            c for c in ("monto_total", "importe_total", "importe") if c in r["valores"]
        )
        return importe_es(r["valores"][campo]) + " " + r["valores"]["moneda"]

    def ref(ident: str, clave: str) -> str:
        fid = ident + "-" + clave
        if fid not in frags:
            raise ValueError("fragmento_referencia_no_construido")
        return fid

    def add(categoria, pregunta, contexto, afirmaciones, refs, motivo=None, variantes=None):
        if categoria != "sin_evidencia" and not afirmaciones:
            raise ValueError("respondible_sin_afirmaciones")
        result.append(
            {
                "pregunta_id": f"Q{len(result) + 1:02d}",
                "categoria": categoria,
                "clase": "sin_evidencia"
                if categoria == "sin_evidencia"
                else "respondible",
                "pregunta": pregunta,
                "documentos_autorizados": list(contexto),
                "respuesta_referencia": " ".join(afirmaciones)
                if afirmaciones
                else None,
                "afirmaciones_esperadas": afirmaciones,
                "fragmentos_relevantes": refs,
                "motivo_abstencion": motivo,
                # REV-E90-19: precisiones que se aceptan sin exigirlas, fijadas antes de medir.
                "variantes_admitidas": variantes,
                "consenso": None,
            }
        )

    t1, t2, t3, t4 = (contextos["E90-CONTEXTO-T" + n] for n in "1234")
    pre1, oc1, fac1, _com1 = t1
    ctr2, _oc2, fac2 = t2
    pre3, oc3, fac3 = t3
    oc4, fac4, com4 = t4
    p1 = v(oc1, "proveedor")
    p2 = v(ctr2, "contraparte")
    p3 = v(pre3, "proveedor")
    p4 = v(oc4, "proveedor")

    # Directas (10): tipos, formatos y contextos distintos, con capturas y adversariales.
    add(
        "directa",
        f"¿Cuál es el importe total de la orden de compra emitida a {p1}?",
        t1,
        [
            f"El importe total de la orden de compra es {importe_es(v(oc1, 'importe_total'))}.",
            f"Está expresado en {moneda(oc1)}.",
        ],
        [ref(oc1, "CAMPO-importe_total"), ref(oc1, "CAMPO-moneda")],
    )
    add(
        "directa",
        f"¿Qué número tiene la factura de {p1} por {registros[fac1]['servicio']}?",
        t1,
        [f"La factura tiene el número {v(fac1, 'numero')}."],
        [ref(fac1, "CAMPO-numero")],
    )
    add(
        "directa",
        f"¿Hasta qué fecha está vigente el contrato con {p2}?",
        t2,
        [f"El contrato está vigente hasta el {fecha_es(v(ctr2, 'vigencia_hasta'))}."],
        [ref(ctr2, "CAMPO-vigencia_hasta")],
    )
    add(
        "directa",
        f"¿En qué fecha se emitió el presupuesto de {p3}?",
        t3,
        [f"El presupuesto se emitió el {fecha_es(v(pre3, 'fecha'))}."],
        [ref(pre3, "CAMPO-fecha")],
    )
    add(
        "directa",
        f"¿Por qué importe se dejó constancia de la recepción de los servicios de {p4}?",
        t4,
        [
            f"La constancia registra un importe de {importe_es(v(com4, 'importe'))}.",
            f"Está expresado en {moneda(com4)}.",
        ],
        [ref(com4, "CAMPO-importe"), ref(com4, "CAMPO-moneda")],
    )
    for ident, pregunta in (
        (
            "E90-M-CTR-09",
            "¿Cuál es el precio total del contrato de servicios técnicos y en qué moneda está expresado?",
        ),
        (
            "E90-M-CTR-03",
            "¿Cuál es el monto total de honorarios del encargo de proyecto y dirección?",
        ),
    ):
        add(
            "directa",
            pregunta,
            [ident],
            [
                f"El monto total es {importe_es(v(ident, 'monto_total'))}.",
                f"Está expresado en {moneda(ident)}.",
            ],
            [ref(ident, "CAMPO-monto_total"), ref(ident, "CAMPO-moneda")],
        )
    for ident, pregunta, campo, texto in (
        (
            "E90-M-OC-10",
            "¿Qué número tiene la orden de compra?",
            "numero",
            "La orden de compra tiene el número {}.",
        ),
        (
            "E90-M-PRE-02",
            "¿Cuál es el CUIT del proveedor que emitió la cotización?",
            "cuit_proveedor",
            "El CUIT del proveedor es {}.",
        ),
        (
            "E90-M-FAC-06",
            "¿Cuál es el CUIT del emisor de la factura?",
            "cuit_emisor",
            "El CUIT del emisor es {}.",
        ),
    ):
        add(
            "directa",
            pregunta,
            [ident],
            [texto.format(v(ident, campo))],
            [ref(ident, "CAMPO-" + campo)],
        )

    # Condiciones (8): cinco sobre cláusulas contractuales y tres comerciales.
    add(
        "condiciones",
        "¿En qué casos puede una de las partes dar por terminado el contrato?",
        ["E90-M-CTR-06"],
        [
            "Ante un incumplimiento material no subsanado.",
            "Ante una imposibilidad grave de continuidad.",
            "Por otra causal pactada.",
            "La parte debe notificar el motivo.",
        ],
        [ref("E90-M-CTR-06", "CLAUSULA-7"), ref("E90-M-CTR-06", "ORIGEN-7")],
        variantes=EXTRACTO_DE_PROCEDENCIA,
    )
    add(
        "condiciones",
        "Si el proveedor incurre en un incumplimiento notificable, ¿qué puede pedirle el comprador y qué puede hacer con lo que reciba?",
        ["E90-M-CTR-08"],
        [
            "El comprador puede solicitar un plan de rectificación con las acciones y plazos propuestos.",
            "Puede aceptar el plan o rechazarlo con razones.",
            "Ante un rechazo, puede requerir una versión revisada.",
        ],
        [ref("E90-M-CTR-08", "CLAUSULA-7"), ref("E90-M-CTR-08", "ORIGEN-7")],
        variantes=EXTRACTO_DE_PROCEDENCIA,
    )
    add(
        "condiciones",
        "¿Cómo se resuelven las controversias en el contrato de proyecto y dirección?",
        ["E90-M-CTR-03"],
        [
            "Primero se prevé una mediación.",
            "Si la controversia persiste, se prevé una pericia arbitral.",
        ],
        [ref("E90-M-CTR-03", "CLAUSULA-8")],
    )
    add(
        "condiciones",
        f"En el contrato con {p2}, ¿de cuántos días dispone el cliente para revisar un entregable sujeto a aceptación?",
        t2,
        ["El cliente dispone de un período de prueba de quince días."],
        [ref(ctr2, "CLAUSULA-3")],
    )
    add(
        "condiciones",
        "¿Qué información queda excluida de la obligación de reserva?",
        ["E90-M-CTR-01"],
        [
            "La información ya pública.",
            "La recibida legítimamente de un tercero.",
            "La desarrollada de manera independiente.",
        ],
        [ref("E90-M-CTR-01", "CLAUSULA-6")],
    )
    c3 = registros[pre3]["condiciones"]
    add(
        "condiciones",
        f"¿Por cuánto tiempo es válida la oferta de {p3} y cómo se paga?",
        t3,
        [
            f"La oferta es válida por {c3['validez_dias']} días corridos.",
            f"La forma de pago es {c3['forma_pago']}.",
        ],
        [ref(pre3, "COND-validez"), ref(pre3, "COND-forma_pago")],
    )
    c1 = registros[oc1]["condiciones"]
    add(
        "condiciones",
        f"¿Dónde y en qué fecha debe entregarse lo pedido a {p1}?",
        t1,
        [
            f"Debe entregarse en {c1['lugar_entrega']}.",
            f"La fecha de entrega comprometida es el {fecha_es(c1['fecha_entrega'])}.",
        ],
        [ref(oc1, "COND-lugar"), ref(oc1, "COND-fecha_entrega")],
    )
    add(
        "condiciones",
        f"¿Cuándo vence el pago de la factura de {p2}?",
        t2,
        [
            f"El pago vence el {fecha_es(registros[fac2]['condiciones']['vencimiento'])}."
        ],
        [ref(fac2, "COND-vencimiento")],
    )

    # Multidocumento (6): comparación o síntesis entre documentos del trámite.
    add(
        "multidocumento",
        f"¿El importe de la orden de compra coincide con el del presupuesto de {p1}?",
        t1,
        [
            "Sí, los importes coinciden.",
            f"El presupuesto y la orden de compra indican {monto(pre1)}.",
        ],
        [ref(pre1, "CAMPO-importe_total"), ref(oc1, "CAMPO-importe_total")],
    )
    diferencia = Decimal(v(fac3, "importe_total")) - Decimal(v(oc3, "importe_total"))
    add(
        "multidocumento",
        f"¿La factura de {p3} se emitió por el mismo importe que la orden de compra?",
        t3,
        [
            "No, los importes no coinciden.",
            f"La orden de compra indica {monto(oc3)}.",
            f"La factura indica {monto(fac3)}.",
            f"La factura supera a la orden en {importe_es(str(diferencia))} {v(fac3, 'moneda')}.",
        ],
        [ref(oc3, "CAMPO-importe_total"), ref(fac3, "CAMPO-importe_total")],
    )
    add(
        "multidocumento",
        f"¿Qué documento previo cita la factura de {p1} y forma parte del trámite?",
        t1,
        [
            f"La factura cita la orden de compra {v(oc1, 'numero')}.",
            "Esa orden de compra forma parte del trámite.",
        ],
        [ref(fac1, "REF"), ref(oc1, "CAMPO-numero")],
    )
    add(
        "multidocumento",
        f"¿La factura de {p2} se emitió dentro de la vigencia del contrato?",
        t2,
        [
            "Sí, se emitió dentro de la vigencia.",
            f"La factura es del {fecha_es(v(fac2, 'fecha_emision'))}.",
            f"El contrato rige desde el {fecha_es(v(ctr2, 'vigencia_desde'))} hasta el {fecha_es(v(ctr2, 'vigencia_hasta'))}.",
        ],
        [
            ref(fac2, "CAMPO-fecha_emision"),
            ref(ctr2, "CAMPO-vigencia_desde"),
            ref(ctr2, "CAMPO-vigencia_hasta"),
        ],
    )
    add(
        "multidocumento",
        "¿Qué factura cita la constancia de conformidad y coincide su importe con el de esa factura?",
        t4,
        [
            f"La constancia cita la factura {v(fac4, 'numero')}.",
            f"Sí, ambos documentos indican {monto(com4)}.",
        ],
        [
            ref(com4, "OBJETO"),
            ref(com4, "CAMPO-importe"),
            ref(fac4, "CAMPO-numero"),
            ref(fac4, "CAMPO-importe_total"),
        ],
    )
    dias = (
        date.fromisoformat(v(oc3, "fecha")) - date.fromisoformat(v(pre3, "fecha"))
    ).days
    add(
        "multidocumento",
        f"¿Cuántos días pasaron entre el presupuesto y la orden de compra de {p3}?",
        t3,
        [
            f"El presupuesto es del {fecha_es(v(pre3, 'fecha'))}.",
            f"La orden de compra es del {fecha_es(v(oc3, 'fecha'))}.",
            f"Pasaron {dias} días.",
        ],
        [ref(pre3, "CAMPO-fecha"), ref(oc3, "CAMPO-fecha")],
    )

    # Sin evidencia (6): cuatro trámites y dos documentos sueltos, sin pistas en la pregunta.
    for contexto, pregunta, motivo in (
        (
            t1,
            f"¿Qué penalidad corresponde si {p1} entrega con demora?",
            "Los documentos fijan lugar y fecha de entrega, pero no establecen penalidades por demora.",
        ),
        (
            t2,
            f"¿Cuál es el monto de la multa por rescisión anticipada del contrato con {p2}?",
            "El contrato no fija una multa por rescisión anticipada.",
        ),
        (
            t3,
            f"¿A qué cuenta bancaria hay que transferir el pago de la factura de {p3}?",
            "La factura indica la forma de pago, pero no informa un número de cuenta.",
        ),
        (
            t4,
            f"¿Qué número de remito acompañó la entrega de los servicios de {p4}?",
            "Ningún documento del trámite menciona un remito.",
        ),
        (
            ["E90-M-CTR-09"],
            "¿Qué garantía de cumplimiento en dinero debe constituir el prestador?",
            "El contrato no exige una garantía de cumplimiento en dinero.",
        ),
        (
            ["E90-M-CTR-03"],
            "¿Cuál es el presupuesto de construcción de la obra?",
            "El monto pactado corresponde solo a honorarios; el contrato no informa el costo de construcción.",
        ),
    ):
        add("sin_evidencia", pregunta, contexto, [], [], motivo)
    if len(result) != 30:
        raise ValueError("recuento_consultas")
    escribir_json(destino, {"consultas": result})
    print("consultas: 30 (10 directas, 8 de condiciones, 6 multidocumento, 6 sin evidencia)")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--raiz", type=Path, default=Path(__file__).resolve().parents[2]
    )
    args = parser.parse_args()
    preparar_consultas(args.raiz.resolve())


if __name__ == "__main__":
    main()
