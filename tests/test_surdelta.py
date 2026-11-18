"""Demo SurDelta: los documentos coinciden con su oráculo y arman los conflictos de la historia.

También comprueba que los CUIT sean el de la cadena de evidencia o los de ajuste del corpus de
evaluación (nunca los de medición).
"""

import json
import re
import unittest
from decimal import Decimal
from pathlib import Path

from pypdf import PdfReader

from generadores.evaluacion.comun import cuit_valido

RAIZ = Path(__file__).resolve().parents[1]
ORACULO = json.loads((RAIZ / "demo/surdelta_videos/oraculo.json").read_text(encoding="utf-8"))
CORPUS = RAIZ / "demo/surdelta_videos/documentos"


def texto(archivo: str) -> str:
    return " ".join(p.extract_text() for p in PdfReader(CORPUS / archivo).pages)


def pesos(valor: str) -> str:
    entero, decimales = f"{Decimal(valor):.2f}".split(".")
    return f"{int(entero):,}".replace(",", ".") + "," + decimales


def fecha(iso: str) -> str:
    anio, mes, dia = iso.split("-")
    return f"{dia}/{mes}/{anio}"


class CorpusSurDeltaTests(unittest.TestCase):
    def campos(self, codigo: str) -> dict[str, str]:
        return ORACULO["documentos"][codigo]["campos"]

    def test_un_archivo_por_documento_del_oraculo(self) -> None:
        self.assertEqual(
            {d["archivo"] for d in ORACULO["documentos"].values()},
            {p.name for p in CORPUS.iterdir()},
        )

    def test_cuit_de_la_cadena_de_evidencia_o_de_ajuste(self) -> None:
        cadena = json.loads((RAIZ / "demo/cadena_evidencia/oraculo.json").read_text(encoding="utf-8"))
        registros = json.loads((RAIZ / "oraculos/registros.json").read_text(encoding="utf-8"))
        ajuste = {
            v
            for r in registros["registros"]
            if r["grupo"] == "ajuste"
            for k, v in r["valores"].items()
            if k.startswith("cuit")
        }
        permitidos = ajuste | {cadena["registro_cuit"]["numero"]}
        for codigo, doc in ORACULO["documentos"].items():
            for campo, valor in doc["campos"].items():
                if campo.startswith("cuit"):
                    self.assertTrue(cuit_valido(valor), codigo)
                    self.assertIn(valor, permitidos, codigo)

    def test_cada_dato_figura_en_su_documento_y_sin_texto_de_control(self) -> None:
        for codigo, doc in ORACULO["documentos"].items():
            contenido = re.sub(r"\s+", " ", texto(doc["archivo"]))
            for campo, valor in doc["campos"].items():
                if campo in ("monto_total", "importe_total"):
                    valor = pesos(valor)
                elif re.fullmatch(r"\d{4}-\d{2}-\d{2}", valor):
                    valor = fecha(valor)
                elif campo == "moneda":
                    valor = "(ARS)"
                self.assertIn(valor, contenido, f"{codigo}.{campo}")
            for prohibido in ("regla", "R3", "R4", "se espera", "advertencia", "oráculo"):
                self.assertNotIn(prohibido, contenido, codigo)

    def test_conflictos_de_la_historia(self) -> None:
        contrato, orden = self.campos("SD-C1"), self.campos("SD-O1")
        v1, v2 = self.campos("SD-F1"), self.campos("SD-F2")
        # R4: la v1 supera la orden y la v2 no; R3: la orden es anterior a la vigencia.
        self.assertGreater(Decimal(v1["importe_total"]), Decimal(orden["importe_total"]))
        self.assertEqual(Decimal(v2["importe_total"]), Decimal(orden["importe_total"]))
        self.assertLess(orden["fecha"], contrato["vigencia_desde"])
        # RF-06: la v2 tiene otro número y el mismo CUIT del emisor que la v1.
        self.assertNotEqual(v1["numero"], v2["numero"])
        self.assertEqual(v1["cuit_emisor"], v2["cuit_emisor"])
        self.assertEqual(v1["cuit_emisor"], orden["cuit_proveedor"])
        self.assertEqual(orden["cuit_proveedor"], contrato["cuit_contraparte"])
        # Compra: la factura llega con el CUIT de otra empresa (R1).
        self.assertNotEqual(
            self.campos("SD-F3")["cuit_emisor"], self.campos("SD-O2")["cuit_proveedor"]
        )


if __name__ == "__main__":
    unittest.main()
