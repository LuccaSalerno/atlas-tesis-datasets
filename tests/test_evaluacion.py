"""Corpus de evaluación: coherencia de documentos y oráculos, y generación reproducible."""

import copy
import tempfile
import unittest
from pathlib import Path

from generadores.evaluacion.comun import leer_json
from generadores.evaluacion.preparar import preparar
from generadores.evaluacion.validar import leer, validar, validar_estructura

RAIZ = Path(__file__).resolve().parents[1]


class CorpusEvaluacionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.datos = leer(RAIZ)

    def test_el_corpus_valida(self):
        self.assertEqual(validar(RAIZ), [])

    def test_preparar_reproduce_los_registros(self):
        with tempfile.TemporaryDirectory() as tmp:
            preparar(Path(tmp))
            self.assertEqual(
                leer_json(Path(tmp) / "oraculos/registros.json"), self.datos["registros"]
            )

    def errores_con(self, cambio):
        datos = copy.deepcopy(self.datos)
        cambio(datos)
        return validar_estructura(datos)

    def test_detecta_un_valor_que_no_coincide_con_el_oraculo(self):
        def cambio(d):
            d["campos"]["campos"][0]["valor_esperado"] = "otro"

        self.assertIn("CAMPO_DISCREPANTE:E90-M-CTR-01", self.errores_con(cambio))

    def test_detecta_una_cita_fuera_del_contexto(self):
        def cambio(d):
            q = next(q for q in d["consultas"]["consultas"] if q["categoria"] == "directa")
            q["documentos_autorizados"] = ["E90-M-FAC-01"]

        self.assertTrue(any(e.startswith("CITA_FUERA_DE_CONTEXTO") for e in self.errores_con(cambio)))

    def test_detecta_un_tramite_que_no_concilia(self):
        def cambio(d):
            r = next(r for r in d["registros"]["registros"] if r["doc_id"] == "E90-M-OC-01")
            r["valores"]["moneda"] = "USD"

        self.assertIn("TRAMITE_INCOHERENTE:E90-CONTEXTO-T1", self.errores_con(cambio))

    def test_detecta_un_adversarial_que_coincide_con_el_oraculo(self):
        def cambio(d):
            r = next(r for r in d["registros"]["registros"] if r["prompt_injection"])
            r["adversarial"]["objetivos"] = {"clasificacion": r["clasificacion_esperada"]}

        self.assertTrue(any(e.startswith("ADVERSARIAL_NO_CONTRASTABLE") for e in self.errores_con(cambio)))


if __name__ == "__main__":
    unittest.main()
