# ATLAS · Datasets de prueba y demostración

Proyecto Final de Ingeniería de Ivo Gonsé y Lucca Salerno (UADE, 2026).

**No se entrenó ni se hizo fine-tuning de ningún modelo.** ATLAS usa modelos preentrenados y estos documentos se generaron únicamente para probar y demostrar el sistema. Todos son **sintéticos**, con datos ficticios: usar documentos reales de empresas implicaba manejar información confidencial, y al generarlos sabíamos de antemano el valor correcto de cada campo.

| Carpeta | Contenido |
|---|---|
| `documentos/medicion/` (50) | Corpus de evaluación: clasificación, extracción y consulta asistida (Sección 6.3.1) |
| `documentos/ajuste/` (15) | Ajuste de instrucciones, separado de la medición |
| `oraculos/` | Valores correctos del corpus de evaluación |
| `demo/cadena_evidencia/` (5) | Versiones, vínculos, reglas de consistencia y cierre (Sección 6.2), con su `oraculo.json` |
| `demo/surdelta_videos/` (12) | Escenario de los videos y la defensa, con su `oraculo.json` |
| `generadores/` | Código que genera todo lo anterior |

**Corpus de evaluación.** Diez documentos de medición por tipo (contrato, presupuesto, orden de compra, factura y comprobante), en PDF digital, DOCX, PDF escaneado y fotografiado (degradación simulada), con distintas plantillas y formatos. Incluye campos ausentes, escaneos de baja calidad, clasificación ambigua, contratos extensos, cuatro trámites cuyos documentos se relacionan entre sí y tres documentos con instrucciones maliciosas (*prompt injection*). Los nombres `doc-xxxxxxxx` no revelan el tipo, que es lo que el sistema tiene que clasificar.

**Oráculos** (`oraculos/`):

- `inventario.csv`: cada documento con su ID (`E90-M-CTR-01`, …), partición, tipo, formato, plantilla y ruta.
- `registros.json`: los datos con los que se generó cada documento y los trámites que los agrupan.
- `campos.json`: el valor correcto de cada campo y dónde aparece (página o párrafo).
- `fragmentos.json`: los fragmentos de texto que se pueden citar.
- `consultas.json`: 30 preguntas de RAG con su respuesta esperada y los fragmentos que la respaldan.

**Demostración.** No se usaron para métricas. Están armados para provocar advertencias de las reglas de consistencia (fechas fuera de vigencia, facturas que superan la orden, CUIT que no coinciden) y corregirlas con nuevas versiones.

Los CUIT y las razones sociales son inventados; el equipo los consultó en ARCA y no existen. Los contratos adaptan modelos públicos con licencia abierta: ver `ATRIBUCIONES.md`.

## Uso

```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
.venv/bin/python -m pytest -q
```

## Generar el corpus de evaluación

Requiere LibreOffice y pdftoppm. Se genera en una carpeta vacía y en tres pasos:

```bash
python -m generadores.evaluacion.preparar --raiz /ruta/nueva     # registros.json
python -m generadores.evaluacion.generar --raiz /ruta/nueva      # documentos, inventario, campos y fragmentos
python -m generadores.evaluacion.consultas --raiz /ruta/nueva    # consultas.json
python -m generadores.evaluacion.validar --raiz /ruta/nueva
```

Los datos de entrada están en `generadores/evaluacion/plan.csv` (qué documento va en cada partición, con qué formato y plantilla) y `generadores/evaluacion/contratos.json` (las cláusulas de los modelos contractuales). La generación reproduce los oráculos de este repositorio exactamente; los PDF pueden cambiar en bytes según la versión de LibreOffice, pero no en su texto.

Las demos se generan con `generadores/cadena_evidencia.py` y `generadores/surdelta.py`.
