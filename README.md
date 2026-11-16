# ATLAS · Datasets de prueba y demostración

Proyecto Final de Ingeniería de Ivo Gonsé y Lucca Salerno (UADE, 2026).

**No se entrenó ni se hizo fine-tuning de ningún modelo.** ATLAS usa modelos preentrenados y estos documentos se generaron únicamente para probar y demostrar el sistema. Todos son **sintéticos**, con datos ficticios: usar documentos reales de empresas implicaba manejar información confidencial, y al generarlos sabíamos de antemano el valor correcto de cada campo.

| Carpeta | Uso | Documentos |
|---|---|---:|
| `documentos/medicion/` | Medición de clasificación, extracción y consulta asistida (Sección 6.3.1) | 50 |
| `documentos/ajuste/` | Ajuste de instrucciones, separado de la medición | 15 |
| `demo/cadena_evidencia/` | Pruebas de versiones, vínculos, reglas de consistencia y cierre (Sección 6.2) | 5 |
| `demo/surdelta_videos/` | Escenario de los videos y la defensa | 12 |

**Corpus de evaluación.** Diez documentos de medición por tipo (contrato, presupuesto, orden de compra, factura y comprobante), en PDF digital, DOCX, PDF escaneado y fotografiado (degradación simulada), con distintas plantillas y formatos. Incluye casos difíciles: campos ausentes, escaneos de baja calidad, clasificación ambigua, contratos extensos y tres documentos con instrucciones maliciosas (*prompt injection*). Los nombres `doc-xxxxxxxx` son neutros para no revelar el tipo, que es lo que el sistema tiene que clasificar.

**Demostración.** No se usaron para métricas. Están armados para provocar advertencias de las reglas de consistencia (fechas fuera de vigencia, facturas que superan la orden, CUIT que no coinciden) y corregirlas con nuevas versiones.

Los contratos de evaluación adaptan modelos públicos con licencia abierta. Cada documento cita su fuente y los cambios realizados (ver `ATRIBUCIONES.md`).
