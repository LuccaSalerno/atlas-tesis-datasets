#!/usr/bin/env python3
"""Genera los DOCX de la cadena de evidencia (contrato, órdenes y facturas en dos versiones).

Datos ficticios; el CUIT 30-96472518-3 no existe en ARCA.
Ejecución: python generadores/cadena_evidencia.py; luego exportar cada DOCX a PDF digital.
"""
from pathlib import Path
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT/'demo'/'cadena_evidencia'/'documentos'
OUT.mkdir(parents=True,exist_ok=True)
CUIT = '30-96472518-3'  # Candidato únicamente: validación oficial ARCA pendiente
MARCA = 'DOCUMENTO FICTICIO - SIN VALIDEZ COMERCIAL NI FISCAL'
PROVEEDOR = 'Servicios Integrales del Sur S.A.'
CLIENTE = 'Laboratorio ATLAS de Procesos S.A.'


def doc_base(titulo, code):
    d=Document(); sec=d.sections[0]
    sec.top_margin=Cm(2.4); sec.bottom_margin=Cm(2.3)
    sec.left_margin=Cm(2.55); sec.right_margin=Cm(2.55)
    norm=d.styles['Normal']; norm.font.name='Liberation Sans'; norm.font.size=Pt(10.4)
    norm.paragraph_format.space_after=Pt(8); norm.paragraph_format.line_spacing=1.15
    for style_name,size in [('Title',18),('Heading 1',13),('Heading 2',11.5)]:
        s=d.styles[style_name]; s.font.name='Liberation Sans'; s.font.size=Pt(size); s.font.bold=True
        s.font.color.rgb=RGBColor(27,47,70)
    header=sec.header.paragraphs[0]
    header.text=f'ATLAS | CORPUS SINTÉTICO | {code}'
    header.style='Caption'; header.alignment=WD_ALIGN_PARAGRAPH.RIGHT
    footer=sec.footer.paragraphs[0]
    footer.alignment=WD_ALIGN_PARAGRAPH.CENTER
    run=footer.add_run(MARCA+'  |  Página ')
    run.font.size=Pt(7.4); run.font.color.rgb=RGBColor(149,40,40)
    fld=OxmlElement('w:fldSimple'); fld.set(qn('w:instr'),'PAGE'); footer._p.append(fld)
    label=d.add_paragraph(); label.alignment=WD_ALIGN_PARAGRAPH.CENTER
    rr=label.add_run(MARCA); rr.bold=True; rr.font.size=Pt(10); rr.font.color.rgb=RGBColor(154,37,37)
    d.add_heading(titulo,0)
    d.add_paragraph(f'Identificador documental: {code}  |  Versión: 1.0  |  Uso exclusivo: demostración y pruebas académicas.')
    return d

def P(d,t): d.add_paragraph(t)
def H(d,t): d.add_heading(t,level=1)
def H2(d,t): d.add_heading(t,level=2)
def bullet(d,t): d.add_paragraph(t, style='List Bullet')
def page(d,heading): d.add_page_break(); H(d,heading)
def tbl(d,items):
    tab=d.add_table(rows=0, cols=2);tab.style='Table Grid';tab.alignment=WD_TABLE_ALIGNMENT.CENTER
    for key,value in items:
        cells=tab.add_row().cells;cells[0].text=key;cells[1].text=value
    return tab

# CONTRATO: 10 páginas construidas explícitamente. Información de oráculo localizada por página.
d=doc_base('Contrato marco de mantenimiento y soporte documental','CTR-ATLAS-2026-001 / C-v1')
H(d,'Carátula y celebración')
P(d,'En la Ciudad Autónoma de Buenos Aires, a los 20 días del mes de febrero de 2026, se confecciona exclusivamente con fines académicos este documento de prueba entre las partes ficticias que se identifican a continuación. No existe manifestación real de consentimiento, contratación, facturación ni representación societaria.')
tbl(d,[('Parte requirente',CLIENTE+' (entidad completamente ficticia)'),('Parte prestadora / contraparte',PROVEEDOR+' (entidad completamente ficticia)'),('CUIT de la contraparte (dato de prueba)',CUIT),('Fecha de celebración','20/02/2026'),('Número de contrato','CTR-ATLAS-2026-001'),('Naturaleza','Instrumento sintético no suscrito, sin efectos jurídicos')])
P(d,'La identificación tributaria incorporada es un candidato generado para ensayos de extracción, coincidencia entre documentos y validación de campos. Su dígito verificador es correcto; la comprobación de inexistencia de coincidencias conocidas en la consulta pública de ARCA debe completarse antes de congelar el corpus.')
H2(d,'Estructura')
P(d,'Este texto contiene definiciones, objeto, obligaciones recíprocas, régimen económico, período de prestación, seguimiento, reserva, incidencias, resolución y anexo operativo. Las páginas sirven como procedencia documental para ensayar localizadores de extracción y consultas con cita de origen.')

page(d,'1. Definiciones y criterios de interpretación')
P(d,'A los efectos de este caso ficticio, se denomina servicio a las prestaciones de mantenimiento preventivo, soporte correctivo y elaboración de reportes descritas en el objeto. Se denomina requerimiento al pedido técnico registrado por la parte requirente, evidencia al registro documental de la atención efectuada y orden de compra al documento administrativo que autoriza una provisión vinculada a este contrato.')
P(d,'Las partes ficticias acuerdan utilizar un único identificador de referencia para correlacionar tickets, actas de conformidad y órdenes de compra. La correlación administrativa no altera por sí sola ni el alcance ni el período de prestación. Toda aceptación requiere revisión y conformidad humana según el flujo operativo de la parte requirente.')
H2(d,'1.1. Orden documental')
P(d,'La documentación relacionada comprende el contrato, las órdenes de compra, los comprobantes de prestación y las facturas de prueba. Cada instrumento conserva su identidad y puede tener versiones sucesivas. Ante una inconsistencia, se examinan los documentos originales, sus fechas de emisión y su autorización, en lugar de suponer que una extracción automática resuelve el conflicto.')
H2(d,'1.2. Interpretación de la prueba')
P(d,'No se incorporan cláusulas especiales provenientes de contratos reales. Los términos fueron redactados como material de ensayo y no constituyen una recomendación profesional sobre contratación. Los importes y los identificadores contenidos en documentos asociados deben cotejarse mediante reglas explícitas y datos confirmados.')

page(d,'2. Objeto de los servicios')
P(d,'La prestación representada incluye inspección preventiva de equipos informáticos de oficina, relevamiento de incidencias técnicas, soporte de segundo nivel, revisión de respaldos y elaboración de un reporte de mantenimiento. El alcance se limita al parque de equipos ficticio detallado en el anexo y no comprende suministro de hardware, enlaces de telecomunicaciones ni desarrollo de sistemas a medida.')
H2(d,'2.1. Prestaciones comprendidas')
for t in ['Diagnóstico inicial de treinta puestos de trabajo de laboratorio y elaboración del inventario técnico de ensayo.','Una revisión preventiva trimestral de estaciones de trabajo y recursos compartidos identificados por el cliente.','Soporte remoto para incidentes registrados por el referente operativo dentro del horario de atención indicado.','Informe mensual de actividades, hallazgos, tiempos de atención y recomendaciones sujetas a aprobación humana.','Registro de los cambios efectuados, su motivo y su relación con los requerimientos habilitados.']:
    bullet(d,t)
H2(d,'2.2. Exclusiones')
P(d,'No se autorizan intervenciones sobre entornos productivos reales ni operaciones con información confidencial. Las tareas documentadas constituyen ejemplos descriptivos. La clasificación automática, el reconocimiento de texto y la consulta asistida que ATLAS pueda aplicar a este archivo no representan obligaciones de la contraparte ficticia.')

page(d,'3. Obligaciones de la parte prestadora')
P(d,'La parte prestadora ficticia designaría un responsable técnico de coordinación, dispondría de un canal documentado para registrar incidentes y mantendría trazabilidad entre solicitud, actividad realizada y resultado informado. La ejecución de cualquier intervención estaría condicionada a las autorizaciones definidas por la parte requirente.')
H2(d,'3.1. Registro y comunicación')
P(d,'Cada incidente tendría un número interno, una fecha de apertura, una prioridad, la descripción del síntoma, la intervención propuesta y una constancia de cierre. Un cambio de prioridad debe quedar justificado y visible en el historial. Los datos necesarios para evaluar la prestación se mantendrían accesibles para las personas con autorización, sin difundir contenido fuera del ámbito correspondiente.')
H2(d,'3.2. Calidad y resguardo')
P(d,'La prestadora documentaría las tareas de diagnóstico, verificaría el funcionamiento posterior y elevaría las intervenciones que requieren conformidad previa. El incumplimiento de un objetivo de atención no se transforma automáticamente en autorización para ejecutar tareas adicionales. Las observaciones deben revisarse por una persona designada.')
P(d,'La parte prestadora no obtiene facultades para aprobar gastos, modificar órdenes de compra ni emitir decisiones documentales en nombre del cliente. Las solicitudes de modificación de alcance se documentan por separado y no modifican versiones históricas.')

page(d,'4. Condiciones económicas y forma de pago')
P(d,'El monto total de referencia del presente contrato marco sintético asciende a PESOS ARGENTINOS UN MILLÓN DOSCIENTOS CINCUENTA MIL ($ 1.250.000,00), con IVA incluido. La moneda pactada para todas las prestaciones comprendidas es ARS (pesos argentinos). Se trata de un importe nominal de prueba que no genera créditos, deudas ni obligaciones comerciales.')
tbl(d,[('Monto total','ARS 1.250.000,00'),('Moneda','ARS - pesos argentinos'),('Tratamiento impositivo','Importe total IVA incluido'),('Base neta de referencia','ARS 1.033.057,85'),('IVA de referencia (21 %)','ARS 216.942,15'),('Suma neto + IVA','ARS 1.250.000,00')])
H2(d,'4.1. Importe fijo y autorizaciones')
P(d,'El valor indicado corresponde al paquete completo de prestaciones y no debe interpretarse como una tarifa por hora. Las órdenes de compra documentan autorizaciones administrativas asociadas y pueden presentar inconsistencias de fecha o importe: dichas discrepancias se muestran como advertencias para revisión, sin sustituir el criterio del aprobador.')
H2(d,'4.2. Conformidad')
P(d,'En este escenario ficticio, la conformidad se registraría después de revisar el detalle de la prestación y de sus documentos vinculados. Una factura de prueba no equivale por sí sola a conformidad de servicio ni obliga al cierre de un trámite. No hay medios de pago, cuentas bancarias ni datos fiscales reales en este corpus.')

page(d,'5. Plazo de ejecución y vigencia')
P(d,'La vigencia del contrato se establece expresamente desde el 01/03/2026 (primero de marzo de dos mil veintiséis) hasta el 28/02/2027 (veintiocho de febrero de dos mil veintisiete), inclusive. Una orden de compra cuya fecha de emisión quede fuera de este intervalo debe generar una discrepancia verificable al compararse con el contrato vinculado, sin modificar automáticamente ningún documento.')
tbl(d,[('Inicio de vigencia','01/03/2026'),('Fin de vigencia','28/02/2027'),('Tratamiento de extremos','Ambas fechas inclusive'),('Prórroga','Solo mediante instrumento nuevo y confirmación humana')])
H2(d,'5.1. Calendario')
P(d,'Las revisiones preventivas se planifican al inicio de cada trimestre operativo, siempre dentro del período indicado. Las actividades reprogramadas conservan el historial del requerimiento original. La fecha de una orden de compra se interpreta como la fecha de su emisión, no como la fecha en que ATLAS procesa o incorpora el archivo.')
H2(d,'5.2. Modificaciones')
P(d,'Si cambiara el período de vigencia, se elaboraría una nueva versión documental y se solicitaría el circuito de autorización correspondiente. Una sugerencia producida por IA no modifica las fechas validadas y una decisión tomada sobre una versión anterior no se extiende automáticamente a las nuevas condiciones.')

page(d,'6. Seguimiento y evidencia de cumplimiento')
P(d,'El seguimiento se organiza mediante reportes de mantenimiento, constancias de tareas y observaciones realizadas por el referente designado. Cada documento generado durante la ejecución conserva el identificador del requerimiento al que corresponde y una referencia a la intervención efectuada.')
H2(d,'6.1. Indicadores operativos de ejemplo')
P(d,'El reporte puede incluir cantidad de incidentes, categoría, fecha de inicio, fecha de resolución y descripción de la intervención. Los indicadores no sustituyen el acta de aceptación. Toda corrección posterior debe preservar el dato anterior y la persona que autorizó el cambio, permitiendo reconstruir la secuencia.')
H2(d,'6.2. Tratamiento de observaciones')
P(d,'Ante un desacuerdo sobre la calidad de una tarea, el referente registra el motivo y solicita aclaraciones. Una observación no exige por sí misma la creación de una nueva versión de contrato: solo si cambia el contenido de un documento se incorpora una versión nueva, conforme al circuito definido por la parte requirente.')
P(d,'La información generada durante el servicio queda vinculada a la solicitud que la originó. No se confunde un indicador técnico con una decisión administrativa ni con una evidencia de validez fiscal de la documentación.')

page(d,'7. Confidencialidad y protección de información')
P(d,'Las partes del escenario declaran que cualquier material al que se diera acceso para una prestación real debería tratarse conforme a los permisos otorgados, con uso limitado a la finalidad autorizada. En esta instancia académica no se entregan ni procesan datos empresariales reales, información personal de terceros ni documentación contractual vigente.')
H2(d,'7.1. Acceso restringido')
P(d,'Los registros, archivos y fragmentos utilizados para pruebas se mantienen dentro de los entornos de demostración o evaluación expresamente autorizados. El personal operativo no debe acceder al contenido documental fuera de las responsabilidades asignadas. Las credenciales y los secretos técnicos no forman parte de los documentos.')
H2(d,'7.2. Tratamiento de incidencias')
P(d,'Una incidencia de acceso o divulgación se documentaría con fecha, naturaleza, recursos afectados y acciones de contención. La conservación y eventual eliminación técnica de datos dependerían de las políticas y autorizaciones del entorno; esta cláusula no pretende establecer un plazo legal universal de conservación.')
H2(d,'7.3. Uso académico')
P(d,'El documento representa exclusivamente texto de entrada para procesos de OCR, extracción y consulta asistida. No autoriza su reutilización como plantilla contractual lista para firma ni justifica atribuir validez jurídica a firmas simuladas que pudieran mostrarse en una aplicación.')

page(d,'8. Incumplimientos, rescisión y cierre')
P(d,'En una relación real, cualquier rescisión debería sujetarse al régimen contractual y normativo aplicable. Para este escenario de ficción, se registra que la parte que detecta un incumplimiento comunicaría los hechos, identificaría las prestaciones involucradas y habilitaría un plazo de respuesta antes de proponer el cierre de la relación.')
H2(d,'8.1. Causales de ejemplo')
for t in ['Incumplimiento reiterado de actividades expresamente comprendidas en el objeto.','Imposibilidad técnica persistente, acreditada con evidencia y comunicada a la contraparte.','Acuerdo documentado de ambas partes ficticias para suspender nuevas tareas.','Incumplimiento de controles de acceso o confidencialidad durante la ejecución.']:
    bullet(d,t)
H2(d,'8.2. Efectos documentales')
P(d,'El cierre administrativo exige distinguir documentos vigentes, observaciones pendientes y decisiones efectivamente registradas. Ninguna advertencia automatizada produce por sí sola la rescisión. Las versiones anteriores, sus registros y el fundamento de cada decisión se preservan para mantener trazabilidad histórica.')
H2(d,'8.3. Ausencia de firma')
P(d,'Este instrumento no fue suscrito por ninguna persona ni entidad real. Los espacios de aprobación que pudiera representar una interfaz de demostración corresponden a firmas simuladas y no expresan consentimiento jurídicamente vinculante.')

page(d,'Anexo A. Inventario y protocolo de aceptación sintéticos')
P(d,'El anexo describe activos y actividades de laboratorio enteramente ficticios, concebidos para ilustrar el tipo de información que acompañaría a un contrato de mantenimiento. No se identifica domicilio real, responsable identificable ni infraestructura de terceros.')
tbl(d,[('Parque de ensayo','30 estaciones de trabajo ficticias'),('Recursos compartidos','2 impresoras y 1 repositorio documental simulado'),('Gestión de solicitudes','Tickets con identificador y secuencia temporal'),('Reporte esperado','Detalle de actividad, observaciones y referencia al requerimiento'),('Aceptación','Revisión y confirmación por referente humano autorizado')])
H2(d,'A.1. Control de documentación')
P(d,'Para la demostración se prevé vincular este contrato con órdenes de compra y facturas también sintéticas. Los documentos vinculados registran sus propios números, fechas, importes y versiones. Las reglas comparan únicamente datos validados de documentos confirmados como parte del mismo trámite.')
H2(d,'A.2. Identificación y cierre del instrumento de prueba')
P(d,'Código de corpus: C-v1. Número de referencia: CTR-ATLAS-2026-001. Estado del documento: material sintético de prueba, sin firmas, sin efecto comercial y sin validez fiscal. Esta copia puede utilizarse para ensayo de clasificación y extracción, pero su identificación tributaria candidata debe cotejarse en ARCA antes de la congelación del corpus.')
P(d,'Fin del documento C-v1. Toda modificación posterior del contenido deberá dar lugar a una nueva versión controlada y a un nuevo valor SHA-256.')
d.save(OUT/'C-v1_Contrato_Mantenimiento_ATLAS.docx')

# Cuatro documentos relacionados: se usan en las Etapas 1.7+ y EVD-01; no son documentos fiscales.
# Número documental estable entre versiones del mismo documento (DO-2026-10-23-CIERRE D4,
# DIF-84): RF-06 relaciona O-v2 con O-v1 y F-v2 con F-v1 por su número.
NUMERO = {'Orden_Compra': 'OC-2026-0042', 'Factura': 'FC-A-0001-00000123'}

def auxiliary(kind, code, title, date, amount, extra=None):
    d=doc_base(title,code)
    P(d,f'Fecha de emisión: {date}. Todos los sujetos, importes y referencias de este documento son ficticios. No constituye un comprobante fiscal, una autorización de compra ni una instrucción de pago.')
    tbl(d,[('Número',NUMERO[kind]),('Proveedor / emisor',PROVEEDOR),('CUIT proveedor / emisor (dato de prueba)',CUIT),('Cliente / destinatario',CLIENTE),('Fecha',date),('Importe total IVA incluido',f'ARS {amount}'),('Moneda','ARS - pesos argentinos'),('Contrato de referencia','CTR-ATLAS-2026-001')])
    H(d,'Detalle del concepto')
    P(d,'Prestaciones de mantenimiento preventivo y soporte documental de laboratorio, vinculadas al contrato sintético de referencia. El número y los datos se incluyen solo para verificar trazabilidad documental, control humano y reglas de consistencia.')
    H2(d,'Desglose impositivo de referencia')
    amountnum=int(amount.replace('.','').replace(',00',''))
    from decimal import Decimal, ROUND_HALF_UP
    net=(Decimal(amountnum)/Decimal('1.21')).quantize(Decimal('0.01'),rounding=ROUND_HALF_UP)
    vat=Decimal(amountnum)-net
    fmt=lambda x:f'{x:,.2f}'.replace(',','X').replace('.',',').replace('X','.')
    P(d,f'Total de prueba IVA incluido: ARS {fmt(Decimal(amountnum))}. Neto de referencia: ARS {fmt(net)}; IVA de referencia (21 %): ARS {fmt(vat)}. No hay CAE, punto de venta fiscal, QR fiscal, firma ni datos bancarios.')
    if extra: P(d,extra)
    P(d,'Referencia administrativa no vinculante. Los datos aquí contenidos deberán confirmarse mediante una validación humana antes de usarse como entradas de reglas de consistencia.')
    d.save(OUT/(code.replace('/','_')+'_'+kind+'.docx'))

auxiliary('Orden_Compra','O-v1','Orden de compra sintética (primera versión)','15/03/2027','1.250.000,00','Esta fecha de emisión está FUERA de la vigencia establecida por el contrato de referencia; se espera una advertencia de la regla R3.')
auxiliary('Orden_Compra','O-v2','Orden de compra sintética (versión corregida)','15/09/2026','1.250.000,00','Esta versión corrige únicamente la fecha de emisión respecto de O-v1. La fecha está DENTRO de la vigencia contractual.')
auxiliary('Factura','F-v1','Factura ficticia de ensayo (primera versión)','30/09/2026','1.312.500,00','El total supera en ARS 62.500 el total de la orden vinculada, por lo que se espera una advertencia de la regla R4.')
auxiliary('Factura','F-v2','Factura ficticia de ensayo (versión corregida)','30/09/2026','1.250.000,00','El total corregido coincide con el importe de la orden vinculada. Esta pieza no constituye una factura fiscal real.')
print('Fuentes DOCX generadas en',OUT)
