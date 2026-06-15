"""
Genera el manual de usuario del Validador Documental con temática UDEA.
Ejecutar: python scripts/generar_manual.py
Salida:   manual_validador_documental.pdf
"""

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    PageBreak, HRFlowable, KeepTogether,
)
from reportlab.pdfgen import canvas
from reportlab.platypus.flowables import Flowable
import io, os

# ── Paleta UDEA ───────────────────────────────────────────────────────────────
VERDE       = colors.HexColor("#006633")
VERDE_OSC   = colors.HexColor("#004d26")
VERDE_CLARO = colors.HexColor("#e8f4ee")
VERDE_MED   = colors.HexColor("#339966")
GRIS_OSC    = colors.HexColor("#333333")
GRIS_MED    = colors.HexColor("#666666")
GRIS_CLARO  = colors.HexColor("#f5f5f5")
BLANCO      = colors.white
DORADO      = colors.HexColor("#c8a84b")
ROJO        = colors.HexColor("#cc3333")
AZUL_INFO   = colors.HexColor("#1a6696")

W, H = A4

OUTPUT = "manual_validador_documental.pdf"

# ── Estilos ───────────────────────────────────────────────────────────────────
def estilos():
    return {
        "titulo_portada": ParagraphStyle(
            "titulo_portada", fontName="Helvetica-Bold",
            fontSize=26, textColor=BLANCO, alignment=TA_CENTER, leading=32,
        ),
        "sub_portada": ParagraphStyle(
            "sub_portada", fontName="Helvetica",
            fontSize=13, textColor=VERDE_CLARO, alignment=TA_CENTER, leading=18,
        ),
        "version": ParagraphStyle(
            "version", fontName="Helvetica",
            fontSize=10, textColor=DORADO, alignment=TA_CENTER,
        ),
        "h1": ParagraphStyle(
            "h1", fontName="Helvetica-Bold",
            fontSize=16, textColor=VERDE_OSC, spaceBefore=18, spaceAfter=8, leading=20,
        ),
        "h2": ParagraphStyle(
            "h2", fontName="Helvetica-Bold",
            fontSize=12, textColor=VERDE, spaceBefore=14, spaceAfter=6, leading=16,
        ),
        "h3": ParagraphStyle(
            "h3", fontName="Helvetica-Bold",
            fontSize=10, textColor=GRIS_OSC, spaceBefore=10, spaceAfter=4, leading=14,
        ),
        "body": ParagraphStyle(
            "body", fontName="Helvetica",
            fontSize=9.5, textColor=GRIS_OSC, leading=14, spaceAfter=6,
            alignment=TA_JUSTIFY,
        ),
        "code": ParagraphStyle(
            "code", fontName="Courier",
            fontSize=8.5, textColor=VERDE_OSC, backColor=VERDE_CLARO,
            leftIndent=10, rightIndent=10, leading=13, spaceBefore=4, spaceAfter=4,
        ),
        "bullet": ParagraphStyle(
            "bullet", fontName="Helvetica",
            fontSize=9.5, textColor=GRIS_OSC, leading=14,
            leftIndent=16, bulletIndent=6, spaceAfter=3,
        ),
        "caption": ParagraphStyle(
            "caption", fontName="Helvetica-Oblique",
            fontSize=8, textColor=GRIS_MED, alignment=TA_CENTER, spaceAfter=8,
        ),
        "toc": ParagraphStyle(
            "toc", fontName="Helvetica",
            fontSize=10, textColor=GRIS_OSC, leading=18,
        ),
        "toc_bold": ParagraphStyle(
            "toc_bold", fontName="Helvetica-Bold",
            fontSize=10, textColor=VERDE_OSC, leading=20,
        ),
        "aviso": ParagraphStyle(
            "aviso", fontName="Helvetica",
            fontSize=9, textColor=GRIS_OSC, leading=13,
            leftIndent=8,
        ),
    }

E = estilos()

# ── Flowable: caja coloreada ──────────────────────────────────────────────────
class CajaColor(Flowable):
    def __init__(self, texto, color_fondo=VERDE_CLARO, color_borde=VERDE,
                 color_texto=VERDE_OSC, icono="", altura=None):
        super().__init__()
        self.texto       = texto
        self.color_fondo = color_fondo
        self.color_borde = color_borde
        self.color_texto = color_texto
        self.icono       = icono
        self._altura     = altura or 0.85*cm
        self.width       = 0
        self.height      = self._altura

    def wrap(self, availWidth, availHeight):
        self.width = availWidth
        return availWidth, self._altura

    def draw(self):
        c = self.canv
        c.setFillColor(self.color_fondo)
        c.setStrokeColor(self.color_borde)
        c.roundRect(0, 0, self.width, self._altura, 5, fill=1, stroke=1)
        c.setFillColor(self.color_texto)
        c.setFont("Helvetica-Bold", 9)
        txt = f"{self.icono}  {self.texto}" if self.icono else self.texto
        c.drawString(10, self._altura / 2 - 4, txt)


class LineaVerde(Flowable):
    def __init__(self, grosor=2):
        super().__init__()
        self.grosor = grosor
        self.height = grosor + 4

    def wrap(self, aW, aH):
        self.width = aW
        return aW, self.height

    def draw(self):
        c = self.canv
        c.setStrokeColor(VERDE)
        c.setLineWidth(self.grosor)
        c.line(0, self.height / 2, self.width, self.height / 2)


# ── Header / Footer ───────────────────────────────────────────────────────────
def _encabezado_pie(canv, doc):
    canv.saveState()
    # Encabezado
    canv.setFillColor(VERDE_OSC)
    canv.rect(0, H - 1.1*cm, W, 1.1*cm, fill=1, stroke=0)
    canv.setFillColor(BLANCO)
    canv.setFont("Helvetica-Bold", 8)
    canv.drawString(1.5*cm, H - 0.72*cm, "VALIDADOR DOCUMENTAL SHAREPOINT")
    canv.setFont("Helvetica", 8)
    canv.drawRightString(W - 1.5*cm, H - 0.72*cm, "Manual de Usuario v2.0")

    # Franja dorada
    canv.setFillColor(DORADO)
    canv.rect(0, H - 1.25*cm, W, 0.15*cm, fill=1, stroke=0)

    # Pie
    canv.setFillColor(GRIS_CLARO)
    canv.rect(0, 0, W, 1.0*cm, fill=1, stroke=0)
    canv.setFillColor(VERDE)
    canv.rect(0, 0.95*cm, W, 0.05*cm, fill=1, stroke=0)
    canv.setFillColor(GRIS_MED)
    canv.setFont("Helvetica", 7.5)
    canv.drawString(1.5*cm, 0.32*cm, "Universidad de Antioquia — Programa de Fortalecimiento Empresarial")
    canv.setFont("Helvetica-Bold", 8)
    canv.drawRightString(W - 1.5*cm, 0.32*cm, f"Pag. {doc.page}")
    canv.restoreState()


def _portada(canv, doc):
    canv.saveState()
    # Fondo verde principal
    canv.setFillColor(VERDE_OSC)
    canv.rect(0, 0, W, H, fill=1, stroke=0)

    # Franja dorada superior
    canv.setFillColor(DORADO)
    canv.rect(0, H - 0.6*cm, W, 0.6*cm, fill=1, stroke=0)

    # Franja verde media decorativa
    canv.setFillColor(VERDE_MED)
    canv.rect(0, H * 0.42, W, 0.4*cm, fill=1, stroke=0)

    # Bloque blanco inferior decorativo
    canv.setFillColor(BLANCO)
    canv.setFillAlpha(0.06)
    canv.rect(0, 0, W, H * 0.18, fill=1, stroke=0)
    canv.setFillAlpha(1)

    # Franja dorada inferior
    canv.setFillColor(DORADO)
    canv.rect(0, 0.6*cm, W, 0.3*cm, fill=1, stroke=0)

    # Pie portada
    canv.setFillColor(VERDE_OSC)
    canv.rect(0, 0, W, 0.6*cm, fill=1, stroke=0)
    canv.setFillColor(DORADO)
    canv.setFont("Helvetica", 7.5)
    canv.drawCentredString(W / 2, 0.2*cm, "Universidad de Antioquia  |  2026")

    # Cuadro decorativo izquierdo
    canv.setFillColor(VERDE_MED)
    canv.setFillAlpha(0.3)
    canv.rect(0, H * 0.55, 0.8*cm, H * 0.35, fill=1, stroke=0)
    canv.setFillAlpha(1)

    # Cuadro decorativo derecho
    canv.setFillColor(VERDE_MED)
    canv.setFillAlpha(0.2)
    canv.rect(W - 0.8*cm, H * 0.55, 0.8*cm, H * 0.35, fill=1, stroke=0)
    canv.setFillAlpha(1)

    # Icono documento (rectángulos simulando archivos)
    cx, cy = W / 2, H * 0.70
    for i, (dx, dy, col) in enumerate([
        (-2.2*cm, 0, VERDE_MED), (0, 0.5*cm, BLANCO), (2.2*cm, 0, VERDE_MED)
    ]):
        canv.setFillColor(col)
        canv.setFillAlpha(0.85 if i == 1 else 0.5)
        canv.roundRect(cx + dx - 0.8*cm, cy + dy - 1.0*cm, 1.6*cm, 2.0*cm, 3, fill=1, stroke=0)
        canv.setFillColor(VERDE_OSC if i == 1 else VERDE_CLARO)
        canv.setFillAlpha(0.9)
        for ln in range(4):
            canv.rect(cx + dx - 0.5*cm, cy + dy + 0.35*cm - ln * 0.28*cm, 1.0*cm, 0.07*cm, fill=1, stroke=0)
        canv.setFillAlpha(1)

    # Checkmark verde
    canv.setFillColor(DORADO)
    canv.setFont("Helvetica-Bold", 28)
    canv.drawCentredString(W / 2, H * 0.68, "OK")

    canv.restoreState()


# ── Tabla utilitaria ──────────────────────────────────────────────────────────
def tabla(datos, col_widths, header=True, zebra=True):
    t = Table(datos, colWidths=col_widths, repeatRows=1 if header else 0)
    estilo = [
        ("FONTNAME",    (0, 0), (-1, -1), "Helvetica"),
        ("FONTSIZE",    (0, 0), (-1, -1), 8.5),
        ("ROWBACKGROUND", (0, 0), (-1, 0), VERDE),
        ("TEXTCOLOR",   (0, 0), (-1, 0), BLANCO),
        ("FONTNAME",    (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE",    (0, 0), (-1, 0), 9),
        ("ALIGN",       (0, 0), (-1, -1), "LEFT"),
        ("VALIGN",      (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING",  (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("GRID",        (0, 0), (-1, -1), 0.3, colors.HexColor("#cccccc")),
        ("ROWBACKGROUND", (0, 1), (-1, 1), VERDE_CLARO if zebra else BLANCO),
    ]
    if zebra:
        for i in range(2, len(datos), 2):
            estilo.append(("ROWBACKGROUND", (0, i), (-1, i), GRIS_CLARO))
    t.setStyle(TableStyle(estilo))
    return t


# ── Paso visual numerado ──────────────────────────────────────────────────────
def paso(numero, titulo, elementos):
    items = []
    # Badge circular con número
    badge_data = [[f"  {numero}  "]]
    badge = Table(badge_data, colWidths=[1.0*cm])
    badge.setStyle(TableStyle([
        ("BACKGROUND",    (0,0), (-1,-1), VERDE),
        ("TEXTCOLOR",     (0,0), (-1,-1), BLANCO),
        ("FONTNAME",      (0,0), (-1,-1), "Helvetica-Bold"),
        ("FONTSIZE",      (0,0), (-1,-1), 11),
        ("ALIGN",         (0,0), (-1,-1), "CENTER"),
        ("VALIGN",        (0,0), (-1,-1), "MIDDLE"),
        ("TOPPADDING",    (0,0), (-1,-1), 4),
        ("BOTTOMPADDING", (0,0), (-1,-1), 4),
        ("ROUNDEDCORNERS",(0,0), (-1,-1), [8,8,8,8]),
    ]))
    tit = Paragraph(f"<b>{titulo}</b>", ParagraphStyle(
        "paso_tit", fontName="Helvetica-Bold", fontSize=11,
        textColor=VERDE_OSC, leading=14,
    ))
    header_row = Table([[badge, tit]], colWidths=[1.3*cm, 13.5*cm])
    header_row.setStyle(TableStyle([
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("LEFTPADDING", (0,0), (-1,-1), 0),
        ("RIGHTPADDING", (0,0), (-1,-1), 0),
        ("TOPPADDING", (0,0), (-1,-1), 0),
        ("BOTTOMPADDING", (0,0), (-1,-1), 4),
    ]))
    items.append(header_row)
    items.extend(elementos)
    items.append(Spacer(1, 0.3*cm))

    # Envolver en un contenedor con borde verde
    contenedor = []
    contenedor.append(HRFlowable(width="100%", thickness=1.5, color=VERDE_MED,
                                  spaceAfter=6, spaceBefore=0))
    contenedor.extend(items)
    contenedor.append(HRFlowable(width="100%", thickness=0.5, color=VERDE_MED,
                                  spaceAfter=2, spaceBefore=4))
    return contenedor + [Spacer(1, 0.4*cm)]


# ── Caja de información / aviso ───────────────────────────────────────────────
def aviso(texto, tipo="info"):
    colores_tipo = {
        "info":    (colors.HexColor("#dbeeff"), AZUL_INFO,  "INFO"),
        "alerta":  (colors.HexColor("#fff3cd"), colors.HexColor("#856404"), "AVISO"),
        "error":   (colors.HexColor("#fde8e8"), ROJO,       "IMPORTANTE"),
        "ok":      (VERDE_CLARO,                VERDE,      "OK"),
    }
    fondo, borde, etiq = colores_tipo.get(tipo, colores_tipo["info"])
    contenido = Paragraph(f"<b>{etiq}:</b> {texto}", ParagraphStyle(
        "av", fontName="Helvetica", fontSize=9, textColor=GRIS_OSC,
        leading=13, leftIndent=4,
    ))
    t = Table([[contenido]], colWidths=[W - 4*cm])
    t.setStyle(TableStyle([
        ("BACKGROUND",    (0,0), (-1,-1), fondo),
        ("BOX",           (0,0), (-1,-1), 1.5, borde),
        ("LEFTPADDING",   (0,0), (-1,-1), 10),
        ("RIGHTPADDING",  (0,0), (-1,-1), 10),
        ("TOPPADDING",    (0,0), (-1,-1), 7),
        ("BOTTOMPADDING", (0,0), (-1,-1), 7),
    ]))
    return [t, Spacer(1, 0.3*cm)]


# ── CONTENIDO ────────────────────────────────────────────────────────────────
def build_story():
    story = []
    p = lambda txt, st="body": Paragraph(txt, E[st])
    sp = lambda h=0.3: Spacer(1, h*cm)

    # ── PORTADA ──────────────────────────────────────────────────────────────
    story.append(Spacer(1, 5.5*cm))
    story.append(p("VALIDADOR DOCUMENTAL", "titulo_portada"))
    story.append(sp(0.4))
    story.append(p("SharePoint · FastAPI · OpenAI", "sub_portada"))
    story.append(sp(1.5))
    story.append(p("Manual de Usuario — v2.0", "version"))
    story.append(sp(0.3))
    story.append(p("Programa de Fortalecimiento Empresarial", "version"))
    story.append(PageBreak())

    # ── TABLA DE CONTENIDO ───────────────────────────────────────────────────
    story.append(sp(0.5))
    story.append(p("Tabla de Contenido", "h1"))
    story.append(LineaVerde())
    story.append(sp(0.3))

    toc_items = [
        ("1.", "Descripcion General del Sistema",    "3"),
        ("2.", "Requisitos e Instalacion",           "4"),
        ("3.", "Arranque del Servidor",              "5"),
        ("4.", "Uso desde Swagger UI — Paso a Paso","5"),
        ("5.", "Flujos de Validacion (00 al 03)",    "8"),
        ("6.", "Estructura del Archivo de Entrada",  "10"),
        ("7.", "Checklist de Salida",                "11"),
        ("8.", "Ejecucion de Flujos Independientes", "12"),
        ("9.", "Filas Pendientes por Timeout",       "13"),
        ("10.","Variables de Entorno",               "14"),
        ("11.","Preguntas Frecuentes",               "15"),
    ]
    for num, titulo, pag in toc_items:
        fila = Table(
            [[p(f"<b>{num}</b>", "toc_bold"), p(titulo, "toc"),
              p(pag, "toc")]],
            colWidths=[1.0*cm, 12.5*cm, 1.5*cm],
        )
        fila.setStyle(TableStyle([
            ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
            ("LINEBELOW", (0,0), (-1,-1), 0.3, colors.HexColor("#dddddd")),
            ("TOPPADDING", (0,0), (-1,-1), 4),
            ("BOTTOMPADDING", (0,0), (-1,-1), 4),
            ("LEFTPADDING", (0,0), (0,0), 0),
            ("ALIGN", (2,0), (2,0), "RIGHT"),
        ]))
        story.append(fila)
    story.append(PageBreak())

    # ── 1. DESCRIPCION GENERAL ───────────────────────────────────────────────
    story.append(p("1. Descripcion General del Sistema", "h1"))
    story.append(LineaVerde())
    story.append(sp(0.3))
    story.append(p(
        "El <b>Validador Documental</b> es una aplicacion web desarrollada en "
        "<b>FastAPI</b> que automatiza la verificacion de la documentacion de "
        "beneficiarios almacenada en <b>Microsoft SharePoint</b>. El sistema recibe "
        "un archivo Excel con la matriz de registros, descarga selectivamente las "
        "carpetas de cada beneficiario y produce un checklist de validacion en Excel "
        "con el estado de cada documento.",
    ))
    story.append(sp(0.3))

    # Diagrama de flujo general
    flujo_data = [
        ["ENTRADA", "PROCESO", "SALIDA"],
        ["Matriz Excel\n(beneficiarios)", "Descarga SharePoint\nValidacion documentos\nAnalisis IA (opcional)", "Checklist Excel\n+ Pendientes si hay\ntimeouts"],
    ]
    ft = Table(flujo_data, colWidths=[4.5*cm, 6.5*cm, 4.5*cm])
    ft.setStyle(TableStyle([
        ("BACKGROUND",    (0,0), (-1,0), VERDE),
        ("TEXTCOLOR",     (0,0), (-1,0), BLANCO),
        ("FONTNAME",      (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE",      (0,0), (-1,-1), 9),
        ("ALIGN",         (0,0), (-1,-1), "CENTER"),
        ("VALIGN",        (0,0), (-1,-1), "MIDDLE"),
        ("BACKGROUND",    (0,1), (0,1), VERDE_CLARO),
        ("BACKGROUND",    (1,1), (1,1), colors.HexColor("#fff9e6")),
        ("BACKGROUND",    (2,1), (2,1), VERDE_CLARO),
        ("GRID",          (0,0), (-1,-1), 1, VERDE_MED),
        ("TOPPADDING",    (0,0), (-1,-1), 10),
        ("BOTTOMPADDING", (0,0), (-1,-1), 10),
        ("FONTNAME",      (0,1), (-1,-1), "Helvetica"),
    ]))
    story.append(ft)
    story.append(p("Figura 1. Flujo general del sistema.", "caption"))
    story.append(sp(0.3))

    story.append(p("<b>Carpetas validadas por beneficiario:</b>"))
    carpetas = [
        ["Carpeta", "Nombre", "Que valida"],
        ["00", "Documentacion", "Cedula, RUT, Comercio, Tenencia segun modalidad"],
        ["01", "Visita 1 — Caracterizacion", "Acta compromiso, acta visita, fotos/videos, tratamiento datos, gestor"],
        ["02", "Visita 2 — Diagnostico", "Acta visita 2, diagnostico empresarial, plan de negocio"],
        ["03", "Capacitacion", "Encuestas, lista grupal, reporte individual, modulos TX/RX"],
    ]
    story.append(tabla(carpetas, [1.5*cm, 5.0*cm, 9.0*cm]))
    story.append(sp(0.2))
    story.append(PageBreak())

    # ── 2. INSTALACION ───────────────────────────────────────────────────────
    story.append(p("2. Requisitos e Instalacion", "h1"))
    story.append(LineaVerde())
    story.append(sp(0.3))
    story.append(p("<b>Requisitos previos:</b>"))
    for req in ["Python 3.10 o superior", "Acceso a internet (para descargas de SharePoint y OpenAI)", "Archivo Excel de la matriz de beneficiarios"]:
        story.append(p(f"&#8226;  {req}", "bullet"))
    story.append(sp(0.3))

    pasos_inst = [
        ("1", "Clonar o descomprimir el proyecto",
         [p("Ubique la carpeta <b>validador_documental</b> en su equipo.")]),
        ("2", "Crear el entorno virtual",
         [p("<font face='Courier' size='8' color='#006633'>python -m venv .venv</font>", "code"),
          p("<font face='Courier' size='8' color='#006633'>.venv\\Scripts\\activate</font>", "code")]),
        ("3", "Instalar dependencias",
         [p("<font face='Courier' size='8' color='#006633'>pip install -r requirements.txt</font>", "code")]),
        ("4", "Configurar variables de entorno",
         [p("<font face='Courier' size='8' color='#006633'>copy .env.example .env</font>", "code"),
          p("Edite el archivo <b>.env</b> con su API Key de OpenAI si desea activar el analisis con IA.")]),
    ]
    for num, tit, elems in pasos_inst:
        story.extend(paso(num, tit, elems))

    story.append(PageBreak())

    # ── 3. ARRANQUE ──────────────────────────────────────────────────────────
    story.append(p("3. Arranque del Servidor", "h1"))
    story.append(LineaVerde())
    story.append(sp(0.3))
    story.append(p("Desde la carpeta del proyecto, ejecute:"))
    story.append(p("<font face='Courier' size='9' color='#006633'>python server.py</font>", "code"))
    story.append(sp(0.2))

    opciones = [
        ["Opcion", "Comando", "Uso"],
        ["Servidor basico",     "python server.py",              "Produccion / uso normal"],
        ["Recarga automatica",  "python server.py --reload",     "Desarrollo (recarga al cambiar codigo)"],
        ["Puerto personalizado","python server.py --port 9000",  "Cuando el puerto 8000 esta ocupado"],
    ]
    story.append(tabla(opciones, [4.5*cm, 6.0*cm, 5.0*cm]))
    story.extend(aviso("El servidor estara disponible en http://127.0.0.1:8000 — abra esta URL en su navegador para acceder a Swagger UI.", "ok"))
    story.append(PageBreak())

    # ── 4. SWAGGER ───────────────────────────────────────────────────────────
    story.append(p("4. Uso desde Swagger UI — Paso a Paso", "h1"))
    story.append(LineaVerde())
    story.append(sp(0.3))
    story.append(p(
        "Swagger UI es la interfaz web integrada que permite ejecutar todos los endpoints "
        "sin necesidad de Postman ni codigo adicional. Acceda a:"
    ))
    story.append(p("<font face='Courier' size='10' color='#006633'>http://127.0.0.1:8000/docs</font>", "code"))
    story.append(sp(0.3))

    # Mapa de endpoints
    endpoints = [
        ["Metodo", "Endpoint", "Que hace"],
        ["GET",  "/health",                       "Verifica que el servidor esta activo"],
        ["POST", "/validacion/validar",            "Sube la matriz e inicia la validacion"],
        ["GET",  "/validacion/estado/{job_id}",    "Consulta el progreso del trabajo"],
        ["GET",  "/validacion/resultado/{job_id}", "Descarga el checklist Excel resultante"],
        ["GET",  "/validacion/pendientes/{job_id}","Descarga las filas que fallaron por timeout"],
        ["POST", "/descarga/descargar",            "Prueba la conexion con una URL de SharePoint"],
    ]
    t_ep = Table(endpoints[0:1] + endpoints[1:], colWidths=[1.5*cm, 6.5*cm, 7.5*cm])
    t_ep.setStyle(TableStyle([
        ("BACKGROUND",    (0,0), (-1,0), VERDE),
        ("TEXTCOLOR",     (0,0), (-1,0), BLANCO),
        ("FONTNAME",      (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE",      (0,0), (-1,-1), 8.5),
        ("ALIGN",         (0,0), (0,-1), "CENTER"),
        ("VALIGN",        (0,0), (-1,-1), "MIDDLE"),
        ("GRID",          (0,0), (-1,-1), 0.3, colors.HexColor("#cccccc")),
        ("TOPPADDING",    (0,0), (-1,-1), 5),
        ("BOTTOMPADDING", (0,0), (-1,-1), 5),
        ("LEFTPADDING",   (0,0), (-1,-1), 8),
        ("BACKGROUND",    (0,2), (0,2), colors.HexColor("#e6f0ff")),
        ("BACKGROUND",    (0,4), (0,4), colors.HexColor("#e6f0ff")),
        ("BACKGROUND",    (0,6), (0,6), colors.HexColor("#e6f0ff")),
        ("TEXTCOLOR",     (0,1), (0,-1), AZUL_INFO),
        ("FONTNAME",      (0,1), (0,-1), "Helvetica-Bold"),
        ("TEXTCOLOR",     (0,2), (0,2), colors.HexColor("#cc6600")),
        ("TEXTCOLOR",     (0,4), (0,4), colors.HexColor("#cc6600")),
        ("TEXTCOLOR",     (0,6), (0,6), colors.HexColor("#cc6600")),
    ]))
    story.append(t_ep)
    story.append(sp(0.4))

    # Pasos Swagger
    story.extend(paso("1", "Verificar que el servidor esta activo", [
        p("Expanda <b>GET /health</b> → clic en <b>Try it out</b> → <b>Execute</b>"),
        p("Respuesta esperada:"),
        p('<font face="Courier" size="8" color="#006633">{ "status": "ok", "version": "2.0.0" }</font>', "code"),
    ]))

    story.extend(paso("2", "Lanzar la validacion — POST /validacion/validar", [
        p("1. Expanda <b>POST /validacion/validar</b> → clic en <b>Try it out</b>"),
        p("2. En el campo <b>archivo</b> → clic en <b>Choose File</b> → seleccione el Excel de la matriz"),
        p("3. En el campo <b>flujos</b> escriba los flujos a ejecutar (dejar vacio para todos):"),
        tabla([
            ["Que ejecutar",              "Valor del campo flujos"],
            ["Todo el flujo completo",    "(dejar vacio)"],
            ["Solo documentacion",        "00"],
            ["Solo primera visita",       "01"],
            ["Documentacion + visita 1",  "00,01"],
            ["Las dos visitas",           "01,02"],
            ["Todo excepto capacitacion", "00,01,02"],
        ], [7.5*cm, 6.0*cm]),
        sp(0.2),
        p("4. Clic en <b>Execute</b>"),
        p("5. Copie el <b>job_id</b> de la respuesta — lo necesitara en los siguientes pasos:"),
        p('<font face="Courier" size="8" color="#006633">{ "job_id": "a1b2c3d4-...", "estado": "iniciado", "flujos": ["00","01","02","03"] }</font>', "code"),
        *aviso("El proceso corre en background. La respuesta llega en menos de 1 segundo aunque la validacion tarde varios minutos.", "alerta"),
    ]))

    story.extend(paso("3", "Consultar el progreso — GET /validacion/estado/{job_id}", [
        p("1. Expanda <b>GET /validacion/estado/{job_id}</b> → <b>Try it out</b>"),
        p("2. Pegue el <b>job_id</b> copiado en el paso anterior"),
        p("3. Clic en <b>Execute</b> — repita hasta que <b>estado</b> sea <b>completado</b>"),
        sp(0.2),
        tabla([
            ["Estado",       "Significado"],
            ["iniciado",     "Job registrado, aun no comenzo a procesar filas"],
            ["procesando",   "Validando filas — vea filas_procesadas / filas_total"],
            ["completado",   "Todas las filas procesadas — checklist disponible"],
            ["error",        "Falla critica (archivo invalido, SharePoint caido)"],
        ], [4.0*cm, 11.5*cm]),
    ]))

    story.extend(paso("4", "Descargar el checklist — GET /validacion/resultado/{job_id}", [
        p("1. Expanda <b>GET /validacion/resultado/{job_id}</b> → <b>Try it out</b>"),
        p("2. Pegue el <b>job_id</b>"),
        p("3. Clic en <b>Execute</b>"),
        p("4. En la seccion <b>Response body</b>, haga clic en <b>Download file</b>"),
        p("5. Guarde el archivo <b>checklist_validacion.xlsx</b>"),
        *aviso("Solo disponible cuando el estado es 'completado'. Si se llama antes, retorna HTTP 400.", "info"),
    ]))

    story.extend(paso("5", "Verificar filas pendientes — GET /validacion/pendientes/{job_id}", [
        p("Si alguna fila fallo por timeout o error de red, el campo <b>filas_omitidas</b> en "
          "/estado sera mayor que 0."),
        p("1. Expanda <b>GET /validacion/pendientes/{job_id}</b> → <b>Try it out</b>"),
        p("2. Pegue el <b>job_id</b> → <b>Execute</b>"),
        p("3. Descargue <b>pendientes_reintento.xlsx</b>"),
        p("4. Suba ese archivo como nueva matriz para reprocesar solo las filas fallidas."),
        *aviso("Si no hubo filas pendientes, este endpoint retorna HTTP 404 — es normal.", "ok"),
    ]))

    story.append(PageBreak())

    # ── 5. FLUJOS ────────────────────────────────────────────────────────────
    story.append(p("5. Flujos de Validacion — Detalle por Carpeta", "h1"))
    story.append(LineaVerde())
    story.append(sp(0.3))

    flujos_det = [
        ("00", "Documentacion", VERDE, [
            ["Documento",  "M1",          "M2",          "M3",          "M4"],
            ["CEDULA",     "Obligatorio", "Obligatorio", "Obligatorio", "Obligatorio"],
            ["COMERCIO",   "No aplica",   "Opcional",    "Obligatorio", "Obligatorio"],
            ["RUT",        "Opcional",    "Opcional",    "Obligatorio", "Obligatorio"],
            ["TENENCIA",   "Opcional",    "Obligatorio", "Obligatorio", "Obligatorio"],
        ]),
        ("01", "Visita 1 — Caracterizacion", VERDE_MED, [
            ["Documento",          "Requerido", "Que verifica la IA"],
            ["ACTA_COMPROMISO",    "Si",        "Firmas y fechas correctas"],
            ["ACTA_VISITA_1",      "Si",        "Extrae nombre y cedula del gestor"],
            ["TRATAMIENTO_DATOS",  "Si",        "Autorizacion firmada presente"],
            ["Fotos / Videos",     "Min. 4",    "Conteo de archivos multimedia"],
            ["GESTOR",             "Si",        "Verifica gestor en base de datos interna"],
        ]),
        ("02", "Visita 2 — Diagnostico", AZUL_INFO, [
            ["Documento",    "Requerido", "Que verifica la IA"],
            ["ACTA_VISITA_2","Si",        "Contenido y firmas del acta"],
            ["DIAGNOSTICO",  "Si",        "Documento completo y coherente"],
            ["PLAN_NEGOCIO", "Si",        "Campos obligatorios completados"],
        ]),
        ("03", "Capacitacion", colors.HexColor("#1F5C6B"), [
            ["Elemento",       "Descripcion"],
            ["Encuestas",      "Archivos ENCUESTA_XXXXX — uno por participante"],
            ["GRUPAL",         "Lista de asistencia grupal del modulo"],
            ["Individual",     "Reporte ID_ / Formacion / Cumplimiento por asistente"],
            ["Modulos TX/RX",  "Archivos T{n}_R{n}_... — lista de asistencia por modulo"],
            ["Asistencia IA",  "GPT-4.1 verifica nombres de encuestas en TX_RX"],
        ]),
    ]

    for cod, nombre, color, datos in flujos_det:
        story.append(p(f"Carpeta {cod} — {nombre}", "h2"))
        ancho_cols = [4.5*cm, 2.5*cm, 8.5*cm] if len(datos[0]) == 3 else \
                     [4.5*cm, 3.0*cm, 3.0*cm, 3.0*cm, 2.0*cm] if len(datos[0]) == 5 else \
                     [5.0*cm, 10.5*cm]
        t = Table(datos, colWidths=ancho_cols)
        t.setStyle(TableStyle([
            ("BACKGROUND",    (0,0), (-1,0), color),
            ("TEXTCOLOR",     (0,0), (-1,0), BLANCO),
            ("FONTNAME",      (0,0), (-1,0), "Helvetica-Bold"),
            ("FONTSIZE",      (0,0), (-1,-1), 8.5),
            ("ALIGN",         (0,0), (-1,-1), "LEFT"),
            ("VALIGN",        (0,0), (-1,-1), "MIDDLE"),
            ("GRID",          (0,0), (-1,-1), 0.3, colors.HexColor("#cccccc")),
            ("TOPPADDING",    (0,0), (-1,-1), 5),
            ("BOTTOMPADDING", (0,0), (-1,-1), 5),
            ("LEFTPADDING",   (0,0), (-1,-1), 8),
            ("ROWBACKGROUND", (0,1), (-1,1), VERDE_CLARO),
            ("ROWBACKGROUND", (0,3), (-1,3), GRIS_CLARO),
        ]))
        story.append(t)
        story.append(sp(0.4))

    story.append(PageBreak())

    # ── 6. ESTRUCTURA ENTRADA ────────────────────────────────────────────────
    story.append(p("6. Estructura del Archivo de Entrada (Matriz)", "h1"))
    story.append(LineaVerde())
    story.append(sp(0.3))
    story.append(p(
        "El archivo Excel de entrada debe tener una hoja activa con las siguientes columnas. "
        "Los nombres deben coincidir exactamente (sensible a mayusculas)."
    ))
    story.append(sp(0.3))
    cols_entrada = [
        ["Columna",              "Requerida", "Descripcion"],
        ["ID_unico",             "Si",        "Identificador unico del beneficiario"],
        ["modalidad",            "Si",        "M1, M2, M3 o M4"],
        ["unidad_doc",           "Si",        "Numero de cedula o NIT del beneficiario"],
        ["carpetas_link",        "Si",        "URL anonima de la carpeta en SharePoint"],
        ["integrante_nombre1",   "No",        "Primer nombre del integrante"],
        ["integrante_nombre2",   "No",        "Segundo nombre del integrante"],
        ["integrante_apellido1", "No",        "Primer apellido del integrante"],
        ["integrante_apellido2", "No",        "Segundo apellido del integrante"],
    ]
    story.append(tabla(cols_entrada, [5.0*cm, 2.5*cm, 8.0*cm]))
    story.append(sp(0.3))
    story.extend(aviso(
        "La columna carpetas_link debe contener la URL de la carpeta raiz del beneficiario en SharePoint, "
        "no la URL de una subcarpeta especifica. El sistema busca automaticamente las subcarpetas 00, 01, 02 y 03.",
        "alerta"
    ))
    story.extend(aviso(
        "El proceso es reanudable: si se interrumpe, al volver a ejecutar con el mismo archivo de salida "
        "continua desde donde quedo, saltando los ID_unico ya procesados.",
        "info"
    ))
    story.append(PageBreak())

    # ── 7. CHECKLIST SALIDA ──────────────────────────────────────────────────
    story.append(p("7. Checklist de Salida — Interpretacion", "h1"))
    story.append(LineaVerde())
    story.append(sp(0.3))
    story.append(p(
        "El checklist Excel de salida tiene una fila por beneficiario y columnas agrupadas "
        "por carpeta. Los encabezados de grupo usan colores para identificar cada seccion."
    ))
    story.append(sp(0.3))

    grupos = [
        ["Grupo (color encabezado)",       "Columnas incluidas"],
        ["Identificacion (azul)",          "ID_unico, modalidad, unidad_doc"],
        ["00 Documentacion (naranja)",     "CEDULA, COMERCIO, RUT, TENENCIA"],
        ["01 Visita (verde oscuro)",       "01_DOCUMENTOS, 01_FOTOS_VIDEOS, ACTA_COMPROMISO, ACTA_VISITA, GESTOR, TRATAMIENTO_DATOS"],
        ["02 Visita 2 (morado)",           "02_DOCUMENTOS, 02_ACTA_VISITA, 02_DIAGNOSTICO, 02_PLAN_NEGOCIO"],
        ["03 Capacitacion (teal/verde)",   "03_ENCUESTAS, 03_GRUPAL, 03_INDIVIDUAL, 03_MODULOS, 03_ASISTENCIA"],
        ["General (gris)",                 "observaciones"],
    ]
    story.append(tabla(grupos, [6.5*cm, 9.0*cm]))
    story.append(sp(0.4))

    story.append(p("<b>Significado de los valores en las celdas:</b>"))
    valores = [
        ["Valor",                 "Significado",                           "Color celda"],
        ["OK",                   "Documento obligatorio presente",         "Sin color"],
        ["FALTA (obligatorio)",  "Documento obligatorio ausente",          "Rojo"],
        ["Presente",             "Documento opcional encontrado",          "Sin color"],
        ["Ausente",              "Documento opcional no encontrado",       "Sin color"],
        ["N/A",                  "No aplica para esta modalidad",          "Sin color"],
        ["—",                    "Flujo no ejecutado (seleccion parcial)", "Sin color"],
        ["ALERTA: ...",          "La IA detecto un problema en el doc.",   "Amarillo"],
        ["NO REGISTRADO — ...",  "El gestor no esta en la base de datos",  "Rojo"],
    ]
    story.append(tabla(valores, [4.5*cm, 7.5*cm, 3.5*cm]))
    story.append(PageBreak())

    # ── 8. FLUJOS INDEPENDIENTES ─────────────────────────────────────────────
    story.append(p("8. Ejecucion de Flujos Independientes", "h1"))
    story.append(LineaVerde())
    story.append(sp(0.3))
    story.append(p(
        "El sistema permite ejecutar uno o varios flujos de validacion sin necesidad de "
        "procesar todas las carpetas. Esto es util cuando solo se requiere verificar "
        "un aspecto especifico de la documentacion o cuando se quiere reutilizar el resultado "
        "de una ejecucion anterior."
    ))
    story.append(sp(0.3))
    story.extend(aviso(
        "Las columnas de los flujos no ejecutados apareceran con el valor '—' en el checklist, "
        "indicando que no fueron evaluadas en esa ejecucion.", "info"
    ))
    story.append(sp(0.3))

    ejemplos = [
        ["Que ejecutar",                  "Campo 'flujos' en Swagger"],
        ["Todo el flujo completo",        "(dejar vacio)"],
        ["Solo documentacion",            "00"],
        ["Solo primera visita",           "01"],
        ["Solo segunda visita",           "02"],
        ["Solo capacitacion",             "03"],
        ["Documentacion + primera visita","00,01"],
        ["Las dos visitas",               "01,02"],
        ["Todo excepto capacitacion",     "00,01,02"],
        ["Documentacion + capacitacion",  "00,03"],
    ]
    story.append(tabla(ejemplos, [8.0*cm, 7.5*cm]))
    story.append(sp(0.4))
    story.extend(aviso(
        "Caso de uso tipico: ejecutar primero '00,01' para validar documentacion basica. "
        "Cuando esten listas las visitas 2 y capacitacion, ejecutar '02,03' sobre la misma matriz — "
        "el sistema retomara sin repetir los IDs ya procesados.", "alerta"
    ))
    story.append(PageBreak())

    # ── 9. PENDIENTES POR TIMEOUT ────────────────────────────────────────────
    story.append(p("9. Filas Pendientes por Timeout o Error de Red", "h1"))
    story.append(LineaVerde())
    story.append(sp(0.3))
    story.append(p(
        "Cuando una fila no puede procesarse por problemas de conectividad con SharePoint "
        "(timeout, error de red, servidor no disponible), el sistema la omite, "
        "continua con la siguiente fila y al finalizar genera un archivo "
        "<b>pendientes_reintento.xlsx</b> con las filas que no se ejecutaron."
    ))
    story.append(sp(0.3))

    flujo_timeout = [
        ["Evento",                        "Que hace el sistema"],
        ["Timeout durante descarga",      "Omite la fila, limpia temporales, pasa a la siguiente"],
        ["Error de red / conexion",       "Igual que timeout — registra en log y continua"],
        ["Error critico inesperado",      "Omite la fila — no detiene el job completo"],
        ["Al finalizar con omisiones",    "Genera pendientes_reintento.xlsx en la carpeta del job"],
        ["Campo filas_omitidas en /estado","Indica cuantas filas fueron omitidas"],
    ]
    story.append(tabla(flujo_timeout, [6.5*cm, 9.0*cm]))
    story.append(sp(0.4))

    story.append(p("<b>Como reprocesar las filas pendientes:</b>", "h3"))
    story.extend(paso("1", "Verificar que hay pendientes",
        [p("En <b>GET /validacion/estado/{job_id}</b> revise el campo <b>filas_omitidas</b>. Si es 0, no hay pendientes.")]
    ))
    story.extend(paso("2", "Descargar pendientes_reintento.xlsx",
        [p("Ejecute <b>GET /validacion/pendientes/{job_id}</b> y descargue el archivo.")]
    ))
    story.extend(paso("3", "Subir como nueva matriz",
        [p("Suba <b>pendientes_reintento.xlsx</b> como archivo de entrada en un nuevo <b>POST /validacion/validar</b>.")]
    ))
    story.extend(paso("4", "Combinar checklists",
        [p("El nuevo checklist contendra solo los beneficiarios pendientes. Puede combinar ambos archivos Excel manualmente.")]
    ))
    story.append(PageBreak())

    # ── 10. VARIABLES DE ENTORNO ─────────────────────────────────────────────
    story.append(p("10. Variables de Entorno (.env)", "h1"))
    story.append(LineaVerde())
    story.append(sp(0.3))
    story.append(p(
        "Copie <b>.env.example</b> a <b>.env</b> y ajuste los valores segun su entorno. "
        "Todas las variables tienen valores por defecto funcionales — solo es obligatorio "
        "configurar <b>OPENAI_API_KEY</b> si desea activar el analisis con IA."
    ))
    story.append(sp(0.3))
    env_vars = [
        ["Variable",                "Valor por defecto",   "Descripcion"],
        ["JOBS_DIR",                "jobs",                "Directorio donde se guardan los trabajos y resultados"],
        ["DOWNLOADS_DIR",           "descargas_sharepoint","Directorio temporal de descargas de SharePoint"],
        ["MAX_WORKERS_LISTADO",     "8",                   "Hilos paralelos para listar carpetas en SharePoint"],
        ["MAX_WORKERS_DESCARGA",    "12",                  "Hilos paralelos para descargar archivos"],
        ["CHECKLIST_LOTE_GUARDADO", "50",                  "Numero de filas antes de guardar el checklist a disco"],
        ["TIMEOUT_DESCARGA",        "120",                 "Segundos maximos de espera por archivo de SharePoint"],
        ["IA_HABILITADO",           "false",               "Activa el analisis semantico con OpenAI (true/false)"],
        ["OPENAI_API_KEY",          "(vacio)",             "API Key de OpenAI — requerida si IA_HABILITADO=true"],
        ["OPENAI_MODEL",            "gpt-4o",              "Modelo OpenAI para analisis visual de documentos"],
    ]
    story.append(tabla(env_vars, [4.5*cm, 3.5*cm, 7.5*cm]))
    story.append(sp(0.4))
    story.extend(aviso(
        "Para activar el analisis con IA, configure IA_HABILITADO=true y proporcione su OPENAI_API_KEY. "
        "El modelo gpt-4o analiza visualmente los documentos escaneados para verificar firmas, fechas y contenido.",
        "info"
    ))
    story.append(PageBreak())

    # ── 11. FAQ ──────────────────────────────────────────────────────────────
    story.append(p("11. Preguntas Frecuentes", "h1"))
    story.append(LineaVerde())
    story.append(sp(0.3))

    faqs = [
        ("El endpoint /resultado retorna HTTP 400",
         "El job aun no ha terminado. Consulte /estado y espere a que el estado sea 'completado'."),
        ("Aparece 'N/A' en todas las columnas de un flujo",
         "La carpeta correspondiente no fue encontrada en SharePoint para ese beneficiario. "
         "Verifique que la URL en carpetas_link sea correcta y que la carpeta exista."),
        ("Aparece '—' en las columnas de un flujo",
         "Ese flujo no fue incluido en la ejecucion. Use el campo 'flujos' en Swagger para incluirlo, "
         "o deje el campo vacio para ejecutar todos."),
        ("Sale 'Error IA (ver log)' en una columna de revision",
         "Ocurrio un error al comunicarse con OpenAI. Verifique su OPENAI_API_KEY y conexion a internet. "
         "Revise el log del servidor para el detalle del error."),
        ("El proceso se queda pegado en una fila",
         "SharePoint puede tardar o no responder. El sistema tiene un timeout de 120 segundos por archivo — "
         "pasado ese tiempo, omite la fila y la incluye en pendientes_reintento.xlsx."),
        ("Como reprocesar solo algunos beneficiarios",
         "Cree un Excel con solo las filas que desea reprocesar y subalo como nueva matriz. "
         "Si ya fueron procesados en un checklist anterior, use un archivo de salida diferente."),
        ("El gestor aparece como 'NO REGISTRADO'",
         "La IA extrajo el nombre o cedula del gestor del acta de visita, pero no coincide con "
         "ningun registro en la base de datos interna. Verifique el documento y la base de gestores."),
    ]

    for pregunta, respuesta in faqs:
        caja_faq = Table([
            [Paragraph(f"P: {pregunta}", ParagraphStyle("q", fontName="Helvetica-Bold", fontSize=9, textColor=VERDE_OSC, leading=13))],
            [Paragraph(f"R: {respuesta}", ParagraphStyle("a", fontName="Helvetica", fontSize=9, textColor=GRIS_OSC, leading=13, leftIndent=6))],
        ], colWidths=[W - 4*cm])
        caja_faq.setStyle(TableStyle([
            ("BACKGROUND",    (0,0), (-1,0), VERDE_CLARO),
            ("BACKGROUND",    (0,1), (-1,1), BLANCO),
            ("BOX",           (0,0), (-1,-1), 0.8, VERDE_MED),
            ("LINEBELOW",     (0,0), (-1,0), 0.5, VERDE_MED),
            ("TOPPADDING",    (0,0), (-1,-1), 7),
            ("BOTTOMPADDING", (0,0), (-1,-1), 7),
            ("LEFTPADDING",   (0,0), (-1,-1), 10),
            ("RIGHTPADDING",  (0,0), (-1,-1), 10),
        ]))
        story.append(caja_faq)
        story.append(sp(0.3))

    # ── PIE FINAL ────────────────────────────────────────────────────────────
    story.append(sp(0.5))
    story.append(LineaVerde(grosor=1))
    story.append(sp(0.3))
    story.append(Paragraph(
        "Universidad de Antioquia  —  Programa de Fortalecimiento Empresarial  —  2026",
        ParagraphStyle("pie_final", fontName="Helvetica", fontSize=8,
                       textColor=GRIS_MED, alignment=TA_CENTER)
    ))

    return story


# ── GENERAR PDF ───────────────────────────────────────────────────────────────
def generar():
    doc = SimpleDocTemplate(
        OUTPUT,
        pagesize=A4,
        leftMargin=2*cm, rightMargin=2*cm,
        topMargin=1.8*cm, bottomMargin=1.6*cm,
        title="Manual de Usuario — Validador Documental",
        author="Universidad de Antioquia",
        subject="Validacion documental SharePoint",
    )

    def _primera_pagina(c, d):
        _portada(c, d)

    def _paginas_siguientes(c, d):
        _encabezado_pie(c, d)

    story = build_story()
    doc.build(story, onFirstPage=_primera_pagina, onLaterPages=_paginas_siguientes)
    print(f"PDF generado: {OUTPUT}")


if __name__ == "__main__":
    generar()
