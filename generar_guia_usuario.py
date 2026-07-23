"""Genera el PDF 'Guia del Usuario' del bot de facturas.

Uso:
    python3 generar_guia_usuario.py

Salida: Guia_Robot_Facturas.pdf en la carpeta del proyecto.
"""
from pathlib import Path

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, PageBreak, Table, TableStyle,
    KeepTogether, HRFlowable,
)


# ── Paleta de colores ─────────────────────────────────────────────────────────
AZUL_OSCURO = colors.HexColor("#1F3A5F")
AZUL_CLARO = colors.HexColor("#3E6DA3")
GRIS_FONDO = colors.HexColor("#F3F5F8")
GRIS_TEXTO = colors.HexColor("#44474D")
VERDE_OK = colors.HexColor("#2E8B57")
NARANJA = colors.HexColor("#E07B2F")


def _footer(canvas, doc):
    canvas.saveState()
    canvas.setStrokeColor(colors.HexColor("#D5DBE3"))
    canvas.setLineWidth(0.4)
    canvas.line(2 * cm, 1.5 * cm, A4[0] - 2 * cm, 1.5 * cm)
    canvas.setFillColor(colors.HexColor("#8A909B"))
    canvas.setFont("Helvetica", 8)
    canvas.drawString(2 * cm, 1.1 * cm, "Robot de Facturas — Guia del Usuario")
    canvas.drawRightString(A4[0] - 2 * cm, 1.1 * cm, f"Pagina {doc.page}")
    canvas.restoreState()


def _cover(canvas, doc):
    width, height = A4
    # Fondo superior azul
    canvas.setFillColor(AZUL_OSCURO)
    canvas.rect(0, height - 10 * cm, width, 10 * cm, fill=1, stroke=0)
    # Titulo
    canvas.setFillColor(colors.white)
    canvas.setFont("Helvetica-Bold", 34)
    canvas.drawCentredString(width / 2, height - 5.2 * cm, "Robot de Facturas")
    canvas.setFont("Helvetica", 16)
    canvas.drawCentredString(width / 2, height - 6.2 * cm, "Guia del usuario")
    # Linea decorativa
    canvas.setStrokeColor(NARANJA)
    canvas.setLineWidth(3)
    canvas.line(width / 2 - 2 * cm, height - 7 * cm, width / 2 + 2 * cm, height - 7 * cm)
    # Subtitulo
    canvas.setFillColor(colors.HexColor("#D5DBE3"))
    canvas.setFont("Helvetica-Oblique", 12)
    canvas.drawCentredString(
        width / 2, height - 8.2 * cm,
        "Tu asistente automatico para facturas electronicas FEL"
    )
    # Pie
    canvas.setFillColor(AZUL_CLARO)
    canvas.rect(0, 0, width, 2 * cm, fill=1, stroke=0)
    canvas.setFillColor(colors.white)
    canvas.setFont("Helvetica", 10)
    canvas.drawCentredString(
        width / 2, 0.9 * cm,
        "Envia una foto o PDF — recibe un reporte Excel listo"
    )


def _section_title(texto, estilos):
    return Paragraph(
        f'<font color="{AZUL_OSCURO.hexval()}"><b>{texto}</b></font>', estilos["h1"]
    )


def _subsection(texto, estilos):
    return Paragraph(
        f'<font color="{AZUL_CLARO.hexval()}"><b>{texto}</b></font>', estilos["h2"]
    )


def _paragrafo(texto, estilos):
    return Paragraph(texto, estilos["body"])


def _bullet(texto, estilos):
    return Paragraph(f"• {texto}", estilos["bullet"])


def _step_box(numero, titulo, texto, estilos):
    data = [[
        Paragraph(f'<font color="white" size="20"><b>{numero}</b></font>',
                  ParagraphStyle("n", alignment=TA_CENTER)),
        [
            Paragraph(
                f'<b><font color="{AZUL_OSCURO.hexval()}">{titulo}</font></b>',
                estilos["h3"]),
            Paragraph(texto, estilos["body"]),
        ],
    ]]
    t = Table(data, colWidths=[1.6 * cm, 14 * cm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (0, 0), AZUL_CLARO),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (0, 0), (0, 0), "CENTER"),
        ("LEFTPADDING", (1, 0), (1, 0), 12),
        ("RIGHTPADDING", (1, 0), (1, 0), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
        ("LINEAFTER", (0, 0), (0, 0), 0.5, AZUL_OSCURO),
    ]))
    return KeepTogether([t, Spacer(1, 0.35 * cm)])


def _info_box(titulo, texto, estilos, color_fondo=GRIS_FONDO, color_titulo=AZUL_OSCURO):
    data = [
        [Paragraph(f'<b><font color="{color_titulo.hexval()}">{titulo}</font></b>',
                   estilos["h3"])],
        [Paragraph(texto, estilos["body"])],
    ]
    t = Table(data, colWidths=[16 * cm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), color_fondo),
        ("LEFTPADDING", (0, 0), (-1, -1), 12),
        ("RIGHTPADDING", (0, 0), (-1, -1), 12),
        ("TOPPADDING", (0, 0), (0, 0), 8),
        ("BOTTOMPADDING", (0, 0), (0, 0), 2),
        ("TOPPADDING", (0, 1), (0, 1), 2),
        ("BOTTOMPADDING", (0, 1), (0, 1), 10),
        ("LINEBEFORE", (0, 0), (0, -1), 3, color_titulo),
        ("ROUNDEDCORNERS", (0, 0), (-1, -1), [0, 0, 0, 0]),
    ]))
    return KeepTogether([t, Spacer(1, 0.35 * cm)])


def _tabla_comandos(estilos):
    datos = [
        ["Comando", "Para que sirve"],
        ["/start o /help", "Ver este menu dentro del bot"],
        ["/perfil", "Ver y configurar tus datos (codigo, CECO, nombre, apellido)"],
        ["/reporte", "Generar el Excel con todas las facturas del periodo"],
        ["/buscar", "Buscar facturas por empresa, NIT, fecha, monto"],
        ["/eliminar", "Borrar una factura del registro"],
        ["/tipo",
         "Cambiar el tipo de una factura (Desayuno, Almuerzo, Cena, Gasolina, Hospedaje)"],
        ["/odometro",
         "Registrar el kilometraje del carro en cada factura de gasolina"],
        ["/catalogo", "Ver las empresas que el bot ya conoce"],
        ["/agregar", "Agregar una empresa nueva al catalogo"],
        ["/pendientes", "Ver las facturas que todavia no estan confirmadas"],
        ["/fallidas", "Ver facturas que dieron error y reintentar"],
        ["/estado", "Ver si el bot esta sano y cuantas facturas has procesado"],
    ]
    t = Table(datos, colWidths=[4.2 * cm, 11.8 * cm], repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), AZUL_OSCURO),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 10),
        ("FONTSIZE", (0, 1), (-1, -1), 9.5),
        ("FONTNAME", (0, 1), (0, -1), "Helvetica-Bold"),
        ("TEXTCOLOR", (0, 1), (0, -1), AZUL_OSCURO),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, GRIS_FONDO]),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
        ("LINEBELOW", (0, 0), (-1, 0), 1, colors.white),
    ]))
    return t


def _tabla_reporte(estilos):
    datos = [
        ["Dato del Excel", "De donde sale"],
        ["Fecha del gasto", "Fecha de emision que aparece en la factura"],
        ["Departamento", "Direccion del emisor (ej. ALTA VERAPAZ, GUATEMALA)"],
        ["Tipo de gasto", "Alimentos, Hospedaje o Gasolina (el bot lo detecta)"],
        ["Concepto", "Desayuno, Almuerzo, Cena, Hospedaje o Gasolina"],
        ["Kilometraje", "Solo en gasolina: diferencia de odometro entre cargas"],
        ["Monto factura", "Total que aparece en la factura"],
        ["Monto politica",
         "Maximo que la politica reconoce segun departamento y tipo"],
        ["Maximo reembolsable", "El menor entre el monto factura y el monto politica"],
        ["Codigo de empleado", "Lo que configuraste con /perfil"],
        ["CECO", "Lo que configuraste con /perfil o lo que das al pedir /reporte"],
    ]
    t = Table(datos, colWidths=[5.3 * cm, 10.7 * cm], repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), VERDE_OK),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 10),
        ("FONTSIZE", (0, 1), (-1, -1), 9.5),
        ("FONTNAME", (0, 1), (0, -1), "Helvetica-Bold"),
        ("TEXTCOLOR", (0, 1), (0, -1), AZUL_OSCURO),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, GRIS_FONDO]),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
    ]))
    return t


def construir(path_salida: Path):
    doc = SimpleDocTemplate(
        str(path_salida),
        pagesize=A4,
        leftMargin=2 * cm, rightMargin=2 * cm,
        topMargin=2 * cm, bottomMargin=2 * cm,
        title="Guia del Usuario — Robot de Facturas",
        author="Robot de Facturas",
    )

    base = getSampleStyleSheet()
    estilos = {
        "h1": ParagraphStyle(
            "h1", parent=base["Heading1"], fontName="Helvetica-Bold",
            fontSize=18, leading=22, spaceBefore=4, spaceAfter=10,
            textColor=AZUL_OSCURO,
        ),
        "h2": ParagraphStyle(
            "h2", parent=base["Heading2"], fontName="Helvetica-Bold",
            fontSize=13, leading=16, spaceBefore=10, spaceAfter=4,
            textColor=AZUL_CLARO,
        ),
        "h3": ParagraphStyle(
            "h3", parent=base["Heading3"], fontName="Helvetica-Bold",
            fontSize=11, leading=14, spaceBefore=0, spaceAfter=2,
        ),
        "body": ParagraphStyle(
            "body", parent=base["BodyText"], fontName="Helvetica",
            fontSize=10.5, leading=15, textColor=GRIS_TEXTO, alignment=TA_JUSTIFY,
            spaceAfter=4,
        ),
        "bullet": ParagraphStyle(
            "bullet", parent=base["BodyText"], fontName="Helvetica",
            fontSize=10.5, leading=15, textColor=GRIS_TEXTO,
            leftIndent=14, spaceAfter=3,
        ),
        "quote": ParagraphStyle(
            "quote", parent=base["BodyText"], fontName="Helvetica-Oblique",
            fontSize=10, leading=14, textColor=GRIS_TEXTO, leftIndent=14,
            spaceAfter=6,
        ),
    }

    elementos = []

    # ── Portada (se dibuja via _cover) ──
    elementos.append(PageBreak())

    # ── 1. Que es este bot ──
    elementos.append(_section_title("1. Que es este bot", estilos))
    elementos.append(_paragrafo(
        "El <b>Robot de Facturas</b> es un asistente dentro de Telegram. Su trabajo es "
        "recibir tus facturas electronicas (las FEL de Guatemala), leerlas, validarlas "
        "con el SAT, organizarlas y al final entregarte un <b>reporte de gastos en "
        "Excel listo para liquidar</b>. Tu solo tomas la foto o reenvias el PDF y el bot "
        "se encarga del resto.",
        estilos,
    ))
    elementos.append(Spacer(1, 0.2 * cm))
    elementos.append(_info_box(
        "En una frase",
        "Tomas foto de la factura -> el bot la verifica con SAT -> la archiva -> al final del "
        "periodo le pides tu reporte y recibes un Excel con todos los datos ya "
        "clasificados y un ZIP con los PDFs oficiales.",
        estilos,
    ))

    # ── 2. Para que sirve ──
    elementos.append(_subsection("Para que te sirve", estilos))
    for b in [
        "No tener que copiar manualmente los datos de cada factura al Excel de liquidacion.",
        "Evitar errores: el bot valida cada factura directamente en la pagina del SAT.",
        "Clasificar automaticamente si es Desayuno, Almuerzo, Cena, Gasolina u Hospedaje.",
        "Controlar el kilometraje del carro cuando cargas combustible.",
        "No duplicar facturas por error (aunque mandes dos veces la misma foto).",
        "Guardar los PDFs oficiales del SAT, organizados por usuario.",
    ]:
        elementos.append(_bullet(b, estilos))
    elementos.append(PageBreak())

    # ── 3. Como se usa ──
    elementos.append(_section_title("2. Como se usa, paso a paso", estilos))
    elementos.append(_step_box(
        "1", "Configura tu perfil una sola vez",
        "Escribe <b>/perfil nombre Juan</b>, <b>/perfil apellido Perez</b>, "
        "<b>/perfil codigo 12345</b> y <b>/perfil ceco 349065</b>. Esos datos se usan "
        "para el reporte y para que el archivo Excel salga con tu nombre.",
        estilos,
    ))
    elementos.append(_step_box(
        "2", "Envia la factura al bot",
        "Puedes mandarle una foto con la camara del celular o reenviar el PDF que te "
        "llega por correo. Tambien funciona si la fotografias desde una pantalla. "
        "Asegurate de que se vean claro el NIT, la serie y el numero de DTE.",
        estilos,
    ))
    elementos.append(_step_box(
        "3", "El bot lee y valida con el SAT",
        "El bot extrae los datos clave de la factura, resuelve el captcha del SAT "
        "por ti y baja el PDF oficial. Si todo cuadra, te envia la confirmacion con el "
        "monto, la fecha, el establecimiento y el departamento.",
        estilos,
    ))
    elementos.append(_step_box(
        "4", "Responde preguntas rapidas si hace falta",
        "Si es gasolina, te pregunta el kilometraje del odometro. Si no pudo clasificar "
        "el tipo de gasto (desayuno, almuerzo, cena...), te muestra botones para que "
        "elijas. Siempre puedes dar 'Saltar' y arreglarlo despues.",
        estilos,
    ))
    elementos.append(_step_box(
        "5", "Al final del periodo pide tu reporte",
        "Cuando quieras liquidar, escribe <b>/reporte</b>. El bot te pregunta por el "
        "rango de fechas y te devuelve el Excel ya lleno, un archivo de texto con el "
        "detalle y un ZIP con todos los PDFs oficiales del SAT listos para adjuntar.",
        estilos,
    ))

    elementos.append(PageBreak())

    # ── 4. Que hace el bot detras de la foto ──
    elementos.append(_section_title(
        "3. Que hace el bot detras de cada factura", estilos
    ))
    elementos.append(_paragrafo(
        "Cuando le mandas una foto o un PDF, el bot sigue este camino para que tu no "
        "tengas que hacer nada mas:",
        estilos,
    ))
    pasos_internos = [
        ("Prepara la imagen",
         "Si esta torcida, la endereza. Si quedo oscura, la aclara. Si esta borrosa, "
         "te avisa antes de seguir para que vuelvas a tomarla."),
        ("Encuentra los datos importantes",
         "Busca en la factura el NIT del emisor (no el tuyo, el de quien te la "
         "emitio), la serie y el numero de DTE. Usa varias tecnicas para que incluso "
         "una foto un poco torcida funcione."),
        ("Consulta el SAT",
         "Entra al portal del SAT, resuelve el captcha, valida que la factura existe "
         "y descarga el PDF oficial. Si el SAT tarda, reintenta varias veces sin que "
         "tu hagas nada."),
        ("Revisa que no este duplicada",
         "Si esa misma factura ya la habias enviado antes, te lo dice y no la "
         "registra dos veces."),
        ("Clasifica el tipo de gasto",
         "Mira el nombre del establecimiento, la hora y los productos de la factura "
         "para decidir si fue Desayuno, Almuerzo, Cena, Gasolina u Hospedaje. "
         "Si tiene dudas, te pregunta."),
        ("Guarda todo en tu carpeta personal",
         "El PDF oficial se archiva en tu carpeta privada, y los datos se anotan en un "
         "registro para cuando quieras generar el reporte."),
    ]
    for i, (tit, desc) in enumerate(pasos_internos, 1):
        elementos.append(_step_box(str(i), tit, desc, estilos))

    # ── 5. Como se arma el reporte ──
    elementos.append(PageBreak())
    elementos.append(_section_title("4. Como arma el reporte en Excel", estilos))
    elementos.append(_paragrafo(
        "Cuando pides <b>/reporte</b>, el bot toma todas tus facturas del periodo que "
        "elegiste, las ordena por fecha y llena la plantilla oficial. El archivo te "
        "llega con un nombre claro como:",
        estilos,
    ))
    elementos.append(_info_box(
        "Ejemplo de nombre del archivo",
        "<font face='Courier'>12345 Juan Perez 14-17 Abril.xlsx</font><br/>"
        "(codigo de empleado, tu nombre y apellido, y el rango de fechas del reporte)",
        estilos, color_fondo=colors.HexColor("#FFF4EA"), color_titulo=NARANJA,
    ))
    elementos.append(_paragrafo(
        "El reporte se entrega junto con dos cosas mas: un archivo de texto con el "
        "detalle de cada factura para copiar y pegar donde necesites, y un ZIP con "
        "todos los PDFs oficiales del SAT listos para adjuntar en el sistema de "
        "liquidacion.",
        estilos,
    ))
    elementos.append(Spacer(1, 0.25 * cm))
    elementos.append(_subsection("Que datos lleva el Excel", estilos))
    elementos.append(_tabla_reporte(estilos))

    elementos.append(PageBreak())

    # ── 6. Casos especiales: gasolina y comidas ──
    elementos.append(_section_title("5. Casos especiales", estilos))

    elementos.append(_subsection("Gasolina y kilometraje", estilos))
    elementos.append(_paragrafo(
        "Cada vez que cargas combustible el bot te pregunta el numero del odometro. "
        "Con eso calcula automaticamente cuantos kilometros recorriste desde la carga "
        "anterior y lo pone en la columna de kilometraje del reporte.",
        estilos,
    ))
    elementos.append(_paragrafo(
        "Si quieres, puedes fijar un odometro inicial escribiendo "
        "<b>/odometro inicial 120000</b>. Si no lo haces, el bot usa como punto de "
        "partida el valor mas bajo que hayas registrado. Tambien puedes corregir "
        "cualquier kilometraje despues con <b>/odometro</b>.",
        estilos,
    ))
    elementos.append(Spacer(1, 0.2 * cm))

    elementos.append(_subsection("Comidas por dia", estilos))
    elementos.append(_paragrafo(
        "La politica reconoce un Desayuno, un Almuerzo y una Cena por dia. Si envias "
        "dos facturas del mismo tipo en la misma fecha, el bot te avisa para que "
        "verifiques. No te bloquea, solo te recuerda.",
        estilos,
    ))
    elementos.append(Spacer(1, 0.2 * cm))

    elementos.append(_subsection("Si algo quedo mal clasificado", estilos))
    elementos.append(_paragrafo(
        "Si ves que una factura tiene el tipo equivocado (por ejemplo un restaurante "
        "dentro de un hotel que quedo como 'Hospedaje'), escribe <b>/tipo</b>, elige "
        "la factura y selecciona el tipo correcto. No necesitas eliminarla ni "
        "volverla a enviar.",
        estilos,
    ))

    elementos.append(PageBreak())

    # ── 7. Comandos ──
    elementos.append(_section_title("6. Todos los comandos disponibles", estilos))
    elementos.append(_paragrafo(
        "Estos son los comandos que puedes escribir en el chat. Todos empiezan con "
        "la barra (/). No hace falta que los memorices: siempre puedes abrir <b>/help</b> "
        "y te los muestra.",
        estilos,
    ))
    elementos.append(Spacer(1, 0.2 * cm))
    elementos.append(_tabla_comandos(estilos))

    elementos.append(PageBreak())

    # ── 8. Preguntas frecuentes ──
    elementos.append(_section_title("7. Preguntas frecuentes", estilos))
    faqs = [
        ("Y si la foto sale borrosa?",
         "El bot te avisa y te pide que tomes otra. Para mejores resultados, asegurate "
         "de que la luz sea pareja y que la factura se vea completa."),
        ("Puedo mandar un PDF en vez de foto?",
         "Si, perfecto. Incluso es mas rapido porque el bot no tiene que leer la "
         "imagen, lee el texto directo del PDF."),
        ("Que pasa si el SAT esta caido?",
         "El bot reintenta automaticamente varias veces. Si no lo logra, la factura "
         "queda en /fallidas para que la vuelvas a probar despues."),
        ("Puedo corregir una factura ya enviada?",
         "Si. Con <b>/tipo</b> cambias el tipo de gasto, con <b>/odometro</b> arreglas "
         "el kilometraje, y con <b>/eliminar</b> la quitas del registro."),
        ("Mis facturas estan en un lugar seguro?",
         "Cada usuario tiene su carpeta privada en el servidor. Solo tu reporte y "
         "tus PDFs pueden salir desde tu chat."),
        ("Que pasa si mando la misma factura dos veces?",
         "El bot se da cuenta y no la duplica. Te avisa que ya estaba registrada "
         "y te reenvia la confirmacion."),
        ("Puedo generar varios reportes al mes?",
         "Si. Puedes pedir reportes por semana, por quincena o del mes completo. "
         "El archivo se arma sobre las facturas del rango que elijas."),
    ]
    for q, a in faqs:
        elementos.append(Paragraph(
            f'<b><font color="{AZUL_OSCURO.hexval()}">{q}</font></b>', estilos["h3"]
        ))
        elementos.append(_paragrafo(a, estilos))
        elementos.append(Spacer(1, 0.15 * cm))

    elementos.append(HRFlowable(width="100%", thickness=0.5,
                                color=colors.HexColor("#D5DBE3"),
                                spaceBefore=10, spaceAfter=10))
    elementos.append(Paragraph(
        f'<para alignment="center"><font color="{GRIS_TEXTO.hexval()}" size="9">'
        f'Este documento explica el funcionamiento del Robot de Facturas '
        f'desde el punto de vista del usuario.<br/>'
        f'Para soporte tecnico, contacta al administrador.'
        f'</font></para>',
        estilos["body"],
    ))

    # Construir con portada custom en primera pagina
    doc.build(
        elementos,
        onFirstPage=_cover,
        onLaterPages=_footer,
    )


if __name__ == "__main__":
    salida = Path(__file__).parent / "Guia_Robot_Facturas.pdf"
    construir(salida)
    print(f"OK -> {salida}")
