from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import datetime
import os

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "..", "templates", "plantilla_base.docx")

# ── Helpers ────────────────────────────────────────────────────────────────────

def _fmt_fecha(iso: str) -> str:
    """'2026-04-24' → '24 de abril de 2026'"""
    MESES = [
        "", "enero","febrero","marzo","abril","mayo","junio",
        "julio","agosto","septiembre","octubre","noviembre","diciembre"
    ]
    try:
        d = datetime.fromisoformat(iso.replace("T", " ").split(" ")[0])
        return f"{d.day} de {MESES[d.month]} de {d.year}"
    except Exception:
        return iso

def _fmt_fecha_hora(iso: str) -> str:
    """'2026-04-30T18:00' → '30 de abril de 2026 a las 18:00 horas'"""
    MESES = [
        "", "enero","febrero","marzo","abril","mayo","junio",
        "julio","agosto","septiembre","octubre","noviembre","diciembre"
    ]
    try:
        dt = datetime.fromisoformat(iso.replace("T", " "))
        return f"{dt.day} de {MESES[dt.month]} de {dt.year} a las {dt.strftime('%H:%M')} horas"
    except Exception:
        return iso

def _metodologia_texto(producto: dict) -> str:
    met = producto.get("metodologia_evaluacion", "libre")
    if met == "mensual":
        return (
            "Evaluación mensual: el precio ofertado puede ser diferente para cada mes "
            "del producto ($/kWh, hasta dos decimales). BIA ENERGY presentará oferta de "
            "reserva mensual y evaluará cada mes de manera independiente."
        )
    elif met == "anual":
        return (
            "Evaluación anual: el precio ofertado debe ser único para el año del "
            "producto ($/kWh, hasta dos decimales). BIA ENERGY presentará oferta de "
            "reserva para el periodo solicitado y evaluará las ofertas de manera anual."
        )
    else:
        return producto.get(
            "descripcion_metodologia_libre",
            "Metodología definida por BIA ENERGY conforme al pliego de condiciones."
        )

# ── Reemplazo de marcadores en párrafos y tablas ───────────────────────────────

def _reemplazar_en_parrafos(parrafos, marcadores: dict):
    for p in parrafos:
        texto_completo = "".join(r.text for r in p.runs)
        for clave, valor in marcadores.items():
            if clave in texto_completo:
                texto_completo = texto_completo.replace(clave, str(valor))
        # Escribir el resultado en el primer run y limpiar los demás
        if p.runs:
            p.runs[0].text = texto_completo
            for r in p.runs[1:]:
                r.text = ""

def _reemplazar_en_doc(doc: Document, marcadores: dict):
    _reemplazar_en_parrafos(doc.paragraphs, marcadores)
    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                _reemplazar_en_parrafos(celda.paragraphs, marcadores)
    for seccion in doc.sections:
        _reemplazar_en_parrafos(seccion.header.paragraphs, marcadores)
        _reemplazar_en_parrafos(seccion.footer.paragraphs, marcadores)

# ── Construcción de la tabla de productos ─────────────────────────────────────

def _agregar_tabla_productos(doc: Document, productos: list):
    doc.add_heading("Descripción de los Productos a Contratar", level=2)

    for prod in productos:
        doc.add_heading(f"Producto {prod['numero']}", level=3)

        tabla = doc.add_table(rows=1, cols=2)
        tabla.style = "Table Grid"

        def _fila(campo, valor):
            fila = tabla.add_row()
            fila.cells[0].text = campo
            fila.cells[1].text = str(valor)
            fila.cells[0].paragraphs[0].runs[0].bold = True

        _fila("Modalidad de suministro", prod["modalidad_suministro"])
        _fila("Duración del suministro",
              f"{_fmt_fecha(prod['duracion_inicio'])} a {_fmt_fecha(prod['duracion_fin'])}")
        _fila("Tamaño del negocio jurídico (MWh)", f"{prod['tamano_mwh']:,.0f}")
        _fila("Indexador", prod["indexador"])
        _fila("Metodología de evaluación", _metodologia_texto(prod))
        _fila("Garantía vendedor", prod["garantia_vendedor"])
        _fila("Garantía comprador", prod["garantia_comprador"])

        doc.add_paragraph("")  # espacio entre productos

# ── Función principal ──────────────────────────────────────────────────────────

def generar_pliego(data: dict, output_path: str):
    cr = data["cronograma"]
    ct = data["contacto"]

    # Marcadores planos para reemplazo en plantilla
    marcadores = {
        "{{entidad}}":               data["entidad"],
        "{{nit_entidad}}":           data["nit_entidad"],
        "{{objeto_contrato}}":       data["objeto_contrato"],
        "{{codigo_convocatoria}}":   data["codigo_convocatoria"],
        "{{periodo_inicio}}":        _fmt_fecha(data["periodo_contrato_inicio"]),
        "{{periodo_fin}}":           _fmt_fecha(data["periodo_contrato_fin"]),
        "{{contacto_nombre}}":       ct["nombre"],
        "{{contacto_telefono}}":     ct["telefono"],
        "{{contacto_email}}":        ct["email"],
        "{{garantia_seriedad}}":     data.get("tipo_garantia_seriedad", ""),
        # Cronograma
        "{{fecha_publicacion_sicep}}":   _fmt_fecha(cr["publicacion_sicep"]),
        "{{fecha_limite_consultas}}":    _fmt_fecha_hora(cr["limite_consultas"]),
        "{{fecha_pliegos_definitivos}}": _fmt_fecha(cr["pliegos_definitivos"]),
        "{{fecha_limite_oferta}}":       _fmt_fecha_hora(cr["limite_oferta"]),
        "{{fecha_habilitados_asic}}":    _fmt_fecha(cr["habilitados_asic"]),
        "{{fecha_audiencia_publica}}":   _fmt_fecha_hora(cr["audiencia_publica"]),
        "{{fecha_max_formalizacion}}":   _fmt_fecha(cr["max_formalizacion"]),
        "{{fecha_max_registro_asic}}":   _fmt_fecha(cr["max_registro_asic"]),
    }

    # Cargar plantilla o crear documento nuevo si no existe
    if os.path.exists(TEMPLATE_PATH):
        doc = Document(TEMPLATE_PATH)
        _reemplazar_en_doc(doc, marcadores)
    else:
        doc = _crear_doc_desde_cero(data, marcadores)

    # Agregar tabla de productos
    _agregar_tabla_productos(doc, data["productos"])

    doc.save(output_path)


# ── Documento desde cero (cuando no hay plantilla) ────────────────────────────

def _crear_doc_desde_cero(data: dict, marcadores: dict) -> Document:
    doc = Document()
    cr  = data["cronograma"]
    ct  = data["contacto"]

    # Encabezado
    titulo = doc.add_heading("PLIEGO DE CONDICIONES DEFINITIVOS", level=1)
    titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph(f"Código de convocatoria: {data['codigo_convocatoria']}")
    doc.add_paragraph(
        f"Período a contratar: {marcadores['{{periodo_inicio}}']} "
        f"a {marcadores['{{periodo_fin}}']}"
    )

    # Objeto
    doc.add_heading("1. Objeto", level=2)
    doc.add_paragraph(
        f"La presente convocatoria tiene por objeto: {data['objeto_contrato']}."
    )

    # Contacto
    doc.add_heading("2. Funcionario encargado", level=2)
    doc.add_paragraph(f"Nombre: {ct['nombre']}")
    doc.add_paragraph(f"Teléfono: {ct['telefono']}")
    doc.add_paragraph(f"Correo: {ct['email']}")

    # Cronograma
    doc.add_heading("3. Cronograma del proceso", level=2)
    tabla_cr = doc.add_table(rows=1, cols=2)
    tabla_cr.style = "Table Grid"
    encabezado = tabla_cr.rows[0].cells
    encabezado[0].text = "Actividad"
    encabezado[1].text = "Fecha"
    for celda in encabezado:
        celda.paragraphs[0].runs[0].bold = True

    filas_cr = [
        ("Publicación aviso SICEP",           marcadores["{{fecha_publicacion_sicep}}"]),
        ("Límite consultas a pliegos",         marcadores["{{fecha_limite_consultas}}"]),
        ("Publicación pliegos definitivos",    marcadores["{{fecha_pliegos_definitivos}}"]),
        ("Límite entrega oferta",              marcadores["{{fecha_limite_oferta}}"]),
        ("Remisión habilitados al ASIC",       marcadores["{{fecha_habilitados_asic}}"]),
        ("Audiencia pública",                  marcadores["{{fecha_audiencia_publica}}"]),
        ("Máxima formalización de ofertas",    marcadores["{{fecha_max_formalizacion}}"]),
        ("Máximo registro ante ASIC",          marcadores["{{fecha_max_registro_asic}}"]),
    ]
    for actividad, fecha in filas_cr:
        f = tabla_cr.add_row()
        f.cells[0].text = actividad
        f.cells[1].text = fecha

    doc.add_paragraph("")

    # Garantía de seriedad
    doc.add_heading("4. Garantía de seriedad de la oferta", level=2)
    doc.add_paragraph(data.get("tipo_garantia_seriedad", ""))

    return doc
