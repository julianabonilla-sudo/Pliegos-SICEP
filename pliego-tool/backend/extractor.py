import subprocess, tempfile, os, json
import anthropic
from dotenv import load_dotenv

load_dotenv()

# ── Constantes fijas BIA Energy ────────────────────────────────────────────────
ENTIDAD       = "BIA ENERGY SAS ESP"
NIT_ENTIDAD   = "901.588.412-3"
OBJETO        = (
    "Suministro de energía y potencia eléctrica con destino "
    "al mercado regulado de BIA ENERGY"
)

# ── Conversión .doc → .docx si es necesario ────────────────────────────────────
def _asegurar_docx(path: str) -> str:
    if path.endswith(".doc"):
        out_dir = tempfile.mkdtemp()
        subprocess.run(
            ["soffice", "--headless", "--convert-to", "docx", "--outdir", out_dir, path],
            check=True, capture_output=True
        )
        nombre = os.path.splitext(os.path.basename(path))[0] + ".docx"
        return os.path.join(out_dir, nombre)
    return path

# ── Extracción de texto del .docx ──────────────────────────────────────────────
def _extraer_texto(path: str) -> str:
    result = subprocess.run(
        ["pandoc", path, "-t", "plain"],
        capture_output=True, text=True
    )
    if result.returncode != 0:
        raise RuntimeError(f"pandoc falló: {result.stderr}")
    return result.stdout

# ── Prompt para Claude ─────────────────────────────────────────────────────────
PROMPT = """Eres un experto en pliegos de convocatorias públicas de energía eléctrica en Colombia (CREG).

Los siguientes campos son FIJOS y NO debes extraerlos del texto:
  entidad: "BIA ENERGY SAS ESP"
  nit_entidad: "901.588.412-3"
  objeto_contrato: "Suministro de energía y potencia eléctrica con destino al mercado regulado de BIA ENERGY"

Extrae ÚNICAMENTE las variables que se listan a continuación del texto del pliego.
Responde SOLO con un objeto JSON válido, sin texto adicional, sin markdown.

ESQUEMA REQUERIDO:
{
  "codigo_convocatoria": "string",
  "periodo_contrato_inicio": "YYYY-MM-DD",
  "periodo_contrato_fin": "YYYY-MM-DD",
  "cronograma": {
    "publicacion_sicep": "YYYY-MM-DD",
    "limite_consultas": "YYYY-MM-DDTHH:MM",
    "pliegos_definitivos": "YYYY-MM-DD",
    "limite_oferta": "YYYY-MM-DDTHH:MM",
    "habilitados_asic": "YYYY-MM-DD",
    "audiencia_publica": "YYYY-MM-DDTHH:MM",
    "max_formalizacion": "YYYY-MM-DD",
    "max_registro_asic": "YYYY-MM-DD"
  },
  "contacto": {
    "nombre": "string",
    "telefono": "string",
    "email": "string"
  },
  "requiere_garantia_seriedad": true,
  "tipo_garantia_seriedad": "string — describe el instrumento y condiciones",
  "productos": [
    {
      "numero": 1,
      "modalidad_suministro": "string",
      "duracion_inicio": "YYYY-MM-DD",
      "duracion_fin": "YYYY-MM-DD",
      "tamano_mwh": 0.0,
      "indexador": "string",
      "metodologia_evaluacion": "mensual | anual | libre",
      "descripcion_metodologia_libre": null,
      "garantia_vendedor": "string",
      "garantia_comprador": "string"
    }
  ]
}

REGLAS para metodologia_evaluacion:
- Si la sección de metodología del producto dice "evaluará cada mes de manera independiente"
  o "precio ofertado podrá ser diferente para cada mes" → usa "mensual"
- Si dice "evaluará las ofertas de manera anual" o "precio ofertado debe ser único para el año" → usa "anual"
- Si no encaja claramente en ninguno de los dos → usa "libre" y pon la descripción en
  descripcion_metodologia_libre (máx. 2 oraciones)

TEXTO DEL PLIEGO:
{texto}
"""

# ── Función principal ──────────────────────────────────────────────────────────
def extraer_variables_pliego(path_docx: str) -> dict:
    path_docx = _asegurar_docx(path_docx)
    texto     = _extraer_texto(path_docx)

    client  = anthropic.Anthropic()
    message = client.messages.create(
        model      = "claude-sonnet-4-6",
        max_tokens = 2048,
        messages   = [
            {
                "role": "user",
                "content": PROMPT.replace("{texto}", texto[:20000])
            }
        ],
    )

    raw = message.content[0].text.strip()
    # Quitar posibles bloques ```json ... ```
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]
    extracted = json.loads(raw)

    # Inyectar los campos fijos (siempre sobreescriben lo que venga)
    extracted["entidad"]        = ENTIDAD
    extracted["nit_entidad"]    = NIT_ENTIDAD
    extracted["objeto_contrato"] = OBJETO

    return extracted
