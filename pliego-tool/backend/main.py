from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel
from typing import Optional
import tempfile, os, json

from extractor import extraer_variables_pliego
from generator import generar_pliego

app = FastAPI(title="Pliego Automatizador - BIA Energy")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# ── Modelos ────────────────────────────────────────────────────────────────────

class Producto(BaseModel):
    numero: int
    modalidad_suministro: str
    duracion_inicio: str          # YYYY-MM-DD
    duracion_fin: str             # YYYY-MM-DD
    tamano_mwh: float
    indexador: str
    metodologia_evaluacion: str   # "mensual" | "anual" | "libre"
    descripcion_metodologia_libre: Optional[str] = None
    garantia_vendedor: str
    garantia_comprador: str

class Cronograma(BaseModel):
    publicacion_sicep: str
    limite_consultas: str
    pliegos_definitivos: str
    limite_oferta: str
    habilitados_asic: str
    audiencia_publica: str
    max_formalizacion: str
    max_registro_asic: str

class Contacto(BaseModel):
    nombre: str
    telefono: str
    email: str

class PliegoData(BaseModel):
    # Fijos — siempre BIA Energy
    entidad: str = "BIA ENERGY SAS ESP"
    nit_entidad: str = "901.588.412-3"
    objeto_contrato: str = (
        "Suministro de energía y potencia eléctrica con destino "
        "al mercado regulado de BIA ENERGY"
    )
    # Variables por convocatoria
    codigo_convocatoria: str
    periodo_contrato_inicio: str
    periodo_contrato_fin: str
    mes_indexacion: str              # Ej: "marzo de 2026" — único para toda la convocatoria
    cronograma: Cronograma
    contacto: Contacto
    tipo_pliego: str = "definitivos"          # "definitivos" | "preliminares"
    requiere_garantia_seriedad: bool = True
    tipo_garantia_seriedad: str
    garantia_cumplimiento_tipo: str = "pagare_garantia"  # "pagare" | "pagare_garantia"
    # Productos (1..N)
    productos: list[Producto]


# ── Endpoints ──────────────────────────────────────────────────────────────────

@app.post("/extraer")
async def extraer(file: UploadFile = File(...)):
    """Recibe un .docx/.doc y devuelve las variables extraídas por Claude."""
    if not file.filename.endswith((".docx", ".doc")):
        raise HTTPException(400, "Solo se aceptan archivos .docx o .doc")

    suffix = ".docx" if file.filename.endswith(".docx") else ".doc"
    with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
        tmp.write(await file.read())
        tmp_path = tmp.name

    try:
        variables = extraer_variables_pliego(tmp_path)
        return variables
    except Exception as e:
        raise HTTPException(500, f"Error al extraer variables: {str(e)}")
    finally:
        os.unlink(tmp_path)


@app.post("/generar")
async def generar(data: PliegoData):
    """Recibe las variables (editadas por el usuario) y genera el .docx final."""
    output_path = tempfile.mktemp(suffix=".docx")
    try:
        generar_pliego(data.dict(), output_path)
        return FileResponse(
            output_path,
            filename=f"PLIEGO_{data.codigo_convocatoria}.docx",
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
    except Exception as e:
        raise HTTPException(500, f"Error al generar pliego: {str(e)}")


@app.get("/health")
def health():
    return {"status": "ok"}
