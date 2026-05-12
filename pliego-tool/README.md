# Automatizador de Pliegos — BIA Energy

Herramienta para extraer variables de pliegos .docx y generar nuevos documentos.

## Estructura

```
pliego-tool/
├── backend/
│   ├── main.py          # API FastAPI
│   ├── extractor.py     # Extracción con Claude
│   ├── generator.py     # Generación del .docx
│   └── requirements.txt
├── frontend/
│   └── index.html       # UI (abrir directo en browser)
└── templates/
    └── plantilla_base.docx   # (opcional) plantilla con marcadores {{variable}}
```

## Instalación

```bash
# 1. Instalar dependencias Python
cd backend
pip install -r requirements.txt

# 2. Configurar API key de Anthropic
export ANTHROPIC_API_KEY="sk-ant-..."

# 3. Instalar pandoc (para convertir .docx a texto)
# macOS:
brew install pandoc
# Ubuntu/Debian:
sudo apt install pandoc

# 4. (Opcional) LibreOffice para convertir .doc legacy
# macOS:
brew install --cask libreoffice
# Ubuntu:
sudo apt install libreoffice
```

## Uso

```bash
# Iniciar backend
cd backend
uvicorn main:app --reload --port 8000

# Abrir frontend
# Abre frontend/index.html en el browser (doble clic o con Live Server en VSCode/Cursor)
```

## Flujo de la herramienta

1. **Cargar pliego** — arrastra un .docx existente → Claude extrae las variables
2. **Revisar variables** — edita lo que necesites (cronograma, contacto, garantías)
3. **Productos** — para cada producto elige la metodología:
   - 📅 **Mensual**: precio diferente por mes, reserva mensual, % variable
   - 📆 **Anual**: precio único por año, reserva anual, % fijo
   - ✏️ **Libre**: escribes la metodología manualmente
4. **Generar** — descarga el .docx final

## Variables fijas (nunca cambian)

| Campo | Valor |
|-------|-------|
| `entidad` | BIA ENERGY SAS ESP |
| `nit_entidad` | 901.588.412-3 |
| `objeto_contrato` | Suministro de energía y potencia eléctrica con destino al mercado regulado de BIA ENERGY |

## Plantilla personalizada (opcional)

Si quieres usar tu propio formato Word, crea `templates/plantilla_base.docx`
con estos marcadores en el texto:

```
{{codigo_convocatoria}}     {{periodo_inicio}}          {{periodo_fin}}
{{contacto_nombre}}         {{contacto_telefono}}       {{contacto_email}}
{{garantia_seriedad}}
{{fecha_publicacion_sicep}} {{fecha_limite_consultas}}  {{fecha_pliegos_definitivos}}
{{fecha_limite_oferta}}     {{fecha_habilitados_asic}}  {{fecha_audiencia_publica}}
{{fecha_max_formalizacion}} {{fecha_max_registro_asic}}
```

Si no existe la plantilla, el generador crea el documento desde cero.
