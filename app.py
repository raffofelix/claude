"""
ICB Prospecta — Web App
Investiga un cliente gastronómico y entrega un Excel con los Top 15 SKUs de ICB Food Service.

Requisitos:
    pip install flask anthropic pandas openpyxl requests

Variables de entorno:
    ANTHROPIC_API_KEY   — clave API de Anthropic (obligatoria)
    PORT                — puerto del servidor (default: 5000)

Uso:
    python app.py
    Abrir http://localhost:5000 en cualquier navegador o celular de la misma red.
"""

import os, io, json, glob, tempfile, urllib.request
from datetime import datetime
from pathlib import Path
from flask import Flask, request, jsonify, send_file, render_template
import anthropic
import pandas as pd

# ── Importar generador de Excel ──────────────────────────────────────────────
import sys
sys.path.insert(0, str(Path(__file__).parent))
from generar_excel import seleccionar_top15, generar_excel as _generar_excel

app = Flask(__name__)

# ── Configuración ────────────────────────────────────────────────────────────
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
CATALOGO_URL = (
    "https://1drv.ms/x/c/91cea5c809fe1373/"
    "IQAoUK8ItuU8R4k0BCHZhvIBAZubGKc0-FvXZuOJ-2-296o?e=JAf6qD&download=1"
)
CATALOGO_CACHE = Path(tempfile.gettempdir()) / "icb_catalogo.xlsx"

# ── Catálogo ─────────────────────────────────────────────────────────────────

def obtener_catalogo() -> pd.DataFrame:
    """Descarga o usa caché del catálogo ICB."""
    ruta = str(CATALOGO_CACHE)
    if not CATALOGO_CACHE.exists():
        try:
            urllib.request.urlretrieve(CATALOGO_URL, ruta)
        except Exception as e:
            raise RuntimeError(f"No se pudo descargar el catálogo ICB: {e}")
    xl = pd.ExcelFile(ruta)
    frames = []
    for sheet in xl.sheet_names:
        df = pd.read_excel(xl, sheet_name=sheet)
        df["BASE"] = sheet.upper()
        frames.append(df)
    df = pd.concat(frames, ignore_index=True)
    df["CÓDIGO"]    = df["CÓDIGO"].astype(str)
    df["PRECIO"]    = pd.to_numeric(df.get("PRECIO"),    errors="coerce").fillna(0)
    df["PRECIO UN"] = pd.to_numeric(df.get("PRECIO UN"), errors="coerce").fillna(0)
    return df


def limpiar_cache_catalogo():
    """Fuerza re-descarga en el próximo uso."""
    if CATALOGO_CACHE.exists():
        CATALOGO_CACHE.unlink()

# ── Prompt del sistema ───────────────────────────────────────────────────────

SYSTEM_PROMPT = """Eres un experto asesor de ventas de ICB Food Service Chile, el mayor distribuidor
de insumos gastronómicos del país (más de 1.500 SKUs, clientes desde restaurantes hasta casinos).

Cuando te den el nombre de un cliente gastronómico chileno, debes:

1. INVESTIGAR el negocio con búsquedas web: tipo de establecimiento, qué venden, carta/menú,
   tamaño, cuántos locales, canales (delivery, catering, etc.).

2. CLASIFICAR el cliente en uno de estos tipos:
   PASTELERÍA/BAKERY | CAFÉ ESPECIALIDAD | RESTAURANTE ALMUERZO | RESTAURANTE FINE DINING |
   HOTEL/CATERING | FAST FOOD/HAMBURGUESERÍA | PIZZERÍA | SUSHI/JAPONÉS |
   CASINO/COLECTIVIDAD | HELADERÍA | BAR/COCTELERÍA | CAFÉ TRADICIONAL

3. SELECCIONAR las 15 familias de productos ICB más relevantes, con este criterio:
   - P1 (prioridad 1): insumos core, el cliente no puede operar sin ellos
   - P2 (prioridad 2): diferenciadores de calidad o experiencia
   - P3 (prioridad 3): complementos útiles

   Familias disponibles en el catálogo ICB:
   ACEITES Y VINAGRES, ARROZ, AZUCAR Y ENDULZANTES, BOLLERIA PARA HORNEAR, BOLLERÍA LISTA,
   CAFÉ EN GRANO, CAFÉ MOLIDO, CAFÉ INSTANTANEO, CAPSULA LAVAZZA BLUE, CAPSULAS COMPATIBLES,
   CARNES (CERDO/VACUNO/POLLO/PAVO/CORDERO), CERTIFIED ANGUS BEEF, COBERTURA, CONDIMENTOS,
   CONFITERIA, CREMAS, DELIVERY, ELABORADOS CERDO/POLLO/VACUNO, FRUTOS SECOS, FRUTAS CONGELADAS,
   HAMBURGUESAS, HARINAS Y SEMOLAS, HELADOS, HUEVOS, INSUMO DE REPOSTERIA, JARABES Y SYRUPS,
   JUGOS Y PULPAS, KETCHUP, LECHES, MANTEQUILLAS, MARGARINA, MAYONESA, MIEL Y MERMELADAS,
   MOSTAZA, MOZZARELLA, OTRAS CECINAS, OTRAS SALSAS, PAPAS PRE FRITAS, PASTAS,
   PARMESANO / MADURO, POSTRES EN POLVO, PRE MEZCLA, PULPAS, QUESO CREMA, QUESOS ESPECIALIDAD,
   RELLENOS, SABORIZANTE, SALMONES, SAL, SALSAS DE TOMATE, SALSAS DULCES, SNACKS,
   TE E INFUSIONES, TORTAS, VASOS Y TAPAS

4. Para cada familia, indica el código de producto preferido si puedes inferirlo
   (ej: para CAFÉ EN GRANO de un local premium → código Lavazza Orgánico 104519130).
   Si no puedes inferir el código, deja null.

RESPONDE ÚNICAMENTE con este JSON (sin texto adicional, sin markdown):
{
  "cliente": "nombre del cliente",
  "tipo": "TIPO DE NEGOCIO",
  "perfil": "descripción breve del negocio en 1-2 oraciones",
  "angulo_entrada": "por dónde entrar en la primera visita",
  "pregunta_discovery": "1 pregunta clave para hacer al contacto",
  "familias": [
    {
      "familia": "NOMBRE_FAMILIA",
      "prioridad": "1",
      "motivo": "por qué este producto para este cliente específico",
      "codigo_preferido": "123456789 o null"
    }
  ]
}"""

# ── Investigación con Claude ─────────────────────────────────────────────────

def investigar_cliente(nombre_cliente: str) -> dict:
    """Llama a Claude para investigar el cliente y retorna la selección de familias."""
    if not ANTHROPIC_API_KEY:
        raise RuntimeError("Falta la variable de entorno ANTHROPIC_API_KEY.")

    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    response = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=4096,
        system=SYSTEM_PROMPT,
        tools=[{"type": "web_search_20250305", "name": "web_search", "max_uses": 5}],
        messages=[{
            "role": "user",
            "content": (
                f"Investiga este cliente de ICB Food Service Chile: {nombre_cliente}\n\n"
                "Busca su carta, tipo de negocio y retorna el JSON con las 15 familias ICB "
                "más relevantes para ofrecerle."
            )
        }]
    )

    # Extraer texto final del response
    texto = ""
    for block in response.content:
        if hasattr(block, "text"):
            texto += block.text

    # Parsear JSON de la respuesta
    texto = texto.strip()
    if texto.startswith("```"):
        texto = texto.split("```")[1]
        if texto.startswith("json"):
            texto = texto[4:]
    texto = texto.strip()

    return json.loads(texto)


# ── Rutas ─────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/investigar", methods=["POST"])
def api_investigar():
    """Investiga el cliente y retorna el Excel como descarga."""
    data = request.get_json()
    nombre = (data or {}).get("cliente", "").strip()
    if not nombre:
        return jsonify({"error": "Falta el nombre del cliente."}), 400

    try:
        # 1. Investigar con Claude
        resultado = investigar_cliente(nombre)

        # 2. Cargar catálogo
        df = obtener_catalogo()

        # 3. Generar Excel
        nombre_archivo = (
            nombre.replace(" ", "_").replace("/", "-")[:40]
            + f"_Top15_ICB_{datetime.today().strftime('%Y%m')}.xlsx"
        )
        ruta_out = Path(tempfile.gettempdir()) / nombre_archivo

        skus = seleccionar_top15(df, resultado["familias"])
        _generar_excel(skus, resultado["cliente"], resultado["tipo"], str(ruta_out))

        # 4. Retornar metadata + nombre de archivo para descarga
        return jsonify({
            "ok": True,
            "archivo": nombre_archivo,
            "perfil": resultado.get("perfil", ""),
            "tipo": resultado.get("tipo", ""),
            "angulo_entrada": resultado.get("angulo_entrada", ""),
            "pregunta_discovery": resultado.get("pregunta_discovery", ""),
            "n_skus": len(skus),
        })

    except json.JSONDecodeError as e:
        return jsonify({"error": f"Claude devolvió una respuesta inesperada. Intenta de nuevo."}), 500
    except RuntimeError as e:
        return jsonify({"error": str(e)}), 500
    except Exception as e:
        return jsonify({"error": f"Error inesperado: {str(e)}"}), 500


@app.route("/api/descargar/<nombre_archivo>")
def api_descargar(nombre_archivo: str):
    """Descarga el Excel generado."""
    ruta = Path(tempfile.gettempdir()) / nombre_archivo
    if not ruta.exists():
        return jsonify({"error": "Archivo no encontrado. Vuelve a generar."}), 404
    return send_file(
        str(ruta),
        as_attachment=True,
        download_name=nombre_archivo,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@app.route("/api/actualizar-catalogo", methods=["POST"])
def api_actualizar_catalogo():
    """Fuerza re-descarga del catálogo desde OneDrive."""
    limpiar_cache_catalogo()
    try:
        obtener_catalogo()
        return jsonify({"ok": True, "mensaje": "Catálogo actualizado correctamente."})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ── Main ─────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print(f"\n  ICB Prospecta corriendo en http://localhost:{port}")
    print(f"  En tu celular (misma red): http://<IP-de-tu-PC>:{port}\n")
    app.run(host="0.0.0.0", port=port, debug=False)
