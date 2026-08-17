# 🌿 EnviroAssist AI

Asistente inteligente para el inventario biótico (fauna y flora) en Estudios de
Impacto Ambiental (EsIA). Cruza automáticamente los datos de campo (BDIEET del
MITECO, listados Anthos) con el LESRPE/CEEA nacional y el catálogo autonómico
correspondiente, y redacta el apartado de inventario biótico con un LLM local.

Trabajo Final de Máster — Máster en Inteligencia Artificial Aplicada.

---

## Estructura del proyecto

```
EnviroAssistAI/
├── data/
│   ├── normativa/
│   │   └── lesrpe_2025.json     ← LESRPE/CEEA (1010 taxones, MITECO junio 2025)
│   ├── campo/
│   │   └── inventario_demo.csv  ← Inventario de demo para pruebas
│   └── informes/                ← Salidas generadas (.md, .docx, .xlsx)
├── src/
│   ├── app_ieet.py              ← Interfaz web (Streamlit) — punto de entrada
│   ├── modulo_fauna.py          ← Análisis de fauna (BDIEET) por cuadrícula UTM
│   ├── modulo_flora.py          ← Análisis de flora (Anthos) por cuadrícula UTM
│   ├── normativa_autonomica.py  ← Carga y extracción de catálogos autonómicos
│   ├── exportar_word.py         ← Exportación a .docx
│   └── exportar_excel.py        ← Exportación a .xlsx
├── requirements.txt
└── README.md
```

---

## Requisitos previos

- **Python 3.13**
- **LM Studio** con el modelo `microsoft/phi-4` (cuantización Q3_K_L, 7.93 GB)
  descargado y el servidor local activo en `http://127.0.0.1:1234`.

> ⚠️ LM Studio tiene que estar abierto y con el servidor en marcha **cada vez**
> que uses la app. Puede dejarse minimizado en segundo plano.

No se necesita ninguna API key de OpenAI, Anthropic ni de ningún otro proveedor
externo: todo el procesamiento con LLM (extracción de especies desde PDF y
redacción del inventario biótico) se ejecuta contra el servidor local de LM
Studio. El SDK de OpenAI se usa únicamente como cliente para hablar con ese
servidor local, que expone una interfaz compatible con la API de OpenAI — no
se realiza ninguna llamada a servidores externos.

---

## Instalación

```bash
# 1. Clonar / descargar el proyecto
cd EnviroAssistAI

# 2. Crear entorno virtual
py -3.13 -m venv venv
venv\Scripts\activate        # Windows
# source venv/bin/activate   # Linux/Mac

# 3. Instalar dependencias
pip install -r src/requirements.txt
```

### Configurar LM Studio

1. Abre LM Studio y descarga `microsoft/phi-4` (cuantización Q3_K_L).
2. Ve a la pestaña **Developer** y selecciona `microsoft/phi-4` en el modelo del servidor.
3. Activa **Status: Running**. Confirma que aparece `Reachable at: http://127.0.0.1:1234`.

---

## Uso

Con LM Studio arrancado y el entorno virtual activado:

```bash
streamlit run src/app_ieet.py
```

Se abre automáticamente en `http://localhost:8501`. Si no se abre solo, entra
manualmente a esa dirección.

### Flujo típico

1. **Panel lateral**: nombre del proyecto, ubicación, comunidad autónoma y tipo de proyecto.
2. **Pestaña ⚙️ Normativa autonómica**: selecciona la CCAA y sube el catálogo
   oficial (Excel, CSV, PDF o Word — el formato más fiable depende del
   catálogo concreto, no hay un formato universalmente mejor).
3. **Pestaña 🦅 Fauna — BDIEET/IEET**: sube el BDIEET, indica la(s) cuadrícula(s)
   UTM y analiza. El sistema cruza automáticamente con la normativa cargada y
   genera la redacción con phi-4.
4. **Pestaña 🌿 Flora — Anthos/IEET**: mismo proceso con un listado Anthos.
5. Descarga el informe en `.md`, `.docx` o `.xlsx`.

---

## Arquitectura

El cruce normativo (qué especies están protegidas y con qué categoría) es una
operación **determinista de cruce de tablas con Pandas, sin intervención del
LLM** — esto evita que una alucinación del modelo pueda alterar una
clasificación normativa. El LLM (phi-4, temperatura 0.15) interviene
únicamente para:

- Extraer especies y categorías de catálogos autonómicos en PDF o Word sin tablas.
- Redactar el texto del inventario biótico a partir de los resultados ya verificados del cruce.

Todos los documentos generados incluyen el aviso: *"Borrador generado con IA.
Revisar y validar antes de presentación oficial."* — el sistema está pensado
como apoyo al técnico ambiental, no como sustituto de su criterio profesional.

---

## Limitaciones conocidas

- Prototipo probado exhaustivamente con el catálogo CREA Madrid (Decreto
  18/1992); otras comunidades autónomas siguen el mismo proceso pero con
  menor volumen de pruebas.
- No incluye la redacción de otros apartados del EsIA (medio físico,
  socioeconómico, valoración de impactos, medidas correctoras).
- No hay integración con SIG: la cuadrícula UTM se introduce manualmente.

Ver el TFM completo para el detalle de diseño, evaluación y líneas futuras.
