"""
EnviroAssist AI — Módulo Flora
Analiza flora por cuadrícula UTM. Solo análisis — exportación en app_ieet.py.
"""

import io
import json
from pathlib import Path
from openai import OpenAI
from dotenv import load_dotenv
import pandas as pd
load_dotenv()

SRC_DIR       = Path(__file__).parent
DATA_DIR      = SRC_DIR.parent / "data"
OUTPUT_DIR    = DATA_DIR / "informes"
LLM_MODEL     = "microsoft/phi-4"
LM_STUDIO_URL = "http://localhost:1234/v1"
client = OpenAI(base_url=LM_STUDIO_URL, api_key="lm-studio")


def _es_anthos(ruta: Path) -> bool:
    for enc in ["utf-8", "latin-1", "cp1252"]:
        try:
            with open(str(ruta), encoding=enc, errors="ignore") as f:
                texto = f.read(300).lower()
            if any(k in texto for k in ["taxon", "cuadrí", "cuadri", "listado", "anthos"]):
                return True
        except Exception:
            pass
    return False


class AnalizadorFlora:
    def __init__(self):
        self.taxones: list = []
        self.fuente: str = ""
        self.cuadricula: str = ""
        self.ccaa: str = ""

    def cargar_anthos(self, ruta: Path, cuadricula: str = "") -> list:
        self.cuadricula = cuadricula.strip().upper()
        taxones = []
        for enc in ["utf-8", "latin-1", "cp1252", "iso8859-16", "windows-1252"]:
            try:
                with open(str(ruta), encoding=enc, errors="ignore") as f:
                    lineas = f.readlines()
                for l in lineas:
                    l = l.strip()
                    if not l or len(l) < 4:
                        continue
                    l_low = l.lower()
                    if any(k in l_low for k in ["listado","cuadrícula","cuadricula","taxon","ficha"]):
                        if "ficha del taxon" in l_low:
                            nombre = l.replace("Ficha del taxon","").replace("ficha del taxon","").strip()
                            if nombre:
                                taxones.append(nombre)
                        continue
                    if l.startswith("30T") or l.startswith("29T"):
                        continue
                    if l[0].isupper() and " " in l:
                        taxones.append(l)
                if taxones:
                    break
            except Exception:
                continue
        self.taxones = taxones
        self.fuente = f"Anthos — cuadrícula {self.cuadricula}"
        print(f"✅ Anthos: {len(self.taxones)} taxones")
        return self.taxones

    def cargar_csv_flora(self, ruta: Path, cuadricula: str = "") -> list:
        self.cuadricula = cuadricula.strip().upper()
        if _es_anthos(ruta):
            return self.cargar_anthos(ruta, cuadricula)
        raw = ruta.read_bytes()
        df = None
        for enc in ["latin-1","cp1252","utf-8-sig","utf-8"]:
            try:
                text = raw.decode(enc)
                for sep in [";",",","\t"]:
                    try:
                        df_t = pd.read_csv(io.StringIO(text), sep=sep, on_bad_lines="skip")
                        if len(df_t.columns) > 1:
                            df = df_t
                            break
                    except Exception:
                        continue
                if df is not None:
                    break
            except Exception:
                continue
        if df is None:
            raise ValueError(f"No se pudo leer {ruta.name}")
        df.columns = [c.strip().lower() for c in df.columns]
        col = next((c for c in df.columns if c in ["nombre","especie","taxon","nombre_cientifico","species"]), df.columns[0])
        self.taxones = df[col].dropna().str.strip().unique().tolist()
        self.fuente = ruta.name
        return self.taxones

    def generar_redaccion(self, proyecto: str, alertas: list) -> str:
        cuad = self.cuadricula or "—"
        # Muestra de taxones REALES presentes en la cuadrícula, priorizando
        # los protegidos, para que phi-4 describa la flora que de verdad hay
        # (y no invente especies de matorral mediterráneo genérico que a
        # veces contradicen el hábitat real, p.ej. taxones de humedal).
        especies_protegidas = {a["especie"] for a in alertas}
        protegidos = [t for t in self.taxones if t in especies_protegidas]
        resto = [t for t in self.taxones if t not in especies_protegidas]
        muestra = (protegidos + resto)[:15]
        muestra_txt = "; ".join(muestra) or "sin datos"
        alertas_txt = "\n".join([
            f"- {a['especie']} ({a.get('nombre_comun','')}): "
            f"LESRPE={a.get('cat_lesrpe','—')} | {self.ccaa}={a.get('cat_auto_nombre','—')}"
            for a in alertas[:15]
        ]) or "No se detectaron taxones protegidos."
        prompt = f"""Eres un consultor ambiental especialista en botánica con 15 años de experiencia en EsIA.
Redacta el inventario florístico para el EsIA del proyecto "{proyecto}".

Fuente: {self.fuente} | Cuadrícula UTM: {cuad}
Total taxones: {len(self.taxones)}

Taxones reales presentes en esta cuadrícula (usa EXCLUSIVAMENTE estos como
ejemplos; no inventes especies nuevas ni describas un tipo de hábitat o
ecosistema distinto del que reflejan realmente estos taxones). No se dispone
de nombre común para ellos — nómbralos siempre por su nombre científico en
cursiva, sin inventar un nombre común en español:
{muestra_txt}

Taxones con estatus de protección:
{alertas_txt}

Estructura: Metodología (menciona Anthos/IEET), Resultados florísticos,
Taxones protegidos (cita RD 139/2011 y normativa autonómica), Valoración.
Redacta SIEMPRE en prosa narrativa, con frases completas — NUNCA uses listas
con viñetas ni guiones para enumerar taxones, intégralos dentro de las frases
como lo haría un consultor ambiental en un EsIA real.
Extensión: 300-400 palabras. Tono técnico de EsIA."""
        try:
            resp = client.chat.completions.create(
                model=LLM_MODEL,
                messages=[{"role": "user", "content": prompt}],
                temperature=0.15, max_tokens=1000,
            )
            return resp.choices[0].message.content
        except Exception as e:
            return f"[Error LM Studio: {e}]"
