"""
EnviroAssist AI — Módulo Fauna
Analiza especies de fauna por cuadrícula UTM 10x10 km (BDIEET/IEET).
Solo análisis — la exportación se hace en app_ieet.py.
"""

import io
import json
import math
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


def _leer_archivo(ruta: Path) -> pd.DataFrame:
    ext = ruta.suffix.lower()
    if ext in (".xlsx", ".xls"):
        return pd.read_excel(ruta, engine="openpyxl")
    raw = ruta.read_bytes()
    for enc in ["latin-1", "cp1252", "utf-8-sig", "utf-8"]:
        try:
            text = raw.decode(enc)
            for sep in [";", ",", "\t"]:
                try:
                    df = pd.read_csv(io.StringIO(text), sep=sep,
                                     on_bad_lines="skip", low_memory=False)
                    if len(df.columns) > 1 and len(df) > 0:
                        return df
                except Exception:
                    continue
        except Exception:
            continue
    raise ValueError(f"No se pudo leer {ruta.name}")


class AnalizadorFauna:
    def __init__(self):
        self.df_bdieet: pd.DataFrame = None
        self.df: pd.DataFrame = None
        self.cuadriculas: list = []
        self.estadisticas: dict = {}
        self.ccaa: str = ""

    def cargar_bdieet(self, ruta: Path) -> pd.DataFrame:
        ruta = Path(ruta)
        print(f"📂 Cargando BDIEET: {ruta.name}")
        df = _leer_archivo(ruta)
        df.columns = [c.strip().lower() for c in df.columns]
        print(f"✅ {len(df)} registros | cols: {list(df.columns[:5])}...")
        self.df_bdieet = df
        return df

    def filtrar_cuadriculas(self, cuadriculas: list) -> pd.DataFrame:
        if self.df_bdieet is None:
            raise ValueError("Carga el BDIEET primero.")
        col_cuad = next(
            (c for c in self.df_bdieet.columns
             if any(k in c for k in ["cutm","cuadricula","grid","utm"])),
            None
        )
        if col_cuad is None:
            raise ValueError("No se encontró columna de cuadrícula UTM.")
        self.cuadriculas = [c.strip().upper() for c in cuadriculas]
        mask = self.df_bdieet[col_cuad].astype(str).str.upper().str.strip().isin(self.cuadriculas)
        self.df = self.df_bdieet[mask].copy()
        if "genero" in self.df.columns and "especie" in self.df.columns:
            self.df["nombre_cientifico"] = (
                self.df["genero"].fillna("").str.strip() + " " +
                self.df["especie"].fillna("").str.strip()
            ).str.strip()
        elif "nombre" in self.df.columns:
            self.df["nombre_cientifico"] = self.df["nombre"].fillna("").str.strip()
        else:
            self.df["nombre_cientifico"] = "Desconocida"
        if "grupo" not in self.df.columns:
            for alt in ["clase", "orden"]:
                if alt in self.df.columns:
                    self.df["grupo"] = self.df[alt].fillna("No especificado")
                    break
        self.df = self.df.drop_duplicates(subset=["nombre_cientifico"])
        self.df = self.df[self.df["nombre_cientifico"].str.strip() != ""]
        print(f"✅ {len(self.df)} especies en {self.cuadriculas}")
        return self.df

    def calcular_estadisticas(self, alertas: list | None = None) -> dict:
        if self.df is None or len(self.df) == 0:
            return {}
        grupos = self.df.groupby("grupo")["nombre_cientifico"].count().to_dict() if "grupo" in self.df.columns else {}
        # Muestra de especies REALES por grupo (no solo el recuento), para que
        # la redacción con phi-4 use ejemplos que existen de verdad en la
        # cuadrícula analizada, en vez de inventarlos y a veces colocarlos en
        # un grupo taxonómico equivocado (p.ej. un ave listada como mamífero).
        # Las especies protegidas se priorizan para que entren siempre en la
        # muestra, sea cual sea su posición en el listado original.
        especies_protegidas = {a["especie"] for a in (alertas or [])}
        muestra_por_grupo = {}
        if "grupo" in self.df.columns:
            for grupo, sub in self.df.groupby("grupo"):
                nombres = sub["nombre_cientifico"].tolist()
                protegidas = [n for n in nombres if n in especies_protegidas]
                resto = [n for n in nombres if n not in especies_protegidas]
                muestra_por_grupo[grupo] = (protegidas + resto)[:8]
        self.estadisticas = {
            "n_especies": len(self.df),
            "cuadriculas": self.cuadriculas,
            "por_grupo": grupos,
            "muestra_por_grupo": muestra_por_grupo,
        }
        return self.estadisticas

    def generar_redaccion(self, proyecto: str, alertas: list) -> str:
        stats = self.estadisticas
        cuads = ", ".join(self.cuadriculas)
        alertas_por_especie = {a["especie"]: a for a in alertas}
        lineas_grupo = []
        for g, n in stats.get("por_grupo", {}).items():
            muestra = stats.get("muestra_por_grupo", {}).get(g, [])
            ejemplos = []
            for esp in muestra:
                comun = alertas_por_especie.get(esp, {}).get("nombre_comun", "")
                ejemplos.append(f"{esp} ({comun})" if comun else f"{esp} (sin nombre común registrado — usa solo el nombre científico)")
            lineas_grupo.append(f"- {g}: {n} especies. Ejemplos reales presentes en esta cuadrícula: {'; '.join(ejemplos) or 'sin datos'}")
        grupos_txt = "\n".join(lineas_grupo)
        alertas_txt = "\n".join([
            f"- {a['especie']} ({a.get('nombre_comun','')}): "
            f"LESRPE={a.get('cat_lesrpe','—')} | {self.ccaa}={a.get('cat_auto_nombre','—')}"
            for a in alertas[:20]
        ]) or "No se detectaron especies con estatus de protección."
        prompt = f"""Eres un consultor ambiental senior con 15 años de experiencia en EsIA en España.
Redacta el inventario faunístico para el EsIA del proyecto "{proyecto}".

Cuadrículas IEET: {cuads}
Total especies: {stats.get('n_especies', 0)}

Distribución y ejemplos reales por grupo taxonómico (usa EXCLUSIVAMENTE estas
especies como ejemplos en cada grupo; no inventes especies nuevas ni cambies
ninguna de grupo taxonómico — respeta el grupo indicado para cada una. Si una
especie aparece marcada como "sin nombre común registrado", nómbrala solo por
su nombre científico en cursiva — NUNCA inventes ni supongas un nombre común
para ella):
{grupos_txt}

Especies protegidas (LESRPE/CEEA + {self.ccaa}):
{alertas_txt}

Estructura: Metodología (menciona IEET/BDIEET del MITECO), Resultados por grupos,
Especies de interés (cita RD 139/2011 y normativa autonómica cuando corresponda),
Valoración de la sensibilidad faunística.
Redacta SIEMPRE en prosa narrativa, con frases completas — NUNCA uses listas
con viñetas ni guiones para enumerar especies, intégralas dentro de las
frases como lo haría un consultor ambiental en un EsIA real.
Extensión: 400-500 palabras. Tono técnico de EsIA."""
        try:
            resp = client.chat.completions.create(
                model=LLM_MODEL,
                messages=[{"role": "user", "content": prompt}],
                temperature=0.15, max_tokens=1300,
            )
            return resp.choices[0].message.content
        except Exception as e:
            return f"[Error LM Studio: {e}]"
