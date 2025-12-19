# app.py — Valorador de CV Docente/Investigador (UCCuyo)
# FIX CLAVE:
# - Para títulos estructurales (Doctorado/Maestría/Especialización/Grado/Profesorado):
#   toma BLOQUES recortados (hasta el próximo título) para que "Actualidad" de otro ítem no contamine.
# - No puntúa posgrados "en curso" (Actualidad/En curso/Cursando/etc.).
# - Tolera "null" intermedio típico del CVar.
# - Ítems no estructurales: conteo por regex del criteria.json + topes.

import streamlit as st
import re, json, io
import pandas as pd
from docx import Document as DocxDocument
from docx.enum.text import WD_ALIGN_PARAGRAPH

# PDF opcional
try:
    import pdfplumber
    HAVE_PDF = True
except Exception:
    HAVE_PDF = False

st.set_page_config(page_title="Valorador de CV - UCCuyo (DOCX/PDF)", layout="wide")
st.title("Universidad Católica de Cuyo — Valorador de CV Docente/Investigador")
st.caption("Exporta Excel y Word + categoría automática según puntaje total (criteria.json).")

# =========================
# Cargar criteria.json
# =========================
@st.cache_data(show_spinner=False)
def load_json(path: str):
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except json.JSONDecodeError as e:
        st.error(f"criteria.json inválido: {e.msg} (línea {e.lineno}, columna {e.colno}).")
        st.info("Tip: revisá comillas, comas finales y backslashes en regex (en JSON deben ser \\\\).")
        st.stop()
    except FileNotFoundError:
        st.error("No se encontró criteria.json en el repositorio (debe estar junto a app.py).")
        st.stop()
    except Exception as e:
        st.error(f"Error leyendo criteria.json: {e}")
        st.stop()

criteria = load_json("criteria.json")

# =========================
# Extracción de texto
# =========================
def extract_text_docx(file) -> str:
    doc = DocxDocument(file)
    text = "\n".join(p.text for p in doc.paragraphs)
    for t in doc.tables:
        for row in t.rows:
            text += "\n" + " | ".join(c.text for c in row.cells)
    return text

def extract_text_pdf(file) -> str:
    if not HAVE_PDF:
        raise RuntimeError("Falta pdfplumber. Agregalo en requirements.txt: pdfplumber")
    chunks = []
    with pdfplumber.open(file) as pdf:
        for p in pdf.pages:
            chunks.append(p.extract_text() or "")
    return "\n".join(chunks)

# =========================
# Normalización
# =========================
def normalize_text(text: str) -> str:
    if not text:
        return ""
    text = text.replace("\u2013", "-").replace("\u2014", "-").replace("\u2212", "-")
    text = re.sub(r"(\w)-\n(\w)", r"\1\2", text)  # guionado por corte de línea
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text

# =========================
# Regex flags desde criteria.json
# =========================
def flags_from_meta(criteria_dict) -> int:
    meta = criteria_dict.get("meta", {})
    f = meta.get("regex_flags_default", "is")
    flags = 0
    if "i" in f:
        flags |= re.IGNORECASE
    if "s" in f:
        flags |= re.DOTALL
    if "m" in f:
        flags |= re.MULTILINE
    return flags

DEFAULT_FLAGS = flags_from_meta(criteria)

@st.cache_data(show_spinner=False)
def compile_pattern(pattern: str, flags: int):
    return re.compile(pattern, flags)

def match_count(pattern: str, text: str) -> int:
    if not pattern:
        return 0
    try:
        rx = compile_pattern(pattern, DEFAULT_FLAGS)
        return sum(1 for _ in rx.finditer(text))
    except re.error as e:
        st.warning(f"Regex inválida: {e} | patrón: {pattern[:160]}...")
        return 0

def clip(v, cap):
    if cap is None:
        return v
    try:
        cap_val = float(cap)
    except Exception:
        return v
    return min(float(v), cap_val)

# =========================
# Reglas de finalización (NO EN CURSO)
# =========================
INPROGRESS_RX = re.compile(
    r"(?i)\b(Actualidad|En\s+curso|Cursando|No\s+finalizad[oa]|Sin\s+finalizar|Incompleto|"
    r"Doctorand[oa]|Maestrand[oa]|Especializand[oa])\b"
)

# Evidencia de finalización:
FINISH_RX = re.compile(
    r"(?is)("
    r"Situaci[oó]n\s+del\s+nivel\s*:\s*Completo|"
    r"A[nñ]o\s+de\s+(finalizaci[oó]n|obtenci[oó]n|graduaci[oó]n)\s*:\s*(?:\d{2}/)?(?:19\d{2}|20\d{2})|"
    r"\b\d{2}/\d{4}\s*[-–]\s*\d{2}/\d{4}\b|"
    r")"
)

# =========================
# BLOQUES recortados por próximo título (FIX)
# =========================
# Un “ancla global” que marca el inicio de cualquier título/posgrado (para cortar bloques).
GLOBAL_ANCHOR_RX = re.compile(
    r"(?i)\b("
    r"Doctorado|Doctor\s+en|Doctor\s+de\s+la\s+Universidad|"
    r"Maestr[ií]a|Mag[ií]ster|"
    r"Especializaci[oó]n|Especialista|"
    r"Licenciad[oa]\s+en|Licenciatura\s+en|"
    r"T[eé]cnic[ao]\s+Universitari[ao]|Tecnicatura|"
    r"Profesorado|Profesor\s+Universitari[oa]|Docente\s+Universitario|Profesor\s+en"
    r")\b"
)

def get_local_block(text: str, start_idx: int, max_chars: int = 2000) -> str:
    """
    Devuelve el bloque desde start_idx hasta el próximo título (ancla global) o hasta max_chars.
    Esto evita que "Actualidad" de un posgrado posterior contamine un título finalizado previo.
    """
    end_limit = min(len(text), start_idx + max_chars)
    tail = text[start_idx:end_limit]

    # Buscar el próximo ancla global DESPUÉS del primer carácter del bloque
    m_next = GLOBAL_ANCHOR_RX.search(tail, pos=1)
    if m_next:
        tail = tail[:m_next.start()]

    # Limpiar ruido CVar
    tail = re.sub(r"\bnull\b", " ", tail, flags=re.IGNORECASE)
    tail = re.sub(r"[ \t]+", " ", tail)
    return tail.strip()

def has_completed_title(title_anchor_regex: str, text: str) -> int:
    """
    Devuelve 1 si existe AL MENOS UN título/posgrado finalizado válido (no en curso).
    (Nunca >1, porque el ítem tiene tope y semántica 1/0.)
    """
    rx_anchor = re.compile(title_anchor_regex, re.IGNORECASE)
    for m in rx_anchor.finditer(text):
        block = get_local_block(text, m.start(), max_chars=2500)

        # excluir si está en curso
        if INPROGRESS_RX.search(block):
            continue

        # exigir evidencia de finalización
        if FINISH_RX.search(block):
            return 1

    return 0

# =========================
# Categorización por puntaje (criteria.json)
# =========================
def obtener_categoria(total: float, criteria_dict):
    categorias = criteria_dict.get("categorias", {})
    mejor_clave = "Sin categoría"
    mejor_desc = ""
    mejor_min = None
    for clave, info in categorias.items():
        min_pts = float(info.get("min_points", 0) or 0)
        if total >= min_pts and (mejor_min is None or min_pts > mejor_min):
            mejor_min = min_pts
            mejor_clave = clave
            mejor_desc = info.get("descripcion", "")
    return mejor_clave, mejor_desc

# =========================
# UI
# =========================
uploaded = st.file_uploader("Cargar CV (.docx o .pdf)", type=["docx", "pdf"])

if uploaded:
    ext = uploaded.name.split(".")[-1].lower()
    try:
        raw_text = extract_text_docx(uploaded) if ext == "docx" else extract_text_pdf(uploaded)
        raw_text = normalize_text(raw_text)
    except Exception as e:
        st.error(str(e))
        st.stop()

    st.success(f"Archivo cargado: {uploaded.name}")

    with st.expander("Ver texto extraído (debug)"):
        st.text_area("Texto", raw_text, height=260)

    results = {}
    total = 0.0

    # =========================
    # Cálculo por sección
    # =========================
    for section, cfg in criteria.get("sections", {}).items():
        st.markdown(f"### {section}")
        rows = []
        subtotal_raw = 0.0

        items = cfg.get("items", {})
        for item, icfg in items.items():
            pattern = icfg.get("pattern", "")
            unit = float(icfg.get("unit_points", 0) or 0)
            item_cap = float(icfg.get("max_points", 0) or 0)

            # --- REGLAS ESPECIALES: títulos estructurales ---
            if section == "Formación académica y complementaria":
                if item == "Doctorado (finalizado)":
                    c = has_completed_title(r"\b(Doctorado|Doctor\s+en|Doctor\s+de\s+la\s+Universidad)\b", raw_text)

                elif item == "Maestría (finalizada)":
                    c = has_completed_title(r"\b(Maestr[ií]a|Mag[ií]ster)\b", raw_text)

                elif item == "Especialización (finalizada)":
                    c = has_completed_title(r"\b(Especializaci[oó]n|Especialista)\b", raw_text)

                elif item == "Título de grado (finalizado)":
                    # Incluye variantes típicas del CVar (como Vinader):
                    # "LICENCIADO EN ...", "Técnica Universitaria en ...", etc.
                    c = has_completed_title(
                        r"\b("
                        r"Licenciad[oa]\s+en|Licenciatura\s+en|"
                        r"T[eé]cnic[ao]\s+Universitari[ao]\s+en|Tecnicatura\s+en|"
                        r")\b",
                        raw_text
                    )

                elif item == "Profesorado/Docencia universitaria (finalizado)":
                    c = has_completed_title(r"\b(Profesorado|Profesor\s+Universitari[oa]|Docente\s+Universitario|Profesor\s+en)\b", raw_text)

                else:
                    c = match_count(pattern, raw_text)
            else:
                c = match_count(pattern, raw_text)

            pts_raw = c * unit
            pts = clip(pts_raw, item_cap)

            rows.append({
                "Ítem": item,
                "Ocurrencias": int(c),
                "Puntaje (tope ítem)": float(pts),
                "Tope ítem": float(item_cap)
            })
            subtotal_raw += float(pts)

        df = pd.DataFrame(rows)
        section_cap = float(cfg.get("max_points", 0) or 0)
        subtotal = clip(subtotal_raw, section_cap)

        st.dataframe(df, use_container_width=True)
        st.info(f"Subtotal {section}: {subtotal:.1f} / máx {section_cap:.0f}")

        results[section] = {"df": df, "subtotal": float(subtotal)}
        total += float(subtotal)

    # =========================
    # Categoría
    # =========================
    clave_cat, desc_cat = obtener_categoria(total, criteria)
    categoria_label = "Sin categoría" if clave_cat == "Sin categoría" else f"Categoría {clave_cat}"

    st.markdown("---")
    st.subheader("Puntaje total y categoría")
    st.metric("Total acumulado", f"{total:.1f}")
    st.metric("Categoría alcanzada", categoria_label)
    if desc_cat:
        st.info(f"Descripción de la categoría: {desc_cat}")

    # =========================
    # Exportaciones
    # =========================
    st.markdown("---")
    st.subheader("Exportar resultados")

    # Excel
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        for sec, data in results.items():
            data["df"].to_excel(writer, sheet_name=sec[:31], index=False)

        resumen = pd.DataFrame({
            "Sección": list(results.keys()),
            "Subtotal": [results[s]["subtotal"] for s in results]
        })
        resumen.loc[len(resumen)] = ["TOTAL", float(resumen["Subtotal"].sum())]
        resumen.loc[len(resumen)] = ["CATEGORÍA", categoria_label]
        resumen.to_excel(writer, sheet_name="RESUMEN", index=False)

    st.download_button(
        "Descargar Excel",
        data=out.getvalue(),
        file_name="valoracion_cv.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

    # Word
    def export_word(results_dict, total_pts, cat_label, cat_desc):
        doc = DocxDocument()
        p = doc.add_paragraph("Universidad Católica de Cuyo — Secretaría de Investigación")
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph("Informe de valoración de CV").alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph("")
        doc.add_paragraph(f"Puntaje total: {total_pts:.1f}")
        doc.add_paragraph(f"Categoría alcanzada: {cat_label}")
        if cat_desc:
            doc.add_paragraph(cat_desc)

        for sec, data in results_dict.items():
            doc.add_heading(sec, level=2)
            df_sec = data["df"]
            if df_sec.empty:
                doc.add_paragraph("Sin ítems detectados.")
            else:
                tbl = doc.add_table(rows=1, cols=len(df_sec.columns))
                hdr = tbl.rows[0].cells
                for i, c in enumerate(df_sec.columns):
                    hdr[i].text = str(c)
                for _, row in df_sec.iterrows():
                    cells = tbl.add_row().cells
                    for i, c in enumerate(df_sec.columns):
                        cells[i].text = str(row[c])
            doc.add_paragraph(f"Subtotal sección: {data['subtotal']:.1f}")

        bio = io.BytesIO()
        doc.save(bio)
        return bio.getvalue()

    st.download_button(
        "Descargar informe Word",
        data=export_word(results, total, categoria_label, desc_cat),
        file_name="informe_valoracion_cv.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True
    )

else:
    st.info("Subí un archivo para iniciar la valoración.")
