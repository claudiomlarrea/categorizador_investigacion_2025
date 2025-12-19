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
st.title("Universidad Católica de Cuyo — Valorador de CV Docente")
st.caption("Incluye exportación a Excel y Word + categoría automática según puntaje total.")

@st.cache_data(show_spinner=False)
def load_json(path):
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except json.JSONDecodeError as e:
        st.error(f"criteria.json inválido: {e.msg} (línea {e.lineno}, columna {e.colno}).")
        st.info("Tip: revisá comillas, comas finales y backslashes en regex (en JSON deben ser \\\\).")
        st.stop()
    except FileNotFoundError:
        st.error("No se encontró criteria.json en el repositorio (debe estar en la misma carpeta que app.py).")
        st.stop()
    except Exception as e:
        st.error(f"Error leyendo criteria.json: {e}")
        st.stop()

criteria = load_json("criteria.json")


# =========================
# Extracción de texto
# =========================
def extract_text_docx(file):
    doc = DocxDocument(file)
    text = "\n".join(p.text for p in doc.paragraphs)
    for t in doc.tables:
        for row in t.rows:
            text += "\n" + " | ".join(c.text for c in row.cells)
    return text

def extract_text_pdf(file):
    if not HAVE_PDF:
        raise RuntimeError("Falta pdfplumber. Agregalo en requirements.txt: pdfplumber")
    chunks = []
    with pdfplumber.open(file) as pdf:
        for p in pdf.pages:
            chunks.append(p.extract_text() or "")
    return "\n".join(chunks)


# =========================
# Helpers
# =========================
def match_count(pattern, text):
    return len(re.findall(pattern, text, re.IGNORECASE)) if pattern else 0

def clip(v, cap):
    return min(v, cap) if cap else v

def normalize_spaces(s: str) -> str:
    s = s.replace("\u00A0", " ")
    s = re.sub(r"[ \t]+", " ", s)
    s = re.sub(r"\n{3,}", "\n\n", s)
    return s.strip()

def obtener_categoria(total, criteria_dict):
    categorias = criteria_dict.get("categorias", {})
    mejor_clave = "Sin categoría"
    mejor_desc = ""
    mejor_min = None

    for clave, info in categorias.items():
        min_pts = info.get("min_points", 0)
        if total >= min_pts and (mejor_min is None or min_pts > mejor_min):
            mejor_min = min_pts
            mejor_clave = clave
            mejor_desc = info.get("descripcion", "")

    return mejor_clave, mejor_desc


# ==========================================================
# 1) Recorte robusto de la sección "FORMACION ACADÉMICA"
# ==========================================================
FORMACION_HEADERS = [
    r"FORMACI[ÓO]N ACAD[ÉE]MICA",
    r"FORMACION ACADEMICA",
    r"FORMACI[ÓO]N\s+ACAD[ÉE]MICA",
]

# Cortes probables (siguiente sección). Preferimos MAYÚSCULAS típicas.
NEXT_SECTION_MARKERS = [
    r"\n\s*ANTECEDENTES\b",
    r"\n\s*PRODUCCI[ÓO]N\b",
    r"\n\s*PUBLICACIONES\b",
    r"\n\s*ACTIVIDADES\b",
    r"\n\s*EXPERIENCIA\b",
    r"\n\s*CARGOS\b",
    r"\n\s*FORMACI[ÓO]N COMPLEMENTARIA\b",
    r"\n\s*CURSOS\b",
    r"\n\s*IDIOMAS\b",
]

def extract_formacion_academica_block(full_text: str) -> str:
    txt = normalize_spaces(full_text)
    # buscar el primer header de formación
    start_idx = None
    for h in FORMACION_HEADERS:
        m = re.search(h, txt, flags=re.IGNORECASE)
        if m:
            start_idx = m.end()
            break
    if start_idx is None:
        return ""

    tail = txt[start_idx:]
    # cortar en el primer marcador de siguiente sección
    end_idx = len(tail)
    for mk in NEXT_SECTION_MARKERS:
        m2 = re.search(mk, tail, flags=re.IGNORECASE)
        if m2:
            end_idx = min(end_idx, m2.start())
    block = tail[:end_idx].strip()
    return block


# ==========================================================
# 2) Parseo de entradas y reglas "finalizado vs en curso"
# ==========================================================
RE_IN_PROGRESS = re.compile(
    r"\b(Actualidad|En\s+curso|Cursando|En\s+desarrollo|Vigente|Actualmente)\b",
    re.IGNORECASE
)

RE_ENDS_WITH_ACTUALIDAD = re.compile(
    r"(\d{2}/\d{4}|\d{4})\s*-\s*Actualidad\b",
    re.IGNORECASE
)

RE_FINISH_YEAR = re.compile(
    r"A[nñ]o\s+de\s+(finalizaci[oó]n|obtenci[oó]n|graduaci[oó]n)\s*:\s*([0-3]?\d/\d{4}|\d{4})",
    re.IGNORECASE
)

RE_SITUACION_COMPLETO = re.compile(
    r"Situaci[oó]n\s+del\s+nivel\s*:\s*Completo",
    re.IGNORECASE
)

# “Título”/denominación en CVar: suele ir primera línea del bloque
def split_entries(block: str) -> list[str]:
    if not block:
        return []
    # separar por doble salto (en CVAR suele haber “bloques”)
    parts = re.split(r"\n\s*\n", block)
    entries = []
    for p in parts:
        p = p.strip()
        if len(p) < 3:
            continue
        entries.append(p)
    return entries

def entry_is_completed(entry: str) -> bool:
    # Si hay cualquier indicador claro de “en curso”, NO puntúa
    if RE_IN_PROGRESS.search(entry) or RE_ENDS_WITH_ACTUALIDAD.search(entry):
        return False
    # Para puntuar: debe tener año/mes-año de finalización u “Situación: Completo”
    if RE_FINISH_YEAR.search(entry) or RE_SITUACION_COMPLETO.search(entry):
        return True
    return False

def get_finish_token(entry: str) -> str:
    m = RE_FINISH_YEAR.search(entry)
    if m:
        return m.group(2).strip()
    if RE_SITUACION_COMPLETO.search(entry):
        return "COMPLETO"
    return ""

def get_first_line_title(entry: str) -> str:
    # primera línea “significativa” (sin null)
    lines = [l.strip() for l in entry.split("\n") if l.strip()]
    for l in lines:
        if l.lower() == "null":
            continue
        return l
    return (lines[0] if lines else "").strip()

def norm_key(s: str) -> str:
    s = s.lower().strip()
    s = re.sub(r"\s+", " ", s)
    s = re.sub(r"[\"'`´]", "", s)
    return s

def classify_entry(entry: str) -> str:
    # Posgrados primero
    if re.search(r"\bDoctorado\b|\bDoctor\s+en\b|\bDoctor\s+de\s+la\s+Universidad\b", entry, re.IGNORECASE):
        return "doctorado"
    if re.search(r"\bMaestr[ií]a\b|\bMag[ií]ster\b", entry, re.IGNORECASE):
        return "maestria"
    if re.search(r"\bEspecializaci[oó]n\b|\bEspecialista\b", entry, re.IGNORECASE):
        return "especializacion"
    if re.search(r"\bPosdoctorado\b|\bPostdoctorado\b", entry, re.IGNORECASE):
        return "posdoc"

    # Profesorado universitario
    if re.search(r"\bProfesorado\b|\bProfesor\s+en\b", entry, re.IGNORECASE):
        return "profesorado"

    # Grado: regla práctica (NO posgrado) + tiene “Año de finalización” + no “Actualidad”
    # Esto captura: Contadora Pública, Abogado, Bioquímico, Ingeniero, Médico, Licenciado/a, etc.
    if RE_FINISH_YEAR.search(entry) and not (RE_IN_PROGRESS.search(entry) or RE_ENDS_WITH_ACTUALIDAD.search(entry)):
        return "grado"

    return "otro"


def counts_from_formacion(block: str) -> dict:
    """
    Devuelve conteos robustos SOLO de formación académica.
    - Posgrados puntúan SOLO si entry_is_completed(entry)
    - Grado puntúa si tiene año de finalización y no en curso
    - Dedup por (tipo, titulo_normalizado, fin_token)
    """
    entries = split_entries(block)
    seen = set()

    counts = {
        "doctorado": 0,
        "maestria": 0,
        "especializacion": 0,
        "grado": 0,
        "profesorado": 0,
        "posdoc": 0,
    }

    for e in entries:
        tipo = classify_entry(e)
        if tipo not in counts:
            continue

        titulo = get_first_line_title(e)
        fin = get_finish_token(e)
        key = (tipo, norm_key(titulo), norm_key(fin))

        if key in seen:
            continue

        # Reglas de completitud
        if tipo in ("doctorado", "maestria", "especializacion"):
            if not entry_is_completed(e):
                continue

        if tipo == "grado":
            # grado: exige año de finalización (ya lo exige classify_entry)
            if not entry_is_completed(e) and not RE_FINISH_YEAR.search(e):
                continue

        # profesorado: si tiene año de finalización o “Completo”, puntúa.
        if tipo == "profesorado":
            if not (RE_FINISH_YEAR.search(e) or RE_SITUACION_COMPLETO.search(e)):
                continue
            if RE_IN_PROGRESS.search(e) or RE_ENDS_WITH_ACTUALIDAD.search(e):
                continue

        # posdoc: puede ser en curso o finalizado (según tu criterio actual)
        # lo dejamos como estaba: puntúa si aparece (sin exigir finalización)
        # Si querés exigir finalización también, avisame y lo ajusto.

        seen.add(key)
        counts[tipo] += 1

    return counts


# =========================
# UI
# =========================
uploaded = st.file_uploader("Cargar CV (.docx o .pdf)", type=["docx", "pdf"])

if uploaded:
    ext = uploaded.name.split(".")[-1].lower()
    try:
        raw_text = extract_text_docx(uploaded) if ext == "docx" else extract_text_pdf(uploaded)
    except Exception as e:
        st.error(str(e))
        st.stop()

    raw_text = normalize_spaces(raw_text)

    st.success(f"Archivo cargado: {uploaded.name}")

    # Debug general
    with st.expander("Ver texto extraído (debug)"):
        st.text_area("Texto", raw_text, height=240)

    # Debug de formación
    form_block = extract_formacion_academica_block(raw_text)
    with st.expander("Ver sección de Formación académica (debug)"):
        st.text_area("FORMACIÓN ACADÉMICA (recorte)", form_block if form_block else "[No se encontró la sección]", height=240)

    form_counts = counts_from_formacion(form_block)

    results = {}
    total = 0.0

    # =========================
    # Cálculo de puntajes por sección
    # =========================
    for section, cfg in criteria["sections"].items():
        st.markdown(f"### {section}")
        rows = []
        subtotal_raw = 0.0

        for item, icfg in cfg.get("items", {}).items():
            pattern = icfg.get("pattern", "")

            # --- Overrides robustos SOLO para la sección de Formación ---
            c = None
            if section.lower().startswith("formación académica"):
                item_l = item.lower()

                # doctorado finalizado
                if "doctorado" in item_l:
                    c = form_counts.get("doctorado", 0)

                # maestría finalizada
                elif "maestr" in item_l or "magíster" in item_l or "magister" in item_l:
                    c = form_counts.get("maestria", 0)

                # especialización finalizada
                elif "especializ" in item_l or "especialista" in item_l:
                    c = form_counts.get("especializacion", 0)

                # título de grado finalizado
                elif "título de grado" in item_l or "titulo de grado" in item_l or "grado" == item_l.strip():
                    c = form_counts.get("grado", 0)

                # profesorado
                elif "profesorado" in item_l or "docencia universitaria" in item_l:
                    c = form_counts.get("profesorado", 0)

                # posdoc
                elif "posdoctorado" in item_l or "postdoctorado" in item_l:
                    c = form_counts.get("posdoc", 0)

            # si no aplicó override -> lógica original por regex global
            if c is None:
                c = match_count(pattern, raw_text)

            pts = clip(c * icfg.get("unit_points", 0), icfg.get("max_points", 0))
            rows.append({
                "Ítem": item,
                "Ocurrencias": c,
                "Puntaje (tope ítem)": pts,
                "Tope ítem": icfg.get("max_points", 0)
            })
            subtotal_raw += pts

        df = pd.DataFrame(rows)
        subtotal = clip(subtotal_raw, cfg.get("max_points", 0))
        st.dataframe(df, use_container_width=True)
        st.info(f"Subtotal {section}: {subtotal} / máx {cfg.get('max_points', 0)}")
        results[section] = {"df": df, "subtotal": subtotal}
        total += subtotal

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
        resumen.loc[len(resumen)] = ["TOTAL", resumen["Subtotal"].sum()]
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
