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


# =========================
# Cargar criteria.json
# =========================
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
# Normalización (CLAVE)
# =========================
def normalize_text(text: str) -> str:
    if not text:
        return ""
    # unificar guiones típicos de PDF
    text = text.replace("\u2013", "-").replace("\u2014", "-").replace("\u2212", "-")
    # quitar guionado por salto de línea: "inves-\ntigación" -> "investigación"
    text = re.sub(r"(\w)-\n(\w)", r"\1\2", text)
    # normalizar saltos de línea
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    # compactar espacios
    text = re.sub(r"[ \t]+", " ", text)
    # compactar líneas vacías
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text


# =========================
# Flags regex desde criteria.json
# =========================
def flags_from_meta(criteria_dict) -> int:
    meta = criteria_dict.get("meta", {})
    f = meta.get("regex_flags_default", "is")  # default i + s
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
    return min(v, cap_val)


# =========================
# Detector robusto: TITULACIÓN FINALIZADA por bloque
# (evita que "Actualidad" se cuele por ventana chica)
# =========================
INPROGRESS_RX = re.compile(r"(?is)\b(Actualidad|En\s+curso|Cursando|No\s+finalizad[oa]|Incompleto|Pendiente)\b")
FINISH_RX = re.compile(
    r"(?is)\b(Situaci[oó]n\s+del\s+nivel\s*:\s*Completo|A[nñ]o\s+de\s+(finalizaci[oó]n|obtenci[oó]n|graduaci[oó]n)\s*:\s*(?:\d{2}/)?(?:19\d{2}|20\d{2}))\b"
)

def count_completed_by_blocks(title_anchor_regex: str, text: str, max_block_chars: int = 1400) -> int:
    """
    Encuentra ocurrencias del ancla (p.ej. 'Doctorado', 'Maestría', 'Licenciado en', 'Técnic[ao] Universitari[ao]')
    y evalúa el bloque cercano:
    - NO cuenta si aparece "Actualidad/En curso/Cursando..."
    - SOLO cuenta si aparece evidencia de finalización (Año de finalización / Situación Completo, etc.)
    """
    rx_anchor = re.compile(title_anchor_regex, re.IGNORECASE)
    count = 0

    for m in rx_anchor.finditer(text):
        start = m.start()
        end = min(len(text), m.start() + max_block_chars)
        block = text[start:end]

        if INPROGRESS_RX.search(block):
            continue
        if not FINISH_RX.search(block):
            continue

        count += 1

    return count


# =========================
# Categorización según puntaje
# =========================
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
        st.text_area("Texto", raw_text, height=240)

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

            # --- OVERRIDE robusto SOLO para formación finalizada ---
            if section == "Formación académica y complementaria":
                if item == "Doctorado (finalizado)":
                    c = count_completed_by_blocks(r"\b(Doctorado|Doctor\s+en|Doctor\s+de\s+la\s+Universidad)\b", raw_text)
                elif item == "Maestría (finalizada)":
                    c = count_completed_by_blocks(r"\b(Maestr[ií]a|Mag[ií]ster)\b", raw_text)
                elif item == "Especialización (finalizada)":
                    c = count_completed_by_blocks(r"\b(Especializaci[oó]n|Especialista)\b", raw_text)
                elif item == "Profesorado/Docencia universitaria (finalizado)":
                    c = count_completed_by_blocks(r"\b(Profesorado|Docente\s+Universitario|Profesor\s+Universitari[oa]|Profesor\s+en)\b", raw_text)
                elif item == "Título de grado (finalizado)":
                    # Incluye: Licenciado en..., Licenciatura en..., y títulos técnicos universitarios
                    c1 = count_completed_by_blocks(r"\b(Licenciad[oa]\s+en|Licenciatura\s+en)\b", raw_text)
                    c2 = count_completed_by_blocks(r"\b(T[eé]cnic[ao]\s+Universitari[ao]|Tecnicatura)\b", raw_text)
                    # Mantiene también profesiones típicas si están con año de finalización
                    c3 = count_completed_by_blocks(r"\b(Abogad[oa]|M[eé]dic[oa]|Veterinari[oa]|Bioqu[ií]mic[oa]|Contador[oa]?|Ingenier[oa]|Arquitect[oa]|Bromatolog[íi]a|Bromat[oó]log[oa])\b", raw_text)
                    c = c1 + c2 + c3
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
            sheet = sec[:31]
            data["df"].to_excel(writer, sheet_name=sheet, index=False)

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
