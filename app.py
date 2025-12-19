import streamlit as st
import re, json, io, unicodedata
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
# Normalización de texto
# =========================
def normalize_text(text: str) -> str:
    if not text:
        return ""

    # Unificar guiones típicos de PDF
    text = text.replace("\u2013", "-").replace("\u2014", "-").replace("\u2212", "-")

    # Quitar guionado por salto de línea: "inves-\ntigación" -> "investigación"
    text = re.sub(r"(\w)-\n(\w)", r"\1\2", text)

    # Normalizar saltos de línea
    text = text.replace("\r\n", "\n").replace("\r", "\n")

    # Espacios
    text = re.sub(r"[ \t]+", " ", text)

    # Compactar líneas vacías
    text = re.sub(r"\n{3,}", "\n\n", text)

    return text


def strip_accents_lower(s: str) -> str:
    """
    Para PDFs con 'Psicologìa', 'Año', etc.
    """
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return s.lower()


# =========================
# Regex flags desde criteria.json
# =========================
def flags_from_meta(criteria_dict) -> int:
    meta = criteria_dict.get("meta", {})
    f = meta.get("regex_flags_default", "is")  # default: i + s
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
        st.warning(f"Regex inválida: {e} | patrón: {pattern[:120]}...")
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
# Detectores robustos (NO dependen de saltos de línea)
# =========================

# Bloque de "finalización" (ya sin tildes)
END_RX = re.compile(
    r"ano de (finalizacion|obtencion|graduacion)\s*:\s*(\d{2}/)?(19|20)\d{2}",
    re.IGNORECASE,
)

def count_posgrado_finalizado_block(norm_text: str, kind: str) -> int:
    """
    kind: 'doctorado'|'maestria'|'especializacion'|'profesorado'
    Reglas:
    - Debe existir el término (doctorado/maestria/...) y, dentro de 0..900 chars,
      un marcador de finalización (ano de finalizacion...) o 'situacion del nivel: completo'
    - Y NO debe haber 'actualidad' entre el término y el marcador.
    - Dedup por bloque.
    """
    if not norm_text:
        return 0

    if kind == "doctorado":
        title_rx = re.compile(r"\bdoctorad[oa]\b", re.IGNORECASE)
    elif kind == "maestria":
        title_rx = re.compile(r"\bmaestria\b|\bmagister\b", re.IGNORECASE)
    elif kind == "especializacion":
        title_rx = re.compile(r"\bespecializacion\b|\bespecialista\b", re.IGNORECASE)
    else:
        title_rx = re.compile(r"\bprofesorado\b|\bprofesor\b", re.IGNORECASE)

    completo_rx = re.compile(r"situacion del nivel\s*:\s*completo", re.IGNORECASE)

    vistos = set()
    count = 0

    for mt in title_rx.finditer(norm_text):
        window = norm_text[mt.start(): min(len(norm_text), mt.start() + 900)]

        m_end = END_RX.search(window)
        m_comp = completo_rx.search(window)

        if not (m_end or m_comp):
            continue

        marker_idx = None
        if m_end and m_comp:
            marker_idx = min(m_end.start(), m_comp.start())
        elif m_end:
            marker_idx = m_end.start()
        else:
            marker_idx = m_comp.start()

        between = window[:marker_idx]
        if "actualidad" in between:
            continue

        key = window[:200].strip()
        if key in vistos:
            continue
        vistos.add(key)
        count += 1

    return min(count, 3)


def count_titulo_grado_finalizado_block(norm_text: str) -> int:
    """
    Detecta títulos de grado por BLOQUE:
      (titulo de grado) .... ano de finalizacion: ...
    sin depender de '\n'.
    Rechaza si 'actualidad' aparece dentro del bloque antes del ancla.

    Dedup por (titulo_normalizado).
    """
    if not norm_text:
        return 0

    # Títulos de grado típicos + tecnicaturas universitarias
    grado_title = (
        r"(?:"
        r"licenciad[oa]\s+en\s+[a-z0-9 .,'\"-]{2,90}|"
        r"contador[ae]?\s+public[oa]|"
        r"contadora\s+publica|contador\s+publico|"
        r"ingenier[oa]\s+en\s+[a-z0-9 .,'\"-]{2,90}|ingenier[oa]|"
        r"abogad[oa]|"
        r"medic[oa]|"
        r"bioquimic[oa]|"
        r"farmaceutic[oa]|"
        r"odontolog[oa]|"
        r"psicologi[aia]|psicolog[oa]|"
        r"arquitect[oa]|"
        r"veterinar[ioia]|"
        r"nutricionist[ae]|"
        r"enfermer[oa]|"
        r"tecnic[oa]\s+universitari[oa]\s+en\s+[a-z0-9 .,'\"-]{2,90}|"
        r"tecnic[oa]\s+universitari[oa]|"
        r"tecnico\s+universitario"
        r")"
    )

    block_rx = re.compile(
        rf"(?P<title>{grado_title})"
        rf"[\s\S]{{0,260}}?"
        rf"(?P<end>ano de (?:finalizacion|obtencion|graduacion)\s*:\s*(?:\d{{2}}/)?(?:19|20)\d{{2}})",
        re.IGNORECASE,
    )

    vistos = set()
    count = 0

    for m in block_rx.finditer(norm_text):
        block = norm_text[m.start(): m.end()]

        # si aparece 'actualidad' dentro del bloque, NO cuenta
        if "actualidad" in block:
            continue

        title = re.sub(r"\s+", " ", (m.group("title") or "").strip().lower())
        if not title:
            continue

        # Evitar que cuente posgrados por error (si el título trae "maestria/doctorado")
        if "maestria" in title or "doctorado" in title or "especializacion" in title:
            continue

        if title in vistos:
            continue
        vistos.add(title)

        count += 1

    return min(count, 6)


# =========================
# UI
# =========================
uploaded = st.file_uploader("Cargar CV (.docx o .pdf)", type=["docx", "pdf"])

if uploaded:
    ext = uploaded.name.split(".")[-1].lower()
    try:
        raw_text = extract_text_docx(uploaded) if ext == "docx" else extract_text_pdf(uploaded)
        raw_text = normalize_text(raw_text)
        norm_text = strip_accents_lower(raw_text)
    except Exception as e:
        st.error(str(e))
        st.stop()

    st.success(f"Archivo cargado: {uploaded.name}")

    with st.expander("Ver texto extraído (debug)"):
        st.text_area("Texto", raw_text, height=240)

    results = {}
    total = 0.0

    for section, cfg in criteria.get("sections", {}).items():
        st.markdown(f"### {section}")
        rows = []
        subtotal_raw = 0.0

        items = cfg.get("items", {})
        for item, icfg in items.items():
            pattern = icfg.get("pattern", "")
            unit = float(icfg.get("unit_points", 0) or 0)
            item_cap = float(icfg.get("max_points", 0) or 0)

            sec_is_form = section.strip().lower() == "formación académica y complementaria"

            if sec_is_form and item.strip().lower().startswith("doctorado"):
                c = count_posgrado_finalizado_block(norm_text, "doctorado")
            elif sec_is_form and item.strip().lower().startswith("maestr"):
                c = count_posgrado_finalizado_block(norm_text, "maestria")
            elif sec_is_form and item.strip().lower().startswith("especial"):
                c = count_posgrado_finalizado_block(norm_text, "especializacion")
            elif sec_is_form and item.strip().lower().startswith("profesorado"):
                c = count_posgrado_finalizado_block(norm_text, "profesorado")
            elif sec_is_form and item.strip().lower().startswith("título de grado"):
                c = count_titulo_grado_finalizado_block(norm_text)
            else:
                # para el resto, mantenemos el matching original contra raw_text
                c = match_count(pattern, raw_text)

            pts_raw = c * unit
            pts = clip(pts_raw, item_cap)

            rows.append({
                "Ítem": item,
                "Ocurrencias": c,
                "Puntaje (tope ítem)": pts,
                "Tope ítem": item_cap
            })
            subtotal_raw += pts

        df = pd.DataFrame(rows)
        section_cap = float(cfg.get("max_points", 0) or 0)
        subtotal = clip(subtotal_raw, section_cap)

        st.dataframe(df, use_container_width=True)
        st.info(f"Subtotal {section}: {subtotal:.1f} / máx {section_cap:.0f}")

        results[section] = {"df": df, "subtotal": subtotal}
        total += subtotal

    clave_cat, desc_cat = obtener_categoria(total, criteria)
    categoria_label = "Sin categoría" if clave_cat == "Sin categoría" else f"Categoría {clave_cat}"

    st.markdown("---")
    st.subheader("Puntaje total y categoría")
    st.metric("Total acumulado", f"{total:.1f}")
    st.metric("Categoría alcanzada", categoria_label)
    if desc_cat:
        st.info(f"Descripción de la categoría: {desc_cat}")

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
