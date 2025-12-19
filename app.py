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

    # Compactar líneas vacías (cuando existen)
    text = re.sub(r"\n{3,}", "\n\n", text)

    return text


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
# Helpers robustos "finalizado"
# =========================
def _is_garbage_line(line: str) -> bool:
    if not line:
        return True
    l = line.strip()
    if not l:
        return True
    # basura típica del CVar
    if re.match(r"^(null|universidad|facultad|sede|instituto)\b", l, re.IGNORECASE):
        return True
    # líneas claramente institucionales
    if re.search(r"\bUNIVERSIDAD\b|\bFACULTAD\b|\bSEDE\b|\bINSTITUTO\b", l, re.IGNORECASE):
        return True
    return False


def _find_title_line_before(pos: int, text: str, max_lines_back=25):
    """
    Busca una línea candidata a "título" inmediatamente antes de 'pos'
    (usado para anclar el bloque real del título, sin arrastrar Actualidad de arriba).
    """
    pre = text[:pos]
    lines = pre.split("\n")
    tail = lines[-max_lines_back:] if len(lines) >= max_lines_back else lines
    tail = [ln.strip() for ln in tail if ln.strip()]

    # De atrás hacia adelante: tomar primera línea "no basura" que parezca título
    for ln in reversed(tail):
        if _is_garbage_line(ln):
            continue
        # evitar agarrar la línea "Año de finalización..."
        if re.search(r"A[nñ]o de (finalizaci[oó]n|obtenci[oó]n|graduaci[oó]n)\s*:", ln, re.IGNORECASE):
            continue
        # evitar rangos de fechas solos
        if re.match(r"^\d{2}/\d{4}\s*-\s*(Actualidad|\d{2}/\d{4}|\d{4})$", ln, re.IGNORECASE):
            continue
        return ln

    return ""


def count_titulo_grado_finalizado(text: str) -> int:
    """
    Cuenta títulos de grado finalizados usando ancla 'Año de finalización',
    y verificando que 'Actualidad' NO esté entre la línea del título y el ancla.
    Dedup por título para evitar inflación por repeticiones del PDF.
    """
    if not text:
        return 0

    end_rx = re.compile(
        r"A[nñ]o de (finalizaci[oó]n|obtenci[oó]n|graduaci[oó]n)\s*:\s*(\d{2}/)?(19|20)\d{2}",
        re.IGNORECASE,
    )

    # keywords típicas de grado (incluye tecnicaturas universitarias)
    grado_kw = re.compile(
        r"\b("
        r"licenciad[oa]|"
        r"contador[ae]?\s+public[oa]|"
        r"ingenier[oa]|"
        r"abogad[oa]|"
        r"medic[oa]|"
        r"bioquimic[oa]|"
        r"farmaceutic[oa]|"
        r"odontolog[oa]|"
        r"psicolog[iíì]a|psicolog[oa]|"
        r"arquitect[oa]|"
        r"veterinar[ioia]|"
        r"nutricionist[ae]|"
        r"enfermer[oa]|"
        r"tecnic[oa]\s+universitari[oa]|tecnico\s+universitario"
        r")\b",
        re.IGNORECASE,
    )

    vistos = set()
    found = 0

    for m in end_rx.finditer(text):
        title_line = _find_title_line_before(m.start(), text)
        if not title_line:
            continue

        # Confirmar que el título sea "de grado"
        if not grado_kw.search(title_line):
            continue

        # Buscar la posición real de esa línea para recortar bloque local
        # (rfind cerca del final para evitar agarrar una ocurrencia vieja)
        search_from = max(0, m.start() - 2000)
        segment = text[search_from:m.start()]
        idx_local = segment.lower().rfind(title_line.lower())
        if idx_local == -1:
            title_pos = m.start()
        else:
            title_pos = search_from + idx_local

        # CLAVE: mirar SOLO entre título y el ancla de finalización
        between = text[title_pos:m.end()]
        if re.search(r"\bActualidad\b", between, re.IGNORECASE):
            continue

        key = re.sub(r"\s+", " ", title_line.strip().lower())
        if key in vistos:
            continue
        vistos.add(key)

        found += 1

    return min(found, 6)


def count_posgrado_finalizado(text: str, titulo_rx: str) -> int:
    """
    Cuenta posgrados FINALIZADOS (Doctorado/Maestría/Especialización/Profesorado)
    solo si:
      - aparece el título (titulo_rx)
      - y en los ~600 chars siguientes existe:
           'Situación del nivel: Completo'  o  'Año de finalización: ...'
      - y NO aparece 'Actualidad' entre título y marcador de finalización.
    Dedup por título de la línea.
    """
    if not text:
        return 0

    title_rx = re.compile(titulo_rx, re.IGNORECASE)
    completo_rx = re.compile(r"Situaci[oó]n del nivel\s*:\s*Completo", re.IGNORECASE)
    end_rx = re.compile(
        r"A[nñ]o de (finalizaci[oó]n|obtenci[oó]n|graduaci[oó]n)\s*:\s*(\d{2}/)?(19|20)\d{2}",
        re.IGNORECASE,
    )

    vistos = set()
    count = 0

    for mt in title_rx.finditer(text):
        # hallar línea título real
        # tomamos desde el inicio de la línea hasta el fin de la línea
        line_start = text.rfind("\n", 0, mt.start())
        line_start = 0 if line_start == -1 else line_start + 1
        line_end = text.find("\n", mt.start())
        line_end = len(text) if line_end == -1 else line_end
        title_line = text[line_start:line_end].strip()

        if _is_garbage_line(title_line):
            continue

        # mirar ventana hacia adelante para encontrar marcador de finalización
        fwd_end = min(len(text), mt.start() + 900)
        window = text[mt.start():fwd_end]

        m_comp = completo_rx.search(window)
        m_end = end_rx.search(window)

        if not (m_comp or m_end):
            continue

        # recortar ENTRE título y el primer marcador
        marker_pos = None
        if m_comp and m_end:
            marker_pos = mt.start() + min(m_comp.start(), m_end.start())
        elif m_comp:
            marker_pos = mt.start() + m_comp.start()
        else:
            marker_pos = mt.start() + m_end.start()

        between = text[mt.start():marker_pos]
        if re.search(r"\bActualidad\b", between, re.IGNORECASE):
            continue

        key = re.sub(r"\s+", " ", title_line.strip().lower())
        if key in vistos:
            continue
        vistos.add(key)

        count += 1

    return min(count, 3)


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

            # -------------------------
            # OVERRIDE robusto Formación
            # -------------------------
            sec_is_form = section.strip().lower() == "formación académica y complementaria"

            if sec_is_form and item.strip().lower().startswith("doctorado"):
                c = count_posgrado_finalizado(raw_text, r"\bDoctorad[oa]\b")
            elif sec_is_form and item.strip().lower().startswith("maestr"):
                c = count_posgrado_finalizado(raw_text, r"\bMaestr[iíì]a\b|\bMag[iíì]ster\b")
            elif sec_is_form and item.strip().lower().startswith("especial"):
                c = count_posgrado_finalizado(raw_text, r"\bEspecializaci[oó]n\b|\bEspecialista\b")
            elif sec_is_form and item.strip().lower().startswith("profesorado"):
                c = count_posgrado_finalizado(raw_text, r"\bProfesorado\b|\bProfesor\s+en\b|\bProfesor\b")
            elif sec_is_form and item.strip().lower().startswith("título de grado"):
                c = count_titulo_grado_finalizado(raw_text)
            else:
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
