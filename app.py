import streamlit as st
import re, json, io
import pandas as pd
import unicodedata
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


def strip_accents(s: str) -> str:
    if not s:
        return ""
    # NFKD separa letras de diacríticos, luego se eliminan marcas
    return "".join(
        ch for ch in unicodedata.normalize("NFKD", s)
        if not unicodedata.combining(ch)
    )


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
        st.warning(f"Regex inválida: {e} | patrón: {pattern[:140]}...")
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
# Recorte de sección "Formación académica"
# =========================
def slice_section(text: str, start_patterns, end_patterns, max_len_after=25000) -> str:
    if not text:
        return ""

    start_idx = None
    for sp in start_patterns:
        m = re.search(sp, text, re.IGNORECASE)
        if m:
            start_idx = m.start()
            break

    if start_idx is None:
        return text

    tail = text[start_idx:start_idx + max_len_after]

    end_idx = None
    for ep in end_patterns:
        m2 = re.search(ep, tail, re.IGNORECASE)
        if m2:
            if m2.start() > 20:
                end_idx = m2.start()
                break

    return tail[:end_idx] if end_idx else tail


# =========================
# Detección robusta: finalizado vs en curso
# =========================
IN_PROGRESS_RX = re.compile(
    r"\bactualidad\b|\ben curso\b|\bcursando\b|\bno finalizad[oa]\b|\bsin finalizar\b",
    re.IGNORECASE
)

FINALIZATION_RX = re.compile(
    r"anio de (finalizacion|obtencion|graduacion)\s*:\s*(\d{2}/\d{4}|(19|20)\d{2})",
    re.IGNORECASE
)

SITUACION_COMPLETO_RX = re.compile(
    r"situacion del nivel\s*:\s*completo",
    re.IGNORECASE
)


def extract_entry_block(text: str, idx: int, max_chars: int = 1600) -> str:
    if not text:
        return ""
    start = max(0, idx - 50)
    end = min(len(text), idx + max_chars)
    chunk = text[start:end]
    cut = re.search(r"\n\s*\n", chunk)
    return chunk[:cut.start()] if cut else chunk


def normalize_key(s: str) -> str:
    s = (s or "").lower().strip()
    s = re.sub(r"\s+", " ", s)
    return s


def count_completed_by_keyword(text_plain: str, keyword_rx: re.Pattern) -> int:
    """
    text_plain debe venir YA sin acentos (strip_accents).
    Regla:
      - Si en el bloque aparece 'actualidad/en curso/cursando' => NO cuenta.
      - Debe tener 'anio de finalizacion/obtencion/graduacion: ...' o 'situacion del nivel: completo'
      - Deduplicación por (primera línea + año)
    """
    if not text_plain:
        return 0

    seen = set()
    count = 0

    for m in keyword_rx.finditer(text_plain):
        entry = extract_entry_block(text_plain, m.start())

        if IN_PROGRESS_RX.search(entry):
            continue

        has_final = FINALIZATION_RX.search(entry) or SITUACION_COMPLETO_RX.search(entry)
        if not has_final:
            continue

        title_hint = entry.split("\n")[0][:140]
        year = ""
        mfy = FINALIZATION_RX.search(entry)
        if mfy:
            year = mfy.group(2)

        key = normalize_key(f"{title_hint}::{year}")
        if key in seen:
            continue

        seen.add(key)
        count += 1

    return count


# =========================
# Patrones (EN TEXTO SIN ACENTOS)
# =========================
RX_DOCTORADO = re.compile(r"\bdoctorad[oa]\b|\bdoctor en\b|\bph\.?\s?d\b", re.IGNORECASE)
RX_MAESTRIA = re.compile(r"\bmaestria\b|\bmagister\b", re.IGNORECASE)
RX_ESPECIALIZACION = re.compile(r"\bespecializacion\b|\bespecialista\b", re.IGNORECASE)
RX_PROFESORADO = re.compile(r"\bprofesorado\b|\bprofesor en\b|\bdocencia universitaria\b", re.IGNORECASE)

# Grado: contempla "Contadora Publica", "Licenciada en Psicologia", "Tecnica Universitaria ..."
RX_GRADO = re.compile(
    r"\blicenciad[oa]\b|"
    r"\bcontador(?:a)?\b|\bcontador(?:a)? publica\b|"
    r"\babogad[oa]\b|"
    r"\bingenier[oa]\b|"
    r"\bmedic[oa]\b|"
    r"\bbioquimic[oa]\b|"
    r"\bfarmaceutic[oa]\b|"
    r"\bodontolog[oa]\b|"
    r"\bpsicolog(?:o|a)\b|\bpsicologia\b|"
    r"\barquitect[oa]\b|"
    r"\bveterinar(?:io|ia)\b|"
    r"\bkinesiolog(?:o|a)\b|"
    r"\bnutricionist[ao]\b|"
    r"\btecnic[oa]\s+universitari[ao]\b",
    re.IGNORECASE
)

RX_POSDOC = re.compile(r"\bposdoctorad[oa]\b|\bpostdoctorad[oa]\b|\bpostdoc\b", re.IGNORECASE)


def count_posdoc_any(text_plain: str) -> int:
    if not text_plain:
        return 0
    seen = set()
    c = 0
    for m in RX_POSDOC.finditer(text_plain):
        entry = extract_entry_block(text_plain, m.start())
        key = normalize_key(entry[:220])
        if key in seen:
            continue
        seen.add(key)
        c += 1
    return c


# =========================
# Categorización por puntaje
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

    # versión sin acentos para matching robusto
    raw_plain = strip_accents(raw_text)

    with st.expander("Ver texto extraído (debug)"):
        st.text_area("Texto", raw_text, height=240)

    # Recorte de "Formación académica" (en original y sin acentos)
    formacion_text = slice_section(
        raw_text,
        start_patterns=[
            r"\bFORMACI[ÓO]N ACAD[ÉE]MICA\b",
            r"\bFormaci[óo]n acad[ée]mica\b",
            r"\bFORMACION ACADEMICA\b",
        ],
        end_patterns=[
            r"\bCARGOS\b",
            r"\bANTECEDENTES\b",
            r"\bPRODUCCI[ÓO]N\b",
            r"\bPUBLICACIONES\b",
            r"\bFINANCIAMIENTO\b",
            r"\bPROYECTOS\b",
            r"\bEXTENSI[ÓO]N\b",
            r"\bEVALUACI[ÓO]N\b",
        ],
        max_len_after=35000
    )
    formacion_plain = strip_accents(formacion_text)

    with st.expander("Ver sección de Formación académica (debug)"):
        st.text_area("Formación académica (recorte)", formacion_text, height=220)

    results = {}
    total = 0.0

    # =========================
    # Cálculo por sección
    # =========================
    for section, cfg in criteria.get("sections", {}).items():
        st.markdown(f"### {section}")
        rows = []
        subtotal_raw = 0.0

        section_name = (section or "").strip().lower()
        is_formacion = section_name.startswith("formación académica") or section_name.startswith("formacion academica")

        items = cfg.get("items", {})
        for item, icfg in items.items():
            item_name = (item or "").strip().lower()
            pattern = icfg.get("pattern", "")
            unit = float(icfg.get("unit_points", 0) or 0)
            item_cap = float(icfg.get("max_points", 0) or 0)

            # Overrides SOLO para Formación (no rompe el resto)
            if is_formacion and "doctorado" in item_name and "final" in item_name:
                c = count_completed_by_keyword(formacion_plain, RX_DOCTORADO)
            elif is_formacion and "maestr" in item_name and "final" in item_name:
                c = count_completed_by_keyword(formacion_plain, RX_MAESTRIA)
            elif is_formacion and "especializ" in item_name and "final" in item_name:
                c = count_completed_by_keyword(formacion_plain, RX_ESPECIALIZACION)
            elif is_formacion and ("título de grado" in item_name or "titulo de grado" in item_name) and "final" in item_name:
                c = count_completed_by_keyword(formacion_plain, RX_GRADO)
            elif is_formacion and ("profesor" in item_name or "docencia universitaria" in item_name) and "final" in item_name:
                c = count_completed_by_keyword(formacion_plain, RX_PROFESORADO)
            elif is_formacion and ("posdoctor" in item_name or "postdoctor" in item_name):
                c = count_posdoc_any(formacion_plain)
            else:
                # default criteria.json (usa texto original)
                c = match_count(pattern, raw_text)

            pts_raw = c * unit
            pts = clip(pts_raw, item_cap)

            rows.append({
                "Ítem": item,
                "Ocurrencias": int(c),
                "Puntaje (tope ítem)": float(pts),
                "Tope ítem": float(item_cap),
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
            sheet = (sec[:31] if sec else "SECCION")
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
