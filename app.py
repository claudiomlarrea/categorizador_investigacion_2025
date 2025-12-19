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

    # Unificar guiones típicos
    text = text.replace("\u2013", "-").replace("\u2014", "-").replace("\u2212", "-")

    # Quitar guionado por salto de línea: "inves-\ntigación" -> "investigación"
    text = re.sub(r"(\w)-\n(\w)", r"\1\2", text)

    # Normalizar saltos de línea
    text = text.replace("\r\n", "\n").replace("\r", "\n")

    # Normalizar espacios
    text = re.sub(r"[ \t]+", " ", text)

    # Compactar líneas vacías
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
# Categoría según puntaje
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

# ==========================================================
# PARSEO ROBUSTO: Formación académica (FINALIZADO vs ACTUALIDAD)
# ==========================================================
HEADING_FORMACION = [
    r"\bFORMACI[ÓO]N ACAD[ÉE]MICA\b",
    r"\bFormaci[óo]n acad[ée]mica\b",
]
STOP_HEADINGS = [
    r"\bANTECEDENTES\b",
    r"\bPRODUCCI[ÓO]N\b",
    r"\bPUBLICACIONES\b",
    r"\bFINANCIAMIENTO\b",
    r"\bPROYECTOS\b",
    r"\bCARGOS\b",
    r"\bACTIVIDADES\b",
]

def extract_section_block(text: str, heading_patterns, stop_patterns, max_len=25000) -> str:
    """
    Extrae un bloque desde el primer heading hasta el primer stop heading posterior.
    Si no encuentra, devuelve string vacío.
    """
    if not text:
        return ""

    # buscar inicio
    start_idx = None
    for hp in heading_patterns:
        m = re.search(hp, text, re.IGNORECASE)
        if m:
            start_idx = m.start()
            break
    if start_idx is None:
        return ""

    tail = text[start_idx:start_idx + max_len]

    # buscar fin (primer stop heading, pero dejando que no corte en la misma línea inmediata)
    end_idx = None
    for sp in stop_patterns:
        m2 = re.search(sp, tail, re.IGNORECASE)
        if m2:
            # Evitar cortar si el stop aparece muy cerca del inicio (ruido)
            if m2.start() > 150:
                end_idx = m2.start()
                break

    return tail[:end_idx] if end_idx else tail

def has_completion_marker(window: str) -> bool:
    """
    SOLO consideramos finalizado si hay marcador explícito.
    IMPORTANTE: NO usar "año suelto" porque rompe con '2018 - Actualidad'.
    """
    if not window:
        return False

    # Situación del nivel completo (CVar clásico)
    if re.search(r"Situaci[oó]n del nivel\s*:?\s*Completo", window, re.IGNORECASE):
        return True

    # Año de finalización/obtención/graduación (acepta MM/YYYY o YYYY)
    if re.search(
        r"A[nñ]o de (finalizaci[oó]n|obtenci[oó]n|graduaci[oó]n)\s*:\s*(\d{2}/)?(19|20)\d{2}",
        window,
        re.IGNORECASE,
    ):
        return True

    # Algunos CV vienen como "Finalizado" / "Graduado" explícito
    if re.search(r"\b(finalizado|finalizada|graduado|graduada|egresado|egresada|titulad[oa])\b", window, re.IGNORECASE):
        return True

    return False

def has_actualidad_marker(window: str) -> bool:
    if not window:
        return False
    return bool(re.search(r"\bActualidad\b", window, re.IGNORECASE))

def count_completed_degree_by_keyword(block: str, keyword_regex: str, window_back=120, window_forward=420) -> int:
    """
    Cuenta ocurrencias de un tipo de título (doctorado/maestría/especialización/etc.)
    pero SOLO si en la ventana hay marcador de completitud y NO hay 'Actualidad'.
    """
    if not block:
        return 0

    count = 0
    for m in re.finditer(keyword_regex, block, re.IGNORECASE):
        start = max(0, m.start() - window_back)
        end = min(len(block), m.end() + window_forward)
        window = block[start:end]

        # si el ítem es en curso, NO contar
        if has_actualidad_marker(window):
            # aunque tenga un año (inicio), no cuenta
            continue

        # debe tener marcador explícito de finalización
        if not has_completion_marker(window):
            continue

        count += 1

    return count

def count_completed_grado(block: str) -> int:
    """
    Grado finalizado: buscamos 'Año de finalización: ...' y verificamos
    que en los ~250 caracteres previos haya un título típico de grado
    (Licenciado/a, Contador/a, Médico/a, Ingeniero/a, Abogado/a, Bioquímico/a,
    Farmacéutico/a, Arquitecto/a, Odontólogo/a, etc.)
    También incluye 'Técnica/o Universitaria/o' (como el caso Periodismo).
    Excluye si cerca aparece 'Actualidad'.
    """
    if not block:
        return 0

    # localizar cada "Año de finalización"
    end_markers = list(re.finditer(
        r"A[nñ]o de (finalizaci[oó]n|obtenci[oó]n|graduaci[oó]n)\s*:\s*(\d{2}/)?(19|20)\d{2}",
        block,
        re.IGNORECASE
    ))

    if not end_markers:
        return 0

    grado_keywords = re.compile(
        r"\b("
        r"Licenciad[oa]|"
        r"Contador[ae]?\s+P[úu]blic[oa]|"
        r"M[ée]dic[oa]|"
        r"Ingenier[oa]|"
        r"Abogad[oa]|"
        r"Bioqu[ií]mic[oa]|"
        r"Farmac[ée]utic[oa]|"
        r"Arquitect[oa]|"
        r"Odont[óo]log[oa]|"
        r"Kinesi[oó]log[oa]|"
        r"Nutricionist[ae]|"
        r"Psic[oó]log[oa]|"
        r"Enfermer[oa]|"
        r"Veterinar[ioia]|"
        r"Traductor[ae]?|"
        r"Profesor(?:\s+en)?|Profesorado\b|"   # usuario pidió profesorados como grado
        r"T[ée]cnic[oa]\s+Universitari[oa]|"
        r"T[ée]cnico\s+Universitario"
        r")\b",
        re.IGNORECASE
    )

    found = 0
    # DEDUP suave: muchos CV repiten el bloque en tablas; tomamos máximo razonable 6
    for m in end_markers:
        # ventana alrededor del año
        start = max(0, m.start() - 260)
        end = min(len(block), m.end() + 120)
        window = block[start:end]

        if has_actualidad_marker(window):
            continue

        # debe haber un keyword de grado cerca
        if not grado_keywords.search(window):
            continue

        found += 1

    # Evitar inflar por repetición de PDF: si detecta > 6, recorta a 6 (ajustable)
    return min(found, 6)

def compute_formacion_counts(full_text: str) -> dict:
    """
    Devuelve conteos robustos para los ítems críticos de formación.
    """
    block = extract_section_block(full_text, HEADING_FORMACION, STOP_HEADINGS)

    return {
        "doctorado_final": count_completed_degree_by_keyword(block, r"\bDoctorado\b|\bDoctor en\b"),
        "maestria_final": count_completed_degree_by_keyword(block, r"\bMaestr[ií]a\b|\bMag[ií]ster\b"),
        "especializacion_final": count_completed_degree_by_keyword(block, r"\bEspecializaci[oó]n\b|\bEspecialista\b"),
        "profesorado_final": count_completed_degree_by_keyword(block, r"\bProfesorado\b|\bProfesor en\b|\bDocencia universitaria\b"),
        "grado_final": count_completed_grado(block),
    }, block

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

    # Debug general
    with st.expander("Ver texto extraído (debug)"):
        st.text_area("Texto", raw_text, height=260)

    # Conteos robustos de formación
    form_counts, form_block = compute_formacion_counts(raw_text)

    with st.expander("Ver sección de Formación académica (debug)"):
        if form_block.strip():
            st.text_area("Bloque Formación académica", form_block, height=260)
            st.write("Conteos robustos (finalizados):", form_counts)
        else:
            st.warning("No se pudo extraer el bloque de 'Formación académica' (heading no encontrado).")

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

            # --- OVERRIDE CLAVE: Formación académica FINALIZADA ---
            # Identificamos por nombre del ítem (NO tocamos criteria.json)
            item_l = (item or "").strip().lower()
            section_l = (section or "").strip().lower()

            c = None
            if "formación académica" in section_l or "formacion academica" in section_l:
                if item_l.startswith("doctorado") and "final" in item_l:
                    c = form_counts["doctorado_final"]
                elif item_l.startswith("maestr") and "final" in item_l:
                    c = form_counts["maestria_final"]
                elif item_l.startswith("especial") and "final" in item_l:
                    c = form_counts["especializacion_final"]
                elif (item_l.startswith("profesor") or "docencia universitaria" in item_l) and "final" in item_l:
                    c = form_counts["profesorado_final"]
                elif ("título de grado" in item_l or "titulo de grado" in item_l) and "final" in item_l:
                    c = form_counts["grado_final"]

            # si no aplica override, usamos criteria.json
            if c is None:
                c = match_count(pattern, raw_text)

            pts_raw = c * unit
            pts = clip(pts_raw, item_cap)

            rows.append({
                "Ítem": item,
                "Ocurrencias": int(c),
                "Puntaje (tope ítem)": float(pts),
                "Tope ítem": float(item_cap),
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
