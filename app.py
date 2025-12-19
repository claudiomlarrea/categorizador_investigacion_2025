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


# =========================
# Config
# =========================
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


def fold(s: str) -> str:
    """Lower + sin tildes para comparar robusto."""
    if s is None:
        return ""
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
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


# =========================================================
# 1) EXTRAER BLOQUE "FORMACION ACADEMICA" (robusto)
# =========================================================
def extract_formacion_academica_block(text: str) -> str:
    """
    Extrae el bloque desde 'FORMACION ACADEMICA' (con o sin tildes)
    hasta el próximo encabezado fuerte (todo mayúsculas) o el final.
    """
    t = text or ""
    tf = fold(t)

    # localizar inicio
    m = re.search(r"\bformacion academica\b", tf)
    if not m:
        return ""

    start = m.start()
    # desde allí, tomamos el substring original equivalente
    sub = t[start:]

    # cortar cuando aparece un encabezado en mayúsculas típico (otra sección)
    # (en PDFs suele venir en mayúsculas y con saltos)
    cut = re.search(
        r"\n\s*[A-ZÁÉÍÓÚÑÜ]{6,}(?:\s+[A-ZÁÉÍÓÚÑÜ]{3,})*\s*\n",
        sub
    )
    if cut:
        return sub[:cut.start()].strip()

    return sub.strip()


# =========================================================
# 2) PARSEAR ENTRADAS DE FORMACIÓN Y CLASIFICAR
# =========================================================
FINAL_LABEL_RX = re.compile(
    r"(anio|año)\s+de\s+(finalizacion|finalización|obtencion|obtención|graduacion|graduación)\s*:\s*([0-3]?\d/[01]?\d/(?:19|20)\d{2}|[01]?\d/(?:19|20)\d{2}|(?:19|20)\d{2})",
    re.IGNORECASE
)

RANGO_CERRADO_RX = re.compile(
    r"(?:(?:0?[1-9]|1[0-2])/(?:19|20)\d{2}|(?:19|20)\d{2})\s*-\s*(?:(?:0?[1-9]|1[0-2])/(?:19|20)\d{2}|(?:19|20)\d{2})",
    re.IGNORECASE
)

RANGO_ACTUALIDAD_RX = re.compile(
    r"(?:(?:0?[1-9]|1[0-2])/(?:19|20)\d{2}|(?:19|20)\d{2})\s*-\s*actualidad\b",
    re.IGNORECASE
)

EN_CURSO_RX = re.compile(r"\b(actualidad|en curso|cursando)\b", re.IGNORECASE)


def split_entries(block: str) -> list[str]:
    # separa por líneas en blanco, pero conserva entradas aunque vengan “pegadas”
    parts = [p.strip() for p in re.split(r"\n\s*\n", block) if p.strip()]
    return parts


def entry_title_line(entry: str) -> str:
    # primera línea “real”
    for ln in entry.split("\n"):
        ln2 = ln.strip()
        if ln2:
            return ln2
    return entry.strip()[:80]


def is_in_course(entry: str) -> bool:
    if EN_CURSO_RX.search(entry):
        return True
    if RANGO_ACTUALIDAD_RX.search(entry):
        return True
    return False


def is_finalized(entry: str) -> bool:
    """
    Regla FINALIZADO:
    - NO debe estar en curso/Actualidad
    - y debe tener:
      a) "Año de finalización/obtención/graduación: ...", o
      b) rango cerrado "YYYY - YYYY" / "MM/YYYY - MM/YYYY" (sin Actualidad)
    """
    if is_in_course(entry):
        return False
    if FINAL_LABEL_RX.search(entry):
        return True
    if RANGO_CERRADO_RX.search(entry) and not RANGO_ACTUALIDAD_RX.search(entry):
        return True
    return False


def classify_entry(entry: str) -> str | None:
    """
    Devuelve tipo: doctorado|maestria|especializacion|posdoc|profesorado|grado|otros|None
    """
    ef = fold(entry)

    # posdoc
    if "posdoctor" in ef or "postdoctor" in ef:
        return "posdoc"

    # doctorado
    if re.search(r"\bdoctorad", ef) or re.search(r"\bdoctor\b", ef):
        # ojo: "doctor" puede aparecer en otro contexto, pero en FORMACIÓN suele ser título
        return "doctorado"

    # maestría
    if "maestr" in ef or "magister" in ef or "máster" in ef or "master" in ef:
        return "maestria"

    # especialización / especialista
    if "especializ" in ef or "especialista" in ef:
        return "especializacion"

    # profesorado (como título de grado/uni)
    if "profesorado" in ef or re.search(r"\bprofesor\s+en\b", ef):
        return "profesorado"

    # grado: heurística fuerte (Farah/Young/Vinader)
    # Si hay "Año de finalización" y NO es posgrado, casi seguro es grado/tecnicatura
    if FINAL_LABEL_RX.search(entry):
        # tecnicaturas / títulos de grado típicos
        if re.search(r"\b(tecnic|tecnica|tecnico)\s+universitari", ef):
            return "grado"
        if re.search(r"\blicenci", ef):
            return "grado"
        if re.search(r"\bcontador(a)?\b", ef):
            return "grado"
        if re.search(r"\babogad", ef):
            return "grado"
        if re.search(r"\bingenier", ef):
            return "grado"
        if re.search(r"\bmedic", ef) or re.search(r"\bodontolog", ef) or re.search(r"\bbioquim", ef):
            return "grado"
        if re.search(r"\bfarmac", ef) or re.search(r"\barquitect", ef):
            return "grado"
        # fallback: si tiene año de finalización y no cayó antes, lo consideramos grado
        return "grado"

    # rango cerrado sin “Año de finalización”
    if RANGO_CERRADO_RX.search(entry) and not RANGO_ACTUALIDAD_RX.search(entry):
        # Si no es posgrado por palabra, lo dejamos como otros (no puntúa acá)
        return "otros"

    return None


def count_formacion_titles(block: str) -> dict:
    """
    Cuenta entradas FINALIZADAS por tipo, deduplicando por título.
    """
    counts = {
        "doctorado_fin": 0,
        "maestria_fin": 0,
        "especializacion_fin": 0,
        "grado_fin": 0,
        "profesorado_fin": 0,
        "posdoc_any": 0,  # en curso o finalizado
    }

    seen = {k: set() for k in counts.keys()}

    for entry in split_entries(block):
        tline = entry_title_line(entry)
        tkey = fold(tline)

        tipo = classify_entry(entry)
        if tipo is None:
            continue

        if tipo == "posdoc":
            if tkey not in seen["posdoc_any"]:
                seen["posdoc_any"].add(tkey)
                counts["posdoc_any"] += 1
            continue

        fin = is_finalized(entry)

        if tipo == "doctorado" and fin:
            if tkey not in seen["doctorado_fin"]:
                seen["doctorado_fin"].add(tkey)
                counts["doctorado_fin"] += 1

        elif tipo == "maestria" and fin:
            if tkey not in seen["maestria_fin"]:
                seen["maestria_fin"].add(tkey)
                counts["maestria_fin"] += 1

        elif tipo == "especializacion" and fin:
            if tkey not in seen["especializacion_fin"]:
                seen["especializacion_fin"].add(tkey)
                counts["especializacion_fin"] += 1

        elif tipo == "grado" and fin:
            if tkey not in seen["grado_fin"]:
                seen["grado_fin"].add(tkey)
                counts["grado_fin"] += 1

        elif tipo == "profesorado" and fin:
            if tkey not in seen["profesorado_fin"]:
                seen["profesorado_fin"].add(tkey)
                counts["profesorado_fin"] += 1

    return counts


# =========================================================
# UI
# =========================================================
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

    # Bloque formación (para debug)
    form_block = extract_formacion_academica_block(raw_text)
    form_counts = count_formacion_titles(form_block) if form_block else {
        "doctorado_fin": 0,
        "maestria_fin": 0,
        "especializacion_fin": 0,
        "grado_fin": 0,
        "profesorado_fin": 0,
        "posdoc_any": 0,
    }

    with st.expander("Ver texto extraído (debug)"):
        st.text_area("Texto", raw_text, height=240)

    with st.expander("Ver sección de Formación académica (debug)"):
        if form_block:
            st.text_area("FORMACION ACADEMICA (bloque detectado)", form_block, height=220)
            st.caption(f"Conteos estructurados (finalizados): {form_counts}")
        else:
            st.warning("No se detectó el bloque 'FORMACION ACADEMICA' en este CV.")

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

            c = None

            # =========================================================
            # OVERRIDE ROBUSTO SOLO PARA "Formación académica y complementaria"
            # (evita que regex puntúe 'Actualidad' y arregla Farah/Young/Vinader)
            # =========================================================
            if fold(section) == fold("Formación académica y complementaria"):
                it = fold(item)

                # Doctorado finalizado
                if "doctor" in it and "final" in it:
                    c = form_counts["doctorado_fin"]

                # Maestría finalizada
                elif "maestr" in it and "final" in it:
                    c = form_counts["maestria_fin"]

                # Especialización finalizada
                elif ("especializ" in it or "especialista" in it) and "final" in it:
                    c = form_counts["especializacion_fin"]

                # Título de grado finalizado
                elif ("titulo de grado" in it) or (it.strip() == "grado") or ("grado" in it and "final" in it):
                    c = form_counts["grado_fin"]

                # Profesorado / docencia universitaria finalizado
                elif ("profesor" in it or "docencia universitaria" in it) and "final" in it:
                    c = form_counts["profesorado_fin"]

                # Posdoctorado (en curso o finalizado)
                elif ("posdoctor" in it or "postdoctor" in it):
                    c = form_counts["posdoc_any"]

            # fallback: regex del criteria.json
            if c is None:
                c = match_count(pattern, raw_text)

            pts_raw = c * unit
            pts = clip(pts_raw, item_cap)

            rows.append({
                "Ítem": item,
                "Ocurrencias": int(c),
                "Puntaje (tope ítem)": float(pts),
                "Tope ítem": float(item_cap)
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
