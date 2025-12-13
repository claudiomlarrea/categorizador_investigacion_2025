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
st.caption("Incluye exportación a Excel y Word + categoría automática según puntaje total (respetando categoría histórica).")

@st.cache_data
def load_json(path):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)

criteria = load_json("criteria.json")


# === Funciones de extracción de texto ===
def extract_text_docx(file):
    doc = DocxDocument(file)
    text = "\n".join(p.text for p in doc.paragraphs)
    for t in doc.tables:
        for row in t.rows:
            text += "\n" + " | ".join(c.text for c in row.cells)
    return text


def extract_text_pdf(file):
    if not HAVE_PDF:
        raise RuntimeError("Instalá pdfplumber: pip install pdfplumber")
    chunks = []
    with pdfplumber.open(file) as pdf:
        for p in pdf.pages:
            chunks.append(p.extract_text() or "")
    return "\n".join(chunks)


# === Matching genérico ===
def match_count(pattern, text):
    return len(re.findall(pattern, text, re.IGNORECASE)) if pattern else 0


def clip(v, cap):
    return min(v, cap) if cap else v


# === Detección de título / posgrado COMPLETO ===
def titulacion_completa(titulo_regex, text, window_back=250, window_forward=450):
    """
    Cuenta cuántas titulaciones completas hay (Doctorado, Maestría, Especialización,
    Profesorado, etc.) según las reglas:
    - Debe aparecer el título (según regex).
    - En una ventana de texto alrededor:
        * NO debe aparecer 'Actualidad' (titulación en curso).
        * Debe aparecer AL MENOS UNO de:
            - 'Situación del nivel: Completo'
            - 'Año de finalización:' / 'Año de obtención:' / 'Año de graduación:' con año
            - Algún año 19xx o 20xx (para formatos viejos)
    """
    count = 0
    for m in re.finditer(titulo_regex, text, re.IGNORECASE):
        start = max(0, m.start() - window_back)
        end = min(len(text), m.end() + window_forward)
        window = text[start:end]

        # 1) Excluir si aparece "Actualidad" (en curso)
        if re.search(r"Actualidad", window, re.IGNORECASE):
            continue

        # 2) Indicadores de finalización / completitud
        tiene_situacion_completa = re.search(
            r"Situaci[oó]n del nivel:? *Completo", window, re.IGNORECASE
        )

        tiene_anio_finalizacion = re.search(
            r"A[nñ]o de (finalizaci[oó]n|obtenci[oó]n|graduaci[oó]n)\s*:\s*(19|20)\d{2}",
            window,
            re.IGNORECASE,
        )

        tiene_anio_suelto = re.search(r"(19|20)\d{2}", window)

        if not (tiene_situacion_completa or tiene_anio_finalizacion or tiene_anio_suelto):
            continue

        count += 1

    return count


# === Categorización basada en criteria.json (por puntaje) ===
def obtener_categoria_por_puntaje(total, criteria_dict):
    """
    Devuelve (clave_categoria, descripcion_categoria) usando el bloque 'categorias'
    de criteria.json. Elige la categoría con mayor min_points <= total.
    """
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


# === Categorización histórica (piso) ===
def _cat_rank(clave):
    """
    Rank para comparar categorías: I es mejor que II, etc.
    Menor rank = mejor categoría.
    """
    orden = {"I": 1, "II": 2, "III": 3, "IV": 4, "V": 5, "VI": 6}
    return orden.get(clave, 999)


def detectar_categoria_historica(text, criteria_dict):
    """
    Busca una categoría histórica explícita en el CV (piso).
    Por defecto busca 'Investigador Categoría I/II/III/IV/V/VI'.
    Puede ser sobreescrito desde criteria.json en la clave 'categoria_historica'.
    """
    cfg = criteria_dict.get("categoria_historica", {})
    pattern = cfg.get(
        "pattern",
        r"Investigador\s+Categor[ií]a\s*(I|II|III|IV|V|VI)\b"
    )

    # Captura todas las apariciones y se queda con la MEJOR (mínimo rank)
    encontrados = re.findall(pattern, text, flags=re.IGNORECASE)
    if not encontrados:
        return None

    # Normalizar a romano en mayúsculas
    encontrados = [e.upper() for e in encontrados]
    return sorted(encontrados, key=_cat_rank)[0]


def resolver_categoria_final(cat_hist, cat_puntaje):
    """
    Regla: la categoría final es la MEJOR entre:
    - categoría histórica (piso)
    - categoría por puntaje
    """
    if cat_hist and cat_puntaje and cat_puntaje != "Sin categoría":
        return sorted([cat_hist, cat_puntaje], key=_cat_rank)[0]
    return cat_hist or (None if cat_puntaje == "Sin categoría" else cat_puntaje)


uploaded = st.file_uploader("Cargar CV (.docx o .pdf)", type=["docx", "pdf"])
if uploaded:
    ext = uploaded.name.split(".")[-1].lower()
    try:
        raw_text = extract_text_docx(uploaded) if ext == "docx" else extract_text_pdf(uploaded)
    except Exception as e:
        st.error(str(e))
        st.stop()

    st.success(f"Archivo cargado: {uploaded.name}")
    with st.expander("Ver texto extraído (debug)"):
        st.text_area("Texto", raw_text, height=220)

    # === Categoría histórica (piso) ===
    categoria_historica = detectar_categoria_historica(raw_text, criteria)

    results = {}
    total = 0.0

    # === Cálculo de puntajes por sección ===
    for section, cfg in criteria["sections"].items():
        st.markdown(f"### {section}")
        rows = []
        subtotal_raw = 0.0

        for item, icfg in cfg.get("items", {}).items():
            pattern = icfg.get("pattern", "")

            # Lógica especial para titulaciones completas
            if section == "Formación académica y complementaria" and item in [
                "Doctorado",
                "Maestría",
                "Especialización",
                "Profesorados universitarios"
            ]:
                if item == "Doctorado":
                    c = titulacion_completa(r"Doctorado", raw_text)
                elif item == "Maestría":
                    c = titulacion_completa(r"Maestr[ií]a|Mag[íi]ster", raw_text)
                elif item == "Especialización":
                    c = titulacion_completa(r"Especializaci[oó]n|Especialista", raw_text)
                else:  # Profesorados universitarios
                    c = titulacion_completa(r"Profesor en|Profesorado en", raw_text)
            else:
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

    # === Categoría por puntaje (como antes) ===
    cat_puntaje, desc_cat_puntaje = obtener_categoria_por_puntaje(total, criteria)

    # === Categoría final (NUEVO: respeta piso histórico) ===
    cat_final = resolver_categoria_final(categoria_historica, cat_puntaje)

    if cat_final is None:
        categoria_label_final = "Sin categoría"
        desc_cat_final = desc_cat_puntaje if cat_puntaje != "Sin categoría" else ""
    else:
        categoria_label_final = f"Categoría {cat_final}"
        # la descripción la tomamos del bloque categorias (si existe)
        desc_cat_final = criteria.get("categorias", {}).get(cat_final, {}).get("descripcion", "") or desc_cat_puntaje

    # Labels auxiliares para mostrar
    categoria_label_hist = f"Categoría {categoria_historica}" if categoria_historica else "No detectada"
    categoria_label_puntaje = "Sin categoría" if cat_puntaje == "Sin categoría" else f"Categoría {cat_puntaje}"

    st.markdown("---")
    st.subheader("Puntaje total y categoría")
    st.metric("Total acumulado", f"{total:.1f}")
    st.metric("Categoría histórica (piso)", categoria_label_hist)
    st.metric("Categoría por puntaje", categoria_label_puntaje)
    st.metric("Categoría final", categoria_label_final)

    if desc_cat_final:
        st.info(f"Descripción de la categoría: {desc_cat_final}")

    # === Exportaciones ===
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
        resumen.loc[len(resumen)] = ["CATEGORÍA HISTÓRICA (PISO)", categoria_label_hist]
        resumen.loc[len(resumen)] = ["CATEGORÍA POR PUNTAJE", categoria_label_puntaje]
        resumen.loc[len(resumen)] = ["CATEGORÍA FINAL", categoria_label_final]
        resumen.to_excel(writer, sheet_name="RESUMEN", index=False)

    st.download_button(
        "Descargar Excel",
        data=out.getvalue(),
        file_name="valoracion_cv.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

    # Word
    def export_word(results_dict, total_pts, cat_hist_label, cat_puntaje_label, cat_final_label, cat_desc_final):
        doc = DocxDocument()
        p = doc.add_paragraph("Universidad Católica de Cuyo — Secretaría de Investigación")
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph("Informe de valoración de CV").alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph("")
        doc.add_paragraph(f"Puntaje total: {total_pts:.1f}")
        doc.add_paragraph(f"Categoría histórica (piso): {cat_hist_label}")
        doc.add_paragraph(f"Categoría por puntaje: {cat_puntaje_label}")
        doc.add_paragraph(f"Categoría final: {cat_final_label}")
        if cat_desc_final:
            doc.add_paragraph(cat_desc_final)

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
        data=export_word(results, total, categoria_label_hist, categoria_label_puntaje, categoria_label_final, desc_cat_final),
        file_name="informe_valoracion_cv.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True
    )

else:
    st.info("Subí un archivo para iniciar la valoración.")
