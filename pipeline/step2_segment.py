# pipeline/step2_segment.py
import json
import re
from pathlib import Path


def load_text(path: str) -> str:
    return Path(path).read_text(encoding="utf-8", errors="ignore")


def load_section_markers(path: str) -> dict:
    return json.loads(Path(path).read_text(encoding="utf-8", errors="ignore"))


def split_sections(text: str, markers: dict) -> dict:
    """
    Detecta secciones por encabezados conocidos (tolerantes a variantes).
    Devuelve dict: { "Sección": "texto de la sección" }
    """
    upper_text = text.upper()

    # Buscar posiciones de encabezados
    hits = []
    for section, names in markers.items():
        for name in names:
            name_u = name.upper().strip()
            idx = upper_text.find(name_u)
            if idx != -1:
                hits.append((idx, section, name_u))

    hits.sort(key=lambda x: x[0])

    sections = {}
    for i, (start, section, _name) in enumerate(hits):
        end = hits[i + 1][0] if i + 1 < len(hits) else len(text)
        chunk = text[start:end].strip()
        # Si el mismo nombre aparece dos veces, nos quedamos con el más largo (suele ser el real)
        if section not in sections or len(chunk) > len(sections[section]):
            sections[section] = chunk

    return sections


# Separadores típicos de entradas en CVAR/SIGEVA (conservador)
ENTRY_SPLIT_RE = re.compile(
    r"\n(?=(\d{4}\s*[-–]|"
    r"\d{2}/\d{4}\s*[-–]|"
    r"Evento:|"
    r"Rol:|"
    r"Tesista:|"
    r"Becario/a:|"
    r"Becario:|"
    r"Pasante:))",
    re.IGNORECASE
)


def split_entries(section_text: str) -> list[str]:
    """
    Parte una sección en "entradas" (bloques) sin perder información.
    """
    # Quita el encabezado si está en la primera línea (para que no sea una 'entrada')
    lines = section_text.splitlines()
    if lines:
        first = lines[0].strip()
        if len(first) < 80 and first.upper() == first:
            section_text = "\n".join(lines[1:]).strip()

    if not section_text:
        return []

    parts = ENTRY_SPLIT_RE.split("\n" + section_text)  # fuerza match al inicio si aplica
    # parts alterna texto y grupo capturado: re.split con grupo captura el separador
    merged = []
    current = ""

    for p in parts:
        if not p:
            continue
        # si parece "inicio de entrada", cerramos y empezamos
        if re.match(r"^(\d{4}\s*[-–]|\d{2}/\d{4}\s*[-–]|Evento:|Rol:|Tesista:|Becario|Pasante:)", p, re.IGNORECASE):
            if current.strip():
                merged.append(current.strip())
            current = p
        else:
            current += "\n" + p if current else p

    if current.strip():
        merged.append(current.strip())

    # Limpieza final: eliminar bloques demasiado chicos
    merged = [m for m in merged if len(m) > 20]

    return merged


def make_id(prefix: str, i: int) -> str:
    prefix = re.sub(r"[^A-Z0-9]+", "_", prefix.upper()).strip("_")
    prefix = prefix[:10] if prefix else "SEC"
    return f"{prefix}_{i:03d}"


def build_items(sections: dict) -> dict:
    """
    Devuelve:
    {
      "Sección": [
         {"id": "...", "raw_text": "...", "source_section": "..."}
