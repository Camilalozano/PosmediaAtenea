"""Aplicación Streamlit para prediligenciar informes de supervisión ATENEA.

La extracción es deliberadamente conservadora: un dato que no se puede respaldar
con texto de la minuta queda marcado como no especificado. El documento maestro se
lee, pero nunca se modifica.
"""

from __future__ import annotations

import csv
import hashlib
import io
import json
import re
import unicodedata
import zipfile
from copy import deepcopy
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree as ET

import streamlit as st
from pypdf import PdfReader


MASTER = Path(__file__).with_name("Plantilla_maestra_informe_supervision_ATENEA.docx")
NOT_FOUND = "No especificado en la minuta suministrada"
NO_CHANGES = "No se indican modificaciones en la minuta suministrada"
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
ALLOWED = {".pdf"}


@dataclass
class Field:
    value: str = NOT_FOUND
    source: str = "No localizado"
    confidence: str = "bajo"


def normalize(text: str) -> str:
    """Compacta espacios sin cambiar letras ni puntuación contractual."""
    return re.sub(r"[ \t]+", " ", re.sub(r"\r", "", text)).strip()


def safe_name(value: str) -> str:
    value = unicodedata.normalize("NFKD", value).encode("ascii", "ignore").decode()
    return re.sub(r"[^A-Za-z0-9_-]+", "_", value).strip("_")[:80] or "PENDIENTE"


def pdf_text(data: bytes) -> tuple[str, list[str]]:
    reader = PdfReader(io.BytesIO(data))
    if reader.is_encrypted:
        try:
            reader.decrypt("")
        except Exception as exc:
            raise ValueError("El PDF está cifrado y no puede leerse.") from exc
    pages = [page.extract_text() or "" for page in reader.pages]
    text = "\n\n".join(pages)
    if len(re.sub(r"\s", "", text)) < 100:
        raise ValueError("El PDF no contiene texto legible; aplique OCR antes de cargarlo.")
    return text, pages


def find_field(pages: list[str], patterns: Iterable[str], flags: int = re.I) -> Field:
    for page_no, page in enumerate(pages, 1):
        compact = normalize(page)
        for pattern in patterns:
            match = re.search(pattern, compact, flags)
            if match:
                value = normalize(match.group(1)).strip(" :;.-")
                if value:
                    excerpt = normalize(match.group(0))[:240]
                    return Field(value, f"Página {page_no}: {excerpt}", "alto")
    return Field()


def extract_obligations(pages: list[str]) -> list[Field]:
    """Extrae listas numeradas cercanas a compromisos/obligaciones de la contraparte."""
    text = "\n".join(pages)
    headings = list(re.finditer(
        r"(?i)(?:compromisos|obligaciones)\s+(?:espec[ií]fic[oa]s?|de\s+la\s+(?:ies|instituci[oó]n|contraparte|asociad[oa]))",
        text,
    ))
    candidates: list[list[Field]] = []
    for heading in headings:
        block = text[heading.end(): heading.end() + 24000]
        stop = re.search(r"\n\s*(?:CL[AÁ]USULA|OBLIGACIONES?\s+DE\s+ATENEA|COMPROMISOS?\s+DE\s+ATENEA)", block, re.I)
        if stop:
            block = block[:stop.start()]
        matches = list(re.finditer(r"(?ms)(?:^|\n)\s*(\d{1,2})[.)]\s*(.+?)(?=(?:\n\s*\d{1,2}[.)]\s)|\Z)", block))
        found = []
        for match in matches:
            body = normalize(match.group(2).replace("\n", " "))
            if 10 < len(body) < 2500:
                page_no = text[:heading.end() + match.start()].count("\n\n") + 1
                found.append(Field(f"{match.group(1)}. {body}", f"Sección de compromisos, página aproximada {page_no}", "alto"))
        if found:
            candidates.append(found)
    return max(candidates, key=len, default=[])


def extract_contract(pages: list[str]) -> dict:
    fields = {
        "numero_contrato_convenio": find_field(pages, [r"(?:convenio|contrato)\s+(?:de\s+asociaci[oó]n\s+)?(?:n[oº°.]|n[uú]mero)?\s*[:#-]?\s*([A-ZÁ-Ú]{2,15}[- ]\d{2,6}[-/]\d{4})"]),
        "tipo_instrumento": find_field(pages, [r"\b((?:convenio|contrato)(?:\s+de\s+[a-zá-ú ]{2,40})?)\s+(?:n[oº°.]|n[uú]mero)"]),
        "nombre_contratista_asociado": find_field(pages, [r"(?:asociad[oa]|contratista|cooperante)\s*:\s*([^\n]{3,150}?)(?=\s+(?:nit|identificaci[oó]n|representante|objeto)\b|$)"]),
        "nombre_identitario": Field(),
        "nit_contratista": find_field(pages, [r"\bNIT\s*[:.]?\s*([0-9][0-9.\- ]{6,20})"]),
        "representante_legal": find_field(pages, [r"representante\s+legal\s*:?\s*([A-ZÁ-ÚÑ][A-ZÁ-ÚÑ ]{4,100})"]),
        "identificacion_representante": find_field(pages, [r"representante\s+legal.{0,160}?(?:c[eé]dula|C\.?C\.?)\s*(?:n[oº°.]|n[uú]mero)?\s*([0-9.]{5,20})"]),
        "nombre_supervisor": find_field(pages, [r"supervisor(?:a)?\s*:?\s*([A-ZÁ-ÚÑ][A-ZÁ-ÚÑ ]{4,100})"]),
        "cargo_supervisor": find_field(pages, [r"supervisi[oó]n.{0,100}?(Gerente\s+de\s+Educaci[oó]n\s+Posmedia|Subgerente[^,.;\n]{2,80}|Gerente[^,.;\n]{2,80})"]),
        "fecha_suscripcion": find_field(pages, [r"fecha\s+de\s+suscripci[oó]n\s*:?\s*([^.;\n]{5,50})"]),
        "fecha_inicio": find_field(pages, [r"fecha\s+de\s+inicio\s*:?\s*([^.;\n]{5,50})"]),
        "fecha_terminacion": find_field(pages, [r"fecha\s+de\s+terminaci[oó]n\s*:?\s*([^.;\n]{5,60})", r"hasta\s+el\s+(\d{1,2}\s+de\s+[a-zá-ú]+\s+de\s+20\d{2})"]),
        "plazo": find_field(pages, [r"(?:plazo|duraci[oó]n)\s*:?\s*([^.;\n]{4,150})"]),
        "lugar_ejecucion": find_field(pages, [r"lugar\s+de\s+ejecuci[oó]n\s*:?\s*([^.;\n]{3,100})"]),
        "valor_total": find_field(pages, [r"valor\s+(?:total\s+)?(?:del\s+(?:convenio|contrato))?\s*:?\s*(\$\s*[0-9.,]+)"]),
        "aporte_atenea": find_field(pages, [r"aporte\s+(?:de\s+)?(?:la\s+agencia\s+)?atenea\s*:?\s*(\$\s*[0-9.,]+)"]),
        "aporte_contraparte": find_field(pages, [r"aporte\s+(?:de\s+)?(?:la\s+)?(?:ies|contraparte|asociad[oa])\s*:?\s*(\$\s*[0-9.,]+)"]),
        "modificaciones": Field(NO_CHANGES, "Documento inicial suministrado; no se localizaron otrosíes", "medio"),
        "objeto": find_field(pages, [r"(?:objeto(?:\s+del\s+(?:convenio|contrato))?)\s*:?\s*(.{20,1200}?)(?=\s+(?:CL[AÁ]USULA|PLAZO|VALOR|OBLIGACIONES|COMPROMISOS)\b)"]),
    }
    # El supervisor puede estar definido exclusivamente por cargo; no lo convertimos en persona.
    if fields["nombre_supervisor"].value == fields["cargo_supervisor"].value:
        fields["nombre_supervisor"] = Field()
    obligations = extract_obligations(pages)
    return {"fields": fields, "obligations": obligations}


def set_cell(cell: ET.Element, text: str) -> None:
    paragraphs = cell.findall(f"{W}p")
    if not paragraphs:
        paragraphs = [ET.SubElement(cell, f"{W}p")]
    first = paragraphs[0]
    for child in list(cell):
        if child.tag == f"{W}p" and child is not first:
            cell.remove(child)
    runs = first.findall(f"{W}r")
    run = runs[0] if runs else ET.SubElement(first, f"{W}r")
    for other in runs[1:]:
        first.remove(other)
    for node in list(run):
        if node.tag != f"{W}rPr":
            run.remove(node)
    node = ET.SubElement(run, f"{W}t")
    node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    node.text = text


def build_docx(master: Path, result: dict) -> bytes:
    """Edita solo celdas variables en una copia binaria del maestro."""
    with zipfile.ZipFile(master) as source:
        files = {name: source.read(name) for name in source.namelist()}
    root = ET.fromstring(files["word/document.xml"])
    tables = root.findall(f".//{W}tbl")
    rows = tables[0].findall(f"{W}tr")
    mapping = {
        0: "numero_contrato_convenio", 1: "nombre_contratista_asociado", 2: "nombre_identitario",
        3: "nombre_supervisor", 4: "cargo_supervisor", 5: "fecha_terminacion", 6: "modificaciones", 12: "objeto",
    }
    for row_no, key in mapping.items():
        cells = rows[row_no].findall(f"{W}tc")
        target = cells[-1] if row_no != 12 else rows[row_no + 1].findall(f"{W}tc")[0]
        set_cell(target, result["fields"][key].value)
    # Los campos de ejecución deben quedar vacíos, pues una minuta no prueba ejecución.
    for row_no in range(7, 12):
        set_cell(rows[row_no].findall(f"{W}tc")[-1], "")
    obligation_table = tables[1]
    old_rows = obligation_table.findall(f"{W}tr")[1:]
    model = deepcopy(old_rows[0])
    for row in old_rows:
        obligation_table.remove(row)
    obligations = result["obligations"] or [Field(NOT_FOUND)]
    for obligation in obligations:
        row = deepcopy(model)
        cells = row.findall(f"{W}tc")
        set_cell(cells[0], obligation.value)
        set_cell(cells[1], "")
        set_cell(cells[2], "")
        obligation_table.append(row)
    # No prejuzgar cumplimiento ni reutilizar datos residuales del maestro.
    for row in tables[2].findall(f"{W}tr")[2:]:
        for cell in row.findall(f"{W}tc")[1:]:
            set_cell(cell, "")
    files["word/document.xml"] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as target:
        for name, data in files.items():
            target.writestr(name, data)
    return output.getvalue()


def pending(result: dict) -> list[str]:
    return [key for key, field in result["fields"].items() if field.value == NOT_FOUND]


def audit_json(result: dict) -> bytes:
    serial = {"campos": {k: asdict(v) for k, v in result["fields"].items()},
              "obligaciones_compromisos_especificos": [asdict(v) for v in result["obligations"]]}
    return json.dumps(serial, ensure_ascii=False, indent=2).encode()


def process(data: bytes, filename: str) -> dict:
    text, pages = pdf_text(data)
    if not re.search(r"(?i)\b(?:contrato|convenio|minuta)\b", text):
        raise ValueError("No corresponde claramente a una minuta, contrato o convenio.")
    result = extract_contract(pages)
    number = result["fields"]["numero_contrato_convenio"].value
    contractor = result["fields"]["nombre_contratista_asociado"].value
    stem = f"Informe_supervision_{safe_name(number)}_{safe_name(contractor)}"
    return {"origin": filename, "result": result, "name": stem + ".docx",
            "docx": build_docx(MASTER, result), "audit": audit_json(result)}


def csv_summary(rows: list[dict]) -> bytes:
    output = io.StringIO()
    fields = ["archivo de origen", "número contractual", "contratista o asociado", "estado",
              "archivo generado", "campos pendientes", "observaciones"]
    writer = csv.DictWriter(output, fieldnames=fields)
    writer.writeheader(); writer.writerows(rows)
    return output.getvalue().encode("utf-8-sig")


def zip_inputs(data: bytes) -> list[tuple[str, bytes]]:
    seen, inputs = set(), []
    with zipfile.ZipFile(io.BytesIO(data)) as bundle:
        for info in bundle.infolist():
            name = Path(info.filename)
            if info.is_dir() or name.name.startswith((".", "~$")) or "__MACOSX" in name.parts:
                continue
            content = bundle.read(info)
            digest = hashlib.sha256(content).digest()
            if digest not in seen:
                seen.add(digest); inputs.append((name.name, content))
    return inputs


def main() -> None:
    st.set_page_config(page_title="Informes de supervisión ATENEA", page_icon="📄", layout="wide")
    st.title("Prediligenciamiento de informes de supervisión ATENEA")
    st.info("La herramienta extrae datos verificables de la minuta. El resultado requiere revisión y aprobación humana del supervisor.")
    uploaded = st.file_uploader("Cargue una minuta, contrato, convenio o ZIP", type=["pdf", "zip"])
    if uploaded is None:
        st.write("Carga una minuta, contrato o convenio, o un archivo `.zip` con varios documentos para procesamiento masivo. Extraeré la información verificable, diligenciaré una copia de la plantilla oficial por cada contrato y señalaré los campos que requieran revisión.")
        return
    items = zip_inputs(uploaded.getvalue()) if uploaded.name.lower().endswith(".zip") else [(uploaded.name, uploaded.getvalue())]
    products, rows = [], []
    for name, data in items:
        if Path(name).suffix.lower() not in ALLOWED:
            rows.append({"archivo de origen": name, "número contractual": "", "contratista o asociado": "",
                         "estado": "No procesado: formato no compatible", "archivo generado": "",
                         "campos pendientes": "", "observaciones": "Solo se admiten PDF con texto."}); continue
        try:
            product = process(data, name); products.append(product)
            result = product["result"]; missing = pending(result)
            rows.append({"archivo de origen": name, "número contractual": result["fields"]["numero_contrato_convenio"].value,
                         "contratista o asociado": result["fields"]["nombre_contratista_asociado"].value,
                         "estado": "Procesado con campos pendientes" if missing else "Procesado",
                         "archivo generado": product["name"], "campos pendientes": ", ".join(missing),
                         "observaciones": f"{len(result['obligations'])} obligaciones extraídas"})
        except Exception as exc:
            status = "No procesado: archivo ilegible" if "legible" in str(exc) or "cifrado" in str(exc) else "No procesado: no corresponde a una minuta, contrato o convenio"
            rows.append({"archivo de origen": name, "número contractual": "", "contratista o asociado": "",
                         "estado": status, "archivo generado": "", "campos pendientes": "", "observaciones": str(exc)})
    st.dataframe(rows, use_container_width=True)
    if len(items) == 1 and products:
        product = products[0]
        st.download_button("Descargar informe Word editable", product["docx"], product["name"],
                           "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        st.download_button("Descargar trazabilidad (fuentes y confianza)", product["audit"], product["name"].replace(".docx", "_trazabilidad.json"), "application/json")
    elif products:
        bundle = io.BytesIO()
        with zipfile.ZipFile(bundle, "w", zipfile.ZIP_DEFLATED) as archive:
            for product in products:
                archive.writestr(product["name"], product["docx"])
                archive.writestr(product["name"].replace(".docx", "_trazabilidad.json"), product["audit"])
            archive.writestr("Resumen_procesamiento_masivo.csv", csv_summary(rows))
        st.download_button("Descargar resultados", bundle.getvalue(), "Informes_supervision_ATENEA.zip", "application/zip")


if __name__ == "__main__":
    main()
