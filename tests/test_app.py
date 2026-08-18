import io
import sys
import types
import unittest
import zipfile
from pathlib import Path

# Las pruebas de transformación DOCX no necesitan las dependencias de interfaz/PDF.
sys.modules.setdefault("streamlit", types.SimpleNamespace())
sys.modules.setdefault("pypdf", types.SimpleNamespace(PdfReader=object))
import app


class AppTests(unittest.TestCase):
    def test_extracts_fields_without_inventing_missing_values(self):
        result = app.extract_contract([
            "CONVENIO Nº ATENEA-582-2025\nASOCIADO: UNIVERSIDAD EAN NIT: 123-4\n"
            "FECHA DE TERMINACIÓN: 30 de junio de 2031\n"
            "OBJETO: Implementación del Programa Jóvenes a la E CLÁUSULA SEGUNDA"
        ])
        self.assertEqual(result["fields"]["numero_contrato_convenio"].value, "ATENEA-582-2025")
        self.assertEqual(result["fields"]["nombre_contratista_asociado"].value, "UNIVERSIDAD EAN")
        self.assertEqual(result["fields"]["nombre_supervisor"].value, app.NOT_FOUND)

    def test_docx_clears_execution_residue_and_preserves_media(self):
        fields = {key: app.Field() for key in [
            "numero_contrato_convenio", "nombre_contratista_asociado", "nombre_identitario",
            "nombre_supervisor", "cargo_supervisor", "fecha_terminacion", "modificaciones", "objeto"
        ]}
        fields["numero_contrato_convenio"] = app.Field("ATENEA-TEST-2026")
        output = app.build_docx(Path("Plantilla_maestra_informe_supervision_ATENEA.docx"),
                                {"fields": fields, "obligations": [app.Field("1. Obligación verificable")]})
        with zipfile.ZipFile(io.BytesIO(output)) as docx:
            xml = docx.read("word/document.xml").decode()
            self.assertIn("ATENEA-TEST-2026", xml)
            self.assertNotIn("SICORE ATENEA", xml)
            self.assertIn("word/media/image1.png", docx.namelist())


if __name__ == "__main__":
    unittest.main()
