# Prediligenciamiento de informes de supervisión ATENEA

Aplicación Streamlit que recibe una minuta, contrato o convenio en PDF (o un ZIP),
extrae exclusivamente información respaldada por el documento y crea una copia
editable de `Plantilla_maestra_informe_supervision_ATENEA.docx`.

## Ejecución

```bash
python -m pip install -r requirements.txt
streamlit run app.py
```

La aplicación no consulta Internet ni altera el maestro. Los campos de ejecución,
evidencia, periodo, seguridad social, avance y SECOP II permanecen vacíos si solo se
aporta una minuta. Cada resultado incluye un JSON de trazabilidad con fuente y nivel
de confianza. Los lotes producen un ZIP con un Word y una trazabilidad por contrato,
más un resumen CSV.

> El archivo generado es un prediligenciamiento y no sustituye la revisión ni la
> aprobación del supervisor.
