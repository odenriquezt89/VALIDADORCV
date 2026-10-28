import io
import pytest
from fastapi.testclient import TestClient
from reportlab.pdfgen import canvas
from api import app, VacancyRequirements, build_response

client = TestClient(app)

def create_pdf(text: str) -> bytes:
    buf = io.BytesIO()
    c = canvas.Canvas(buf)
    y = 750
    for line in text.split("\n"):
        c.drawString(50, y, line)
        y -= 15
    c.save()
    buf.seek(0)
    return buf.getvalue()

@pytest.fixture
def vacante():
    return VacancyRequirements(
        codigo_vacante="VAC-TEST",
        nombre_puesto="Coordinador Médico",
        meses_experiencia=24,
        requiere_tarjeta=True,
        skills_obligatorias=["Resolución 3100", "Auditoría médica"]
    )

def test_knockout_sin_tarjeta(vacante):
    res = build_response("Candidato A", "Médico con 4 años de experiencia.", vacante)
    assert res.fase_1_knockout == "DESCARTADO"
    assert res.score_cv == 0.0

def test_candidato_cien_puntos(vacante):
    cv = """
    Médico Coordinador con Tarjeta Profesional vigente.
    4 años de experiencia laboral. Manejo de Resolución 3100 y Auditoría médica.
    Reducción del 20% en glosas, 500 pacientes atendidos y ahorro de $ 10000000.
    Enero 2022 a Diciembre 2025.
    """
    res = build_response("Candidato Top", cv, vacante)
    assert res.fase_1_knockout == "APROBADO"
    assert res.score_cv == 100.0

def test_seguridad_api_key(vacante):
    pdf = create_pdf("Texto")
    r_no_key = client.post("/evaluate/cv", files={"cv_file": ("a.pdf", pdf, "application/pdf")}, data={"vacante_data": vacante.model_dump_json()})
    assert r_no_key.status_code == 401

    r_inact = client.post("/evaluate/cv", files={"cv_file": ("a.pdf", pdf, "application/pdf")}, data={"vacante_data": vacante.model_dump_json()}, headers={"X-API-KEY": "key_live_evaluador_inactivo"})
    assert r_inact.status_code == 403

def test_batch_ordenamiento(vacante):
    p1 = create_pdf("Médico con Tarjeta Profesional. 2 años de experiencia. Resolución 3100.")
    p2 = create_pdf("Médico Coordinador con Tarjeta Profesional. 4 años de experiencia. Resolución 3100 y Auditoría médica. 30% mejora, 200 pacientes, $ 5000. Enero 2024.")
    files = [("cv_files", ("basico.pdf", p1, "application/pdf")), ("cv_files", ("experto.pdf", p2, "application/pdf"))]
    r = client.post("/evaluate/batch", files=files, data={"vacante_data": vacante.model_dump_json()}, headers={"X-API-KEY": "key_live_evaluador_alpha_2026"})
    assert r.status_code == 200
    body = r.json()
    assert body["candidatos_ordenados"][0]["candidato"] == "Experto"
