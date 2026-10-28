"""Pruebas de humo de la interfaz Streamlit y de la persistencia por convocatoria."""
import sqlite3

import pytest
from streamlit.testing.v1 import AppTest

from core import build_candidate_record, analyze_cv_for_tables, init_db, load_convocatoria, save_convocatoria_results

VACANTE = {
    "codigo_vacante": "CONV-TEST-01",
    "nombre_puesto": "Coordinador Médico",
    "meses_experiencia": 24,
    "requiere_tarjeta": True,
    "skills_obligatorias": ["Resolución 3100", "Auditoría médica"],
}
CV_TOP = ("Médico Coordinador con Tarjeta Profesional vigente. 4 años de experiencia. "
          "Resolución 3100 y Auditoría médica. 20% menos glosas, 500 pacientes, $ 1000. Enero 2022.")
CV_MEDIO = "Médico con Tarjeta Profesional. 3 años de experiencia. Resolución 3100. 100 pacientes."
CV_SIN_TARJETA = "Médico general con 5 años de experiencia."


@pytest.fixture
def db(tmp_path, monkeypatch):
    path = str(tmp_path / "test.db")
    monkeypatch.setenv("DATABASE_PATH", path)
    init_db(path)
    return path


def _candidatos():
    return [build_candidate_record("Ana Top", CV_TOP, VACANTE),
            build_candidate_record("Luis Medio", CV_MEDIO, VACANTE),
            build_candidate_record("Pedro Sin Tarjeta", CV_SIN_TARJETA, VACANTE)]


def test_tablas_reflejan_puntajes_del_candidato():
    cand = build_candidate_record("Ana Top", CV_TOP, VACANTE)
    tablas = analyze_cv_for_tables(CV_TOP, VACANTE, cand)
    assert list(tablas["modulos"]["Puntaje"]) == [cand["m2"], cand["m3"], cand["m4"], cand["m5"]]
    assert set(tablas) == {"datos_generales", "skills", "modulos", "logros"}


def test_meses_no_se_confunden_con_anios():
    cand = build_candidate_record("X", "Tarjeta profesional. 30 meses de experiencia. Trabajó 1 año en urgencias.", VACANTE)
    assert cand["meses_detectados"] == 30


def test_guardado_sin_duplicados_y_recarga(db):
    cands = _candidatos()
    for c in cands:
        c.update(score_entrevista=80.0, score_final=round(c["score_cv"] * .6 + 32, 1))
    save_convocatoria_results(VACANTE, cands, "Ana Top", "ana@test.com", db)
    save_convocatoria_results(VACANTE, cands, "Ana Top", "ana@test.com", db)
    with sqlite3.connect(db) as conn:
        assert conn.execute("SELECT COUNT(*) FROM candidatos_historico").fetchone()[0] == 3
    data = load_convocatoria("CONV-TEST-01", db)
    assert data["vacante"]["nombre_puesto"] == "Coordinador Médico"
    assert {c["nombre"] for c in data["evaluaciones"]} == {"Ana Top", "Luis Medio", "Pedro Sin Tarjeta"}


def test_app_flujo_completo(db):
    at = AppTest.from_file("app.py", default_timeout=60)
    at.run()
    assert not at.exception
    at.text_input[0].input("admin")
    at.text_input[1].input("AdminPass2026!")
    at.button[0].click().run()
    assert not at.exception
    assert at.session_state["user"]["username"] == "admin"

    # Inyecta una criba ya realizada y recorre todas las pestañas
    at.session_state["vacante_activa"] = VACANTE
    at.session_state["evaluaciones"] = _candidatos()
    at.run()
    assert not at.exception

    # Registra una entrevista completa para el primer aprobado
    at.session_state["entrevistas"] = {"Ana Top": 80.0}
    at.session_state["entrevistas_detalle"] = {"Ana Top": [
        {"pregunta_num": i, "tema": "T", "enunciado": "P", "puntaje": 20 if i < 5 else 0, "observacion": "ok"}
        for i in range(1, 6)]}
    at.run()
    assert not at.exception

    # Cierra la convocatoria dos veces: no debe duplicar registros
    for _ in range(2):
        next(b for b in at.button if b.key == "btn_cierre").click().run()
        assert not at.exception
    with sqlite3.connect(db) as conn:
        rows = conn.execute("SELECT nombre, seleccionado FROM candidatos_historico WHERE vacante_codigo = ?",
                            ("CONV-TEST-01",)).fetchall()
    assert len(rows) == 3
    assert ("Ana Top", 1) in rows
    assert at.session_state["archivada"] is True
