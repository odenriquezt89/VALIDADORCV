import json
import secrets
from datetime import datetime
from typing import List, Optional
from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Security, Request, status
from fastapi.security import APIKeyHeader
from pydantic import BaseModel, Field
from slowapi import Limiter, _rate_limit_exceeded_handler
from slowapi.util import get_remote_address
from slowapi.errors import RateLimitExceeded
from core import get_db_path, get_db_connection, init_db, extract_text_from_pdf_bytes, evaluate_candidate

DB_PATH = get_db_path()
init_db(DB_PATH)

FALLBACK_TEST_KEYS = {
    "key_live_evaluador_alpha_2026": {"cliente": "Evaluador Alpha", "activo": True},
    "key_live_evaluador_inactivo": {"cliente": "Cuenta Vencida", "activo": False}
}

API_KEY_HEADER = APIKeyHeader(name="X-API-KEY", auto_error=False)
limiter = Limiter(key_func=get_remote_address)

def verify_api_key(api_key: str = Security(API_KEY_HEADER)) -> dict:
    if not api_key:
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="Falta la cabecera 'X-API-KEY'.")

    conn = get_db_connection(DB_PATH)
    row = conn.execute("SELECT username, activo FROM usuarios WHERE api_key = ?", (api_key,)).fetchone()
    conn.close()
    if row:
        if row["activo"] != 1:
            raise HTTPException(status_code=status.HTTP_403_FORBIDDEN, detail="Suscripción suspendida o inactiva.")
        return {"cliente": row["username"], "activo": True}

    for k, val in FALLBACK_TEST_KEYS.items():
        if secrets.compare_digest(api_key, k):
            if not val["activo"]:
                raise HTTPException(status_code=status.HTTP_403_FORBIDDEN, detail="Suscripción suspendida o inactiva.")
            return val

    raise HTTPException(status_code=status.HTTP_403_FORBIDDEN, detail="API Key inválida o no registrada.")

app = FastAPI(title="TalentValidator API", version="1.0.0")
app.state.limiter = limiter
app.add_exception_handler(RateLimitExceeded, _rate_limit_exceeded_handler)

class VacancyRequirements(BaseModel):
    codigo_vacante: str = Field(..., examples=["VAC-2026-01"])
    nombre_puesto: str = Field(..., examples=["Coordinador de Auditoría"])
    meses_experiencia: int = Field(default=24, ge=0)
    requiere_tarjeta: bool = Field(default=True)
    skills_obligatorias: List[str] = Field(default_factory=list)

class ModuleScores(BaseModel):
    m2_trayectoria: float
    m3_skills: float
    m4_logros: float
    m5_formato: float

class EvaluationResponse(BaseModel):
    candidato: str
    codigo_vacante: str
    fase_1_knockout: str
    motivo_descarte: Optional[str] = None
    score_cv: float
    desglose_modulos: ModuleScores
    fortalezas: List[str]
    debilidades: List[str]
    timestamp: str

class BatchEvaluationResponse(BaseModel):
    cliente_autenticado: str
    total_procesados: int
    aprobados_knockout: int
    descartados_knockout: int
    candidatos_ordenados: List[EvaluationResponse]

def build_response(name: str, text: str, vac: VacancyRequirements) -> EvaluationResponse:
    res = evaluate_candidate(text, vac.model_dump())
    return EvaluationResponse(
        candidato=name,
        codigo_vacante=vac.codigo_vacante,
        fase_1_knockout="APROBADO" if res["knockout"] else "DESCARTADO",
        motivo_descarte=res["motivo_descarte"],
        score_cv=res["score_cv"],
        desglose_modulos=ModuleScores(
            m2_trayectoria=res["modulos"]["m2"],
            m3_skills=res["modulos"]["m3"],
            m4_logros=res["modulos"]["m4"],
            m5_formato=res["modulos"]["m5"]
        ),
        fortalezas=res["fortalezas"],
        debilidades=res["debilidades"],
        timestamp=datetime.now().isoformat()
    )

@app.post("/evaluate/cv", response_model=EvaluationResponse)
@limiter.limit("30/minute")
async def evaluate_single_cv(
    request: Request,
    cv_file: UploadFile = File(...),
    vacante_data: str = Form(...),
    client_info: dict = Security(verify_api_key)
):
    if not cv_file.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Solo se admiten documentos en formato PDF.")
    try:
        vac = VacancyRequirements(**json.loads(vacante_data))
    except Exception as e:
        raise HTTPException(status_code=422, detail=f"JSON de vacante inválido: {str(e)}")

    raw_text = extract_text_from_pdf_bytes(await cv_file.read())
    name = cv_file.filename.replace(".pdf", "").replace("_", " ").title()
    return build_response(name, raw_text, vac)

@app.post("/evaluate/batch", response_model=BatchEvaluationResponse)
@limiter.limit("10/minute")
async def evaluate_batch_cv(
    request: Request,
    cv_files: List[UploadFile] = File(...),
    vacante_data: str = Form(...),
    client_info: dict = Security(verify_api_key)
):
    if len(cv_files) > 10:
        raise HTTPException(status_code=400, detail="El procesamiento admite un máximo de 10 archivos.")
    try:
        vac = VacancyRequirements(**json.loads(vacante_data))
    except Exception as e:
        raise HTTPException(status_code=422, detail=f"JSON de vacante inválido: {str(e)}")

    evals = []
    for f in cv_files:
        if f.filename.lower().endswith(".pdf"):
            txt = extract_text_from_pdf_bytes(await f.read())
            name = f.filename.replace(".pdf", "").replace("_", " ").title()
            evals.append(build_response(name, txt, vac))

    evals.sort(key=lambda x: (x.fase_1_knockout == "APROBADO", x.score_cv), reverse=True)
    aprob = sum(1 for e in evals if e.fase_1_knockout == "APROBADO")
    return BatchEvaluationResponse(
        cliente_autenticado=client_info["cliente"],
        total_procesados=len(evals),
        aprobados_knockout=aprob,
        descartados_knockout=len(evals) - aprob,
        candidatos_ordenados=evals
    )
