from __future__ import annotations

import io
import json
import os
from datetime import datetime, timedelta
from threading import Lock

import gspread
import pandas as pd
from flask import Flask, jsonify, render_template, request, send_file
from google.oauth2.service_account import Credentials

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
]

SPREADSHEET_ID = os.getenv("SPREADSHEET_ID", "")
BASE_SHEET_NAME = os.getenv("BASE_SHEET_NAME", "BASE_COLABORADORES")
RESPONSES_SHEET_NAME = os.getenv("RESPONSES_SHEET_NAME", "RESPOSTAS_FERIAS")
ADMIN_TOKEN = os.getenv("ADMIN_TOKEN", "RH_BRA_2026")
CACHE_MINUTES = int(os.getenv("CACHE_MINUTES", "10"))
PORT = int(os.getenv("PORT", "10000"))

app = Flask(__name__)
_base_lock = Lock()
_write_lock = Lock()
_base_cache: dict[str, dict] = {}
_base_cache_expires_at: datetime | None = None

BASE_FIELDS = ["MATRICULA", "CPF", "NOME", "UNIDADE", "MES_FERIAS"]
RESPONSE_HEADERS = [
    "DATA_HORA",
    "MATRICULA",
    "NOME",
    "UNIDADE",
    "MES_FERIAS",
    "TIPO_FERIAS",
    "EMAIL",
    "TELEFONE",
    "OBSERVACOES",
    "CIENCIA",
]


def normalize(value) -> str:
    return str("" if value is None else value).strip()


def only_digits(value) -> str:
    return "".join(ch for ch in str("" if value is None else value) if ch.isdigit())


def normalize_header(value) -> str:
    text = normalize(value).upper()
    replacements = {
        "Á": "A", "À": "A", "Â": "A", "Ã": "A",
        "É": "E", "Ê": "E",
        "Í": "I",
        "Ó": "O", "Ô": "O", "Õ": "O",
        "Ú": "U",
        "Ç": "C",
    }
    for src, dst in replacements.items():
        text = text.replace(src, dst)
    text = text.replace(" ", "_")
    return text


def email_valid(email: str) -> bool:
    return "@" in email and "." in email.split("@")[-1] and " " not in email


def phone_valid(phone: str) -> bool:
    digits = only_digits(phone)
    return len(digits) in (10, 11)


def get_google_credentials() -> Credentials:
    raw_json = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
    if not raw_json:
        raise RuntimeError(
            "A variável GOOGLE_SERVICE_ACCOUNT_JSON não foi configurada no Render."
        )

    try:
        info = json.loads(raw_json)
    except json.JSONDecodeError as exc:
        raise RuntimeError(
            "A variável GOOGLE_SERVICE_ACCOUNT_JSON não contém um JSON válido."
        ) from exc

    return Credentials.from_service_account_info(info, scopes=SCOPES)


def get_client() -> gspread.Client:
    return gspread.authorize(get_google_credentials())


def get_spreadsheet():
    if not SPREADSHEET_ID:
        raise RuntimeError("A variável SPREADSHEET_ID não foi configurada no Render.")
    return get_client().open_by_key(SPREADSHEET_ID)


def get_worksheet(name: str):
    try:
        return get_spreadsheet().worksheet(name)
    except gspread.WorksheetNotFound as exc:
        raise RuntimeError(f'A aba "{name}" não foi encontrada na planilha.') from exc


def ensure_responses_sheet() -> None:
    ss = get_spreadsheet()
    try:
        ws = ss.worksheet(RESPONSES_SHEET_NAME)
    except gspread.WorksheetNotFound:
        ws = ss.add_worksheet(title=RESPONSES_SHEET_NAME, rows=1000, cols=len(RESPONSE_HEADERS))

    current = [normalize_header(v) for v in ws.row_values(1)[: len(RESPONSE_HEADERS)]]
    expected = [normalize_header(v) for v in RESPONSE_HEADERS]
    if current != expected:
        ws.update("A1:J1", [RESPONSE_HEADERS])


def load_base(force: bool = False) -> dict[str, dict]:
    global _base_cache, _base_cache_expires_at

    with _base_lock:
        now = datetime.now()
        if not force and _base_cache and _base_cache_expires_at and now < _base_cache_expires_at:
            return _base_cache

        ws = get_worksheet(BASE_SHEET_NAME)
        values = ws.get_all_values()

        if not values:
            raise RuntimeError('A aba "BASE_COLABORADORES" está vazia.')

        raw_headers = values[0]
        headers = [normalize_header(h) for h in raw_headers]

        missing = [field for field in BASE_FIELDS if field not in headers]
        if missing:
            raise RuntimeError(
                "Cabeçalhos obrigatórios não encontrados na aba BASE_COLABORADORES: "
                + ", ".join(missing)
            )

        idx = {header: headers.index(header) for header in BASE_FIELDS}

        result: dict[str, dict] = {}
        matriculas_duplicadas: list[str] = []
        cpfs_duplicados: list[str] = []

        for row in values[1:]:
            matricula = normalize(row[idx["MATRICULA"]] if idx["MATRICULA"] < len(row) else "")
            cpf = only_digits(row[idx["CPF"]] if idx["CPF"] < len(row) else "")

            if not matricula:
                continue

            collaborator = {
                "matricula": matricula,
                "cpf": cpf,
                "nome": normalize(row[idx["NOME"]] if idx["NOME"] < len(row) else ""),
                "unidade": normalize(row[idx["UNIDADE"]] if idx["UNIDADE"] < len(row) else ""),
                "mes_ferias": normalize(row[idx["MES_FERIAS"]] if idx["MES_FERIAS"] < len(row) else ""),
            }

            chave_matricula = f"matricula:{matricula}"
            if chave_matricula in result:
                matriculas_duplicadas.append(matricula)
            else:
                result[chave_matricula] = collaborator

            if cpf:
                chave_cpf = f"cpf:{cpf}"
                if chave_cpf in result:
                    cpfs_duplicados.append(cpf)
                else:
                    result[chave_cpf] = collaborator

        if matriculas_duplicadas:
            raise RuntimeError(
                "Existem matrículas duplicadas na aba BASE_COLABORADORES: "
                + ", ".join(matriculas_duplicadas[:10])
            )

        if cpfs_duplicados:
            raise RuntimeError(
                "Existem CPFs duplicados na aba BASE_COLABORADORES: "
                + ", ".join(cpfs_duplicados[:10])
            )

        _base_cache = result
        _base_cache_expires_at = now + timedelta(minutes=CACHE_MINUTES)
        return _base_cache


def find_collaborator(base: dict[str, dict], identificador: str) -> dict | None:
    identificador_normalizado = normalize(identificador)
    identificador_numerico = only_digits(identificador)

    if not identificador_normalizado:
        return None

    collaborator = base.get(f"matricula:{identificador_normalizado}")
    if collaborator:
        return collaborator

    if identificador_numerico:
        collaborator = base.get(f"matricula:{identificador_numerico}")
        if collaborator:
            return collaborator

        collaborator = base.get(f"cpf:{identificador_numerico}")
        if collaborator:
            return collaborator

    return None


def answer_exists(matricula: str) -> bool:
    ws = get_worksheet(RESPONSES_SHEET_NAME)
    values = ws.col_values(2)
    if len(values) <= 1:
        return False

    target = only_digits(matricula) or normalize(matricula)
    for value in values[1:]:
        current = only_digits(value) or normalize(value)
        if current == target:
            return True
    return False


def append_answer(row: list[str]) -> None:
    ws = get_worksheet(RESPONSES_SHEET_NAME)
    ws.append_row(row, value_input_option="USER_ENTERED")


@app.get("/")
def index():
    return render_template("index.html")


@app.get("/health")
def health():
    return jsonify({"ok": True})


@app.post("/api/consultar")
def api_consultar():
    try:
        payload = request.get_json(silent=True) or {}
        identificador = normalize(payload.get("identificador"))

        if not identificador:
            return jsonify({
                "sucesso": False,
                "mensagem": "Informe a matrícula ou o CPF."
            }), 400

        base = load_base()
        collaborator = find_collaborator(base, identificador)

        if not collaborator:
            return jsonify({
                "sucesso": False,
                "mensagem": "Matrícula ou CPF não localizado. Verifique os números digitados."
            }), 404

        if answer_exists(collaborator["matricula"]):
            return jsonify({
                "sucesso": False,
                "mensagem": "Já existe uma resposta registrada para este colaborador.",
                "bloqueado": True,
            }), 409

        return jsonify({"sucesso": True, "colaborador": collaborator})
    except Exception as exc:
        return jsonify({"sucesso": False, "mensagem": f"Erro ao consultar: {exc}"}), 500


@app.post("/api/enviar")
def api_enviar():
    payload = request.get_json(silent=True) or {}
    matricula = normalize(payload.get("matricula"))
    tipo_ferias = normalize(payload.get("tipo_ferias"))
    email = normalize(payload.get("email"))
    telefone = normalize(payload.get("telefone"))
    observacoes = normalize(payload.get("observacoes"))
    ciencia = bool(payload.get("ciencia"))

    if not matricula:
        return jsonify({"sucesso": False, "mensagem": "Informe a matrícula."}), 400
    if not tipo_ferias:
        return jsonify({"sucesso": False, "mensagem": "Selecione o tipo de férias."}), 400
    if not email:
        return jsonify({"sucesso": False, "mensagem": "Informe o e-mail."}), 400
    if not email_valid(email):
        return jsonify({"sucesso": False, "mensagem": "Informe um e-mail válido."}), 400
    if not telefone:
        return jsonify({"sucesso": False, "mensagem": "Informe o telefone com DDD."}), 400
    if not phone_valid(telefone):
        return jsonify({"sucesso": False, "mensagem": "Informe um telefone válido com DDD."}), 400
    if not ciencia:
        return jsonify({"sucesso": False, "mensagem": "É obrigatório marcar a declaração de ciência."}), 400

    try:
        base = load_base()
        collaborator = find_collaborator(base, matricula)
        if not collaborator:
            return jsonify({"sucesso": False, "mensagem": "Matrícula não localizada."}), 404

        with _write_lock:
            if answer_exists(collaborator["matricula"]):
                return jsonify({"sucesso": False, "mensagem": "Já existe uma resposta registrada para esta matrícula."}), 409

            append_answer([
                datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
                collaborator["matricula"],
                collaborator["nome"],
                collaborator["unidade"],
                collaborator["mes_ferias"],
                tipo_ferias,
                email,
                telefone,
                observacoes,
                "SIM",
            ])

        return jsonify({"sucesso": True, "mensagem": "Resposta enviada com sucesso."})
    except Exception as exc:
        return jsonify({"sucesso": False, "mensagem": f"Erro ao enviar: {exc}"}), 500


@app.post("/api/recarregar-base")
def api_recarregar_base():
    token = request.headers.get("X-Admin-Token", "")
    if token != ADMIN_TOKEN:
        return jsonify({"sucesso": False, "mensagem": "Acesso não autorizado."}), 403
    try:
        load_base(force=True)
        return jsonify({"sucesso": True, "mensagem": "Base recarregada com sucesso."})
    except Exception as exc:
        return jsonify({"sucesso": False, "mensagem": f"Erro ao recarregar base: {exc}"}), 500


@app.get("/admin/exportar")
def admin_exportar():
    token = request.args.get("token", "")
    if token != ADMIN_TOKEN:
        return "Acesso não autorizado.", 403

    ws = get_worksheet(RESPONSES_SHEET_NAME)
    values = ws.get_all_values()
    if not values:
        values = [RESPONSE_HEADERS]

    headers = values[0]
    rows = values[1:]
    df = pd.DataFrame(rows, columns=headers)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=RESPONSES_SHEET_NAME)
    output.seek(0)
    return send_file(
        output,
        as_attachment=True,
        download_name="respostas_ferias.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    ensure_responses_sheet()
    load_base(force=True)
    app.run(host="0.0.0.0", port=PORT, debug=False)
