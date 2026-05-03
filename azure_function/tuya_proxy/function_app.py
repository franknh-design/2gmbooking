"""
tuya_proxy / function_app.py
============================
Azure Function HTTP-trigger som fungerer som sikker proxy mellom
2GM Booking (JavaScript) og Tuya Smart Lock API.

CLIENT_ID og CLIENT_SECRET lagres som Azure Function App Settings
(miljøvariabler) — aldri i kildekoden.

Endepunkter:
  POST /api/tuya/create_pin    — Opprett temporær PIN på en lås
  POST /api/tuya/delete_pin    — Slett en PIN fra en lås
  GET  /api/tuya/list_pins     — List aktive PIN-koder på en lås
  GET  /api/tuya/health        — Enkel liveness-sjekk

Krav (requirements.txt):
  azure-functions
  requests
  pycryptodome

CORS:
  Tillatte origins konfigureres i Azure Portal under
  Function App → API → CORS. Legg til:
    https://booking.2gm.no
    https://franknh-design.github.io
"""

import azure.functions as func
import json
import os
import logging

from .tuya_client import TuyaLockClient, TuyaError

app = func.FunctionApp(http_auth_level=func.AuthLevel.FUNCTION)

# ============================================================
# Hjelpefunksjoner
# ============================================================

def _json_response(data: dict, status: int = 200) -> func.HttpResponse:
    return func.HttpResponse(
        json.dumps(data, ensure_ascii=False),
        status_code=status,
        mimetype="application/json",
        headers={"Access-Control-Allow-Origin": "*"},
    )

def _error(msg: str, status: int = 400) -> func.HttpResponse:
    return _json_response({"success": False, "error": msg}, status)

def _get_client() -> TuyaLockClient:
    client_id = os.environ.get("TUYA_CLIENT_ID")
    client_secret = os.environ.get("TUYA_CLIENT_SECRET")
    base_url = os.environ.get("TUYA_BASE_URL", "https://openapi.tuyaeu.com")
    if not client_id or not client_secret:
        raise TuyaError("TUYA_CLIENT_ID / TUYA_CLIENT_SECRET mangler i miljøvariabler")
    return TuyaLockClient(client_id, client_secret, base_url)


# ============================================================
# POST /api/tuya/create_pin
# Body (JSON):
#   device_id   str   — Tuya device_id for låsen
#   pin         str   — 6-sifret PIN
#   name        str   — Navn/beskrivelse (f.eks. gjestens navn)
#   valid_from  int?  — Unix timestamp start (standard: nå)
#   valid_to    int?  — Unix timestamp slutt
#   days        int?  — Antall dager (brukes hvis valid_to mangler, standard: 30)
# ============================================================
@app.route(route="tuya/create_pin", methods=["POST", "OPTIONS"])
def create_pin(req: func.HttpRequest) -> func.HttpResponse:
    if req.method == "OPTIONS":
        return func.HttpResponse(status_code=204, headers={"Access-Control-Allow-Origin": "*", "Access-Control-Allow-Headers": "Content-Type,x-functions-key"})
    try:
        body = req.get_json()
    except Exception:
        return _error("Ugyldig JSON i request body")

    device_id = (body.get("device_id") or "").strip()
    pin = (body.get("pin") or "").strip()
    name = (body.get("name") or "Gjest").strip()
    valid_from = body.get("valid_from")
    valid_to = body.get("valid_to")
    days = int(body.get("days", 30))

    if not device_id:
        return _error("device_id er påkrevd")
    if not pin or len(pin) != 6 or not pin.isdigit():
        return _error("pin må være nøyaktig 6 siffer")

    try:
        client = _get_client()
        result = client.create_password(
            device_id=device_id,
            pin=pin,
            name=name,
            valid_from=valid_from,
            valid_to=valid_to,
            days=days,
        )
        return _json_response({"success": True, "result": result})
    except TuyaError as e:
        logging.error(f"[Tuya] create_pin feilet: {e}")
        return _error(str(e), 502)
    except Exception as e:
        logging.exception("[Tuya] Uventet feil i create_pin")
        return _error("Intern serverfeil", 500)


# ============================================================
# POST /api/tuya/delete_pin
# Body (JSON):
#   device_id   str   — Tuya device_id
#   password_id int   — ID på koden som skal slettes
# ============================================================
@app.route(route="tuya/delete_pin", methods=["POST", "OPTIONS"])
def delete_pin(req: func.HttpRequest) -> func.HttpResponse:
    if req.method == "OPTIONS":
        return func.HttpResponse(status_code=204, headers={"Access-Control-Allow-Origin": "*", "Access-Control-Allow-Headers": "Content-Type,x-functions-key"})
    try:
        body = req.get_json()
    except Exception:
        return _error("Ugyldig JSON i request body")

    device_id = (body.get("device_id") or "").strip()
    password_id = body.get("password_id")

    if not device_id:
        return _error("device_id er påkrevd")
    if not password_id:
        return _error("password_id er påkrevd")

    try:
        client = _get_client()
        client.delete_password(device_id=device_id, password_id=int(password_id))
        return _json_response({"success": True})
    except TuyaError as e:
        logging.error(f"[Tuya] delete_pin feilet: {e}")
        return _error(str(e), 502)
    except Exception as e:
        logging.exception("[Tuya] Uventet feil i delete_pin")
        return _error("Intern serverfeil", 500)


# ============================================================
# GET /api/tuya/list_pins?device_id=xxx
# ============================================================
@app.route(route="tuya/list_pins", methods=["GET", "OPTIONS"])
def list_pins(req: func.HttpRequest) -> func.HttpResponse:
    if req.method == "OPTIONS":
        return func.HttpResponse(status_code=204, headers={"Access-Control-Allow-Origin": "*", "Access-Control-Allow-Headers": "Content-Type,x-functions-key"})

    device_id = (req.params.get("device_id") or "").strip()
    if not device_id:
        return _error("device_id er påkrevd som query-parameter")

    try:
        client = _get_client()
        passwords = client.list_passwords(device_id=device_id, valid_only=True)
        return _json_response({"success": True, "result": passwords})
    except TuyaError as e:
        logging.error(f"[Tuya] list_pins feilet: {e}")
        return _error(str(e), 502)
    except Exception as e:
        logging.exception("[Tuya] Uventet feil i list_pins")
        return _error("Intern serverfeil", 500)


# ============================================================
# GET /api/tuya/health
# ============================================================
@app.route(route="tuya/health", methods=["GET"])
def health(req: func.HttpRequest) -> func.HttpResponse:
    configured = bool(os.environ.get("TUYA_CLIENT_ID") and os.environ.get("TUYA_CLIENT_SECRET"))
    return _json_response({"status": "ok", "configured": configured})
