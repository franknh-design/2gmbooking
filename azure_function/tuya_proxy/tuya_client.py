"""
tuya_client.py
==============
Tuya Smart Lock API-klient.
Brukes av Azure Function-proxyen.
Kan også brukes direkte (se tuya_lock.py for CLI-eksempler).
"""

import hashlib
import hmac
import time
import uuid
import json
import requests
from Crypto.Cipher import AES
from Crypto.Util.Padding import pad


class TuyaError(Exception):
    """Feil fra Tuya API eller kryptering."""
    pass


class TuyaLockClient:
    """
    Klient for Tuya Smart Lock temporære PIN-koder.

    Args:
        client_id:     Tuya IoT Platform Client ID
        client_secret: Tuya IoT Platform Client Secret
        base_url:      API base URL (standard: https://openapi.tuyaeu.com)
    """

    def __init__(self, client_id: str, client_secret: str, base_url: str = "https://openapi.tuyaeu.com"):
        self.client_id = client_id
        self.client_secret = client_secret
        self.base_url = base_url.rstrip("/")

    # --------------------------------------------------------
    # SIGNERING OG TOKEN
    # --------------------------------------------------------

    def _sign(self, t: str, nonce: str, access_token: str, method: str, path: str, body_str: str = "") -> str:
        content_hash = hashlib.sha256(body_str.encode("utf-8")).hexdigest()
        str_to_sign = "\n".join([method, content_hash, "", path])
        message = self.client_id + access_token + t + nonce + str_to_sign
        return hmac.new(
            self.client_secret.encode("utf-8"),
            message.encode("utf-8"),
            hashlib.sha256,
        ).hexdigest().upper()

    def _headers(self, method: str, path: str, body_str: str = "", token: str = "") -> dict:
        t = str(int(time.time() * 1000))
        nonce = uuid.uuid4().hex
        sign = self._sign(t, nonce, token, method, path, body_str)
        h = {
            "client_id": self.client_id,
            "sign": sign,
            "t": t,
            "sign_method": "HMAC-SHA256",
            "nonce": nonce,
        }
        if token:
            h["access_token"] = token
        return h

    def get_token(self) -> str:
        path = "/v1.0/token?grant_type=1"
        r = requests.get(
            self.base_url + path,
            headers=self._headers("GET", path),
            timeout=10,
        )
        data = r.json()
        if not data.get("success"):
            raise TuyaError(f"Token feilet: {data.get('msg', data)}")
        return data["result"]["access_token"]

    # --------------------------------------------------------
    # HTTP-HJELPERE
    # --------------------------------------------------------

    def _get(self, token: str, path: str) -> dict:
        r = requests.get(
            self.base_url + path,
            headers=self._headers("GET", path, token=token),
            timeout=10,
        )
        return r.json()

    def _post(self, token: str, path: str, body: dict) -> dict:
        body_str = json.dumps(body, separators=(",", ":"))
        headers = self._headers("POST", path, body_str=body_str, token=token)
        headers["Content-Type"] = "application/json"
        r = requests.post(
            self.base_url + path,
            headers=headers,
            data=body_str,
            timeout=10,
        )
        return r.json()

    def _delete(self, token: str, path: str) -> dict:
        r = requests.delete(
            self.base_url + path,
            headers=self._headers("DELETE", path, token=token),
            timeout=10,
        )
        return r.json()

    # --------------------------------------------------------
    # KRYPTERING
    # --------------------------------------------------------

    def _decrypt_ticket_key(self, ticket_key_hex: str) -> bytes:
        key = self.client_secret[:32].encode("utf-8").ljust(32, b"\0")
        encrypted = bytes.fromhex(ticket_key_hex)
        cipher = AES.new(key, AES.MODE_ECB)
        decrypted = b""
        for i in range(0, len(encrypted), 16):
            decrypted += cipher.decrypt(encrypted[i : i + 16])
        return decrypted[:16]

    def _encrypt_pin(self, pin: str, decrypted_key: bytes) -> str:
        if len(pin) != 6 or not pin.isdigit():
            raise TuyaError("PIN må være nøyaktig 6 siffer")
        key = decrypted_key[:16]
        padded = pad(pin.encode("utf-8"), AES.block_size, style="pkcs7")
        cipher = AES.new(key, AES.MODE_ECB)
        return cipher.encrypt(padded).hex().upper()

    def _get_ticket(self, token: str, device_id: str) -> tuple[str, str]:
        resp = self._post(token, f"/v1.0/devices/{device_id}/door-lock/password-ticket", {})
        if not resp.get("success"):
            raise TuyaError(f"Ticket feilet: {resp.get('msg', resp)}")
        return resp["result"]["ticket_id"], resp["result"]["ticket_key"]

    # --------------------------------------------------------
    # OFFENTLIGE METODER
    # --------------------------------------------------------

    def create_password(
        self,
        device_id: str,
        pin: str,
        name: str,
        valid_from: int | None = None,
        valid_to: int | None = None,
        days: int = 30,
    ) -> dict:
        """
        Oppretter en temporær PIN-kode på låsen.

        Returns:
            dict med password_id, name, pin, valid_from, valid_to
        """
        now = int(time.time())
        if valid_from is None:
            valid_from = now
        if valid_to is None:
            valid_to = now + (days * 24 * 60 * 60)

        token = self.get_token()
        ticket_id, ticket_key_hex = self._get_ticket(token, device_id)
        decrypted_key = self._decrypt_ticket_key(ticket_key_hex)
        encrypted_pin = self._encrypt_pin(pin, decrypted_key)

        body = {
            "name": name,
            "password": encrypted_pin,
            "password_type": "ticket",
            "ticket_id": ticket_id,
            "effective_time": valid_from,
            "invalid_time": valid_to,
            "time_zone": "Europe/Oslo",
            "type": 0,
        }

        result = self._post(token, f"/v1.0/devices/{device_id}/door-lock/temp-password", body)
        if not result.get("success"):
            raise TuyaError(f"Opprett kode feilet: {result.get('msg', result)}")

        return {
            "password_id": result["result"]["id"],
            "name": name,
            "pin": pin,
            "valid_from": valid_from,
            "valid_to": valid_to,
        }

    def list_passwords(self, device_id: str, valid_only: bool = True) -> list:
        """
        Returnerer liste med temporære PIN-koder for enheten.
        """
        token = self.get_token()
        path = f"/v1.0/devices/{device_id}/door-lock/temp-passwords"
        if valid_only:
            path += "?valid=true"
        result = self._get(token, path)
        if not result.get("success"):
            raise TuyaError(f"List koder feilet: {result.get('msg', result)}")
        return result["result"]

    def delete_password(self, device_id: str, password_id: int) -> bool:
        """
        Sletter en temporær PIN-kode fra låsen.
        """
        token = self.get_token()
        result = self._delete(token, f"/v1.0/devices/{device_id}/door-lock/temp-passwords/{password_id}")
        if not result.get("success"):
            raise TuyaError(f"Slett kode feilet: {result.get('msg', result)}")
        return True
