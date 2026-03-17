# src/soma_app/infra/soma_api_client.py
from __future__ import annotations

import logging
import os
import time
from dataclasses import dataclass
from typing import Any, Dict, Optional, Union

import requests

try:
    from soma_app.infra.trace import step, log_kv  # type: ignore
except Exception:  # pragma: no cover
    step = None  # type: ignore
    log_kv = None  # type: ignore

log = logging.getLogger("soma_app.infra.soma_api")

JsonType = Union[Dict[str, Any], list, str, int, float, bool, None]


def _mask_secret(s: str, keep: int = 3) -> str:
    s = (s or "").strip()
    if not s:
        return ""
    if len(s) <= keep:
        return "*" * len(s)
    return ("*" * (len(s) - keep)) + s[-keep:]


@dataclass
class SomaApiError(RuntimeError):
    status_code: int
    method: str
    url: str
    message: str
    response_text: str = ""

    def __str__(self) -> str:
        base = f"{self.status_code} {self.method} {self.url} | {self.message}"
        if self.response_text:
            return f"{base} | body={self.response_text}"
        return base


class SomaApiClient:
    """
    Cliente HTTP para Soma API.

    Modos de autenticação suportados:
      A) Token de sessão PRONTO (recomendado no teu caso):
         - define SOMA_SESSION_TOKEN no .env
         - o cliente usa esse token no header X-Session-Token
         - NÃO chama /v1/sessions

      B) Abrir sessão via /v1/sessions (se permitido no teu ambiente):
         - login/password -> token
    """

    def __init__(
        self,
        *,
        base_url: str,
        login: str = "",
        password: str = "",
        session_token: Optional[str] = None,
        close_on_exit: bool = False,
        timeout_seconds: int = 40,
        max_retries: int = 2,
        backoff_seconds: float = 0.9,
    ) -> None:
        self.base_url = (base_url or "").rstrip("/")
        self.login = (login or "").strip()
        self.password = (password or "").strip()

        self.timeout_seconds = int(timeout_seconds or 40)
        self.max_retries = int(max_retries or 0)
        self.backoff_seconds = float(backoff_seconds or 0.0)

        self._http = requests.Session()

        self.session_token: Optional[str] = (session_token or "").strip() or None
        self._opened_by_client: bool = False  # só True quando nós chamamos /v1/sessions
        self._close_on_exit: bool = bool(close_on_exit)

    # -------------------------
    # sessão
    # -------------------------
    def use_session_token(self, token: str, *, close_on_exit: bool = False) -> None:
        t = (token or "").strip()
        if not t:
            raise ValueError("session token vazio")
        self.session_token = t
        self._opened_by_client = False
        self._close_on_exit = bool(close_on_exit)
        log.info("SOMA API token externo configurado | token=%s", _mask_secret(t))

    def open_session(self) -> str:
        # ✅ Se já tens token, NÃO chama /v1/sessions
        if self.session_token:
            log.info("SOMA API: a usar token já configurado | token=%s", _mask_secret(self.session_token))
            return self.session_token

        if not self.login or not self.password:
            raise RuntimeError(
                "Sem SOMA_SESSION_TOKEN e sem credenciais. "
                "Define SOMA_SESSION_TOKEN no .env (preferido), ou SOMA_API_LOGIN/SOMA_API_PASSWORD."
            )

        url = f"{self.base_url}/v1/sessions"
        payload = {"login": self.login, "password": self.password}

        data = self._request_json("POST", url, json_body=payload, needs_session=False, allow_write=False)

        token = self._extract_token(data)
        if not token and isinstance(data, str) and data.strip():
            token = data.strip()

        if not token:
            raise RuntimeError("Não consegui obter token de sessão no retorno do /v1/sessions.")

        self.session_token = token
        self._opened_by_client = True

        log.info(
            "SOMA API sessão aberta | base_url=%s | login=%s | token=%s",
            self.base_url,
            _mask_secret(self.login),
            _mask_secret(token),
        )
        return token

    def close_session(self) -> None:
        if not self.session_token:
            return

        # ✅ Só fecha se (a) nós abrimos a sessão, ou (b) o user pediu close_on_exit
        if (not self._opened_by_client) and (not self._close_on_exit):
            return

        url = f"{self.base_url}/v1/sessions/{self.session_token}"
        try:
            self._request_json("DELETE", url, json_body=None, needs_session=False, allow_write=False)
        except Exception:
            log.exception("Falha ao fechar sessão na Soma API.")
        finally:
            self.session_token = None
            self._opened_by_client = False

    # -------------------------
    # requests públicos
    # -------------------------
    def get_json(self, path: str, *, params: Optional[Dict[str, Any]] = None) -> JsonType:
        return self._request_json(
            "GET",
            f"{self.base_url}{path}",
            params=params,
            json_body=None,
            needs_session=True,
            allow_write=False,
            extra_headers=None,
        )

    def post_json(
        self,
        path: str,
        *,
        payload: Dict[str, Any],
        allow_write: bool = True,
        idempotency_key: Optional[str] = None,
    ) -> JsonType:
        extra = {}
        if idempotency_key:
            extra["Idempotency-Key"] = idempotency_key

        return self._request_json(
            "POST",
            f"{self.base_url}{path}",
            params=None,
            json_body=payload,
            needs_session=True,
            allow_write=allow_write,
            extra_headers=extra or None,
        )

    # -------------------------
    # internals
    # -------------------------
    def _headers(self, *, needs_session: bool, allow_write: bool) -> Dict[str, str]:
        h: Dict[str, str] = {"Accept": "application/json"}
        if needs_session:
            if not self.session_token:
                raise RuntimeError("Sessão não iniciada. Define SOMA_SESSION_TOKEN ou chama open_session().")
            h["X-Session-Token"] = self.session_token
        if allow_write:
            h["X-Allow-Write"] = "true"
        return h

    def _request_json(
        self,
        method: str,
        url: str,
        *,
        params: Optional[Dict[str, Any]] = None,
        json_body: Optional[Dict[str, Any]] = None,
        needs_session: bool,
        allow_write: bool,
        extra_headers: Optional[Dict[str, str]] = None,
    ) -> JsonType:
        headers = self._headers(needs_session=needs_session, allow_write=allow_write)
        if extra_headers:
            headers.update({k: v for k, v in extra_headers.items() if v})

        # correlação (opcional)
        run_id = (os.getenv("RUN_ID") or "").strip()
        row_no = (os.getenv("ROW_NUMBER") or "").strip()
        if run_id:
            headers["X-Run-Id"] = run_id
        if row_no:
            headers["X-Row-Number"] = row_no

        ctx = None
        if callable(step):
            ctx = step(log, "api.request", method=method, url=url)
            ctx.__enter__()

        t0 = time.time()
        last_exc: Optional[Exception] = None

        try:
            for attempt in range(self.max_retries + 1):
                try:
                    resp = self._http.request(
                        method=method,
                        url=url,
                        params=params,
                        json=json_body,
                        headers=headers,
                        timeout=self.timeout_seconds,
                    )
                    elapsed_ms = int((time.time() - t0) * 1000)

                    if callable(log_kv):
                        log_kv(
                            log,
                            logging.INFO,
                            "API response",
                            method=method,
                            url=url,
                            status=resp.status_code,
                            elapsed_ms=elapsed_ms,
                            params=params or {},
                        )
                    else:
                        log.info("API %s %s -> %s (%sms)", method, url, resp.status_code, elapsed_ms)

                    if 200 <= resp.status_code < 300:
                        txt = (resp.text or "").strip()
                        if not txt:
                            return None
                        try:
                            return resp.json()
                        except Exception:
                            return txt

                    if resp.status_code in {429, 500, 502, 503, 504} and attempt < self.max_retries:
                        time.sleep(max(0.2, self.backoff_seconds * (2**attempt)))
                        continue

                    body = (resp.text or "").strip()
                    raise SomaApiError(
                        status_code=resp.status_code,
                        method=method,
                        url=url,
                        message="Request falhou",
                        response_text=body[:800],
                    )

                except SomaApiError:
                    raise
                except Exception as e:
                    last_exc = e
                    if attempt < self.max_retries:
                        time.sleep(max(0.2, self.backoff_seconds * (2**attempt)))
                        continue
                    raise

            if last_exc:
                raise last_exc
            raise RuntimeError("Erro desconhecido no request API.")
        finally:
            if ctx is not None:
                try:
                    ctx.__exit__(None, None, None)
                except Exception:
                    pass

    @staticmethod
    def _extract_token(data: JsonType) -> Optional[str]:
        if isinstance(data, dict):
            for k in (
                "token",
                "session_token",
                "sessionToken",
                "x_session_token",
                "xSessionToken",
                "X-Session-Token",
            ):
                v = data.get(k)
                if isinstance(v, str) and v.strip():
                    return v.strip()
        return None