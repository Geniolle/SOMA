from __future__ import annotations

import logging
from typing import Any

from soma_app.domain.models import ContaOrdemRow

logger = logging.getLogger(__name__)


def _safe_err(e: Exception) -> str:
    s = str(e).strip().replace("\n", " ")
    return s[:180] if s else type(e).__name__


def _is_auth_error(e: Exception) -> bool:
    status_code = getattr(e, "status_code", None)
    if status_code in {401, 403}:
        return True
    msg = str(e).lower()
    return ("sessão inválida" in msg) or ("sessao invalida" in msg) or ("unauthorized" in msg)


class EntradasSaidasAdapter:
    """
    Usa a API como primária e, em caso de falha de autenticação, desativa a API
    para o resto do run e passa a usar o fallback (Selenium) se disponível.
    """

    def __init__(self, primary: Any, fallback: Any | None = None):
        self.primary = primary
        self.fallback = fallback
        self._api_enabled = True

    def _disable_api_if_needed(self, e: Exception, *, op: str) -> None:
        if self.fallback is None:
            return
        if self._api_enabled and _is_auth_error(e):
            self._api_enabled = False
            logger.warning("API desativada para Entradas/Saídas após falha de autenticação (%s).", op)

    def create_and_get_doc_id(self, row: ContaOrdemRow) -> str:
        if (not self._api_enabled) and self.fallback is not None:
            return self.fallback.create_and_get_doc_id(row)
        try:
            return self.primary.create_and_get_doc_id(row)
        except Exception as e:
            self._disable_api_if_needed(e, op="create")
            if self.fallback is None:
                raise
            logger.warning("API falhou (create Entradas/Saídas). Vou tentar Selenium. err=%s", _safe_err(e))
            return self.fallback.create_and_get_doc_id(row)

    def recover_doc_id(self, row: ContaOrdemRow) -> str | None:
        if (not self._api_enabled) and self.fallback is not None:
            return self.fallback.recover_doc_id(row)
        try:
            return self.primary.recover_doc_id(row)
        except Exception as e:
            self._disable_api_if_needed(e, op="recover")
            if self.fallback is None:
                raise
            logger.warning("API falhou (recover DOC). Vou tentar Selenium. err=%s", _safe_err(e))
            return self.fallback.recover_doc_id(row)

    def precheck_duplicate(self, row: ContaOrdemRow) -> str | None:
        if (not self._api_enabled) and self.fallback is not None:
            return self.fallback.precheck_duplicate(row)
        try:
            return self.primary.precheck_duplicate(row)
        except Exception as e:
            self._disable_api_if_needed(e, op="precheck_duplicate")
            if self.fallback is None:
                raise
            logger.warning("API falhou (precheck duplicate). Vou tentar Selenium. err=%s", _safe_err(e))
            return self.fallback.precheck_duplicate(row)

    def fetch_dados_doc(self, doc_id: str) -> str:
        if (not self._api_enabled) and self.fallback is not None and hasattr(self.fallback, "fetch_dados_doc"):
            return self.fallback.fetch_dados_doc(doc_id)
        return self.primary.fetch_dados_doc(doc_id)


class TransferenciasAdapter:
    def __init__(self, primary: Any, fallback: Any | None = None):
        self.primary = primary
        self.fallback = fallback
        self._api_enabled = True

    def _disable_api_if_needed(self, e: Exception) -> None:
        if self.fallback is None:
            return
        if self._api_enabled and _is_auth_error(e):
            self._api_enabled = False
            logger.warning("API desativada para Transferências após falha de autenticação.")

    def run(self, row: ContaOrdemRow) -> str:
        if (not self._api_enabled) and self.fallback is not None:
            return self.fallback.run(row)
        try:
            return self.primary.run(row)
        except Exception as e:
            self._disable_api_if_needed(e)
            if self.fallback is None:
                raise
            logger.warning("API falhou (Transferência). Vou tentar Selenium. err=%s", _safe_err(e))
            return self.fallback.run(row)
