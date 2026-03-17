# src/soma_app/automation/api/iv_api.py
from __future__ import annotations

import logging
import os
import re
from typing import Any, Dict, Iterable, Optional

from soma_app.domain.models import ContaOrdemRow, TipoMovimento
from soma_app.infra.soma_api_client import SomaApiClient, JsonType

log = logging.getLogger("soma_app.automation.api.iv")


def _as_str(v: Any) -> str:
    return "" if v is None else str(v).strip()


def _extract_doc_id(obj: Any) -> Optional[str]:
    if isinstance(obj, str):
        m = re.search(r"\b\d{3,}\b", obj)
        return m.group(0) if m else (obj.strip() or None)

    if isinstance(obj, dict):
        for k in ("doc_id", "docId", "doc", "documento", "numero_documento", "numero", "id", "ID"):
            v = obj.get(k)
            if isinstance(v, (int, float)):
                vv = str(int(v))
                if vv:
                    return vv
            if isinstance(v, str) and v.strip():
                m = re.search(r"\b\d{3,}\b", v)
                return m.group(0) if m else v.strip()

        for v in obj.values():
            got = _extract_doc_id(v)
            if got:
                return got

    if isinstance(obj, list):
        for it in obj:
            got = _extract_doc_id(it)
            if got:
                return got

    return None


def _iter_items(data: JsonType) -> Iterable[Dict[str, Any]]:
    if isinstance(data, list):
        for it in data:
            if isinstance(it, dict):
                yield it
        return

    if isinstance(data, dict):
        for k in ("items", "data", "results", "rows"):
            v = data.get(k)
            if isinstance(v, list):
                for it in v:
                    if isinstance(it, dict):
                        yield it
                return

        for v in data.values():
            if isinstance(v, list):
                for it in v:
                    if isinstance(it, dict):
                        yield it
                return


def _idempotency(kind: str, row_number: int) -> str:
    run_id = (os.getenv("RUN_ID") or "run").strip()
    return f"{run_id}:{row_number}:{kind}"


class EntradasSaidasApi:
    def __init__(self, api: SomaApiClient):
        self.api = api

    def create_and_get_doc_id(self, row: ContaOrdemRow) -> str:
        kind = "saida" if row.tipo == TipoMovimento.SAIDA else "entrada"
        key = _idempotency(f"entradas-saidas:{kind}", int(getattr(row, "row_number", 0) or 0))

        # ✅ dry_run é obrigatório (erro 422)
        payload: Dict[str, Any] = {
            "dry_run": False,
            "kind": kind,
            "descricao": _as_str(getattr(row, "descricao_soma", "")),
            "descricao_soma": _as_str(getattr(row, "descricao_soma", "")),
            "plano_conta": _as_str(getattr(row, "plano_conta", "")),
            "centro_custo": _as_str(getattr(row, "centro_custo", "")),
            "valor": _as_str(getattr(row, "importancia", "")),
            "data": _as_str(getattr(row, "data_mov", "")),
            "forma_pagamento": _as_str(getattr(row, "forma_pagamento", "")),
            "caixa": _as_str(getattr(row, "caixa", "")),
            "obs": _as_str(getattr(row, "descricao_soma", "")),
        }

        data = self.api.post_json("/v1/ivv/entradas-saidas", payload=payload, allow_write=True, idempotency_key=key)
        doc = _extract_doc_id(data)
        if not doc:
            raise RuntimeError(f"API: não consegui extrair DOC do create entradas-saidas. retorno={str(data)[:400]}")
        return doc

    def recover_doc_id(self, row: ContaOrdemRow) -> str:
        kind = "saida" if row.tipo == TipoMovimento.SAIDA else "entrada"

        params = {
            "kind": kind,
            "pesquisa": _as_str(getattr(row, "descricao_soma", "")),
            "filtro": "descricao",
            "use_period": True,
            "start_date": _as_str(getattr(row, "data_mov", "")),
            "end_date": _as_str(getattr(row, "data_mov", "")),
            "limit": 5,
            "offset": 0,
        }

        data = self.api.get_json("/v1/ivv/entradas-saidas", params=params)
        for it in _iter_items(data):
            doc = _extract_doc_id(it)
            if doc:
                return doc

        raise RuntimeError("API: recover_doc_id não encontrou resultado na listagem (pesquisa + período).")

    def fetch_dados_doc(self, doc_id: str) -> str:
        doc_id = (doc_id or "").strip()
        if not doc_id:
            raise ValueError("doc_id vazio")

        # tenta by-id (se existir no backend)
        try:
            data = self.api.get_json(f"/v1/ivv/entradas-saidas/{doc_id}", params=None)
            if isinstance(data, dict):
                for k in ("dados_doc", "dadosDoc", "dados", "detalhes", "info", "observacao", "obs", "descricao"):
                    v = data.get(k)
                    if isinstance(v, str) and v.strip():
                        return v.strip()
        except Exception:
            pass

        # fallback: pesquisar
        for filtro in ("id", "descricao"):
            params = {"pesquisa": doc_id, "filtro": filtro, "limit": 5, "offset": 0}
            data2 = self.api.get_json("/v1/ivv/entradas-saidas", params=params)
            for it in _iter_items(data2):
                for k in ("dados_doc", "dadosDoc", "dados", "detalhes", "info", "observacao", "obs", "descricao"):
                    v = it.get(k)
                    if isinstance(v, str) and v.strip():
                        return v.strip()

        raise RuntimeError("API: não encontrei DADOS DOC para este documento.")


class TransferenciasApi:
    def __init__(self, api: SomaApiClient):
        self.api = api

    def run(self, row: ContaOrdemRow) -> str:
        key = _idempotency("transferencia", int(getattr(row, "row_number", 0) or 0))

        payload: Dict[str, Any] = {
            "dry_run": False,  # ✅ por consistência (se também for obrigatório aqui)
            "caixa_saida": _as_str(getattr(row, "caixa_saida", "")),
            "caixa_entrada": _as_str(getattr(row, "caixa", "")),
            "valor": _as_str(getattr(row, "importancia", "")),
            "data": _as_str(getattr(row, "data_mov", "")),
            "descricao": _as_str(getattr(row, "descricao_soma", "")),
        }

        data = self.api.post_json(
            "/v1/ivv/transferencias-caixas",
            payload=payload,
            allow_write=True,
            idempotency_key=key,
        )
        doc = _extract_doc_id(data)
        return doc or "Transferido"