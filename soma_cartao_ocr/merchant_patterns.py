#!/usr/bin/env python3
"""Aprendizado e validação de padrões de comerciantes."""

import json
from collections import defaultdict
from pathlib import Path
from typing import Any, TYPE_CHECKING

import numpy as np

if TYPE_CHECKING:
    from main import Movement


class MerchantPatternLearner:
    """Aprende padrões de comerciantes do histórico de movimentos validados."""

    def __init__(self, history_path: Path | str | None = None):
        """
        Inicializa o aprendiz de padrões.

        Args:
            history_path: Caminho para arquivo JSON com padrões históricos
        """
        self.history_path = Path(history_path) if history_path else Path("merchant_patterns.json")
        self.patterns: dict[str, dict[str, Any]] = {}
        self.load_patterns()

    def learn_from_movement(self, movement: "Movement"):
        """
        Aprende de um movimento validado.

        Args:
            movement: Movimento com status = "VÁLIDO"
        """
        from main import parse_money

        if movement.status != "VÁLIDO":
            return  # Só aprender de movimentos válidos

        desc_upper = movement.descricao.upper().strip()

        if not desc_upper:
            return

        if desc_upper not in self.patterns:
            self.patterns[desc_upper] = {
                "count": 0,
                "typical_amounts": [],
                "amounts_history": [],
                "dates": [],
                "confidence_scores": [],
                "countries": [],
                "first_seen": None,
                "last_seen": None,
            }

        pattern = self.patterns[desc_upper]
        pattern["count"] += 1

        # Registrar montante
        debit = parse_money(movement.debito_eur)
        credit = parse_money(movement.credito_eur)
        amount = debit if debit else credit

        if amount is not None:
            pattern["amounts_history"].append(amount)
            # Manter últimas 50 transações
            pattern["amounts_history"] = pattern["amounts_history"][-50:]

        # Registrar data
        if movement.data_movimento:
            pattern["dates"].append(movement.data_movimento)

        # Registrar confiança
        pattern["confidence_scores"].append(movement.confidence)

        # Registrar país
        if movement.pais:
            pattern["countries"].append(movement.pais)

        # Atualizar last_seen
        pattern["last_seen"] = movement.data_movimento

        if pattern["first_seen"] is None:
            pattern["first_seen"] = movement.data_movimento

    def get_expected_amount_range(
        self,
        description: str,
        confidence_level: float = 0.95
    ) -> tuple[float, float]:
        """
        Retorna intervalo típico de valores para um comerciante.

        Args:
            description: Descrição do comerciante
            confidence_level: Nível de confiança (0.95 = intervalo de 95%)

        Returns:
            (valor_mínimo, valor_máximo)
        """
        pattern = self.patterns.get(description.upper().strip())

        if not pattern or not pattern["amounts_history"]:
            return (0.0, 10000.0)  # Default para comerciante desconhecido

        amounts = pattern["amounts_history"]
        mean = np.mean(amounts)
        std = np.std(amounts)

        if std == 0:  # Todos os valores são iguais
            return (mean * 0.8, mean * 1.2)

        # Z-score para intervalo de confiança
        z_score = 1.96 if confidence_level == 0.95 else 2.576

        return (
            max(0, mean - z_score * std),
            mean + z_score * std
        )

    def is_expected_amount(
        self,
        description: str,
        amount: float,
        strict: bool = False
    ) -> bool:
        """
        Verifica se o montante é típico para o comerciante.

        Args:
            description: Descrição do comerciante
            amount: Montante a validar
            strict: Se True, usa intervalo mais estreito

        Returns:
            True se o montante é esperado
        """
        confidence = 0.99 if strict else 0.95
        low, high = self.get_expected_amount_range(description, confidence)

        return low <= amount <= high

    def get_merchant_confidence(self, description: str) -> float:
        """
        Retorna nível de confiança/conhecimento sobre um comerciante.

        Args:
            description: Descrição do comerciante

        Returns:
            Score 0.0-1.0
        """
        pattern = self.patterns.get(description.upper().strip())

        if not pattern:
            return 0.0

        # Confiança baseada em:
        # 1. Número de ocorrências (máximo 50)
        count_score = min(pattern["count"] / 50, 1.0)

        # 2. Consistência de montantes (baixo desvio padrão)
        if len(pattern["amounts_history"]) > 1:
            amounts = pattern["amounts_history"]
            mean = np.mean(amounts)
            std = np.std(amounts)
            cv = std / mean if mean > 0 else 0  # Coeficiente de variação
            consistency_score = max(0, 1.0 - cv)  # Mais estável = score maior
        else:
            consistency_score = 0.5

        # 3. Confiança média do OCR quando visto
        ocr_confidence = np.mean(pattern["confidence_scores"]) if pattern["confidence_scores"] else 0.5

        # Score combinado (ponderado)
        score = (count_score * 0.4 + consistency_score * 0.35 + ocr_confidence * 0.25)

        return min(max(score, 0.0), 1.0)

    def get_merchant_info(self, description: str) -> dict[str, Any]:
        """
        Retorna informações completas sobre um comerciante.

        Args:
            description: Descrição do comerciante

        Returns:
            Dicionário com informações
        """
        pattern = self.patterns.get(description.upper().strip())

        if not pattern:
            return {
                "known": False,
                "description": description,
            }

        amounts = pattern["amounts_history"]
        return {
            "known": True,
            "description": description,
            "occurrences": pattern["count"],
            "typical_amount_range": self.get_expected_amount_range(description),
            "average_amount": np.mean(amounts) if amounts else None,
            "amount_std": np.std(amounts) if amounts else None,
            "confidence_score": self.get_merchant_confidence(description),
            "avg_ocr_confidence": np.mean(pattern["confidence_scores"]) if pattern["confidence_scores"] else None,
            "countries": list(set(pattern["countries"])) if pattern["countries"] else [],
            "first_seen": pattern["first_seen"],
            "last_seen": pattern["last_seen"],
        }

    def find_similar_merchants(
        self,
        description: str,
        max_results: int = 5
    ) -> list[tuple[str, float]]:
        """
        Encontra comerciantes similares usando string matching.

        Args:
            description: Descrição a buscar
            max_results: Máximo de resultados

        Returns:
            Lista de (descrição, score_similaridade)
        """
        from difflib import SequenceMatcher

        desc_upper = description.upper().strip()
        similarities = []

        for known_desc in self.patterns.keys():
            ratio = SequenceMatcher(None, desc_upper, known_desc).ratio()
            if ratio > 0.5:  # Pelo menos 50% similar
                similarities.append((known_desc, ratio))

        # Ordenar por similaridade (maior primeiro)
        similarities.sort(key=lambda x: x[1], reverse=True)

        return similarities[:max_results]

    def save_patterns(self):
        """Persiste padrões em JSON."""
        # Converter numpy types para JSON-serializáveis
        serializable = {}

        for desc, pattern in self.patterns.items():
            serializable[desc] = {
                "count": int(pattern["count"]),
                "typical_amounts": [float(x) for x in pattern["typical_amounts"]],
                "amounts_history": [float(x) for x in pattern["amounts_history"]],
                "dates": pattern["dates"],
                "confidence_scores": [float(x) for x in pattern["confidence_scores"]],
                "countries": pattern["countries"],
                "first_seen": pattern["first_seen"],
                "last_seen": pattern["last_seen"],
            }

        self.history_path.parent.mkdir(parents=True, exist_ok=True)

        with open(self.history_path, "w", encoding="utf-8") as f:
            json.dump(serializable, f, indent=2, ensure_ascii=False)

    def load_patterns(self):
        """Carrega padrões históricos de arquivo."""
        if not self.history_path.exists():
            self.patterns = {}
            return

        try:
            with open(self.history_path, "r", encoding="utf-8") as f:
                loaded = json.load(f)
                self.patterns = loaded
        except (json.JSONDecodeError, IOError):
            self.patterns = {}

    def get_statistics(self) -> dict[str, Any]:
        """
        Retorna estatísticas sobre o corpus de padrões.

        Returns:
            Dicionário com estatísticas
        """
        if not self.patterns:
            return {"merchants_tracked": 0}

        total_transactions = sum(p["count"] for p in self.patterns.values())
        all_amounts = []
        all_confidences = []

        for pattern in self.patterns.values():
            all_amounts.extend(pattern["amounts_history"])
            all_confidences.extend(pattern["confidence_scores"])

        return {
            "merchants_tracked": len(self.patterns),
            "total_transactions": total_transactions,
            "avg_transactions_per_merchant": total_transactions / len(self.patterns),
            "avg_transaction_amount": np.mean(all_amounts) if all_amounts else None,
            "avg_ocr_confidence": np.mean(all_confidences) if all_confidences else None,
            "most_frequent_merchants": sorted(
                [(d, p["count"]) for d, p in self.patterns.items()],
                key=lambda x: x[1],
                reverse=True
            )[:10],
        }
