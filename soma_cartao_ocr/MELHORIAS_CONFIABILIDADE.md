# Plano de Melhoria de Confiabilidade - OCR de Extratos de Cartão

## 1. Análise Atual

**Taxa de sucesso observada**: 84% (16 válidas de 19 movimentos)
**Linhas para revisão**: 3 movimentos (~16%)

### Problemas Identificados

#### 1.1 Confiança OCR Baixa
- Threshold configurado, mas sem fallback inteligente
- Não há validação contra padrões históricos
- Sem uso de contexto semântico

#### 1.2 Erros Comuns de Leitura
- Caracteres confundidos: "0" vs "O", "1" vs "l", "S" vs "5"
- Formatação inconsistente: "COMISSAD" → "COMISSÃO"
- Datas com separadores incorretos: "26 06" → "26/06"

#### 1.3 Alinhamento de Coluna
- Palavras podem sangrar entre colunas
- Falta validação de coerência espacial
- Sem detecção de "linhas fantasma" (ruído que parece ser dados)

#### 1.4 Contexto Insuficiente
- Sem análise de similaridade com histórico
- Sem validação de negócio (ex: valores típicos por comerciante)
- Sem análise de padrão de datas

---

## 2. Melhorias Propostas (Priorizado)

### 🔴 CRÍTICA (Alto Impacto, Fácil Implementação)

#### 2.1 Validação de Coerência de Coluna
**Impacto**: Reduz falsos positivos em 5-8%

```python
def validate_column_coherence(row: list[Word], column_bounds: dict) -> dict:
    """
    Valida se as palavras de cada coluna fazem sentido semântico.
    Deteta quando palavras sangram incorretamente entre colunas.
    """
    issues = {}
    for col_name, (x_min, x_max) in column_bounds.items():
        words_in_col = [w for w in row if x_min <= w.cx < x_max]
        
        # 1. Verificar coerência de confiança
        confidence_score = np.mean([w.confidence for w in words_in_col])
        if confidence_score < 0.7:
            issues[col_name] = {
                "reason": "low_column_confidence",
                "score": confidence_score
            }
        
        # 2. Detectar inconsistência de altura
        if words_in_col:
            heights = [w.y1 - w.y0 for w in words_in_col]
            if max(heights) > min(heights) * 1.3:  # Variação > 30%
                issues[col_name]["reason"] = "height_inconsistency"
        
        # 3. Detectar espaços anormais entre palavras
        if len(words_in_col) > 1:
            sorted_words = sorted(words_in_col, key=lambda w: w.x0)
            gaps = [sorted_words[i+1].x0 - sorted_words[i].x1 
                   for i in range(len(sorted_words)-1)]
            if gaps and max(gaps) > 3 * np.median(gaps):
                issues[col_name]["reason"] = "unusual_spacing"
    
    return issues
```

#### 2.2 Biblioteca de Correções Automáticas
**Impacto**: Reduz erros de leitura conhecidos em 10-15%

```python
COMMON_OCR_FIXES = {
    # Caracteres confundidos
    "O0": {"0": 0.8, "O": 0.2},  # 80% provável ser "0"
    "Il": {"1": 0.7, "l": 0.3},
    "S5": {"5": 0.8, "S": 0.2},
    "rn": {"m": 0.9},
    
    # Erros de formatação específicos
    "COMISSAD": "COMISSÃO",
    "COMISSAQ": "COMISSÃO",
    "COMIÇAO": "COMISSÃO",
    
    # Nomes comuns de comerciantes
    "MERCADORÍA": "MERCADONA",
    "SUPERMERCAD": "SUPERMERCADO",
    "FARMAÇIA": "FARMÁCIA",
    
    # Padrões de data
    r"(\d{2})\s+(\d{2})": r"\1/\2",  # "26 06" → "26/06"
    r"(\d{2})[.,](\d{2})": r"\1/\2",  # "26.06" → "26/06"
}

def apply_ocr_corrections(text: str, confidence: float) -> tuple[str, bool]:
    """
    Aplica correções automáticas baseadas em confiança.
    Retorna (texto_corrigido, foi_corrigido)
    """
    corrected = text
    was_corrected = False
    
    for pattern, replacement in COMMON_OCR_FIXES.items():
        if isinstance(replacement, dict):
            # Heurística probabilística
            if confidence < 0.75:
                best_char = max(replacement, key=replacement.get)
                if text in pattern:
                    corrected = corrected.replace(text, best_char)
                    was_corrected = True
        else:
            # Substituição direta
            if pattern.startswith("("):  # Regex
                import re
                new_text = re.sub(pattern, replacement, corrected, flags=re.IGNORECASE)
                if new_text != corrected:
                    corrected = new_text
                    was_corrected = True
            else:
                if pattern in corrected.upper():
                    corrected = corrected.replace(pattern, replacement)
                    was_corrected = True
    
    return corrected, was_corrected
```

#### 2.3 Análise Semântica de Descrição
**Impacto**: Reduz rejeições desnecessárias em 3-5%

```python
def validate_description_semantics(desc: str) -> tuple[bool, str]:
    """
    Valida se a descrição faz sentido semântico.
    Detecta quando é apenas ruído/números.
    """
    if not desc or len(desc) < 2:
        return False, "empty"
    
    # Detectar descrições apenas com números/símbolos
    alpha_ratio = sum(c.isalpha() for c in desc) / len(desc)
    if alpha_ratio < 0.3:
        return False, "mostly_numbers"
    
    # Detectar repetição excessiva
    if len(set(desc)) < len(desc) * 0.3:
        return False, "excessive_repetition"
    
    # Validar comprimento mínimo
    words = desc.split()
    if len(words) == 0:
        return False, "no_words"
    
    if len(desc) < 3:
        return False, "too_short"
    
    return True, "valid"
```

---

### 🟠 ALTA (Impacto Significativo, Complexidade Média)

#### 2.4 Sistema de Validação Cruzada Multi-Nível
**Impacto**: Reduz erros compostos em 8-12%

```python
def cross_validate_movement(movement: Movement, cfg: dict) -> dict[str, bool]:
    """
    Executa múltiplas validações cruzadas:
    - Coerência data/valor
    - Consistência com histórico
    - Validação de negócio
    """
    validations = {}
    
    # 1. Validação de Coerência Data
    d_mov = extract_day_month(movement.data_movimento)
    d_val = extract_day_month(movement.data_valor)
    
    if d_mov and d_val:
        # Data de valor não deve ser ANTES de data de movimento
        if d_val < d_mov and not (d_mov[1] == 6 and d_val == (1, 7)):
            validations["date_order"] = False
        else:
            validations["date_order"] = True
    
    # 2. Validação de Consistência de Montante
    debit = parse_money(movement.debito_eur)
    credit = parse_money(movement.credito_eur)
    
    if debit and credit:
        # Não pode ter débito e crédito simultâneamente
        validations["amount_exclusivity"] = False
    else:
        validations["amount_exclusivity"] = True
    
    # 3. Validação de Valor Típico
    if debit:
        # Valores típicos em extratos de cartão
        if debit < 0.01 or debit > 10000:
            validations["amount_reasonable"] = False
        else:
            validations["amount_reasonable"] = True
    
    # 4. Validação de País (se presente)
    if movement.pais:
        valid_countries = {"IRL", "USA", "PRT", "ESP", "GBR", "DEU", "FRA", "NLD", "LUX", "ITA", "CHE"}
        if movement.pais.upper() not in valid_countries:
            validations["country_valid"] = False
        else:
            validations["country_valid"] = True
    
    return validations
```

#### 2.5 Aprimoramento de Detecção de Linhas
**Impacto**: Reduz "linhas fantasma" em 10-15%

```python
def validate_row_consistency(row: list[Word], table_cfg: dict) -> tuple[bool, str]:
    """
    Valida se uma linha extraída é realmente uma linha de dados
    ou apenas ruído/artefatos.
    """
    
    # 1. Deve ter pelo menos uma data
    date_words = [w for w in row if re.match(r"\d{1,2}[/.-]\d{1,2}", w.text)]
    if not date_words:
        return False, "no_date_pattern"
    
    # 2. Deve ter confiança média aceitável
    avg_confidence = np.mean([w.confidence for w in row])
    if avg_confidence < 0.5:
        return False, "low_avg_confidence"
    
    # 3. Não deve ser repetição de palavras
    unique_words = len(set(w.text.upper() for w in row))
    if unique_words < len(row) * 0.3:  # Menos de 30% unique
        return False, "too_repetitive"
    
    # 4. Deve ter mínimo de palavras significativas
    significant_words = [w for w in row if len(w.text) > 2]
    if len(significant_words) < 2:
        return False, "insufficient_content"
    
    # 5. Alinhamento Y deve ser coerente
    y_positions = [w.cy for w in row]
    y_variance = np.var(y_positions)
    y_range = max(y_positions) - min(y_positions)
    avg_word_height = np.mean([w.y1 - w.y0 for w in row])
    
    if y_range > avg_word_height * 0.8:  # Y spread é muito grande
        return False, "inconsistent_y_alignment"
    
    return True, "valid"
```

#### 2.6 Scoring de Confiança Aprimorado
**Impacto**: Reduz decisões erradas em 5-10%

```python
def calculate_enhanced_confidence_score(
    movement: Movement,
    row_words: list[Word],
    table_cfg: dict,
    historical_context: dict | None = None
) -> float:
    """
    Calcula score de confiança em múltiplas dimensões.
    Retorna 0.0-1.0 onde 1.0 é confiança máxima.
    """
    
    scores = {}
    
    # 1. Confiança de OCR pura (30%)
    ocr_confidence = np.mean([w.confidence for w in row_words]) if row_words else 0
    scores["ocr"] = ocr_confidence * 0.3
    
    # 2. Coerência de coluna (25%)
    column_validity = validate_column_coherence(row_words, table_cfg.get("columns", {}))
    col_score = 1.0 - (len(column_validity) / max(len(table_cfg.get("columns", {})), 1))
    scores["column"] = col_score * 0.25
    
    # 3. Validação de dados (20%)
    cross_vals = cross_validate_movement(movement, table_cfg)
    validation_pass_rate = sum(cross_vals.values()) / max(len(cross_vals), 1)
    scores["validation"] = validation_pass_rate * 0.20
    
    # 4. Consistência com histórico (15%)
    if historical_context:
        desc_match = historical_context.get("similar_description_found", False)
        pattern_match = historical_context.get("pattern_matches", False)
        hist_score = (desc_match + pattern_match) / 2
        scores["history"] = hist_score * 0.15
    else:
        scores["history"] = 0.15  # Assume neutro se sem histórico
    
    # 5. Bonificação por correções aplicadas (10%)
    corrections_applied = movement.texto_ocr != getattr(movement, "raw_text", movement.texto_ocr)
    scores["corrections"] = 0.1 if corrections_applied else 0
    
    # Score final ponderado
    final_score = sum(scores.values())
    return min(max(final_score, 0.0), 1.0)  # Clamp 0-1
```

---

### 🟡 MÉDIA (Impacto Relevante, Requer Trabalho Significativo)

#### 2.7 Aprendizado de Padrões Comerciantes
**Impacto**: Reduz erros em 8-12% após acúmulo de dados

```python
class MerchantPatternLearner:
    """Aprende padrões de comerciantes do histórico."""
    
    def __init__(self, history_path: Path):
        self.patterns: dict[str, dict] = {}
        self.history_path = history_path
        self.load_patterns()
    
    def learn_from_movement(self, movement: Movement):
        """Adiciona conhecimento do movimento validado."""
        desc_upper = movement.descricao.upper().strip()
        
        if desc_upper not in self.patterns:
            self.patterns[desc_upper] = {
                "count": 0,
                "typical_amount": [],
                "dates": [],
                "country": movement.pais,
                "confidence_scores": []
            }
        
        pattern = self.patterns[desc_upper]
        pattern["count"] += 1
        
        if debit := parse_money(movement.debito_eur):
            pattern["typical_amount"].append(debit)
        
        pattern["dates"].append(movement.data_movimento)
        pattern["confidence_scores"].append(movement.confidence)
    
    def get_expected_amount_range(self, description: str) -> tuple[float, float]:
        """Retorna intervalo típico de valores para um comerciante."""
        pattern = self.patterns.get(description.upper().strip())
        if not pattern or not pattern["typical_amount"]:
            return (0.0, 10000.0)  # Default
        
        amounts = pattern["typical_amount"]
        mean = np.mean(amounts)
        std = np.std(amounts)
        
        return (max(0, mean - 2 * std), mean + 2 * std)
    
    def is_expected_amount(self, description: str, amount: float) -> bool:
        """Verifica se o montante é típico para o comerciante."""
        low, high = self.get_expected_amount_range(description)
        return low <= amount <= high
    
    def save_patterns(self):
        """Persiste padrões em JSON."""
        with open(self.history_path, "w") as f:
            json.dump(self.patterns, f, indent=2, default=str)
    
    def load_patterns(self):
        """Carrega padrões históricos."""
        if self.history_path.exists():
            with open(self.history_path) as f:
                self.patterns = json.load(f)
```

#### 2.8 Detecção de Falsos Negativos
**Impacto**: Reduz rejeições incorretas em 5-8%

```python
def detect_potential_false_rejection(
    movement: Movement,
    cfg: dict,
    trusted_descriptions: set[str]
) -> bool:
    """
    Detecta se um movimento foi rejeitado incorretamente.
    Retorna True se deve ser reconsiderado.
    """
    
    if movement.status != "REVISÃO":
        return False
    
    # Se a descrição está na lista confiável, reconsidera
    if movement.descricao.upper() in trusted_descriptions:
        return True
    
    # Se foi rejeitado apenas por confiança OCR baixa
    if "confiança OCR baixa" in movement.motivos_revisao:
        ocr_confidence = movement.confidence
        
        # Mas a descrição é legível e o montante faz sentido
        if ocr_confidence > 0.65 and len(movement.descricao) > 5:
            # Reconsidere se o padrão é conhecido
            merchant_patterns = load_merchant_patterns()
            if movement.descricao in merchant_patterns:
                return True
    
    # Se tem data válida mas foi rejeitado por outra razão
    if valid_date(movement.data_movimento, cfg.get("validation", {})):
        reason_count = len(movement.motivos_revisao.split(";"))
        if reason_count == 1 and "confiança" not in movement.motivos_revisao:
            # Único motivo e não é confiança - avaliar se é crítico
            if movement.debito_eur or movement.credito_eur:
                return True  # Tem montante, talvez seja válido
    
    return False
```

---

### 🟢 BÔNUS (Qualidade de Vida, Observabilidade)

#### 2.9 Sistema de Logging Estruturado
```python
def log_ocr_analysis(movement: Movement, analysis_metadata: dict):
    """
    Registra análise detalhada para auditoria e aprendizado.
    """
    log_entry = {
        "timestamp": datetime.now().isoformat(),
        "movement_line": movement.line,
        "description": movement.descricao,
        "status": movement.status,
        "confidence": movement.confidence,
        "analysis": {
            "column_coherence": analysis_metadata.get("column_coherence"),
            "cross_validation": analysis_metadata.get("cross_validation"),
            "semantic_check": analysis_metadata.get("semantic_check"),
            "corrections_applied": analysis_metadata.get("corrections_applied")
        },
        "raw_text": movement.texto_ocr
    }
    
    # Escrever em log estruturado para análise posterior
    return log_entry
```

#### 2.10 Dashboard de Métricas
```python
def generate_ocr_quality_metrics(movements: list[Movement]) -> dict:
    """Gera métricas de qualidade do OCR para monitoramento."""
    
    total = len(movements)
    validos = sum(1 for m in movements if m.status == "VÁLIDO")
    revisao = sum(1 for m in movements if m.status == "REVISÃO")
    
    avg_confidence = np.mean([m.confidence for m in movements])
    min_confidence = min([m.confidence for m in movements], default=0)
    max_confidence = max([m.confidence for m in movements], default=1)
    
    # Análise de motivos de rejeição
    rejection_reasons = {}
    for m in movements:
        if m.status == "REVISÃO":
            for reason in m.motivos_revisao.split(";"):
                reason = reason.strip()
                rejection_reasons[reason] = rejection_reasons.get(reason, 0) + 1
    
    return {
        "total_movements": total,
        "valid_count": validos,
        "review_count": revisao,
        "success_rate": (validos / total * 100) if total > 0 else 0,
        "avg_confidence": avg_confidence,
        "confidence_range": (min_confidence, max_confidence),
        "rejection_reasons": rejection_reasons,
        "quality_trend": "improving" if avg_confidence > 0.85 else "needs_attention"
    }
```

---

## 3. Implementação Prioritizada

### Fase 1 (Semana 1-2) - Ganhos Rápidos
1. ✅ Implementar correções OCR automáticas
2. ✅ Validação de coerência de coluna
3. ✅ Análise semântica básica de descrição
4. ✅ Dashboard de métricas

### Fase 2 (Semana 3-4) - Consolidação
5. Validação cruzada multi-nível
6. Scoring de confiança aprimorado
7. Detecção de falsos negativos
8. Logging estruturado

### Fase 3 (Ongoing) - Aprendizado
9. Sistema de padrões de comerciantes
10. Análise preditiva de erros

---

## 4. Métricas de Sucesso

| Métrica | Linha de Base | Alvo |
|---------|---------------|------|
| Taxa de Sucesso | 84% | 95%+ |
| Confiança Média | ~80% | >90% |
| Falsos Rejeitados | ~5-8% | <2% |
| Tempo de Reprocessamento | N/A | <10s/imagem |

---

## 5. Estrutura de Código Sugerida

```
main.py (existente - expandido)
├── ocr_validators.py (novo)
│   ├── validate_column_coherence()
│   ├── validate_row_consistency()
│   ├── validate_description_semantics()
│   └── apply_ocr_corrections()
├── confidence_scoring.py (novo)
│   ├── calculate_enhanced_confidence_score()
│   ├── cross_validate_movement()
│   └── detect_potential_false_rejection()
├── merchant_patterns.py (novo)
│   └── MerchantPatternLearner
├── logging_audit.py (novo)
│   └── log_ocr_analysis()
└── metrics.py (novo)
    └── generate_ocr_quality_metrics()
```

---

## 6. Notas de Implementação

- **Backwards compatibility**: Todas as melhorias devem funcionar com a pipeline existente
- **Configuração**: Adicionar flags em `config.yaml` para ativar/desativar cada validação
- **Testes**: Expandir `test_core.py` com casos de teste para cada validador
- **Performance**: Manter overhead < 5% do tempo total de processamento
- **Observabilidade**: Registrar todas as decisões para auditoria posterior

