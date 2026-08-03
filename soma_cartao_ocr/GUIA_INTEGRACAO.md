# Guia de Integração - Melhorias de Confiabilidade

Este guia mostra como integrar os novos módulos de validação e scoring no pipeline existente.

## 📋 Índice

1. [Visão Geral](#visão-geral)
2. [Módulos Disponíveis](#módulos-disponíveis)
3. [Exemplos de Uso](#exemplos-de-uso)
4. [Integração com main.py](#integração-com-mainpy)
5. [Configuração](#configuração)
6. [Exemplos Práticos](#exemplos-práticos)

---

## Visão Geral

Os novos módulos fornecem validação em múltiplas dimensões:

```
                    ┌─────────────────────────────────────┐
                    │  Texto OCR Raw (Google Cloud Vision) │
                    └────────────────┬────────────────────┘
                                     │
                ┌────────────────────┼────────────────────┐
                │                    │                    │
                ▼                    ▼                    ▼
         ┌─────────────┐      ┌─────────────┐     ┌────────────┐
         │  OCR        │      │  Validação  │     │  Scoring   │
         │  Validators │      │  Cruzada    │     │  Avançado  │
         └──────┬──────┘      └──────┬──────┘     └─────┬──────┘
                │                    │                   │
                └────────────────────┼───────────────────┘
                                     │
                                     ▼
                          ┌──────────────────────┐
                          │  Decision: VÁLIDO/   │
                          │  REVISÃO (improved)  │
                          └──────────────────────┘
```

---

## Módulos Disponíveis

### 1. `ocr_validators.py`
Validadores de baixo nível para qualidade de dados.

**Funções principais:**
- `apply_ocr_corrections()` - Corrige erros comuns de OCR
- `validate_column_coherence()` - Valida alinhamento de coluna
- `validate_row_consistency()` - Detecta linhas fantasma
- `validate_description_semantics()` - Valida sentido semântico
- `detect_phantom_row()` - Identifica artefatos de OCR

### 2. `confidence_scoring.py`
Sistema de scoring de confiança multi-dimensional.

**Funções principais:**
- `calculate_enhanced_confidence_score()` - Score ponderado com breakdown
- `cross_validate_movement()` - Validações cruzadas de negócio
- `detect_potential_false_rejection()` - Identifica rejeições incorretas
- `calculate_rejection_confidence_threshold()` - Threshold dinâmico

### 3. `merchant_patterns.py`
Aprendizado de padrões históricos de comerciantes.

**Classe principal:** `MerchantPatternLearner`
- `learn_from_movement()` - Aprende de movimentos válidos
- `get_merchant_confidence()` - Score de conhecimento
- `get_expected_amount_range()` - Intervalo típico de valores
- `find_similar_merchants()` - Busca padrões similares
- `get_merchant_info()` - Informações completas

### 4. `metrics.py`
Análise de qualidade e tendências.

**Funções principais:**
- `generate_ocr_quality_metrics()` - Métricas agregadas
- `format_metrics_report()` - Relatório formatado
- `identify_problematic_merchants()` - Comerciantes problemáticos
- `get_recommendation()` - Recomendações de ação

---

## Exemplos de Uso

### Exemplo 1: Aplicar Correções OCR

```python
from ocr_validators import apply_ocr_corrections

# Texto com erro comum
texto_errado = "COMISSAD"
confidence = 0.72

texto_corrigido, foi_corrigido = apply_ocr_corrections(texto_errado, confidence)
print(f"Original: {texto_errado}")
print(f"Corrigido: {texto_corrigido}")
print(f"Aplicada correção: {foi_corrigido}")

# Output:
# Original: COMISSAD
# Corrigido: COMISSÃO
# Aplicada correção: True
```

### Exemplo 2: Validar Coerência de Coluna

```python
from ocr_validators import validate_column_coherence
from main import Word  # ou importar Word de ocr_validators

# Criar palavras fictícias para teste
words = [
    Word("26/06", 100, 50, 150, 65, 0.95),
    Word("MERCADONA", 200, 50, 350, 65, 0.85),
    Word("25.50", 400, 50, 450, 65, 0.60),
]

# Definir colunas (com ratios de 0 a 1)
columns = {
    "data_movimento": (0.0, 0.2),
    "descricao": (0.2, 0.6),
    "debito_eur": (0.6, 1.0),
}

issues = validate_column_coherence(words, columns, image_width=500)

if issues:
    print(f"Problemas encontrados: {issues}")
else:
    print("Colunas coerentes!")
```

### Exemplo 3: Calcular Score de Confiança Aprimorado

```python
from confidence_scoring import calculate_enhanced_confidence_score
from main import Movement

# Movimento exemplo
movement = Movement(
    line=1,
    data_movimento="26/06",
    data_valor="25/06",
    descricao="MERCADONA",
    pais="ESP",
    moeda_original="",
    taxa_cambio="",
    debito_eur="35.15",
    credito_eur="",
    confidence=0.82,
    status="VÁLIDO",
    motivos_revisao="",
    texto_ocr="26/06 25/06 MERCADONA 35.15"
)

# Configuração
cfg = {
    "validation": {
        "year": 2026,
        "allowed_months": [6, 7],
        "minimum_confidence": 0.75,
    },
    "table": {
        "width": 1000,
        "columns": {
            "data_movimento": (0.0, 0.15),
            "data_valor": (0.15, 0.3),
            "descricao": (0.3, 0.65),
            "debito_eur": (0.65, 0.85),
        }
    }
}

# Calcular score
breakdown = calculate_enhanced_confidence_score(
    movement=movement,
    row_words=[],  # Vazio neste exemplo
    table_cfg=cfg["table"],
    cfg=cfg,
    verbose=True
)

print(f"Score final: {breakdown.final_score:.2%}")
print(f"OCR contribution: {breakdown.ocr_score:.2%}")
print(f"Detalhes: {breakdown.factors}")
```

### Exemplo 4: Usar MerchantPatternLearner

```python
from merchant_patterns import MerchantPatternLearner
from pathlib import Path

# Inicializar
learner = MerchantPatternLearner("data/merchant_patterns.json")

# Aprender de um movimento válido
for movimento_valido in movimentos_validados:
    learner.learn_from_movement(movimento_valido)

# Salvar padrões
learner.save_patterns()

# Consultar informações de comerciante
info = learner.get_merchant_info("MERCADONA")
print(f"Conhecimento sobre MERCADONA:")
print(f"  • Ocorrências: {info['occurrences']}")
print(f"  • Intervalo típico: €{info['typical_amount_range']}")
print(f"  • Confiança: {info['confidence_score']:.1%}")

# Validar montante
desc = "MERCADONA"
amount = 45.50
is_valid = learner.is_expected_amount(desc, amount)
print(f"Montante €{amount} típico para {desc}? {is_valid}")

# Encontrar comerciantes similares
similar = learner.find_similar_merchants("MERCADOR")
print(f"Similares a MERCADOR: {similar}")
```

### Exemplo 5: Gerar Métricas de Qualidade

```python
from metrics import (
    generate_ocr_quality_metrics,
    format_metrics_report,
    identify_problematic_merchants,
)

# Gerar métricas de um lote de movimentos
metrics = generate_ocr_quality_metrics(movimentos)

# Imprimir relatório
print(format_metrics_report(metrics))

# Identificar problemas
problem_merchants = identify_problematic_merchants(movimentos)
for merchant, stats in problem_merchants.items():
    print(f"{merchant}: {stats['rejection_rate']:.0%} rejeição")

# Exportar para JSON
import json
with open("metrics.json", "w") as f:
    json.dump(metrics.__dict__, f, indent=2)
```

---

## Integração com main.py

### Passo 1: Adicionar Imports

No início de `main.py`, adicione:

```python
from ocr_validators import (
    apply_ocr_corrections,
    validate_column_coherence,
    validate_row_consistency,
    validate_description_semantics,
)
from confidence_scoring import (
    calculate_enhanced_confidence_score,
    cross_validate_movement,
    detect_potential_false_rejection,
)
from merchant_patterns import MerchantPatternLearner
from metrics import (
    generate_ocr_quality_metrics,
    format_metrics_report,
)
```

### Passo 2: Integrar Validações no Pipeline

Modificar a função `build_movements()` em `main.py`:

```python
def build_movements(
    rows: list[list[Word]],
    width: int,
    cfg: dict[str, Any],
    trusted_texts: set[str] | list[str] | None = None
) -> list[Movement]:
    """Build movements com validações aprimoradas."""
    
    movements = []
    allowed_months = list(cfg.get("validation", {}).get("allowed_months", [6, 7]))
    trusted_set = {str(t).strip().upper() for t in trusted_texts if t and str(t).strip()} if trusted_texts else set()
    
    # Carregar padrões de comerciantes se habilitado
    use_patterns = cfg.get("features", {}).get("use_merchant_patterns", True)
    merchant_learner = MerchantPatternLearner() if use_patterns else None

    for index, row in enumerate(rows, 1):
        fields = split_columns(row, width, cfg["table"]["columns"])
        raw = " ".join(word.text for word in row)
        
        # ✨ NOVA: Validar consistência da linha
        is_valid_row, row_reason, row_details = validate_row_consistency(row)
        if not is_valid_row:
            continue

        # ... resto do processamento de datas ...

        desc_raw = normalize_description_text(fields["descricao"])
        desc_clean, pais_code = extract_country_and_clean_desc(desc_raw, fields.get("pais", ""))
        
        # ✨ NOVA: Aplicar correções automáticas de OCR
        if cfg.get("features", {}).get("auto_correct_ocr", True):
            desc_clean, _ = apply_ocr_corrections(desc_clean, np.mean([w.confidence for w in row]))
        
        fields["descricao"] = desc_clean
        fields["pais"] = pais_code
        
        # ✨ NOVA: Validar semântica
        desc_valid, sem_reason = validate_description_semantics(desc_clean)
        
        # ... resto do processamento ...
        
        # Construir movimento
        confidence = sum(word.confidence for word in row) / max(len(row), 1)
        
        reasons = []
        # ... validações existentes ...
        
        # ✨ NOVA: Validação de confiança aprimorada
        if cfg.get("features", {}).get("enhanced_confidence_scoring", True):
            breakdown = calculate_enhanced_confidence_score(
                Movement(...),  # Movimento temporário
                row,
                cfg.get("table", {}),
                cfg=cfg,
            )
            enhanced_confidence = breakdown.final_score
            
            if enhanced_confidence < float(cfg["validation"]["minimum_confidence"]):
                reasons.append(f"confiança aprimorada baixa ({enhanced_confidence:.1%})")
        
        movements.append(Movement(
            index,
            fields["data_movimento"],
            fields["data_valor"],
            fields["descricao"],
            fields["pais"],
            fields["moeda_original"],
            fields["taxa_cambio"],
            fields["debito_eur"],
            fields["credito_eur"],
            round(confidence, 4),
            "REVISÃO" if reasons else "VÁLIDO",
            "; ".join(reasons),
            raw
        ))

    # ✨ NOVA: Aprender de movimentos válidos
    if merchant_learner:
        for m in movements:
            if m.status == "VÁLIDO":
                merchant_learner.learn_from_movement(m)
        merchant_learner.save_patterns()

    return movements
```

### Passo 3: Gerar Métricas de Qualidade

Ao final de `main()`:

```python
def main() -> int:
    # ... código existente ...
    
    movements = build_movements(rows, image.shape[1], cfg, trusted_texts=trusted_texts)
    
    # ✨ NOVA: Gerar métricas
    if cfg.get("features", {}).get("generate_metrics", True):
        metrics = generate_ocr_quality_metrics(movements)
        
        # Salvar métricas
        metrics_file = out_dir / "metricas.json"
        import json
        with open(metrics_file, "w") as f:
            json.dump(metrics.__dict__, f, default=str, indent=2)
        
        # Imprimir relatório
        print(format_metrics_report(metrics))
        
        # Adicionar aos metadados
        metadata.update({
            "quality_metrics": {
                "success_rate": f"{metrics.success_rate:.1f}%",
                "avg_confidence": f"{metrics.avg_confidence:.1%}",
                "quality_trend": metrics.quality_trend,
            }
        })
    
    # ... resto do código ...
```

---

## Configuração

### Adicionar ao `config.yaml`

```yaml
# Novos recursos de validação
features:
  # Corrigir automaticamente erros comuns de OCR
  auto_correct_ocr: true
  
  # Usar scoring de confiança aprimorado (multi-dimensional)
  enhanced_confidence_scoring: true
  
  # Usar padrões históricos de comerciantes
  use_merchant_patterns: true
  
  # Gerar e registrar métricas de qualidade
  generate_metrics: true
  
  # Tentar recuperar falsos negativos
  recover_false_rejections: true

# Configuração de validação
validation:
  year: 2026
  allowed_months: [6, 7]
  minimum_confidence: 0.75
  require_two_dates: true
  require_exactly_one_amount: true

# Configuração de detecção de linhas
table:
  width: 1000
  row_tolerance_ratio: 0.01
  header_terms:
    - data
    - movimento
    - descrição
    - débito
    - crédito
  columns:
    data_movimento: [0.0, 0.15]
    data_valor: [0.15, 0.3]
    descricao: [0.3, 0.65]
    pais: [0.65, 0.75]
    debito_eur: [0.75, 0.9]
    credito_eur: [0.9, 1.0]

# Padrões de comerciantes
merchant_patterns:
  enabled: true
  history_file: "data/merchant_patterns.json"
```

---

## Exemplos Práticos

### Cenário 1: Melhorar Taxa de Sucesso

**Problema**: Taxa de sucesso em 84%, precisa chegar a 95%

**Solução**:
1. Ativar `auto_correct_ocr`
2. Usar `enhanced_confidence_scoring`
3. Ativar `recover_false_rejections`
4. Analisar `identify_problematic_merchants()` para padrões

**Impacto esperado**: +8-12% de sucesso

### Cenário 2: Reduzir Rejeições Falsas

**Problema**: Comerciantes conhecidos sendo rejeitados

**Solução**:
1. Usar `MerchantPatternLearner` para construir histórico
2. Usar `detect_potential_false_rejection()` para reconhecer
3. Configurar `trusted_descriptions` com comerciantes confiáveis

**Impacto esperado**: Eliminar 90% de falsos negativos

### Cenário 3: Detectar Problemas de Qualidade de Imagem

**Problema**: Algumas imagens têm qualidade ruim

**Solução**:
1. Gerar `metrics` para cada lote
2. Analisar `confidence_std` (deve ser < 0.15)
3. Se alto, melhorar pré-processamento (aumento de contraste, etc.)

**Impacto esperado**: Melhor diagnóstico de problemas upstream

---

## Testes Unitários

Exemplo de teste:

```python
import unittest
from ocr_validators import apply_ocr_corrections

class TestOCRValidators(unittest.TestCase):
    def test_apply_ocr_corrections(self):
        """Testa aplicação de correções."""
        text, corrected = apply_ocr_corrections("COMISSAD", 0.72)
        self.assertEqual(text, "COMISSÃO")
        self.assertTrue(corrected)

if __name__ == "__main__":
    unittest.main()
```

---

## Troubleshooting

| Problema | Solução |
|----------|---------|
| Muitas rejeições ainda | Aumentar `minimum_confidence` gradualmente; revisar `top_rejection_reasons` |
| Alto desvio padrão de confiança | Melhorar qualidade de imagem; ajustar pré-processamento |
| Padrões de comerciantes vazios | Garantir que `learn_from_movement()` é chamado; verificar status="VÁLIDO" |
| Performance lenta | Desativar features não essenciais; otimizar `validate_column_coherence()` |

---

## Próximos Passos

1. ✅ Implementar básica: correções OCR + validação
2. ⏳ Fase 2: Scoring completo + padrões
3. ⏳ Fase 3: Análise preditiva + alertas

