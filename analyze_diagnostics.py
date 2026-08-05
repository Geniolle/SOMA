#!/usr/bin/env python3
import json
import sys
from pathlib import Path

# Force UTF-8 output
if sys.stdout.encoding != 'utf-8':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

diagnostics_dir = Path("artifacts/diagnostics")
if not diagnostics_dir.exists():
    print(f"Erro: {diagnostics_dir} não existe")
    sys.exit(1)

# Procurar por arquivos JSON de wait_any_present
json_files = sorted(diagnostics_dir.glob("wait_any_present_timeout*.json"), key=lambda p: p.stat().st_mtime, reverse=True)

if not json_files:
    print("Nenhum arquivo de diagnóstico wait_any_present encontrado")
    print("\nArquivos disponíveis:")
    for f in sorted(diagnostics_dir.glob("*.json"), reverse=True)[:10]:
        print(f"  - {f.name}")
    sys.exit(0)

print(f"Analisando {len(json_files)} arquivo(s) de diagnóstico...\n")

for json_file in json_files[:3]:  # Analisar os 3 mais recentes
    print(f"\n{'='*80}")
    print(f"Arquivo: {json_file.name}")
    print('='*80)

    try:
        data = json.loads(json_file.read_text(encoding="utf-8"))

        for i, entry in enumerate(data, 1):
            locator = entry.get("locator", [])
            method = locator[0] if len(locator) > 0 else "unknown"
            selector = locator[1] if len(locator) > 1 else "unknown"
            found = entry.get("found", False)
            count = entry.get("count", 0)

            status = "✓ ENCONTRADO" if found else "✗ não encontrado"

            print(f"\n[{i}] {status} (count={count})")
            print(f"    Método: {method}")
            print(f"    Selector: {selector}")

            if found:
                print(f"    Tag: {entry.get('tag')}")
                print(f"    Text: {entry.get('text')}")
                print(f"    Value: {entry.get('value')}")
                print(f"    Name: {entry.get('name')}")
                print(f"    ID: {entry.get('id')}")
                print(f"    Class: {entry.get('class')}")
                print(f"    Displayed: {entry.get('displayed')}")
                print(f"    Enabled: {entry.get('enabled')}")
                if entry.get('outer_html'):
                    html = entry.get('outer_html')[:200]
                    print(f"    HTML: {html}...")
            else:
                error = entry.get("error", "unknown")
                print(f"    Erro: {error}")

    except json.JSONDecodeError as e:
        print(f"Erro ao ler JSON: {e}")
        continue
    except Exception as e:
        print(f"Erro: {e}")
        continue

print(f"\n{'='*80}")
print("Para ver o HTML completo, abra em um navegador:")
for f in sorted(diagnostics_dir.glob("wait_any_present_timeout*.html"), reverse=True)[:1]:
    print(f"  {f}")
