# Quick Reference - relatorios.py

## TL;DR

O sistema agora:
- ✅ Abre o workbook correto em rede UNC
- ✅ Atualiza dados Power Query ANTES de ler
- ✅ Extrai Celulose E Biomassa separadamente
- ✅ Envia 6 tabelas por e-mail HTML
- ✅ Anexa o arquivo Excel

## Executar

```bash
# Opção 1: Batch
executar.bat

# Opção 2: Python direto
python relatorios.py

# Opção 3: Pipeline com refresh
python run_pipeline.py
```

## Configurar (arquivo .env)

```
EXCEL_PATH=\\servidor\pasta\Report Operacional V2.xlsx
EMAIL_REMETENTE=seu.email@eldorado.com.br
DESTINATARIOS=email1@eldorado.com.br,email2@eldorado.com.br
ANEXAR_PLANILHA=true
```

## Abas Processadas

| Aba | Produto | Linhas |
|-----|---------|--------|
| Planilha1 | Celulose | 55 |
| Planilha2 | Biomassa | 26 |
| Volume Real Diário | Métricas (4 blocos) | 14+14+14+61 |

## Logs

Cada execução gera:
- `relatórios.log` - histórico de execução
- `report_temp_*.png` - imagens das tabelas

## Troubleshooting

| Problema | Solução |
|----------|---------|
| "Arquivo não encontrado" | Verifique caminho UNC e acesso à rede |
| "Cabeçalho não encontrado" | Workbook mudou de estrutura, verifique abas |
| "E-mail não enviado" | Verifique EMAIL_REMETENTE e DESTINATARIOS no .env |
| "Unicode error" | Arquivo log usa UTF-8, tudo ok |

## Código Chave

```python
# Importar e executar
import relatorios
relatorios.executar()

# Ou apenas carregar dados
from pathlib import Path
tabelas = relatorios.carregar_e_segmentar(
    relatorios.EXCEL_PATH,
    relatorios.SHEET_CONFIG
)
for titulo, df in tabelas:
    print(f"{titulo}: {df.shape}")
```

## Mantido Para Compatibilidade

- Outras abas e extractores ainda funcionam
- Não quebrando mudanças, apenas adições
- Todas as funções de e-mail/imagem intactas

---

**Última atualização:** 25/05/2026
