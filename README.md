# Envio de Emails - Indicadores Eldorado

Este repositório contém um script Python para extrair tabelas de uma planilha Excel, gerar imagens separadas para cada bloco identificado e enviar um e-mail via Outlook com as imagens inline e a planilha anexada.

## O que o projeto faz

- Lê a planilha `data/Indicadores Produção.xlsx`
- Identifica blocos de indicadores e sub-tabelas
- Gera imagens PNG separadas para cada tabela
- Monta um e-mail HTML com as imagens embutidas
- Envia a mensagem usando Outlook COM

## Estrutura do projeto

- `relatórios.py` — script principal
- `.env` — variáveis de configuração privadas
- `.gitignore` — exclui arquivos sensíveis/localmente gerados
- `data/` — pasta com a planilha Excel de entrada

## Configuração

1. Crie e ative seu ambiente Python:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

2. Instale dependências:

```powershell
pip install pandas openpyxl dataframe_image pillow pywin32
```

3. Configure o `.env` com seus valores:

```env
EMAIL_REMETENTE=seu.email@empresa.com
ASSINATURA_NOME=Seu Nome Completo
ASSINATURA_CARGO=Seu Cargo
ASSINATURA_TEL=+55 (xx) x xxxx-xxxx
DESTINATARIOS=dest1@empresa.com,dest2@empresa.com
```

> Importante: não commite o `.env` e não inclua dados sensíveis no repositório.

## Uso

Execute o script:

```powershell
python relatórios.py
```

## Boas práticas

- Mantenha o `.env` local apenas no seu computador
- Não commit o arquivo Excel `data/Indicadores Produção.xlsx` se houver dados confidenciais
- Use `.gitignore` para bloquear arquivos de dados, ambientes virtuais e o `.env`

## Observações

- O envio é feito via Outlook instalado localmente no Windows
- O Outlook precisa estar configurado com a conta correta para enviar
- O arquivo Excel deve estar na pasta `data/`
