
"""
=============================================================================
ORQUESTRADOR DO PIPELINE OPERACIONAL  v3.2.0
=============================================================================
v3.2.0 — Correção crítica:
  - PYTHON_SGF e PYTHON_EMAIL agora são configuráveis via .env de forma
    independente. Isso resolve o ModuleNotFoundError quando o orquestrador
    é iniciado de dentro de um .venv diferente do script de extração SGF.
  - Fallback: se PYTHON_SGF não estiver definido, usa sys.executable
    (comportamento anterior, mantido para compatibilidade).

Variáveis de ambiente adicionadas:
  PYTHON_SGF   — Python do .venv do projeto SGF   (com selenium instalado)
  PYTHON_EMAIL — Python do .venv do projeto e-mail (com as libs de e-mail)
=============================================================================
"""

import os
import sys
import subprocess
import re
import time
from datetime import datetime
from pathlib import Path


# ══════════════════════════════════════════════════════════════
# UTILITÁRIOS
# ══════════════════════════════════════════════════════════════

def _carregar_dotenv(arquivo: str = ".env", niveis: int = 3) -> None:
    diretorio = Path(__file__).resolve().parent
    for _ in range(niveis):
        caminho = diretorio / arquivo
        if caminho.exists():
            for linha in caminho.read_text(encoding="utf-8").splitlines():
                linha = linha.strip()
                if not linha or linha.startswith("#") or "=" not in linha:
                    continue
                chave, valor = linha.split("=", 1)
                chave = chave.strip()
                valor = valor.strip().strip('"').strip("'")
                if chave and chave not in os.environ:
                    os.environ[chave] = valor
            return
        diretorio = diretorio.parent


_carregar_dotenv()

# ── Configurações ────────────────────────────────────────────
BASE_DIR     = Path(__file__).resolve().parent
SGF_SCRIPT   = Path(os.getenv("SGF_SCRIPT",   str(BASE_DIR / "sgf_pre_email.py")))
EMAIL_SCRIPT = Path(os.getenv("EMAIL_SCRIPT", str(BASE_DIR / "relatorios.py")))
LOG_SGF      = Path(os.getenv("LOG_SGF",      str(BASE_DIR / "log_sgf.txt")))
PIPELINE_LOG = Path(os.getenv("PIPELINE_LOG", str(BASE_DIR / "pipeline.log")))
EXCEL_PATH   = Path(os.getenv("EXCEL_PATH", "")).expanduser() if os.getenv("EXCEL_PATH") else None

# ── Executáveis Python por script ────────────────────────────
# PYTHON_SGF   → .venv do projeto SGF (precisa do selenium)
# PYTHON_EMAIL → .venv do projeto de e-mail
# Fallback para sys.executable caso não estejam definidos.
PYTHON_SGF   = os.getenv("PYTHON_SGF",   sys.executable)
PYTHON_EMAIL = os.getenv("PYTHON_EMAIL", sys.executable)


# ══════════════════════════════════════════════════════════════
# LOG DO ORQUESTRADOR
# ══════════════════════════════════════════════════════════════

def _log(nivel: str, msg: str) -> None:
    ts    = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    linha = f"[{ts}] [{nivel:<7}] {msg}"
    print(linha)
    with open(PIPELINE_LOG, "a", encoding="utf-8") as f:
        f.write(linha + "\n")

def log_info(m):  _log("INFO",  m)
def log_ok(m):    _log("OK",    m)
def log_warn(m):  _log("WARN",  m)
def log_erro(m):  _log("ERRO",  m)

def _cabecalho() -> None:
    linha = "═" * 60
    bloco = (
        f"\n{linha}\n"
        f"  PIPELINE — {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n"
        f"  Python SGF   — {PYTHON_SGF}\n"
        f"  Python Email — {PYTHON_EMAIL}\n"
        f"{linha}\n"
    )
    with open(PIPELINE_LOG, "a", encoding="utf-8") as f:
        f.write(bloco)
    print(bloco, end="")


# ══════════════════════════════════════════════════════════════
# AGUARDAR ARQUIVO DE REDE COM RETRY
# ══════════════════════════════════════════════════════════════

def _aguardar_arquivo_rede(
    caminho: Path,
    tentativas: int = 3,
    intervalo_s: int = 60,
) -> bool:
    if caminho is None:
        log_warn("EXCEL_PATH não configurado — pulando verificação de rede.")
        return True

    for t in range(1, tentativas + 1):
        if caminho.exists():
            log_ok(f"Arquivo de rede acessível: {caminho}")
            return True
        if t < tentativas:
            log_warn(
                f"Arquivo de rede inacessível (tentativa {t}/{tentativas}): {caminho}\n"
                f"         Aguardando {intervalo_s}s antes de tentar novamente..."
            )
            time.sleep(intervalo_s)
        else:
            log_erro(
                f"Arquivo de rede permanece inacessível após {tentativas} tentativas: {caminho}"
            )

    return False


# ══════════════════════════════════════════════════════════════
# PARSE DO LOG SGF
# ══════════════════════════════════════════════════════════════

def _analisar_log_sgf() -> dict:
    resultado = {
        "sucesso":     True,
        "ids_ok":      [],
        "ids_erro":    [],
        "linhas_erro": [],
    }

    if not LOG_SGF.exists():
        resultado["sucesso"] = False
        resultado["linhas_erro"].append("Arquivo de log da extração não encontrado.")
        return resultado

    texto = LOG_SGF.read_text(encoding="utf-8", errors="replace")

    def _extrair_ids(m):
        if not m:
            return []
        return re.findall(r"\b\d+\b", m.group(1))

    resultado["ids_ok"]   = _extrair_ids(re.search(r"Sucesso\s*:\s*([^\n]+)",  texto))
    resultado["ids_erro"] = _extrair_ids(re.search(r"Com erro\s*:\s*([^\n]+)", texto))

    m_stat = re.search(r"Status\s*:\s*([^\n]+)", texto)
    if m_stat and "ERROS" in m_stat.group(1).upper():
        resultado["sucesso"] = False

    if resultado["ids_erro"]:
        resultado["sucesso"] = False

    erros = [l for l in texto.splitlines() if "[ERRO" in l or "[WARN" in l]
    resultado["linhas_erro"] = erros[:10]

    return resultado


# ══════════════════════════════════════════════════════════════
# INJEÇÃO DO BANNER DE AVISO NO SCRIPT DE E-MAIL
# ══════════════════════════════════════════════════════════════

_SENTINELA_AVISO = "# __PIPELINE_AVISO_INJETADO__"

def _injetar_aviso_email(analise: dict) -> bool:
    if analise["sucesso"]:
        return False

    ids_erro_str = ", ".join(analise["ids_erro"]) or "—"
    linhas_log   = "<br>".join(analise["linhas_erro"])

    banner_html = (
        '<div style="background:#fff3cd;border-left:4px solid #e6a817;'
        'padding:12px 16px;margin-bottom:20px;border-radius:4px;">'
        '<p style="margin:0;font-size:13px;font-weight:700;color:#856404;">'
        '⚠ Aviso — Falha parcial na extração de dados</p>'
        f'<p style="margin:6px 0 0;font-size:12px;color:#664d03;">'
        f'IDs com erro: <strong>{ids_erro_str}</strong>. '
        'Os dados destes relatórios podem estar desatualizados ou ausentes.</p>'
        + (
            f'<details style="margin-top:8px;">'
            f'<summary style="font-size:11px;color:#664d03;cursor:pointer;">Ver log de erros</summary>'
            f'<pre style="font-size:10px;color:#664d03;margin:6px 0 0;'
            f'white-space:pre-wrap;">{linhas_log}</pre></details>'
            if linhas_log else ""
        )
        + '</div>'
    )

    conteudo = EMAIL_SCRIPT.read_text(encoding="utf-8")

    if _SENTINELA_AVISO in conteudo:
        return False

    padrao = re.compile(r'^(TEXTO_CORPO\s*=\s*["\(])', re.MULTILINE)
    m = padrao.search(conteudo)
    if not m:
        log_warn("Não foi possível localizar TEXTO_CORPO no script de e-mail. Aviso não injetado.")
        return False

    insercao = (
        f"\n{_SENTINELA_AVISO}\n"
        f"_AVISO_EXTRACAO = (\n"
        f"    '{banner_html}'\n"
        f")\n"
        f"TEXTO_CORPO = _AVISO_EXTRACAO + TEXTO_CORPO\n\n"
    )

    novo_conteudo = conteudo[:m.start()] + insercao + conteudo[m.start():]
    EMAIL_SCRIPT.write_text(novo_conteudo, encoding="utf-8")
    log_info("Banner de aviso injetado no script de e-mail.")
    return True


def _reverter_injecao() -> None:
    if not EMAIL_SCRIPT.exists():
        return
    conteudo = EMAIL_SCRIPT.read_text(encoding="utf-8")
    if _SENTINELA_AVISO not in conteudo:
        return

    padrao = re.compile(
        rf"\n{re.escape(_SENTINELA_AVISO)}\n.*?TEXTO_CORPO = _AVISO_EXTRACAO \+ TEXTO_CORPO\n\n",
        re.DOTALL,
    )
    novo_conteudo = padrao.sub("", conteudo)
    EMAIL_SCRIPT.write_text(novo_conteudo, encoding="utf-8")
    log_ok("Injeção revertida — script de e-mail restaurado.")


# ══════════════════════════════════════════════════════════════
# EXECUÇÃO DOS SUBPROCESSOS
# ══════════════════════════════════════════════════════════════

def _executar_script(script: Path, nome: str, python_exe: str) -> int:
    if not script.exists():
        log_erro(f"Script não encontrado: {script}")
        return -1

    log_info(f"▶ Iniciando: {nome}  ({script})")
    log_info(f"  Executável Python: {python_exe}")
    t0 = time.time()

    resultado = subprocess.run(
        [python_exe, str(script)],
        cwd=str(script.parent),
        capture_output=False,
    )

    duracao = time.time() - t0
    codigo  = resultado.returncode

    if codigo == 0:
        log_ok(f"✔ {nome} concluído em {duracao:.1f}s (código {codigo})")
    else:
        log_erro(f"✘ {nome} encerrou com código {codigo} em {duracao:.1f}s")

    return codigo


# ══════════════════════════════════════════════════════════════
# PIPELINE PRINCIPAL
# ══════════════════════════════════════════════════════════════

def main() -> None:
    _cabecalho()
    codigo_email = -1

    # ── ETAPA 0: Verificar acessibilidade do arquivo de rede ─
    log_info("━" * 40)
    log_info("ETAPA 0 — Verificação do arquivo de rede")
    log_info("━" * 40)

    rede_ok = _aguardar_arquivo_rede(EXCEL_PATH, tentativas=3, intervalo_s=60)
    if not rede_ok:
        log_erro(
            "Arquivo de rede inacessível após todas as tentativas. "
            "O e-mail será enviado sem as imagens atualizadas."
        )

    # ── ETAPA 1: Extração SGF ─────────────────────────────────
    log_info("━" * 40)
    log_info("ETAPA 1 — Extração SGF")
    log_info("━" * 40)

    codigo_sgf = _executar_script(SGF_SCRIPT, "Extração SGF", PYTHON_SGF)

    # ── ETAPA 2: Análise do log de extração ──────────────────
    log_info("Analisando resultado da extração...")
    analise = _analisar_log_sgf()

    if analise["sucesso"]:
        log_ok(f"Extração OK — IDs processados: {analise['ids_ok']}")
    else:
        log_warn(
            f"Extração com falhas — "
            f"OK: {analise['ids_ok']} | Erro: {analise['ids_erro']}"
        )

    # ── ETAPA 3: Envio de e-mail ──────────────────────────────
    log_info("━" * 40)
    log_info("ETAPA 2 — Envio de E-mail")
    log_info("━" * 40)

    injetou = False
    try:
        injetou = _injetar_aviso_email(analise)
        codigo_email = _executar_script(EMAIL_SCRIPT, "Envio de E-mail", PYTHON_EMAIL)
    finally:
        if injetou:
            _reverter_injecao()

    # ── Resumo ────────────────────────────────────────────────
    log_info("━" * 40)
    log_info("RESUMO DO PIPELINE")
    log_info("━" * 40)
    log_info(f"  Python SGF   : {PYTHON_SGF}")
    log_info(f"  Python Email : {PYTHON_EMAIL}")
    log_info(f"  Rede OK      : {'✔' if rede_ok else '✘'}")
    log_info(f"  Extração     : {'✔ OK' if codigo_sgf == 0 else '✘ ERRO'} (código {codigo_sgf})")
    log_info(f"  E-mail       : {'✔ OK' if codigo_email == 0 else '✘ ERRO'} (código {codigo_email})")

    sucesso_geral = (codigo_sgf == 0 and codigo_email == 0)
    if sucesso_geral:
        log_ok("Pipeline concluído com sucesso.")
    else:
        log_warn("Pipeline concluído com um ou mais erros. Verifique os logs acima.")

    sys.exit(0 if sucesso_geral else 1)


if __name__ == "__main__":
    main()
