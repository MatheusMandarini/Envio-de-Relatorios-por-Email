"""
=============================================================================
ORQUESTRADOR DO PIPELINE OPERACIONAL  v3.4.0
=============================================================================
v3.4.0 — Melhorias:
  - Refresh do Excel movido para APÓS a extração SGF (Etapa 2).
  - Flags de linha de comando para pular etapas individualmente:
      --skip-rede      Pula verificação de acessibilidade do arquivo de rede
      --skip-sgf       Pula a extração SGF
      --skip-refresh   Pula o refresh do Excel
      --skip-email     Pula o envio de e-mail
    Exemplo:
      python pipeline_orquestrador.py --skip-sgf
      python pipeline_orquestrador.py --skip-sgf --skip-refresh

v3.3.0 — Refresh automático do Report Operacional via Excel COM (xlwings).
v3.2.0 — PYTHON_SGF e PYTHON_EMAIL configuráveis via .env de forma independente.

Variáveis de ambiente:
  PYTHON_SGF            — Python do .venv do projeto SGF (com selenium)
  PYTHON_EMAIL          — Python do .venv do projeto de e-mail
  EXCEL_PATH            — Caminho completo para o Report Operacional V2.xlsx
  EXCEL_REFRESH_TIMEOUT — Tempo máx. de espera pelo refresh em segundos (padrão 300)
=============================================================================
"""

import os
import sys
import argparse
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

EXCEL_REFRESH_TIMEOUT = int(os.getenv("EXCEL_REFRESH_TIMEOUT", "300"))

# ── Executáveis Python por script ────────────────────────────
PYTHON_SGF   = os.getenv("PYTHON_SGF",   sys.executable)
PYTHON_EMAIL = os.getenv("PYTHON_EMAIL", sys.executable)


# ══════════════════════════════════════════════════════════════
# ARGUMENTOS DE LINHA DE COMANDO
# ══════════════════════════════════════════════════════════════

def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Orquestrador do Pipeline Operacional",
        formatter_class=argparse.RawTextHelpFormatter,
    )
    parser.add_argument(
        "--skip-rede",
        action="store_true",
        help="Pula a verificação de acessibilidade do arquivo de rede",
    )
    parser.add_argument(
        "--skip-sgf",
        action="store_true",
        help="Pula a extração SGF (Etapa 1)",
    )
    parser.add_argument(
        "--skip-refresh",
        action="store_true",
        help="Pula o refresh do Report Operacional no Excel (Etapa 2)",
    )
    parser.add_argument(
        "--skip-email",
        action="store_true",
        help="Pula o envio de e-mail (Etapa 3)",
    )
    return parser.parse_args()


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
def log_skip(m):  _log("SKIP",  m)

def _cabecalho(args: argparse.Namespace) -> None:
    puladas = [
        etapa for flag, etapa in [
            (args.skip_rede,    "rede"),
            (args.skip_sgf,     "sgf"),
            (args.skip_refresh, "refresh"),
            (args.skip_email,   "email"),
        ] if flag
    ]
    info_puladas = f"  Etapas puladas : {', '.join(puladas)}\n" if puladas else ""

    linha = "═" * 60
    bloco = (
        f"\n{linha}\n"
        f"  PIPELINE — {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n"
        f"  Python SGF   — {PYTHON_SGF}\n"
        f"  Python Email — {PYTHON_EMAIL}\n"
        f"{info_puladas}"
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
# REFRESH DO EXCEL VIA COM (Power Query / xlwings)
# ══════════════════════════════════════════════════════════════

def _refresh_excel(caminho: Path, timeout_s: int = 300) -> bool:
    app = None
    wb = None
    excel_pid = None
    if caminho is None:
        log_warn("EXCEL_PATH não configurado — refresh ignorado.")
        return True

    if not caminho.exists():
        log_erro(f"Arquivo Excel não encontrado para refresh: {caminho}")
        return False

    try:
        import xlwings as xw
    except ImportError:
        log_erro(
            "Biblioteca 'xlwings' não encontrada.\n"
            "         Instale com: pip install xlwings\n"
            "         Refresh do Excel ignorado."
        )
        return False

    # ── NOVO: aguarda o arquivo ser liberado ──────────────────────────────────
    LOCK_TIMEOUT   = int(os.getenv("LOCK_TIMEOUT",   "600"))  # máx 10 min esperando
    LOCK_INTERVALO = int(os.getenv("LOCK_INTERVALO", "30"))   # testa a cada 30s

    aguardado_lock = 0
    while True:
        try:
            # Tenta abrir em modo exclusivo — falha se outro processo tiver o arquivo
            with open(caminho, "r+b"):
                pass
            log_ok("Arquivo disponível (sem bloqueio detectado).")
            break
        except (PermissionError, OSError):
            if aguardado_lock >= LOCK_TIMEOUT:
                log_erro(
                    f"Arquivo permanece bloqueado após {LOCK_TIMEOUT}s. "
                    "Refresh abortado."
                )
                return False
            log_warn(
                f"Arquivo em uso por outro processo. "
                f"Aguardando {LOCK_INTERVALO}s... "
                f"({aguardado_lock}/{LOCK_TIMEOUT}s)"
            )
            time.sleep(LOCK_INTERVALO)
            aguardado_lock += LOCK_INTERVALO
    # ─────────────────────────────────────────────────────────────────────────

    log_info(f"Abrindo Excel em background para refresh: {caminho.name}")
    app = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts  = False
        app.screen_updating = False
        excel_pid = app.pid
        log_info(f"Excel iniciado (PID={excel_pid})")
        wb = app.books.open(str(caminho))
        log_info("Disparando RefreshAll (Power Query)...")
        wb.api.RefreshAll()

        intervalo = 5
        aguardado = 0
        while aguardado < timeout_s:
            time.sleep(intervalo)
            aguardado += intervalo

            ainda_atualizando = False
            for sheet in wb.sheets:
                for qt in sheet.api.QueryTables:
                    if qt.Refreshing:
                        ainda_atualizando = True
                        break
                if ainda_atualizando:
                    break
                for lo in sheet.api.ListObjects:
                    try:
                        if lo.QueryTable.Refreshing:
                            ainda_atualizando = True
                            break
                    except Exception:
                        pass
                if ainda_atualizando:
                    break

            if not ainda_atualizando:
                log_ok(f"Refresh concluído em ~{aguardado}s.")
                break
        else:
            log_warn(
                f"Timeout de {timeout_s}s atingido aguardando o refresh. "
                "A planilha pode estar parcialmente atualizada."
            )

        wb.save()
        wb.close()
        log_ok(f"Arquivo salvo e fechado: {caminho.name}")
        return True

    except Exception as exc:
        log_erro(f"Falha no refresh do Excel: {exc}")
        return False

    finally:
        try:
            if wb is not None:
                wb.close()
        except Exception:
            pass

        try:
            if app is not None:
                app.quit()
        except Exception:
            pass

        # garante encerramento do Excel criado pelo xlwings
        try:
            if excel_pid:
                subprocess.run(
                    ["taskkill", "/PID", str(excel_pid), "/F", "/T"],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                )
                log_info(f"Excel encerrado à força (PID={excel_pid})")
        except Exception:
            pass


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
    args = _parse_args()
    _cabecalho(args)

    codigo_sgf   = 0   # 0 = OK por padrão quando pulado
    codigo_email = 0
    rede_ok      = True
    refresh_ok   = True

    # ── ETAPA 0: Verificar acessibilidade do arquivo de rede ──
    log_info("━" * 40)
    log_info("ETAPA 0 — Verificação do arquivo de rede")
    log_info("━" * 40)

    if args.skip_rede:
        log_skip("Verificação de rede PULADA (--skip-rede).")
    else:
        rede_ok = _aguardar_arquivo_rede(EXCEL_PATH, tentativas=3, intervalo_s=60)
        if not rede_ok:
            log_erro(
                "Arquivo de rede inacessível após todas as tentativas. "
                "O pipeline continuará sem atualizar a planilha."
            )

    # ── ETAPA 1: Extração SGF ─────────────────────────────────
    log_info("━" * 40)
    log_info("ETAPA 1 — Extração SGF")
    log_info("━" * 40)

    if args.skip_sgf:
        log_skip("Extração SGF PULADA (--skip-sgf).")
    else:
        codigo_sgf = _executar_script(SGF_SCRIPT, "Extração SGF", PYTHON_SGF)

    # ── ETAPA 2: Refresh do Report Operacional (Power Query) ──
    log_info("━" * 40)
    log_info("ETAPA 2 — Refresh do Report Operacional V2.xlsx")
    log_info("━" * 40)

    if args.skip_refresh:
        log_skip("Refresh do Excel PULADO (--skip-refresh).")
    elif not rede_ok or EXCEL_PATH is None:
        log_warn("Refresh ignorado — arquivo de rede inacessível.")
        refresh_ok = False
    else:
        refresh_ok = _refresh_excel(EXCEL_PATH, timeout_s=EXCEL_REFRESH_TIMEOUT)
        if not refresh_ok:
            log_warn(
                "Refresh do Excel falhou. "
                "O e-mail usará os dados da última atualização disponível."
            )

    # ── ETAPA 2.5: Análise do log de extração ─────────────────
    if not args.skip_sgf:
        log_info("Analisando resultado da extração...")
        analise = _analisar_log_sgf()
        if analise["sucesso"]:
            log_ok(f"Extração OK — IDs processados: {analise['ids_ok']}")
        else:
            log_warn(
                f"Extração com falhas — "
                f"OK: {analise['ids_ok']} | Erro: {analise['ids_erro']}"
            )
    else:
        # SGF pulado: considera extração ok para não bloquear o e-mail
        analise = {"sucesso": True, "ids_ok": [], "ids_erro": [], "linhas_erro": []}

    # ── ETAPA 3: Envio de e-mail ──────────────────────────────
    log_info("━" * 40)
    log_info("ETAPA 3 — Envio de E-mail")
    log_info("━" * 40)

    if args.skip_email:
        log_skip("Envio de e-mail PULADO (--skip-email).")
    else:
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

    def _status(pulado, ok, codigo=None):
        if pulado:
            return "— PULADO"
        base = "✔ OK" if ok else "✘ ERRO"
        return f"{base} (código {codigo})" if codigo is not None else base

    log_info(f"  Rede         : {_status(args.skip_rede,    rede_ok)}")
    log_info(f"  Extração SGF : {_status(args.skip_sgf,     codigo_sgf == 0, codigo_sgf)}")
    log_info(f"  Refresh Excel: {_status(args.skip_refresh, refresh_ok)}")
    log_info(f"  E-mail       : {_status(args.skip_email,   codigo_email == 0, codigo_email)}")

    # Considera sucesso geral apenas nas etapas que foram executadas
    falhas = []
    if not args.skip_sgf   and codigo_sgf   != 0: falhas.append("Extração SGF")
    if not args.skip_email and codigo_email != 0: falhas.append("Envio de E-mail")

    if not falhas:
        log_ok("Pipeline concluído com sucesso.")
        sys.exit(0)
    else:
        log_warn(f"Pipeline concluído com erros em: {', '.join(falhas)}. Verifique os logs.")
        sys.exit(1)


if __name__ == "__main__":
    try:
        main()

    except KeyboardInterrupt:
        log_warn("Execução interrompida pelo usuário (Ctrl+C).")

        # mata qualquer Excel iniciado pelo pipeline
        try:
            os.system("taskkill /F /IM EXCEL.EXE >nul 2>&1")
        except Exception:
            pass

        sys.exit(1)