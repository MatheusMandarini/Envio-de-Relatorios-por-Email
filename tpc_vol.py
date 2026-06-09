"""
=============================================================================
SISTEMA DE AUTOMAÇÃO CORPORATIVA — ENVIO DE REPORT OPERACIONAL POR E-MAIL
=============================================================================
Autor      : Matheus Menezes
Empresa    : Eldorado Brasil
Versão     : 3.1.0  (corrigido)

Correções v3.1.0:
  - Saída de log agora vai para ARQUIVO + CONSOLE simultaneamente (não some mais)
  - EXCEL_PATH pode ser caminho de rede (UNC) — configure no .env
  - Timeout explícito ao abrir arquivo Excel
  - Colunas vazias/espaço removidas do cabeçalho das tabelas
  - Filtro _tem_numero corrigido para funcionar após descarte de col[0]
  - Mensagens de progresso mais detalhadas para facilitar diagnóstico
=============================================================================
"""

import os
import sys
import time
import traceback
import numpy as np
from datetime import datetime
from pathlib import Path


# ===========================================================================
# LOG DUPLO: arquivo + console (resolve o "travamento invisível")
# ===========================================================================

class DualLog:
    """Escreve simultaneamente em arquivo e no console original."""
    def __init__(self, filename: str, mode: str = "a"):
        self._console = sys.__stdout__   # console original
        self._file    = open(filename, mode, encoding="utf-8")

    def write(self, msg):
        self._console.write(msg)
        self._console.flush()
        self._file.write(msg)
        self._file.flush()

    def flush(self):
        self._console.flush()
        self._file.flush()

    def close(self):
        self._file.close()

_dual_log = DualLog("relatórios.log")
sys.stdout = _dual_log
sys.stderr = _dual_log


# ===========================================================================
# .env
# ===========================================================================

def carregar_env(path: str = ".env") -> None:
    if not os.path.exists(path):
        return
    with open(path, encoding="utf-8") as f:
        for linha in f:
            linha = linha.strip()
            if not linha or linha.startswith("#") or "=" not in linha:
                continue
            chave, valor = linha.split("=", 1)
            chave = chave.strip()
            valor = valor.strip().strip('"').strip("'")
            if chave and chave not in os.environ:
                os.environ[chave] = valor

carregar_env()

import pandas as pd
import dataframe_image as dfi
from PIL import Image


# ===========================================================================
# CONFIGURAÇÕES
# ===========================================================================

# Caminho padrão — altere no .env como EXCEL_PATH=\\servidor\pasta\arquivo.xlsx
EXCEL_PATH = os.getenv("TPC_EXCEL_PATH", r"C:\Users\ext.matheusmm\Desktop\Relatório Estoque 28_05_2026.xlsx")
DESTINATARIOS = [e.strip() for e in os.getenv("TPC_DESTINATARIOS", "").split(",") if e.strip()]
ANEXAR_PLANILHA = os.getenv("TPC_ANEXAR_PLANILHA", "true").lower() == "true"

EMAIL_REMETENTE  = os.getenv("EMAIL_REMETENTE", "")
ASSINATURA_NOME  = os.getenv("ASSINATURA_NOME", "")
ASSINATURA_CARGO = os.getenv("ASSINATURA_CARGO", "")
ASSINATURA_TEL   = os.getenv("ASSINATURA_TEL", "")
ASSINATURA_EMAIL = EMAIL_REMETENTE


_dest_raw    = os.getenv("DESTINATARIOS", "")



SAUDACAO    = "Prezados,"
TEXTO_CORPO = (
    "Segue em anexo o <strong>Estoque Diário de Madeira/TPC</strong> da Eldorado Brasil, "
    "com os dados consolidados do período."
)

IMAGE_BASE = "report_temp"

SHEET_CONFIG: list[dict] = [
    {
        "sheet":    "Resumo",
        "titulo":   "Resumo de Estoque",
        "extrator": "extrator_resumo_estoque",
        "ativo":    True,
   },
]



# ===========================================================================
# UTILITÁRIOS
# ===========================================================================

def _is_nan(v) -> bool:
    try:
        return np.isnan(v)
    except (TypeError, ValueError):
        return False


def _str_val(v) -> str:
    if _is_nan(v):
        return ""
    s = str(v).strip()
    if "00:00:00" in s:
        try:
            return pd.to_datetime(v).strftime("%d/%m/%Y")
        except Exception:
            pass
    return s


def _fmt_num(v) -> str:
    if isinstance(v, (int, float)) and not _is_nan(v):
        if v == int(v):
            return f"{int(v):,}".replace(",", ".")
        return f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    s = _str_val(v)
    return s if s else "—"


def _limpar_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.dropna(how="all", axis=0).dropna(how="all", axis=1)
    return df.reset_index(drop=True)


def _formatar_df(df: pd.DataFrame) -> pd.DataFrame:
    return df.apply(lambda col: col.map(_fmt_num))


def _refresh_excel(path: str | Path, timeout_s: int | None = None) -> bool:
    caminho = Path(path)
    if not caminho.exists():
        print(f"[AVISO] Arquivo Excel não encontrado para refresh: {caminho}")
        return False

    try:
        import xlwings as xw
    except ImportError:
        print(
            "[AVISO] Biblioteca 'xlwings' não instalada. "
            "Refresh do Excel será ignorado."
        )
        return False

    timeout_s = timeout_s or int(os.getenv("EXCEL_REFRESH_TIMEOUT", "300"))
    print(f"[LOG] Abrindo Excel em background para refresh: {caminho.name}")

    app = None
    wb = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        wb = app.books.open(str(caminho), update_links=False)
        print("[LOG] Disparando RefreshAll (Power Query)...")
        wb.api.RefreshAll()

        inicio = time.time()
        while True:
            refreshing = False
            try:
                for qt in wb.api.QueryTables:
                    if getattr(qt, "Refreshing", False):
                        refreshing = True
                        break
            except Exception:
                pass

            if not refreshing:
                try:
                    for lo in wb.api.ListObjects:
                        qt = getattr(lo, "QueryTable", None)
                        if qt is not None and getattr(qt, "Refreshing", False):
                            refreshing = True
                            break
                except Exception:
                    pass

            if not refreshing:
                break

            if time.time() - inicio > timeout_s:
                print(
                    f"[AVISO] Timeout de {timeout_s}s atingido aguardando o refresh. "
                    "Continuando com os dados atuais."
                )
                return False

            time.sleep(2)

        print("[LOG] Refresh concluído.")
        return True

    except Exception as exc:
        print(f"[AVISO] Erro no refresh do Excel: {exc}")
        return False

    finally:
        if wb is not None:
            try:
                wb.close(save_changes=False)
            except Exception:
                pass
        if app is not None:
            try:
                app.quit()
            except Exception:
                pass


# ===========================================================================
# EXTRATORES
# ===========================================================================

def extrator_volume_real_diario(df: pd.DataFrame) -> list[tuple[str, pd.DataFrame]]:
    """
    Detecta blocos onde a linha seguinte ao título contém datas ou 'Rótulos de Coluna'.
    Extrai cada bloco como uma sub-tabela.
    """
    resultados = []
    n_rows = len(df)
    blocos = []

    for i in range(n_rows - 1):
        v1 = _str_val(df.iloc[i, 1]) if df.shape[1] > 1 else ""
        v2 = df.iloc[i + 1, 2]       if df.shape[1] > 2 else ""

        if not v1:
            continue

        if hasattr(v2, "strftime") or "Rótulos" in _str_val(v2):
            blocos.append((i, v1))

    if not blocos:
        print("[AVISO] Nenhum bloco detectado na aba 'Volume Real Diário'.")
        return resultados

    print(f"[LOG]   Blocos detectados: {[nome for _, nome in blocos]}")

    for b_idx, (linha_titulo, nome) in enumerate(blocos):
        linha_fim = blocos[b_idx + 1][0] if b_idx + 1 < len(blocos) else n_rows
        linha_cab = linha_titulo + 1

        # Montar cabeçalhos da linha de datas
        cab_row = df.iloc[linha_cab].tolist()
        cabecalhos = []
        for v in cab_row:
            if hasattr(v, "strftime"):
                cabecalhos.append(v.strftime("%d/%m"))
            elif not _is_nan(v) and str(v).strip() and str(v).strip() != " ":
                cabecalhos.append(str(v).strip())
            else:
                cabecalhos.append("")   # espaço em branco → coluna vazia

        dados = df.iloc[linha_cab + 1 : linha_fim].copy().reset_index(drop=True)
        if dados.empty:
            continue

        # Descartar col[0] se inteiramente NaN
        if dados.iloc[:, 0].isna().all():
            dados      = dados.iloc[:, 1:].copy().reset_index(drop=True)
            cabecalhos = cabecalhos[1:]

        # Alinhar colunas
        n_cab = len(cabecalhos)
        n_col = len(dados.columns)
        if n_col < n_cab:
            cabecalhos = cabecalhos[:n_col]
        elif n_col > n_cab:
            cabecalhos += [f"Col_{k}" for k in range(n_cab, n_col)]

        dados.columns = cabecalhos

        # Renomear primeira coluna se vazia
        first = dados.columns[0]
        if first in ("", " ") or first.startswith("Col_"):
            dados = dados.rename(columns={first: "Fazenda/Categoria"})

        # Remover colunas cujo cabeçalho é vazio (ex: coluna de espaço " ")
        cols_validas = [c for c in dados.columns if c.strip() not in ("", " ") or c == "Fazenda/Categoria"]
        dados = dados[cols_validas].copy()

        # Filtrar linhas: descartar NaN, linha de ano (ex: "2026")
        def _eh_util(v):
            s = str(v).strip()
            if _is_nan(v) or s in ("", "nan"):
                return False
            if s.isdigit() and len(s) == 4:
                return False
            return True

        mask_util = dados.iloc[:, 0].apply(_eh_util)
        dados = dados[mask_util.values].copy()

        if dados.empty:
            continue

        # Manter somente linhas com ao menos um valor numérico (exceto col[0])
        def _tem_numero(row):
            return any(
                isinstance(v, (int, float)) and not _is_nan(v)
                for v in row.tolist()[1:]
            )

        dados = dados[dados.apply(_tem_numero, axis=1)].copy().reset_index(drop=True)

        if dados.empty:
            continue

        # Selecionar colunas de data (excluir "Col_*" sem nome útil)
        id_cols   = [dados.columns[0]]
        date_cols = [
            c for c in dados.columns[1:]
            if c.strip() and not c.startswith("Col_") and c != "Fazenda/Categoria"
        ]
        sub = dados[id_cols + date_cols].copy()
        sub = _formatar_df(sub)
        sub = sub.replace("nan", "—")

        print(f"[LOG]   Bloco '{nome}': {len(sub)} linhas × {len(sub.columns)} colunas")
        resultados.append((f"Volume Real Diário — {nome}", sub))

    return resultados

def extrator_resumo_estoque(df: pd.DataFrame) -> list[tuple[str, pd.DataFrame]]:
    """
    Extrai os dois blocos da aba Resumo:
    Bloco 1 (cols H–N): Estoque Total por Classe TPC
    Bloco 2 (cols P–X): Estoque Total por Densidade
    Cabeçalho na linha 8 (índice 8), dados nas linhas 9–20.
    """
    resultados = []
    # Localiza a linha de cabeçalho pela presença de "Fazendas"
    header_row = None
    for i in range(len(df)):
        vals = [_str_val(v) for v in df.iloc[i].tolist()]
        if "Fazendas" in vals:
            header_row = i
            break

    if header_row is None:
        print("[AVISO] Cabeçalho não encontrado na aba 'Resumo'.")
        return resultados

    cab = df.iloc[header_row].tolist()
    dados = df.iloc[header_row + 1:].copy().reset_index(drop=True)

    # Detecta onde cada bloco começa (células com "Fazendas" no cabeçalho)
    blocos_ini = [i for i, v in enumerate(cab) if _str_val(v) == "Fazendas"]

    titulos_bloco = []
    for bi in blocos_ini:
        # Título real está 1 linha acima do cabeçalho (linha "Estoque Total / Classe TPC")
        t1 = _str_val(df.iloc[header_row - 1, bi]) if header_row >= 1 else ""
        t2 = _str_val(df.iloc[header_row - 1, bi + 1]) if header_row >= 1 else ""
        titulo = f"{t1} — {t2}".strip(" —") if t2 else t1
        titulos_bloco.append(titulo or f"Bloco col {bi}")

    for idx, bi in enumerate(blocos_ini):
        # Fim do bloco = início do próximo ou fim do df
        bf = blocos_ini[idx + 1] - 1 if idx + 1 < len(blocos_ini) else len(cab)
        cols_idx = list(range(bi, bf))

        cabecalhos = [_str_val(cab[i]) or f"Col_{i}" for i in cols_idx]
        sub = dados.iloc[:, cols_idx].copy()
        sub.columns = cabecalhos

        # Filtra linhas com valor na col "Fazendas" e exclui "Total Geral"
        mask = sub["Fazendas"].apply(
            lambda v: not _is_nan(v)
                    and str(v).strip() not in ("", "nan", "Total Geral")
        )
        sub = sub[mask.values].copy().reset_index(drop=True)

        # Adiciona linha de Total Geral ao final
        total_mask = dados.iloc[:, bi].apply(
            lambda v: _str_val(v) == "Total Geral"
        )
        if total_mask.any():
            total_row = dados[total_mask].iloc[[0], cols_idx].copy()
            total_row.columns = cabecalhos
            sub = pd.concat([sub, total_row], ignore_index=True)

        if sub.empty:
            continue

        sub = _formatar_df(sub)
        sub = sub.replace("nan", "—")
        print(f"[LOG]   Bloco '{titulos_bloco[idx]}': {len(sub)} linhas × {len(sub.columns)} colunas")
        resultados.append((f"Resumo — {titulos_bloco[idx]}", sub))

    return resultados

def extrator_meta_diaria_bi(df: pd.DataFrame) -> list[tuple[str, pd.DataFrame]]:
    resultados = []
    header_row = None
    for i, row in df.iterrows():
        vals = [_str_val(v) for v in row.tolist()]
        if "Frente" in vals and "FAZENDA" in vals:
            header_row = i
            break
    if header_row is None:
        return resultados

    cabecalhos = [_str_val(v) for v in df.iloc[header_row].tolist()]
    dados = df.iloc[header_row + 1:].copy()
    dados.columns = range(len(dados.columns))

    cols_premissas = ["Frente", "FAZENDA", "Produto", "Cx Carga", "DMT", "Asfalto", "Terra"]
    idx_premissas  = [i for i, c in enumerate(cabecalhos) if c in cols_premissas]

    if idx_premissas:
        sub = dados.iloc[:, idx_premissas].copy()
        sub.columns = [cabecalhos[i] for i in idx_premissas]
        sub = _limpar_df(sub)
        if not sub.empty:
            sub = _formatar_df(sub)
            resultados.append(("Meta Diária BI — Premissas por Fazenda", sub))

    idx_datas   = [i for i, c in enumerate(cabecalhos) if "/" in c and len(c) == 10][-15:]
    nomes_datas = [cabecalhos[i] for i in idx_datas]

    if idx_datas:
        idx_id  = [i for i, c in enumerate(cabecalhos) if c in ("FAZENDA", "Produto")]
        id_cols = [cabecalhos[i] for i in idx_id]
        sub = dados.iloc[:, idx_id + idx_datas].copy()
        sub.columns = id_cols + nomes_datas
        sub = _limpar_df(sub)
        sub = sub[sub.iloc[:, len(id_cols):].apply(
            lambda r: r.map(lambda v: not _is_nan(v) and str(v).strip() not in ("", "0")),
            axis=1,
        ).any(axis=1)]
        if not sub.empty:
            sub = _formatar_df(sub)
            resultados.append(("Meta Diária BI — Programação por Data (últimos 15 dias)", sub))

    return resultados


def extrator_forecast_volume(df: pd.DataFrame) -> list[tuple[str, pd.DataFrame]]:
    resultados = []
    header_row = None
    for i, row in df.iterrows():
        vals = [_str_val(v) for v in row.tolist()]
        if "Fazenda" in vals and "Real" in vals:
            header_row = i
            break
    if header_row is None:
        return resultados

    cab   = [_str_val(v) for v in df.iloc[header_row].tolist()]
    dados = df.iloc[header_row + 1:].copy()
    idx_validos = [i for i, c in enumerate(cab) if c and i < len(dados.columns)]
    sub = dados.iloc[:, idx_validos].copy()
    sub.columns = [cab[i] for i in idx_validos]
    sub = _limpar_df(sub)

    if "Fazenda" in sub.columns:
        mask = sub["Fazenda"].apply(
            lambda v: not _is_nan(v) and str(v).strip() not in ("", "nan")
        )
        sub = sub[mask.values]

    if not sub.empty:
        sub = _formatar_df(sub)
        resultados.append(("Forecast de Volume — Média Diária por Fazenda", sub))

    return resultados


def extrator_micro_logistica(df: pd.DataFrame) -> list[tuple[str, pd.DataFrame]]:
    resultados = []
    header_row = None
    for i, row in df.iterrows():
        vals = [_str_val(v) for v in row.tolist()]
        if "Frente" in vals and "Fazenda" in vals and "DMT" in vals:
            header_row = i
            break
    if header_row is None:
        return resultados

    cab   = [_str_val(v) for v in df.iloc[header_row].tolist()]
    dados = df.iloc[header_row + 1:].copy()
    dados.columns = range(len(dados.columns))

    cols_alvo = [
        "Frente", "Premissas", "Produto", "Fazenda",
        "Média", "DMT", "Terra", "Asfalto",
        ">=40 D", ">= 35 < 40", "< 35D", "Vol. Disp.",
    ]
    idx_map = {}
    for i, c in enumerate(cab):
        if c in cols_alvo and c not in idx_map:
            idx_map[c] = i

    idx_sel = list(idx_map.values())
    nomes   = list(idx_map.keys())
    if not idx_sel:
        return resultados

    sub = dados.iloc[:, idx_sel].copy()
    sub.columns = nomes
    sub = _limpar_df(sub)

    if "Premissas" in sub.columns:
        mask = sub["Premissas"].apply(
            lambda v: not _is_nan(v) and str(v).strip() not in ("", "nan")
        )
        sub = sub[mask.values]

    if not sub.empty:
        sub = _formatar_df(sub)
        resultados.append(("Micro Logística — Premissas de Meta por Fazenda", sub))

    return resultados


EXTRATORES = {
    "extrator_meta_diaria_bi":      extrator_meta_diaria_bi,
    "extrator_volume_real_diario":  extrator_volume_real_diario,
    "extrator_forecast_volume":     extrator_forecast_volume,
    "extrator_micro_logistica":     extrator_micro_logistica,
}


# ===========================================================================
# CARREGAMENTO
# ===========================================================================

def carregar_e_segmentar(path: str, config: list[dict]) -> list[tuple[str, pd.DataFrame]]:
    print(f"[LOG] Verificando existência do arquivo: {path}")
    if not os.path.exists(path):
        raise FileNotFoundError(
            f"Arquivo não encontrado: {path}\n"
            "Verifique se o caminho de rede está mapeado e acessível."
        )

    print(f"[LOG] Abrindo planilha...")
    todas: list[tuple[str, pd.DataFrame]] = []

    for cfg in config:
        if not cfg.get("ativo", True):
            print(f"[LOG] Aba desativada, pulando: {cfg['sheet']}")
            continue

        sheet         = cfg["sheet"]
        extrator_nome = cfg.get("extrator", "")
        fn_extrator   = EXTRATORES.get(extrator_nome)

        if fn_extrator is None:
            print(f"[AVISO] Extrator '{extrator_nome}' não encontrado. Aba ignorada.")
            continue

        print(f"[LOG] Lendo aba: '{sheet}'...")
        try:
            df = pd.read_excel(path, sheet_name=sheet, header=None, engine="openpyxl")
        except Exception as exc:
            print(f"[AVISO] Erro ao ler aba '{sheet}': {exc}. Pulando.")
            continue

        print(f"[LOG]   Aba carregada: {df.shape[0]} linhas × {df.shape[1]} colunas")
        subtabelas = fn_extrator(df)
        print(f"[LOG]   {len(subtabelas)} sub-tabela(s) extraída(s) de '{sheet}'")
        todas.extend(subtabelas)

    print(f"[LOG] Total de sub-tabelas: {len(todas)}")
    return todas


# ===========================================================================
# IMAGENS
# ===========================================================================

def gerar_imagem(df: pd.DataFrame, titulo: str, caminho: str) -> None:
    print(f"[LOG] Gerando imagem: '{titulo}'...")
    df = df.reset_index(drop=True)

    # Desduplicar colunas
    if df.columns.duplicated().any():
        seen: dict = {}
        novas = []
        for c in df.columns:
            if c in seen:
                seen[c] += 1
                novas.append(f"{c}.{seen[c]}")
            else:
                seen[c] = 0
                novas.append(c)
        df.columns = novas

    COR_HEADER  = "#1e5f3b"
    COR_STRIPE  = "#e8f5ee"
    COR_BORDER  = "#a8d5b5"

    estilo = (
        df.style
        .set_caption(titulo)
        .set_properties(**{
            "font-family": "Calibri, Arial, sans-serif",
            "font-size":   "13px",       # era 11px
            "text-align":  "center",
            "border":      f"1px solid {COR_BORDER}",
            "padding":     "7px 10px",   # era 5px 8px
            "white-space": "nowrap",
        })
        .set_table_styles([
            {
                "selector": "caption",
                "props": [
                    ("font-size", "12px"), ("font-weight", "bold"),
                    ("color", COR_HEADER), ("text-align", "left"),
                    ("padding-bottom", "4px"),
                    ("font-family", "Calibri, Arial, sans-serif"),
                ],
            },
            {
                "selector": "th",
                "props": [
                    ("background-color", COR_HEADER), ("color", "white"),
                    ("font-weight", "bold"), ("font-size", "13px"),
                    ("text-align", "center"), ("padding", "8px 10px"),
                    ("border", f"1px solid {COR_HEADER}"),
                ],
            },
            {
                "selector": "table",
                "props": [("border-collapse", "collapse"), ("width", "100%")],
            },
        ])
        .apply(
            lambda row: [
                f"background-color: {COR_STRIPE}" if row.name % 2 == 0 else ""
                for _ in row
            ],
            axis=1,
        )
        .hide(axis="index")
    )

    dfi.export(estilo, caminho, dpi=300, table_conversion="matplotlib", max_rows=-1)

    # Redimensionar se necessário
    with Image.open(caminho) as img:
        max_w = 640
        if img.width > max_w:
            scale    = max_w / img.width
            new_size = (max_w, int(img.height * scale))
            img      = img.resize(new_size, Image.LANCZOS)
            img.save(caminho, format="PNG", optimize=True)

    print(f"[LOG]   Imagem salva: {caminho}")


def gerar_todas_imagens(
    tabelas: list[tuple[str, pd.DataFrame]],
    base: str,
) -> list[tuple[str, str]]:
    resultado = []
    for idx, (titulo, df) in enumerate(tabelas, start=1):
        nome_seguro = "".join(c if c.isalnum() or c in "-_" else "_" for c in titulo)
        caminho     = f"{base}_{idx:02d}_{nome_seguro}.png"
        gerar_imagem(df, titulo, caminho)
        resultado.append((titulo, caminho))
    return resultado


# ===========================================================================
# E-MAIL HTML
# ===========================================================================

def montar_email_html(
    assunto, saudacao, texto_corpo, data_atual,
    assinatura_nome, assinatura_cargo, assinatura_email_sig, assinatura_tel,
    imagens, caminho_anexo,
):
    caminhos   = [c for _, c in imagens]
    blocos_img = ""
    for idx, (titulo, _) in enumerate(imagens, start=1):
        blocos_img += f"""
          <div style="margin-bottom:16px;">
            <p style="margin:0 0 4px;font-size:12px;font-weight:700;color:#1e5f3b;">{titulo}</p>
            <img src="cid:tabela_{idx}" alt="{titulo}"
                style="max-width:100%;height:auto;border:1px solid #c8e6d4;
                        border-radius:3px;display:block;" />
          </div>"""

    html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head><meta charset="UTF-8"><title>{assunto}</title></head>
<body style="margin:0;padding:0;background-color:#f4f4f4;font-family:Calibri,Arial,sans-serif;">

  <table width="100%" cellpadding="0" cellspacing="0" style="background-color:#f4f4f4;padding:24px 0;">
    <tr><td align="center">
      <table width="680" cellpadding="0" cellspacing="0"
             style="background-color:#ffffff;border-radius:6px;
                    border:1px solid #dde8e2;overflow:hidden;">

        <!-- CABEÇALHO -->
        <tr>
          <td style="background-color:#1e5f3b;padding:14px 24px;">
            <table width="100%" cellpadding="0" cellspacing="0"><tr>
              <td>
                <span style="font-size:15px;font-weight:700;color:#ffffff;">Eldorado Brasil</span>
                <span style="font-size:11px;color:#a8d5b5;margin-left:8px;">· Estoque Diário de Madeira/TPC</span>
              </td>
              <td align="right">
                <span style="font-size:11px;color:#a8d5b5;">{data_atual}</span>
              </td>
            </tr></table>
          </td>
        </tr>

        <!-- CORPO -->
        <tr>
          <td style="padding:20px 24px 12px;">
            <p style="margin:0 0 6px;font-size:14px;color:#333333;">{saudacao}</p>
            <p style="margin:0 0 16px;font-size:13px;color:#555555;line-height:1.6;">{texto_corpo}</p>
          </td>
        </tr>

        <!-- TABELAS -->
        <tr>
          <td style="padding:0 24px 16px;">{blocos_img}</td>
        </tr>

        <!-- DIVISÓRIA -->
        <tr>
          <td style="padding:0 24px;">
            <hr style="border:none;border-top:1px solid #d4e8db;margin:0;">
          </td>
        </tr>

        <!-- ASSINATURA -->
        <tr>
        <td style="padding:14px 24px 20px;">
            
            <table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;">
            <tr>

                <!-- COLUNA ESQUERDA -->
                <td width="65%" style="vertical-align:top;padding-right:28px;">

                <div style="font-size:14px;font-weight:700;color:#1e5f3b;line-height:18px;">
                    {assinatura_nome}
                </div>

                <div style="font-size:12px;color:#555555;line-height:18px;">
                    {assinatura_cargo}
                </div>

                <div style="height:10px;"></div>

                <div style="font-size:12px;line-height:20px;">
                    <a href="mailto:{assinatura_email_sig}" 
                    style="color:#1e5f3b;text-decoration:underline;">
                    {assinatura_email_sig}
                    </a>
                </div>

                <div style="font-size:12px;line-height:20px;">
                    <a href="https://eldoradobrasil.com.br"
                    style="color:#1e5f3b;text-decoration:underline;">
                    eldoradobrasil.com.br
                    </a>
                </div>

                <div style="font-size:12px;line-height:20px;">
                    <a href="https://linhaetica.eldoradobrasil.com.br"
                    style="color:#1e5f3b;text-decoration:underline;">
                    linhaetica.eldoradobrasil.com.br
                    </a>
                </div>

                </td>

                <!-- DIVISÓRIA -->
                <td width="1" 
                    style="background:#d4e8db;font-size:0;line-height:0;">
                &nbsp;
                </td>

                <!-- COLUNA DIREITA -->
                <td width="35%" style="vertical-align:top;padding-left:28px;">

                <div style="font-size:12px;color:#555555;line-height:18px;">
                    Rod. Br 158 – Km 231
                </div>

                <div style="font-size:12px;color:#555555;line-height:18px;">
                    Três Lagoas - MS - Brasil
                </div>

                <div style="font-size:12px;color:#555555;line-height:18px;">
                    CEP: 79641-300
                </div>

                <div style="height:6px;"></div>

                <div style="font-size:12px;color:#555555;line-height:18px;">
                    +55 (17) 9 9767-4167
                </div>

                </td>

            </tr>
            </table>

        </td>
        </tr>


        <!-- RODAPÉ -->
        <tr>
          <td style="background-color:#f0f7f3;padding:10px 24px;border-top:1px solid #dde8e2;text-align:center;">
            <span style="font-size:10px;color:#7a9e82;">
              E-mail gerado automaticamente. Por favor, não responda diretamente.
            </span>
          </td>
        </tr>

      </table>
    </td></tr>
  </table>

</body></html>"""

    return html, caminhos, []


# ===========================================================================
# ENVIO OUTLOOK
# ===========================================================================

def enviar_email(html_body, assunto, email_remetente, destinatarios, imagem_paths, anexos):
    try:
        import win32com.client
    except ImportError as exc:
        raise RuntimeError(
            "PyWin32 é necessário. Instale com: pip install pywin32"
        ) from exc

    print("[LOG] Criando e-mail no Outlook...")
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail    = outlook.CreateItem(0)
    mail.Subject  = assunto
    mail.To       = ";".join(destinatarios)
    mail.HTMLBody = html_body

    if email_remetente:
        try:
            mail.SentOnBehalfOfName = email_remetente
        except Exception:
            print("[AVISO] Não foi possível configurar SentOnBehalfOfName.")

    for idx, caminho in enumerate(imagem_paths, start=1):
        if os.path.exists(caminho):
            att = mail.Attachments.Add(os.path.abspath(caminho))
            try:
                att.PropertyAccessor.SetProperty(
                    "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
                    f"tabela_{idx}",
                )
            except Exception:
                print(f"[AVISO] Não foi possível definir CID para: {caminho}")
        else:
            print(f"[AVISO] Imagem não encontrada: {caminho}")

    for caminho in anexos:
        if os.path.exists(caminho):
            mail.Attachments.Add(os.path.abspath(caminho))
        else:
            print(f"[AVISO] Anexo não encontrado: {caminho}")

    mail.Send()
    print(f"[LOG] E-mail enviado para: {', '.join(destinatarios)}")


# ===========================================================================
# LIMPEZA
# ===========================================================================

def limpar_temporarios(caminhos: list[str]) -> None:
    for c in caminhos:
        try:
            if os.path.exists(c):
                os.remove(c)
                print(f"[LOG] Temporário removido: {c}")
        except OSError as exc:
            print(f"[AVISO] Não foi possível remover {c}: {exc}")


# ===========================================================================
# EXECUÇÃO PRINCIPAL
# ===========================================================================

def executar() -> None:
    inicio = time.time()
    print("=" * 60)
    print("  SISTEMA DE ENVIO AUTOMÁTICO — REPORT OPERACIONAL ELDORADO")
    print(f"  Início: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    print("=" * 60)

    data_atual = datetime.now().strftime("%d/%m/%Y")
    data_curta = datetime.now().strftime("%d/%m")
    assunto    = f"Estoque Diário de Madeira/TPC — {data_curta}"
    imagem_paths: list[str] = []

    try:
        if not EMAIL_REMETENTE:
            raise RuntimeError(
                "EMAIL_REMETENTE não definido no .env."
            )
        if not DESTINATARIOS:
            raise RuntimeError(
                "DESTINATARIOS não definido no .env (separe por vírgula)."
            )
        
        print(f"[LOG] Destinatários configurados ({len(DESTINATARIOS)}):")
        for d in DESTINATARIOS:
            print(f"       → {d}")

        if EXCEL_PATH:
            if not _refresh_excel(EXCEL_PATH):
                print("[AVISO] Não foi possível atualizar a planilha antes da leitura.")

        tabelas = carregar_e_segmentar(EXCEL_PATH, SHEET_CONFIG)
        if not tabelas:
            raise RuntimeError("Nenhuma sub-tabela válida encontrada.")

        imagens      = gerar_todas_imagens(tabelas, IMAGE_BASE)
        imagem_paths = [c for _, c in imagens]

        html_body, imagem_paths, anexos = montar_email_html(
            assunto              = assunto,
            saudacao             = SAUDACAO,
            texto_corpo          = TEXTO_CORPO,
            data_atual           = data_atual,
            assinatura_nome      = ASSINATURA_NOME,
            assinatura_cargo     = ASSINATURA_CARGO,
            assinatura_email_sig = ASSINATURA_EMAIL,
            assinatura_tel       = ASSINATURA_TEL,
            imagens              = imagens,
            caminho_anexo        = EXCEL_PATH if ANEXAR_PLANILHA else "",
        )
        print("[LOG] E-mail montado com sucesso.")

        enviar_email(
            html_body       = html_body,
            assunto         = assunto,
            email_remetente = EMAIL_REMETENTE,
            destinatarios   = DESTINATARIOS,
            imagem_paths    = imagem_paths,
            anexos          = anexos,
        )

    except FileNotFoundError as exc:
        print(f"\n[ERRO] {exc}")
        traceback.print_exc()
    except RuntimeError as exc:
        print(f"\n[ERRO] {exc}")
        traceback.print_exc()
    except Exception as exc:
        print(f"\n[ERRO INESPERADO] {exc}")
        traceback.print_exc()
    finally:
        limpar_temporarios(imagem_paths)
        print(f"\n[LOG] Duração total: {time.time() - inicio:.1f}s")
        print("=" * 60)
EXTRATORES = {
    "extrator_meta_diaria_bi":      extrator_meta_diaria_bi,
    "extrator_volume_real_diario":  extrator_volume_real_diario,
    "extrator_forecast_volume":     extrator_forecast_volume,
    "extrator_micro_logistica":     extrator_micro_logistica,
    "extrator_resumo_estoque":      extrator_resumo_estoque,
}

if __name__ == "__main__":
    executar()