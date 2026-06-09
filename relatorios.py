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
EXCEL_PATH = os.getenv(
    "EXCEL_PATH",
    r"\\brtlgwvfs01eld\Florestal\015_Logística Florestal e Infra Estrutura"
    r"\002 - Logistica Florestal\01. Base de Dados\02. Manuais\09. ChatBot"
    r"\03. Planilha de envio\Relatório de Produção - Logística Florestal.xlsx"
)

EMAIL_REMETENTE  = os.getenv("EMAIL_REMETENTE", "")
ASSINATURA_NOME  = os.getenv("ASSINATURA_NOME", "")
ASSINATURA_CARGO = os.getenv("ASSINATURA_CARGO", "")
ASSINATURA_TEL   = os.getenv("ASSINATURA_TEL", "")
ASSINATURA_EMAIL = EMAIL_REMETENTE
ANEXAR_PLANILHA  = os.getenv("ANEXAR_PLANILHA", "true").lower() == "true"
ANEXAR_GRAFICOS  = os.getenv("ANEXAR_GRAFICOS", "true").lower() == "true"

_dest_raw    = os.getenv("DESTINATARIOS", "")
DESTINATARIOS = [e.strip() for e in _dest_raw.split(",") if e.strip()]

ASSUNTO     = "Relatório de Produção - Logística Florestal"
SAUDACAO    = "Prezados,"
TEXTO_CORPO = (
    "Segue em anexo o <strong>Relatório de Produção</strong> da Eldorado Brasil, "
    "com os dados consolidados do período."
)

IMAGE_BASE = "report_temp"

SHEET_CONFIG: list[dict] = [
    {
        "sheet":    "Tabela - Produção Diária",
        "titulo":   "Produção Diária",
        "extrator": "extrator_subtabelas_por_bloques",
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


def _is_blank(v) -> bool:
    if v is None:
        return True
    if isinstance(v, str) and v.strip() == "":
        return True
    return _is_nan(v)


def _split_header_segments(values: list, max_gap: int = 3) -> list[tuple[int, int]]:
    segments = []
    start = None
    gap = 0

    for idx, value in enumerate(values):
        blank = _is_blank(value)
        if not blank:
            if start is None:
                start = idx
            gap = 0
        elif start is not None:
            gap += 1
            if gap > max_gap:
                end = idx - gap
                if end >= start:
                    segments.append((start, end))
                start = None
                gap = 0

    if start is not None:
        segments.append((start, len(values) - 1))

    return segments


def _extract_segment_title(df: pd.DataFrame, header_row: int, start_col: int, end_col: int) -> str:
    for row_idx in range(header_row - 1, max(header_row - 4, -1), -1):
        linha = df.iloc[row_idx, start_col:end_col + 1].tolist()
        textos = [str(v).strip() for v in linha if not _is_blank(v)]
        if textos:
            for keyword in (
                "Logística de Celulose",
                "Logistica de Celulose",
                "Logística de Biomassa",
                "Logistica de Biomassa",
                "Report Operacional - Celulose",
                "Report Operacional - Biomassa",
            ):
                for texto in textos:
                    if keyword.lower() in texto.lower():
                        return texto
            for texto in reversed(textos):
                if any(k in texto for k in ("Celulose", "Biomassa")):
                    return texto
            return textos[-1]
    return ""


def _find_segment_end(df: pd.DataFrame, header_row: int, start_col: int, end_col: int) -> int:
    row_end = header_row + 1
    while row_end < len(df):
        linha = df.iloc[row_end, start_col:end_col + 1].tolist()
        if all(_is_blank(v) for v in linha):
            break
        if any(str(v).strip() == "Fazenda" for v in linha if not _is_blank(v)):
            break
        row_end += 1
    return row_end


def _exportar_graficos(path: str, sheet_name: str, base_name: str = "chart") -> list[tuple[str, str]]:
    imagens = []
    try:
        import xlwings as xw
    except ImportError:
        print("[AVISO] Biblioteca 'xlwings' não instalada. Ignorando exportação de gráficos.")
        return imagens

    try:
        app = xw.App(visible=False, add_book=False)
        wb = app.books.open(str(path))
        sht = wb.sheets[sheet_name]
        
        try:
            for slicer in sht.api.Slicers:
                slicer.SlicerItems.ClearAll()
        except Exception:
            pass
        
        count = sht.api.ChartObjects().Count
        if count <= 0:
            return imagens

        print(f"[LOG] Exportando {count} gráfico(s) de '{sheet_name}'...")
        for idx in range(1, count + 1):
            chart_obj = sht.api.ChartObjects(idx)
            nome = f"{base_name}_{sheet_name}_{idx:02d}.png"
            nome = "".join(c if c.isalnum() or c in "-_." else "_" for c in nome)
            caminho_absoluto = os.path.abspath(nome)
            exportou = chart_obj.Chart.Export(str(caminho_absoluto), "PNG")
            if exportou and os.path.exists(caminho_absoluto):
                titulo = (
                    "Grafico de Celulose: Report Operacional - Celulose"
                    if chart_obj.Left < 700
                    else "Grafico de Biomassa: Report Operacional - Biomassa"
                )
                imagens.append((titulo, caminho_absoluto))
                print(f"[LOG]   Gráfico exportado: {caminho_absoluto} ({titulo})")
            else:
                print(f"[AVISO] Não exportou gráfico {idx} de '{sheet_name}'.")

    except Exception as exc:
        print(f"[AVISO] Falha ao exportar gráficos de '{sheet_name}': {exc}")
    finally:
        try:
            if wb is not None:
                wb.close(save_changes=False)
        except Exception:
            pass
        try:
            if app is not None:
                app.quit()
        except Exception:
            pass

    return imagens


def _extrair_data_coleta(path: str, sheet_name: str) -> str:
    """Extrai a data de coleta da planilha."""
    try:
        df = pd.read_excel(path, sheet_name=sheet_name, header=None)
        for idx, row in df.iterrows():
            for val in row:
                if pd.notna(val):
                    val_str = str(val).strip()
                    try:
                        dt = pd.to_datetime(val_str, dayfirst=True)
                        if pd.Timestamp.now().year - 2 <= dt.year <= pd.Timestamp.now().year + 1:
                            return dt.strftime("%d/%m/%Y")
                    except (ValueError, TypeError):
                        pass
    except Exception as exc:
        print(f"[AVISO] Falha ao extrair data: {exc}")
    return ""


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

def extrator_producao_celulose_biomassa(df: pd.DataFrame, titulo_aba: str = "") -> list[tuple[str, pd.DataFrame]]:
    """
    Extrator legado para planilhas de Celulose e Biomassa em formato antigo.
    Detecta cabeçalho em linha 2 (F_Cubo415[...]) e extrai dados agregados.
    """
    resultados = []
    
    # Procurar pela linha de cabeçalho
    header_row = None
    for i in range(min(5, len(df))):
        vals = [str(v).strip() for v in df.iloc[i].tolist()]
        if any("F_Cubo415" in v for v in vals):
            header_row = i
            break
    
    if header_row is None:
        print(f"[AVISO] Cabeçalho F_Cubo415 não encontrado em '{titulo_aba}'.")
        return resultados
    
    # Usar a linha de cabeçalho
    cabecalhos_raw = [str(v).strip() for v in df.iloc[header_row].tolist()]
    
    # Extrair dados (começar depois do cabeçalho)
    dados = df.iloc[header_row + 1:].copy()
    dados.columns = range(len(dados.columns))
    
    # Colunas-alvo: Nome da Origem, Fazenda/Projeto, Produto, Volume, DMT
    cols_alvo = {
        "Nome da Origem": None,
        "Fazenda/Projeto": None,
        "Dcr Produto": None,
        "Vol": None,
        "DMT x Vol": None,
        "RPV x Vol": None,
    }
    
    # Localizar índices das colunas-alvo
    for idx, cab in enumerate(cabecalhos_raw):
        for col_alvo in cols_alvo:
            if col_alvo in cab:
                cols_alvo[col_alvo] = idx
                break
    
    # Selecionar apenas colunas encontradas
    indices = [i for i in cols_alvo.values() if i is not None]
    nomes = [k for k, i in cols_alvo.items() if i is not None]
    
    if not indices:
        print(f"[AVISO] Nenhuma coluna-alvo encontrada em '{titulo_aba}'.")
        return resultados
    
    sub = dados.iloc[:, indices].copy()
    sub.columns = nomes
    sub = _limpar_df(sub)
    
    # Remover linhas vazias/inválidas
    if not sub.empty:
        sub = sub[sub.apply(lambda row: not all(pd.isna(v) or str(v).strip() in ("", "nan") for v in row), axis=1)]
    
    if not sub.empty:
        sub = _formatar_df(sub)
        titulo_final = titulo_aba or "Produção"
        resultados.append((titulo_final, sub))
    
    return resultados


def extrator_subtabelas_por_bloques(df: pd.DataFrame, titulo_aba: str = "") -> list[tuple[str, pd.DataFrame]]:
    """
    Extrai múltiplas subtabelas de uma mesma aba de Excel.
    Detecta cabeçalhos repetidos na planilha e segmenta blocos horizontais
    para gerar tabelas independentes.
    """
    resultados = []
    header_rows = []

    for i in range(len(df)):
        valores = [str(v).strip() for v in df.iloc[i].tolist()]
        if "Fazenda" in valores and any("Meta" in v or "Vol." in v or "DMT" in v or "RPV" in v for v in valores if v):
            header_rows.append(i)

    if not header_rows:
        print(f"[AVISO] Nenhum cabeçalho de subtabela encontrado em '{titulo_aba}'.")
        return resultados

    for header_row in header_rows:
        segmentos = _split_header_segments(df.iloc[header_row].tolist(), max_gap=3)
        for start_col, end_col in segmentos:
            cabecalhos = [ _str_val(v) for v in df.iloc[header_row, start_col:end_col + 1].tolist() ]
            if not any(c for c in cabecalhos if c):
                continue
            if not any("Fazenda" in c or "Meta" in c or "Vol" in c or "DMT" in c or "RPV" in c for c in cabecalhos if c):
                continue

            row_end = _find_segment_end(df, header_row, start_col, end_col)
            sub = df.iloc[header_row + 1:row_end, start_col:end_col + 1].copy()
            sub.columns = cabecalhos
            sub = _limpar_df(sub)

            if sub.empty:
                continue

            # Excluir linhas que sejam totalmente vazias ou somente com identificadores inválidos
            mask = sub.apply(lambda row: not all(
                _is_blank(v) or str(v).strip().lower() in ("", "nan")
                for v in row
            ), axis=1)
            sub = sub[mask].reset_index(drop=True)

            if sub.empty:
                continue

            sub = _formatar_df(sub)
            titulo_extra = _extract_segment_title(df, header_row, start_col, end_col)
            titulo_final = titulo_extra or f"{titulo_aba} — Subtabela {len(resultados)+1}"
            resultados.append((titulo_final, sub))

    return resultados


def extrator_volume_real_diario(df: pd.DataFrame) -> list[tuple[str, pd.DataFrame]]:
    """
    Extrator específico para a aba 'Volume Real Diário'.
    Detecta blocos onde:
      - col[1] = nome do bloco
      - col[2] = 'Rótulos de Coluna'
    """

    resultados = []
    n_rows = len(df)
    blocos = []

    # Detectar blocos
    for i in range(n_rows - 1):
        v1 = str(df.iloc[i, 1]).strip() if df.shape[1] > 1 else ""
        v2 = str(df.iloc[i, 2]).strip() if df.shape[1] > 2 else ""

        if v1 and v1 not in ("nan", " ") and "Rótulos" in v2:
            blocos.append((i, v1))

    if not blocos:
        print("[AVISO] Nenhum bloco detectado na aba 'Volume Real Diário'.")
        return resultados

    print(f"[LOG]   Blocos detectados: {[n for _, n in blocos]}")

    for b_idx, (linha_titulo, nome) in enumerate(blocos):

        linha_fim = (
            blocos[b_idx + 1][0]
            if b_idx + 1 < len(blocos)
            else n_rows
        )

        linha_cab = linha_titulo + 1

        # Cabeçalhos
        cab_row = df.iloc[linha_cab].tolist()

        cabecalhos = ["Fazenda/Categoria"]

        for v in cab_row[2:]:
            if hasattr(v, "strftime"):
                cabecalhos.append(v.strftime("%d/%m"))

            elif not _is_nan(v) and str(v).strip() not in ("", "nan"):
                cabecalhos.append(str(v).strip())

            else:
                cabecalhos.append(None)

        # Pular linha do ano (2026)
        inicio_dados = linha_cab + 2

        dados_raw = (
            df.iloc[inicio_dados:linha_fim]
            .copy()
            .reset_index(drop=True)
        )

        if dados_raw.empty:
            continue

        nomes_col = dados_raw.iloc[:, 1].tolist()
        valores = dados_raw.iloc[:, 2:].values.tolist()

        linhas_ok = []

        for nm, vrow in zip(nomes_col, valores):

            s = str(nm).strip()

            if _is_nan(nm) or s in ("", "nan"):
                continue

            if s.isdigit() and len(s) == 4:
                continue

            linhas_ok.append((s, vrow))

        if not linhas_ok:
            continue

        idx_validos = [
            i
            for i, c in enumerate(cabecalhos[1:])
            if c is not None
        ]

        cols_finais = (
            ["Fazenda/Categoria"]
            + [cabecalhos[1:][i] for i in idx_validos]
        )

        rows_finais = []

        for nm, vrow in linhas_ok:

            vals_selecionados = [
                vrow[i] if i < len(vrow) else np.nan
                for i in idx_validos
            ]

            tem_num = any(
                isinstance(v, (int, float))
                and not _is_nan(v)
                for v in vals_selecionados
            )

            if not tem_num:
                continue

            row_fmt = [nm] + [
                _fmt_num(v)
                for v in vals_selecionados
            ]

            rows_finais.append(row_fmt)

        
        if not rows_finais:
            continue

        sub = pd.DataFrame(
            rows_finais,
            columns=cols_finais
        )

        # Mantém Fazenda/Categoria + até os últimos 7 dias
        if len(sub.columns) > 1:
            qtd_dias = min(7, len(sub.columns) - 1)

            colunas = [sub.columns[0]] + list(sub.columns[-qtd_dias:])
            sub = sub[colunas]

        print(
            f"[LOG]   Bloco '{nome}': "
            f"{sub.shape[0]} linhas × "
            f"{sub.shape[1]} colunas"
        )

        resultados.append(
            (
                f"Volume Real Diário — {nome}",
                sub
            )
        )

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
    "extrator_producao_celulose_biomassa": extrator_producao_celulose_biomassa,
    "extrator_subtabelas_por_bloques":     extrator_subtabelas_por_bloques,
    "extrator_meta_diaria_bi":             extrator_meta_diaria_bi,
    "extrator_volume_real_diario":         extrator_volume_real_diario,
    "extrator_forecast_volume":            extrator_forecast_volume,
    "extrator_micro_logistica":            extrator_micro_logistica,
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
        try:
            subtabelas = fn_extrator(df, cfg.get("titulo", sheet))
        except TypeError:
            subtabelas = fn_extrator(df)
        print(f"[LOG]   {len(subtabelas)} sub-tabela(s) extraída(s) de '{sheet}'")
        todas.extend(subtabelas)

    print(f"[LOG] Total de sub-tabelas: {len(todas)}")
    return todas


# ===========================================================================
# IMAGENS
# ===========================================================================

def gerar_imagem(df: pd.DataFrame, titulo: str, caminho: str) -> None:
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

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

    COR_HEADER = "#1e5f3b"
    COR_STRIPE = "#e8f5ee"
    COR_BRANCO = "#ffffff"
    COR_BORDER = "#a8d5b5"

    n_rows, n_cols = df.shape
    col_labels = list(df.columns)
    cell_data  = df.values.tolist()

    # --- Calcular largura de cada coluna pelo conteúdo ---
    CHAR_W   = 0.10   # caracteres mais compactos
    MIN_W    = 0.45   # mínimo (colunas de data tipo "01/05" ficam ~0.7")
    MAX_W    = 4.0    # teto maior para não comprimir colunas
    PAD      = 0.20   # padding lateral

    col_widths = []
    for ci in range(n_cols):
        header_len = len(str(col_labels[ci]))
        data_len   = max((len(str(row[ci])) for row in cell_data), default=0)
        best       = max(header_len, data_len)
        w          = min(MAX_W, max(MIN_W, best * CHAR_W + PAD))
        col_widths.append(w)

    fig_w  = sum(col_widths) + 0.5          # margem lateral maior
    ROW_H  = 0.40   # altura para melhor legibilidade
    fig_h  = (n_rows + 2) * ROW_H + 0.6    # +2 = header + título

    fig, ax = plt.subplots(figsize=(fig_w, fig_h))
    fig.patch.set_facecolor(COR_BRANCO)     # ← fundo BRANCO (sem transparência)
    ax.set_facecolor(COR_BRANCO)
    ax.axis("off")

    tbl = ax.table(
        cellText=cell_data,
        colLabels=col_labels,
        cellLoc="center",
        loc="center",
        colWidths=[w / fig_w for w in col_widths],  # frações da figura
    )
    tbl.auto_set_font_size(False)
    tbl.set_fontsize(11)   # fonte maior

    # Estilo cabeçalho
    for ci in range(n_cols):
        cell = tbl[0, ci]
        cell.set_facecolor(COR_HEADER)
        cell.set_text_props(color="white", fontweight="bold")
        cell.set_edgecolor(COR_BORDER)
        cell.set_height(ROW_H / fig_h)

    # Estilo linhas de dados
    for ri in range(1, n_rows + 1):
        bg = COR_STRIPE if ri % 2 == 1 else COR_BRANCO
        for ci in range(n_cols):
            cell = tbl[ri, ci]
            cell.set_facecolor(bg)
            cell.set_edgecolor(COR_BORDER)
            cell.set_text_props(color="#333333")
            cell.set_height(ROW_H / fig_h)

    

    plt.tight_layout()
    plt.savefig(
        caminho, dpi=160,
        bbox_inches="tight",
        facecolor=COR_BRANCO,   # ← garante branco no PNG final
        format="PNG",
    )
    plt.close(fig)
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
    assunto, saudacao, texto_corpo, data_atual, data_coleta,
    assinatura_nome, assinatura_cargo, assinatura_email_sig, assinatura_tel,
    imagens, graficos, caminho_anexo, tem_graficos=False,
):
    blocos_intercalados = ""
    
    for idx in range(max(len(graficos), len(imagens))):
        if idx < len(graficos):
            titulo_g, _ = graficos[idx]
            blocos_intercalados += f"""
            <div style=\"margin-bottom:24px;text-align:left;\">
                <p style=\"margin:0 0 6px;font-size:13px;font-weight:700;
                        color:#1e5f3b;\">{titulo_g}</p>
                <img src=\"cid:grafico_{idx+1}\" alt=\"{titulo_g}\"
                width=\"750\"
                style=\"width:750px;height:auto;
                        border:1px solid #a8d5b5;border-radius:4px;
                        display:block;\" />
            </div>"""
        
        if idx < len(imagens):
            titulo_i, _ = imagens[idx]
            blocos_intercalados += f"""
            <div style=\"margin-bottom:24px;text-align:left;\">
                <p style=\"margin:0 0 6px;font-size:13px;font-weight:700;
                        color:#1e5f3b;\">{titulo_i}</p>
                <img src=\"cid:tabela_{idx+1}\" alt=\"{titulo_i}\"
                width=\"700\"
                style=\"width:700px;height:auto;
                        border:1px solid #a8d5b5;border-radius:4px;
                        display:block;\" />
            </div>"""

    html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head><meta charset="UTF-8"><title>{assunto}</title></head>
<body style="margin:0;padding:0;background-color:#f0f7f3;font-family:Calibri,Arial,sans-serif;">
  <table width="100%" cellpadding="0" cellspacing="0" style="background-color:#f0f7f3;padding:30px 0;">
    <tr><td align="center">
      <table width="860" cellpadding="0" cellspacing="0"
             style="background-color:#ffffff;border-radius:8px;
                    box-shadow:0 2px 8px rgba(0,0,0,0.10);overflow:hidden;">
        <tr>
          <td style="background-color:#1e5f3b;padding:22px 32px;">
            <table width="100%" cellpadding="0" cellspacing="0"><tr>
              <td>
                <span style="font-size:20px;font-weight:700;color:#ffffff;">Eldorado Brasil</span><br>
                <span style="font-size:13px;color:#a8d5b5;">Relatório de Produção - Logística Florestal</span>
              </td>
              <td align="right"><span style="font-size:12px;color:#a8d5b5;">{data_atual}</span></td>
            </tr></table>
          </td>
        </tr>
        
        <tr>
          <td style="padding:28px 32px 18px;">
            <p style="margin:0 0 14px;font-size:15px;color:#333333;">{saudacao}</p>
            <p style="margin:0 0 22px;font-size:14px;color:#555555;line-height:1.65;">{texto_corpo}</p>
            <p style="margin:0 0 4px;font-size:12px;font-weight:700;color:#1e5f3b;text-transform:uppercase;">
              Indicadores Operacionais
            </p>
          </td>
        </tr>
        <tr>
        <td style="padding:0;">
            <hr style="border:none;border-top:1px solid #c8e6d4;margin:0;">
        </td>
        </tr>
        {f'<tr><td style="padding:0 32px 12px;"><p style="margin:0;font-size:11px;color:#888888;font-style:italic;">Dados coletados em: {data_coleta}</p></td></tr>' if data_coleta else ''}
            <tr><td style="padding:0 32px 18px;">{blocos_intercalados}</td></tr>

        <tr>
            <td style="padding:0;">
                <hr style="border:none;border-top:1px solid #c8e6d4;margin:0;">
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
                    {assinatura_tel}
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

    return html, [], []


# ===========================================================================
# ENVIO OUTLOOK
# ===========================================================================

def _converter_imagem_base64(caminho: str) -> str:
    """Converte imagem para base64 para embedding direto no HTML."""
    try:
        with open(caminho, "rb") as f:
            import base64
            data = base64.b64encode(f.read()).decode("utf-8")
            return f"data:image/png;base64,{data}"
    except Exception as exc:
        print(f"[AVISO] Falha ao converter imagem para base64: {exc}")
        return ""


def enviar_email(html_body, assunto, email_remetente, destinatarios, imagem_paths, num_tabela_images, anexos):
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

    data_atual    = datetime.now().strftime("%d/%m/%Y")
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

        temp_paths: list[str] = []
        chart_items: list[tuple[str, str]] = []

        if EXCEL_PATH:
            if not _refresh_excel(EXCEL_PATH):
                print("[AVISO] Não foi possível atualizar a planilha antes da leitura.")

        if ANEXAR_GRAFICOS:
            for cfg in SHEET_CONFIG:
                if cfg.get("ativo", True):
                    chart_items.extend(_exportar_graficos(EXCEL_PATH, cfg["sheet"], base_name=IMAGE_BASE))

        tabelas = carregar_e_segmentar(EXCEL_PATH, SHEET_CONFIG)
        if not tabelas:
            raise RuntimeError("Nenhuma sub-tabela válida encontrada.")

        imagens      = gerar_todas_imagens(tabelas, IMAGE_BASE)
        imagem_paths = [c for _, c in imagens]
        chart_paths  = [c for _, c in chart_items]
        temp_paths   = imagem_paths + chart_paths
        
        data_coleta = _extrair_data_coleta(EXCEL_PATH, "Tabela - Produção Diária")

        html_body, imagem_paths, anexos = montar_email_html(
            assunto              = ASSUNTO,
            saudacao             = SAUDACAO,
            texto_corpo          = TEXTO_CORPO,
            data_atual           = data_atual,
            data_coleta          = data_coleta,
            assinatura_nome      = ASSINATURA_NOME,
            assinatura_cargo     = ASSINATURA_CARGO,
            assinatura_email_sig = ASSINATURA_EMAIL,
            assinatura_tel       = ASSINATURA_TEL,
            imagens              = imagens,
            graficos             = chart_items,
            caminho_anexo        = EXCEL_PATH if ANEXAR_PLANILHA else "",
            tem_graficos         = bool(chart_items),
        )
        print("[LOG] E-mail montado com sucesso.")

        anexos = []
        if ANEXAR_PLANILHA:
            anexos.append(EXCEL_PATH)

        enviar_email(
            html_body          = html_body,
            assunto            = ASSUNTO,
            email_remetente    = EMAIL_REMETENTE,
            destinatarios      = DESTINATARIOS,
            imagem_paths       = imagem_paths + chart_paths,
            num_tabela_images  = len(imagens),
            anexos             = anexos,
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
        limpar_temporarios(temp_paths if 'temp_paths' in locals() else imagem_paths)
        print(f"\n[LOG] Duração total: {time.time() - inicio:.1f}s")
        print("=" * 60)


if __name__ == "__main__":
    executar()