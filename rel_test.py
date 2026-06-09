"""
=============================================================================
SISTEMA DE AUTOMAÇÃO CORPORATIVA — ENVIO DE REPORT OPERACIONAL POR E-MAIL
=============================================================================
Autor      : Matheus Menezes
Empresa    : Eldorado Brasil
Versão     : 3.4.0

Mudanças v3.4.0:
  - Ranges com linha final DINÂMICA: a última linha preenchida é detectada
    automaticamente via xlwings — sem necessidade de ajuste manual quando
    a planilha crescer.
  - Configuração centralizada em SHEET_CONFIG: para adicionar um novo
    segmento ou aba, basta acrescentar uma entrada no dicionário.
  - O workbook é aberto UMA única vez por aba para resolver todos os ranges,
    melhorando a performance.
  - Colunas continuam fixas (conforme layout real da planilha).
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
# LOG DUPLO: arquivo + console
# ===========================================================================

class DualLog:
    """Escreve simultaneamente em arquivo e no console original."""
    def __init__(self, filename: str, mode: str = "a"):
        self._console = sys.__stdout__
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

try:
    from PIL import Image
except ImportError:
    Image = None


# ===========================================================================
# CONFIGURAÇÕES
# ===========================================================================

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

_dest_raw     = os.getenv("DESTINATARIOS", "")
DESTINATARIOS = [e.strip() for e in _dest_raw.split(",") if e.strip()]

ASSUNTO     = "Relatório de Produção - Logística Florestal"
SAUDACAO    = "Prezados,"
TEXTO_CORPO = (
    "Segue em anexo o <strong>Relatório de Produção</strong> da Eldorado Brasil, "
    "com os dados consolidados do período."
)

IMAGE_BASE = "report_temp"

# ---------------------------------------------------------------------------
# NOMES DAS ABAS
# ---------------------------------------------------------------------------
SHEET_DIARIA = "Tabela - Produção Diária"
SHEET_MENSAL = "Tabela - Produção Mensal"

# ---------------------------------------------------------------------------
# NOMES DOS SEGMENTOS
# ---------------------------------------------------------------------------
SEGMENTO_CELULOSE = "celulose"
SEGMENTO_BIOMASSA = "biomassa"

# ===========================================================================
# CONFIGURAÇÃO CENTRAL DE ABAS E SEGMENTOS
#
# Como funciona:
#   • "col_*_start" e "col_*_end" são as colunas fixas (nunca mudam).
#   • "row_*_start" é a linha inicial fixa.
#   • A linha FINAL é detectada automaticamente — não precisa mexer aqui
#     quando a planilha crescer.
#   • "row_pad" é uma margem extra de linhas abaixo do último dado
#     (útil para gráficos que têm legenda abaixo da última linha de dados).
#
# Para adicionar um novo segmento (ex: "madeira"):
#   1. Defina SEGMENTO_MADEIRA = "madeira"
#   2. Adicione a entrada no dicionário "segments" da aba desejada.
#
# Para adicionar uma nova aba:
#   1. Defina o nome: SHEET_NOVA = "Tabela - Nova Aba"
#   2. Adicione SHEET_NOVA: { "has_charts": ..., "segments": { ... } }
# ===========================================================================

SHEET_CONFIG = {
    SHEET_DIARIA: {
        "has_charts": True,          # possui gráficos para exportar
        "segments": {
            SEGMENTO_CELULOSE: {
                # Colunas do gráfico (fixas)
                "col_grafico_start": "B",
                "col_grafico_end":   "L",
                "row_grafico_start": 3,   # linha inicial do gráfico
                "row_grafico_pad":   3,   # margem extra abaixo (gráficos têm legenda)

                # Colunas da tabela (fixas)
                "col_tabela_start":  "B",
                "col_tabela_end":    "L",
                "row_tabela_start":  5,   # linha inicial da tabela
                "row_tabela_pad":    2,   # margem extra abaixo
            },
            SEGMENTO_BIOMASSA: {
                "col_grafico_start": "Q",
                "col_grafico_end":   "AA",
                "row_grafico_start": 3,
                "row_grafico_pad":   3,

                "col_tabela_start":  "Q",
                "col_tabela_end":    "AA",
                "row_tabela_start":  5,
                "row_tabela_pad":    2,
            },
        },
    },
    SHEET_MENSAL: {
        "has_charts": False,         # sem gráficos — apenas tabelas
        "segments": {
            SEGMENTO_CELULOSE: {
                "col_tabela_start": "B",
                "col_tabela_end":   "J",
                "row_tabela_start": 3,
                "row_tabela_pad":   2,
            },
            SEGMENTO_BIOMASSA: {
                "col_tabela_start": "Q",
                "col_tabela_end":   "Y",
                "row_tabela_start": 3,
                "row_tabela_pad":   2,
            },
        },
    },
}


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


# ===========================================================================
# DETECÇÃO DINÂMICA DA ÚLTIMA LINHA PREENCHIDA  ← NOVO EM v3.4.0
# ===========================================================================

def _col_letra_para_numero(col: str) -> int:
    """Converte letra(s) de coluna Excel para número (A=1, B=2, AA=27...)."""
    col = col.upper().strip()
    num = 0
    for ch in col:
        num = num * 26 + (ord(ch) - ord("A") + 1)
    return num


def _detectar_ultima_linha(
    wb,
    sheet_name:  str,
    col_start:   str,
    col_end:     str,
    row_start:   int,
    row_pad:     int = 2,
) -> int:
    """
    Detecta automaticamente a última linha preenchida dentro do bloco
    de colunas [col_start : col_end], a partir de row_start.

    Estratégia:
      1. Usa o UsedRange da aba como teto máximo (evita varrer linhas vazias
         até o infinito).
      2. Varre de baixo para cima — para na primeira célula não-vazia.
      3. Soma row_pad como margem de segurança.

    Retorna row_start + 1 como mínimo caso não encontre nada.
    """
    sht     = wb.sheets[sheet_name]
    max_row = sht.used_range.last_cell.row

    c_start = _col_letra_para_numero(col_start)
    c_end   = _col_letra_para_numero(col_end)

    for row in range(max_row, row_start - 1, -1):
        for col in range(c_start, c_end + 1):
            try:
                val = sht.cells(row, col).value
                if val is not None and str(val).strip() not in ("", "None"):
                    return row + row_pad
            except Exception:
                pass

    return row_start + 1  # fallback mínimo


def _resolver_range(
    wb,
    sheet_name:  str,
    col_start:   str,
    col_end:     str,
    row_start:   int,
    row_pad:     int = 2,
) -> str:
    """
    Monta o endereço de range com linha final detectada dinamicamente.
    Exemplo: col_start='B', col_end='L', row_start=3, última linha=17
             → retorna 'B3:L19'  (17 + row_pad=2)
    """
    row_end    = _detectar_ultima_linha(wb, sheet_name, col_start, col_end, row_start, row_pad)
    range_addr = f"{col_start}{row_start}:{col_end}{row_end}"
    print(f"[LOG]   Range dinâmico '{sheet_name}' [{col_start}→{col_end}]: {range_addr}")
    return range_addr


# ===========================================================================
# CAPTURA DE IMAGEM VIA XLWINGS (preserva layout 100%)
# ===========================================================================

def _capturar_range_como_imagem(
    path: str,
    sheet_name: str,
    range_address: str,
    output_path: str,
) -> bool:
    """
    Captura um range específico do Excel como PNG via xlwings + CopyPicture.
    Preserva formatação, cores e layout originais do Excel.
    Retorna True se bem-sucedido.
    """
    try:
        import xlwings as xw
        import win32clipboard
        from PIL import Image
        import io
    except ImportError as e:
        print(f"[AVISO] Dependência ausente para captura de imagem: {e}")
        return False

    app = None
    wb  = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        wb  = app.books.open(str(path), update_links=False)
        sht = wb.sheets[sheet_name]

        rng = sht.range(range_address)
        rng.api.CopyPicture(Appearance=1, Format=2)  # xlScreen, xlBitmap

        win32clipboard.OpenClipboard()
        try:
            data = win32clipboard.GetClipboardData(win32clipboard.CF_DIB)
        finally:
            win32clipboard.CloseClipboard()

        img = Image.open(io.BytesIO(data))
        img.save(output_path, "PNG")
        print(f"[LOG]   Range '{range_address}' capturado → {output_path}")
        return True

    except Exception as exc:
        print(f"[AVISO] Falha ao capturar '{range_address}' de '{sheet_name}': {exc}")
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


def _capturar_aba_como_imagem(
    path: str,
    sheet_name: str,
    output_path: str,
) -> bool:
    """
    Detecta automaticamente o used_range da aba e captura como imagem.
    Usado como fallback quando ranges explícitos não estão definidos.
    """
    try:
        import xlwings as xw
    except ImportError:
        return False

    app = None
    wb  = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        wb  = app.books.open(str(path), update_links=False)
        sht = wb.sheets[sheet_name]
        range_addr = sht.used_range.address
        wb.close(save_changes=False)
        wb = None
        app.quit()
        app = None
        return _capturar_range_como_imagem(path, sheet_name, range_addr, output_path)
    except Exception as exc:
        print(f"[AVISO] Falha ao detectar range de '{sheet_name}': {exc}")
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
# EXPORTAÇÃO DE GRÁFICOS POR SEGMENTO
# ===========================================================================

def _exportar_graficos_por_segmento(
    path: str,
    sheet_name: str,
    base_name: str,
) -> dict[str, list[tuple[str, str]]]:
    """
    Exporta gráficos da aba separando-os por segmento (celulose / biomassa)
    com base na posição horizontal do ChartObject:
      • Left < limiar  → celulose
      • Left >= limiar → biomassa

    Retorna dict: {"celulose": [(titulo, path), ...], "biomassa": [...]}
    """
    resultado: dict[str, list[tuple[str, str]]] = {
        SEGMENTO_CELULOSE: [],
        SEGMENTO_BIOMASSA: [],
    }

    try:
        import xlwings as xw
    except ImportError:
        print("[AVISO] xlwings não instalado — exportação de gráficos ignorada.")
        return resultado

    LIMIAR_HORIZONTAL = 700  # pixels: abaixo = celulose, acima = biomassa

    app = None
    wb  = None
    try:
        app = xw.App(visible=False, add_book=False)
        wb  = app.books.open(str(path))
        sht = wb.sheets[sheet_name]

        try:
            for slicer in sht.api.Slicers:
                slicer.SlicerItems.ClearAll()
        except Exception:
            pass

        count = sht.api.ChartObjects().Count
        if count <= 0:
            print(f"[LOG]   Nenhum gráfico em '{sheet_name}'.")
            return resultado

        print(f"[LOG] Exportando {count} gráfico(s) de '{sheet_name}'...")

        for idx in range(1, count + 1):
            chart_obj = sht.api.ChartObjects(idx)
            segmento  = SEGMENTO_CELULOSE if chart_obj.Left < LIMIAR_HORIZONTAL else SEGMENTO_BIOMASSA
            titulo    = (
                "Logística de Celulose — Volume Entregue (M³)"
                if segmento == SEGMENTO_CELULOSE
                else "Logística de Biomassa — Volume Entregue (M³)"
            )

            nome_arquivo = (
                f"{base_name}_grafico_{sheet_name[:15]}_{segmento}_{idx:02d}.png"
            )
            nome_arquivo = "".join(c if c.isalnum() or c in "-_." else "_" for c in nome_arquivo)
            caminho_abs  = os.path.abspath(nome_arquivo)

            exportou = chart_obj.Chart.Export(str(caminho_abs), "PNG")
            if exportou and os.path.exists(caminho_abs):
                resultado[segmento].append((titulo, caminho_abs))
                print(f"[LOG]   Gráfico [{segmento}] exportado: {caminho_abs}")
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

    return resultado


# ===========================================================================
# CAPTURA DE TABELAS POR SEGMENTO (screenshot de range específico)
# ===========================================================================

def _capturar_tabela_segmento(
    path: str,
    sheet_name: str,
    segmento: str,
    range_addr: str,
    base_name: str,
    sufixo: str,
) -> tuple[str, str] | None:
    """
    Captura um range de tabela de um segmento como imagem PNG.
    Retorna (titulo, caminho_png) ou None se falhar.
    """
    titulo = (
        "Logística de Celulose — Tabela de Produção"
        if segmento == SEGMENTO_CELULOSE
        else "Logística de Biomassa — Tabela de Produção"
    )

    nome_arquivo = f"{base_name}_tabela_{sufixo}_{segmento}.png"
    nome_arquivo = "".join(c if c.isalnum() or c in "-_." else "_" for c in nome_arquivo)
    output_path  = os.path.abspath(nome_arquivo)

    print(f"[LOG] Capturando tabela [{segmento}] de '{sheet_name}' (range={range_addr})...")
    ok = _capturar_range_como_imagem(path, sheet_name, range_addr, output_path)

    if ok and os.path.exists(output_path):
        return (titulo, output_path)

    print(f"[AVISO] Captura via xlwings falhou para [{segmento}] em '{sheet_name}'.")
    return None


# ===========================================================================
# REFRESH DO EXCEL
# ===========================================================================

def _refresh_excel(path: str | Path, timeout_s: int | None = None) -> bool:
    caminho = Path(path)
    if not caminho.exists():
        print(f"[AVISO] Arquivo Excel não encontrado para refresh: {caminho}")
        return False

    try:
        import xlwings as xw
    except ImportError:
        print("[AVISO] xlwings não instalado — refresh ignorado.")
        return False

    timeout_s = timeout_s or int(os.getenv("EXCEL_REFRESH_TIMEOUT", "300"))
    print(f"[LOG] Abrindo Excel para refresh: {caminho.name}")

    app = None
    wb  = None
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
                print(f"[AVISO] Timeout de {timeout_s}s atingido. Continuando com dados atuais.")
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
# FALLBACK: GERAÇÃO DE TABELA VIA PANDAS + MATPLOTLIB
# ===========================================================================

def _split_header_segments(values: list, max_gap: int = 3) -> list[tuple[int, int]]:
    segments = []
    start    = None
    gap      = 0
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
                gap   = 0
    if start is not None:
        segments.append((start, len(values) - 1))
    return segments


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


def extrator_subtabelas(df: pd.DataFrame, titulo_aba: str = "") -> list[tuple[str, pd.DataFrame]]:
    """Extrai sub-tabelas da aba detectando cabeçalhos com 'Fazenda'."""
    resultados  = []
    header_rows = []

    for i in range(len(df)):
        valores = [str(v).strip() for v in df.iloc[i].tolist()]
        if "Fazenda" in valores and any(
            kw in v for v in valores if v
            for kw in ("Meta", "Vol.", "DMT", "RPV", "Real")
        ):
            header_rows.append(i)

    if not header_rows:
        print(f"[AVISO] Nenhum cabeçalho encontrado em '{titulo_aba}'.")
        return resultados

    for header_row in header_rows:
        segmentos = _split_header_segments(df.iloc[header_row].tolist(), max_gap=3)
        for start_col, end_col in segmentos:
            cabecalhos = [
                str(v).strip() if not _is_blank(v) else ""
                for v in df.iloc[header_row, start_col:end_col + 1].tolist()
            ]
            if not any("Fazenda" in c or "Meta" in c or "Vol" in c or "DMT" in c for c in cabecalhos if c):
                continue

            row_end = _find_segment_end(df, header_row, start_col, end_col)
            sub = df.iloc[header_row + 1:row_end, start_col:end_col + 1].copy()
            sub.columns = cabecalhos

            sub = sub.dropna(how="all").reset_index(drop=True)
            if sub.empty:
                continue

            sub = sub.apply(lambda col: col.map(_fmt_num))

            meio = df.shape[1] // 2
            segmento_nome = SEGMENTO_CELULOSE if start_col < meio else SEGMENTO_BIOMASSA
            titulo_final  = (
                "Logística de Celulose — Tabela de Produção"
                if segmento_nome == SEGMENTO_CELULOSE
                else "Logística de Biomassa — Tabela de Produção"
            )
            resultados.append((titulo_final, sub, segmento_nome))

    return resultados


def _gerar_imagens_matplotlib(
    tabelas: list[tuple[str, pd.DataFrame, str]],
    base: str,
    sufixo: str,
) -> dict[str, list[tuple[str, str]]]:
    """Gera imagens PNG de tabelas via matplotlib (fallback)."""
    resultado: dict[str, list[tuple[str, str]]] = {
        SEGMENTO_CELULOSE: [],
        SEGMENTO_BIOMASSA: [],
    }

    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    COR_HEADER = "#1e5f3b"
    COR_STRIPE = "#e8f5ee"
    COR_BRANCO = "#ffffff"
    COR_BORDER = "#a8d5b5"

    for sub_idx, (titulo, df, segmento) in enumerate(tabelas, start=1):
        df = df.reset_index(drop=True)

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

        n_rows, n_cols = df.shape
        col_labels = list(df.columns)
        cell_data  = df.values.tolist()

        CHAR_W = 0.10
        MIN_W  = 0.45
        MAX_W  = 4.0
        PAD    = 0.20

        col_widths = []
        for ci in range(n_cols):
            header_len = len(str(col_labels[ci]))
            data_len   = max((len(str(row[ci])) for row in cell_data), default=0)
            w          = min(MAX_W, max(MIN_W, max(header_len, data_len) * CHAR_W + PAD))
            col_widths.append(w)

        fig_w = sum(col_widths) + 0.5
        ROW_H = 0.40
        fig_h = (n_rows + 2) * ROW_H + 0.6

        fig, ax = plt.subplots(figsize=(fig_w, fig_h))
        fig.patch.set_facecolor(COR_BRANCO)
        ax.set_facecolor(COR_BRANCO)
        ax.axis("off")

        tbl = ax.table(
            cellText=cell_data,
            colLabels=col_labels,
            cellLoc="center",
            loc="center",
            colWidths=[w / fig_w for w in col_widths],
        )
        tbl.auto_set_font_size(False)
        tbl.set_fontsize(11)

        for ci in range(n_cols):
            cell = tbl[0, ci]
            cell.set_facecolor(COR_HEADER)
            cell.set_text_props(color="white", fontweight="bold")
            cell.set_edgecolor(COR_BORDER)
            cell.set_height(ROW_H / fig_h)

        for ri in range(1, n_rows + 1):
            bg = COR_STRIPE if ri % 2 == 1 else COR_BRANCO
            for ci in range(n_cols):
                cell = tbl[ri, ci]
                cell.set_facecolor(bg)
                cell.set_edgecolor(COR_BORDER)
                cell.set_text_props(color="#333333")
                cell.set_height(ROW_H / fig_h)

        nome_seg  = "".join(c if c.isalnum() or c in "-_" else "_" for c in segmento)
        caminho   = os.path.abspath(f"{base}_tabela_{sufixo}_{nome_seg}_{sub_idx:02d}.png")

        plt.tight_layout()
        plt.savefig(caminho, dpi=160, bbox_inches="tight", facecolor=COR_BRANCO, format="PNG")
        plt.close(fig)
        print(f"[LOG]   (matplotlib fallback) Salvo: {caminho}")
        resultado[segmento].append((titulo, caminho))

    return resultado


# ===========================================================================
# EXTRAIR DATA DE COLETA
# ===========================================================================

def _extrair_data_coleta(path: str, sheet_name: str) -> str:
    try:
        df = pd.read_excel(path, sheet_name=sheet_name, header=None)
        for idx, row in df.iterrows():
            for val in row:
                if pd.notna(val):
                    try:
                        dt = pd.to_datetime(str(val).strip(), dayfirst=True)
                        if pd.Timestamp.now().year - 2 <= dt.year <= pd.Timestamp.now().year + 1:
                            return dt.strftime("%d/%m/%Y")
                    except (ValueError, TypeError):
                        pass
    except Exception as exc:
        print(f"[AVISO] Falha ao extrair data: {exc}")
    return ""


# ===========================================================================
# MONTAGEM DO HTML DO E-MAIL
# ===========================================================================

def _img_b64(caminho: str) -> str:
    import base64
    try:
        with open(caminho, "rb") as f:
            return f"data:image/png;base64,{base64.b64encode(f.read()).decode('utf-8')}"
    except Exception:
        return ""


def _bloco_imagem_html(titulo: str, caminho: str, largura: int = 900) -> str:
    b64 = _img_b64(caminho)
    if not b64:
        return ""
    return f"""
<div style="margin-bottom:20px;text-align:left;">
  <p style="margin:0 0 5px;font-size:12px;font-weight:700;color:#1e5f3b;
            text-transform:uppercase;letter-spacing:0.3px;">{titulo}</p>
  <img src="{b64}" alt="{titulo}"
       style="max-width:{largura}px;width:100%;height:auto;
              border:1px solid #a8d5b5;border-radius:4px;display:block;" />
</div>"""


def _separador_secao_html(titulo_secao: str) -> str:
    return f"""
<div style="margin:28px 0 14px;padding:8px 16px;
            background-color:#1e5f3b;border-radius:4px;">
  <span style="font-size:12px;font-weight:700;color:#ffffff;
               text-transform:uppercase;letter-spacing:0.8px;">
    {titulo_secao}
  </span>
</div>"""


def montar_blocos_email(
    graficos_diarios: dict[str, list[tuple[str, str]]],
    tabelas_diarias:  dict[str, list[tuple[str, str]]],
    tabelas_mensais:  dict[str, list[tuple[str, str]]],
) -> str:
    """
    Monta o HTML do corpo seguindo a ordem:
      1. Celulose  — Produção Diária   (gráfico → tabela)
      2. Celulose  — Produção Mensal   (tabela)
      3. Biomassa  — Produção Diária   (gráfico → tabela)
      4. Biomassa  — Produção Mensal   (tabela)
    """
    html = ""

    for segmento, label_seg in [
        (SEGMENTO_CELULOSE, "Logística de Celulose"),
        (SEGMENTO_BIOMASSA, "Logística de Biomassa"),
    ]:
        graficos_d = graficos_diarios.get(segmento, [])
        tabelas_d  = tabelas_diarias.get(segmento, [])

        if graficos_d or tabelas_d:
            html += _separador_secao_html(f"{label_seg} — Produção Diária")
            max_items = max(len(graficos_d), len(tabelas_d))
            for i in range(max_items):
                if i < len(graficos_d):
                    tit, pth = graficos_d[i]
                    html += _bloco_imagem_html(tit, pth, largura=920)
                if i < len(tabelas_d):
                    tit, pth = tabelas_d[i]
                    html += _bloco_imagem_html(tit, pth, largura=920)

        tabelas_m = tabelas_mensais.get(segmento, [])
        if tabelas_m:
            html += _separador_secao_html(f"{label_seg} — Produção Mensal")
            for tit, pth in tabelas_m:
                html += _bloco_imagem_html(tit, pth, largura=920)

    return html


# ===========================================================================
# HTML COMPLETO DO E-MAIL
# ===========================================================================

def montar_email_html(
    assunto: str,
    saudacao: str,
    texto_corpo: str,
    data_atual: str,
    data_coleta: str,
    assinatura_nome: str,
    assinatura_cargo: str,
    assinatura_email_sig: str,
    assinatura_tel: str,
    blocos_conteudo: str,
) -> str:

    data_coleta_html = (
        f'<tr><td style="padding:4px 32px 0;">'
        f'<p style="margin:0;font-size:11px;color:#888888;font-style:italic;">'
        f'Dados coletados em: {data_coleta}</p></td></tr>'
        if data_coleta else ""
    )

    return f"""<!DOCTYPE html>
<html lang="pt-BR">
<head><meta charset="UTF-8"><title>{assunto}</title></head>
<body style="margin:0;padding:0;background-color:#f0f7f3;
             font-family:Calibri,Arial,sans-serif;">
  <table width="100%" cellpadding="0" cellspacing="0"
         style="background-color:#f0f7f3;padding:30px 0;">
    <tr><td align="center">
      <table width="980" cellpadding="0" cellspacing="0"
             style="background-color:#ffffff;border-radius:8px;
                    box-shadow:0 2px 8px rgba(0,0,0,0.10);overflow:hidden;">

        <!-- CABEÇALHO -->
        <tr>
          <td style="background-color:#1e5f3b;padding:22px 32px;">
            <table width="100%" cellpadding="0" cellspacing="0"><tr>
              <td>
                <span style="font-size:20px;font-weight:700;color:#ffffff;">
                  Eldorado Brasil</span><br>
                <span style="font-size:13px;color:#a8d5b5;">
                  Relatório de Produção — Logística Florestal</span>
              </td>
              <td align="right">
                <span style="font-size:12px;color:#a8d5b5;">{data_atual}</span>
              </td>
            </tr></table>
          </td>
        </tr>

        <!-- INTRODUÇÃO -->
        <tr>
          <td style="padding:26px 32px 14px;">
            <p style="margin:0 0 12px;font-size:15px;color:#333333;">{saudacao}</p>
            <p style="margin:0 0 20px;font-size:14px;color:#555555;line-height:1.65;">
              {texto_corpo}</p>
            <p style="margin:0;font-size:11px;font-weight:700;color:#1e5f3b;
                      text-transform:uppercase;letter-spacing:0.5px;">
              Indicadores Operacionais</p>
          </td>
        </tr>
        <tr>
          <td><hr style="border:none;border-top:1px solid #c8e6d4;margin:0;"></td>
        </tr>

        {data_coleta_html}

        <!-- CONTEÚDO: GRÁFICOS E TABELAS -->
        <tr>
          <td style="padding:4px 32px 20px;">
            {blocos_conteudo}
          </td>
        </tr>

        <tr>
          <td><hr style="border:none;border-top:1px solid #c8e6d4;margin:0;"></td>
        </tr>

        <!-- ASSINATURA -->
        <tr>
          <td style="padding:14px 24px 20px;">
            <table width="100%" cellpadding="0" cellspacing="0"
                   style="border-collapse:collapse;">
              <tr>
                <td width="65%" style="vertical-align:top;padding-right:28px;">
                  <div style="font-size:14px;font-weight:700;color:#1e5f3b;
                              line-height:18px;">{assinatura_nome}</div>
                  <div style="font-size:12px;color:#555555;line-height:18px;">
                    {assinatura_cargo}</div>
                  <div style="height:8px;"></div>
                  <div style="font-size:12px;line-height:20px;">
                    <a href="mailto:{assinatura_email_sig}"
                       style="color:#1e5f3b;text-decoration:underline;">
                      {assinatura_email_sig}</a></div>
                  <div style="font-size:12px;line-height:20px;">
                    <a href="https://eldoradobrasil.com.br"
                       style="color:#1e5f3b;text-decoration:underline;">
                      eldoradobrasil.com.br</a></div>
                  <div style="font-size:12px;line-height:20px;">
                    <a href="https://linhaetica.eldoradobrasil.com.br"
                       style="color:#1e5f3b;text-decoration:underline;">
                      linhaetica.eldoradobrasil.com.br</a></div>
                </td>
                <td width="1" style="background:#d4e8db;font-size:0;
                                     line-height:0;">&nbsp;</td>
                <td width="35%" style="vertical-align:top;padding-left:28px;">
                  <div style="font-size:12px;color:#555555;line-height:18px;">
                    Rod. Br 158 – Km 231</div>
                  <div style="font-size:12px;color:#555555;line-height:18px;">
                    Três Lagoas - MS - Brasil</div>
                  <div style="font-size:12px;color:#555555;line-height:18px;">
                    CEP: 79641-300</div>
                  <div style="height:6px;"></div>
                  <div style="font-size:12px;color:#555555;line-height:18px;">
                    {assinatura_tel}</div>
                </td>
              </tr>
            </table>
          </td>
        </tr>

        <!-- RODAPÉ -->
        <tr>
          <td style="background-color:#f0f7f3;padding:10px 24px;
                     border-top:1px solid #dde8e2;text-align:center;">
            <span style="font-size:10px;color:#7a9e82;">
              E-mail gerado automaticamente. Por favor, não responda diretamente.
            </span>
          </td>
        </tr>

      </table>
    </td></tr>
  </table>
</body></html>"""


# ===========================================================================
# ENVIO VIA OUTLOOK
# ===========================================================================

def enviar_email(
    html_body: str,
    assunto: str,
    email_remetente: str,
    destinatarios: list[str],
    anexos: list[str],
) -> None:
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
        if os.path.exists(str(caminho)):
            mail.Attachments.Add(os.path.abspath(str(caminho)))
        else:
            print(f"[AVISO] Anexo não encontrado: {caminho}")

    mail.Send()
    print(f"[LOG] E-mail enviado para: {', '.join(destinatarios)}")


# ===========================================================================
# LIMPEZA DE TEMPORÁRIOS
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
    temp_paths: list[str] = []

    try:
        # ------------------------------------------------------------------
        # Validações
        # ------------------------------------------------------------------
        if not EMAIL_REMETENTE:
            raise RuntimeError("EMAIL_REMETENTE não definido no .env.")
        if not DESTINATARIOS:
            raise RuntimeError("DESTINATARIOS não definido no .env (separe por vírgula).")

        print(f"[LOG] Destinatários ({len(DESTINATARIOS)}):")
        for d in DESTINATARIOS:
            print(f"       → {d}")

        # ------------------------------------------------------------------
        # Refresh da planilha
        # ------------------------------------------------------------------
        if EXCEL_PATH and Path(EXCEL_PATH).exists():
            if not _refresh_excel(EXCEL_PATH):
                print("[AVISO] Não foi possível atualizar a planilha antes da leitura.")
        else:
            print(f"[AVISO] Planilha não encontrada em: {EXCEL_PATH}")

        # ------------------------------------------------------------------
        # Estruturas de resultado
        # ------------------------------------------------------------------
        graficos_diarios: dict[str, list[tuple[str, str]]] = {
            SEGMENTO_CELULOSE: [],
            SEGMENTO_BIOMASSA: [],
        }
        tabelas_diarias: dict[str, list[tuple[str, str]]] = {
            SEGMENTO_CELULOSE: [],
            SEGMENTO_BIOMASSA: [],
        }
        tabelas_mensais: dict[str, list[tuple[str, str]]] = {
            SEGMENTO_CELULOSE: [],
            SEGMENTO_BIOMASSA: [],
        }

        # ==================================================================
        # PROCESSAMENTO GENÉRICO — orientado por SHEET_CONFIG     ← v3.4.0
        #
        # Para cada aba definida em SHEET_CONFIG:
        #   1. Abre o workbook UMA vez para resolver todos os ranges
        #      (detecta a última linha automaticamente).
        #   2. Fecha o workbook.
        #   3. Exporta gráficos (se houver).
        #   4. Captura tabelas usando os ranges já resolvidos.
        # ==================================================================

        try:
            import xlwings as xw
            _xlwings_ok = True
        except ImportError:
            print("[AVISO] xlwings não disponível — usando apenas fallback matplotlib.")
            _xlwings_ok = False

        for sheet_name, sheet_cfg in SHEET_CONFIG.items():
            has_charts = sheet_cfg.get("has_charts", False)
            sufixo     = "diaria" if sheet_name == SHEET_DIARIA else "mensal"
            destino    = tabelas_diarias if sheet_name == SHEET_DIARIA else tabelas_mensais

            print(f"\n[LOG] ══ Processando aba: '{sheet_name}' ══")

            # ── Passo 1: resolve ranges dinamicamente (UMA abertura por aba) ──
            resolved: dict[str, dict[str, str]] = {}

            if _xlwings_ok:
                _app = None
                _wb  = None
                try:
                    _app = xw.App(visible=False, add_book=False)
                    _app.display_alerts = False
                    _wb  = _app.books.open(str(EXCEL_PATH), update_links=False)

                    for seg, seg_cfg in sheet_cfg["segments"].items():
                        resolved[seg] = {}

                        # Tabela (sempre presente)
                        resolved[seg]["tabela"] = _resolver_range(
                            _wb,
                            sheet_name,
                            seg_cfg["col_tabela_start"],
                            seg_cfg["col_tabela_end"],
                            seg_cfg["row_tabela_start"],
                            seg_cfg.get("row_tabela_pad", 2),
                        )

                        # Gráfico (somente abas com has_charts=True)
                        if has_charts and "col_grafico_start" in seg_cfg:
                            resolved[seg]["grafico"] = _resolver_range(
                                _wb,
                                sheet_name,
                                seg_cfg["col_grafico_start"],
                                seg_cfg["col_grafico_end"],
                                seg_cfg["row_grafico_start"],
                                seg_cfg.get("row_grafico_pad", 3),
                            )

                except Exception as exc:
                    print(f"[AVISO] Não foi possível resolver ranges de '{sheet_name}': {exc}")
                finally:
                    try:
                        _wb.close(save_changes=False)
                    except Exception:
                        pass
                    try:
                        _app.quit()
                    except Exception:
                        pass

            # ── Passo 2: exporta gráficos (se a aba tiver) ───────────────────
            if has_charts:
                graficos_result = _exportar_graficos_por_segmento(
                    EXCEL_PATH, sheet_name, IMAGE_BASE
                )
                for seg in sheet_cfg["segments"]:
                    graficos_diarios[seg] = graficos_result.get(seg, [])
                    temp_paths.extend(p for _, p in graficos_diarios[seg])

            # ── Passo 3: captura tabelas com ranges resolvidos ───────────────
            for seg in sheet_cfg["segments"]:
                range_addr = resolved.get(seg, {}).get("tabela")

                if range_addr:
                    item = _capturar_tabela_segmento(
                        path       = EXCEL_PATH,
                        sheet_name = sheet_name,
                        segmento   = seg,
                        range_addr = range_addr,
                        base_name  = IMAGE_BASE,
                        sufixo     = sufixo,
                    )
                    if item:
                        destino[seg].append(item)
                        temp_paths.append(item[1])
                        continue  # sucesso — pula fallback

                # Fallback matplotlib (xlwings indisponível ou falhou)
                print(f"[LOG] Usando fallback matplotlib para [{seg}] em '{sheet_name}'...")
                try:
                    df_raw   = pd.read_excel(EXCEL_PATH, sheet_name=sheet_name, header=None)
                    subtabs  = [t for t in extrator_subtabelas(df_raw, sheet_name) if t[2] == seg]
                    fallback = _gerar_imagens_matplotlib(subtabs, IMAGE_BASE, f"{sufixo}_{seg}")
                    destino[seg] = fallback.get(seg, [])
                    temp_paths.extend(p for _, p in destino[seg])
                except Exception as exc:
                    print(f"[AVISO] Fallback matplotlib falhou para '{sheet_name}' [{seg}]: {exc}")

        # ------------------------------------------------------------------
        # Montar HTML do e-mail
        # ------------------------------------------------------------------
        blocos_html = montar_blocos_email(
            graficos_diarios = graficos_diarios,
            tabelas_diarias  = tabelas_diarias,
            tabelas_mensais  = tabelas_mensais,
        )

        if not blocos_html.strip():
            raise RuntimeError("Nenhum conteúdo gerado para o e-mail.")

        data_coleta = _extrair_data_coleta(EXCEL_PATH, SHEET_DIARIA)

        html_body = montar_email_html(
            assunto              = ASSUNTO,
            saudacao             = SAUDACAO,
            texto_corpo          = TEXTO_CORPO,
            data_atual           = data_atual,
            data_coleta          = data_coleta,
            assinatura_nome      = ASSINATURA_NOME,
            assinatura_cargo     = ASSINATURA_CARGO,
            assinatura_email_sig = ASSINATURA_EMAIL,
            assinatura_tel       = ASSINATURA_TEL,
            blocos_conteudo      = blocos_html,
        )
        print("\n[LOG] E-mail montado com sucesso.")

        # ------------------------------------------------------------------
        # Anexos e envio
        # ------------------------------------------------------------------
        anexos = []
        if ANEXAR_PLANILHA and Path(EXCEL_PATH).exists():
            anexos.append(EXCEL_PATH)

        enviar_email(
            html_body       = html_body,
            assunto         = ASSUNTO,
            email_remetente = EMAIL_REMETENTE,
            destinatarios   = DESTINATARIOS,
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
        limpar_temporarios(temp_paths)
        print(f"\n[LOG] Duração total: {time.time() - inicio:.1f}s")
        print("=" * 60)


if __name__ == "__main__":
    executar()