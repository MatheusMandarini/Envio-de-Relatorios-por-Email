"""
=============================================================================
SISTEMA DE AUTOMAÇÃO CORPORATIVA — ENVIO DE INDICADORES POR E-MAIL
=============================================================================
Autor      : Matheus Menezes
Empresa    : Eldorado Brasil
Versão     : 2.0.0

Funcionamento:
    1. Lê a aba "Indicadores Semanal_Flash" da planilha Excel
    2. Detecta automaticamente cada bloco de linhas (Vol Real, DMT, RPV...)
    3. Dentro de cada bloco, extrai cada sub-tabela lateral separadamente
    4. Gera uma imagem PNG por sub-tabela (sem "nan" nas capturas)
    5. Envia e-mail HTML com imagens inline + planilha anexada

Instalação das dependências:
    pip install pandas openpyxl dataframe_image pillow
=============================================================================
"""

import os
import time
import traceback
import numpy as np
from datetime import datetime


def carregar_env(path: str = ".env") -> None:
    if not os.path.exists(path):
        return
    with open(path, encoding="utf-8") as arquivo:
        for linha in arquivo:
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
# CONFIGURACOES CENTRALIZADAS — edite apenas esta secao
# ===========================================================================

EXCEL_PATH = r"C:\Users\ext.matheusmm\Documents\Envio de Emails\data\Indicadores Produção.xlsx"
SHEET_NAME = "Indicadores Semanal_Flash"

# Grupos de colunas que formam cada sub-tabela dentro de um bloco de linhas.
# Cada tupla é (col_inicio_inclusivo, col_fim_exclusivo), base 0.
# Inspecione sua planilha e ajuste se adicionar/remover colunas.
SUB_COL_GROUPS = [
    (1, 8),   # Vol Real / DMT total / RPV Pond
    (9, 13),  # Cx Carga / DMT Terra / Km
    (17, 21), # Qtde Viagens / DMT Asf / Dens. Industrial
    (25, 29), # Idade Floresta / TPC Pond / Dens.Campo
]

IMAGE_BASE = "indicadores_temp"

EMAIL_REMETENTE  = os.getenv("EMAIL_REMETENTE", "")
ASSINATURA_NOME  = os.getenv("ASSINATURA_NOME", "")
ASSINATURA_CARGO = os.getenv("ASSINATURA_CARGO", "")
ASSINATURA_EMAIL = EMAIL_REMETENTE
ASSINATURA_TEL   = os.getenv("ASSINATURA_TEL", "")

_destinatarios_raw = os.getenv("DESTINATARIOS", "")
DESTINATARIOS = [
    email.strip()
    for email in _destinatarios_raw.split(",")
    if email.strip()
]

ASSUNTO          = "Indicadores Semanal Flash — Eldorado Brasil"
SAUDACAO         = "Prezados,"
TEXTO_CORPO      = (
    "Segue em anexo o <strong>Painel Semanal de Indicadores de Transporte</strong> "
    "da Eldorado Brasil, com os dados consolidados do período."
)


# ===========================================================================
# FUNCOES
# ===========================================================================

def _is_nan(v) -> bool:
    """Retorna True se o valor for NaN (float ou numpy)."""
    try:
        return np.isnan(v)
    except (TypeError, ValueError):
        return False


def _str_val(v) -> str:
    if _is_nan(v):
        return ""
    return str(v).strip()


def _is_table_anchor_row(row: pd.Series) -> bool:
    """
    Identifica linhas que iniciam um novo bloco de tabelas.
    A planilha usa 'Mês Flash' na posição de coluna 3 (base 0)
    e 'Período' na posição 4 como marcadores de cabeçalho principal.
    """
    vals = row.tolist()
    return _str_val(vals[3]) == "Mês Flash" or _str_val(vals[4]) == "Período"


def detectar_blocos(df: pd.DataFrame) -> list[tuple[int, int]]:
    """
    Encontra os intervalos de linhas de cada bloco principal de tabelas.
    Retorna lista de (linha_inicio, linha_fim_exclusiva).
    """
    blocos: list[tuple[int, int]] = []
    i = 0
    n = len(df)
    while i < n:
        if _is_table_anchor_row(df.iloc[i]):
            j = i + 1
            while j < n and not _is_table_anchor_row(df.iloc[j]):
                j += 1
            blocos.append((i, j))
            i = j
        else:
            i += 1
    return blocos


def extrair_subtabela(
    df: pd.DataFrame,
    row_start: int,
    row_end: int,
    col_start: int,
    col_end: int,
) -> pd.DataFrame | None:
    """
    Extrai uma sub-tabela do DataFrame bruto e a formata para exibição.

    - Remove colunas e linhas inteiramente vazias
    - Usa a primeira linha do bloco como título da tabela
    - Usa as duas linhas seguintes como cabeçalho multi-nível (mesclado em uma linha)
    - Formata números com 2 casas decimais

    Retorna None se a sub-tabela não tiver dados úteis.
    """
    bloco = df.iloc[row_start:row_end, col_start:col_end].copy()
    bloco = bloco.dropna(how="all", axis=1).dropna(how="all", axis=0)

    if bloco.empty or len(bloco) < 2:
        return None

    # --- Título da sub-tabela (primeira célula não-vazia do bloco) ---
    titulo = ""
    for val in bloco.iloc[0]:
        t = _str_val(val)
        if t and t.lower() not in ("mês flash", "período"):
            titulo = t
            break

    # --- Montar cabeçalho: combinar linhas 1 e 2 do bloco (índices 0 e 1) ---
    linha_cab1 = bloco.iloc[0].tolist()
    linha_cab2 = bloco.iloc[1].tolist() if len(bloco) > 1 else [""] * len(bloco.columns)
    linha_cab3 = bloco.iloc[2].tolist() if len(bloco) > 2 else [""] * len(bloco.columns)

    cabecalhos = []
    for v1, v2, v3 in zip(linha_cab1, linha_cab2, linha_cab3):
        partes = [_str_val(v) for v in (v1, v2, v3) if _str_val(v)]
        cabecalhos.append(" | ".join(partes) if partes else "—")

    # Garantir unicidade nos cabeçalhos
    seen: dict[str, int] = {}
    cabs_unicos: list[str] = []
    for c in cabecalhos:
        if c in seen:
            seen[c] += 1
            cabs_unicos.append(f"{c}_{seen[c]}")
        else:
            seen[c] = 0
            cabs_unicos.append(c)

    # Dados: a partir da 4ª linha do bloco (pulando as 3 linhas de cabeçalho)
    n_header_rows = 3
    dados = bloco.iloc[n_header_rows:].copy()
    dados.columns = cabs_unicos
    dados = dados.dropna(how="all", axis=0).reset_index(drop=True)

    if dados.empty:
        return None

    # Formatar números
    def fmt(v):
        if isinstance(v, (int, float)) and not _is_nan(v):
            return f"{v:,.2f}"
        s = _str_val(v)
        return s if s else "—"

    dados = dados.apply(lambda col: col.map(fmt))

    # Substituir strings vazias por "—"
    dados = dados.replace("", "—")

    # Guardar o título como atributo para uso na imagem
    dados.attrs["titulo"] = titulo
    return dados


def carregar_e_segmentar(path: str, sheet: str) -> list[tuple[str, pd.DataFrame]]:
    """
    Lê a planilha e retorna uma lista de (titulo, DataFrame) para cada
    sub-tabela encontrada, sem NaN.
    """
    print(f"[LOG] Abrindo planilha: {path}")
    if not os.path.exists(path):
        raise FileNotFoundError(f"Arquivo não encontrado: {path}")

    try:
        df = pd.read_excel(path, sheet_name=sheet, header=None, engine="openpyxl")
    except Exception as exc:
        raise RuntimeError(f"Erro ao ler a aba '{sheet}': {exc}") from exc

    blocos = detectar_blocos(df)
    if not blocos:
        raise RuntimeError("Nenhum bloco de tabela detectado na planilha.")

    print(f"[LOG] {len(blocos)} bloco(s) detectado(s): {blocos}")

    tabelas: list[tuple[str, pd.DataFrame]] = []
    for b_start, b_end in blocos:
        for c_start, c_end in SUB_COL_GROUPS:
            sub = extrair_subtabela(df, b_start, b_end, c_start, c_end)
            if sub is not None and not sub.empty:
                titulo = sub.attrs.get("titulo", f"Tabela {len(tabelas) + 1}")
                tabelas.append((titulo, sub))
                print(f"[LOG]   Sub-tabela extraída: '{titulo}' "
                      f"({sub.shape[0]} linhas x {sub.shape[1]} colunas)")

    print(f"[LOG] Total de sub-tabelas: {len(tabelas)}")
    return tabelas


def gerar_imagem(df: pd.DataFrame, titulo: str, caminho_imagem: str) -> None:
    """
    Converte o DataFrame em imagem PNG com estilo corporativo.
    O título da sub-tabela é adicionado como primeira linha destacada.
    """
    print(f"[LOG] Gerando imagem: '{titulo}' → {caminho_imagem}")

    try:
        estilo = (
            df.style
            .set_caption(titulo)
            .set_properties(**{
                "font-family": "Calibri, Arial, sans-serif",
                "font-size":   "10px",
                "text-align":  "center",
                "border":      "1px solid #cccccc",
                "padding":     "5px 8px",
                "white-space": "nowrap",
            })
            .set_table_styles([
                {
                    "selector": "caption",
                    "props": [
                        ("font-size",        "12px"),
                        ("font-weight",      "bold"),
                        ("color",            "#1a3e5c"),
                        ("text-align",       "left"),
                        ("padding-bottom",   "4px"),
                        ("font-family",      "Calibri, Arial, sans-serif"),
                    ],
                },
                {
                    "selector": "th",
                    "props": [
                        ("background-color", "#1a3e5c"),
                        ("color",            "white"),
                        ("font-weight",      "bold"),
                        ("font-size",        "10px"),
                        ("text-align",       "center"),
                        ("padding",          "6px 8px"),
                        ("border",           "1px solid #0f2840"),
                    ],
                },
                {
                    "selector": "tr:nth-child(even) td",
                    "props": [("background-color", "#eaf1f8")],
                },
                {
                    "selector": "table",
                    "props": [("border-collapse", "collapse"), ("width", "100%")],
                },
            ])
            .hide(axis="index")
        )

        dfi.export(
            estilo,
            caminho_imagem,
            dpi=90,
            table_conversion="matplotlib",
        )

        # Reduz o tamanho final da imagem para melhorar a visualização no e-mail.
        with Image.open(caminho_imagem) as img:
            max_width = 680
            if img.width > max_width:
                scale = max_width / img.width
                new_size = (max_width, int(img.height * scale))
                img = img.resize(new_size, Image.LANCZOS)
                img.save(caminho_imagem, format="PNG", optimize=True)

    except Exception as exc:
        raise RuntimeError(f"Erro ao gerar imagem '{titulo}': {exc}") from exc

    print(f"[LOG] Imagem criada: {caminho_imagem}")


def gerar_todas_imagens(
    tabelas: list[tuple[str, pd.DataFrame]],
    base_path: str,
) -> list[tuple[str, str]]:
    """
    Gera uma imagem por sub-tabela.
    Retorna lista de (titulo, caminho_imagem).
    """
    resultado: list[tuple[str, str]] = []
    for idx, (titulo, df) in enumerate(tabelas, start=1):
        # Nome de arquivo seguro
        nome_seguro = "".join(c if c.isalnum() or c in "-_" else "_" for c in titulo)
        caminho = f"{base_path}_{idx:02d}_{nome_seguro}.png"
        gerar_imagem(df, titulo, caminho)
        resultado.append((titulo, caminho))
    return resultado


def montar_email_html(
    assunto: str,
    saudacao: str,
    texto_corpo: str,
    data_atual: str,
    assinatura_nome: str,
    assinatura_cargo: str,
    assinatura_email_sig: str,
    assinatura_tel: str,
    imagens: list[tuple[str, str]],
    caminho_anexo: str,
) -> tuple[str, list[str], list[str]]:
    """
    Monta o HTML do e-mail.
    Retorna (html_body, lista_caminhos_imagens, lista_anexos).
    """
    caminhos = [c for _, c in imagens]

    blocos_imagem = ""
    for idx, (titulo, _) in enumerate(imagens, start=1):
        blocos_imagem += f"""
              <div style="margin-bottom:24px;text-align:center;">
                <p style="margin:0 0 6px;font-size:13px;font-weight:700;
                          color:#1a3e5c;text-align:left;">{titulo}</p>
                <img src="cid:tabela_{idx}" alt="{titulo}"
                     width="560"
                     style="max-width:560px;width:100%;height:auto;border:1px solid #dde3ea;
                            border-radius:4px;display:block;margin:0 auto;" />
              </div>"""

    html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>{assunto}</title>
</head>
<body style="margin:0;padding:0;background-color:#f4f6f9;
             font-family:Calibri,Arial,sans-serif;">

  <table width="100%" cellpadding="0" cellspacing="0"
         style="background-color:#f4f6f9;padding:30px 0;">
    <tr>
      <td align="center">
        <table width="700" cellpadding="0" cellspacing="0"
               style="background-color:#ffffff;border-radius:8px;
                      box-shadow:0 2px 8px rgba(0,0,0,0.10);overflow:hidden;">

          <!-- CABECALHO -->
          <tr>
            <td style="background-color:#1a3e5c;padding:22px 32px;">
              <table width="100%" cellpadding="0" cellspacing="0">
                <tr>
                  <td>
                    <span style="font-size:20px;font-weight:700;
                                 color:#ffffff;letter-spacing:0.4px;">
                      Eldorado Brasil
                    </span><br>
                    <span style="font-size:13px;color:#a8c8e8;">
                      Painel Semanal de Indicadores — Transporte
                    </span>
                  </td>
                  <td align="right">
                    <span style="font-size:12px;color:#a8c8e8;">{data_atual}</span>
                  </td>
                </tr>
              </table>
            </td>
          </tr>

          <!-- CORPO -->
          <tr>
            <td style="padding:28px 32px 18px;">
              <p style="margin:0 0 14px;font-size:15px;color:#333333;">{saudacao}</p>
              <p style="margin:0 0 22px;font-size:14px;color:#555555;line-height:1.65;">
                {texto_corpo}
              </p>
              <p style="margin:0 0 6px;font-size:12px;font-weight:700;color:#1a3e5c;
                        text-transform:uppercase;letter-spacing:1px;">
                Indicadores Semanal Flash
              </p>
              <p style="margin:0 0 20px;font-size:12px;color:#888888;line-height:1.55;">
                Cada indicador abaixo foi gerado como imagem separada.
                A planilha completa está anexada a este e-mail.
              </p>
            </td>
          </tr>

          <!-- IMAGENS -->
          <tr>
            <td style="padding:0 32px 18px;">
              {blocos_imagem}
            </td>
          </tr>

          <!-- DIVISOR -->
          <tr>
            <td style="padding:0 32px;">
              <hr style="border:none;border-top:1px solid #e8ecf0;margin:0;">
            </td>
          </tr>

          <!-- ASSINATURA -->
          <tr>
            <td style="padding:18px 32px 26px;">
              <table cellpadding="0" cellspacing="0">
                <tr>
                  <td style="width:4px;background-color:#1a3e5c;
                             border-radius:2px;">&nbsp;</td>
                  <td style="padding-left:14px;">
                    <span style="display:block;font-size:14px;font-weight:700;
                                 color:#1a3e5c;">{assinatura_nome}</span>
                    <span style="display:block;font-size:13px;color:#555555;
                                 margin-top:2px;">{assinatura_cargo}</span>
                    <span style="display:block;font-size:12px;color:#888888;
                                 margin-top:5px;">
                      {assinatura_email_sig} &nbsp;|&nbsp; {assinatura_tel}
                    </span>
                  </td>
                </tr>
              </table>
            </td>
          </tr>

          <!-- RODAPE -->
          <tr>
            <td style="background-color:#f0f4f8;padding:13px 32px;text-align:center;">
              <span style="font-size:11px;color:#aaaaaa;">
                E-mail gerado automaticamente pelo sistema de automação de
                indicadores. Por favor, não responda diretamente.
              </span>
            </td>
          </tr>

        </table>
      </td>
    </tr>
  </table>

</body>
</html>"""

    return html, caminhos, [caminho_anexo]


def enviar_email(
    html_body: str,
    assunto: str,
    email_remetente: str,
    destinatarios: list,
    imagem_paths: list[str],
    anexos: list[str],
) -> None:
    try:
        import win32com.client
    except ImportError as exc:
        raise RuntimeError(
            "PyWin32 é necessário para enviar pelo Outlook. "
            "Instale com: pip install pywin32"
        ) from exc

    print("[LOG] Criando e-mail no Outlook...")
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.Subject = assunto
    mail.To = ";".join(destinatarios)
    mail.HTMLBody = html_body

    if email_remetente:
        try:
            mail.SentOnBehalfOfName = email_remetente
        except Exception:
            print("[AVISO] Não foi possível configurar SentOnBehalfOfName. "
                  "O Outlook usará a conta padrão.")

    for idx, caminho in enumerate(imagem_paths, start=1):
        if os.path.exists(caminho):
            attachment = mail.Attachments.Add(os.path.abspath(caminho))
            try:
                prop_accessor = attachment.PropertyAccessor
                prop_accessor.SetProperty(
                    "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
                    f"tabela_{idx}",
                )
            except Exception:
                print(f"[AVISO] Não foi possível definir CID para: {caminho}")
        else:
            print(f"[AVISO] Imagem não encontrada, será pulada: {caminho}")

    for caminho in anexos:
        if os.path.exists(caminho):
            mail.Attachments.Add(os.path.abspath(caminho))
        else:
            print(f"[AVISO] Anexo não encontrado, será pulado: {caminho}")

    mail.Send()
    print(f"[LOG] E-mail enviado para: {', '.join(destinatarios)}")


def limpar_temporarios(caminhos: list[str]) -> None:
    for caminho in caminhos:
        try:
            if os.path.exists(caminho):
                os.remove(caminho)
                print(f"[LOG] Temporário removido: {caminho}")
        except OSError as exc:
            print(f"[AVISO] Não foi possível remover {caminho}: {exc}")


# ===========================================================================
# EXECUCAO PRINCIPAL
# ===========================================================================

def executar() -> None:
    inicio = time.time()
    print("=" * 60)
    print("  SISTEMA DE ENVIO AUTOMÁTICO DE INDICADORES — ELDORADO")
    print(f"  Início: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    print("=" * 60)

    data_atual = datetime.now().strftime("%d/%m/%Y")
    imagem_paths: list[str] = []

    try:
        if not EMAIL_REMETENTE:
            raise RuntimeError(
                "Variável de ambiente EMAIL_REMETENTE não definida. "
                "Configure o arquivo .env com o valor correto."
            )

        if not DESTINATARIOS:
            raise RuntimeError(
                "Variável de ambiente DESTINATARIOS não definida. "
                "Adicione os destinatários no arquivo .env, separados por vírgula."
            )

        # 1. Carregar e segmentar a planilha em sub-tabelas limpas
        tabelas = carregar_e_segmentar(EXCEL_PATH, SHEET_NAME)
        if not tabelas:
            raise RuntimeError("Nenhuma sub-tabela válida encontrada na planilha.")

        # 2. Gerar uma imagem PNG por sub-tabela
        imagens = gerar_todas_imagens(tabelas, IMAGE_BASE)
        imagem_paths = [c for _, c in imagens]

        # 3. Montar e-mail HTML
        html_body, imagem_paths, anexos = montar_email_html(
            assunto             = ASSUNTO,
            saudacao            = SAUDACAO,
            texto_corpo         = TEXTO_CORPO,
            data_atual          = data_atual,
            assinatura_nome     = ASSINATURA_NOME,
            assinatura_cargo    = ASSINATURA_CARGO,
            assinatura_email_sig= ASSINATURA_EMAIL,
            assinatura_tel      = ASSINATURA_TEL,
            imagens             = imagens,
            caminho_anexo       = EXCEL_PATH,
        )
        print("[LOG] E-mail montado com sucesso")

        # 4. Enviar via Outlook
        enviar_email(
            html_body       = html_body,
            assunto         = ASSUNTO,
            email_remetente = EMAIL_REMETENTE,
            destinatarios   = DESTINATARIOS,
            imagem_paths    = imagem_paths,
            anexos          = anexos,
        )

    except FileNotFoundError as exc:
        print(f"\n[ERRO] Arquivo não encontrado: {exc}")
        traceback.print_exc()
    except RuntimeError as exc:
        print(f"\n[ERRO] {exc}")
        traceback.print_exc()
    except Exception as exc:
        print(f"\n[ERRO INESPERADO] {exc}")
        traceback.print_exc()
    finally:
        limpar_temporarios(imagem_paths)
        duracao = time.time() - inicio
        print(f"\n[LOG] Duração total: {duracao:.1f}s")
        print("=" * 60)


if __name__ == "__main__":
    executar()


# ===========================================================================
# EXTENSOES FUTURAS (descomente para ativar)
# ===========================================================================

# A) Agendamento diário automático
# ─────────────────────────────────
# import schedule
# schedule.every().day.at("07:00").do(executar)
# while True:
#     schedule.run_pending()
#     time.sleep(60)

# B) Ler destinatários de aba do Excel
# ──────────────────────────────────────
# def carregar_destinatarios_excel(path):
#     df = pd.read_excel(path, sheet_name="Destinatarios")
#     return df["email"].dropna().str.strip().tolist()
# DESTINATARIOS = carregar_destinatarios_excel(EXCEL_PATH)

# C) Múltiplas abas — um e-mail por aba
# ──────────────────────────────────────
# ABAS = [
#     {"sheet": "Indicadores Semanal_Flash",    "assunto": "Flash Semanal"},
#     {"sheet": "Indicadores Mensal_Fechamento", "assunto": "Fechamento Mensal"},
# ]
# for cfg in ABAS:
#     SHEET_NAME = cfg["sheet"]
#     ASSUNTO    = cfg["assunto"]
#     executar()
#     time.sleep(5)

# D) Envio individual (um e-mail por destinatário)
# ─────────────────────────────────────────────────
# for dest in DESTINATARIOS:
#     executar_para([dest])
#     time.sleep(2)