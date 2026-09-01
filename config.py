"""
config.py
=========
Centraliza todas as configurações do pipeline de despesas.

Secrets (API key, flag de envio) são lidos do arquivo .env via python-dotenv.
O mapa de e-mails é carregado de emails.json — edite esse arquivo para
adicionar/remover destinatários sem tocar em código.

Hierarquia de precedência para variáveis de ambiente:
1. Variável já definida no sistema operacional
2. Arquivo .env na raiz do projeto
3. Default hardcoded abaixo (somente para valores não-sensíveis)
"""

import json
import os
from pathlib import Path
from dotenv import load_dotenv

# ======================================================
# ENV
# ======================================================

load_dotenv(Path(__file__).parent / ".env", override=False)

# ======================================================
# CAMINHOS
# ======================================================

BASE_DIR = os.getenv(
    "BASE_DIR",
    r"C:\Users\jcostaya\OneDrive - Johnson Controls\Documentos\Despesas DB\DB Geral",
)

GESTORES_DIR = os.getenv(
    "GESTORES_DIR",
    r"C:\Users\jcostaya\Johnson Controls\ADT - Equipe Financeiro - Documentos\Despesas - Gestores",
)

INPUT_FILE = "Despesas Gerais.xlsx"
SHEET_NAME = 0
FORCE_RESP_COL_NAME = "RESPONSÁVEL"
DATE_FMT = "%Y-%m-%d"

# ======================================================
# POWER BI
# ======================================================

BI_LINK = (
    "https://app.powerbi.com/Redirect?"
    "action=OpenReport"
    "&appId=e362eba8-af56-4e8d-83b8-95c780bb6504"
    "&reportObjectId=1f7a083f-188b-49ea-a0c7-8912fb644807"
)

# ======================================================
# SENDGRID
# ======================================================

SENDGRID_API_KEY = os.getenv("SENDGRID_API_KEY", "").strip()
SENDER_EMAIL = os.getenv("SENDER_EMAIL", "financeiro@adt.com.br")

_send_raw = os.getenv("SEND_EMAILS", "true").strip().lower()
SEND_EMAILS: bool = _send_raw in ("true", "1", "yes")

# ======================================================
# IMAGENS INLINE DO E-MAIL
# ======================================================

LOGO_PATH = os.getenv(
    "LOGO_PATH",
    r"C:\Users\jcostaya\Downloads\Logo ADT Nova 3.png",
)

LOGO_CONTENT_ID = "logo-adt"

INLINE_IMAGES: list[dict[str, str]] = [
    {
        "cid": "hero_bg",
        "path": os.getenv(
            "HERO_BANNER_PATH",
            r"C:\Users\jcostaya\OneDrive - Johnson Controls\Documentos\Despesas DB\Mail\Mail ADT Background.png",
        ),
    },
    {
        "cid": "pbi_icon",
        "path": os.getenv(
            "POWERBI_ICON_PATH",
            r"C:\Users\jcostaya\Downloads\icons8-power-bi-2021-48.png",
            
        ),
    },
    {
        "cid": "excel_icon",
        "path": os.getenv(
            "EXCEL_ICON_PATH",
            r"C:\Users\jcostaya\OneDrive - Johnson Controls\Documentos\Despesas DB\Mail\excel_icon.png",
            
        ),
    },
]

# ======================================================
# ASSUNTO
# ======================================================

SUBJECT_TEMPLATE = (
    "Relatorio de Despesas - "
    "{responsavel} ({mes_anterior_pt} - FY{fy:02d})"
)

# ======================================================
# DESTINATÁRIOS
# ======================================================

RESP_GERAL = {"Hebert Hiroyuki", "Rafael Leite"}
EXTRA_GERAL = {"Izabella Barretti"}

FULL_BASE_RECIPIENTS = RESP_GERAL | EXTRA_GERAL

ALLOWED_COMMENTS_RECIPIENTS = {
    "Rafael Leite",
    "Hebert Hiroyuki",
    "Izabella Barretti",
}

# ======================================================
# MAPA DE E-MAILS
# ======================================================

def _load_email_map() -> dict[str, str]:
    """
    Carrega o mapa nome -> e-mail de emails.json.
    Retorna dicionário vazio com aviso se o arquivo
    não existir ou for inválido.
    """
    path = Path(__file__).parent / "emails.json"

    if not path.exists():
        print(f"[WARN] emails.json não encontrado em {path}")
        return {}

    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        return data.get("destinatarios", {})
    except Exception as exc:
        print(f"[WARN] Falha ao carregar emails.json: {exc}")
        return {}

EMAIL_MAP: dict[str, str] = _load_email_map()

# ======================================================
# COMENTÁRIOS
# ======================================================

COMENTARIOS_SHEET = "Comentario"

COMENTARIOS_FILE = os.getenv(
    "COMENTARIOS_FILE",
    r"C:\Users\jcostaya\OneDrive - Johnson Controls"
    r"\Documentos\Despesas DB\DB Geral\Comentários DB.xlsx",
)

# ======================================================
# DEBUG
# ======================================================

DEBUG_SAMPLE_CHECK: bool = os.getenv(
    "DEBUG_SAMPLE_CHECK", "false"
).lower() in ("true", "1")

# ======================================================
# CALENDÁRIO FISCAL
# ======================================================

FY_START_MONTH = 10  # Outubro

MESES_PT = [
    "Janeiro", "Fevereiro", "Março", "Abril",
    "Maio", "Junho", "Julho", "Agosto",
    "Setembro", "Outubro", "Novembro", "Dezembro",
]

# ======================================================
# TEMPLATE HTML
# ======================================================

BODY_TEMPLATE = """
  <!doctype html>
  <html lang="pt-br">
  <head>
    <meta charset="utf-8">
    <title>Notificação – Base de Despesas</title>
  </head>

  <body style="margin:0; padding:0; background-color:#f4f4f4;">

  <!-- WRAPPER -->
  <table width="100%" cellpadding="0" cellspacing="0" bgcolor="#f4f4f4">
    <tr>
      <td align="center" style="padding:30px 0;">

        <!-- CONTAINER -->
        <table width="640" cellpadding="0" cellspacing="0" style="background:#ffffff; border-radius:12px; overflow:hidden; font-family:Tahoma, Arial, Helvetica, sans-serif;">

          <!-- HERO -->
          <tr>
            <td align="center">
              <img src="cid:hero_bg"
                  alt="Notificação de despesas"
                  width="640"
                  style="display:block; width:640px; max-width:100%; height:auto; border:0; margin:0;" />
            </td>
          </tr>

          <!-- CONTENT -->
          <tr>
            <td style="padding:32px; text-align:center; color:#333333;">

              <h2 style="margin:0 0 16px; font-size:24px; font-weight:normal;">
                Olá {responsavel},
              </h2>

              <p style="margin:0 0 16px; font-size:17px; line-height:1.55;">
                Informamos que a base de despesas do mês já está disponível.<br>
              </p>

              <p style="margin:0 0 20px; font-size:16px;">
                Para acessar o dashboard e a base de dados completa:
              </p>

  <!-- ICONES POWER BI + EXCEL -->
  <table cellpadding="0" cellspacing="0" align="center">
    <tr>
      <!-- POWER BI -->
      <td align="center" style="padding-right:10px;">
        <a href="https://app.powerbi.com/Redirect?action=OpenReport&amp;appId=e362eba8-af56-4e8d-83b8-95c780bb6504&amp;reportObjectId=1f7a083f-188b-49ea-a0c7-8912fb644807&amp;ctid=a1f1e214-7ded-45b6-81a1-9e8ae3459641&amp;reportPage=e1c594aa699aadc09741&amp;pbi_source=appShareLink"
          target="_blank">
          <img src="cid:pbi_icon"
              alt="Power BI"
              width="48"
              height="48"
              style="display:block; border:0;" />
        </a>
      </td>

      <!-- EXCEL -->
      <td align="center" style="padding-left:10px;">
        <a href="https://apps.jci.com/:f:/r/sites/ADT-EquipeFinanceiro/Shared%20Documents/Despesas%20-%20Gestores?csf=1&web=1&e=ik3kX4"
          target="_blank">
          <img src="cid:excel_icon"
              alt="Excel"
              width="48"
              height="48"
              style="display:block; border:0;" />
        </a>
      </td>
    </tr>
  </table>

              <!-- ESPAÇADOR (SUBSTITUI MARGIN-TOP) -->
              <table width="100%" cellpadding="0" cellspacing="0">
                <tr>
                  <td height="16" style="font-size:0; line-height:0; background:#ffffff">
                    &nbsp;
                  </td>
                </tr>
              </table>


              <!-- UPDATE BOX -->
  <table width="100%" cellpadding="0" cellspacing="0" style="margin-top:6px; background:#f2f2f2; border-radius:12px;">
    <tr>
      <td style="padding:20px 22px; font-size:16px; color:#333; text-align:left;">

        <div style="text-align:center;">
          <strong style="font-size:18px;">Atualizações 🛠️</strong><br><br>
        </div>

        <!-- BLOCO CENTRALIZADO -->
        <table width="85%" cellpadding="0" cellspacing="0" align="center">
          <tr>
            <td width="20" valign="top" style="padding-left:8px;">✅</td>
            <td style="padding-left:6px;">Correções de performance.</td>
          </tr>
        </table>

      </td>
    </tr>
  </table>


              <!-- FOOTER -->
              <p style="margin:28px 0 0; font-size:15px; color:#555555;">
                Qualquer dúvida, entre em contato com
                <a href="mailto:izabella.barretti@jci.com"
                  style="color:#0066cc; text-decoration:none;">
                  izabella.barretti@jci.com
                </a>
              </p>

            </td>
          </tr>

        </table>
        <!-- /CONTAINER -->

      </td>
    </tr>
  </table>
  <!-- /WRAPPER -->

  </body>
  </html>
"""