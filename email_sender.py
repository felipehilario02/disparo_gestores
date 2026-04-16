"""
email_sender.py
===============

Envio de e-mails via SendGrid com suporte a:
  - Multiplas imagens inline (CID) definidas em config.INLINE_IMAGES
  - Conteudo HTML com links externos (ex: SharePoint, Power BI)

Nao envia anexos.
O acesso aos arquivos gerados deve ser feito exclusivamente via SharePoint,
por meio dos links presentes no template HTML.
"""

import base64
import mimetypes
import os

from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import (
    Attachment,
    ContentId,
    Disposition,
    FileContent,
    FileName,
    FileType,
    Mail,
)

from config import (
    INLINE_IMAGES,
    SEND_EMAILS,
    SENDGRID_API_KEY,
    SENDER_EMAIL,
)
from logger import get_logger


log = get_logger(__name__)


# ─────────────────────────────────────────────
# Helpers internos
# ─────────────────────────────────────────────

def _encode_file(file_path: str) -> tuple[str, str]:
    """
    Retorna (conteudo_base64, mime_type) de um arquivo.
    """
    with open(file_path, "rb") as f:
        encoded = base64.b64encode(f.read()).decode()

    mime, _ = mimetypes.guess_type(file_path)
    return encoded, mime or "application/octet-stream"


def _attach_inline_images(message: Mail) -> None:
    """
    Anexa todas as imagens declaradas em INLINE_IMAGES como inline (CID).

    Cada item de INLINE_IMAGES deve conter:
        {
          "cid":  "pbi_icon",
          "path": "caminho/para/imagem.png"
        }

    Arquivos nao encontrados sao ignorados com warning.
    """
    for img in INLINE_IMAGES:
        cid = img.get("cid")
        path = img.get("path")

        if not cid or not path or not os.path.isfile(path):
            log.warning(
                "Imagem inline nao encontrada — cid='%s' path='%s'",
                cid,
                path,
            )
            continue

        encoded, mime = _encode_file(path)

        attachment = Attachment(
            FileContent(encoded),
            FileName(os.path.basename(path)),
            FileType(mime),
            Disposition("inline"),
            ContentId(cid),
        )

        message.add_attachment(attachment)
        log.debug("Imagem inline anexada com sucesso: cid='%s'", cid)


# ─────────────────────────────────────────────
# Funcao publica
# ─────────────────────────────────────────────

def send_email_sendgrid(
    to_addr: str,
    subject: str,
    html_body: str,
) -> None:
    """
    Envia um e-mail via SendGrid utilizando HTML e imagens inline (CID).

    - Nao envia anexos
    - Utiliza links no corpo do e-mail para acesso aos arquivos no SharePoint

    Se SEND_EMAILS=False, apenas registra no log (modo teste).
    """
    if not SEND_EMAILS:
        log.info(
            "(TEST) E-mail NAO enviado para %s | Assunto: %s",
            to_addr,
            subject,
        )
        return

    api_key = (SENDGRID_API_KEY or os.getenv("SENDGRID_API_KEY", "")).strip()
    if not api_key:
        raise RuntimeError(
            "SENDGRID_API_KEY nao definido. Configure no arquivo .env."
        )

    message = Mail(
        from_email=SENDER_EMAIL,
        to_emails=to_addr,
        subject=subject,
        html_content=html_body,
    )

    # Anexa imagens inline (CID)
    _attach_inline_images(message)

    sg = SendGridAPIClient(api_key)
    response = sg.send(message)

    log.info(
        "E-mail enviado para %s | Status SendGrid: %d",
        to_addr,
        response.status_code,
    )