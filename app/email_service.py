import os
from app.logging_config import get_logger
from importlib import import_module

try:
    win32 = import_module("win32com.client")
except ImportError:
    win32 = None

logger = get_logger(__name__)

# Indica ao chamador se o Outlook COM está disponível nesta plataforma
OUTLOOK_DISPONIVEL = win32 is not None

# Destinatário padrão do RH (comunicado de devolução)
EMAIL_RH = "rh@somagrupo.com.br"

_STYLE = """
    font-family: 'Segoe UI', Arial, sans-serif;
    font-size: 14px;
    color: #1e293b;
"""

_TABLE_STYLE = """
    border-collapse: collapse;
    width: 100%;
    margin-top: 16px;
    font-size: 13px;
"""

_TH_STYLE = """
    background: #1e40af;
    color: white;
    padding: 8px 14px;
    text-align: left;
    font-weight: 600;
"""

_TD_LABEL_STYLE = """
    background: #f1f5f9;
    padding: 8px 14px;
    font-weight: 600;
    color: #334155;
    width: 200px;
    border-bottom: 1px solid #e2e8f0;
"""

_TD_VALUE_STYLE = """
    padding: 8px 14px;
    color: #1e293b;
    border-bottom: 1px solid #e2e8f0;
"""


def _row(label, value):
    v = value or "—"
    return (
        f'<tr>'
        f'<td style="{_TD_LABEL_STYLE}">{label}</td>'
        f'<td style="{_TD_VALUE_STYLE}">{v}</td>'
        f'</tr>'
    )


def _html_base(titulo, subtitulo, cor_topo, linhas_tabela, rodape_extra=""):
    return f"""
<div style="{_STYLE}">
  <div style="background: linear-gradient(135deg, #0f172a, {cor_topo}); padding: 20px 24px; border-radius: 8px 8px 0 0;">
    <div style="font-size: 18px; font-weight: 800; letter-spacing: 2px; color: white;">AZZAS<span style="color: #3b82f6;">.</span> Tech</div>
    <div style="font-size: 11px; color: #94a3b8; margin-top: 3px;">Sistema de Gestão de Devoluções de Equipamentos</div>
  </div>
  <div style="border: 1px solid #e2e8f0; border-top: none; padding: 24px; border-radius: 0 0 8px 8px; background: #ffffff;">
    <h2 style="font-size: 16px; font-weight: 700; color: #0f172a; margin: 0 0 4px 0;">{titulo}</h2>
    <p style="font-size: 12px; color: #64748b; margin: 0 0 16px 0;">{subtitulo}</p>
    <table style="{_TABLE_STYLE}">
      <thead>
        <tr>
          <th style="{_TH_STYLE}">Campo</th>
          <th style="{_TH_STYLE}">Valor</th>
        </tr>
      </thead>
      <tbody>
        {linhas_tabela}
      </tbody>
    </table>
    {rodape_extra}
    <div style="margin-top: 24px; padding-top: 16px; border-top: 1px solid #e2e8f0; font-size: 11px; color: #94a3b8;">
      Email gerado automaticamente pelo Sistema AZZAS TI &nbsp;·&nbsp; Não responda este email.
    </div>
  </div>
</div>
"""


def _get_outlook_mail():
    """Cria e retorna um item de email do Outlook."""
    if win32 is None:
        raise RuntimeError("win32com não disponível nesta plataforma.")
    try:
        outlook = win32.Dispatch("outlook.application")
        return outlook.CreateItem(0)
    except Exception as err:
        logger.exception("Erro ao iniciar Outlook")
        raise RuntimeError(
            "Não foi possível iniciar o Outlook. Verifique se está instalado e configurado."
        ) from err


def enviar_email(dados):
    """Envia email de devolução padrão (equipamento OK ou Pendente)."""
    if win32 is None:
        logger.warning("win32com não disponível; enviar_email ignorado.")
        return

    email_resp   = (dados.get("email_responsavel") or "").strip()
    gestor_email = (dados.get("gestor_email") or "").strip()

    mail = _get_outlook_mail()
    to_list = []
    if email_resp:
        to_list.append(email_resp)
    if gestor_email and gestor_email not in to_list:
        to_list.append(gestor_email)
    mail.To = "; ".join(to_list) if to_list else ""
    mail.Subject = f"[AZZAS TI] Devolução de Equipamento — {dados.get('patrimonio', '')} | {dados.get('nome', '')}"

    linhas = (
        _row("Data / Hora",           dados.get("data"))
        + _row("Setor",               dados.get("departamento"))
        + _row("Diretoria",           dados.get("diretoria"))
        + _row("Login de Rede",       dados.get("usuario"))
         + _row("Email para a Dell",  email_resp or None)
        + _row("Entregue Por",        dados.get("nome"))
        + _row("Recebido Por",        dados.get("recebido_por"))
        + _row("TAG (Patrimônio)", dados.get("patrimonio"))
        + _row("Tipo",             dados.get("tipo"))
        + _row("Marca",            dados.get("marca"))
        + _row("Modelo",           dados.get("modelo"))
        + _row("Processador",      dados.get("processador"))
        + _row("Memória",          dados.get("memoria"))
        + _row("Armazenamento",    dados.get("armazenamento"))
        + _row("Possui Carregador",dados.get("possui_carregador"))
        + _row("Unidade",          dados.get("unidade"))
        + _row("Motivo",           dados.get("motivo"))
        + _row("Observações",      dados.get("observacoes"))
        + _row("Movido p/ Estoque",dados.get("movido_para_estoque"))
        + _row("Situação",         dados.get("status"))
    )

    mail.HTMLBody = _html_base(
        titulo="Registro de Devolução de Equipamento",
        subtitulo="Um novo equipamento foi devolvido ao setor de TI.",
        cor_topo="#1a2e6c",
        linhas_tabela=linhas,
    )

    try:
        mail.Display()
        logger.info("Email de devolução aberto para %s", dados.get("nome"))
    except Exception as err:
        logger.exception("Erro ao exibir email em enviar_email")
        raise RuntimeError("Não foi possível exibir o email no Outlook.") from err


def email_dano(dados):
    """Envia email de alerta para equipamento danificado, com foto anexada se disponível."""
    if win32 is None:
        logger.warning("win32com não disponível; email_dano ignorado.")
        return

    email_resp   = (dados.get("email_responsavel") or "").strip()
    gestor_email = (dados.get("gestor_email") or "").strip()

    mail = _get_outlook_mail()
    to_list = []
    if email_resp:
        to_list.append(email_resp)
    if gestor_email and gestor_email not in to_list:
        to_list.append(gestor_email)
    mail.To = "; ".join(to_list) if to_list else ""
    mail.Subject = f"[AZZAS TI] ⚠ EQUIPAMENTO DANIFICADO — {dados.get('patrimonio', '')} | {dados.get('nome', '')}"

    alerta = (
        '<div style="background:#fee2e2; border-left: 4px solid #dc2626; '
        'padding: 12px 16px; border-radius: 4px; margin-top: 16px; '
        'color: #7f1d1d; font-size: 13px; font-weight: 600;">'
        '&#9888; Este equipamento foi registrado como DANIFICADO. '
        'Verificar necessidade de abertura de chamado técnico.'
        '</div>'
    )

    linhas = (
        _row("Data / Hora",           dados.get("data"))
        + _row("Setor",               dados.get("departamento"))
        + _row("Diretoria",           dados.get("diretoria"))
        + _row("Login de Rede",       dados.get("usuario"))
        + _row("Email para a Dell",  email_resp or None)
        + _row("Entregue Por",        dados.get("nome"))
        + _row("Recebido Por",        dados.get("recebido_por"))
        + _row("TAG (Patrimônio)",    dados.get("patrimonio"))
        + _row("Tipo",                dados.get("tipo"))
        + _row("Marca",               dados.get("marca"))
        + _row("Modelo",              dados.get("modelo"))
        + _row("Processador",         dados.get("processador"))
        + _row("Memória",             dados.get("memoria"))
        + _row("Armazenamento",       dados.get("armazenamento"))
        + _row("Possui Carregador",   dados.get("possui_carregador"))
        + _row("Unidade",             dados.get("unidade"))
        + _row("Motivo do Dano",      dados.get("motivo"))
        + _row("Observações",         dados.get("observacoes"))
        + _row("Movido p/ Estoque",   dados.get("movido_para_estoque"))
    )

    mail.HTMLBody = _html_base(
        titulo="&#9888; Equipamento Danificado — Ação Necessária",
        subtitulo="Um equipamento com dano foi registrado. Verifique as informações abaixo.",
        cor_topo="#7f1d1d",
        linhas_tabela=linhas,
        rodape_extra=alerta,
    )

    # Anexar foto de dano se disponível
    if dados.get("foto"):
        try:
            foto_abs = os.path.abspath(dados["foto"])
            if os.path.exists(foto_abs):
                mail.Attachments.Add(foto_abs)
                logger.info("Foto de dano anexada: %s", foto_abs)
        except Exception as err:
            logger.warning("Não foi possível anexar a foto: %s", err)

    try:
        mail.Display()
        logger.info("Email de dano aberto para %s", dados.get("nome"))
    except Exception as err:
        logger.exception("Erro ao exibir email em email_dano")
        raise RuntimeError("Não foi possível exibir o email de dano no Outlook.") from err


# ── Destinatário da Dell para cotação de reparo ──────────────────
EMAIL_DELL = os.getenv("EMAIL_DELL", "pon.commercial@dell.com")
CNPJ_EMPRESA = os.getenv("CNPJ_EMPRESA", "09.611.669/0009-41")


def email_cotacao_dell(dados):
    """Cria rascunho no Outlook com cotação de reparo endereçado à Dell.

    Chamado automaticamente quando o equipamento devolvido é da marca Dell
    e está marcado como Danificado.
    """
    if win32 is None:
        logger.warning("win32com não disponível; email_cotacao_dell ignorado.")
        return

    modelo        = dados.get("modelo", "").strip()
    serial        = dados.get("serial", dados.get("patrimonio", "")).strip()
    patrimonio    = dados.get("patrimonio", "").strip()
    email_resp    = (dados.get("email_responsavel") or "").strip()
    gestor_email  = (dados.get("gestor_email") or "").strip()
    cotacao_para  = dados.get("recebido_por") or dados.get("nome") or ""
    if email_resp:
        cotacao_para = f"{cotacao_para} <{email_resp}>".strip(" <>")
        cotacao_para = cotacao_para if cotacao_para != f"<{email_resp}>" else email_resp

    mail = _get_outlook_mail()
    mail.To      = EMAIL_DELL
    cc_list      = []
    if email_resp and email_resp not in cc_list:
        cc_list.append(email_resp)
    if gestor_email and gestor_email not in cc_list:
        cc_list.append(gestor_email)
    if cc_list:
        mail.CC = "; ".join(cc_list)
    mail.Subject = (
        f"Solicitação de Cotação de Reparo — Dell {modelo} | "
        f"Service Tag: {serial} | TAG: {patrimonio}"
    )

    corpo = (
        f"Boa tarde, prezados.\n\n"
        f"Solicito cotação de reparo do equipamento Dell {modelo} "
        f"- Service Tag: {serial}\n\n"
        f"Cotação para: {cotacao_para}\n\n"
        f"Segue em anexo as imagens.\n\n\n\n"
        f"TAG INTERNA {patrimonio}\n"
        f"O chamado em questão é para uma empresa de CNPJ : {CNPJ_EMPRESA}"
    )
    mail.Body = corpo

    # Anexar foto de dano se disponível
    if dados.get("foto"):
        try:
            foto_abs = os.path.abspath(dados["foto"])
            if os.path.exists(foto_abs):
                mail.Attachments.Add(foto_abs)
                logger.info("Foto de dano anexada ao rascunho Dell: %s", foto_abs)
        except Exception as err:
            logger.warning("Não foi possível anexar a foto ao rascunho Dell: %s", err)

    try:
        # Display() abre como rascunho editável — o usuário revisa antes de enviar
        mail.Display()
        logger.info(
            "Rascunho de cotação Dell aberto — modelo: %s | serial: %s",
            modelo, serial,
        )
    except Exception as err:
        logger.exception("Erro ao exibir rascunho Dell em email_cotacao_dell")
        raise RuntimeError("Não foi possível abrir o rascunho no Outlook.") from err


def enviar_email_rh(dados, para=None):
    """Abre rascunho no Outlook com comunicado ao RH sobre devolução do colaborador."""
    if win32 is None:
        logger.warning("win32com não disponível; enviar_email_rh ignorado.")
        return

    destino      = para or EMAIL_RH
    email_resp   = (dados.get("email_responsavel") or "").strip()
    gestor_email = (dados.get("gestor_email") or "").strip()

    mail = _get_outlook_mail()
    mail.To = destino
    cc_list = []
    if email_resp and email_resp not in cc_list:
        cc_list.append(email_resp)
    if gestor_email and gestor_email not in cc_list:
        cc_list.append(gestor_email)
    if cc_list:
        mail.CC = "; ".join(cc_list)
    mail.Subject = (
        f"[AZZAS TI] Devolução de Equipamento — {dados.get('nome', '')} | "
        f"{dados.get('matricula', '')} | {dados.get('departamento', '')}"
    )

    linhas = (
        _row("Data / Hora",       dados.get("data"))
        + _row("Nome Completo",   dados.get("nome"))
        + _row("Matrícula",       dados.get("matricula"))
        + _row("Setor",           dados.get("departamento"))
        + _row("Diretoria",       dados.get("diretoria"))
        + _row("Login de Rede",   dados.get("usuario"))
        + _row("Unidade",         dados.get("unidade"))
        + _row("Tipo",            dados.get("tipo"))
        + _row("Marca",           dados.get("marca"))
        + _row("Modelo",          dados.get("modelo"))
        + _row("TAG (Patrimônio)", dados.get("patrimonio"))
        + _row("Serial / S.Tag",  dados.get("serial"))
        + _row("Situação",        dados.get("status"))
        + _row("Motivo",          dados.get("motivo"))
        + _row("Recebido Por",    dados.get("recebido_por"))
        + _row("Observações",     dados.get("observacoes"))
    )

    mail.HTMLBody = _html_base(
        titulo="Devolução de Equipamento — Comunicado ao RH",
        subtitulo=f"O colaborador {dados.get('nome', '')} devolveu um equipamento ao setor de TI.",
        cor_topo="#1e3a8a",
        linhas_tabela=linhas,
    )

    try:
        mail.Display()
        logger.info("Email RH aberto para %s (%s)", dados.get("nome"), destino)
    except Exception as err:
        logger.exception("Erro ao exibir email RH")
        raise RuntimeError("Não foi possível exibir o email para o RH no Outlook.") from err
