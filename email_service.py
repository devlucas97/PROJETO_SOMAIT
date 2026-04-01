import logging
import sys

try:
    import win32com.client as win32
except ImportError:
    win32 = None
    # Módulo win32com não disponível em Linux; operação de email será mock/sem efeito.

logger = logging.getLogger(__name__)
if not logger.hasHandlers():
    handler = logging.StreamHandler()
    handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(name)s: %(message)s"))
    logger.addHandler(handler)
logger.setLevel(logging.INFO)
 
def enviar_email(dados):
    if win32 is None:
        logger.warning("win32com não disponível; enviar_email ignorado.")
        return

    try:
        outlook = win32.Dispatch("outlook.application")
        mail = outlook.CreateItem(0)
    except Exception as err:
        logger.exception("Erro ao iniciar Outlook em enviar_email")
        raise RuntimeError("Não foi possível iniciar o Outlook. Verifique se está instalado e configurado.") from err
 
    mail.Subject = f"Devolução de Notebook - {dados['nome']}"
 
    try:
        mail.HTMLBody = f"""
<h2 style="color:#2E86C1;">Devolução de Notebook</h2>
<table border="1" cellpadding="6">
<tr><td><b>Colaborador</b></td><td>{dados['nome']}</td></tr>
<tr><td><b>Matrícula</b></td><td>{dados['matricula']}</td></tr>
<tr><td><b>Departamento</b></td><td>{dados['departamento']}</td></tr>
<tr><td><b>Patrimônio</b></td><td>{dados['patrimonio']}</td></tr>
<tr><td><b>Modelo</b></td><td>{dados['modelo']}</td></tr>
<tr><td><b>Serial</b></td><td>{dados['serial']}</td></tr>
<tr><td><b>Status</b></td><td>{dados['status']}</td></tr>
</table>
        """
        mail.Display()
        logger.info("Email de devolução aberto com sucesso para %s", dados.get('nome'))
    except Exception as err:
        logger.exception("Erro ao montar/exibir email em enviar_email")
        raise RuntimeError("Não foi possível preparar o email no Outlook.") from err
 
def email_dano(dados):
    if win32 is None:
        logger.warning("win32com não disponível; email_dano ignorado.")
        return

    try:
        outlook = win32.Dispatch("outlook.application")
        mail = outlook.CreateItem(0)
    except Exception as err:
        raise RuntimeError("Não foi possível iniciar o Outlook. Verifique se está instalado e configurado.") from err
 
    mail.To = "lucas.wagner@somagrupo.com.br"
    mail.Subject = "Notebook danificado - Abertura de chamado"
 
    try:
        mail.Body = f"""
Notebook com dano:
 
Patrimônio: {dados['patrimonio']}
Modelo: {dados['modelo']}
Serial: {dados['serial']}
"""
        # Anexar foto se disponível
        if dados.get("foto"):
            mail.Attachments.Add(dados["foto"])
 
        mail.Display()
        logger.info("Email de dano aberto com sucesso para %s", dados.get('nome'))
    except Exception as err:
        logger.exception("Erro ao montar/exibir email em email_dano")
        raise RuntimeError("Não foi possível preparar o email de dano no Outlook.") from err