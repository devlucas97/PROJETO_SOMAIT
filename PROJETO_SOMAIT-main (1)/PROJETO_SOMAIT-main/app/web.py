import json
import os
import secrets
import time
import functools
from collections import Counter
from flask import (
    Flask, render_template, request, redirect, url_for,
    flash, send_from_directory, session, abort,
)
from werkzeug.security import check_password_hash, generate_password_hash
from werkzeug.utils import secure_filename
from app.logging_config import get_logger
from app.runtime_paths import ensure_runtime_dir, get_bundle_path, iter_config_paths
from app.database import (
    criar, inserir, listar, buscar_por_id,
    estatisticas, listar_filtrado, atualizar, excluir,
    registrar_email_enviado, atualizar_chamado_dell,
    sincronizar_planilha_completa, diagnosticar_config_planilha,
    DELL_WF_STATUSES,
)

# Import com fallback para compatibilidade Linux/Windows
try:
    from app.email_service import enviar_email, email_dano, email_cotacao_dell, enviar_email_rh, OUTLOOK_DISPONIVEL, EMAIL_RH
except ImportError:
    OUTLOOK_DISPONIVEL = False
    EMAIL_RH = "rh@somagrupo.com.br"
    enviar_email = email_dano = email_cotacao_dell = enviar_email_rh = None

# Configuração de logging
logger = get_logger(__name__)

# ── Configuração do app ──────────────────────────────────────────
app = Flask(
    __name__,
    template_folder=get_bundle_path("app", "templates"),
    static_folder=get_bundle_path("app", "static"),
)
app.secret_key = os.getenv("FLASK_SECRET_KEY", "chave_secreta_dev_2026")

UPLOAD_MAX_MB = int(os.getenv("UPLOAD_MAX_MB", "5"))
app.config["MAX_CONTENT_LENGTH"] = UPLOAD_MAX_MB * 1024 * 1024

ALLOWED_IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".webp"}

UPLOAD_DIR = ensure_runtime_dir("uploads")

BRAND_NAME = "SOMALABS"
BRAND_AREA = "Operações de TI"
BRAND_PRODUCT = "Gestão de Devoluções e Ativos"
DEFAULT_APP_USERNAME = "admin"
DEFAULT_APP_PASSWORD = "azzas2026"

# Credenciais de acesso (variáveis de ambiente, com fallback seguro para dev)
APP_USERNAME = os.getenv("APP_USERNAME", DEFAULT_APP_USERNAME)
_raw_password = os.getenv("APP_PASSWORD", DEFAULT_APP_PASSWORD)
# Armazena hash da senha em memória — nunca compare senhas em texto puro
APP_PASSWORD_HASH = generate_password_hash(_raw_password)
del _raw_password

# Paginação
REGISTROS_POR_PAGINA = int(os.getenv("REGISTROS_POR_PAGINA", "30"))

# ── Proteção CSRF ────────────────────────────────────────────────
def _gerar_csrf_token():
    if "_csrf_token" not in session:
        session["_csrf_token"] = secrets.token_hex(32)
    return session["_csrf_token"]


@app.before_request
def _verificar_csrf():
    if request.method in ("GET", "HEAD", "OPTIONS"):
        return
    token = request.form.get("_csrf_token") or request.headers.get("X-CSRF-Token")
    if not token or token != session.get("_csrf_token"):
        logger.warning("CSRF token inválido para %s %s", request.method, request.path)
        abort(403)


@app.context_processor
def inject_brand_context():
    return {
        "brand_name": BRAND_NAME,
        "brand_area": BRAND_AREA,
        "brand_product": BRAND_PRODUCT,
        "csrf_token": _gerar_csrf_token(),
    }


@app.template_filter("basename")
def basename_filter(path):
    return os.path.basename(str(path)) if path else ""


def _cfg():
    """Carrega config.json do projeto; retorna dict vazio se não existir."""
    for config_file in iter_config_paths():
        if not os.path.exists(config_file):
            continue
        try:
            with open(config_file, encoding="utf-8") as f:
                return json.load(f)
        except Exception as err:
            logger.warning("Falha ao ler config.json na web: %s", err)
    return {}


def _log_security_startup_warnings():
    if APP_USERNAME == DEFAULT_APP_USERNAME:
        logger.warning("APP_USERNAME está com o valor padrão. Defina credenciais próprias para uso contínuo.")
    password_from_env = os.getenv("APP_PASSWORD")
    if password_from_env is None:
        logger.warning("APP_PASSWORD não foi definido no ambiente. A aplicação está usando a senha padrão de desenvolvimento.")
    elif password_from_env == DEFAULT_APP_PASSWORD:
        logger.warning("APP_PASSWORD está com o valor padrão de desenvolvimento. Troque a senha antes de compartilhar o acesso.")
    if app.secret_key == "chave_secreta_dev_2026":
        logger.warning("FLASK_SECRET_KEY está com o valor padrão de desenvolvimento. Gere uma chave própria para evitar sessões previsíveis.")


def _salvar_cfg(cfg):
    with open(_CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


def _percentual(parte, total):
    return int(round((parte / total) * 100)) if total else 0


def _montar_painel_executivo(registros, busca="", status_filtro="", data_inicio="", data_fim=""):
    total = len(registros)
    status_counter = Counter((r.get("status") or "Sem status") for r in registros)
    setor_counter = Counter((r.get("departamento") or "Não informado") for r in registros)
    marca_counter = Counter((r.get("marca") or "Não informada") for r in registros)

    danificados = status_counter.get("Danificado", 0)
    pendentes = status_counter.get("Pendente", 0)
    em_fluxo_dell = sum(status_counter.get(s, 0) for s in DELL_WF_STATUSES)
    pendencias = danificados + pendentes + em_fluxo_dell
    equipamentos_dell = sum(1 for r in registros if (r.get("marca") or "").strip().lower() == "dell")
    com_foto = sum(1 for r in registros if r.get("foto"))
    sem_carregador = sum(1 for r in registros if (r.get("possui_carregador") or "").strip().lower() == "não")
    movidos_estoque = sum(1 for r in registros if (r.get("movido_para_estoque") or "").strip().lower() == "sim")
    gestores_em_cc = sum(1 for r in registros if (r.get("gestor_email") or "").strip())
    chamados_dell = sum(1 for r in registros if (r.get("chamado_dell") or "").strip())
    concluidos = status_counter.get("OK", 0) + status_counter.get("Concluído", 0)

    status_breakdown = []
    ordered_statuses = [
        "OK",
        "Danificado",
        "Pendente",
        "Aguardando Cotação",
        "Cotação Recebida",
        "Reparo Aprovado",
        "Em Reparo",
        "Concluído",
    ]
    tone_by_status = {
        "OK": "ok",
        "Danificado": "risk",
        "Pendente": "warning",
        "Aguardando Cotação": "info",
        "Cotação Recebida": "info",
        "Reparo Aprovado": "info",
        "Em Reparo": "info",
        "Concluído": "ok",
    }

    for label in ordered_statuses:
        count = status_counter.get(label, 0)
        if count:
            status_breakdown.append({
                "label": label,
                "count": count,
                "percent": _percentual(count, total),
                "tone": tone_by_status.get(label, "neutral"),
            })

    for label, count in status_counter.most_common():
        if label not in ordered_statuses:
            status_breakdown.append({
                "label": label,
                "count": count,
                "percent": _percentual(count, total),
                "tone": "neutral",
            })

    filtros_ativos = any([busca, status_filtro, data_inicio, data_fim])
    return {
        "total": total,
        "pendencias": pendencias,
        "integridade_pct": _percentual(status_counter.get("OK", 0), total),
        "conclusao_pct": _percentual(concluidos, total),
        "dell_pct": _percentual(equipamentos_dell, total),
        "base_label": "Base filtrada" if filtros_ativos else "Base consolidada",
        "ultima_atualizacao": registros[0].get("data") if registros else None,
        "equipamentos_dell": equipamentos_dell,
        "com_foto": com_foto,
        "sem_carregador": sem_carregador,
        "movidos_estoque": movidos_estoque,
        "gestores_em_cc": gestores_em_cc,
        "chamados_dell": chamados_dell,
        "top_setores": setor_counter.most_common(4),
        "top_marcas": marca_counter.most_common(4),
        "status_breakdown": status_breakdown,
    }


criar()
_log_security_startup_warnings()


# ── Autenticação ─────────────────────────────────────────────────
LOGIN_MAX_TENTATIVAS = int(os.getenv("LOGIN_MAX_TENTATIVAS", "5"))
LOGIN_BLOQUEIO_SEGUNDOS = int(os.getenv("LOGIN_BLOQUEIO_SEGUNDOS", "60"))
_login_tentativas: dict[str, list] = {}  # ip -> [timestamps]


def _login_bloqueado(ip: str) -> bool:
    """Retorna True se o IP excedeu o limite de tentativas."""
    agora = time.monotonic()
    tentativas = _login_tentativas.get(ip, [])
    # Manter apenas tentativas dentro da janela
    tentativas = [t for t in tentativas if agora - t < LOGIN_BLOQUEIO_SEGUNDOS]
    _login_tentativas[ip] = tentativas
    return len(tentativas) >= LOGIN_MAX_TENTATIVAS


def _registrar_tentativa_login(ip: str):
    agora = time.monotonic()
    _login_tentativas.setdefault(ip, []).append(agora)


def login_required(f):
    @functools.wraps(f)
    def decorated(*args, **kwargs):
        if not session.get("autenticado"):
            flash("Faça login para continuar.", "warning")
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated


@app.route("/login", methods=["GET", "POST"])
def login():
    if session.get("autenticado"):
        return redirect(url_for("index"))
    if request.method == "POST":
        ip = request.remote_addr or "unknown"
        if _login_bloqueado(ip):
            flash(f"Muitas tentativas. Aguarde {LOGIN_BLOQUEIO_SEGUNDOS}s e tente novamente.", "error")
            logger.warning("Login bloqueado por rate limit: %s", ip)
            return render_template("login.html"), 429
        usuario = request.form.get("usuario", "").strip()
        senha   = request.form.get("senha", "")
        if usuario == APP_USERNAME and check_password_hash(APP_PASSWORD_HASH, senha):
            session["autenticado"] = True
            session["usuario_logado"] = usuario
            _login_tentativas.pop(ip, None)
            logger.info("Login bem-sucedido: %s", usuario)
            return redirect(url_for("index"))
        _registrar_tentativa_login(ip)
        flash("Usuário ou senha incorretos.", "error")
    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    flash("Sessão encerrada.", "info")
    return redirect(url_for("login"))


# ── Dashboard / Listagem ─────────────────────────────────────────
@app.route("/")
@login_required
def index():
    try:
        busca        = request.args.get("busca", "").strip()
        search_id    = request.args.get("search_id", "").strip()
        status_filtro = request.args.get("status", "").strip()
        data_inicio  = request.args.get("data_inicio", "").strip()
        data_fim     = request.args.get("data_fim", "").strip()
        pagina       = max(1, int(request.args.get("pagina", 1)))

        todos = listar_filtrado(
            status=status_filtro or None,
            busca=busca or None,
            data_inicio=data_inicio or None,
            data_fim=data_fim or None,
        )

        if search_id:
            if search_id.isdigit():
                id_busca = int(search_id)
                todos = [registro for registro in todos if registro.get("id") == id_busca]
                busca = busca or search_id
            else:
                flash("O filtro por ID aceita apenas números inteiros.", "warning")

        stats = estatisticas()
        painel_exec = _montar_painel_executivo(
            todos,
            busca=busca,
            status_filtro=status_filtro,
            data_inicio=data_inicio,
            data_fim=data_fim,
        )

        total_registros = len(todos)
        total_paginas   = max(1, (total_registros + REGISTROS_POR_PAGINA - 1) // REGISTROS_POR_PAGINA)
        pagina          = min(pagina, total_paginas)
        inicio          = (pagina - 1) * REGISTROS_POR_PAGINA
        dados           = todos[inicio: inicio + REGISTROS_POR_PAGINA]

        logger.info("Listagem p.%d/%d (%d registros)", pagina, total_paginas, total_registros)
        return render_template(
            "index.html",
            dados=dados,
            busca=busca,
            search_id=search_id,
            status_filtro=status_filtro,
            data_inicio=data_inicio,
            data_fim=data_fim,
            stats=stats,
            painel_exec=painel_exec,
            pagina=pagina,
            total_paginas=total_paginas,
            total_registros=total_registros,
        )
    except Exception:
        logger.exception("Erro ao listar devoluções")
        flash("Erro ao carregar dados", "error")
        return render_template(
            "index.html",
            dados=[], busca="", search_id="", status_filtro="", data_inicio="", data_fim="",
            stats={"total": 0, "ok": 0, "danificado": 0, "pendente": 0},
            painel_exec=_montar_painel_executivo([]),
            pagina=1, total_paginas=1, total_registros=0,
        )


# ── Nova devolução ───────────────────────────────────────────────
@app.route("/nova", methods=["GET", "POST"])
@login_required
def nova_devolucao():
    if request.method == "POST":
        try:
            dados = {
                "usuario":             request.form.get("usuario"),
                "nome":                request.form.get("nome"),
                "matricula":           request.form.get("matricula"),
                "departamento":        request.form.get("departamento"),
                "patrimonio":          request.form.get("patrimonio"),
                "modelo":              request.form.get("modelo"),
                "serial":              request.form.get("serial"),
                "status":              request.form.get("status"),
                "motivo":              request.form.get("motivo"),
                "foto":                None,
                "diretoria":           request.form.get("diretoria"),
                "tipo":                request.form.get("tipo"),
                "marca":               request.form.get("marca"),
                "processador":         request.form.get("processador"),
                "memoria":             request.form.get("memoria"),
                "armazenamento":       request.form.get("armazenamento"),
                "possui_carregador":   request.form.get("possui_carregador"),
                "recebido_por":        request.form.get("recebido_por"),
                "unidade":             request.form.get("unidade"),
                "observacoes":         request.form.get("observacoes"),
                "movido_para_estoque": request.form.get("movido_para_estoque"),
                "email_responsavel":   request.form.get("email_responsavel"),
                "gestor_email":        request.form.get("gestor_email"),
            }

            campos_obrigatorios = ["departamento", "usuario", "tipo", "marca", "modelo", "patrimonio", "nome"]
            # Serial obrigatório quando marca for Dell
            marca_lower = (dados.get("marca") or "").strip().lower()
            if marca_lower == "dell" and not dados.get("serial"):
                campos_obrigatorios.append("serial")

            if not all(dados.get(c) for c in campos_obrigatorios):
                if "serial" in campos_obrigatorios and not dados.get("serial"):
                    flash("Service Tag (Serial) é obrigatório para equipamentos Dell.", "error")
                else:
                    flash("Todos os campos obrigatórios devem ser preenchidos.", "error")
                return redirect(url_for("nova_devolucao"))

            # Upload de foto (apenas para Danificado)
            if "foto" in request.files and dados["status"] == "Danificado":
                arquivo = request.files["foto"]
                if arquivo and arquivo.filename:
                    try:
                        nome_seguro = secure_filename(arquivo.filename)
                        ext = os.path.splitext(nome_seguro)[1].lower()
                        if ext not in ALLOWED_IMAGE_EXTENSIONS:
                            flash("Formato de imagem inválido. Use PNG, JPG, JPEG ou WEBP.", "error")
                            return redirect(url_for("nova_devolucao"))
                        if not (arquivo.mimetype and arquivo.mimetype.startswith("image/")):
                            flash("Arquivo enviado não é uma imagem válida.", "error")
                            return redirect(url_for("nova_devolucao"))
                        nome_arquivo = secure_filename(f"{dados['patrimonio']}{ext}")
                        destino = os.path.join(UPLOAD_DIR, nome_arquivo)
                        arquivo.save(destino)
                        dados["foto"] = destino
                    except Exception as err:
                        logger.warning("Erro ao salvar foto: %s", err)

            registro = inserir(dados)
            novo_id  = registro["id"] if registro else None
            sync_error = registro.get("_sync_error") if registro else None
            logger.info("Devolução registrada: %s", dados["patrimonio"])

            # Envio de email
            if OUTLOOK_DISPONIVEL and novo_id:
                try:
                    cfg = _cfg()
                    email_rh_dest = cfg.get("email_rh") or EMAIL_RH
                    destino_log = "RH"
                    if dados["status"] == "Danificado":
                        email_dano(dados)
                        destino_log = "Dell"
                    enviar_email_rh(dados, para=email_rh_dest)
                    if dados["status"] == "Danificado":
                        destino_log += " + RH"
                    email_sync_error = registrar_email_enviado(novo_id, destino_log)
                    flash("Devolução registrada e rascunhos abertos no Outlook!", "success")
                    if email_sync_error:
                        flash(email_sync_error, "warning")
                except RuntimeError as err:
                    logger.warning("Email não enviado: %s", err)
                    flash("Devolução registrada, mas o Outlook não pôde abrir o rascunho.", "warning")
            elif OUTLOOK_DISPONIVEL:
                flash("Devolução registrada. (Erro ao recuperar ID para log de email.)", "info")
            else:
                flash("Devolução registrada. (Email automático disponível apenas no Windows com Outlook.)", "info")

            if sync_error:
                flash(sync_error, "warning")

            return redirect(url_for("index"))

        except Exception:
            logger.exception("Erro ao processar nova devolução")
            flash("Erro ao processar a devolução.", "error")
            return redirect(url_for("nova_devolucao"))

    return render_template("nova.html")


# ── Editar devolução ─────────────────────────────────────────────
@app.route("/editar/<int:id_devolucao>", methods=["GET", "POST"])
@login_required
def editar_devolucao(id_devolucao):
    registro = buscar_por_id(id_devolucao)
    if not registro:
        flash("Registro não encontrado.", "error")
        return redirect(url_for("index"))

    if request.method == "POST":
        try:
            dados = {
                "usuario":             request.form.get("usuario"),
                "nome":                request.form.get("nome"),
                "matricula":           request.form.get("matricula"),
                "departamento":        request.form.get("departamento"),
                "patrimonio":          request.form.get("patrimonio"),
                "modelo":              request.form.get("modelo"),
                "serial":              request.form.get("serial"),
                "status":              request.form.get("status"),
                "motivo":              request.form.get("motivo"),
                "diretoria":           request.form.get("diretoria"),
                "tipo":                request.form.get("tipo"),
                "marca":               request.form.get("marca"),
                "processador":         request.form.get("processador"),
                "memoria":             request.form.get("memoria"),
                "armazenamento":       request.form.get("armazenamento"),
                "possui_carregador":   request.form.get("possui_carregador"),
                "recebido_por":        request.form.get("recebido_por"),
                "unidade":             request.form.get("unidade"),
                "observacoes":         request.form.get("observacoes"),
                "movido_para_estoque": request.form.get("movido_para_estoque"),
                "email_responsavel":   request.form.get("email_responsavel"),
                "gestor_email":        request.form.get("gestor_email"),
                "chamado_dell":        request.form.get("chamado_dell"),
            }

            campos_obrigatorios = ["departamento", "usuario", "tipo", "marca", "modelo", "patrimonio", "nome"]
            if not all(dados.get(c) for c in campos_obrigatorios):
                flash("Todos os campos obrigatórios devem ser preenchidos.", "error")
                return redirect(url_for("editar_devolucao", id_devolucao=id_devolucao))

            # Upload de foto (substituição)
            if "foto" in request.files:
                arquivo = request.files["foto"]
                if arquivo and arquivo.filename:
                    try:
                        nome_seguro = secure_filename(arquivo.filename)
                        ext = os.path.splitext(nome_seguro)[1].lower()
                        if ext not in ALLOWED_IMAGE_EXTENSIONS:
                            flash("Formato de imagem inválido. Use PNG, JPG, JPEG ou WEBP.", "error")
                            return redirect(url_for("editar_devolucao", id_devolucao=id_devolucao))
                        if not (arquivo.mimetype and arquivo.mimetype.startswith("image/")):
                            flash("Arquivo enviado não é uma imagem válida.", "error")
                            return redirect(url_for("editar_devolucao", id_devolucao=id_devolucao))
                        nome_arquivo = secure_filename(f"{dados['patrimonio']}{ext}")
                        destino = os.path.join(UPLOAD_DIR, nome_arquivo)
                        arquivo.save(destino)
                        dados["foto"] = destino
                        logger.info("Foto atualizada para registro %d: %s", id_devolucao, destino)
                    except Exception as err:
                        logger.warning("Erro ao salvar foto na edição: %s", err)

            sync_error = atualizar(id_devolucao, dados)
            logger.info("Registro %d atualizado: %s", id_devolucao, dados["patrimonio"])
            flash(f"Registro #{id_devolucao} atualizado com sucesso.", "success")
            if sync_error:
                flash(sync_error, "warning")
            return redirect(url_for("index"))

        except Exception:
            logger.exception("Erro ao atualizar registro %d", id_devolucao)
            flash("Erro ao atualizar o registro.", "error")
            return redirect(url_for("editar_devolucao", id_devolucao=id_devolucao))

    # Registro já é dict graças ao Row factory
    return render_template("editar.html", r=registro)


# ── Excluir devolução ────────────────────────────────────────────
@app.route("/excluir/<int:id_devolucao>", methods=["POST"])
@login_required
def excluir_devolucao(id_devolucao):
    try:
        sync_error = excluir(id_devolucao)
        logger.info("Registro %d excluído.", id_devolucao)
        flash(f"Registro #{id_devolucao} excluído.", "success")
        if sync_error:
            flash(sync_error, "warning")
    except Exception:
        logger.exception("Erro ao excluir registro %d", id_devolucao)
        flash("Erro ao excluir o registro.", "error")
    return redirect(url_for("index"))


# ── Atualizar chamado Dell ───────────────────────────────────────
@app.route("/chamado_dell/<int:id_devolucao>", methods=["POST"])
@login_required
def salvar_chamado_dell(id_devolucao):
    chamado = request.form.get("chamado", "").strip()
    try:
        sync_error = atualizar_chamado_dell(id_devolucao, chamado or None)
        logger.info("Chamado Dell atualizado para registro %d: %s", id_devolucao, chamado)
        flash(f"Nº chamado Dell atualizado para #{ id_devolucao }.", "success")
        if sync_error:
            flash(sync_error, "warning")
    except Exception:
        logger.exception("Erro ao atualizar chamado Dell %d", id_devolucao)
        flash("Erro ao salvar o chamado Dell.", "error")
    return redirect(request.referrer or url_for("index"))


# ── Configurações ────────────────────────────────────────────────
@app.route("/configuracoes", methods=["GET", "POST"])
@login_required
def configuracoes():
    cfg = _cfg()
    planilha_diag = diagnosticar_config_planilha(cfg.get("planilha_empresa", ""))
    if request.method == "POST":
        # Strip de aspas extras que o usuário possa ter colado ao redor do caminho
        planilha_raw = request.form.get("planilha_empresa", "").strip().strip('"').strip("'")
        cfg["planilha_empresa"] = planilha_raw
        cfg["email_rh"]         = request.form.get("email_rh", "").strip()
        cfg.pop("aba_planilha_empresa", None)
        _salvar_cfg(cfg)
        planilha_diag = diagnosticar_config_planilha(planilha_raw)
        flash("✅ Configurações salvas com sucesso.", "success")
        if planilha_diag["bloqueada"]:
            flash(
                f"⚠ Integração Excel bloqueada neste ambiente. {planilha_diag['detalhe']}",
                "warning",
            )
        return redirect(url_for("configuracoes"))
    return render_template("configuracoes.html", cfg=cfg, planilha_diag=planilha_diag)


@app.route("/sincronizar_planilha", methods=["POST"])
@login_required
def sincronizar_planilha():
    """Força sincronização completa de todos os registros na planilha corporativa."""
    diagnostico = diagnosticar_config_planilha()
    if diagnostico["bloqueada"]:
        flash(
            f"❌ Integração Excel bloqueada neste ambiente. {diagnostico['detalhe']}",
            "error",
        )
        return redirect(url_for("configuracoes"))
    try:
        sincronizar_planilha_completa()
        flash("✅ Planilha sincronizada com sucesso — todos os registros foram gravados.", "success")
    except Exception as err:
        logger.error("Erro ao sincronizar planilha: %s", err)
        flash(f"❌ Erro ao sincronizar planilha: {err}", "error")
    return redirect(url_for("configuracoes"))


# ── Fotos ────────────────────────────────────────────────────────
@app.route("/foto/<filename>")
@login_required
def servir_foto(filename):
    nome_arquivo = secure_filename(filename)
    caminho = os.path.join(UPLOAD_DIR, nome_arquivo)
    if os.path.exists(caminho):
        return send_from_directory(UPLOAD_DIR, nome_arquivo)
    flash("Foto não encontrada.", "error")
    return redirect(url_for("index"))


@app.route("/visualizar/<filename>")
@login_required
def visualizar_foto(filename):
    nome_arquivo = secure_filename(filename)
    caminho = os.path.join(UPLOAD_DIR, nome_arquivo)
    if os.path.exists(caminho):
        return render_template(
            "visualizar_foto.html",
            foto_url=url_for("servir_foto", filename=nome_arquivo),
            nome_arquivo=nome_arquivo,
        )
    flash("Foto não encontrada.", "error")
    return redirect(url_for("index"))


if __name__ == "__main__":
    host  = os.getenv("FLASK_HOST", "127.0.0.1")
    port  = int(os.getenv("FLASK_PORT", "5000"))
    debug = os.getenv("FLASK_DEBUG", "true").strip().lower() in {"1", "true", "yes", "on"}
    logger.info("Iniciando servidor Flask em http://%s:%s", host, port)
    app.run(debug=debug, host=host, port=port)
