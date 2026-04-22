import json
import os
import pytest
import app.database as database


@pytest.fixture
def client(monkeypatch, tmp_path):
    db_file = tmp_path / "test_database.db"
    monkeypatch.setattr(database, "DB", str(db_file))
    database.criar()

    import app.web as web

    with web.app.test_client() as c:
        # Simula sessão autenticada com token CSRF
        with c.session_transaction() as sess:
            sess["autenticado"] = True
            sess["usuario_logado"] = "test"
            sess["_csrf_token"] = "test_csrf_token"
        yield c


def _with_csrf(data: dict) -> dict:
    """Adiciona token CSRF aos dados do formulário de teste."""
    return {**data, "_csrf_token": "test_csrf_token"}


def test_index_sem_dados(client):
    response = client.get("/")
    assert response.status_code == 200
    assert b"Nenhum registro" not in response.data


def test_busca_por_id(client):
    dados = {
        "usuario": "test.user",
        "nome": "Test User",
        "matricula": "12345",
        "departamento": "TI",
        "patrimonio": "PT-0003",
        "modelo": "ModeloY",
        "serial": "SN0003",
        "status": "Devolvido",
        "motivo": "Teste.",
        "foto": None,
    }
    database.inserir(dados)

    response = client.get("/", query_string={"search_id": "1"})
    assert response.status_code == 200
    assert b"PT-0003" in response.data


def test_nova_devolucao_persiste_registro(client, monkeypatch):
    import app.web as web

    monkeypatch.setattr(web, "OUTLOOK_DISPONIVEL", False)

    response = client.post(
        "/nova",
        data=_with_csrf({
            "usuario": "joao.silva",
            "nome": "Joao Silva",
            "matricula": "12345",
            "departamento": "TI",
            "patrimonio": "PT-0100",
            "modelo": "Latitude 5520",
            "serial": "ABC1234",
            "status": "OK",
            "motivo": "Desligamento",
            "diretoria": "Operacoes",
            "tipo": "Notebook",
            "marca": "Dell",
            "processador": "Intel Core i5",
            "memoria": "16 GB",
            "armazenamento": "SSD 256 GB",
            "possui_carregador": "Sim",
            "recebido_por": "Analista TI",
            "unidade": "Sede SP",
            "observacoes": "Registro de teste",
            "movido_para_estoque": "Sim",
            "email_responsavel": "responsavel@empresa.com",
            "gestor_email": "gestor@empresa.com",
        }),
        follow_redirects=False,
    )

    assert response.status_code == 302
    assert response.headers["Location"].endswith("/")

    registros = database.listar()
    assert len(registros) == 1
    assert registros[0]["patrimonio"] == "PT-0100"
    assert registros[0]["status"] == "OK"


def test_nova_devolucao_ok_abre_apenas_email_rh(client, monkeypatch):
    import app.web as web

    chamadas = []
    monkeypatch.setattr(web, "OUTLOOK_DISPONIVEL", True)
    monkeypatch.setattr(web, "enviar_email", lambda dados: chamadas.append("ti"))
    monkeypatch.setattr(web, "email_dano", lambda dados: chamadas.append("dano"))
    monkeypatch.setattr(web, "email_cotacao_dell", lambda dados: chamadas.append("dell"))
    monkeypatch.setattr(web, "enviar_email_rh", lambda dados, para=None: chamadas.append("rh"))
    monkeypatch.setattr(web, "registrar_email_enviado", lambda id_devolucao, destino: None)

    response = client.post(
        "/nova",
        data=_with_csrf({
            "usuario": "joao.silva",
            "nome": "Joao Silva",
            "matricula": "12345",
            "departamento": "TI",
            "patrimonio": "PT-0102",
            "modelo": "Latitude 5520",
            "serial": "ABC1234",
            "status": "OK",
            "motivo": "Desligamento",
            "diretoria": "Operacoes",
            "tipo": "Notebook",
            "marca": "Dell",
        }),
        follow_redirects=False,
    )

    assert response.status_code == 302
    assert chamadas == ["rh"]


def test_nova_devolucao_danificado_abre_email_dell_e_rh(client, monkeypatch):
    import app.web as web

    chamadas = []
    monkeypatch.setattr(web, "OUTLOOK_DISPONIVEL", True)
    monkeypatch.setattr(web, "email_dano", lambda dados: chamadas.append("dell"))
    monkeypatch.setattr(web, "email_cotacao_dell", lambda dados: chamadas.append("cotacao"))
    monkeypatch.setattr(web, "enviar_email_rh", lambda dados, para=None: chamadas.append("rh"))
    monkeypatch.setattr(web, "registrar_email_enviado", lambda id_devolucao, destino: None)

    response = client.post(
        "/nova",
        data=_with_csrf({
            "usuario": "joao.silva",
            "nome": "Joao Silva",
            "matricula": "12345",
            "departamento": "TI",
            "patrimonio": "PT-0103",
            "modelo": "Latitude 5520",
            "serial": "ABC1234",
            "status": "Danificado",
            "motivo": "Queda",
            "diretoria": "Operacoes",
            "tipo": "Notebook",
            "marca": "Dell",
        }),
        follow_redirects=False,
    )

    assert response.status_code == 302
    assert chamadas == ["dell", "rh"]


def test_nova_devolucao_exibe_aviso_quando_excel_esta_bloqueado(client, monkeypatch):
    import app.web as web

    monkeypatch.setattr(web, "OUTLOOK_DISPONIVEL", False)
    monkeypatch.setattr(web, "inserir", lambda dados: {
        "id": 1,
        "patrimonio": dados["patrimonio"],
        "_sync_error": "A devolução foi salva, mas a planilha Excel está bloqueada para escrita.",
    })

    response = client.post(
        "/nova",
        data=_with_csrf({
            "usuario": "joao.silva",
            "nome": "Joao Silva",
            "matricula": "12345",
            "departamento": "TI",
            "patrimonio": "PT-0101",
            "modelo": "Latitude 5520",
            "serial": "ABC1234",
            "status": "OK",
            "motivo": "Desligamento",
            "diretoria": "Operacoes",
            "tipo": "Notebook",
            "marca": "Dell",
        }),
        follow_redirects=True,
    )

    assert response.status_code == 200
    assert b"planilha Excel" in response.data


def test_configuracoes_post_salva_arquivo_de_config(client, monkeypatch, tmp_path):
    import app.web as web

    config_path = tmp_path / "config.json"
    monkeypatch.setattr(web, "iter_config_paths", lambda: iter([str(config_path)]))

    response = client.post(
        "/configuracoes",
        data=_with_csrf({
            "planilha_empresa": "/tmp/empresa.xlsx",
            "email_rh": "rh@empresa.com",
        }),
        follow_redirects=False,
    )

    assert response.status_code == 302
    assert config_path.exists()

    saved = json.loads(config_path.read_text(encoding="utf-8"))
    assert saved["planilha_empresa"] == "/tmp/empresa.xlsx"
    assert saved["email_rh"] == "rh@empresa.com"


def test_configuracoes_exibe_alerta_para_caminho_incompativel(client, monkeypatch):
    import app.web as web

    caminho_windows = r"C:\Compartilhado\TI\Devolucoes.xlsx"
    monkeypatch.setattr(web, "_cfg", lambda: {
        "planilha_empresa": caminho_windows,
        "email_rh": "rh@somagrupo.com.br",
    })
    monkeypatch.setattr(web, "diagnosticar_config_planilha", lambda caminho=None: {
        "caminho": caminho_windows,
        "ambiente": "Linux / container",
        "ok": False,
        "bloqueada": True,
        "nivel": "error",
        "mensagem": "Caminho Windows detectado em ambiente Linux / container.",
        "detalhe": "Use um caminho acessível dentro do container.",
    })

    response = client.get("/configuracoes")

    assert response.status_code == 200
    assert b"Caminho Windows detectado" in response.data
    assert b"Linux / container" in response.data


def test_editar_registro_via_web(client, monkeypatch):
    import app.web as web

    monkeypatch.setattr(web, "OUTLOOK_DISPONIVEL", False)

    # Cria um registro
    client.post(
        "/nova",
        data=_with_csrf({
            "usuario": "joao.silva",
            "nome": "Joao Silva",
            "matricula": "12345",
            "departamento": "TI",
            "patrimonio": "PT-EDIT",
            "modelo": "Latitude 5520",
            "serial": "ABC1234",
            "status": "OK",
            "motivo": "Teste",
            "diretoria": "Operacoes",
            "tipo": "Notebook",
            "marca": "Dell",
        }),
        follow_redirects=False,
    )

    registros = database.listar()
    reg_id = registros[0]["id"]

    # Edita
    response = client.post(
        f"/editar/{reg_id}",
        data=_with_csrf({
            "usuario": "joao.silva",
            "nome": "Joao Silva Updated",
            "matricula": "12345",
            "departamento": "RH",
            "patrimonio": "PT-EDIT",
            "modelo": "Latitude 5520",
            "serial": "ABC1234",
            "status": "Danificado",
            "motivo": "Queda",
            "diretoria": "Operacoes",
            "tipo": "Notebook",
            "marca": "Dell",
        }),
        follow_redirects=False,
    )

    assert response.status_code == 302
    atualizado = database.buscar_por_id(reg_id)
    assert atualizado["nome"] == "Joao Silva Updated"
    assert atualizado["status"] == "Danificado"


def test_excluir_registro_via_web(client, monkeypatch):
    import app.web as web

    monkeypatch.setattr(web, "OUTLOOK_DISPONIVEL", False)

    client.post(
        "/nova",
        data=_with_csrf({
            "usuario": "joao.silva",
            "nome": "Joao Silva",
            "matricula": "12345",
            "departamento": "TI",
            "patrimonio": "PT-DEL",
            "modelo": "Latitude 5520",
            "serial": "ABC1234",
            "status": "OK",
            "motivo": "Teste",
            "diretoria": "Operacoes",
            "tipo": "Notebook",
            "marca": "Dell",
        }),
        follow_redirects=False,
    )

    registros = database.listar()
    reg_id = registros[0]["id"]

    response = client.post(
        f"/excluir/{reg_id}",
        data=_with_csrf({}),
        follow_redirects=False,
    )

    assert response.status_code == 302
    assert database.buscar_por_id(reg_id) is None


def test_chamado_dell_via_web(client, monkeypatch):
    import app.web as web

    monkeypatch.setattr(web, "OUTLOOK_DISPONIVEL", False)

    client.post(
        "/nova",
        data=_with_csrf({
            "usuario": "joao.silva",
            "nome": "Joao Silva",
            "matricula": "12345",
            "departamento": "TI",
            "patrimonio": "PT-DELL",
            "modelo": "Latitude 5520",
            "serial": "ABC1234",
            "status": "Danificado",
            "motivo": "Dano",
            "diretoria": "Operacoes",
            "tipo": "Notebook",
            "marca": "Dell",
        }),
        follow_redirects=False,
    )

    registros = database.listar()
    reg_id = registros[0]["id"]

    response = client.post(
        f"/chamado_dell/{reg_id}",
        data=_with_csrf({"chamado": "OS-99999"}),
        follow_redirects=False,
    )

    assert response.status_code == 302
    atualizado = database.buscar_por_id(reg_id)
    assert atualizado["chamado_dell"] == "OS-99999"

