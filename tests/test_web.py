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
        # Simula sessão autenticada
        with c.session_transaction() as sess:
            sess["autenticado"] = True
            sess["usuario_logado"] = "test"
        yield c


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

