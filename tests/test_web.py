import os
import pytest
import Projeto.database as database


@pytest.fixture
def client(monkeypatch, tmp_path):
    db_file = tmp_path / "test_database.db"
    excel_file = tmp_path / "test_devolucoes.xlsx"
    monkeypatch.setattr(database, "DB", str(db_file))
    monkeypatch.setattr(database, "EXCEL_FILE", str(excel_file))
    database.criar()

    import Projeto.web as web

    with web.app.test_client() as c:
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


def test_import_export_excel(client, tmp_path):
    # Inserir registro via app / database
    dados = {
        "usuario": "test.user",
        "nome": "Test User",
        "matricula": "12345",
        "departamento": "TI",
        "patrimonio": "PT-0004",
        "modelo": "ModeloZ",
        "serial": "SN0004",
        "status": "Aguardando",
        "motivo": "Teste.",
        "foto": None,
    }
    database.inserir(dados)

    # Exporta Excel via rota
    response = client.get("/export-excel")
    assert response.status_code == 200
    assert response.headers["Content-Type"] == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
