from app.database import criar, inserir, listar


def main():
    criar()

    amostras = [
        {
            "usuario": "luiz.silva",
            "nome": "Luiz Silva",
            "matricula": "12345",
            "departamento": "TI",
            "patrimonio": "PT-1001",
            "modelo": "Dell Latitude 5410",
            "serial": "SN123456789",
            "status": "OK",
            "tipo": "Notebook",
            "marca": "Dell",
            "motivo": "Desligamento",
            "foto": None,
        },
        {
            "usuario": "mariana.rodrigues",
            "nome": "Mariana Rodrigues",
            "matricula": "12346",
            "departamento": "RH",
            "patrimonio": "PT-1002",
            "modelo": "HP ProBook 450",
            "serial": "SN987654321",
            "status": "Pendente",
            "tipo": "Notebook",
            "marca": "HP",
            "motivo": "Troca de equipamento",
            "foto": None,
        },
        {
            "usuario": "pedro.santos",
            "nome": "Pedro Santos",
            "matricula": "12347",
            "departamento": "Financeiro",
            "patrimonio": "PT-1003",
            "modelo": "Lenovo ThinkPad T14",
            "serial": "SN456789123",
            "status": "Danificado",
            "tipo": "Notebook",
            "marca": "Lenovo",
            "motivo": "Tela quebrada",
            "foto": None,
        },
        {
            "usuario": "sara.lima",
            "nome": "Sara Lima",
            "matricula": "12348",
            "departamento": "Marketing",
            "patrimonio": "PT-1004",
            "modelo": "Acer Aspire 5",
            "serial": "SN321654987",
            "status": "OK",
            "tipo": "Notebook",
            "marca": "Acer",
            "motivo": "Desligamento",
            "foto": None,
        },
    ]

    for item in amostras:
        inserir(item)

    registros = listar()
    print(f"{len(registros)} registros inseridos no banco de dados")
    for r in registros:
        print(r)


if __name__ == "__main__":
    main()
