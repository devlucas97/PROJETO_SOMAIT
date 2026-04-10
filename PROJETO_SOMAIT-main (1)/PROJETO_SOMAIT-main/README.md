# SOMALABS — Gestão de Devoluções e Ativos

Sistema corporativo para gerenciamento de devoluções de equipamentos de TI, com dashboard executivo, fluxo Dell, evidências fotográficas, integração com planilha Excel e rascunhos de e-mail via Outlook.

---

## Funcionalidades

| Recurso | Descrição |
|---------|-----------|
| **Dashboard executivo** | KPIs, filtros por status/data/texto, paginação |
| **Registro de devoluções** | Formulário completo com campos corporativos |
| **Fluxo Dell** | Workflow dedicado (cotação → reparo → conclusão) |
| **Evidências fotográficas** | Upload, visualização e validação de fotos de dano |
| **Integração Excel** | Sincronização automática com planilha corporativa (3 abas) |
| **E-mails via Outlook** | Rascunhos automáticos para RH, Dell e responsáveis (Windows) |
| **Banco SQLite** | Armazenamento local, sem dependência de servidor |
| **Interface Qt (legado)** | Desktop com fallback automático para web |

---

## Estrutura do Projeto

```
PROJETO_SOMAIT/
│
├── main.py                         # Ponto de entrada (Qt + fallback web)
├── config.json                     # Configurações de integração (runtime)
├── requirements.txt                # Dependências Python
├── .env.example                    # Variáveis de ambiente (template)
├── .gitignore                      # Arquivos ignorados pelo Git
├── README.md                       # Esta documentação
│
├── app/                            # Pacote principal da aplicação
│   ├── __init__.py
│   ├── web.py                      # Aplicação Flask (rotas, views, API)
│   ├── database.py                 # Camada de dados (SQLite + Excel)
│   ├── email_service.py            # Integração Outlook COM (Windows)
│   ├── ui_main.py                  # Interface Qt (legado)
│   ├── logging_config.py           # Configuração de logging
│   ├── static/
│   │   └── css/
│   │       └── app.css             # Estilos da interface web
│   └── templates/
│       ├── base.html               # Template base (header, footer)
│       ├── login.html              # Tela de login
│       ├── index.html              # Dashboard principal
│       ├── nova.html               # Formulário de nova devolução
│       ├── editar.html             # Formulário de edição
│       ├── configuracoes.html      # Painel de configurações
│       └── visualizar_foto.html    # Visualizador de foto
│
├── scripts/                        # Utilitários e scripts auxiliares
│   ├── seed_database.py            # Popular banco com dados de teste
│   ├── iniciar_web_windows.bat     # Inicializador Windows (duplo clique)
│   └── main.spec                   # Configuração PyInstaller (.exe)
│
├── tests/                          # Testes automatizados (pytest)
│   ├── test_database.py            # Testes da camada de dados
│   └── test_web.py                 # Testes da aplicação web
│
└── uploads/                        # Fotos de dano (gerado em runtime)
```

---

## Requisitos

| Plataforma | Requisitos |
|------------|------------|
| **Windows** | Python 3.10+, Outlook clássico (para e-mails), navegador moderno |
| **Linux / Mac / Container** | Python 3.10+, navegador moderno |

> **Nota:** A integração com Outlook (COM) funciona exclusivamente no Windows.

---

## Instalação

```bash
# 1. Clonar o repositório
git clone <url-do-repositorio>
cd PROJETO_SOMAIT

# 2. Criar ambiente virtual
python -m venv .venv

# 3. Ativar ambiente virtual
# Windows:
.venv\Scripts\activate
# Linux/Mac:
source .venv/bin/activate

# 4. Instalar dependências
pip install -r requirements.txt
```

> Se houver problemas com `pywin32` no Windows:
> ```bash
> python -m pip install --upgrade pywin32
> ```

---

## Execução

### Web (recomendado)

```bash
python -m app.web
```

Acesse: **http://127.0.0.1:5000**

Login padrão: `admin` / `azzas2026`

No Windows, também é possível usar duplo clique em `scripts/iniciar_web_windows.bat`.

### Desktop Qt (legado)

```bash
python main.py
```

Detecta automaticamente se há ambiente gráfico. Caso contrário, inicia o servidor web como fallback.

No Windows, também é possível usar duplo clique em `scripts/iniciar_desktop_windows.bat`.
Para iniciar sem console, use `scripts/iniciar_desktop_windows.bat --quiet`.

Para iniciar pelo terminal com logs visíveis, use `scripts\iniciar_desktop_windows.bat --console`.

### Ícone do desktop no Windows

É possível definir um ícone personalizado para o modo desktop.

1. Salve o arquivo de ícone como `assets/app.ico` na raiz do projeto.
2. Se você abrir pelo executável gerado com PyInstaller, gere novamente o `.exe` usando `scripts/main.spec`.
3. Se você abrir por um atalho da área de trabalho apontando para o `.bat` ou para o `.exe`, recrie o atalho ou ajuste o campo **Alterar Ícone** nas propriedades do atalho.

Sem o arquivo `assets/app.ico`, o projeto continua funcionando e usa o ícone padrão do Windows.

## Build do desktop no Windows

Para gerar um executável portátil do modo desktop:

```bat
scripts\build_desktop_portable_windows.bat
```

Resultado esperado:

- `dist\SOMALABSDesktop.exe`: executável gerado pelo PyInstaller
- `dist\SOMALABSDesktop-portable\`: pasta pronta para copiar para outro computador Windows

A pasta portátil pode incluir automaticamente:

- `config.json`
- `database.db`
- `uploads\`
- `assets\`

No computador de destino, basta abrir `SOMALABSDesktop.exe`.

Pré-requisitos no computador de destino:

- Windows
- Outlook clássico instalado e configurado, se for usar geração de emails

Observação: o executável é portátil. Para criar um instalador `.exe` com assistente de instalação, o caminho mais simples é usar Inno Setup ou NSIS sobre a pasta `dist\SOMALABSDesktop-portable\`.

---

## Dados de teste

```bash
python scripts/seed_database.py
```

Insere registros de exemplo para validar dashboard, filtros e sincronização.

---

## Configuração

### Variáveis de ambiente

Copie `.env.example` para `.env` e ajuste os valores:

```bash
cp .env.example .env
```

Troque obrigatoriamente `APP_PASSWORD` e `FLASK_SECRET_KEY` antes de compartilhar o uso da aplicação.

### Planilha Excel

Em **Configurações** na interface web, informe o caminho da planilha corporativa. No Windows:

```
C:\Users\seu.usuario\Downloads\devolucoes.xlsx
```

> Caminhos Windows são bloqueados em ambiente Linux/container por segurança.

### E-mail (Windows)

O Outlook clássico deve estar instalado e configurado. Abra-o ao menos uma vez antes de usar a integração.

---

## Testes

```bash
python -m pytest tests/ -q
```

---

## Troubleshooting

| Problema | Solução |
|----------|---------|
| E-mails não funcionam | Rode no Windows host com Outlook configurado |
| Planilha não sincroniza no Linux | Use caminho acessível no container, não `C:\...` |
| `libGL.so.1: cannot open shared object file` | Use `python -m app.web` em vez do modo Qt |
| `No module named 'win32com'` (Linux/Mac) | Esperado — use a versão web sem Outlook |
| Falha ao salvar foto | Verifique permissões da pasta `uploads/` |

---

## Segurança

- Login obrigatório com hash de senha (werkzeug)
- Aviso em log quando credenciais ou secret key padrão estiverem em uso
- Validação de campos obrigatórios no servidor
- Sanitização de nomes de arquivo para uploads
- Validação de MIME-type para imagens
- Bloqueio de caminhos incompatíveis com o ambiente
- Tratamento de exceções em operações críticas
