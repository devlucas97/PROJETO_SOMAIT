# 📱 Sistema de Devoluções de Notebook

Sistema completo para gerenciamento de devoluções e incidentes com notebooks. Suporta interface desktop (Qt) no Windows e versão web (Flask) em qualquer plataforma.

## ✨ Funcionalidades

- ✅ Registro de devoluções com status (OK, Danificado, Pendente)
- 📧 Notificações automáticas por email (Windows com Outlook)
- 📸 Upload de fotos de danos
- 📊 Histórico completo de devoluções
- 💾 Banco de dados SQLite integrado
- 🎨 Interface amigável (Desktop Qt ou Web Flask)
- 📝 Logging completo de operações

## 📋 Requisitos

### Windows (Desktop)
- Python 3.8+
- Outlook instalado e configurado

### Linux/Mac (Web)
- Python 3.8+
- Qualquer navegador web moderno

## 🚀 Instalação

### 1. Clonar/Baixar o projeto

```bash
cd projeto_dio
```

### 2. Criar ambiente virtual

```bash
python -m venv .venv
```

### 3. Ativar ambiente virtual

**Windows:**
```bash
.venv\Scripts\activate
```

**Linux/Mac:**
```bash
source .venv/bin/activate
```

### 4. Instalar dependências

```bash
pip install -r requirements.txt
```

> **Nota sobre Windows:** Se encontrar erros com `pywin32`, execute:
> ```bash
> python -m pip install --upgrade pywin32
> python -m pip install --upgrade pyinstaller
> ```

---

## 💻 Execução

### Opção 1: Interface Desktop (Windows)

Executa a aplicação gráfica Qt com suporte a Outlook:

```bash
python main.py
```

**Funcionalidades:**
- Interface gráfica confortável
- Upload direto de fotos com preview
- Envio de emails via Outlook integrado
- Histórico visual em tabela

### Opção 2: Versão Web (Qualquer plataforma)

Executa servidor Flask e abre no navegador:

```bash
python web.py
```

Em seguida, acesse: **http://127.0.0.1:5000**

**Funcionalidades:**
- Acesso via navegador web
- Mesmo banco de dados (SQLite)
- Formulário responsivo
- Compatível com Linux, Mac e Windows
- Email funciona em Windows (com Outlook local)

---

## 📁 Estrutura do Projeto

```
projeto_dio/
├── main.py               # Entrada para versão Desktop (Qt)
├── web.py               # Servidor Flask (Web)
├── database.py          # Lógica de banco de dados
├── email_service.py     # Funções de email (Outlook)
├── ui_main.py           # Interface gráfica (Qt)
├── seed_database.py     # Script para popular DB com dados de teste
├── requirements.txt     # Dependências do projeto
├── README.md           # Este arquivo
├── database.db         # Banco SQLite (criado automaticamente)
├── uploads/            # Fotos de danos (criada automaticamente)
└── templates/
    ├── index.html      # Página inicial (Web)
    └── nova.html       # Formulário de nova devolução (Web)
```

---

## 🧪 Testando com Dados Fictícios

Para preencher o banco com dados de teste:

```bash
python seed_database.py
```

Isso criará 4 registros de exemplo no banco de dados para você testar toda a funcionalidade.

---

## 📧 Configuração de Email

### Windows (Desktop)
Outlook deve estar instalado e configurado com conta de email. O `email_service.py` usa COM (Component Object Model) para integração nativa:

```python
import win32com.client as win32
outlook = win32.Dispatch("outlook.application")
```

### Linux/Mac (Web)
Email **não funciona** em modo web no Linux/Mac (restrição de plataforma). A aplicação registra a devolução normalmente, mas email é opcional e não bloqueia.

---

## 🐛 Troubleshooting

### "libGL.so.1: cannot open shared object file" (Linux/Mac)
**Solução:** Use a versão Web (`python web.py`) em vez da Desktop.

### "Outlook não iniciado" (Windows)
**Solução:** Verifique se Outlook está instalado e configurado. Abra Outlook manualmente primeira vez.

### ImportError: No module named 'win32com' (Linux/Mac)
**Solução:** Normal - `pywin32` é Windows-only. Use a versão Web.

### Falha ao salvar foto
**Solução:** Verifique permissões da pasta `uploads/`. Crie manualmente se necessário.

---

## 📝 Campos do Formulário

| Campo | Descrição | Obrigatório |
|-------|-----------|------------|
| Usuário | Nome de usuário (login) | ✅ Sim |
| Nome Completo | Nome da pessoa | ✅ Sim |
| Matrícula | ID do funcionário | ✅ Sim |
| Departamento | Setor da empresa | ✅ Sim |
| Patrimônio | ID/Número do notebook | ✅ Sim |
| Modelo | Dell, HP, Lenovo, etc | ✅ Sim |
| Serial | Número de série do device | ✅ Sim |
| Status | OK / Danificado / Pendente | ✅ Sim |
| Motivo | Descrição do problema | ❌ Não |
| Foto | Comprovação de dano | ❌ Não |

---

## 📊 Banco de Dados

Tabela `devolucoes`:
```sql
sqlite> .schema devolucoes
CREATE TABLE devolucoes (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    data TEXT,
    usuario TEXT,
    nome TEXT,
    matricula TEXT,
    departamento TEXT,
    patrimonio TEXT,
    modelo TEXT,
    serial TEXT,
    status TEXT,
    motivo TEXT
);
```

---

## 🔐 Segurança

- ✅ Validação de campos obrigatórios
- ✅ Tratamento de exceções em operações críticas
- ✅ Logging completo de ações
- ✅ Isolamento de erros de email (não bloqueia DB)
- ❌ **Não produção:** Usar senha forte em `app.secret_key` (web.py)

---

## 📈 Próximas Melhorias (Opcional)

- [ ] Autenticação de usuário (login)
- [ ] Editar/Deletar registros
- [ ] Exportar dados (CSV, PDF)
- [ ] Integração com LDAP para usuários corporativos
- [ ] Dashboard com estatísticas
- [ ] Notificações por SMS
- [ ] Deploy em servidor (Docker, AWS, Heroku)

---

## 👨‍💻 Desenvolvido por

Projeto de Sistema de Devoluções

- Versão Desktop: PySide6 (Qt)
- Versão Web: Flask
- Banco: SQLite
- Email: win32com (Outlook/Windows)

---

## 📄 Licença

MIT License - Livre para usar e modificar.

---

## 📞 Suporte

Para erros ou dúvidas:
1. Verifique os logs no console
2. Consulte `troubleshooting` acima
3. Abra uma issue no GitHub
