import os
import sys
from PySide6.QtGui import QPixmap
from PySide6.QtWidgets import QFileDialog
import getpass
import logging
from PySide6.QtWidgets import *
from PySide6.QtCore import Qt
 
from Projeto.database import inserir, listar
from Projeto.email_service import enviar_email, email_dano

logger = logging.getLogger(__name__)
if not logger.hasHandlers():
    handler = logging.StreamHandler()
    handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(name)s: %(message)s"))
    logger.addHandler(handler)
logger.setLevel(logging.INFO)
 
 
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
 
        # Criar pasta de uploads
        global UPLOAD_DIR
        UPLOAD_DIR = "uploads"
        os.makedirs(UPLOAD_DIR, exist_ok=True)
        self.caminho_foto = None

        main_widget = QWidget()
        self.setCentralWidget(main_widget)
 
        layout = QHBoxLayout(main_widget)
 
        # Sidebar
        sidebar = QVBoxLayout()
 
        btn_novo = QPushButton("Nova Devolução")
        btn_hist = QPushButton("Histórico")
 
        btn_novo.clicked.connect(self.show_nova)
        btn_hist.clicked.connect(self.show_historico)
 
        sidebar.addWidget(btn_novo)
        sidebar.addWidget(btn_hist)
        sidebar.addStretch()
 
        # Container de páginas
        self.stack = QStackedWidget()
 
        self.page_nova = self.create_nova_page()
        self.page_hist = self.create_hist_page()
 
        self.stack.addWidget(self.page_nova)
        self.stack.addWidget(self.page_hist)
 
        layout.addLayout(sidebar, 1)
        layout.addWidget(self.stack, 4)
 
    # --------------------------
    # Tela Nova Devolução
    # --------------------------
    def create_nova_page(self):
        widget = QWidget()
        layout = QFormLayout(widget)
 
        self.inputs = {}
 
        campos = ["nome", "matricula", "departamento", "patrimonio", "modelo", "serial"]
 
        for campo in campos:
            line = QLineEdit()
            line.setPlaceholderText(f"Informe {campo}")
            layout.addRow(campo.capitalize(), line)
            self.inputs[campo] = line
 
        self.status_ok = QRadioButton("OK")
        self.status_dano = QRadioButton("Danificado")
        self.status_ok.setChecked(True)
 
        status_layout = QHBoxLayout()
        status_layout.addWidget(self.status_ok)
        status_layout.addWidget(self.status_dano)
 
        layout.addRow("Status", status_layout)
 
        # Botão de upload
        self.btn_upload = QPushButton("Selecionar Foto de Dano")
        self.btn_upload.clicked.connect(self.selecionar_foto)
        layout.addRow("Foto de Dano", self.btn_upload)

        # Preview da imagem
        self.label_foto = QLabel()
        self.label_foto.setFixedSize(200, 150)
        self.label_foto.setStyleSheet("border: 1px solid gray;")
        layout.addRow("", self.label_foto)

    def create_hist_page(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)
 
        self.table = QTableWidget()
        layout.addWidget(self.table)
 
        self.load_table()
 
        return widget
 
    def load_table(self):
        dados = listar()
 
        self.table.setRowCount(len(dados))
        self.table.setColumnCount(6)
        self.table.clearContents()
 
        self.table.setHorizontalHeaderLabels([
            "Data", "Nome", "Patrimônio", "Modelo", "Status", "Usuário"
        ])
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
 
        for row, linha in enumerate(dados):
            self.table.setItem(row, 0, QTableWidgetItem(linha[1]))
            self.table.setItem(row, 1, QTableWidgetItem(linha[3]))
            self.table.setItem(row, 2, QTableWidgetItem(linha[6]))
            self.table.setItem(row, 3, QTableWidgetItem(linha[7]))
            self.table.setItem(row, 4, QTableWidgetItem(linha[9]))
            self.table.setItem(row, 5, QTableWidgetItem(linha[2]))
 
    # --------------------------
    # Ações
    # --------------------------
    def processar(self):
        usuario = getpass.getuser()
 
        dados = {k: v.text().strip() for k, v in self.inputs.items()}
        faltantes = [k for k, v in dados.items() if not v]
        if faltantes:
            QMessageBox.warning(
                self,
                "Validação",
                f"Campos obrigatórios não preenchidos: {', '.join(faltantes)}"
            )
            return
 
        dados["usuario"] = usuario
        dados["status"] = "Danificado" if self.status_dano.isChecked() else "OK"
 
        # Salvar foto localmente se houver
        if self.caminho_foto and dados["status"] == "Danificado":
            ext = os.path.splitext(self.caminho_foto)[1]
            nome_arquivo = f"{dados['patrimonio']}{ext}"
            destino = os.path.join(UPLOAD_DIR, nome_arquivo)
            pixmap = QPixmap(self.caminho_foto)
            pixmap.save(destino)
            dados["foto"] = destino
            logger.info(f"Foto salva em: {destino}")
        else:
            dados["foto"] = None

        try:
            # Sempre tenta enviar devolução, e em caso de dano abrir email de chamado
            enviar_email(dados)
        except Exception as err:
            logger.warning("Falha no envio de e-mail de devolução: %s", err)
            QMessageBox.warning(
                self,
                "Email não enviado",
                f"O e-mail de devolução não pôde ser aberto: {err}"
            )

        if dados["status"] == "Danificado":
            resp = QMessageBox.question(
                self,
                "Notebook Danificado",
                "Deseja abrir chamado?",
                QMessageBox.Yes | QMessageBox.No
            )
            if resp == QMessageBox.Yes:
                try:
                    email_dano(dados)
                except Exception as err:
                    logger.warning("Falha no envio de e-mail de dano: %s", err)
                    QMessageBox.warning(
                        self,
                        "Email de dano não enviado",
                        f"O e-mail de dano não pôde ser aberto: {err}"
                    )

        QMessageBox.information(self, "Sucesso", "Processo concluído!")
 
        self.load_table()
        
        for campo in self.inputs.values():
            campo.clear()
        self.status_ok.setChecked(True)

 
    def show_nova(self):
        self.stack.setCurrentWidget(self.page_nova)
 
    def show_historico(self):
        self.stack.setCurrentWidget(self.page_hist)

    def selecionar_foto(self):
        file_dialog = QFileDialog(self)
        file_dialog.setNameFilter("Imagens (*.png *.jpg *.jpeg)")
        if file_dialog.exec():
            file_path = file_dialog.selectedFiles()[0]
            self.caminho_foto = file_path

            pixmap = QPixmap(file_path).scaled(
                self.label_foto.width(),
                self.label_foto.height(),
                Qt.KeepAspectRatio
            )
            self.label_foto.setPixmap(pixmap)