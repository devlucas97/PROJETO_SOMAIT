import json
import os
import sys
from PySide6.QtGui import QPixmap
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QHBoxLayout, QVBoxLayout,
    QFormLayout, QStackedWidget, QPushButton, QLineEdit,
    QLabel, QTableWidget, QTableWidgetItem,
    QHeaderView, QMessageBox, QFileDialog, QGroupBox,
    QScrollArea, QComboBox, QTextEdit,
)
from PySide6.QtCore import Qt
import getpass

from app.logging_config import get_logger
from app.database import inserir, listar, registrar_email_enviado
from app.email_service import (
    EMAIL_RH,
    OUTLOOK_DISPONIVEL,
    email_cotacao_dell,
    enviar_email_rh,
)
from app.runtime_paths import ensure_runtime_dir, get_runtime_path, iter_config_paths

logger = get_logger(__name__)

UPLOAD_DIR = ensure_runtime_dir("uploads")


def _cfg():
    for config_file in iter_config_paths():
        if not os.path.exists(config_file):
            continue
        try:
            with open(config_file, encoding="utf-8") as file_obj:
                return json.load(file_obj)
        except Exception as err:
            logger.warning("Falha ao ler config.json no desktop: %s", err)
    return {}
 
 
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("SOMALABS - Gestao de Devolucoes")
        self.resize(1200, 800)
 
        # Criar pasta de uploads
        global UPLOAD_DIR
        UPLOAD_DIR = ensure_runtime_dir("uploads")
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

    def _create_line_edit(self, placeholder, text=""):
        line = QLineEdit()
        line.setPlaceholderText(placeholder)
        if text:
            line.setText(text)
        return line

    def _create_combo_box(self, options):
        combo = QComboBox()
        combo.addItem("— Selecione —", "")
        for option in options:
            combo.addItem(option, option)
        return combo

    def _field_value(self, widget):
        if isinstance(widget, QLineEdit):
            return widget.text().strip()
        if isinstance(widget, QComboBox):
            return (widget.currentData() or "").strip()
        if isinstance(widget, QTextEdit):
            return widget.toPlainText().strip()
        return ""

    def _reset_form(self):
        for widget in self.inputs.values():
            if isinstance(widget, QLineEdit):
                widget.clear()
            elif isinstance(widget, QComboBox):
                widget.setCurrentIndex(0)
            elif isinstance(widget, QTextEdit):
                widget.clear()
        self.caminho_foto = None
        self.label_foto.clear()
        self.btn_upload.setEnabled(False)
        self.serial_hint_label.setText("Necessário para cotação Dell.")
        self.email_hint_label.setText("Campo opcional apenas para registro. O envio automático segue a regra do status.")

    def _on_status_changed(self):
        status = self._field_value(self.inputs["status"])
        self.btn_upload.setEnabled(status == "Danificado")

    def _on_brand_changed(self):
        marca_lower = self._field_value(self.inputs["marca"]).lower()
        serial_widget = self.inputs["serial"]
        email_widget = self.inputs["email_responsavel"]

        if marca_lower == "dell":
            serial_widget.setPlaceholderText("Obrigatório para Dell")
            self.serial_hint_label.setText("Marca Dell detectada: informe a Service Tag.")
            self.email_hint_label.setText("Campo opcional apenas para registro. O envio automático segue a regra do status.")
        else:
            serial_widget.setPlaceholderText("Número de série do equipamento")
            self.serial_hint_label.setText("Necessário para cotação Dell.")
            self.email_hint_label.setText("Campo opcional apenas para registro. O envio automático segue a regra do status.")
 
    # --------------------------
    # Tela Nova Devolução
    # --------------------------
    def create_nova_page(self):
        widget = QWidget()
        outer_layout = QVBoxLayout(widget)

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setFrameShape(QScrollArea.NoFrame)

        content = QWidget()
        content_layout = QVBoxLayout(content)
        content_layout.setSpacing(14)

        header = QLabel("Registrar devolucao e abrir rascunhos do Outlook")
        header.setStyleSheet("font-size: 18px; font-weight: 700; padding: 4px 0 8px 0;")
        content_layout.addWidget(header)

        self.inputs = {}

        self.required_fields = {
            "nome", "departamento", "patrimonio", "modelo", "tipo", "marca", "status"
        }

        collaborator_box = QGroupBox("Colaborador")
        collaborator_box.setStyleSheet("QGroupBox { font-weight: 700; }")
        collaborator_layout = QFormLayout(collaborator_box)
        collaborator_layout.setSpacing(10)
        self.inputs["nome"] = self._create_line_edit("Informe nome completo")
        self.inputs["matricula"] = self._create_line_edit("Informe a matricula")
        self.inputs["departamento"] = self._create_line_edit("Ex: TI, RH, Financeiro")
        self.inputs["diretoria"] = self._create_line_edit("Ex: Diretoria de Operacoes")
        self.inputs["unidade"] = self._create_line_edit("Ex: Sede SP, Filial RJ")
        self.inputs["recebido_por"] = self._create_line_edit("Responsavel pelo recebimento")
        collaborator_layout.addRow("Entregue por", self.inputs["nome"])
        collaborator_layout.addRow("Matricula", self.inputs["matricula"])
        collaborator_layout.addRow("Departamento", self.inputs["departamento"])
        collaborator_layout.addRow("Diretoria", self.inputs["diretoria"])
        collaborator_layout.addRow("Unidade", self.inputs["unidade"])
        collaborator_layout.addRow("Recebido por", self.inputs["recebido_por"])
        content_layout.addWidget(collaborator_box)

        equipment_box = QGroupBox("Equipamento")
        equipment_box.setStyleSheet("QGroupBox { font-weight: 700; }")
        equipment_layout = QFormLayout(equipment_box)
        equipment_layout.setSpacing(10)
        self.inputs["tipo"] = self._create_combo_box([
            "Notebook", "Desktop", "Monitor", "Tablet", "Celular", "Impressora", "Roteador / Switch", "Outro"
        ])
        self.inputs["marca"] = self._create_line_edit("Ex: Dell, HP, Lenovo, Apple")
        self.inputs["marca"].textChanged.connect(self._on_brand_changed)
        self.inputs["modelo"] = self._create_line_edit("Informe o modelo")
        self.inputs["serial"] = self._create_line_edit("Obrigatorio para Dell")
        self.inputs["patrimonio"] = self._create_line_edit("Informe a tag do equipamento")
        self.inputs["processador"] = self._create_line_edit("Ex: Intel Core i5")
        self.inputs["memoria"] = self._create_line_edit("Ex: 8 GB, 16 GB")
        self.inputs["armazenamento"] = self._create_line_edit("Ex: SSD 256 GB")
        self.inputs["possui_carregador"] = self._create_combo_box(["Sim", "Não"])
        equipment_layout.addRow("Tipo", self.inputs["tipo"])
        equipment_layout.addRow("Marca", self.inputs["marca"])
        equipment_layout.addRow("Modelo", self.inputs["modelo"])
        equipment_layout.addRow("Serial / Service Tag", self.inputs["serial"])
        self.serial_hint_label = QLabel("Necessário para cotação Dell.")
        self.serial_hint_label.setWordWrap(True)
        self.serial_hint_label.setStyleSheet("color: #6b7280; font-size: 11px;")
        equipment_layout.addRow("", self.serial_hint_label)
        equipment_layout.addRow("Patrimonio", self.inputs["patrimonio"])
        equipment_layout.addRow("Processador", self.inputs["processador"])
        equipment_layout.addRow("Memoria", self.inputs["memoria"])
        equipment_layout.addRow("Armazenamento", self.inputs["armazenamento"])
        equipment_layout.addRow("Possui carregador?", self.inputs["possui_carregador"])
        content_layout.addWidget(equipment_box)

        flow_box = QGroupBox("Devolucao e Emails")
        flow_box.setStyleSheet("QGroupBox { font-weight: 700; }")
        flow_layout = QFormLayout(flow_box)
        flow_layout.setSpacing(10)
        self.inputs["email_responsavel"] = self._create_line_edit("Opcional")
        self.inputs["gestor_email"] = self._create_line_edit("Opcional")
        self.inputs["motivo"] = self._create_line_edit("Ex: Desligamento, Troca, Reforma")
        self.inputs["movido_para_estoque"] = self._create_combo_box(["Sim", "Não"])
        self.inputs["observacoes"] = QTextEdit()
        self.inputs["observacoes"].setPlaceholderText("Descreva detalhes relevantes sobre a devolucao")
        self.inputs["observacoes"].setFixedHeight(90)
        flow_layout.addRow("Email do responsavel", self.inputs["email_responsavel"])
        self.email_hint_label = QLabel("Campo opcional apenas para registro. O envio automático segue a regra do status.")
        self.email_hint_label.setWordWrap(True)
        self.email_hint_label.setStyleSheet("color: #6b7280; font-size: 11px;")
        flow_layout.addRow("", self.email_hint_label)
        flow_layout.addRow("Email do gestor", self.inputs["gestor_email"])
        flow_layout.addRow("Motivo", self.inputs["motivo"])
        flow_layout.addRow("Movido para estoque?", self.inputs["movido_para_estoque"])
        flow_layout.addRow("Observacoes", self.inputs["observacoes"])
        content_layout.addWidget(flow_box)

        status_box = QGroupBox("Status e evidencias")
        status_layout_container = QFormLayout(status_box)
        status_layout_container.setSpacing(10)

        self.inputs["status"] = self._create_combo_box([
            "OK", "Danificado", "Pendente"
        ])
        self.inputs["status"].currentIndexChanged.connect(self._on_status_changed)
        status_layout_container.addRow("Status", self.inputs["status"])
 
        # Botão de upload
        self.btn_upload = QPushButton("Selecionar Foto de Dano")
        self.btn_upload.clicked.connect(self.selecionar_foto)
        self.btn_upload.setEnabled(False)
        status_layout_container.addRow("Foto de Dano", self.btn_upload)

        # Preview da imagem
        self.label_foto = QLabel()
        self.label_foto.setFixedSize(200, 150)
        self.label_foto.setStyleSheet("border: 1px solid gray;")
        status_layout_container.addRow("Preview", self.label_foto)

        content_layout.addWidget(status_box)

        btn_processar = QPushButton("Processar Devolução")
        btn_processar.setMinimumHeight(40)
        btn_processar.clicked.connect(self.processar)
        content_layout.addWidget(btn_processar)

        content_layout.addStretch()
        scroll_area.setWidget(content)
        outer_layout.addWidget(scroll_area)

        return widget

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
            self.table.setItem(row, 0, QTableWidgetItem(linha.get("data", "")))
            self.table.setItem(row, 1, QTableWidgetItem(linha.get("nome", "")))
            self.table.setItem(row, 2, QTableWidgetItem(linha.get("patrimonio", "")))
            self.table.setItem(row, 3, QTableWidgetItem(linha.get("modelo", "")))
            self.table.setItem(row, 4, QTableWidgetItem(linha.get("status", "")))
            self.table.setItem(row, 5, QTableWidgetItem(linha.get("usuario", "")))
 
    # --------------------------
    # Ações
    # --------------------------
    def processar(self):
        usuario = getpass.getuser()

        dados = {key: self._field_value(widget) for key, widget in self.inputs.items()}
        faltantes = [campo for campo in self.required_fields if not dados.get(campo)]
        if faltantes:
            QMessageBox.warning(
                self,
                "Validação",
                f"Campos obrigatórios não preenchidos: {', '.join(faltantes)}"
            )
            return

        marca_lower = (dados.get("marca") or "").strip().lower()
        if marca_lower == "dell" and not dados.get("serial"):
            QMessageBox.warning(
                self,
                "Validação",
                "Service Tag (Serial) é obrigatório para equipamentos Dell."
            )
            return

        dados["usuario"] = usuario
        dados["registrado_por"] = usuario

        # Salvar foto localmente antes de gravar, para persistir e anexar corretamente.
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

        registro = inserir(dados)
        novo_id = registro["id"] if registro else None
        sync_error = registro.get("_sync_error") if registro else None
        if registro:
            dados.update(registro)

        if OUTLOOK_DISPONIVEL and novo_id:
            try:
                email_rh_dest = _cfg().get("email_rh") or EMAIL_RH
                destino_log = "RH"
                if dados["status"] == "Danificado" and marca_lower == "dell":
                    email_cotacao_dell(dados)
                    destino_log = "Cotação Dell + RH"

                enviar_email_rh(dados, para=email_rh_dest)
                email_sync_error = registrar_email_enviado(novo_id, destino_log)
                QMessageBox.information(
                    self,
                    "Sucesso",
                    "Processo concluido e rascunhos abertos no Outlook!"
                )
                if email_sync_error:
                    QMessageBox.warning(self, "Sincronização Excel", email_sync_error)
            except RuntimeError as err:
                logger.warning("Falha no envio de e-mail no desktop: %s", err)
                QMessageBox.warning(
                    self,
                    "Email não enviado",
                    f"A devolução foi registrada, mas o Outlook não pôde abrir o rascunho: {err}"
                )
        elif OUTLOOK_DISPONIVEL:
            QMessageBox.information(
                self,
                "Sucesso",
                "Processo concluido. (Erro ao recuperar o ID do registro para log de email.)"
            )
        else:
            QMessageBox.information(
                self,
                "Sucesso",
                "Processo concluido. (Email automatico disponivel apenas no Windows com Outlook.)"
            )

        if sync_error:
            QMessageBox.warning(self, "Sincronização Excel", sync_error)

        logger.info("Devolução registrada no desktop: %s", dados["patrimonio"])

        self.load_table()
        self._reset_form()

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