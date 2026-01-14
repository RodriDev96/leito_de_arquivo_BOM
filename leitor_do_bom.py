import csv
import sys
import os
import json
import shutil

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QFileDialog,
    QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QLineEdit, QTableWidget, QTableWidgetItem,
    QMessageBox, QComboBox, QInputDialog
)
from PySide6.QtCore import QSettings, Qt
from PySide6.QtPrintSupport import QPrinter, QPrintDialog
from PySide6.QtGui import QTextDocument


# =============================
# CONFIGURAÇÕES
# =============================
SENHA_DEV = "1234"

BASE_DIR = os.path.join(
    os.path.expanduser("~"),
    "Documents",
    "LeitorDoBOM"
)

PROD_DIR = os.path.join(BASE_DIR, "produtos")
PROD_FILE = os.path.join(PROD_DIR, "produtos.json")

os.makedirs(PROD_DIR, exist_ok=True)


# =============================
# TEMA
# =============================
DARK_STYLE = """
QWidget {
    background-color: #121212;
    color: #e0e0e0;
    font-size: 12px;
}
QLineEdit, QTableWidget, QComboBox {
    background-color: #1e1e1e;
    border: 1px solid #333;
}
QHeaderView::section {
    background-color: #222;
    padding: 4px;
    border: 1px solid #333;
}
QPushButton {
    background-color: #2c2c2c;
    border: 1px solid #444;
    padding: 6px;
}
QPushButton:hover {
    background-color: #3a3a3a;
}
"""

LIGHT_STYLE = ""


# =============================
# BOM – UTILIDADES
# =============================
def extrair_valor(comment):
    if not comment:
        return "—"
    partes = comment.split()
    return partes[-1] if len(partes) > 1 else "—"


def extrair_tolerancia(comment):
    if not comment:
        return "—"
    txt = comment.upper()
    if "0R" in txt:
        return "±5%"
    if any(x in txt for x in ["R", "K", "M"]):
        return "±1%"
    return "—"


# =============================
# LEITURA TX400
# =============================
def carregar_tx400(arquivo):
    feeders = {}
    comps = []

    with open(arquivo, encoding="utf-8") as f:
        reader = csv.reader(f)
        for row in reader:
            if not row:
                continue

            if row[0] == "Feeder" and len(row) >= 8:
                feeders[row[1]] = {
                    "Comment": row[7],
                    "Valor": extrair_valor(row[7]),
                    "Tolerância": extrair_tolerancia(row[7]),
                }

            elif row[0] == "Comp" and len(row) >= 5:
                comps.append({
                    "Designator": row[4].upper(),
                    "Feeder ID": row[1],
                    "Footprint": row[3],
                })

    return feeders, comps


# =============================
# VALIDAÇÃO CSV
# =============================
def validar_csv_tx400(arquivo):
    erros = []
    feeders = set()
    comp_feeders = []

    try:
        with open(arquivo, encoding="utf-8") as f:
            reader = csv.reader(f)
            for linha, row in enumerate(reader, start=1):
                if not row:
                    continue

                if row[0] == "Feeder":
                    if len(row) < 8:
                        erros.append(f"Linha {linha}: Feeder incompleto")
                    else:
                        feeders.add(row[1])

                elif row[0] == "Comp":
                    if len(row) < 5:
                        erros.append(f"Linha {linha}: Comp incompleto")
                        continue

                    if not row[1]:
                        erros.append(f"Linha {linha}: Feeder ID vazio")

                    if not row[4]:
                        erros.append(f"Linha {linha}: Designator vazio")

                    comp_feeders.append(row[1])

        if not feeders:
            erros.append("Nenhum Feeder encontrado")

        if not comp_feeders:
            erros.append("Nenhum componente encontrado")

        for f in comp_feeders:
            if f not in feeders:
                erros.append(f"Feeder ID inexistente: {f}")

    except Exception as e:
        erros.append(str(e))

    return erros


# =============================
# PRODUTOS
# =============================
def carregar_produtos():
    if not os.path.exists(PROD_FILE):
        return {}
    with open(PROD_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


def salvar_produtos(produtos):
    with open(PROD_FILE, "w", encoding="utf-8") as f:
        json.dump(produtos, f, indent=2, ensure_ascii=False)


# =============================
# JANELA PRINCIPAL
# =============================
class LeitorBOM(QMainWindow):
    def __init__(self):
        super().__init__()

        self.settings = QSettings("TX400", "LeitorDoBOM")
        self.feeders = {}
        self.comps = []
        self.dark_mode = False

        self.setWindowTitle("Leitor do BOM")
        self.setGeometry(100, 100, 1000, 550)

        self._ui()
        self.carregar_preferencias()
        self.atualizar_lista_produtos()

    # =============================
    # UI
    # =============================
    def _ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        top = QHBoxLayout()
        self.combo_produto = QComboBox()
        self.combo_produto.currentTextChanged.connect(self.carregar_produto)

        btn_dev = QPushButton("⚙ Desenvolvedor")
        btn_dev.clicked.connect(self.modo_desenvolvedor)

        self.btn_tema = QPushButton("🌙 Tema escuro")
        self.btn_tema.clicked.connect(self.alternar_tema)

        top.addWidget(QLabel("Produto:"))
        top.addWidget(self.combo_produto)
        top.addStretch()
        top.addWidget(btn_dev)
        top.addWidget(self.btn_tema)
        layout.addLayout(top)

        busca = QHBoxLayout()
        busca.addWidget(QLabel("Buscar:"))

        self.edt_busca = QLineEdit()
        self.edt_busca.returnPressed.connect(self.buscar)

        btn_buscar = QPushButton("Buscar")
        btn_buscar.clicked.connect(self.buscar)

        busca.addWidget(self.edt_busca)
        busca.addWidget(btn_buscar)
        busca.addStretch()
        layout.addLayout(busca)

        btn_pedido = QPushButton("📦 Gerar Pedido / Imprimir")
        btn_pedido.clicked.connect(self.gerar_pedido)
        layout.addWidget(btn_pedido)

        self.tabela = QTableWidget()
        self.tabela.setColumnCount(7)
        self.tabela.setHorizontalHeaderLabels([
            "✔", "Designator", "Valor", "Tolerância",
            "Footprint", "Feeder", "Comentário"
        ])
        self.tabela.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(self.tabela)

    # =============================
    # PRODUTOS
    # =============================
    def atualizar_lista_produtos(self):
        self.combo_produto.blockSignals(True)
        self.combo_produto.clear()
        self.combo_produto.addItem("Selecione...")
        for nome in carregar_produtos():
            self.combo_produto.addItem(nome)
        self.combo_produto.blockSignals(False)

    def carregar_produto(self, nome):
        if nome == "Selecione...":
            return
        produtos = carregar_produtos()
        caminho = os.path.join(PROD_DIR, produtos[nome])
        self.feeders, self.comps = carregar_tx400(caminho)
        self.buscar()

    # =============================
    # MODO DESENVOLVEDOR
    # =============================
    def modo_desenvolvedor(self):
        senha, ok = QInputDialog.getText(
            self, "Acesso restrito", "Senha do desenvolvedor:",
            QLineEdit.Password
        )
        if not ok or senha != SENHA_DEV:
            QMessageBox.critical(self, "Acesso negado", "Senha incorreta.")
            return

        acao, ok = QInputDialog.getItem(
            self, "Modo Desenvolvedor",
            "Escolha a ação:",
            ["Adicionar produto", "Excluir produto"],
            0, False
        )
        if not ok:
            return

        if acao == "Adicionar produto":
            self.adicionar_produto()
        else:
            self.excluir_produto()

    def adicionar_produto(self):
        nome, ok = QInputDialog.getText(self, "Novo produto", "Nome do produto:")
        if not ok or not nome.strip():
            return

        produtos = carregar_produtos()
        if nome in produtos:
            QMessageBox.warning(self, "Erro", "Produto já cadastrado.")
            return

        caminho, _ = QFileDialog.getOpenFileName(
            self, "Selecionar TX400 CSV", "", "CSV (*.csv)"
        )
        if not caminho:
            return

        erros = validar_csv_tx400(caminho)
        if erros:
            QMessageBox.critical(
                self, "CSV inválido",
                "O arquivo possui erros:\n\n" + "\n".join(erros[:10])
            )
            return

        nome_arquivo = nome.lower().replace(" ", "_") + ".csv"
        shutil.copy(caminho, os.path.join(PROD_DIR, nome_arquivo))

        produtos[nome] = nome_arquivo
        salvar_produtos(produtos)

        QMessageBox.information(self, "OK", f"Produto '{nome}' cadastrado.")
        self.atualizar_lista_produtos()

    def excluir_produto(self):
        produtos = carregar_produtos()
        if not produtos:
            QMessageBox.information(self, "Aviso", "Nenhum produto cadastrado.")
            return

        nome, ok = QInputDialog.getItem(
            self, "Excluir produto",
            "Selecione o produto:",
            list(produtos.keys()),
            0, False
        )
        if not ok:
            return

        resp = QMessageBox.question(
            self, "Confirmação",
            f"Excluir o produto '{nome}'?",
            QMessageBox.Yes | QMessageBox.No
        )
        if resp != QMessageBox.Yes:
            return

        try:
            os.remove(os.path.join(PROD_DIR, produtos[nome]))
        except:
            pass

        del produtos[nome]
        salvar_produtos(produtos)

        QMessageBox.information(self, "OK", "Produto excluído.")
        self.atualizar_lista_produtos()

    # =============================
    # BUSCA
    # =============================
    def buscar(self):
        termo = self.edt_busca.text().strip().upper()
        encontrados = self.comps if not termo else [
            c for c in self.comps if termo in c["Designator"]
        ]

        self.tabela.setRowCount(len(encontrados))

        for r, c in enumerate(encontrados):
            f = self.feeders.get(c["Feeder ID"], {})

            chk = QTableWidgetItem()
            chk.setCheckState(Qt.Unchecked)
            self.tabela.setItem(r, 0, chk)

            dados = [
                c["Designator"],
                f.get("Valor", "—"),
                f.get("Tolerância", "—"),
                c["Footprint"],
                c["Feeder ID"],
                f.get("Comment", "—"),
            ]

            for col, v in enumerate(dados, start=1):
                self.tabela.setItem(r, col, QTableWidgetItem(v))

    # =============================
    # PEDIDO / IMPRESSÃO
    # =============================
    def gerar_pedido(self):
        selecionados = []

        for row in range(self.tabela.rowCount()):
            item = self.tabela.item(row, 0)
            if item and item.checkState() == Qt.Checked:
                selecionados.append(row)

        if not selecionados:
            QMessageBox.warning(self, "Aviso", "Nenhum componente selecionado.")
            return

        qtd, ok = QInputDialog.getInt(
            self, "Quantidade", "Quantidade desejada:", 100, 1
        )
        if not ok:
            return

        texto = "PEDIDO DE COMPONENTES\n\n"

        for row in selecionados:
            texto += (
                f"{self.tabela.item(row,1).text()} | "
                f"{self.tabela.item(row,2).text()} | "
                f"{self.tabela.item(row,3).text()} | "
                f"{self.tabela.item(row,4).text()} | "
                f"{self.tabela.item(row,6).text()} | "
                f"QTD: {qtd}\n"
            )

        acao, ok = QInputDialog.getItem(
            self, "Pedido",
            "O que deseja fazer?",
            ["Salvar em TXT", "Imprimir direto"],
            0, False
        )
        if not ok:
            return

        if acao == "Salvar em TXT":
            caminho, _ = QFileDialog.getSaveFileName(
                self, "Salvar pedido", "pedido_componentes.txt", "TXT (*.txt)"
            )
            if not caminho:
                return

            with open(caminho, "w", encoding="utf-8") as f:
                f.write(texto)

            QMessageBox.information(self, "OK", "Pedido salvo com sucesso.")
        else:
            self.imprimir_texto(texto)

    def imprimir_texto(self, texto):
        printer = QPrinter(QPrinter.HighResolution)
        dialog = QPrintDialog(printer, self)

        if dialog.exec() != QPrintDialog.Accepted:
            return

        doc = QTextDocument()
        doc.setPlainText(texto)
        doc.print(printer)

        QMessageBox.information(self, "Impressão", "Pedido enviado para a impressora.")

    # =============================
    # TEMA
    # =============================
    def carregar_preferencias(self):
        self.dark_mode = self.settings.value("dark_mode", False, type=bool)
        self.aplicar_tema()

    def alternar_tema(self):
        self.dark_mode = not self.dark_mode
        self.aplicar_tema()
        self.settings.setValue("dark_mode", self.dark_mode)

    def aplicar_tema(self):
        if self.dark_mode:
            self.setStyleSheet(DARK_STYLE)
            self.btn_tema.setText("🌞 Tema claro")
        else:
            self.setStyleSheet(LIGHT_STYLE)
            self.btn_tema.setText("🌙 Tema escuro")


# =============================
# MAIN
# =============================
if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = LeitorBOM()
    win.show()
    sys.exit(app.exec())
