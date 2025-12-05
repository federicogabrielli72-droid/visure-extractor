import sys
import os
import re
import pdfplumber
import pandas as pd

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QListWidget, QFileDialog, QLabel, QProgressBar,
    QTextEdit, QMessageBox
)
from PyQt5.QtCore import Qt


# ----------------- FUNZIONI DI SUPPORTO ----------------- #

def trova(pattern, testo, flags=re.IGNORECASE):
    """Trova la prima occorrenza di un pattern regex."""
    m = re.search(pattern, testo, flags)
    return m.group(1).strip() if m else ""


def estrai_dati_visura(pdf_path):
    """
    Estrae i dati da una visura camerale in PDF.
    Funziona con layout come C.M. IMPRESA EDILE SRL e NEWTON S.R.L.
    """
    with pdfplumber.open(pdf_path) as pdf:
        testo = "\n".join(page.extract_text() or "" for page in pdf.pages)

    dati = {}

    # Ragione sociale (prima linea grande oppure sezione 'Denominazione')
    ragione = trova(r"VISURA[^\n]*\n([A-Z0-9\.\s'&]+S\.?R\.?L\.?)", testo)
    if not ragione:
        ragione = trova(r"Denominazione:\s*([A-Z0-9\.\s'&]+)", testo)
    dati["ragione_sociale"] = ragione

    # Codice fiscale e Partita IVA
    dati["codice_fiscale"] = trova(r"Codice fiscale(?: e n\.iscr\. al Registro Imprese)?\s*[:\s]*([A-Z0-9]+)", testo)
    dati["piva"] = trova(r"Partita IVA\s*([0-9]+)", testo)

    # Forma giuridica
    dati["forma_giuridica"] = trova(r"Forma giuridica\s*(.+)", testo)

    # PEC
    dati["pec"] = trova(r"(?:PEC|Domicilio digitale\/PEC)\s*([A-Za-z0-9\.\-_@]+)", testo)

    # REA
    dati["rea"] = trova(r"Numero REA\s*([A-Z]+\s*-\s*\d+)", testo)

    # Date
    dati["data_costituzione"] = trova(r"Data atto di costituzione\s*([\d\/\.]+)", testo)
    dati["data_iscrizione"] = trova(r"Data iscrizione\s*([\d\/\.]+)", testo)
    dati["data_ultimo_protocollo"] = trova(r"Data ultimo protocollo\s*([\d\/\.]+)", testo)

    # Stato attività
    dati["stato_attivita"] = trova(r"Stato attività\s*([A-Za-z]+)", testo)
    dati["data_inizio_attivita"] = trova(r"Data inizio attività\s*([\d\/\.]+)", testo)

    # Attività prevalente e ATECO
    dati["attivita_prevalente"] = trova(r"Attività prevalente\s*(.+)", testo)
    dati["codice_ateco"] = trova(r"Codice ATECO(?: 2\.1)?\s*([0-9\.\-]+)", testo)

    # Capitale sociale
    cap = trova(r"Capitale sociale(?: sottoscritto)?\s*([0-9\.\,]+)", testo)
    if not cap:
        cap = trova(r"Capitale sociale in Euro\s*Deliberato:\s*([0-9\.\,]+)", testo)
    dati["capitale_sociale"] = cap

    # Indirizzo sede legale
    match_ind = re.search(
        r"Indirizzo Sede legale\s*(.+?)(?:Domicilio digitale\/PEC|PEC|Numero REA|Partita IVA)",
        testo, re.IGNORECASE | re.DOTALL
    )
    indirizzo_sede = ""
    if match_ind:
        blocco = match_ind.group(1).strip().replace("\n", " ")
        indirizzo_sede = re.sub(r"\s+", " ", blocco)
    dati["indirizzo_sede_legale"] = indirizzo_sede

    # Numero soci / amministratori (se presenti)
    dati["num_soci"] = trova(r"Soci e titolari di diritti su\s*azioni e quote\s*(\d+)", testo)
    dati["num_amministratori"] = trova(r"Amministratori\s*(\d+)", testo)

    # ----------------- ESTRAZIONE SOCI ----------------- #
    soci = []
    pattern_soci_blocchi = re.findall(
        r"([A-Z'À-Ü\s]+)\s+Codice fiscale:\s*([A-Z0-9]+)(?:[\s\S]*?Quota di nominali:\s*([0-9\.\,]+)\s*Euro)?(?:[\s\S]*?(\d{1,3})\s*%)?",
        testo,
        re.IGNORECASE
    )

    for nome, cf_socio, quota, perc in pattern_soci_blocchi:
        soci.append({
            "nome": " ".join(nome.split()).title(),
            "codice_fiscale": cf_socio,
            "quota_euro": quota or "",
            "percentuale": perc or ""
        })

    # ----------------- ESTRAZIONE AMMINISTRATORI ----------------- #
    amministratori = []
    pattern_amm = re.findall(
        r"Amministratore\s+([A-Z'À-Ü\s]+)[\s\S]*?Codice fiscale:\s*([A-Z0-9]+)[\s\S]*?domicilio\s*(.+?)\s*carica",
        testo,
        re.IGNORECASE
    )

    for nome_amm, cf_amm, dom in pattern_amm:
        amministratori.append({
            "nome": " ".join(nome_amm.split()).title(),
            "codice_fiscale": cf_amm,
            "domicilio": " ".join(dom.replace("\n", " ").split())
        })

    return dati, soci, amministratori


# ----------------- INTERFACCIA GRAFICA (PyQt5) ----------------- #

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Estrazione Visure Camerali → Excel")
        self.setMinimumSize(750, 550)

        self.pdf_files = []
        self.output_folder = os.path.expanduser("~")

        self.build_ui()

    def build_ui(self):
        central = QWidget()
        layout = QVBoxLayout()

        # Bottone selezione PDF
        btn_sel_pdf = QPushButton("Seleziona PDF delle Visure…")
        btn_sel_pdf.clicked.connect(self.seleziona_pdf)
        layout.addWidget(btn_sel_pdf)

        # Lista PDF
        self.list_pdf = QListWidget()
        layout.addWidget(QLabel("File selezionati:"))
        layout.addWidget(self.list_pdf)

        # Selezione cartella output
        h = QHBoxLayout()
        self.lbl_output = QLabel(f"Cartella output: {self.output_folder}")
        btn_out = QPushButton("Cambia cartella…")
        btn_out.clicked.connect(self.seleziona_output)
        h.addWidget(self.lbl_output)
        h.addWidget(btn_out)
        layout.addLayout(h)

        # Bottone genera Excel
        btn_run = QPushButton("Genera Excel")
        btn_run.clicked.connect(self.genera_excel)
        layout.addWidget(btn_run)

        # Barra di progresso
        self.progress = QProgressBar()
        self.progress.setValue(0)
        layout.addWidget(self.progress)

        # Log
        layout.addWidget(QLabel("Log:"))
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        layout.addWidget(self.log)

        central.setLayout(layout)
        self.setCentralWidget(central)

    # ----------------- Metodi UI ----------------- #

    def log_msg(self, msg):
        self.log.append(msg)
        self.log.ensureCursorVisible()

    def aggiorna_progress(self, value):
        self.progress.setValue(value)

    def seleziona_pdf(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "Scegli i PDF", "", "File PDF (*.pdf)"
        )
        if files:
            self.pdf_files = files
            self.list_pdf.clear()
            for f in files:
                self.list_pdf.addItem(f)
            self.log_msg(f"{len(files)} PDF selezionati.")

    def seleziona_output(self):
        folder = QFileDialog.getExistingDirectory(
            self, "Seleziona cartella output", self.output_folder
        )
        if folder:
            self.output_folder = folder
            self.lbl_output.setText(f"Cartella output: {folder}")

    def genera_excel(self):
        if not self.pdf_files:
            QMessageBox.warning(self, "Errore", "Seleziona almeno un PDF.")
            return

        output_path = os.path.join(self.output_folder, "database_visure.xlsx")

        self.progress.setValue(0)
        self.log.clear()
        self.log_msg("Inizio elaborazione…")

        imprese = []
        soci_all = []
        amm_all = []

        try:
            total = len(self.pdf_files)
            for idx, pdf in enumerate(self.pdf_files, start=1):
                self.log_msg(f"Leggo: {os.path.basename(pdf)}")

                dati, soci, amm = estrai_dati_visura(pdf)

                id_impresa = len(imprese) + 1
                dati["id_impresa"] = id_impresa
                dati["file_origine"] = os.path.basename(pdf)
                imprese.append(dati)

                for s in soci:
                    s["id_impresa"] = id_impresa
                    soci_all.append(s)

                for a in amm:
                    a["id_impresa"] = id_impresa
                    amm_all.append(a)

                self.aggiorna_progress(int(idx / total * 100))
                self.log_msg("OK")

            # Esporta Excel
            with pd.ExcelWriter(output_path) as writer:
                pd.DataFrame(imprese).to_excel(writer, sheet_name="Imprese", index=False)
                pd.DataFrame(soci_all).to_excel(writer, sheet_name="Soci", index=False)
                pd.DataFrame(amm_all).to_excel(writer, sheet_name="Amministratori", index=False)

            self.log_msg(f"COMPLETATO. File creato: {output_path}")
            QMessageBox.information(self, "Fatto", f"File Excel generato:\n{output_path}")

        except Exception as e:
            self.log_msg(f"ERRORE: {e}")
            QMessageBox.critical(self, "Errore", str(e))


# ----------------- MAIN ----------------- #

def main():
    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
