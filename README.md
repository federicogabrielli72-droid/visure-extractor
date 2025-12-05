# Estrazione Automatica Visure Camerali – App PyQt5

Questa applicazione permette di:

- Selezionare uno o più PDF di visure camerali
- Estrarre automaticamente dati dell'impresa, soci e amministratori
- Generare un file Excel con i fogli:
  - Imprese
  - Soci
  - Amministratori

## Come funziona la compilazione automatica

Il repository contiene un workflow GitHub Actions che:

- gira su Windows
- installa Python 3.x
- installa PyQt5, pdfplumber, pandas
- usa PyInstaller per creare un eseguibile Windows (.exe)
- pubblica l'eseguibile negli Artifacts del workflow

Per ottenere l'eseguibile:

1. Vai su "Actions"
2. Apri l'ultimo workflow "Build EXE"
3. Scarica `app_visure_pyqt.exe` dagli Artifacts

## Come usare l'app

1. Avvia `app_visure_pyqt.exe`
2. Seleziona i PDF di visura camerale
3. Seleziona la cartella di output
4. Clicca "Genera Excel"
5. Verrà creato `database_visure.xlsx`
