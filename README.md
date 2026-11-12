# danea-import-automation ☕

Strumento di automazione con interfaccia grafica (GUI) per l'importazione di fatture e DDT di vari fornitori nel formato compatibile con **Danea Easyfatt**. Traccia lo storico dei prezzi in un database SQLite locale e segnala eventuali aumenti di prezzo durante l'elaborazione.

## 🚀 Caratteristiche

- **Supporto multi-fornitore**: parsing automatico di listini/fatture/DDT in formato PDF, Excel, TXT o CSV per fornitori come *Ferrari, Viridex, Pilato, Monge, Capaldo, Blumen, Verisem, Giardini, Garver, Record, Ricci*.
- **Controllo Storico Prezzi**: salva i prezzi d'acquisto in un database locale (`storico_prezzi.db`) e inserisce una nota di avviso nel file Danea se rileva un aumento rispetto all'ultimo acquisto.
- **Interfaccia Grafica Intuitiva**: seleziona la cartella di lavoro, visualizza i file da elaborare e monitora il processo in tempo reale tramite la console integrata.
- **Privacy al 100%**: Tutti i dati dei prezzi e i file elaborati rimangono esclusivamente sul tuo computer locale.

---

## ☕ Supporta il progetto

Se questo strumento ti ha fatto risparmiare tempo e semplificato il lavoro, puoi offrirmi un caffè per supportare lo sviluppo!

[!["PayPal"](https://img.shields.io/badge/Donate-PayPal-green.svg)](https://www.paypal.me/gates538)
[!["GitHub Sponsors"](https://img.shields.io/badge/GitHub%20Sponsors-Support-EA4AAA?logo=githubsponsors&logoColor=white)](https://github.com/sponsors/gates538)


---

## 🛠️ Installazione e Uso

1. Assicurati di avere installato **Python 3**.
2. Avvia lo script facendo doppio clic su `import_to_danea.py` o eseguendo da terminale:
   ```bash
   python import_to_danea.py
   ```
3. Seleziona la cartella contenente i documenti da convertire e clicca su **Avvia Elaborazione**. I file pronti per l'importazione in Danea verranno salvati nella cartella `Lavorati/`.
