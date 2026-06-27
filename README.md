# 🖨🔥 STAMPA FIRE

**Sysadmin Tools Against Malicious Printer Agon Force Inf Registration Engine**

![versione](https://img.shields.io/badge/versione-0.0.2-blue) ![piattaforma](https://img.shields.io/badge/Windows-10%20%7C%2011-0078D6) ![licenza](https://img.shields.io/badge/licenza-GPLv3-green)

Tool grafico per l'installazione **in batch** di stampanti di rete su Windows 10/11,
configurabile tramite un semplice file Excel.

STAMPA FIRE trasforma il dramma dell'installazione massiva di stampanti in un processo
lineare e ripetibile. Mappi IP, driver e posizioni in un foglio Excel e il tool esegue
in automatico l'intero flusso: ping della macchina, iniezione del pacchetto driver,
creazione della porta TCP/IP e registrazione della stampante nel sistema — tutto
tracciato in file di log con timestamp.

📄 **[Guida di utilizzo](https://denvermotel.github.io/stampa-fire/#guida)** ·
🌐 **[Presentazione](https://denvermotel.github.io/stampa-fire/)**

---

## 🚀 Avvio rapido

**Requisiti:** Windows 10/11.

1. Scarica **`StampaFire.exe`** (64bit, autonomo).
2. Fai doppio click e conferma il prompt **UAC**: i privilegi di amministratore sono
   obbligatori per installare driver e stampanti.
3. Metti un file di configurazione (`printers.xlsx` o `printers.csv`) **accanto
   all'eseguibile** per caricarlo in automatico, oppure selezionalo dal pulsante
   **Sfoglia**. Un modello pronto è scaricabile dalla
   [guida](https://denvermotel.github.io/stampa-fire/#guida).
4. Seleziona le stampanti e premi **Installa selezionate**.

📖 Istruzioni dettagliate (file di configurazione, colonne, esempi):
**[Guida completa](https://denvermotel.github.io/stampa-fire/#guida)**.

---

## 📋 Changelog

Sunto delle novità. Changelog completo in
**[CHANGELOG.md](CHANGELOG.md)**.

### v0.0.2
- Gestione delle **stampanti installate**: visualizzazione e disinstallazione con
  scelta dell'ambito (solo stampante / +porta / +driver)
- **Esportazione** delle stampanti installate in CSV/Excel per reimportarle su un altro PC
- **Importazione da CSV** oltre che da Excel (separatore `;`/`,` rilevato automaticamente)
- **Modelli pronti** `printers_template.xlsx` e `.csv`, scaricabili anche dalla guida
- **Schermata di caricamento** all'avvio (niente più finestra "Non risponde")

### v0.0.1
- Prima versione: GUI Tkinter, configurazione da `.xlsx`, installazione via
  `pnputil` + PowerShell, ping opzionale, barra di avanzamento, log con timestamp,
  elevazione UAC

---

Distribuito sotto licenza GPLv3. Software fornito «così com'è», senza garanzie:
l'utente è l'unico responsabile dell'uso che ne viene fatto.
