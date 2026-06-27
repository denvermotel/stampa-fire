# Changelog

Tutte le modifiche rilevanti a **STAMPA FIRE** sono documentate in questo file.


## [0.0.2]

### Aggiunto
- **Schermata di caricamento (splash)**:
  il file di configurazione viene letto in background, evitando che l'applicazione risulti 
  in stato di "Non risponde" durante avvii lenti o con Excel corposi.
- **Gestione delle stampanti installate**: finestra dedicata per visualizzare le
  stampanti di rete presenti nel sistema e disinstallarle, con scelta dell'ambito
  (solo stampante / anche porta TCP/IP / anche driver). Porta e driver vengono
  rimossi solo se non più usati da altre stampanti.
- **Esportazione** delle stampanti installate in CSV o Excel, per reimportarle su
  un altro PC. I campi `cartella_driver` e `file_inf` (percorso del driver) restano
  vuoti perché non ricavabili dal sistema e vanno compilati manualmente.
- **Importazione da CSV** oltre che da Excel, con rilevamento automatico del
  separatore (`;` o `,`).
- **Modelli pronti** `printers_template.xlsx` e `printers_template.csv` (stessi dati),
  scaricabili anche dalla guida online.
- **Documentazione pubblica**: README e sito con guida (GitHub Pages).

## [0.0.1]

### Aggiunto
- Prima versione: interfaccia grafica Tkinter, lettura della configurazione da
  Excel, installazione in batch tramite `pnputil` e PowerShell.
- Test di connessione (ping) opzionale prima di ogni installazione.
- Barra di avanzamento, log con timestamp e riepilogo per stampante.
- Richiesta automatica dei privilegi di amministratore (elevazione UAC).
