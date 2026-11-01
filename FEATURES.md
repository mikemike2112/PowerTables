# Funzionalità

Metodo dell'inventario: script `tools/inventory-ribbon.ps1` apre Excel (istanza privata, finestra visibile) su una cartella di lavoro vuota creata dallo script,
legge il ribbon con UI Automation (`System.Windows.Automation`: schede, comandi) e scrive un JSON in `tests/work/` (non committato). Schede rilevate:
Home (55 comandi), Inserisci (33), Disegno, Layout di pagina (27), Formule (27), Dati (29), Revisione (19), Visualizza (33), Automatizza, Guida.
Le voci sono raggruppate per scheda del ribbon con priorità. **Si spunta solo ciò che è implementato e coperto da un test automatico** (`dotnet test`).
Non è stato possibile pilotare la UI con mouse/tastiera in modo automatico (niente computer use): la UI è stata avviata e osservata con screenshot (`docs/`), i comportamenti sono provati sul Core e sull'xlsx riaperto in Excel.

## Home
- [x] P0 Copia/Taglia/Incolla con adattamento dei riferimenti relativi (`OpsTests`)
- [x] P0 Annulla/Ripeti (anche unioni, larghezze, riquadri: `FeatureRoundTripTests`)
- [x] P0 Grassetto, corsivo, sottolineato, barrato, tipo e dimensione carattere, colore carattere e riempimento (verificato in Excel)
- [x] P0 Bordi (sottile/medio, lati), allineamento orizzontale/verticale, testo a capo, rientro (verificato in Excel)
- [x] P0 Unisci/dividi celle (verificato in Excel)
- [x] P0 Formati numerici (generale, numero, migliaia, percentuale, scientifico, data, ora, testo, codici personalizzati), +/− decimali
- [x] P0 Larghezza colonne e altezza righe (menu e trascinamento sulle intestazioni; adatta colonna con doppio clic)
- [x] P0 Riempimento automatico (maniglia, Ctrl+D/Ctrl+R): serie numeriche, testo+numero, formule adattate, date
- [x] P0 Ordinamento (A→Z, Z→A, più chiavi nel Core, formule adattate)
- [x] P0 Filtro automatico con elenco valori (salvato come `autoFilter` + righe nascoste)
- [x] P0 Somma automatica
- [ ] P1 Copia formato, Formatta come tabella, Stili cella, Formattazione condizionale (le formattazioni condizionali esistenti nei file sono conservate, non modificabili)
- [ ] P1 Trova e sostituisci, inserisci/elimina righe-colonne-celle (con aggiornamento dei riferimenti), Incolla speciale (solo valori nel Core, senza UI)
- [ ] P2 Orientamento testo, riduci per adattare

## Formule
- [x] P0 Parser e valutatore con riferimenti relativi/assoluti/misti, intervalli, intere colonne/righe, altri fogli, nomi definiti, costanti matrice
- [x] P0 Grafo delle dipendenze, ricalcolo incrementale, rilevamento dei cicli, volatili (`TODAY`, `NOW`, `RAND`, `OFFSET`, `INDIRECT`)
- [x] P0 Errori Excel (`#DIV/0!`, `#N/A`, `#VALUE!`, `#REF!`, `#NAME?`, `#NUM!`, `#NULL!`)
- [x] P0 173 funzioni (matematica, statistica, logica, testo, ricerca, data/ora, finanza di base), 149 confrontate con Excel vero (`COMPAT.md`)
- [x] P0 Formule salvate in forma invariante con prefissi `_xlfn.` corretti; formule matriciali/dinamiche e funzioni sconosciute conservate con valore in cache
- [ ] P1 Riferimenti strutturati di tabella, riferimenti 3D, intersezione, calcolo iterativo, funzioni mancanti (statistiche avanzate, `IRR`, `RATE`, `FILTER`/`SORT`/`UNIQUE` come array dinamici, `LET`/`LAMBDA`)
- [ ] P1 Gestione nomi, Controllo errori, Valuta formula, precedenti/dipendenti

## Dati
- [x] P0 Import ed export CSV (RFC 4180, separatore rilevato, UTF-8, niente formule da dati importati)
- [x] P0 Ordina e filtra
- [ ] P1 Convalida dati (conservata nei file, non modificabile), Rimuovi duplicati, Testo in colonne, Subtotali/struttura

## Visualizza
- [x] P0 Blocca riquadri (riga, colonna, fino alla selezione, sblocca; verificato in Excel)
- [x] P0 Più fogli (nuovo foglio, rinomina con aggiornamento delle formule), griglia on/off, tema scuro/chiaro
- [ ] P1 Zoom, divisione finestra, nascondi/mostra righe-colonne-fogli da UI, elimina/sposta foglio
- [ ] P1 Anteprima interruzioni di pagina, layout di pagina, stampa, esportazione PDF

## Griglia e prestazioni
- [x] P0 Griglia virtualizzata con `FrameworkElement` + `DrawingContext`, store di celle sparso, test a 1.000.000 di righe popolate (`COMPAT.md`)
- [x] P0 Tipi di cella: numero, testo, data (formati), booleano, errore
- [x] P0 Selezione, tastiera (frecce, Ctrl+frecce, Maiusc, Pag, Home/Fine, Tab/Invio, F2, Canc, digitazione), mouse (selezione, trascinamento, intestazioni, ridimensionamento), formula bar, casella nome

## Inserisci / Layout / Revisione / Disegno
- [ ] P1 Grafici, immagini, forme, tabelle pivot, commenti, collegamenti: **conservati intatti nel file** (round-trip provato), non creabili né visibili nell'app
- [ ] P2 Pivot, Power Query, macro (nessuna esecuzione di contenuto attivo), revisione, input penna
