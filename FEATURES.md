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
- [x] P1 Copia formato (pulsante, poi clic/selezione di destinazione), Formattazione condizionale (maggiore/minore/tra/uguale/testo contiene, cancella regole), Formatta come tabella e Tabella (`TablesTests`), Stili cella (preimpostazioni), Formato celle (Ctrl+1) (`DataToolsTests`, `SheetFeaturesTests`)
- [x] P1 Incolla ▾ (Incolla, valori, formati), Trova e seleziona ▾ (Trova, Sostituisci, Vai a), Somma automatica (selezione, colonna sopra, riga a sinistra), Riempi giù/Ctrl+D e Ctrl+R con una sola cella, menu contestuale della griglia e delle schede foglio
- [x] P1 Inserisci/elimina righe e colonne (Home > Celle; aggiorna formule di tutti i fogli, unioni, larghezze, filtro; riferimenti eliminati → `#REF!`; annullabile; `LinesFindTests`)
- [x] P1 Trova e sostituisci (Ctrl+F, Ctrl+H, Home > Trova e seleziona; maiuscole, cella intera, formule incluse, sostituisci tutto in un solo Annulla; `LinesFindTests`)
- [ ] P1 Inserisci/elimina celle singole, Incolla speciale (solo valori nel Core, senza UI)
- [ ] P2 Orientamento testo, riduci per adattare

## Formule
- [x] P0 Parser e valutatore con riferimenti relativi/assoluti/misti, intervalli, intere colonne/righe, altri fogli, nomi definiti, costanti matrice
- [x] P0 Grafo delle dipendenze, ricalcolo incrementale, rilevamento dei cicli, volatili (`TODAY`, `NOW`, `RAND`, `OFFSET`, `INDIRECT`)
- [x] P0 Errori Excel (`#DIV/0!`, `#N/A`, `#VALUE!`, `#REF!`, `#NAME?`, `#NUM!`, `#NULL!`)
- [x] P0 173 funzioni (matematica, statistica, logica, testo, ricerca, data/ora, finanza di base), 149 confrontate con Excel vero (`COMPAT.md`)
- [x] P0 Formule salvate in forma invariante con prefissi `_xlfn.` corretti; formule matriciali/dinamiche e funzioni sconosciute conservate con valore in cache
- [x] P0 Elenco funzioni scrivendo `=` (filtro mentre si digita, Tab completa), selezione dinamica dei riferimenti (clic, trascinamento, frecce, Maiusc+frecce; mai assoluti da soli; il clic in un punto non valido non conferma)
- [x] P0 Scheda Formule: Inserisci funzione (ricerca per nome inglese/italiano) e categorie Finanziarie, Logiche, Testo, Data e ora, Ricerca e riferimento, Matematiche, Statistiche
- [x] P0 Nomi di funzione italiani e separatore `;` con virgola decimale digitati a mano (`=SOMMA(A1;A2;0,5)`): nel file restano nomi inglesi e virgola
- [ ] P1 Riferimenti strutturati di tabella, riferimenti 3D, intersezione, calcolo iterativo, funzioni mancanti (statistiche avanzate, `IRR`, `RATE`, `FILTER`/`SORT`/`UNIQUE` come array dinamici, `LET`/`LAMBDA`)
- [x] P1 Definisci nome e Gestione nomi (`NamesTests`)
- [ ] P1 Controllo errori, Valuta formula, precedenti/dipendenti

## Dati
- [x] P0 Import ed export CSV (RFC 4180, separatore rilevato, UTF-8, niente formule da dati importati)
- [x] P0 Ordina e filtra
- [x] P1 Convalida dati (intero, decimale, lunghezza, elenco), Rimuovi duplicati, Testo in colonne, Cancella filtro (solo criteri) (`SheetFeaturesTests`, `DataToolsTests`, `SheetOpsTests`)
- [ ] P1 Subtotali/struttura, Avanzate (filtro)

## Visualizza
- [x] P0 Blocca riquadri (riga, colonna, fino alla selezione, sblocca; verificato in Excel)
- [x] P0 Più fogli (nuovo foglio, rinomina con aggiornamento delle formule), griglia on/off, tema scuro/chiaro
- [x] P1 Nuovo/elimina/sposta/rinomina foglio annullabili (`SheetOpsTests`), Normale (unica vista)
- [ ] P1 Divisione finestra, nascondi/mostra righe-colonne-fogli da UI, Intestazioni on/off, Nuova finestra
- [x] P1 Margini, orientamento, dimensioni carta, area di stampa, griglia in stampa, Esporta PDF e Stampa (via PDF) (`PdfExportTests`)
- [ ] P1 Anteprima interruzioni di pagina, Layout di pagina (vista), adatta a pagine, intestazioni/piè di pagina

## Griglia e prestazioni
- [x] P0 Griglia virtualizzata con `FrameworkElement` + `DrawingContext`, store di celle sparso, test a 1.000.000 di righe popolate (`COMPAT.md`)
- [x] P0 Tipi di cella: numero, testo, data (formati), booleano, errore
- [x] P0 Selezione, tastiera (frecce, Ctrl+frecce, Maiusc, Pag, Home/Fine, Tab/Invio, F2, Canc, digitazione), mouse (selezione, trascinamento, intestazioni, ridimensionamento), formula bar, casella nome

## Inserisci / Layout / Revisione / Disegno
- [x] P1 Collegamento (esterno http/https/ftp/mailto o interno; Ctrl+clic)
- [ ] P1 Grafici, immagini, forme, tabelle pivot, commenti: **conservati intatti nel file** (round-trip provato), non creabili né visibili nell'app. Rimasti inerti nella ribbon: Tabella pivot, Immagini, Forme, Grafici consigliati/Istogramma/Linee, Commento, Casella di testo, Temi, Analisi di simulazione, Foglio previsione, Nuovo script/Visualizza script, Nuova finestra, Avanzate, Anteprima interruzioni, Layout di pagina (vista), Intestazioni
- [ ] P2 Pivot, Power Query, macro (nessuna esecuzione di contenuto attivo), revisione, input penna

## Stato round-trip (aggiornamento)
- [x] Testo ricco delle celle conservato al salvataggio anche nelle righe riscritte (test su Excel vero).
- I test UI con input reale (UiTests) sono ora con finestra Topmost e ripetizioni su condizione: 3 esecuzioni consecutive isolate 3/3; nella suite completa ne è fallito uno su una singola esecuzione (vedi report).

## Aggiornamento giro "todo"
- [x] Schede Disegno e Revisione e comando Dati > Da Web tolti dall'interfaccia (il contenuto già presente nei file resta conservato al salvataggio).
- [x] Pulsante ingrandisci: un quadrato a finestra, due quadrati a schermo intero, aggiornato a ogni cambio di stato (test UI Automation). Tooltip "Ingrandisci" / "Ripristina giù".
- [x] Ridimensionamento: barra del titolo (ricerca e titolo si stringono), schede in più nel menu «⋯», gruppi del ribbon compattati a sole icone e poi nel menu «»»; larghezza minima 640 (screenshot a massimizzata, 1100, 800 e minima).
- [x] Powers > Scudo privacy: analizza il file salvato (autore, ultimo autore, società, responsabile, proprietà personalizzate, autori di commenti/persone/revisioni, percorsi con nome utente, miniatura), crea una copia `nome (pulito).xlsx` (anonimizza o rimuove commenti), ricontrolla la copia ed elenca le riserve. Non tocca mai il file aperto.
