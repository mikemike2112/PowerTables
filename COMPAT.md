# Compatibilità e misure reali

Tutti i numeri vengono da esecuzioni reali su questo PC (Windows 11, .NET 10, Excel desktop in italiano) il 01/11/2026.

## Round-trip OOXML (Excel vero via COM)

Corpus: **42 file** prodotti da Excel (`tests/corpus/office-made`, script `tools/make-corpus.ps1`), un file per funzionalità più file misti
(tipi/formati, date 1900 e 1904, 25 file di formule, formule matriciali CSE e dinamiche, nomi, celle unite, riquadri bloccati, filtri, tabella,
grafico, formattazione condizionale, convalida, commenti, collegamenti, protezione, 1000 righe, report e budget misti). Author/LastAuthor = "Corpus" (verificato da test).

Ultima esecuzione completa di `dotnet test`: 328 superati, 0 falliti, 0 saltati (1 min 38 s); i test per ogni file del corpus sono:

| Controllo | Esito |
|---|---|
| Excel apre l'originale senza errori | 42/42 |
| apri + salva senza modifiche, Excel riapre senza errore né ripristino | 42/42 (100%) |
| `OpenXmlValidator`: nessun errore nuovo rispetto all'originale (salvato e modificato) | 42/42 (nessun errore nuovo; non serve nessuna eccezione) |
| ogni parte XML identica (forma canonica) e stesso elenco part/relazioni/tipi di contenuto (senza modifiche) | 42/42 |
| contenuto identico letto da Excel (valori, formule, formati, unioni, larghezze, altezze, riquadri, oggetti) | 42/42 |
| con modifica (nuove celle testo/numero/formula): modifica presente in Excel, resto intatto | 42/42 |
| controllo negativo: il rilevatore fallisce su 5 file corrotti (XML troncato, parte mancante, indice sst fuori range, relazione rotta, formula non valida) | 5/5 colti |

Note sul round-trip:
- Un salvataggio senza modifiche non tocca nessuna parte foglio: cambiano solo la serializzazione di `[Content_Types].xml` e dei `.rels` fatta dall'SDK
  (Default/Override riorganizzati, target relativi → assoluti): significato identico, verificato risolvendo tipi e target.
- Salvataggio con modifica: si riscrivono solo le righe toccate; le altre righe del foglio sono copiate byte per byte; `<cols>`, unioni, riquadri e filtri sono rigenerati solo se modificati.
  Se una formula è cambiata viene eliminata `calcChain.xml` (Excel la ricostruisce). Misurato con Excel vero (`format-alignment-wrap.xlsx`): la rimozione di `calcChain` non c'entra, anche una modifica di solo valore in una riga nuova fa rimisurare a Excel le righe con testo a capo senza `customHeight` (altezza 100,8 salvata da Excel con colonna ancora larga → 57,6 reale con larghezza 14,77); le righe con altezza personalizzata restano invariate. Non altera il contenuto.
- Formule matriciali (CSE), tabelle dati e formule dinamiche (Formula2/array dinamici, `cm=`) sono conservate come XML originale con il valore in cache e non vengono ricalcolate.
  Le celle "spill" sono valori normali. Formule con sintassi non gestita (es. riferimenti strutturati `Tabella[[#This Row],[Col]]`) o funzioni non implementate conservano formula e valore in cache.
- Righe riscritte: le stringhe diventano `inlineStr`; il testo ricco (runs formattati) delle celle il cui testo non cambia viene riscritto identico dentro `<is>` (test `RichText_SurvivesEditOfAnotherCellInTheSameRow` su Excel vero: formattazione di primo e ultimo carattere uguale prima/dopo; senza la correzione il test fallisce). Limite: se una cella a testo ricco viene spostata (es. ordinamento) la formattazione dei runs si perde; attributi di cella sconosciuti diversi da `cm`/`vm` nelle righe riscritte non sono preservati.
- Salvataggio atomico (file temporaneo + `File.Replace`).

## Motore formule contro Excel vero

`tests/PowerTables.Tests/engine-cases.txt`: **393 formule** (incluse date 1900, bug 29/02/1900, errori, coercizioni, jolly, casi limite) scritte in un `.xlsx`,
calcolate da Excel (`Value2`/`Formula` invarianti) e confrontate con il motore: **393/393 uguali** all'ultima esecuzione (1 con differenza solo di locale, vedi sotto).
Le **173 funzioni** registrate sono quelle sotto; **149** hanno almeno un caso confrontato con Excel, le altre 24 (RAND, TODAY, NOW, TRANSPOSE, DOLLAR, FIXED, SHEET, FORMULATEXT, ISREF, ISFORMULA, alias STDEV.S/VAR.S/...) sono coperte solo da test interni.
Inoltre `CorpusEngineTests` ricalcola tutti i 42 file del corpus e ritrova i valori in cache scritti da Excel.

Differenze di locale note (Excel italiano): conversione numero→testo e `TEXT` usano il separatore decimale/migliaia del sistema (`1,5`, `VERO`), il motore usa sempre
la sintassi invariante (`1.5`, `TRUE`); `NUMBERVALUE` senza argomenti, `DATEVALUE` con date locali. I test evitano questi casi o li normalizzano.
Le formule si digitano nella UI in sintassi inglese con la virgola (non nomi di funzione localizzati).

Limiti del motore: niente intersezione/unione di riferimenti (spazio, virgola), riferimenti 3D, riferimenti strutturati; sistema di date 1904 letto/scritto ma le funzioni data usano sempre il 1900;
`MATCH`/`VLOOKUP` approssimati su dati non ordinati usano una ricerca binaria sui soli elementi dello stesso tipo (Excel su dati non ordinati ha comportamento non definito);
cicli: le celle del ciclo valgono 0 e sono elencate (`Workbook.Cycles`), calcolo iterativo non supportato; nomi definiti solo globali/di foglio letti dal file ma non scritti se nuovi.

## Test di carico: 1.000.000 di righe popolate (build Release, misurato dall'app con `--bench`)

File: 1.000.001 righe × 6 colonne (ID, testo, importo formattato, data, 2 formule per riga) + 3 formule di riepilogo: **6.000.012 celle, 2.000.003 formule**, 56,7 MB su disco.

| Misura | Valore |
|---|---|
| Apertura (lettura streaming + store sparso) | 7,7 s |
| Memoria gestita dopo l'apertura | 1.263 MB (working set 2.423 MB) |
| Scroll: 100 salti casuali su tutto il foglio, render completo | media 31,8 ms, max 119 ms |
| Scroll: 300 passi sequenziali, render completo | media 11,2 ms, max 26,3 ms |
| Ricalcolo completo (costruzione del grafo + 2.000.003 formule) | 6,8 s |
| Ricalcolo incrementale dopo la modifica di una cella | 251 ms (2 formule; domina una `VLOOKUP` esatta su 1M righe) |
| Generazione + ricalcolo iniziale + salvataggio (Core, Debug) | 4,3 s + 15,5 s + 22,4 s |

Il grafo delle dipendenze viene costruito alla prima modifica o al ricalcolo (non all'apertura). La memoria è alta (~210 byte/cella): le stringhe di formula non sono condivise.

## Confronto visivo

L'esportazione PDF esiste (`PdfExport`, PDFsharp 6.2.4 MIT) ma il confronto PDF con Excel previsto dal brief **non è stato eseguito**. Osservazione visiva diretta: l'app mostra celle,
formati, unioni, riquadri e filtri del corpus; i colori di sfondo/testo espliciti sono rispettati, il testo nero su sfondo non riempito viene invertito nel tema scuro (scelta di leggibilità).

## Limiti noti
Vedere `FEATURES.md` per le voci non spuntate.

## Scudo privacy (misurato)
- Test con file dai dati noti (autore, ultimo autore, società, responsabile, proprietà personalizzata, autore di commento, miniatura, percorso `C:\Users\...` in `xl/workbook.xml`): la copia, scansionata parte per parte, non contiene più nessuno dei valori; l'originale resta identico byte per byte (SHA-256); Excel vero apre la copia senza ripristino con celle e numero di commenti uguali (modalità anonimizza) o senza commenti (modalità rimuovi); un file già pulito dà "niente trovato".
- Riserve sempre elencate: EXIF delle immagini, oggetti incorporati, macro, righe/colonne e fogli nascosti, collegamenti esterni a percorsi locali, testo delle celle e dei commenti. Non analizzato: contenuto delle celle.
- Osservato: Excel scrive nei file un `x15ac:absPath` con il percorso della cartella di salvataggio (nel corpus: `C:\dev\powertables\tests\corpus\office-made\`): non contiene il nome utente, ma lo Scudo lo segnala se contiene `\Users\nome`.

## Inserimento/eliminazione righe e colonne
- Aggiornati: celle, formule di tutti i fogli (incluse quelle su altri fogli), unioni, altezze/larghezze, filtro automatico, testo ricco, nomi definiti (anche `_xlnm.*`), `sqref` e formule di formattazione condizionale e convalida dati, collegamenti. Righe/colonne eliminate: le regole interamente dentro scompaiono, le altre si accorciano. Verificato sul Core con il validatore Open XML (`LinesFindTests`), non confrontato con Excel vero.
- **Non aggiornati** (restano ai vecchi indirizzi, nessun errore di struttura): grafici, pivot, commenti, formule matriciali/opache, CF/convalida con estensione `x14` (`extLst`), regole di altri fogli che puntano a questo. Chi li usa deve ricontrollarli dopo l'operazione.
- Formattazione condizionale, convalida dati e collegamenti sono ora nel modello del foglio (`Sheet.CondFmt/DataVal/Links`) e seguono inserimenti/eliminazioni, anche con Annulla dopo il salvataggio. Le tabelle (`tableParts`) seguono righe/colonne e i nomi delle colonne seguono le intestazioni.

## Giro «comandi inerti»
Verificato sul Core con il validatore Open XML (nessun errore nuovo) e con test automatici; **non confrontato con Excel vero**.
- **Copia formato / Rimuovi duplicati / Testo in colonne**: annullabili con un Ctrl+Z; Rimuovi duplicati confronta senza distinzione di maiuscole e sposta le formule delle righe rimaste; Testo in colonne interpreta i valori come digitati (delimitatori, qualificatore `"`).
- **Nomi** (Definisci nome, Gestione nomi): validazione come Excel (non coincidono con riferimenti di cella), ambito cartella o foglio, `definedNames` riscritto al salvataggio conservando i nomi riservati; rinomina/spostamento/eliminazione fogli rimappano i nomi locali e `Print_Area`.
- **Convalida dati**: numero intero, decimale, lunghezza testo, elenco (voci o intervallo); il valore digitato viene rifiutato con Riprova/Annulla; freccia dell'elenco nella cella attiva. Le formule digitate non sono controllate. Convalide con estensioni `x14`/`mc:AlternateContent` sono conservate ma non modificate.
- **Formattazione condizionale**: regole `cellIs` (maggiore, minore, uguale, diverso, tra...) e `containsText` con colori carattere/riempimento/grassetto (`dxf`), disegnate nella griglia e nel PDF. Altri tipi di regola (barre, scale colori, set di icone, formule) sono conservati e non disegnati.
- **Collegamento**: URL esterni solo `http`, `https`, `ftp`, `mailto` (qualunque altro schema è rifiutato), posizioni interne `Foglio!A1`; Ctrl+clic apre.
- **Formato celle** (Ctrl+1): codice numerico, allineamento, rientro, carattere, colori (RRGGBB), un solo Annulla. **Stili cella**: preimpostazioni applicate come formato (non stili con nome del file).
- **Imposta pagina**: margini, orientamento, formato carta (A4, A3, A5, Lettera, Legale), area di stampa, griglia in stampa, salvati in `pageMargins`/`pageSetup`/`printOptions`/`_xlnm.Print_Area`. Non gestiti: adattamento a pagine, intestazioni e piè di pagina, ripeti titoli, interruzioni manuali.
- **Esporta PDF / Stampa**: area di stampa o area usata, orientamento, carta, margini, divisione in pagine per righe e colonne (prima in basso, poi a destra), riempimenti, bordi, caratteri, allineamenti, formattazione condizionale e stile tabella. Caratteri dalla cartella Fonts di sistema (Calibri, Arial, Times, Courier, Verdana, Georgia, Segoe UI, Consolas; altri → Arial). Stampa apre il PDF nel programma predefinito. Righe nascoste, unioni e rotazione del testo: unioni sì, rotazione no.
- **Tabelle**: creazione (con o senza intestazioni; intestazioni vuote/duplicate rinominate «ColonnaN»), `table*.xml` scritto con stile, aggiornamento area/nomi colonna dopo inserimenti e modifiche delle intestazioni, conversione in intervallo. Gli stili sono 5 (Medio 2, 3, 4, 5, 7); stili di tabella non gestiti sono conservati e disegnati solo da Excel. Niente totali, filtri di tabella né riferimenti strutturati nelle formule.
- **Fogli**: nuovo, elimina (riferimenti → `#REF!`), sposta, rinomina: tutti annullabili. Un foglio eliminato e poi ripristinato con Annulla dopo un salvataggio è riscritto da zero (celle e layout, non gli oggetti del file originale).

## Giro «ultimi pulsanti inerti»
Verificato con il validatore Open XML e con Excel vero via COM (apre senza ripristino: `ObjectsTests.ExcelOpensObjectsAndCommentsWithoutRepair`, `PivotTests`); i test con input reale prendono `ui.lock`.
- **Immagini, Forme, Casella di testo, Grafici**: oggetti in `xl/drawings` (`twoCellAnchor`), nome `PT-xxxxxxxx` per ritrovarli; grafico `col` o `line` con cache dei valori; descrizione dell'oggetto (JSON) nell'attributo `descr` per rileggerlo. Quando si riapre il file gli oggetti creati da PowerTables tornano visibili; **gli oggetti creati da altri programmi restano intatti nel file ma non sono disegnati né elencati**. Annulla dopo il salvataggio toglie anche le parti (immagine/grafico) dal file. Non si spostano né ridimensionano col mouse; ancoraggio alla cella attiva; il PDF **non** contiene questi oggetti.
- **Commento**: `comments*.xml` e VML scritti/aggiornati; i commenti già nel file sono letti come testo semplice, quelli modificati perdono la formattazione interna; i commenti a catena (`threadedComments`) non sono toccati.
- **Tabella pivot minima**: un campo riga, un campo valori (somma o conteggio), nuovo foglio o cella; le celle dell'output sono scritte come valori, `pivotTable`/`pivotCacheDefinition` con `refreshOnLoad` e nessun record (Excel ricalcola all'apertura). Non rilette: dopo la riapertura restano nel file ma l'app non le conosce. Le parti della cache stanno in `/pivotCache/` (percorso scelto dall'SDK, accettato da Excel).
- **Filtro avanzato**: criteri come Excel (testo = «inizia con», `=testo` esatto, `*` `?`, `>` `<` `>=` `<=` `<>`, righe = O, colonne = E); sul posto nasconde le righe (non crea `_FilterDatabase`).
- **Ricerca obiettivo**: secante poi bisezione, tolleranza 1e-7; nessuna soluzione = nessuna modifica. **Foglio previsione**: regressione lineare con limiti al 95% (t approssimato), **non** il modello ETS di Excel.
- **Layout di pagina / Anteprima interruzioni**: visualizzazioni di controllo, non l'editor di Excel. **Intestazioni**: nascoste solo le etichette. **Griglia e Intestazioni** ora sono salvate in `showGridLines`/`showRowColHeaders`.
- **Temi**: cambia il tema dell'interfaccia di PowerTables, non il tema di colori del file (la tavolozza dei colori tema è fissa Office).
- **Scheda Automatizza** (Nuovo script, Visualizza script): tolta dall'interfaccia, niente esecuzione di contenuto attivo. Gli script presenti in un file restano conservati al salvataggio.
