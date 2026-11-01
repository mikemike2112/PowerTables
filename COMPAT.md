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

Il clone non esporta PDF (la stampa è P1), quindi il confronto PDF previsto dal brief **non è stato eseguito**. Osservazione visiva diretta: l'app mostra celle,
formati, unioni, riquadri e filtri del corpus; i colori di sfondo/testo espliciti sono rispettati, il testo nero su sfondo non riempito viene invertito nel tema scuro (scelta di leggibilità).

## Limiti noti
Vedere `FEATURES.md` per le voci non spuntate.
