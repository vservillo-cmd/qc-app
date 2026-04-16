QC TABLET PILOT - PACCHETTO COMPLETO

Contenuto:
- index.html
- style.css
- app.js
- README.txt

Uso:
1. Estrai tutti i file in una cartella.
2. Dalla cartella avvia:
   py -m http.server 8000
3. Apri:
   http://localhost:8000
4. Login demo:
   qualita1 / 1234

Import anagrafiche:
- pulsante "Importa Anagrafiche"
- legge i fogli: Postazioni, Macchiniste, Codici_Prodotto
- ricava le Linee dal foglio Postazioni (colonna Linea)
- se nel file esiste un foglio Misure/Tolleranze, applica l'OK/KO automatico sulle misure

Regole attive:
- Tipo cucitura automatico:
  ML -> CU
  MCD -> IM
  TC -> SORF
- se serve un tipo diverso, usa il campo sotto "Nuovo tipo cucitura"
- per Postazione, Linea e Codice prodotto puoi usare il valore da elenco oppure inserire un nuovo valore sotto
- per privacy nello storico viene mostrato il badge macchinista

Importazione in Excel:
1. Apri Excel
2. Vai su Dati > Da testo/CSV
3. Seleziona qc_controlli.csv
4. Se Excel non separa bene le colonne, imposta delimitatore ;
5. Carica


Aggiornamenti v4:
- Particolare in produzione popolato in automatico dal foglio Postazioni:
  ID_Linea -> valore colonna Linea
- Campo manuale sotto per nuovo particolare
- Esiti misura OK/KO non più selezionabili a mano: sono campi automatici in sola lettura


Aggiornamenti v5:
- Particolare in produzione spostato sulla stessa riga di Codice prodotto
- Particolare in produzione popolato automaticamente da ID_Linea -> colonna Linea
- Esiti misura OK/KO ora visibili in campi readonly con evidenza colore


Aggiornamenti v6:
- Linea lavorazione letta dal foglio "Elenco Linee":
  ID_Linea Excel -> Linea lavorazione software
  Linea Excel -> Particolare in produzione software
- Particolare in produzione aggiornato automaticamente alla selezione della linea
- Esiti misura OK/KO aggiornati in tempo reale e visibili nei campi readonly


Aggiornamenti v8:
- prime due righe rimesse in ordine senza cambiare la struttura del form
- selezionando la Postazione, la Linea lavorazione viene proposta automaticamente dalla tabella Postazioni
- il Particolare in produzione usa la mappatura del foglio Elenco Linee:
  ID_Linea excel -> Linea lavorazione software
  Linea excel -> Particolare in produzione software

Nota importante:
- l'OK/KO automatico delle misure funziona solo se nel file importato è presente un foglio Misure/Tolleranze con i limiti.


Aggiornamenti v9:
- lettura tolleranze allineata alle intestazioni reali del foglio Misure
- filtro Codice prodotto per ID_Linea
- i codici prodotto vengono aggiornati in base alla linea selezionata/proposta dalla postazione


Aggiornamenti v10:
- login reso tollerante agli errori di inizializzazione
- se una parte dell'avvio fallisce, il pulsante Entra continua a funzionare


Aggiornamenti v11:
- finestra attestazione QR resa scrollabile e contenuta nei limiti visibili
- prima riga del form compattata per mostrare su una sola riga:
  Turno, Postazione, Tipo cucitura, Linea lavorazione, Lotto, Codice prodotto, Particolare in produzione
- seconda riga dedicata solo ai campi "nuovo ..."


Aggiornamenti v12:
- prima riga del form ricostruita davvero con griglia dedicata
- Turno, Postazione, Tipo cucitura, Linea lavorazione, Lotto, Codice prodotto e Particolare in produzione ora stanno sulla stessa riga su schermi desktop ampi
- seconda riga separata per i campi 'nuovo ...'
