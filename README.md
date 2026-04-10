# Microsoft Web Scripts

Script da eseguire nella console del browser per scaricare email e trascrizioni dai servizi Microsoft.

> **Nota:** le pagine Microsoft hanno Content Security Policy restrittive che bloccano `fetch` verso domini esterni. Per questo motivo gli script vanno copiati e incollati direttamente nella console, non caricati via fetch.

## Download Email da Outlook Web

Scarica tutte le email di una cartella Outlook in formato `.eml`, compresse in file `.zip`.

**Funzionalita:**
- Selezione dinamica della cartella tramite menu a tendina (incluse sottocartelle)
- Slider per configurare quante email per ZIP (default: 100, max: 5000)
- Preflight: mostra quanti file ZIP verranno generati prima di iniziare
- Se le email stanno in un singolo ZIP, viene scaricato un file unico; altrimenti vengono creati piu file numerati
- Barra di progresso in tempo reale
- Nessuna dipendenza esterna (ZIP generato in puro JS)

**Utilizzo:**

1. Apri [Outlook Web](https://outlook.cloud.microsoft)
2. Premi `F12` → Console
3. Copia il contenuto di [`download_email.js`](download_email.js) e incollalo nella console
4. Premi Invio - apparira una modale per scegliere cartella e dimensione ZIP

## Download Trascrizioni da Microsoft Stream

Estrae la trascrizione completa da un video Microsoft Stream (SharePoint) e la scarica come file `.txt`.

**Utilizzo:**

1. Apri il video su Microsoft Stream e apri il pannello **Transcript**
2. Premi `F12` → Console
3. Copia il contenuto di [`download_transcript.js`](download_transcript.js) e incollalo nella console
4. Premi Invio - il download partira automaticamente
