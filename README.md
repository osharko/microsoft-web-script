# Microsoft Web Scripts

Script per scaricare email e trascrizioni dai servizi Microsoft. Disponibili come **estensione Chrome** o da **console del browser**.

## Estensione Chrome (consigliato)

Evita di usare la console del browser: l'estensione inietta gli script con un click.

**Installazione:**

1. Clona il repo o scarica lo ZIP
2. Apri `chrome://extensions/` e attiva la **Modalita sviluppatore**
3. Clicca **Carica estensione non pacchettizzata** e seleziona la cartella `extension/`
4. L'icona apparira nella toolbar di Chrome

**Utilizzo:** clicca l'icona dell'estensione e scegli lo script da eseguire.

## Utilizzo da Console

> **Nota:** le pagine Microsoft hanno CSP restrittive che bloccano `fetch` verso domini esterni. Gli script vanno copiati e incollati nella console.

### Download Email da Outlook Web

Scarica tutte le email di una cartella Outlook in formato `.eml` compresse in `.zip`.

- Selezione dinamica della cartella tramite menu a tendina (incluse sottocartelle)
- Slider per configurare quante email per ZIP (default: 100, max: totale email della cartella)
- Preflight: mostra quanti file ZIP verranno generati prima di iniziare
- Barra di progresso in tempo reale
- Nessuna dipendenza esterna (ZIP generato in puro JS)

1. Apri [Outlook Web](https://outlook.cloud.microsoft)
2. Premi `F12` → Console
3. Copia il contenuto di [`download_email.js`](download_email.js) e incollalo nella console
4. Premi Invio - apparira una modale per scegliere cartella e dimensione ZIP

### Download Trascrizioni da Microsoft Stream

Estrae la trascrizione completa da un video Microsoft Stream (SharePoint) come file `.txt`.

1. Apri il video su Microsoft Stream e apri il pannello **Transcript**
2. Premi `F12` → Console
3. Copia il contenuto di [`download_transcript.js`](download_transcript.js) e incollalo nella console
4. Premi Invio - il download partira automaticamente
