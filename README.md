# Microsoft Web Scripts

Script da eseguire direttamente nella console del browser per scaricare email e trascrizioni dai servizi Microsoft.

## Download Email da Outlook Web

Scarica tutte le email di una cartella Outlook in formato `.eml`, compresse in un unico file `.zip`.

**Funzionalita:**
- Selezione dinamica della cartella tramite menu a tendina (incluse sottocartelle)
- Slider per configurare quante email includere per ZIP (default: 100, max: 5000)
- Barra di progresso in tempo reale
- Nessuna dipendenza esterna (ZIP generato in puro JS)

**Utilizzo:** apri [Outlook Web](https://outlook.cloud.microsoft), premi `F12` → Console, e incolla:

```js
fetch('https://raw.githubusercontent.com/osharko/microsoft-web-script/main/download_email.js').then(r=>r.text()).then(eval)
```

## Download Trascrizioni da Microsoft Stream

Estrae la trascrizione completa da un video Microsoft Stream (SharePoint) e la scarica come file `.txt`.

**Utilizzo:** apri il video su Microsoft Stream, apri il pannello **Transcript**, premi `F12` → Console, e incolla:

```js
fetch('https://raw.githubusercontent.com/osharko/microsoft-web-script/main/download_transcript.js').then(r=>r.text()).then(eval)
```
