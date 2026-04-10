// === Script per estrarre la trascrizione da Microsoft Stream (SharePoint) ===
// 1. Apri il pannello "Transcript" sul video
// 2. Apri la Console del browser (F12 → Console)
// 3. Incolla ed esegui questo script

(async function downloadTranscript() {
  // Trova il pannello trascrizione
  const transcriptPanel = document.querySelector('[role="complementary"][aria-label="Transcript"]');
  if (!transcriptPanel) {
    alert('Pannello Transcript non trovato! Assicurati che sia aperto.');
    return;
  }

  // Trova il container scrollabile (la lista è virtualizzata)
  const divs = transcriptPanel.querySelectorAll('div');
  const scrollable = Array.from(divs).find(d => d.scrollHeight > d.clientHeight + 50);
  if (!scrollable) {
    alert('Container scrollabile non trovato.');
    return;
  }

  console.log('⏳ Raccolta trascrizione in corso...');

  // Scroll graduale per caricare tutti gli elementi virtualizzati
  const allEntries = new Map();
  let lastScrollTop = -1;
  scrollable.scrollTop = 0;

  const collectVisible = () => {
    const listItems = Array.from(document.querySelectorAll('[role="listitem"]'));
    let currentAuthor = '';
    let currentTimestamp = '';

    for (const item of listItems) {
      const nameEl = item.querySelector('[class*="itemDisplayName"]');
      const timestampEl = item.querySelector('[id^="Header-timestamp"]');

      if (nameEl) {
        currentAuthor = nameEl.textContent.trim();
        currentTimestamp = timestampEl ? timestampEl.textContent.trim() : '';
      } else {
        const text = item.textContent.trim();
        if (text && currentAuthor) {
          const key = currentTimestamp + '|' + currentAuthor + '|' + text.substring(0, 50);
          if (!allEntries.has(key)) {
            allEntries.set(key, { author: currentAuthor, timestamp: currentTimestamp, text });
          }
        }
      }
    }
  };

  for (let i = 0; i < 200; i++) {
    collectVisible();
    scrollable.scrollTop += 500;
    await new Promise(r => setTimeout(r, 150));
    if (scrollable.scrollTop === lastScrollTop) break;
    lastScrollTop = scrollable.scrollTop;
  }
  collectVisible();

  // Torna in cima
  scrollable.scrollTop = 0;

  // Formatta e scarica
  const transcript = Array.from(allEntries.values());
  const formatted = transcript.map(e => `[${e.timestamp}] ${e.author}: ${e.text}`).join('\n');

  const blob = new Blob([formatted], { type: 'text/plain;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `trascrizione_${document.title.replace(/[^a-zA-Z0-9]/g, '_').substring(0, 60)}.txt`;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);

  console.log(`✅ Download completato! ${transcript.length} interventi estratti.`);
})();