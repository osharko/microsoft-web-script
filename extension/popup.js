function injectScript(file) {
  const status = document.getElementById('status');

  chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
    const tab = tabs[0];
    if (!tab) {
      status.textContent = 'Nessun tab attivo trovato.';
      return;
    }

    const url = tab.url || '';
    if (file === 'download_email.js' && !url.includes('outlook.')) {
      status.textContent = 'Apri Outlook Web prima.';
      return;
    }
    if (file === 'download_transcript.js' && !url.includes('microsoft') && !url.includes('sharepoint')) {
      status.textContent = 'Apri un video Stream prima.';
      return;
    }

    chrome.scripting.executeScript({
      target: { tabId: tab.id },
      files: [file]
    }, () => {
      if (chrome.runtime.lastError) {
        status.textContent = 'Errore: ' + chrome.runtime.lastError.message;
      } else {
        status.textContent = 'Script iniettato!';
        setTimeout(() => window.close(), 800);
      }
    });
  });
}

document.getElementById('btnEmail').addEventListener('click', () => injectScript('download_email.js'));
document.getElementById('btnTranscript').addEventListener('click', () => injectScript('download_transcript.js'));
