// ============================================================
// Outlook Web Email Downloader (.eml -> .zip)
// ============================================================
// Apri Outlook Web (outlook.cloud.microsoft), premi F12 -> Console
// Incolla tutto questo script e premi Invio.
// Apparira' una modale per scegliere la cartella e la dimensione ZIP.
// ============================================================

// --- MiniZip: creatore ZIP puro JS (metodo STORE, no dipendenze) ---
class MiniZip {
  constructor() { this.files = []; }

  addFile(name, data) {
    this.files.push({ name, data });
  }

  generate() {
    const localHeaders = [];
    const centralHeaders = [];
    let offset = 0;

    for (const file of this.files) {
      const nameBytes = new TextEncoder().encode(file.name);
      const data = file.data;
      const crc = this._crc32(data);

      const local = new ArrayBuffer(30 + nameBytes.length + data.length);
      const lv = new DataView(local);
      const lu = new Uint8Array(local);
      lv.setUint32(0, 0x04034b50, true);
      lv.setUint16(4, 20, true);
      lv.setUint16(6, 0, true);
      lv.setUint16(8, 0, true);
      lv.setUint16(10, 0, true);
      lv.setUint16(12, 0, true);
      lv.setUint32(14, crc, true);
      lv.setUint32(18, data.length, true);
      lv.setUint32(22, data.length, true);
      lv.setUint16(26, nameBytes.length, true);
      lv.setUint16(28, 0, true);
      lu.set(nameBytes, 30);
      lu.set(data, 30 + nameBytes.length);
      localHeaders.push(lu);

      const central = new ArrayBuffer(46 + nameBytes.length);
      const cv = new DataView(central);
      const cu = new Uint8Array(central);
      cv.setUint32(0, 0x02014b50, true);
      cv.setUint16(4, 20, true);
      cv.setUint16(6, 20, true);
      cv.setUint16(8, 0, true);
      cv.setUint16(10, 0, true);
      cv.setUint16(12, 0, true);
      cv.setUint16(14, 0, true);
      cv.setUint32(16, crc, true);
      cv.setUint32(20, data.length, true);
      cv.setUint32(24, data.length, true);
      cv.setUint16(28, nameBytes.length, true);
      cv.setUint16(30, 0, true);
      cv.setUint16(32, 0, true);
      cv.setUint16(34, 0, true);
      cv.setUint16(36, 0, true);
      cv.setUint32(38, 0, true);
      cv.setUint32(42, offset, true);
      cu.set(nameBytes, 46);
      centralHeaders.push(cu);
      offset += local.byteLength;
    }

    const centralSize = centralHeaders.reduce((s, h) => s + h.length, 0);
    const eocd = new ArrayBuffer(22);
    const ev = new DataView(eocd);
    ev.setUint32(0, 0x06054b50, true);
    ev.setUint16(4, 0, true);
    ev.setUint16(6, 0, true);
    ev.setUint16(8, this.files.length, true);
    ev.setUint16(10, this.files.length, true);
    ev.setUint32(12, centralSize, true);
    ev.setUint32(16, offset, true);
    ev.setUint16(20, 0, true);

    const result = new Uint8Array(offset + centralSize + 22);
    let pos = 0;
    for (const lh of localHeaders) { result.set(lh, pos); pos += lh.length; }
    for (const ch of centralHeaders) { result.set(ch, pos); pos += ch.length; }
    result.set(new Uint8Array(eocd), pos);
    return result;
  }

  _crc32(data) {
    if (!MiniZip._crcTable) {
      const t = new Uint32Array(256);
      for (let i = 0; i < 256; i++) {
        let c = i;
        for (let j = 0; j < 8; j++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
        t[i] = c;
      }
      MiniZip._crcTable = t;
    }
    let crc = 0xFFFFFFFF;
    for (let i = 0; i < data.length; i++)
      crc = MiniZip._crcTable[(crc ^ data[i]) & 0xFF] ^ (crc >>> 8);
    return (crc ^ 0xFFFFFFFF) >>> 0;
  }
}

// --- Funzione per ottenere il token dalla sessione OWA ---
function getOutlookToken() {
  const allKeys = Object.keys(localStorage);
  const tokenKey = allKeys.find(k =>
    k.includes('accesstoken') &&
    k.includes('outlook.office') &&
    k.includes('analytics.readwrite')
  );
  if (!tokenKey) throw new Error('Token non trovato! Assicurati di essere loggato su Outlook Web.');
  return JSON.parse(localStorage.getItem(tokenKey)).secret;
}

// --- Carica le cartelle dall'account ---
async function loadFolders(token) {
  const headers = { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' };
  const folders = [];

  async function fetchChildren(parentId, prefix) {
    const url = parentId
      ? `https://outlook.office.com/api/v2.0/me/mailfolders/${parentId}/childfolders?$select=Id,DisplayName,TotalItemCount&$top=100`
      : `https://outlook.office.com/api/v2.0/me/mailfolders?$select=Id,DisplayName,TotalItemCount&$top=100`;
    const resp = await fetch(url, { headers });
    if (!resp.ok) return;
    const data = await resp.json();
    for (const f of (data.value || [])) {
      const displayPath = prefix ? `${prefix} / ${f.DisplayName}` : f.DisplayName;
      folders.push({ id: f.Id, name: f.DisplayName, path: displayPath, count: f.TotalItemCount });
      await fetchChildren(f.Id, displayPath);
    }
  }

  await fetchChildren(null, '');
  return folders;
}

// --- Mostra la modale di configurazione ---
function showConfigModal(folders) {
  return new Promise((resolve, reject) => {
    const overlay = document.createElement('div');
    overlay.id = 'emlModalOverlay';
    overlay.style.cssText = 'position:fixed;top:0;left:0;width:100%;height:100%;' +
      'background:rgba(0,0,0,0.6);z-index:999998;display:flex;align-items:center;justify-content:center;';

    const modal = document.createElement('div');
    modal.style.cssText = 'background:#1a1a2e;color:#e0e0e0;padding:30px;border-radius:12px;' +
      'min-width:420px;max-width:500px;font-family:monospace;font-size:14px;' +
      'box-shadow:0 8px 32px rgba(0,0,0,0.5);border:1px solid #00ff88;';

    const DEFAULT_EMAILS_PER_ZIP = 100;

    const calcPreflight = (totalEmails, perZip) => {
      const zipCount = Math.ceil(totalEmails / perZip);
      if (zipCount <= 1) return `<span style="color:#00ff88;">1 file ZIP</span>`;
      return `<span style="color:#ffaa00;">${zipCount} file ZIP</span>`;
    };

    const initMax = Math.max(folders[0]?.count || 100, 50);

    modal.innerHTML = `
      <h2 style="margin:0 0 20px;color:#00ff88;font-size:18px;">Outlook Email Downloader</h2>

      <label style="display:block;margin-bottom:6px;color:#aaa;">Cartella:</label>
      <select id="emlFolderSelect" style="width:100%;padding:8px;border-radius:6px;
        background:#0d0d1a;color:#fff;border:1px solid #333;font-family:monospace;font-size:13px;
        margin-bottom:20px;">
        ${folders.map((f, i) => `<option value="${i}">${f.path} (${f.count})</option>`).join('')}
      </select>

      <label style="display:block;margin-bottom:6px;color:#aaa;">
        Email per ZIP: <span id="emlSliderValue" style="color:#00ff88;">${DEFAULT_EMAILS_PER_ZIP}</span>
      </label>
      <input type="range" id="emlSlider" min="50" max="${initMax}" step="50" value="${Math.min(DEFAULT_EMAILS_PER_ZIP, initMax)}"
        style="width:100%;margin-bottom:4px;accent-color:#00ff88;">
      <div style="display:flex;justify-content:space-between;color:#666;font-size:11px;margin-bottom:16px;">
        <span>50</span><span id="emlSliderMax">${initMax}</span>
      </div>

      <div id="emlPreflight" style="background:#0d0d1a;padding:12px;border-radius:8px;
        margin-bottom:20px;border:1px solid #333;text-align:center;font-size:13px;">
      </div>

      <div style="display:flex;gap:10px;justify-content:flex-end;">
        <button id="emlCancel" style="padding:8px 20px;border-radius:6px;border:1px solid #666;
          background:transparent;color:#aaa;cursor:pointer;font-family:monospace;">Annulla</button>
        <button id="emlStart" style="padding:8px 20px;border-radius:6px;border:none;
          background:#00ff88;color:#0d0d1a;cursor:pointer;font-weight:bold;font-family:monospace;">Scarica</button>
      </div>
    `;

    overlay.appendChild(modal);
    document.body.appendChild(overlay);

    const slider = document.getElementById('emlSlider');
    const sliderValue = document.getElementById('emlSliderValue');
    const folderSelect = document.getElementById('emlFolderSelect');
    const preflightDiv = document.getElementById('emlPreflight');

    const updatePreflight = () => {
      const folder = folders[folderSelect.value];
      const perZip = parseInt(slider.value);
      const total = folder.count;
      const zipCount = Math.ceil(total / perZip);
      preflightDiv.innerHTML = `${total} email &rarr; ${calcPreflight(total, perZip)}`;
    };

    const sliderMaxLabel = document.getElementById('emlSliderMax');

    const updateSliderMax = () => {
      const folder = folders[folderSelect.value];
      const newMax = Math.max(folder.count, 50);
      slider.max = newMax;
      sliderMaxLabel.textContent = newMax;
      if (parseInt(slider.value) > newMax) {
        slider.value = newMax;
        sliderValue.textContent = newMax;
      }
    };

    slider.addEventListener('input', () => {
      sliderValue.textContent = slider.value;
      updatePreflight();
    });
    folderSelect.addEventListener('change', () => {
      updateSliderMax();
      updatePreflight();
    });
    updatePreflight();

    document.getElementById('emlCancel').addEventListener('click', () => {
      overlay.remove();
      reject(new Error('Annullato dall\'utente'));
    });

    document.getElementById('emlStart').addEventListener('click', () => {
      const selectedFolder = folders[document.getElementById('emlFolderSelect').value];
      const emailsPerZip = parseInt(slider.value);
      overlay.remove();
      resolve({ folder: selectedFolder, emailsPerZip });
    });
  });
}

// --- Script principale ---
(async () => {
  const API_BATCH_SIZE = 50;
  const DELAY_MS = 100;

  const token = getOutlookToken();
  const headers = { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' };

  // Caricamento cartelle con indicatore
  let loadingDiv = document.createElement('div');
  loadingDiv.style.cssText = 'position:fixed;top:10px;right:10px;z-index:999999;' +
    'background:#1a1a2e;color:#00ff88;padding:15px 20px;border-radius:10px;' +
    'font-family:monospace;font-size:14px;box-shadow:0 4px 20px rgba(0,255,136,0.3);' +
    'border:1px solid #00ff88;';
  loadingDiv.textContent = 'Caricamento cartelle...';
  document.body.appendChild(loadingDiv);

  const folders = await loadFolders(token);
  loadingDiv.remove();

  if (!folders.length) { console.error('Nessuna cartella trovata.'); return; }

  // Mostra la modale e attendi la scelta
  let config;
  try {
    config = await showConfigModal(folders);
  } catch (e) {
    console.log('Download annullato.');
    return;
  }

  const { folder, emailsPerZip } = config;
  const folderId = folder.id;
  const folderName = folder.name;
  const totalItems = folder.count;

  // UI di stato
  let statusDiv = document.createElement('div');
  statusDiv.id = 'emlDownloadStatus';
  statusDiv.style.cssText = 'position:fixed;top:10px;right:10px;z-index:999999;' +
    'background:#1a1a2e;color:#00ff88;padding:20px;border-radius:10px;' +
    'font-family:monospace;font-size:14px;max-width:400px;' +
    'box-shadow:0 4px 20px rgba(0,255,136,0.3);border:1px solid #00ff88;';
  document.body.appendChild(statusDiv);
  const updateUI = (html) => { statusDiv.innerHTML = html; };
  updateUI(`<b>${folderName} Downloader</b><br>${totalItems} email trovate<br>Inizio...`);

  const totalZips = Math.ceil(totalItems / emailsPerZip);
  let globalProcessed = 0, errors = 0, zipPartNum = 0;
  let currentZip = new MiniZip(), emailsInCurrentZip = 0;

  const saveZip = (zip, partNum) => {
    const data = zip.generate();
    const blob = new Blob([data], { type: 'application/zip' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = totalZips > 1
      ? `${folderName}_emails_part${String(partNum).padStart(2, '0')}.zip`
      : `${folderName}_emails.zip`;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { a.remove(); URL.revokeObjectURL(url); }, 5000);
  };

  let skip = 0;
  while (skip < totalItems) {
    const listResp = await fetch(
      `https://outlook.office.com/api/v2.0/me/mailfolders/${folderId}/messages?$top=${API_BATCH_SIZE}&$skip=${skip}&$select=Id,Subject,ReceivedDateTime&$orderby=ReceivedDateTime desc`,
      { headers }
    );
    if (!listResp.ok) { updateUI(`Errore lista: ${listResp.status}`); return; }
    const messages = (await listResp.json()).value;
    if (!messages?.length) break;

    for (const msg of messages) {
      try {
        const mimeResp = await fetch(
          `https://outlook.office.com/api/v2.0/me/messages/${msg.Id}/$value`,
          { headers: { 'Authorization': 'Bearer ' + token } }
        );
        if (mimeResp.ok) {
          const mimeData = new Uint8Array(await mimeResp.arrayBuffer());
          const safeName = (msg.Subject || 'no_subject')
            .replace(/[^a-zA-Z0-9\u00C0-\u00FF _\-\.]/g, '_').substring(0, 80);
          const date = (msg.ReceivedDateTime || 'nodate').substring(0, 10);
          currentZip.addFile(`${date}_${safeName}_${globalProcessed}.eml`, mimeData);
          emailsInCurrentZip++;
        } else { errors++; }
      } catch(e) { errors++; }

      globalProcessed++;
      if (globalProcessed % 5 === 0) {
        const pct = Math.round(globalProcessed / totalItems * 100);
        updateUI(`<b>${folderName} Downloader</b><br>` +
          `Scaricate: ${globalProcessed}/${totalItems} (${pct}%)<br>` +
          `Errori: ${errors} | ZIP: ${zipPartNum + 1}/${totalZips}<br>` +
          `<progress value="${globalProcessed}" max="${totalItems}" style="width:100%"></progress>`);
      }

      if (emailsInCurrentZip >= emailsPerZip) {
        zipPartNum++;
        updateUI(`<b>Generazione ZIP ${zipPartNum}/${totalZips}...</b>`);
        saveZip(currentZip, zipPartNum);
        currentZip = new MiniZip();
        emailsInCurrentZip = 0;
        await new Promise(r => setTimeout(r, 1000));
      }
      await new Promise(r => setTimeout(r, DELAY_MS));
    }
    skip += API_BATCH_SIZE;
  }

  if (emailsInCurrentZip > 0) { zipPartNum++; saveZip(currentZip, zipPartNum); }

  updateUI(`<b>Download completato!</b><br>` +
    `Email: ${globalProcessed} | Errori: ${errors} | ZIP: ${zipPartNum}`);
  console.log(`${folderName}: ${globalProcessed} email in ${zipPartNum} ZIP`);
})();
