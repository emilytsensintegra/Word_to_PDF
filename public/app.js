const form = document.getElementById('convert-form');
const result = document.getElementById('result');
const progressWrap = document.getElementById('progress-wrap');
const progressBar = document.getElementById('progress');
const progressText = document.getElementById('progress-text');
const downloadWrap = document.getElementById('download-wrap');
const downloadBtn = document.getElementById('download');

function showMessage(html, isError = false) {
  result.classList.remove('hidden');
  result.classList.toggle('error', isError);
  result.innerHTML = html;
}

function resetUI() {
  showMessage('', false);
  result.classList.add('hidden');
  downloadWrap.classList.add('hidden');
  setProgress(0, 'Progresso: 0%');
}

function setProgress(val, label) {
  const v = Math.max(0, Math.min(100, Math.floor(val)));
  progressWrap.classList.remove('hidden');
  progressBar.value = v;
  progressText.textContent = label || `Progresso: ${v}%`;
}

form.addEventListener('submit', (e) => {
  e.preventDefault();
  resetUI();

  const fileInput = document.getElementById('file');
  const target = document.getElementById('target').value;
  const file = fileInput.files[0];
  if (!file) {
    showMessage('Selecione um arquivo primeiro.', true);
    return;
  }

  showMessage('Enviando e convertendo, aguarde...');

  const url = `/convert?target=${encodeURIComponent(target)}&filename=${encodeURIComponent(file.name)}`;
  const xhr = new XMLHttpRequest();
  xhr.open('POST', url, true);
  xhr.responseType = 'blob';
  xhr.setRequestHeader('Content-Type', file.type || 'application/octet-stream');

  let animateTimer = null;
  let respTotal = 0;

  xhr.upload.onprogress = (ev) => {
    if (!ev.lengthComputable) return;
    const pct = (ev.loaded / ev.total) * 50; // 0–50% para upload
    setProgress(pct, `Enviando: ${Math.floor(pct)}%`);
  };

  xhr.upload.onload = () => {
    // Simula progresso durante a conversão até ~90%
    let current = Math.max(progressBar.value, 50);
    animateTimer = setInterval(() => {
      if (current < 90) {
        current += 1;
        setProgress(current, `Convertendo: ${current}%`);
      }
    }, 200);
  };

  xhr.onreadystatechange = () => {
    if (xhr.readyState === 2) { // HEADERS_RECEIVED
      const len = xhr.getResponseHeader('Content-Length');
      respTotal = len ? parseInt(len, 10) : 0;
    }
  };

  xhr.onprogress = (ev) => {
    if (!respTotal || !ev.lengthComputable) return;
    // 90–100% para download
    const pct = 90 + (ev.loaded / respTotal) * 10;
    setProgress(pct, `Baixando: ${Math.floor(pct)}%`);
  };

  xhr.onerror = () => {
    if (animateTimer) clearInterval(animateTimer);
    showMessage('Erro de rede durante a conversão.', true);
  };

  xhr.onload = async () => {
    if (animateTimer) clearInterval(animateTimer);

    if (xhr.status < 200 || xhr.status >= 300) {
      try {
        const txt = await xhr.response.text();
        const data = JSON.parse(txt);
        showMessage(data && data.error ? data.error : 'Falha na conversão.', true);
      } catch {
        showMessage('Falha na conversão.', true);
      }
      return;
    }

    setProgress(100, 'Concluído: 100%');

    // Deriva filename
    let filename = 'convertido';
    const disp = xhr.getResponseHeader('Content-Disposition');
    if (disp) {
      const m = /filename=\"?([^\";]+)\"?/i.exec(disp);
      if (m && m[1]) filename = m[1];
    } else if (target === 'pdf') {
      filename = (file.name.replace(/\.[^.]+$/, '')) + '.pdf';
    } else {
      filename = (file.name.replace(/\.[^.]+$/, '')) + '.docx';
    }

    // Exibe botão de download
    const blob = xhr.response;
    const href = URL.createObjectURL(blob);
    downloadBtn.href = href;
    downloadBtn.download = filename;
    downloadWrap.classList.remove('hidden');

    showMessage(`Conversão concluída. Clique em <strong>Baixar arquivo convertido</strong>.`);
  };

  xhr.send(file);
});

