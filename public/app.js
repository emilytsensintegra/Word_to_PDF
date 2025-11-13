const form = document.getElementById('convert-form');
const result = document.getElementById('result');
const progressWrap = document.getElementById('progress-wrap');
const progressBar = document.getElementById('progress');
const progressText = document.getElementById('progress-text');
const downloadWrap = document.getElementById('download-wrap');
const downloadBtn = document.getElementById('download');
const fileInput = document.getElementById('file');
const dropzone = document.getElementById('dropzone');
const fileInfo = document.getElementById('file-info');
const submitBtn = document.getElementById('submit-btn');

const formatBytes = (bytes) => {
  if (bytes === undefined || bytes === null) return '';
  if (bytes === 0) return '0 B';
  const units = ['B', 'KB', 'MB', 'GB'];
  const i = Math.min(units.length - 1, Math.floor(Math.log(bytes) / Math.log(1024)));
  const value = bytes / Math.pow(1024, i);
  return `${value.toFixed(value > 9 ? 0 : 1)} ${units[i]}`;
};

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

function updateFileInfo(file) {
  if (!file) {
    fileInfo.textContent = 'Nenhum arquivo selecionado.';
    dropzone.classList.remove('has-file');
    return;
  }
  dropzone.classList.add('has-file');
  fileInfo.innerHTML = `<strong>${file.name}</strong> - ${formatBytes(file.size)}`;
}

function setBusyState(isBusy) {
  submitBtn.disabled = isBusy;
  dropzone.classList.toggle('is-busy', isBusy);
}

fileInput.addEventListener('change', () => {
  updateFileInfo(fileInput.files[0]);
});

if (dropzone) {
  ['dragenter', 'dragover'].forEach((evtName) => {
    dropzone.addEventListener(evtName, (event) => {
      event.preventDefault();
      event.stopPropagation();
      dropzone.classList.add('is-dragover');
    });
  });

  ['dragleave', 'drop'].forEach((evtName) => {
    dropzone.addEventListener(evtName, (event) => {
      event.preventDefault();
      event.stopPropagation();
      dropzone.classList.remove('is-dragover');
    });
  });

  dropzone.addEventListener('drop', (event) => {
    const fileList = event.dataTransfer && event.dataTransfer.files;
    if (!fileList || !fileList.length) return;
    const file = fileList[0];
    try {
      if (window.DataTransfer) {
        const dt = new DataTransfer();
        dt.items.add(file);
        fileInput.files = dt.files;
      } else {
        fileInput.files = fileList;
      }
    } catch (err) {
      console.warn('Falha ao sincronizar input com dropzone:', err);
    }
    updateFileInfo(file);
  });
}

form.addEventListener('submit', (e) => {
  e.preventDefault();
  resetUI();

  const target = document.getElementById('target').value;
  const file = fileInput.files[0];
  if (!file) {
    showMessage('Selecione um arquivo primeiro.', true);
    return;
  }

  setBusyState(true);
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
    const pct = (ev.loaded / ev.total) * 50; // 0-50% para upload
    setProgress(pct, `Enviando: ${Math.floor(pct)}%`);
  };

  xhr.upload.onload = () => {
    let current = Math.max(progressBar.value, 50);
    animateTimer = setInterval(() => {
      if (current < 90) {
        current += 1;
        setProgress(current, `Convertendo: ${current}%`);
      }
    }, 200);
  };

  xhr.onreadystatechange = () => {
    if (xhr.readyState === 2) {
      const len = xhr.getResponseHeader('Content-Length');
      respTotal = len ? parseInt(len, 10) : 0;
    }
  };

  xhr.onprogress = (ev) => {
    if (!respTotal || !ev.lengthComputable) return;
    const pct = 90 + (ev.loaded / respTotal) * 10;
    setProgress(pct, `Baixando: ${Math.floor(pct)}%`);
  };

  xhr.onerror = () => {
    if (animateTimer) clearInterval(animateTimer);
    setBusyState(false);
    showMessage('Erro de rede durante a conversao.', true);
  };

  xhr.onload = async () => {
    if (animateTimer) clearInterval(animateTimer);
    setBusyState(false);

    if (xhr.status < 200 || xhr.status >= 300) {
      try {
        const txt = await xhr.response.text();
        const data = JSON.parse(txt);
        showMessage(data && data.error ? data.error : 'Falha na conversao.', true);
      } catch {
        showMessage('Falha na conversao.', true);
      }
      return;
    }

    setProgress(100, 'Concluido: 100%');

    let filename = 'convertido';
    const disp = xhr.getResponseHeader('Content-Disposition');
    if (disp) {
      const match = /filename=\"?([^\";]+)\"?/i.exec(disp);
      if (match && match[1]) filename = match[1];
    } else if (target === 'pdf') {
      filename = file.name.replace(/\.[^.]+$/, '') + '.pdf';
    } else {
      filename = file.name.replace(/\.[^.]+$/, '') + '.docx';
    }

    const blob = xhr.response;
    const href = URL.createObjectURL(blob);
    downloadBtn.href = href;
    downloadBtn.download = filename;
    downloadWrap.classList.remove('hidden');

    showMessage('Conversao concluida. Clique em <strong>Baixar arquivo convertido</strong>.');
  };

  xhr.send(file);
});
