function setStatus(message){
  const el = document.getElementById('statusText');
  if(el) el.textContent = message;
  console.log(message);
}

let selectedFile = null;
let lastDownloadUrl = null;

const drop = document.getElementById('dropzone');
const fileInput = document.getElementById('fileInput');
const browseBtn = document.getElementById('browseBtn');
const fileInfo = document.getElementById('fileInfo');
const generateBtn = document.getElementById('generateBtn');
const downloadBtn = document.getElementById('downloadBtn');
const resultsCard = document.getElementById('resultsCard');
const progressWrap = document.getElementById('progressWrap');
const progressInner = document.getElementById('progressInner');
const statusDot = document.getElementById('statusDot');

function showToast(message, timeout=2200){
  let toast = document.getElementById('ui_toast');
  if(!toast){
    toast = document.createElement('div');
    toast.id = 'ui_toast';
    toast.className = 'toast';
    document.body.appendChild(toast);
  }
  toast.textContent = message;
  toast.classList.add('show');
  setTimeout(() => toast.classList.remove('show'), timeout);
}

function setProgress(percent){
  if(progressWrap) progressWrap.style.display = 'block';
  if(progressInner) progressInner.style.width = `${Math.max(0, Math.min(100, percent))}%`;
}

function setProcessing(isProcessing){
  generateBtn.disabled = isProcessing;
  generateBtn.innerHTML = isProcessing
    ? '<span class="spinner"></span> Processing...'
    : '<svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><polyline points="13 2 3 14 12 14 11 22 21 10 12 10 13 2"/></svg> Generate';
}

function resetResults(){
  if(resultsCard) resultsCard.style.display = 'none';
  if(lastDownloadUrl){
    URL.revokeObjectURL(lastDownloadUrl);
    lastDownloadUrl = null;
  }
}

function handleFiles(files){
  if(!files || files.length === 0) return;

  const file = files[0];
  const name = file.name.toLowerCase();
  if(!(name.endsWith('.xlsx') || name.endsWith('.xls'))){
    selectedFile = null;
    resetResults();
    if(statusDot) statusDot.className = 'status-dot';
    setStatus('Invalid Excel file.');
    showToast('Invalid Excel file');
    return;
  }

  selectedFile = file;
  resetResults();
  fileInfo.style.display = 'block';
  document.getElementById('fileName').textContent = file.name;
  if(statusDot) statusDot.className = 'status-dot done';
  setProgress(0);
  setStatus('File ready. Enter output columns and generate results.');
}

async function readErrorMessage(response){
  try{
    const data = await response.json();
    return data.error || 'Request failed.';
  } catch(e){
    return 'Request failed.';
  }
}

async function generateResults(){
  if(!selectedFile){
    setStatus('No file selected.');
    showToast('No file selected');
    return;
  }

  const numOutputs = parseInt(document.getElementById('outputCount').value, 10);
  if(!Number.isInteger(numOutputs) || numOutputs < 1){
    setStatus('Invalid number of outputs.');
    showToast('Invalid number of outputs');
    return;
  }

  resetResults();
  setProcessing(true);
  if(statusDot) statusDot.className = 'status-dot active';

  try{
    setProgress(15);
    setStatus('Uploading file...');

    const formData = new FormData();
    formData.append('file', selectedFile);
    formData.append('num_outputs', String(numOutputs));

    setProgress(35);
    setStatus('Processing...');

    const response = await fetch('/generate', {
      method: 'POST',
      body: formData
    });

    if(!response.ok){
      throw new Error(await readErrorMessage(response));
    }

    setProgress(75);
    setStatus('Generating workbook...');

    const blob = await response.blob();
    const url = URL.createObjectURL(blob);
    lastDownloadUrl = url;

    const filename = 'CALG_Results.xlsx';
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = filename;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();

    document.getElementById('resFile').textContent = filename;
    document.getElementById('resOutputs').textContent = numOutputs;
    if(resultsCard) resultsCard.style.display = 'block';
    if(downloadBtn) {
      downloadBtn.onclick = () => {
        const retry = document.createElement('a');
        retry.href = url;
        retry.download = filename;
        document.body.appendChild(retry);
        retry.click();
        retry.remove();
      };
    }

    setProgress(100);
    if(statusDot) statusDot.className = 'status-dot done';
    setStatus('Download ready.');
    showToast('Results generated');
  } catch(error){
    if(statusDot) statusDot.className = 'status-dot';
    setProgress(0);
    setStatus(error.message || 'Workbook generation failure.');
    showToast(error.message || 'Workbook generation failure', 3600);
  } finally {
    setProcessing(false);
  }
}

browseBtn.addEventListener('click', (event) => {
  event.preventDefault();
  event.stopPropagation();
  fileInput.click();
});

fileInput.addEventListener('change', (event) => handleFiles(event.target.files));

['dragenter', 'dragover'].forEach(eventName => {
  drop.addEventListener(eventName, (event) => {
    event.preventDefault();
    drop.classList.add('dragover');
  });
});

['dragleave', 'drop'].forEach(eventName => {
  drop.addEventListener(eventName, (event) => {
    event.preventDefault();
    drop.classList.remove('dragover');
  });
});

drop.addEventListener('drop', (event) => {
  const transfer = event.dataTransfer;
  if(transfer && transfer.files) handleFiles(transfer.files);
});

generateBtn.addEventListener('click', generateResults);

document.querySelectorAll('.tab').forEach(tab => {
  tab.addEventListener('click', () => {
    document.querySelectorAll('.tab').forEach(item => item.classList.remove('active'));
    tab.classList.add('active');
  });
});

setStatus('Ready. Upload a CALG Excel file to begin.');
