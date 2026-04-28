/* eslint-env browser */

document.addEventListener('DOMContentLoaded', () => {
  const workbookInput = document.getElementById('workbook-input');
  const workbookLabel = document.getElementById('workbook-label');
  const instructionInput = document.getElementById('instruction-input');
  const instructionLabel = document.getElementById('instruction-label');
  const resourceInput = document.getElementById('resource-input');
  const resourceList = document.getElementById('resource-list');
  const instructionText = document.getElementById('instruction-text');
  const runBtn = document.getElementById('run-btn');
  const progressSection = document.getElementById('progress-section');
  const progressFill = document.getElementById('progress-fill');
  const statusText = document.getElementById('status-text');
  const resultSection = document.getElementById('result-section');
  const resultIcon = document.getElementById('result-icon');
  const resultText = document.getElementById('result-text');
  const resultDetail = document.getElementById('result-detail');

  let workbookFile = null;
  let instructionFile = null;
  const resourceFiles = [];

  function updateRunButton() {
    const hasWorkbook = !!workbookFile;
    const hasInstructions = !!instructionFile || instructionText.value.trim().length > 0;
    runBtn.disabled = !(hasWorkbook && hasInstructions);
  }

  workbookInput.addEventListener('change', (e) => {
    workbookFile = e.target.files[0];
    if (workbookFile) {
      workbookLabel.querySelector('.file-text').textContent = workbookFile.name;
      workbookLabel.classList.add('has-file');
    }
    updateRunButton();
  });

  instructionInput.addEventListener('change', (e) => {
    instructionFile = e.target.files[0];
    if (instructionFile) {
      instructionLabel.querySelector('.file-text').textContent = instructionFile.name;
      instructionLabel.classList.add('has-file');
    }
    updateRunButton();
  });

  resourceInput.addEventListener('change', (e) => {
    for (const f of e.target.files) {
      resourceFiles.push(f);
    }
    renderResourceList();
  });

  instructionText.addEventListener('input', updateRunButton);

  function renderResourceList() {
    resourceList.innerHTML = '';
    resourceFiles.forEach((f, i) => {
      const chip = document.createElement('span');
      chip.className = 'resource-chip';
      chip.innerHTML = `📎 ${f.name} <button data-idx="${i}">✕</button>`;
      chip.querySelector('button').addEventListener('click', () => {
        resourceFiles.splice(i, 1);
        renderResourceList();
      });
      resourceList.appendChild(chip);
    });
  }

  runBtn.addEventListener('click', async () => {
    runBtn.disabled = true;
    progressSection.hidden = false;
    resultSection.hidden = true;
    progressFill.style.width = '10%';
    statusText.textContent = 'Connecting to server...';

    try {
      const settings = await chrome.storage.sync.get({
        serverURL: 'http://localhost:8080',
        apiKey: '',
      });

      progressFill.style.width = '30%';
      statusText.textContent = 'Sending to engine...';

      let instrContent = instructionText.value.trim();
      if (instructionFile) {
        instrContent = await readFileAsText(instructionFile);
      }

      const maxLayer = document.getElementById('max-layer').value;
      const dryRun = document.getElementById('dry-run').checked;

      const body = {
        jsonrpc: '2.0',
        id: Math.floor(Math.random() * 999999),
        method: 'tools/call',
        params: {
          name: 'complete_assignment',
          arguments: {
            workbook_path: workbookFile.name,
            instruction_text: instrContent,
            max_layer: maxLayer,
            dry_run: String(dryRun),
          },
        },
      };

      const headers = { 'Content-Type': 'application/json' };
      if (settings.apiKey) {
        headers['X-API-Key'] = settings.apiKey;
      }

      progressFill.style.width = '50%';
      statusText.textContent = 'Engine processing...';

      const resp = await fetch(`${settings.serverURL}/mcp`, {
        method: 'POST',
        headers,
        body: JSON.stringify(body),
      });

      progressFill.style.width = '90%';

      if (!resp.ok) {
        throw new Error(`Server returned ${resp.status}`);
      }

      const data = await resp.json();

      progressFill.style.width = '100%';
      statusText.textContent = 'Done!';

      if (data.error) {
        showResult(false, data.error.message, '');
      } else {
        const content = data.result?.content?.[0]?.text || 'Completed';
        let parsed;
        try { parsed = JSON.parse(content); } catch { parsed = null; }
        if (parsed) {
          showResult(
            parsed.success !== false,
            parsed.success !== false ? 'Assignment Complete' : 'Completed with Issues',
            `Tasks: ${parsed.tasks_completed || '?'}/${parsed.tasks_total || '?'}`
          );
        } else {
          showResult(true, 'Engine Completed', content.substring(0, 200));
        }
      }
    } catch (err) {
      progressFill.style.width = '100%';
      statusText.textContent = 'Error';
      showResult(false, 'Connection Failed', err.message);
    }

    runBtn.disabled = false;
  });

  function showResult(success, title, detail) {
    resultSection.hidden = false;
    resultSection.className = `result-section ${success ? 'success' : 'error'}`;
    resultIcon.textContent = success ? '✅' : '❌';
    resultText.textContent = title;
    resultDetail.textContent = detail;
  }

  function readFileAsText(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result);
      reader.onerror = reject;
      reader.readAsText(file);
    });
  }
});
