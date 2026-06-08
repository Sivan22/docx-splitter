import './styles.css';
import { splitDocx, zipParts } from './core/split.js';
import { mergeDocx } from './core/merge.js';

const $ = (id) => document.getElementById(id);

function download(name, bytes, mime) {
  const blob = new Blob([bytes], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = name;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function showError(el, msg) {
  el.textContent = msg;
  el.hidden = !msg;
}

// --- Tabs ---
function activate(which) {
  $('tab-split').classList.toggle('active', which === 'split');
  $('tab-merge').classList.toggle('active', which === 'merge');
  $('panel-split').hidden = which !== 'split';
  $('panel-merge').hidden = which !== 'merge';
}
$('tab-split').addEventListener('click', () => activate('split'));
$('tab-merge').addEventListener('click', () => activate('merge'));

// --- Drop zone wiring (shared) ---
function wireDrop(zoneId, inputId, onFiles) {
  const zone = $(zoneId);
  const input = $(inputId);
  zone.addEventListener('click', () => input.click());
  zone.addEventListener('dragover', (e) => {
    e.preventDefault();
    zone.classList.add('over');
  });
  zone.addEventListener('dragleave', () => zone.classList.remove('over'));
  zone.addEventListener('drop', (e) => {
    e.preventDefault();
    zone.classList.remove('over');
    onFiles([...e.dataTransfer.files]);
  });
  input.addEventListener('change', () => onFiles([...input.files]));
}

// --- Split flow ---
wireDrop('drop-split', 'file-split', async (files) => {
  showError($('split-error'), '');
  $('split-summary').innerHTML = '';
  const file = files[0];
  if (!file) return;
  const maxWords = Number($('max-words').value);
  try {
    const data = await file.arrayBuffer();
    const { parts, total, summary } = await splitDocx(data, file.name, maxWords);
    const base = file.name.replace(/\.docx$/i, '');
    const { name, bytes } = await zipParts(parts, `${base}_parts.zip`);
    download(name, bytes, 'application/zip');
    $('split-summary').innerHTML =
      `<li><strong>${total} part(s)</strong> — downloaded ${name}</li>` +
      summary.map((s) => `<li>Part ${s.part}: ${s.words} words</li>`).join('');
  } catch (err) {
    showError($('split-error'), err.message);
  }
});

// --- Merge flow ---
wireDrop('drop-merge', 'file-merge', async (files) => {
  showError($('merge-error'), '');
  $('merge-status').textContent = '';
  if (!files.length) return;
  try {
    const payload = await Promise.all(
      files.map(async (f) => ({ name: f.name, data: await f.arrayBuffer() })),
    );
    const { name, bytes } = await mergeDocx(payload);
    download(
      name,
      bytes,
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    );
    $('merge-status').textContent = `Merged ${files.length} part(s) → ${name}`;
  } catch (err) {
    showError($('merge-error'), err.message);
  }
});
