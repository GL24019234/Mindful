import { headingColors, setColor, getLevel, getLabel } from './colorConfig.js';
import { runAllDetection, clearHighlights } from './autoDetect.js';
import { generateTOC } from './ToCgenerator.js';

// ── Available Word highlight colors ──────────────────────────────────────────

const COLORS = [
  { name: 'Yellow',      hex: '#FFFF00' },
  { name: 'Cyan',        hex: '#00FFFF' },
  { name: 'Pink',        hex: '#FF69B4' },
  { name: 'Green',       hex: '#00FF00' },
  { name: 'Orange',      hex: '#FFA500' },
  { name: 'Red',         hex: '#FF0000' },
  { name: 'DarkBlue',    hex: '#00008B' },
  { name: 'Turquoise',   hex: '#40E0D0' },
];

// ── Tab switching ─────────────────────────────────────────────────────────────

document.querySelectorAll('.tab-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
    btn.classList.add('active');
    document.getElementById(`tab-${btn.dataset.tab}`).classList.add('active');
  });
});

// ── Build color swatches ──────────────────────────────────────────────────────

function buildSwatches() {
  const levels = ['section', 'pointHeader', 'subPoint'];

  levels.forEach(level => {
    const container = document.querySelector(`.swatches[data-level="${level}"]`);

    COLORS.forEach(color => {
      const swatch = document.createElement('div');
      swatch.className = 'swatch';
      swatch.style.background = color.hex;
      swatch.title = color.name;

      // mark current selection
      if (headingColors[level] === color.name) {
        swatch.classList.add('selected');
      }

      swatch.addEventListener('click', () => {
        // deselect all swatches in this row
        container.querySelectorAll('.swatch').forEach(s => s.classList.remove('selected'));
        swatch.classList.add('selected');
        setColor(level, color.name);
      });

      container.appendChild(swatch);
    });
  });
}

// ── Status helpers ────────────────────────────────────────────────────────────

function setStatus(elementId, message, type = 'info') {
  const el = document.getElementById(elementId);
  el.textContent = message;
  el.className = `status ${type}`;
}

function clearStatus(elementId) {
  const el = document.getElementById(elementId);
  el.textContent = '';
  el.className = 'status';
}

// ── Detect tab ────────────────────────────────────────────────────────────────

document.getElementById('btn-detect').addEventListener('click', async () => {
  clearStatus('detect-status');
  setStatus('detect-status', 'Running detection...');

  const options = {
    useStyle:   document.getElementById('toggle-style').checked,
    useBold:    document.getElementById('toggle-bold').checked,
    usePattern: document.getElementById('toggle-pattern').checked,
  };

  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      await runAllDetection(context, body, options);
    });
    setStatus('detect-status', 'Detection complete.', 'success');
  } catch (err) {
    setStatus('detect-status', `Error: ${err.message}`, 'error');
  }
});

document.getElementById('btn-clear').addEventListener('click', async () => {
  clearStatus('detect-status');

  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      await clearHighlights(context, body);
    });
    setStatus('detect-status', 'Highlights cleared.', 'success');
  } catch (err) {
    setStatus('detect-status', `Error: ${err.message}`, 'error');
  }
});

// ── Generate tab ──────────────────────────────────────────────────────────────

document.getElementById('btn-preview').addEventListener('click', async () => {
  const list = document.getElementById('preview-list');
  list.innerHTML = '<li class="empty-state">Loading...</li>';

  try {
    await Word.run(async (context) => {
      const body = context.document.body;

      // load paragraphs
      const paragraphs = body.paragraphs;
      context.load(paragraphs, 'items');
      await context.sync();

      // load each paragraph's text and font separately
      for (const para of paragraphs.items) {
        context.load(para, 'text');
        context.load(para.font, 'highlightColor');
      }
      await context.sync();

      const entries = [];
      for (const para of paragraphs.items) {
        const level = getLevel(para.font.highlightColor);
        if (level && para.text.trim().length > 0) {
          entries.push({ text: para.text.trim(), level });
        }
      }

      list.innerHTML = '';

      if (entries.length === 0) {
        list.innerHTML = '<li class="empty-state">No highlighted headings found. Run auto-detect or manually highlight headings first.</li>';
        return;
      }

      entries.forEach(entry => {
        const li = document.createElement('li');
        li.className = `preview-item preview-indent-${entry.level}`;

        const badge = document.createElement('span');
        badge.className = `level-badge level-${entry.level}`;
        badge.textContent = `H${entry.level}`;

        const text = document.createElement('span');
        text.className = 'preview-text';
        text.textContent = entry.text;

        li.appendChild(badge);
        li.appendChild(text);
        list.appendChild(li);
      });
    });
  } catch (err) {
    list.innerHTML = `<li class="empty-state" style="color:#c00">Error: ${err.message}</li>`;
  }
});

document.getElementById('btn-generate').addEventListener('click', async () => {
  clearStatus('generate-status');
  setStatus('generate-status', 'Generating TOC...');

  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      await generateTOC(context, body);
    });
    setStatus('generate-status', 'TOC generated successfully.', 'success');
  } catch (err) {
    setStatus('generate-status', `Error: ${err.message}`, 'error');
  }
});

// ── Init ──────────────────────────────────────────────────────────────────────

Office.onReady(() => {
  buildSwatches();
  // in taskpane.js, inside Office.onReady
document.querySelectorAll('.btn-highlight').forEach(btn => {
  btn.addEventListener('click', async () => {
    const level = btn.dataset.level;
    const color = headingColors[level];

    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      context.load(selection, 'text');
      await context.sync();

      if (selection.text.trim().length === 0) {
        setStatus('config-status', 'No text selected.', 'error');
        return;
      }

      selection.font.highlightColor = color;
      await context.sync();
    });
  });
});
});