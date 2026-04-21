/* ============================================
   AI Novel Template Generator — Core Logic
   Gemini API Integration & Novel Generation
   ============================================ */

// --- State (single run) ---
const SINGLE_FLOW_ID = 'main';
const state = {
  apiKey: '',
  flows: Object.create(null),
  activeFlowId: null,
  speakingSegment: null, // { audioIndex, segmentIndex, flowId } - for stopping TTS
  ttsDemoPlayback: null, // { audio: HTMLAudioElement, url: string } — Settings voice preview
  imageGenerationDisabled: false,
  imageGenerationDisabledReason: '',
};

function newFlowId() {
  return SINGLE_FLOW_ID;
}

function flowDomId(flowId) {
  return String(flowId || 'flow').replace(/[^a-zA-Z0-9_-]/g, '_');
}

function createFlowState(id, label) {
  return {
    id,
    label: label || 'Untitled run',
    novels: [],
    stories: {},
    audioScripts: {},
    audioScriptSegments: {},
    generatedAudio: {},
    generatedAudioBatches: {},
    generatedScenes: {},
    reviewedNovels: new Set(),
    generating: false,
    status: 'idle',
    errorMessage: '',
  };
}

function getFlow(fid) {
  return fid ? state.flows[fid] : null;
}

function getActiveFlow() {
  return state.activeFlowId ? state.flows[state.activeFlowId] : null;
}

function ensureActiveFlow() {
  if (!state.flows[SINGLE_FLOW_ID]) {
    state.flows[SINGLE_FLOW_ID] = createFlowState(SINGLE_FLOW_ID, 'Current run');
  }
  state.activeFlowId = SINGLE_FLOW_ID;
  return state.flows[SINGLE_FLOW_ID];
}

function ensureFlowPanelDom(flowId) {
  const wrap = document.getElementById('flowPanelsContainer');
  if (!wrap || !flowId) return;
  const fd = flowDomId(flowId);
  if (document.getElementById(`flowPanel_${fd}`)) return;
  const panel = document.createElement('div');
  panel.className = 'flow-panel';
  panel.id = `flowPanel_${fd}`;
  panel.dataset.flowId = flowId;
  panel.hidden = true;
  const inner = document.createElement('div');
  inner.id = `novelsContainer_${fd}`;
  panel.appendChild(inner);
  wrap.appendChild(panel);
}

function switchToFlow(flowId) {
  const targetId = getFlow(flowId) ? flowId : SINGLE_FLOW_ID;
  state.activeFlowId = targetId;
  document.querySelectorAll('.flow-panel').forEach((p) => {
    p.hidden = p.dataset.flowId !== targetId;
  });
  refreshGenerateButtonState();
  updateResultsHeaderCount();
}

function updateResultsHeaderCount() {
  const flow = getActiveFlow();
  const countEl = document.getElementById('resultsCount');
  if (countEl && flow) countEl.textContent = `${flow.novels.length} novels generated`;
}

function renderFlowTabs() {
  // Single-flow mode: tabs are removed from UI.
}

function refreshGenerateButtonState() {
  const flow = getActiveFlow();
  const btn = document.getElementById('generateBtn');
  if (!btn) return;
  const busy = !!flow?.generating;
  btn.disabled = busy;
  btn.classList.toggle('loading', busy);
}

function setFlowStatusLine(text) {
  const el = document.getElementById('flowStatusLine');
  if (el) el.textContent = text || '';
}

function canGenerateImages() {
  // Gemini image generation: always allow when provider is Gemini and key is set (Imagen).
  if (getAIProvider() === 'gemini' && getApiKey()) return true;
  // Free API fallback is blocked by CORS on GitHub Pages; disable flag only affects non-Gemini users.
  if (state.imageGenerationDisabled) return false;
  const isGithubPages = typeof location !== 'undefined' && /(^|\.)github\.io$/i.test(location.hostname || '');
  if (isGithubPages) return false;
  return true;
}

// --- UI Styles (covers + chapters) ---
function injectUiStyles() {
  const id = 'novel_text_patches_styles';
  if (document.getElementById(id)) return;
  const style = document.createElement('style');
  style.id = id;
  style.textContent = `
    .novel-cover-thumb {
      width: 44px;
      height: 44px;
      border-radius: 12px;
      object-fit: cover;
      border: 1px solid rgba(255,255,255,0.14);
      box-shadow: 0 8px 18px rgba(0,0,0,0.35);
      margin-left: 10px;
      flex: 0 0 auto;
    }
    .chapters-wrap { display: grid; gap: 12px; }
    .chapter-block { border: 1px solid rgba(120, 128, 160, 0.25); background: rgba(255,255,255,0.03); border-radius: 14px; padding: 12px 12px; }
    .chapter-hdr { display:flex; gap:10px; align-items:baseline; justify-content:space-between; margin-bottom: 8px; }
    .chapter-title { font-weight: 700; letter-spacing: 0.2px; }
    .chapter-meta { opacity: 0.75; font-size: 12px; }
    .chapter-body { white-space: pre-wrap; line-height: 1.65; }
    .download-templates-wrap { position: relative; display: inline-block; }
    .download-templates-dropdown {
      display: none;
      position: absolute;
      top: 100%;
      left: 0;
      margin-top: 4px;
      min-width: 180px;
      background: var(--bg-card, rgba(15, 23, 55, 0.95));
      border: 1px solid var(--border-color, rgba(124, 58, 237, 0.3));
      border-radius: var(--radius-md, 12px);
      box-shadow: var(--shadow-lg, 0 8px 40px rgba(0,0,0,0.5));
      z-index: 100;
      overflow: hidden;
    }
    .download-templates-dropdown.open { display: block; }
    .download-option {
      display: block;
      width: 100%;
      padding: 10px 14px;
      text-align: left;
      background: transparent;
      border: none;
      color: var(--text-primary, #e8ebf4);
      font-size: 0.9rem;
      cursor: pointer;
      font-family: inherit;
    }
    .download-option:hover { background: rgba(124, 58, 237, 0.2); }
  `;
  document.head.appendChild(style);
}

// --- DOM Ready ---
document.addEventListener('DOMContentLoaded', () => {
  // Early fallback binding so "Load example" still works even if initApp throws later.
  const earlyLoadExampleBtn = document.getElementById('loadExampleBtn');
  if (earlyLoadExampleBtn && !earlyLoadExampleBtn.dataset.boundFallback) {
    earlyLoadExampleBtn.dataset.boundFallback = '1';
    earlyLoadExampleBtn.addEventListener('click', (e) => {
      e.preventDefault();
      try {
        loadExampleTemplates();
      } catch (err) {
        console.error('Early load example handler failed:', err);
      }
    });
  }
  try {
    initApp();
  } catch (e) {
    console.error('Init error:', e);
    showToast('App failed to load: ' + (e?.message || 'Unknown error'), 'error');
  }
});

const HISTORY_KEY = 'novel_generation_history_v1';
const HISTORY_MAX = 30;

function loadHistoryRuns() {
  try {
    const raw = localStorage.getItem(HISTORY_KEY);
    const arr = raw ? JSON.parse(raw) : [];
    return Array.isArray(arr) ? arr : [];
  } catch (_) {
    return [];
  }
}

function saveHistoryRuns(runs) {
  try {
    localStorage.setItem(HISTORY_KEY, JSON.stringify((runs || []).slice(0, HISTORY_MAX)));
  } catch (_) {}
}

function addHistoryRun(formData, novels, flowId = null) {
  if (!Array.isArray(novels) || !novels.length) return;
  const now = new Date();
  const run = {
    id: `${now.getTime()}_${Math.random().toString(16).slice(2)}`,
    createdAt: now.toISOString(),
    provider: getAIProvider() || 'gemini',
    numNovels: novels.length,
    masterPrompt: safeStr(formData?.masterPrompt).slice(0, 220),
    novels,
    flowId: flowId || undefined,
    flowLabel: flowId && getFlow(flowId) ? getFlow(flowId).label : undefined,
  };
  const runs = loadHistoryRuns();
  runs.unshift(run);
  saveHistoryRuns(runs);
}

function formatHistoryRunTitle(run) {
  const dt = run?.createdAt ? new Date(run.createdAt) : null;
  const when = dt && !isNaN(dt.getTime()) ? dt.toLocaleString() : 'Unknown time';
  const provider = safeStr(run?.provider) || 'AI';
  const count = run?.numNovels || (Array.isArray(run?.novels) ? run.novels.length : 0);
  const brief = safeStr(run?.masterPrompt) || 'No prompt';
  return `${when} • ${provider} • ${count} novel(s) — ${brief}`;
}

function openHistory() {
  const modal = document.getElementById('historyModal');
  if (!modal) return;
  modal.style.display = 'flex';
  renderHistoryList();
}

function closeHistory() {
  const modal = document.getElementById('historyModal');
  if (!modal) return;
  modal.style.display = 'none';
}

function renderHistoryList() {
  const list = document.getElementById('historyList');
  if (!list) return;
  const runs = loadHistoryRuns();
  if (!runs.length) {
    list.innerHTML = `<div class="form-hint">No saved runs yet. Generate templates to create history.</div>`;
    return;
  }
  list.innerHTML = runs.map(r => `
    <div style="border:1px solid rgba(255,255,255,0.12); border-radius: 12px; padding: 10px 12px; margin: 10px 0; background: rgba(255,255,255,0.03);">
      <div style="font-weight:700; margin-bottom:8px; line-height:1.35;">${escapeHtml(formatHistoryRunTitle(r))}</div>
      <div style="display:flex; gap:10px; flex-wrap:wrap;">
        <button type="button" class="btn btn-primary btn-sm" data-history-load="${escapeHtml(r.id)}">Load</button>
        <button type="button" class="btn btn-secondary btn-sm" data-history-delete="${escapeHtml(r.id)}">Delete</button>
      </div>
    </div>
  `).join('');

  list.querySelectorAll('[data-history-load]').forEach(btn => {
    btn.addEventListener('click', () => {
      const id = btn.getAttribute('data-history-load');
      const run = loadHistoryRuns().find(x => x.id === id);
      if (!run || !Array.isArray(run.novels)) return;
      const flow = ensureActiveFlow();
      flow.label = safeStr(run.masterPrompt).slice(0, 40) || 'History';
      flow.novels = run.novels;
      flow.stories = {};
      flow.audioScripts = {};
      flow.audioScriptSegments = {};
      flow.generatedAudio = {};
      flow.generatedAudioBatches = {};
      flow.generatedScenes = {};
      flow.reviewedNovels = new Set();
      normalizeNovelsForExport(flow.novels);
      stampCollectionAndCategoriesFromForm(flow.novels);
      applyAutoReviewTemplateToNovels(flow.novels);
      ensureFlowPanelDom(flow.id);
      switchToFlow(flow.id);
      renderFlowResults(flow.id);
      showToast('Loaded history run.', 'success');
      closeHistory();
    });
  });
  list.querySelectorAll('[data-history-delete]').forEach(btn => {
    btn.addEventListener('click', () => {
      const id = btn.getAttribute('data-history-delete');
      const runs2 = loadHistoryRuns().filter(x => x.id !== id);
      saveHistoryRuns(runs2);
      renderHistoryList();
      showToast('Deleted history run.', 'info');
    });
  });
}

function initApp() {
  injectUiStyles();

  renderChangelog();

  // Restore from localStorage
  const savedKeys = localStorage.getItem('gemini_api_keys') || localStorage.getItem('gemini_api_key');
  const apiKeyEl = document.getElementById('apiKey');
  if (savedKeys && apiKeyEl) {
    apiKeyEl.value = savedKeys;
    state.apiKey = getApiKeys()[0] || '';
  }
  const savedProvider = localStorage.getItem('ai_provider');
  const aiProviderSel = document.getElementById('aiProvider');
  if (savedProvider && aiProviderSel) aiProviderSel.value = savedProvider;
  const ttsGroup = document.getElementById('geminiTtsKeyGroup');
  if (ttsGroup) ttsGroup.style.display = (aiProviderSel?.value === 'deepseek') ? 'block' : 'none';

  const savedTtsProvider = localStorage.getItem('tts_provider') || 'gemini';
  const ttsProviderSel = document.getElementById('ttsProvider');
  if (ttsProviderSel) ttsProviderSel.value = savedTtsProvider;
  updateTtsProviderUI();

  const savedTtsKey = localStorage.getItem('tts_api_keys') || localStorage.getItem('gemini_tts_key');
  const ttsKeyEl = document.getElementById('ttsApiKey');
  if (savedTtsKey && ttsKeyEl) ttsKeyEl.value = savedTtsKey;
  const savedAi33 = localStorage.getItem('ai33_api_key');
  const ai33El = document.getElementById('ai33ApiKey');
  if (savedAi33 && ai33El) ai33El.value = savedAi33;
  const savedAi33Url = localStorage.getItem('ai33_base_url');
  const ai33UrlEl = document.getElementById('ai33BaseUrl');
  if (savedAi33Url && ai33UrlEl) ai33UrlEl.value = savedAi33Url;

  ['narratorVoice', 'femaleVoice', 'maleVoice'].forEach(id => {
    const el = document.getElementById(id);
    const saved = localStorage.getItem(id);
    if (saved && el) el.value = saved;
    el?.addEventListener('change', (e) => localStorage.setItem(id, e.target.value));
    document.getElementById(`demoVoice_${id}`)?.addEventListener('click', () => playTtsVoiceDemo(id));
  });

  // Settings modal
  document.getElementById('openSettingsBtn')?.addEventListener('click', () => openSettings());
  document.getElementById('openSettingsFromStatus')?.addEventListener('click', () => openSettings());
  document.getElementById('closeSettingsBtn')?.addEventListener('click', () => closeSettings());
  document.getElementById('saveSettingsBtn')?.addEventListener('click', () => { saveSettings(); closeSettings(); });
  document.getElementById('settingsModal')?.addEventListener('click', (e) => {
    if (e.target.id === 'settingsModal') closeSettings();
  });

  // History modal
  document.getElementById('openHistoryBtn')?.addEventListener('click', () => openHistory());
  document.getElementById('closeHistoryBtn')?.addEventListener('click', () => closeHistory());
  document.getElementById('closeHistoryBtn2')?.addEventListener('click', () => closeHistory());
  document.getElementById('clearHistoryBtn')?.addEventListener('click', () => {
    saveHistoryRuns([]);
    renderHistoryList();
    showToast('History cleared.', 'info');
  });
  document.getElementById('historyModal')?.addEventListener('click', (e) => {
    if (e.target.id === 'historyModal') closeHistory();
  });

  document.getElementById('aiProvider')?.addEventListener('change', (e) => {
    localStorage.setItem('ai_provider', e.target.value);
    if (ttsGroup) ttsGroup.style.display = e.target.value === 'deepseek' ? 'block' : 'none';
  });
  document.getElementById('ttsProvider')?.addEventListener('change', () => updateTtsProviderUI());
  document.getElementById('apiKey')?.addEventListener('input', (e) => {
    const keys = getApiKeys();
    state.apiKey = keys[0] || '';
    localStorage.setItem('gemini_api_keys', e.target.value);
  });
  document.getElementById('ttsApiKey')?.addEventListener('input', (e) => localStorage.setItem('tts_api_keys', e.target.value));
  document.getElementById('geminiTtsKey')?.addEventListener('input', (e) => localStorage.setItem('gemini_tts_key', e.target.value));
  document.getElementById('ai33ApiKey')?.addEventListener('input', (e) => localStorage.setItem('ai33_api_key', e.target.value));
  document.getElementById('ai33BaseUrl')?.addEventListener('input', (e) => localStorage.setItem('ai33_base_url', e.target.value));

  // Template metadata: collection + cateogories (persisted)
  const collectionEl = document.getElementById('collectionName');
  if (collectionEl) {
    const saved = localStorage.getItem('template_collection') || '';
    if (saved) collectionEl.value = saved;
    collectionEl.addEventListener('change', (e) => localStorage.setItem('template_collection', e.target.value || ''));
  }
  const cateogoriesEl = document.getElementById('categoryName');
  if (cateogoriesEl) {
    const saved = localStorage.getItem('template_cateogories') || '';
    if (saved) cateogoriesEl.value = saved;
    cateogoriesEl.addEventListener('change', (e) => localStorage.setItem('template_cateogories', e.target.value || ''));
  }

  const authorEl = document.getElementById('authorName');
  if (authorEl) {
    const saved = localStorage.getItem('template_author') || '';
    if (saved) authorEl.value = saved;
    authorEl.addEventListener('input', (e) => localStorage.setItem('template_author', e.target.value || ''));
  }

  updateApiStatusBadge();

  // Event listeners
  const genBtn = document.getElementById('generateBtn');
  if (genBtn) genBtn.addEventListener('click', handleGenerate);
  ensureActiveFlow();
  ensureFlowPanelDom(state.activeFlowId);
  switchToFlow(state.activeFlowId);
  const autoStoryChk = document.getElementById('autoGenerateStoriesAfterTemplates');
  if (autoStoryChk) {
    autoStoryChk.checked = localStorage.getItem('auto_generate_stories_after_templates') === '1';
    autoStoryChk.addEventListener('change', (e) => {
      localStorage.setItem('auto_generate_stories_after_templates', e.target.checked ? '1' : '0');
    });
  }
  const loadExampleBtn = document.getElementById('loadExampleBtn');
  if (loadExampleBtn) {
    // Keep both bindings so the action still works if one listener path is disrupted.
    loadExampleBtn.addEventListener('click', (e) => {
      e.preventDefault();
      loadExampleTemplates();
    });
    loadExampleBtn.onclick = (e) => {
      e.preventDefault();
      loadExampleTemplates();
    };
  }
  document.getElementById('testApiBtn')?.addEventListener('click', handleTestApi);
  document.getElementById('generateThumbnails34AllBtn')?.addEventListener('click', generateThumbnails34ForAll);
  document.getElementById('generateAllStoriesBtn')?.addEventListener('click', handleGenerateAllStories);
  document.getElementById('fillMissingDataBtn')?.addEventListener('click', handleFillMissingData);
  document.getElementById('downloadAllBtn')?.addEventListener('click', () => {
    const flow = ensureActiveFlow();
    if (!flow.novels.length) {
      showToast('No novels to download', 'error');
      return;
    }
    showToast(`Downloading ${flow.novels.length} template .txt file(s)… Allow multiple downloads if your browser asks.`, 'info');
    handleDownloadAll();
  });

  // Download templates: dropdown (All as .txt | Export CSV | Export XLSX)
  const downloadTemplatesBtn = document.getElementById('downloadTemplatesBtn');
  const downloadTemplatesDropdown = document.getElementById('downloadTemplatesDropdown');
  if (downloadTemplatesBtn && downloadTemplatesDropdown) {
    downloadTemplatesBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      downloadTemplatesDropdown.classList.toggle('open');
    });
    document.addEventListener('click', () => downloadTemplatesDropdown.classList.remove('open'));
    document.getElementById('downloadAllTxtBtn')?.addEventListener('click', () => {
      downloadTemplatesDropdown.classList.remove('open');
      handleDownloadAll();
    });
    document.getElementById('downloadExportCsvBtn')?.addEventListener('click', () => {
      downloadTemplatesDropdown.classList.remove('open');
      handleExportCsv();
    });
    document.getElementById('downloadExportXlsxBtn')?.addEventListener('click', () => {
      downloadTemplatesDropdown.classList.remove('open');
      handleExportXlsx();
    });
    document.getElementById('downloadExportXlsxFitBtn')?.addEventListener('click', () => {
      downloadTemplatesDropdown.classList.remove('open');
      handleExportXlsx({ fitColumns: true });
    });
    document.getElementById('downloadExportZipBtn')?.addEventListener('click', () => {
      downloadTemplatesDropdown.classList.remove('open');
      handleExportZipPackage();
    });
  }

  // File upload
  const fileUploadArea = document.getElementById('fileUploadArea');
  const fileInput = document.getElementById('refFileInput');

  fileUploadArea.addEventListener('click', () => fileInput.click());
  fileUploadArea.addEventListener('dragover', (e) => {
    e.preventDefault();
    fileUploadArea.classList.add('dragover');
  });
  fileUploadArea.addEventListener('dragleave', () => {
    fileUploadArea.classList.remove('dragover');
  });
  fileUploadArea.addEventListener('drop', (e) => {
    e.preventDefault();
    fileUploadArea.classList.remove('dragover');
    if (e.dataTransfer.files.length) {
      handleFileUpload(e.dataTransfer.files[0]);
    }
  });
  fileInput.addEventListener('change', (e) => {
    if (e.target.files.length) {
      handleFileUpload(e.target.files[0]);
    }
  });
}

function renderChangelog() {
  const el = document.getElementById('changelogList');
  if (!el) return;

  const items = [
    {
      version: '2026-03-06',
      bullets: [
        'Multi-flow: Generate multiple stories in parallel using multiple API keys',
        'Full story view separated by chapters (with [CHAPTER N] markers)',
        'Auto-generate cover + thumbnail for each template during review',
        'Export: CSV + XLSX (XLSX embeds cover + thumbnail images)',
      ],
    },
  ];

  el.innerHTML = items.map(i => `
    <div style="margin-bottom:10px">
      <div style="font-weight:700; margin-bottom:6px">${i.version}</div>
      <ul style="margin:0; padding-left: 18px; color: rgba(255,255,255,0.88)">
        ${i.bullets.map(b => `<li style="margin: 4px 0">${b}</li>`).join('')}
      </ul>
    </div>
  `).join('');
}

// --- Example templates (output sample so you can try export without generating) ---
function getExampleTemplates() {
  const today = new Date().toISOString().split('T')[0];
  return [
    {
      title: 'Shadows of the Empathic Order',
      synopsis: 'An empath hunts a rogue—then learns the Order she serves may have started the war.',
      overview: 'In a world where emotions manifest as physical powers, a young woman discovers she can absorb others\' pain—and their memories. When the Empathic Order recruits her to hunt a rogue empath, she must choose between duty and the truth behind the war that shattered her family. Each mission pulls her deeper into a cover-up that redefines who the enemy really is.',
      draftScript: 'Dark fantasy. Protagonist has empathy-based power. Conflict: Order vs Void Syndicate. Themes: trauma, healing, belonging.',
      characters: [
        { name: 'Elara', role: 'protagonist', age: '22', description: 'Empath who absorbs pain and memories.', arc: 'From hiding her power to leading a reckoning.', gender: 'female' },
        { name: 'Kael', role: 'antagonist', age: '30', description: 'Rogue empath, former Order knight.', arc: 'Revealed as victim of Order cover-up.', gender: 'male' }
      ],
      authorName: 'Example Author',
      releaseDate: today,
      narratorTone: 'Third-person limited, tense and atmospheric.',
      background: 'Neo-Victorian city where emotion-magic is regulated by the Order.',
      writingLanguage: 'English',
      chapters: [
        { chapterNumber: 1, title: 'The Awakening', summary: 'Power awakens during attack.' },
        { chapterNumber: 2, title: 'The Order', summary: 'Recruited to find Kael.' },
        { chapterNumber: 3, title: 'The Truth', summary: 'Learns the Order lied.' }
      ],
      themes: ['trauma', 'identity', 'power'],
      genre: 'Dark Fantasy',
      category: 'Adult Fiction',
      collection: 'Passion Exclusives',
      cateogories: 'Fantasy',
      premium: 'yes',
      show: 'yes'
    },
    {
      title: 'Midnight at the Inkwell',
      synopsis: 'A ghostwriter decodes a star author\'s manuscript—and finds a real murder in the margins.',
      overview: 'A ghostwriter for a reclusive celebrity novelist uncovers a real murder tied to the author\'s past. To finish the book and stay alive, she must piece together the story from coded manuscripts and dangerous interviews. Every chapter she writes brings her closer to a confession that someone will kill to keep buried.',
      draftScript: 'Mystery thriller. Ghostwriter protagonist. Celebrity author with a secret. Murder plot mirrors the novel-in-progress.',
      characters: [
        { name: 'Maya', role: 'protagonist', age: '28', description: 'Ghostwriter, sharp and observant.', arc: 'From outsider to confronting the past.', gender: 'female' },
        { name: 'Julian Cross', role: 'supporting', age: '55', description: 'Reclusive bestselling author.', arc: 'From enigma to key witness.', gender: 'male' }
      ],
      authorName: 'Example Author',
      releaseDate: today,
      narratorTone: 'First-person, wry and suspenseful.',
      background: 'New York publishing world; Vermont estate.',
      writingLanguage: 'English',
      chapters: [
        { chapterNumber: 1, title: 'The Contract', summary: 'Maya takes the job.' },
        { chapterNumber: 2, title: 'The Manuscript', summary: 'Code hints at a crime.' },
        { chapterNumber: 3, title: 'The Murder', summary: 'A body turns up.' }
      ],
      themes: ['identity', 'truth', 'art'],
      genre: 'Suspense Thriller',
      category: 'Adult Fiction',
      collection: 'Top Picks',
      cateogories: 'Suspense Thriller',
      premium: 'yes',
      show: 'yes'
    }
  ];
}

function stampCollectionAndCategoriesFromForm(novels) {
  const collection = safeStr(document.getElementById('collectionName')?.value);
  const cateogories = safeStr(document.getElementById('categoryName')?.value);
  const formAuthor = safeStr(document.getElementById('authorName')?.value);
  if (!Array.isArray(novels)) return;
  novels.forEach(novel => {
    if (!novel || typeof novel !== 'object') return;
    if (collection) novel.collection = collection;
    if (cateogories) {
      novel.cateogories = normalizeCategoryName(cateogories);
      novel.category = novel.cateogories;
      if (!novel.genre) novel.genre = cateogories;
    }
    if (formAuthor) novel.authorName = formAuthor;
    else if (!safeStr(novel.authorName) || /^(unknown author|anonymous|example author)$/i.test(safeStr(novel.authorName)))
      novel.authorName = pickRandomAuthorName();
    if (!safeStr(novel.premium)) novel.premium = 'yes';
    if (!safeStr(novel.show)) novel.show = 'yes';
  });
}

/** Ensure every novel has data needed for export: synopsis, at least one chapter, premium, show, author, etc. */
function normalizeNovelsForExport(novels) {
  if (!Array.isArray(novels)) return;
  novels.forEach(novel => {
    if (!novel || typeof novel !== 'object') return;
    const rawSyn0 = safeStr(novel.synopsis);
    let ov = safeStr(novel.overview);
    if (!ov || ov === 'N/A') {
      if (rawSyn0.length > 100) ov = rawSyn0;
      else ov = safeStr(novel.draftScript).slice(0, 1800) || rawSyn0 || (novel.title ? `A story: ${novel.title}.` : 'No overview yet.');
    }
    novel.overview = clampText(ov, 2000);

    let hook = rawSyn0;
    if (!hook || hook === 'N/A') hook = novel.overview;
    if (safeStr(hook).length > 100) {
      const parts = safeStr(hook).split(/(?<=[.!?])\s+/);
      hook = (parts[0] || safeStr(hook).slice(0, 100)).trim();
    }
    novel.synopsis = clampText(hook, 100);
    if (!novel.chapters || !Array.isArray(novel.chapters) || novel.chapters.length === 0) {
      const seed = safeStr(novel.overview) || safeStr(novel.synopsis);
      novel.chapters = [
        { chapterNumber: 1, title: 'Chapter 1', summary: seed ? seed.substring(0, 200) + (seed.length > 200 ? '…' : '') : 'Opening.' }
      ];
    }
    const formAuthor = safeStr(document.getElementById('authorName')?.value);
    if (formAuthor) novel.authorName = formAuthor;
    else if (!safeStr(novel.authorName) || /^(unknown author|anonymous|example author)$/i.test(safeStr(novel.authorName)))
      novel.authorName = pickRandomAuthorName();
    if (!safeStr(novel.premium)) novel.premium = 'yes';
    if (!safeStr(novel.show)) novel.show = 'yes';
    if (!safeStr(novel.collection)) novel.collection = safeStr(document.getElementById('collectionName')?.value);
    // Enforce allowed categories (and keep cateogories/category consistent).
    const chosen = normalizeCategoryName(
      safeStr(novel.cateogories || novel.categories || novel.category || novel.genre) ||
      safeStr(document.getElementById('categoryName')?.value)
    );
    novel.cateogories = chosen;
    novel.category = chosen;
    if (!safeStr(novel.thumbnailPrompt)) novel.thumbnailPrompt = buildThumbnailPromptFromNovel(novel);
    if (Array.isArray(novel.themes)) novel.themes = normalizeTagList(novel.themes);
    if (Array.isArray(novel.tags)) novel.tags = normalizeTagList(novel.tags);

    // Chapter outline: limit to 100 chars per chapter line (title + " — " + summary)
    if (Array.isArray(novel.chapters)) {
      novel.chapters = novel.chapters.map(ch => {
        const c = ch && typeof ch === 'object' ? ch : {};
        const limited = clampChapterLine(c.title, c.summary, 100);
        return {
          ...c,
          title: limited.title || safeStr(c.title),
          summary: limited.summary || '',
        };
      });
    }
  });
}

/** Dedupe / normalize character rows so names, roles, and genders stay consistent across the template. */
function synchronizeCharacterSystem(novel) {
  if (!novel || typeof novel !== 'object') return;
  const raw = Array.isArray(novel.characters) ? novel.characters : [];
  const byLower = new Map();
  const normRole = (r) => {
    const s = safeStr(r).toLowerCase();
    if (s.includes('protagonist') || s === 'main' || s.includes('lead')) return 'protagonist';
    if (s.includes('antagonist') || s.includes('villain')) return 'antagonist';
    return 'supporting';
  };
  const normGender = (g) => {
    const s = safeStr(g).toLowerCase();
    if (s === 'male' || s === 'm') return 'male';
    if (s === 'female' || s === 'f') return 'female';
    return 'female';
  };
  for (const c of raw) {
    if (!c || typeof c !== 'object') continue;
    const name = safeStr(c.name).trim();
    if (!name) continue;
    const key = name.toLowerCase();
    const merged = {
      name,
      role: normRole(c.role),
      age: safeStr(c.age) || '',
      description: safeStr(c.description) || '',
      arc: safeStr(c.arc) || '',
      gender: normGender(c.gender),
    };
    if (byLower.has(key)) {
      const prev = byLower.get(key);
      if (merged.description.length > prev.description.length) prev.description = merged.description;
      if (merged.arc.length > prev.arc.length) prev.arc = merged.arc;
      if (!prev.age && merged.age) prev.age = merged.age;
      continue;
    }
    byLower.set(key, merged);
  }
  const order = { protagonist: 0, antagonist: 1, supporting: 2 };
  novel.characters = [...byLower.values()].sort(
    (a, b) => (order[a.role] ?? 9) - (order[b.role] ?? 9) || a.name.localeCompare(b.name),
  );
}

function validateTemplateForAutoReview(novel) {
  const issues = [];
  if (!novel || typeof novel !== 'object') return { ok: false, issues: ['Invalid novel'], synced: true };
  const chars = novel.characters;
  if (!Array.isArray(chars) || chars.length < 2) {
    issues.push('Character system should list at least two characters for voice/export consistency.');
  }
  const seen = new Set();
  (chars || []).forEach((c) => {
    const n = safeStr(c?.name).trim();
    if (!n) {
      issues.push('A character entry is missing a name.');
      return;
    }
    const k = n.toLowerCase();
    if (seen.has(k)) issues.push(`Duplicate character name: "${n}".`);
    seen.add(k);
    if (!safeStr(c.description)) issues.push(`Character "${n}" has no description.`);
  });
  if (!safeStr(novel.draftScript) || novel.draftScript === 'N/A') {
    issues.push('Draft script / core ideas is empty.');
  }
  if (!Array.isArray(novel.chapters) || novel.chapters.length < 2) {
    issues.push('Chapter outline should have at least two chapters.');
  }
  if (isVietnameseWritingLanguage(novel.writingLanguage)) {
    (chars || []).forEach((c) => {
      const nm = safeStr(c?.name);
      if (!looksLikeSinoVietnameseName(nm)) {
        issues.push(`Character "${nm || '(empty)'}" should use Sino-Vietnamese Chinese-style naming (e.g. "Lâm Hải", "Lâm Vi").`);
      }
    });
  }
  const ov = safeStr(novel.overview);
  if (!ov || ov === 'N/A' || ov.length < 120) {
    issues.push('Overview should be a full-meaning summary (aim for 300+ characters).');
  }
  return { ok: issues.length === 0, issues, synced: true };
}

function applyAutoReviewTemplateToNovels(novels) {
  if (!Array.isArray(novels)) return;
  novels.forEach((novel) => {
    if (!novel || typeof novel !== 'object') return;
    enforceSinoVietnameseCharacterNames(novel);
    synchronizeCharacterSystem(novel);
    novel._autoReview = novel._autoReview || {};
    novel._autoReview.template = validateTemplateForAutoReview(novel);
  });
}

function formatAutoReviewIssuesHtml(issues) {
  if (!issues || !issues.length) return '<span class="auto-review-ok">No issues flagged.</span>';
  return `<ul class="auto-review-issues">${issues.map((x) => `<li>${escapeHtml(x)}</li>`).join('')}</ul>`;
}

function validateStoryAgainstTemplateProgrammatic(flowId, index) {
  const flow = getFlow(flowId);
  if (!flow) return { ok: false, issues: ['Missing flow'] };
  const novel = flow.novels[index];
  const storyText = safeStr(flow.stories[index]);
  const issues = [];
  if (!novel) return { ok: false, issues: ['Missing template'] };
  if (!storyText || storyText.length < 400) {
    issues.push('Story text is missing or too short.');
    return { ok: false, issues };
  }
  const outline = novel.chapters || [];
  const expectedNums = outline
    .map((c) => parseInt(c?.chapterNumber, 10))
    .filter((n) => !Number.isNaN(n) && n >= 1)
    .sort((a, b) => a - b);
  const expectedN = expectedNums.length;
  const parsed = parseChaptersFromMarkers(storyText);
  if (expectedN > 0 && parsed.length !== expectedN) {
    issues.push(
      `Chapter count mismatch: full story has ${parsed.length} [CHAPTER] block(s); template outline has ${expectedN}.`,
    );
  }
  if (parsed.length) {
    const parsedNums = parsed
      .map((ch) => parseInt(ch?.number, 10))
      .filter((n) => !Number.isNaN(n) && n >= 1)
      .sort((a, b) => a - b);
    const missing = expectedNums.filter((n) => !parsedNums.includes(n));
    if (missing.length) {
      issues.push(`Missing chapter block(s): ${missing.join(', ')}.`);
    }
    parsed.forEach((ch) => {
      const bodyLen = safeStr(ch?.content).length;
      if (bodyLen < 180) {
        issues.push(`Chapter ${ch.number} is too short (${bodyLen} chars).`);
      }
    });
  }
  const names = (novel.characters || []).map((c) => safeStr(c.name).trim()).filter(Boolean);
  const lower = storyText.toLowerCase();
  names.forEach((fullName) => {
    const tokens = fullName.split(/\s+/).filter((t) => t.length > 1);
    const hit = tokens.some((t) => lower.includes(t.toLowerCase()));
    if (!hit) issues.push(`Character "${fullName}" may be absent from the prose — verify manually.`);
  });
  return { ok: issues.length === 0, issues };
}

function updateStoryButtonReviewState(flowId, index) {
  const flow = getFlow(flowId);
  const novel = flow?.novels?.[index];
  const fd = flowDomId(flowId);
  const btn = document.getElementById(`storyBtn_${fd}_${index}`);
  if (!btn || !novel) return;
  const status = getNovelCompletionStatus(flowId, index, novel);
  const txt = btn.querySelector('.btn-text');
  if (txt) {
    if (status === 'generated finished') txt.textContent = '✅ Generated finished';
    else if (status === 'generated') txt.textContent = '✅ Story Generated';
    else txt.textContent = '📖 Generate Full Story';
  } else {
    if (status === 'generated finished') btn.innerHTML = '<span class="btn-text">✅ Generated finished</span>';
    else if (status === 'generated') btn.innerHTML = '<span class="btn-text">✅ Story Generated</span>';
    else btn.innerHTML = '<span class="btn-text">📖 Generate Full Story</span>';
  }
}

async function callStoryAlignmentReviewLLM(novel, storyText, apiKeyOverride) {
  const outline = (novel.chapters || [])
    .map((ch) => `Ch${ch.chapterNumber}: ${safeStr(ch.title)} — ${safeStr(ch.summary)}`)
    .join('\n');
  const chars = (novel.characters || [])
    .map((c) => `${safeStr(c.name)} (${safeStr(c.role)}): ${safeStr(c.description)}`)
    .join('\n');
  const excerpt = storyText.slice(0, 16000);
  const prompt = `You are a quality-control editor. Compare the story excerpt to the novel template.

Return JSON ONLY with this shape:
{"aligned":boolean,"character_sync_ok":boolean,"issues":string[],"summary":string}

Rules:
- "aligned" true if the story follows the same premise, setting, and chapter intent as the outline.
- "character_sync_ok" true if all listed characters appear to be the same people as in the template (names/roles).
- "issues" max 6 short strings; empty if none.
- "summary" one sentence.

Title: ${safeStr(novel.title)}
Overview: ${safeStr(novel.overview || novel.synopsis)}
Outline:
${outline}

Character system:
${chars || '(none)'}

Story excerpt:
${excerpt}`;

  if (getAIProvider() === 'deepseek') {
    return callDeepSeekAPI(prompt, true);
  }
  if (!apiKeyOverride && !getApiKey()) throw new Error('No API key for alignment review');
  return callGeminiAPIWithKey(prompt, apiKeyOverride, {
    temperature: 0.15,
    maxOutputTokens: 1200,
    responseMimeType: 'application/json',
  });
}

async function runAutoReviewStoryStep(flowId, index, apiKeyOverride) {
  const flow = getFlow(flowId);
  if (!flow) return;
  const novel = flow.novels[index];
  const storyText = flow.stories[index];
  if (!novel || !storyText) return;
  const prog = validateStoryAgainstTemplateProgrammatic(flowId, index);
  let llm = null;
  try {
    if (getApiKey()) {
      llm = await callStoryAlignmentReviewLLM(novel, storyText, apiKeyOverride);
    }
  } catch (e) {
    llm = { qc_error: safeStr(e?.message || e) };
  }
  const llmIssues = llm && !llm.qc_error && Array.isArray(llm.issues) ? llm.issues : [];
  const mergedIssues = [...(prog.issues || []), ...llmIssues];
  if (llm && llm.qc_error) mergedIssues.push(`Alignment review API: ${llm.qc_error}`);
  const llmPass =
    !llm || llm.qc_error
      ? true
      : llm.aligned !== false && llm.character_sync_ok !== false;
  novel._autoReview = novel._autoReview || {};
  novel._autoReview.story = {
    ok: prog.ok && llmPass && mergedIssues.length === 0,
    programmatic: prog,
    llm,
    issues: mergedIssues,
    summary: llm && safeStr(llm.summary),
    at: Date.now(),
  };
  const fd = flowDomId(flowId);
  const el = document.getElementById(`autoReviewStory_${fd}_${index}`);
  if (el) {
    const st = novel._autoReview.story;
    const qcErr = llm && llm.qc_error ? `<p class="auto-review-warn">${escapeHtml(llm.qc_error)}</p>` : '';
    el.className = `auto-review-banner ${st.ok ? 'ok' : 'warn'}`;
    el.innerHTML = `
      <div class="auto-review-title">🔍 Auto-review (full story vs template)</div>
      ${qcErr}
      ${st.summary ? `<p class="auto-review-summary">${escapeHtml(st.summary)}</p>` : ''}
      ${formatAutoReviewIssuesHtml(st.issues)}
    `;
    el.style.display = 'block';
  }
  updateStoryButtonReviewState(flowId, index);
}

function getExampleFullStory(index) {
  const samples = [
    // Shadows of the Empathic Order — full chapter content
    `[CHAPTER 1]
Title: The Awakening

Elara felt the pain before she saw the attacker. It was a sharp, cold blade of fear—not her own. She turned. In the alley, a man in Order colours was advancing on a child. Without thinking, she reached out. The pain flooded into her, and with it, a flash of memory: the Order knight, the cover-up, the lie. When she opened her eyes, the knight was on his knees. The child had fled. Elara ran.

She did not stop until she reached the river. There she knelt, hands in the water, and let the borrowed pain bleed out into the current. It had been like this since she was twelve: other people's feelings found her. Fear, grief, rage. She could pull them in or push them back, but she could not make them stop coming. The Order called it a gift. She had learned to call it a curse.

That night she dreamed of the knight again—not the alley, but years earlier. A room. A report. The words "acceptable losses" and a signature. When she woke, she knew the child had seen the same. Somewhere in the city, someone else now carried a piece of the truth. She did not know yet that the Order would come for her too.

[/CHAPTER 1]

[CHAPTER 2]
Title: The Order

They recruited her the next week. "You have a gift," the commander said. "Help us find the rogue. Help us end this." She took the mission. Find Kael. Bring him in. But every time she touched someone's pain, she saw more of the truth—and less of the Order's version.

Kael had been one of them. A knight. A believer. Then he had stumbled on the same report she had seen in the alley-child's memory. He had asked questions. They had called him rogue. Now they wanted him dead, and they wanted her to do the finding. She walked the districts where the displaced lived, the ones who had lost everything in the war. In every mind she brushed, she found the same thread: the Order had not been the victim. The Order had been the cause.

She sent her reports. She said nothing of what she had seen. At night she practiced holding the pain of others without letting it change her. She was not sure it was possible. By the time they gave her the location of Kael's last sighting, she had already decided she would hear him out before she decided whose side she was on.

[/CHAPTER 2]

[CHAPTER 3]
Title: The Truth

In the ruins where the war had started, she finally faced Kael. "They didn't tell you," he said. "They never tell anyone." The memories she had absorbed from a dozen victims aligned. The Order had started the war. She had been hunting the wrong enemy. Elara made her choice.

She did not draw her blade. She stood between him and the squad that had followed her, and she showed them what she had seen—not in words, but in feeling. She pushed the truth into their minds the way she had once pulled pain from the child in the alley. One by one they felt it: the report, the signatures, the lie. Some of them dropped their weapons. One of them ran. The commander did not. He looked at her and said, "Then you are rogue too."

She and Kael left the ruins before the Order could send reinforcements. They had no plan yet, only the certainty that the truth had to reach the rest of the city. "What do we do now?" Kael asked. Elara thought of the child, and the knight on his knees, and the river. "We find everyone who already knows," she said. "And we make sure they're not alone anymore."

[/CHAPTER 3]`,
    // Midnight at the Inkwell — full chapter content
    `[CHAPTER 1]
Title: The Contract

Maya signed the NDA and took the check. Julian Cross's Vermont estate was as cold as his reputation. "Finish the book," his agent said. "He gives you the outline; you make it sing." The outline was in code. The first page she decoded mentioned a real date, a real place—and a body.

She had thought the code was a quirk. Celebrity authors had quirks. But the more she worked through the cipher, the more the outline read like a confession. A party at the estate. A fall from the balcony. A cover-up that had lasted twenty years. She told herself it was research. That Julian wrote thrillers; of course his outlines were dark. She kept decoding. By the end of the week she had a timeline, names, and a sinking feeling that the book was not fiction.

She met Julian only once that month. He was gaunt, quiet, and he would not look at her when she asked about the outline. "Just write it," he said. "Write it the way it happened." That night she went back to her cottage on the grounds and stared at the decoded pages. The way it happened. She had signed the NDA. She had taken the check. If she walked away now, she would never work again. If she stayed, she had to decide what to do with the truth.

[/CHAPTER 1]

[CHAPTER 2]
Title: The Manuscript

Coded pages kept pointing to the same night: a party, a fall, a cover-up. She interviewed the staff. One of them had seen something. That night, someone left a note on her pillow: Stop asking. She didn't.

The note was typed. No fingerprints. She started locking her door and keeping the decoded manuscript in a bag she never left unattended. The staff had been at the estate for years; loyalty ran deep. But one of them—an older groundskeeper—had looked at her with something like pity when she asked about the balcony. "Some things are better left in the past," he said. She found him again the next day. He was gone. No forwarding address. No one would say where he had gone.

She kept decoding. The manuscript described a woman who had come to the party uninvited. A confrontation. A push. The body had been found in the garden, not the balcony—the outline had been wrong about that, or someone had moved it. She cross-referenced dates. The party had been twenty-two years ago. Julian's first bestseller had come out a year later. The book had been about a writer who got away with murder. She went back to the main house and asked to see Julian again. His agent said he was not well. Maya said she had questions only he could answer. She was still waiting for the meeting when the body turned up in the garden.

[/CHAPTER 2]

[CHAPTER 3]
Title: The Murder

A body turned up in the same spot the manuscript described. The police asked questions. Julian finally talked. "I didn't write it as fiction," he said. "I wrote it as confession." To finish the book and stay alive, Maya had to piece together the story—and decide who to trust.

The body was the groundskeeper. He had been dead for two days. The police treated Maya as a witness. She gave them the decoded manuscript and told them everything she had found. Julian was arrested. His agent was arrested. The story made the front page. Maya's name was in the byline—not as the ghostwriter, but as the one who had broken the case. The publisher still wanted the book. They wanted her to write it, under her own name this time. The truth, they said, would sell.

She went back to Vermont once, after the trial. The estate was empty. She stood in the garden where the groundskeeper had been found and thought about the woman from the party, the one in the manuscript. She had never been identified. Maybe she had never existed. Maybe Julian had made her up to explain the body. Or maybe she was still out there, and the manuscript had been a message. Maya had decoded it. She had told the story. Some nights she wondered if that made her the next target—or the last one who could still choose what happened next. She sat down at her desk and started writing. Not the book they wanted. The one she needed to tell.

[/CHAPTER 3]`
  ];
  return samples[index] || '';
}

function loadExampleTemplates() {
  try {
    const flow = ensureActiveFlow();
    flow.label = 'Example templates';
    const fid = flow.id;
    flow.novels = getExampleTemplates();
    normalizeNovelsForExport(flow.novels);
    stampCollectionAndCategoriesFromForm(flow.novels);
    applyAutoReviewTemplateToNovels(flow.novels);
    flow.stories = {};
    flow.novels.forEach((_, index) => {
      const sample = getExampleFullStory(index);
      if (sample) flow.stories[index] = sample;
    });
    ensureFlowPanelDom(fid);
    switchToFlow(fid);
    const fd = flowDomId(fid);
    const section = document.getElementById('resultsSection');
    const countEl = document.getElementById('resultsCount');
    const container = document.getElementById(`novelsContainer_${fd}`);
    if (!section || !container) {
      showToast('Could not render examples. Please refresh and try again.', 'error');
      return;
    }
    section.classList.add('active');
    if (countEl) countEl.textContent = `${flow.novels.length} example templates`;
    container.innerHTML = '';
    flow.novels.forEach((novel, index) => {
      const card = createNovelCard(novel, index, fid);
      card.style.animationDelay = `${index * 0.1}s`;
      container.appendChild(card);
    });
    attachEditSyncListeners(container, fid);
    flow.novels.forEach((_, index) => {
      const storyText = flow.stories[index];
      if (storyText) {
        const storySection = document.getElementById(`storySection_${fd}_${index}`);
        const storyContent = document.getElementById(`storyContent_${fd}_${index}`);
        if (storySection && storyContent) {
          renderStoryChapters(fid, index, storyText);
          storySection.style.display = 'block';
        }
        const storyBtn = document.getElementById(`storyBtn_${fd}_${index}`);
        if (storyBtn) storyBtn.innerHTML = '<span class="btn-text">✅ Story Generated</span>';
      }
    });
    const firstCard = container?.querySelector('.novel-card');
    if (firstCard) firstCard.classList.add('expanded');
    section.scrollIntoView({ behavior: 'smooth', block: 'start' });
    void (async () => {
      for (let i = 0; i < flow.novels.length; i++) {
        if (!flow.stories[i]) continue;
        try {
          await runAutoReviewStoryStep(fid, i, null);
        } catch (e) {
          console.warn('Example auto-review story', i, e);
        }
      }
    })();
    if (canGenerateImages()) {
      showToast('Example templates loaded. Generating thumbnails…', 'success');
      generateCoversForAllTemplates(fid);
    } else {
      showToast('Example templates loaded. Configure image generation to create thumbnails.', 'success');
    }
  } catch (e) {
    console.error('Load example failed:', e);
    showToast('Load example failed: ' + (e?.message || 'Unknown error'), 'error');
  }
}

// --- Small helpers ---
function safeStr(v) {
  if (v == null) return '';
  if (Array.isArray(v)) return v.map(safeStr).filter(Boolean).join(', ');
  if (typeof v === 'object') return '';
  return String(v).trim();
}

function flowLabelFromPrompt(masterPrompt) {
  const s = safeStr(masterPrompt).replace(/\s+/g, ' ').trim();
  if (!s) return 'New run';
  return s.length > 42 ? s.slice(0, 40) + '…' : s;
}

function clampText(s, maxChars) {
  const str = safeStr(s);
  if (!maxChars || maxChars < 1) return str;
  if (str.length <= maxChars) return str;
  return str.slice(0, Math.max(0, maxChars - 1)).trimEnd() + '…';
}

// --- Category constraints (must match allowed app categories) ---
const ALLOWED_CATEGORIES = [
  'Romance',
  'Werewolf',
  'Mafia',
  'System',
  'Fantasy',
  'Urban',
  'LGBTQ+',
  'YA/TEEN',
  'Paranormal',
  'Mystery/Thriller',
  'Eastern',
  'Games',
  'History',
  'MM Romance',
  'Sci-Fi',
  'War',
  'Emotional Realism',
  'Vampire',
  'Campus',
  'Imagination',
  'Rebirth',
  'Steamy',
  'Folklore Mystery',
  'Male POV',
  'New Adult',
  'Paranormal Urban',
  'Action',
  'Chicklit',
  'Other',
];

function normalizeCategoryName(raw) {
  const s = safeStr(raw);
  if (!s) return '';
  const lower = s.toLowerCase();
  const byLower = new Map(ALLOWED_CATEGORIES.map(c => [c.toLowerCase(), c]));
  if (byLower.has(lower)) return byLower.get(lower);
  // Common variants / fuzzy mapping
  if (lower.includes('sci') || lower.includes('science')) return 'Sci-Fi';
  if (lower.includes('thriller') || lower.includes('mystery') || lower.includes('suspense') || lower.includes('crime')) return 'Mystery/Thriller';
  if (lower.includes('ya') || lower.includes('teen') || lower.includes('young adult')) return 'YA/TEEN';
  if (lower.includes('lgbt')) return 'LGBTQ+';
  if (lower.includes('mm') || (lower.includes('male') && lower.includes('male'))) return 'MM Romance';
  if (lower.includes('paranormal')) return 'Paranormal';
  if (lower.includes('urban')) return 'Urban';
  if (lower.includes('werewolf') || lower.includes('wolf') || lower.includes('lycan')) return 'Werewolf';
  if (lower.includes('mafia') || lower.includes('gang')) return 'Mafia';
  if (lower.includes('romance')) return 'Romance';
  if (lower.includes('fantasy')) return 'Fantasy';
  if (lower.includes('system')) return 'System';
  if (lower.includes('game')) return 'Games';
  if (lower.includes('history') || lower.includes('historical')) return 'History';
  if (lower.includes('war') || lower.includes('military')) return 'War';
  if (lower.includes('eastern') || lower.includes('wuxia') || lower.includes('xianxia') || lower.includes('china') || lower.includes('korea') || lower.includes('japan')) return 'Eastern';
  if (lower.includes('vampire')) return 'Vampire';
  if (lower.includes('campus') || lower.includes('college') || lower.includes('university')) return 'Campus';
  if (lower.includes('rebirth') || lower.includes('reincarnation') || lower.includes('transmigration')) return 'Rebirth';
  if (lower.includes('steamy') || lower.includes('spicy') || lower.includes('explicit')) return 'Steamy';
  if (lower.includes('folklore') && lower.includes('mystery')) return 'Folklore Mystery';
  if (lower.includes('male pov') || lower.includes('male point of view')) return 'Male POV';
  if (lower.includes('emotional realism')) return 'Emotional Realism';
  if (lower.includes('imagination')) return 'Imagination';
  if (lower.includes('new adult')) return 'New Adult';
  if (lower.includes('paranormal urban')) return 'Paranormal Urban';
  if (lower.includes('action')) return 'Action';
  if (lower.includes('chicklit') || lower.includes('chick lit')) return 'Chicklit';
  return 'Other';
}

function clampChapterLine(title, summary, maxChars = 100) {
  const t = safeStr(title);
  const s = safeStr(summary);
  if (!t && !s) return { title: '', summary: '' };
  if (!s) return { title: clampText(t, maxChars), summary: '' };
  const sep = ' — ';
  const base = t ? (t + sep) : '';
  const remaining = Math.max(0, maxChars - base.length);
  if (remaining <= 0) return { title: clampText(t, maxChars), summary: '' };
  return { title: t, summary: clampText(s, remaining) };
}

function limitWords(s, maxWords) {
  const words = safeStr(s).split(/\s+/).filter(Boolean);
  return words.slice(0, Math.max(1, maxWords)).join(' ');
}

// Reject vague/nonsense tag patterns (e.g. "corruption of", "identity vs", "X of", "X vs Y").
function isMeaningfulTag(tag) {
  const s = safeStr(tag).trim();
  if (s.length < 2) return false;
  const lower = s.toLowerCase();
  // Drop phrases ending in " of", " vs", " versus" (incomplete/vague).
  if (/\s+(of|vs\.?|versus)\s*$/.test(lower)) return false;
  // Drop "X of Y" or "X vs Y" when they're generic (only 2–3 words and contain of/vs).
  if (/^.+\s+(of|vs\.?|versus)\s+.+$/.test(lower) && s.split(/\s+/).length <= 3) return false;
  // Drop very generic standalone words that add no meaning.
  const skip = new Set(['the', 'and', 'or', 'vs', 'versus', 'of']);
  const words = lower.split(/\s+/).filter(Boolean);
  if (words.length === 1 && skip.has(words[0])) return false;
  return true;
}

function normalizeTagList(list) {
  const out = (Array.isArray(list) ? list : [])
    .map(t => limitWords(t, 2))
    .map(t => t.replace(/[|—–-]+/g, ' ').replace(/\s+/g, ' ').trim())
    .filter(Boolean)
    .filter(isMeaningfulTag);
  const seen = new Set();
  return out.filter(t => {
    const k = t.toLowerCase();
    if (seen.has(k)) return false;
    seen.add(k);
    return true;
  });
}

// Random author name when none provided (avoid Unknown Author / Anonymous).
const RANDOM_FIRST = ['Emma', 'Marcus', 'Sofia', 'James', 'Luna', 'Alex', 'Mia', 'Leo', 'Zara', 'Noah', 'Ella', 'Finn', 'Ivy', 'Owen', 'Ruby', 'Cole', 'Lily', 'Jake', 'Nina', 'Kai'];
const RANDOM_LAST = ['Vance', 'Webb', 'Cross', 'Reed', 'Blake', 'Hart', 'Shaw', 'Gray', 'Wells', 'Fox', 'Brooks', 'Stone', 'Lane', 'Cole', 'Hayes', 'Marsh', 'Kent', 'Burns', 'Ford', 'Page'];

function pickRandomAuthorName() {
  const first = RANDOM_FIRST[Math.floor(Math.random() * RANDOM_FIRST.length)];
  const last = RANDOM_LAST[Math.floor(Math.random() * RANDOM_LAST.length)];
  return `${first} ${last}`;
}

function isVietnameseWritingLanguage(lang) {
  const s = safeStr(lang).toLowerCase();
  return s.includes('vietnamese') || s.includes('tiếng việt') || s === 'vietnam';
}

const SINO_VIET_SURNAMES = [
  'Lâm', 'Trần', 'Lý', 'Triệu', 'Tô', 'Tạ', 'Chu', 'Hà', 'Vương', 'Dương', 'Tần', 'Mạc', 'Tống', 'Lục', 'Thẩm', 'Bạch', 'Phong',
  'Lam', 'Tran', 'Ly', 'Trieu', 'To', 'Ta', 'Chu', 'Ha', 'Vuong', 'Duong', 'Tan', 'Mac', 'Tong', 'Luc', 'Tham', 'Bach', 'Phong',
];
const SINO_VIET_GIVEN = [
  'Hải', 'Vi', 'Vân', 'Nhi', 'Nguyệt', 'Mộng', 'Uyên', 'Thanh', 'Tâm', 'Yên', 'Diệp', 'Kha', 'Mặc', 'Thiên', 'Tử', 'Kiệt', 'Dật', 'Phi', 'Sương', 'Lạc',
  'Hai', 'Vi', 'Van', 'Nhi', 'Nguyet', 'Mong', 'Uyen', 'Thanh', 'Tam', 'Yen', 'Diep', 'Kha', 'Mac', 'Thien', 'Tu', 'Kiet', 'Dat', 'Phi', 'Suong', 'Lac',
];

function looksLikeSinoVietnameseName(name) {
  const n = safeStr(name).replace(/\s+/g, ' ').trim();
  if (!n) return false;
  const parts = n.split(' ').filter(Boolean);
  if (parts.length < 2 || parts.length > 3) return false;
  const first = parts[0];
  let firstAscii = first;
  try {
    firstAscii = first.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  } catch (_) {
    firstAscii = first;
  }
  return SINO_VIET_SURNAMES.includes(first) || SINO_VIET_SURNAMES.includes(firstAscii);
}

function makeSinoVietnameseName(seed, usedSet) {
  for (let i = 0; i < 100; i++) {
    const s = SINO_VIET_SURNAMES[(seed + i) % SINO_VIET_SURNAMES.length];
    const g1 = SINO_VIET_GIVEN[(seed * 3 + i) % SINO_VIET_GIVEN.length];
    const g2 = SINO_VIET_GIVEN[(seed * 7 + i + 5) % SINO_VIET_GIVEN.length];
    const candidate = `${s} ${g1}` + ((seed + i) % 3 === 0 ? ` ${g2}` : '');
    const key = candidate.toLowerCase();
    if (!usedSet.has(key)) {
      usedSet.add(key);
      return candidate;
    }
  }
  const fallback = `Lâm Vi ${seed + 1}`;
  usedSet.add(fallback.toLowerCase());
  return fallback;
}

function enforceSinoVietnameseCharacterNames(novel) {
  if (!novel || !isVietnameseWritingLanguage(novel.writingLanguage)) return;
  const chars = Array.isArray(novel.characters) ? novel.characters : [];
  const used = new Set(chars.map((c) => safeStr(c?.name).toLowerCase()).filter(Boolean));
  chars.forEach((c, idx) => {
    if (!c || typeof c !== 'object') return;
    const current = safeStr(c.name);
    if (looksLikeSinoVietnameseName(current)) return;
    c.name = makeSinoVietnameseName(idx + 11, used);
  });
}

function getExportCollection() {
  return safeStr(document.getElementById('collectionName')?.value);
}

function getCategoriesForExport(novel) {
  const fixed = safeStr(novel?.cateogories || novel?.categories);
  if (fixed) return fixed;
  const cat = safeStr(novel?.category);
  const genre = safeStr(novel?.genre);
  if (cat && genre && cat !== genre) return `${cat} | ${genre}`;
  return cat || genre || '';
}

function getTagsForExport(novel) {
  if (Array.isArray(novel?.tags) && novel.tags.length) return normalizeTagList(novel.tags).join(', ');
  if (Array.isArray(novel?.themes) && novel.themes.length) return normalizeTagList(novel.themes).join(', ');
  const single = safeStr(novel?.tag);
  return single ? normalizeTagList([single]).join(', ') : '';
}

/** Comma-separated tags for card display / editing */
function getTagsDisplayText(novel) {
  if (Array.isArray(novel?.tags) && novel.tags.length) return normalizeTagList(novel.tags).join(', ');
  if (Array.isArray(novel?.themes) && novel.themes.length) return normalizeTagList(novel.themes).join(', ');
  const single = safeStr(novel?.tag);
  return single || 'N/A';
}

/** Chapter outline as plain text (for export columns): "Ch1: Title — Summary\nCh2: ..." */
function getChapterOutlineText(novel) {
  if (!novel?.chapters || !novel.chapters.length) return '';
  return novel.chapters.map(ch =>
    `Ch${ch.chapterNumber}: ${safeStr(ch.title)} — ${safeStr(ch.summary)}`
  ).join('\n');
}

/** Full story text for a novel (from DOM or state), for export and template .txt */
function getFullStoryText(flowId, index) {
  const fd = flowDomId(flowId);
  const contentEl = document.getElementById(`storyContent_${fd}_${index}`);
  const fromDom = contentEl?.textContent?.trim();
  if (fromDom) return fromDom;
  const flow = getFlow(flowId);
  return safeStr(flow?.stories[index]);
}

/** Parsed chapters with full content for export. Uses story with [CHAPTER N] markers when available for accurate split. */
function getFullStoryChaptersForExport(flowId, index) {
  const flow = getFlow(flowId);
  const rawFromState = safeStr(flow?.stories[index]);
  const fd = flowDomId(flowId);
  const fromDom = document.getElementById(`storyContent_${fd}_${index}`)?.textContent?.trim();
  const textToParse = rawFromState || fromDom || '';
  let chapters = parseChaptersFromMarkers(textToParse);
  if (!chapters.length && textToParse) {
    chapters = [{ number: 1, title: '', content: textToParse }];
  }
  return chapters;
}

function storyIsCompleteForNovel(flowId, index, novel) {
  if (!novel || typeof novel !== 'object') return false;
  const expected = (Array.isArray(novel.chapters) ? novel.chapters : [])
    .map(c => parseInt(c?.chapterNumber, 10))
    .filter(n => !Number.isNaN(n) && n >= 1);
  if (!expected.length) return false;
  const parsed = getFullStoryChaptersForExport(flowId, index);
  const contentByNum = new Map(parsed.map(ch => [parseInt(ch.number, 10), safeStr(ch.content)]));
  // Require every expected chapter number to exist and have some prose.
  return expected.every(n => {
    const txt = contentByNum.get(n) || '';
    return txt.length >= 200; // "full chapter" heuristic
  });
}

function getNovelCompletionStatus(flowId, index, novel) {
  const reviewedOk = !!(novel?._autoReview?.story?.ok);
  if (reviewedOk) return 'generated finished';
  return storyIsCompleteForNovel(flowId, index, novel) ? 'generated' : 'draft';
}

/** Export column order (one row per chapter/episode) */
const EXPORT_HEADERS = [
  'title',
  'author',
  'summary',
  'tags',
  'completionStatus',
  'totalChapter',
  'thumbnailUrl',
  'category',
  'chapterTitle',
  'chapterContent',
];

/** Build one row per chapter for a novel. Each row has same novel metadata + one chapter_outline and chapter_content. */
function buildExportRowsForNovel(flowId, novelIndex, novel, collection) {
  const templateChapters = novel?.chapters || [];
  const storyChapters = getFullStoryChaptersForExport(flowId, novelIndex);
  const outlineByNum = {};
  templateChapters.forEach(ch => {
    // Export expects only the episode/chapter title (no summary/detail).
    outlineByNum[ch.chapterNumber] = safeStr(ch.title);
  });
  const contentByNum = {};
  storyChapters.forEach(ch => {
    contentByNum[ch.number] = safeStr(ch.content);
    if (!outlineByNum[ch.number])
      outlineByNum[ch.number] = safeStr(ch.title) || `Chapter ${ch.number}`;
  });
  const allNums = [...new Set([...Object.keys(outlineByNum).map(Number), ...Object.keys(contentByNum).map(Number)])].sort((a, b) => a - b);
  const rows = [];
  const title = safeStr(novel.title);
  const author = safeStr(novel.authorName || novel.author) || safeStr(document.getElementById('authorName')?.value) || 'Anonymous';
  const summary = clampText(safeStr(novel.overview || novel.synopsis), 500);
  const tags = getTagsForExport(novel);
  const completionStatus = getNovelCompletionStatus(flowId, novelIndex, novel);
  const totalChapter = allNums.length || (Array.isArray(templateChapters) ? templateChapters.length : 0);
  const thumbnailUrl = safeStr(novel.thumbnail || novel.cover || '');
  const category = getCategoriesForExport(novel);

  if (allNums.length === 0) {
    rows.push([title, author, summary, tags, completionStatus, totalChapter, thumbnailUrl, category, '', '']);
    return rows;
  }
  allNums.forEach(num => {
    rows.push([
      title,
      author,
      summary,
      tags,
      completionStatus,
      totalChapter,
      thumbnailUrl,
      category,
      outlineByNum[num] || `Chapter ${num}`,
      contentByNum[num] || '',
    ]);
  });
  return rows;
}

function csvEscape(v) {
  const s = safeStr(v);
  if (s.includes('"') || s.includes(',') || s.includes('\n') || s.includes('\r')) return `"${s.replace(/"/g, '""')}"`;
  return s;
}

function downloadBlob(blob, filename) {
  const a = document.createElement('a');
  const url = URL.createObjectURL(blob);
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1500);
}

function dataUrlToBase64Info(dataUrl) {
  const m = /^data:image\/([a-zA-Z0-9+.-]+);base64,(.+)$/.exec(dataUrl || '');
  if (!m) return null;
  let ext = m[1].toLowerCase();
  if (ext === 'jpeg') ext = 'jpg';
  if (!['png', 'jpg', 'gif', 'webp', 'bmp'].includes(ext)) ext = 'png';
  return { extension: ext, base64: dataUrl };
}

function resizeDataUrl(dataUrl, maxW, maxH, outType = 'image/png') {
  return new Promise((resolve) => {
    if (!dataUrl) return resolve('');
    const img = new Image();
    img.onload = () => {
      const w = img.width || 1;
      const h = img.height || 1;
      const scale = Math.min(maxW / w, maxH / h, 1);
      const tw = Math.max(1, Math.round(w * scale));
      const th = Math.max(1, Math.round(h * scale));
      const canvas = document.createElement('canvas');
      canvas.width = tw;
      canvas.height = th;
      const ctx = canvas.getContext('2d');
      ctx.imageSmoothingEnabled = true;
      ctx.imageSmoothingQuality = 'high';
      ctx.drawImage(img, 0, 0, tw, th);
      try {
        resolve(canvas.toDataURL(outType));
      } catch (_) {
        resolve(dataUrl);
      }
    };
    img.onerror = () => resolve('');
    img.src = dataUrl;
  });
}

function pickCoverDataUrl(flowId, novelIndex, novel) {
  const maybe = novel?.cover || novel?.coverImage || novel?.image || novel?.thumbnail;
  if (typeof maybe === 'string' && maybe.startsWith('data:image/')) return maybe;
  const flow = getFlow(flowId);
  const scenes = flow?.generatedScenes?.[novelIndex];
  if (!scenes) return '';
  const keys = Object.keys(scenes)
    .map(k => parseInt(k, 10))
    .filter(n => !Number.isNaN(n))
    .sort((a, b) => a - b);
  if (!keys.length) return '';
  const first = scenes[keys[0]];
  return (typeof first === 'string' && first.startsWith('data:image/')) ? first : '';
}

function pickThumbnailDataUrl(flowId, novelIndex, novel) {
  const maybe = novel?.thumbnail || novel?.thumb || novel?.cover;
  if (typeof maybe === 'string' && maybe.startsWith('data:image/')) return maybe;
  // Fall back to cover/scenes-derived image if present.
  return pickCoverDataUrl(flowId, novelIndex, novel);
}

// Ensure legacy callers (inline handlers / older exports) can always access it.
try { window.pickThumbnailDataUrl = pickThumbnailDataUrl; } catch (_) {}

/** Before export: ensure every novel has cover + thumbnail so thumbnail/cover columns are filled. */
async function ensureThumbnailsForExport() {
  const flow = ensureActiveFlow();
  const novels = flow.novels || [];
  if (!Array.isArray(novels) || !novels.length) return;
  if (!canGenerateImages()) {
    // Ensure prompt fallback exists even without images
    normalizeNovelsForExport(novels);
    showToast('Image generation is not available; exporting without thumbnails.', 'info');
    return;
  }
  const missing = novels
    .map((n, i) => (typeof n?.cover === 'string' && n.cover.startsWith('data:image/')) ? -1 : i)
    .filter(i => i >= 0);
  if (!missing.length) return;
  showToast(`Generating ${missing.length} missing thumbnail(s) in parallel…`, 'info');
  const queue = missing.slice();
  const concurrency = Math.min(4, queue.length);
  let done = 0;
  const total = missing.length;
  const fid = state.activeFlowId;
  const worker = async () => {
    while (queue.length) {
      const i = queue.shift();
      if (i == null) break;
      try {
        await generateCoverForNovel(fid, i);
        done++;
      } catch (e) {
        console.warn('Thumbnail generation failed for novel', i, e);
        showToast(`Thumbnail generation failed for novel ${i + 1}: ${e?.message || 'Unknown error'}`, 'error');
      }
      await new Promise(r => setTimeout(r, 60));
    }
  };
  await Promise.all(Array.from({ length: concurrency }, () => worker()));
  showToast(`Thumbnails ready for export (${done}/${total}).`, 'success');
}

// --- Export: CSV / XLSX ---
async function handleExportCsv() {
  try {
    const flow = ensureActiveFlow();
    if (!Array.isArray(flow.novels) || !flow.novels.length) {
      showToast('Nothing to export yet. Generate novel templates first.', 'error');
      return;
    }
    const flowId = state.activeFlowId;
    normalizeNovelsForExport(flow.novels);
    const collection = getExportCollection();
    const lines = [EXPORT_HEADERS.map(csvEscape).join(',')];
    for (let i = 0; i < flow.novels.length; i++) {
      const novel = flow.novels[i] || {};
      const rows = buildExportRowsForNovel(flowId, i, novel, collection);
      rows.forEach(row => lines.push(row.map(csvEscape).join(',')));
    }
    const csv = lines.join('\r\n');
    downloadBlob(new Blob(['\uFEFF' + csv], { type: 'text/csv;charset=utf-8' }), 'novels_export.csv');
    showToast('Exported CSV (use Export package .zip for image files)', 'success');
  } catch (e) {
    console.error('CSV export failed', e);
    showToast('CSV export failed: ' + (e?.message || String(e)), 'error');
  }
}

async function handleExportZipPackage() {
  try {
    const flow = ensureActiveFlow();
    if (!Array.isArray(flow.novels) || !flow.novels.length) {
      showToast('Nothing to export yet. Generate novel templates first.', 'error');
      return;
    }
    const flowId = state.activeFlowId;
    normalizeNovelsForExport(flow.novels);
    if (!window.JSZip) {
      showToast('ZIP export library failed to load. Reload and try again.', 'error');
      return;
    }
    await ensureThumbnailsForExport();

    for (let i = 0; i < flow.novels.length; i++) {
      const novel = flow.novels[i] || {};
      if (!novel.thumbnail && novel.cover) {
        novel.thumbnail = await resizeDataUrl(novel.cover, 128, 128, 'image/png');
      }
    }

    const zip = new JSZip();
    const thumbs = zip.folder('thumbnails');
    const covers = zip.folder('covers');

    const collection = getExportCollection();
    const csvLines = [EXPORT_HEADERS.map(csvEscape).join(',')];

    const galleryCards = [];

    for (let i = 0; i < flow.novels.length; i++) {
      const novel = flow.novels[i] || {};
      const coverDataUrl = pickCoverDataUrl(flowId, i, novel);
      const thumbDataUrl = novel.thumbnail || (coverDataUrl ? await resizeDataUrl(coverDataUrl, 128, 128, 'image/png') : '');

      const coverPng = coverDataUrl ? await fetch(coverDataUrl).then(r => r.blob()) : null;
      const thumbPng = thumbDataUrl ? await fetch(thumbDataUrl).then(r => r.blob()) : null;

      const coverName = `novel_${i + 1}.png`;
      const thumbName = `novel_${i + 1}.png`;
      if (coverPng) covers.file(coverName, coverPng);
      if (thumbPng) thumbs.file(thumbName, thumbPng);

      const rows = buildExportRowsForNovel(flowId, i, novel, collection);
      rows.forEach(row => csvLines.push(row.map(csvEscape).join(',')));

      galleryCards.push(`
        <a class="card" href="covers/${coverName}" target="_blank" rel="noopener">
          <img src="thumbnails/${thumbName}" alt="${escapeHtml(safeStr(novel.title) || ('Novel ' + (i + 1)))}"/>
          <div class="t">${escapeHtml(safeStr(novel.title) || ('Novel ' + (i + 1)))}</div>
        </a>
      `);
    }

    zip.file('templates.csv', '\uFEFF' + csvLines.join('\r\n'));
    zip.file('gallery.html', `<!doctype html>
<html>
<head>
  <meta charset="utf-8"/>
  <meta name="viewport" content="width=device-width, initial-scale=1"/>
  <title>Novel thumbnails</title>
  <style>
    body{font-family:system-ui,-apple-system,Segoe UI,Roboto,Arial,sans-serif;background:#0b1020;color:#e8ebf4;margin:0;padding:24px}
    h1{margin:0 0 16px;font-size:18px}
    .grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(160px,1fr));gap:14px}
    .card{display:block;text-decoration:none;color:inherit;border:1px solid rgba(255,255,255,0.12);border-radius:14px;overflow:hidden;background:rgba(255,255,255,0.04)}
    .card img{width:100%;height:160px;object-fit:cover;display:block}
    .card .t{padding:10px 10px;font-size:13px;line-height:1.35;opacity:.92}
  </style>
</head>
<body>
  <h1>Novel thumbnails (click to open cover)</h1>
  <div class="grid">
    ${galleryCards.join('')}
  </div>
</body>
</html>`);

    const blob = await zip.generateAsync({ type: 'blob' });
    downloadBlob(blob, 'novel_templates_package.zip');
    showToast('Exported package (.zip)', 'success');
  } catch (e) {
    console.error('ZIP export failed', e);
    showToast('ZIP export failed: ' + (e?.message || String(e)), 'error');
  }
}

function getExcelCellText(value) {
  if (value == null) return '';
  if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') return String(value);
  if (value.richText && Array.isArray(value.richText)) {
    return value.richText.map((part) => part?.text || '').join('');
  }
  if (value.text) return String(value.text);
  if (value.result != null) return String(value.result);
  return String(value);
}

function autoFitWorksheetColumns(ws, minWidth = 12, maxWidth = 80) {
  if (!ws || !Array.isArray(ws.columns)) return;
  ws.columns.forEach((col) => {
    let maxLen = minWidth;
    col.eachCell({ includeEmpty: true }, (cell) => {
      const text = getExcelCellText(cell?.value).replace(/\r?\n/g, ' ');
      if (!text) return;
      if (text.length > maxLen) maxLen = text.length;
    });
    col.width = Math.max(minWidth, Math.min(maxWidth, maxLen + 2));
  });
}

async function handleExportXlsx(options = {}) {
  const { fitColumns = false } = options || {};
  try {
    const flow = ensureActiveFlow();
    if (!Array.isArray(flow.novels) || !flow.novels.length) {
      showToast('Nothing to export yet. Generate novel templates first.', 'error');
      return;
    }
    const flowId = state.activeFlowId;
    normalizeNovelsForExport(flow.novels);
    if (!window.ExcelJS) {
      showToast('XLSX export library failed to load. Reload and try again.', 'error');
      return;
    }
    const collection = getExportCollection();
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Novels');
    ws.columns = EXPORT_HEADERS.map((h) => ({
      header: h,
      key: h,
      width: h === 'chapterContent' || h === 'summary' ? 60 : h === 'chapterTitle' ? 48 : 20,
    }));
    ws.getRow(1).font = { bold: true };
    ws.getRow(1).alignment = { vertical: 'middle' };
    ws.getRow(1).height = 20;

    let rowNumber = 2;
    for (let i = 0; i < flow.novels.length; i++) {
      const novel = flow.novels[i] || {};
      const rows = buildExportRowsForNovel(flowId, i, novel, collection);
      rows.forEach(row => {
        const rowObj = {};
        EXPORT_HEADERS.forEach((key, idx) => { rowObj[key] = row[idx]; });
        ws.addRow(rowObj);
        ws.getRow(rowNumber).alignment = { vertical: 'top', wrapText: true };
        ws.getRow(rowNumber).height = 24;
        rowNumber++;
      });
    }

    ws.views = [{ state: 'frozen', ySplit: 1 }];
    if (fitColumns) {
      autoFitWorksheetColumns(ws, 10, 80);
    }
    const buf = await wb.xlsx.writeBuffer();
    downloadBlob(
      new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }),
      fitColumns ? 'novels_export_fit_columns.xlsx' : 'novels_export.xlsx'
    );
    showToast(fitColumns ? 'Exported XLSX (fit columns)' : 'Exported XLSX', 'success');
  } catch (e) {
    console.error('XLSX export failed', e);
    showToast('XLSX export failed: ' + (e?.message || String(e)), 'error');
  }
}

// --- File Upload ---
function handleFileUpload(file) {
  if (!file.name.endsWith('.txt')) {
    showToast('Please upload a .txt file', 'error');
    return;
  }
  const reader = new FileReader();
  reader.onload = (e) => {
    document.getElementById('referenceText').value = e.target.result;
    document.getElementById('uploadFileName').textContent = file.name;
    document.getElementById('uploadFileName').style.display = 'block';
    showToast(`Loaded reference: ${file.name}`, 'success');
  };
  reader.readAsText(file);
}

// --- Settings modal ---
function openSettings() {
  document.getElementById('settingsModal')?.classList.add('active');
}
function closeSettings() {
  document.getElementById('settingsModal')?.classList.remove('active');
}
function saveSettings() {
  const ttsProvider = document.getElementById('ttsProvider')?.value || 'gemini';
  localStorage.setItem('tts_provider', ttsProvider);
  const ttsKeys = document.getElementById('ttsApiKey')?.value || '';
  localStorage.setItem('tts_api_keys', ttsKeys);
  if (ttsKeys) localStorage.setItem('gemini_tts_key', ttsKeys.split(/[\s,;\n]+/)[0]?.trim() || '');
  const ai33Key = document.getElementById('ai33ApiKey')?.value || '';
  localStorage.setItem('ai33_api_key', ai33Key);
  const ai33Url = document.getElementById('ai33BaseUrl')?.value || '';
  localStorage.setItem('ai33_base_url', ai33Url || 'https://api.ai33.pro/v1');
  updateApiStatusBadge();
  showToast('Settings saved', 'success');
}
function updateTtsProviderUI() {
  const tts = document.getElementById('ttsProvider')?.value || 'gemini';
  const geminiGroup = document.getElementById('geminiTtsKeyGroupMain');
  const ai33Group = document.getElementById('ai33SettingsGroup');
  const ai33UrlGroup = document.getElementById('ai33BaseUrlGroup');
  if (geminiGroup) geminiGroup.style.display = tts === 'gemini' ? 'block' : 'none';
  if (ai33Group) ai33Group.style.display = tts === 'ai33pro' ? 'block' : 'none';
  if (ai33UrlGroup) ai33UrlGroup.style.display = tts === 'ai33pro' ? 'block' : 'none';
  ['narratorVoice', 'femaleVoice', 'maleVoice'].forEach(id => {
    const sel = document.getElementById(id);
    if (!sel) return;
    const target = tts === 'ai33pro' ? 'ai33' : 'gemini';
    Array.from(sel.querySelectorAll('optgroup')).forEach(grp => {
      grp.style.display = grp.dataset.tts === target ? '' : 'none';
    });
    const visibleOpts = sel.querySelectorAll(`optgroup[data-tts="${target}"] option`);
    const validValues = Array.from(visibleOpts).map(o => o.value);
    if (validValues.length && !validValues.includes(sel.value)) {
      sel.value = validValues[0];
    }
  });
  updateApiStatusBadge();
}
function updateApiStatusBadge() {
  const badge = document.getElementById('apiStatusBadge');
  if (!badge) return;
  const hasText = getApiKeys().length > 0;
  const hasTts = getTtsProvider() === 'ai33pro' ? (document.getElementById('ai33ApiKey')?.value?.trim() || '').length > 0 : getTtsApiKeys().length > 0;
  badge.textContent = hasText ? (hasTts ? 'Configured' : 'TTS: configure in Settings') : 'Configure in Settings';
  badge.style.background = hasText ? 'rgba(16, 185, 129, 0.2)' : 'rgba(90, 97, 128, 0.2)';
}

// --- Get API keys (supports multiple: comma or newline separated) ---
function getApiKeys() {
  const el = document.getElementById('apiKey');
  const raw = (el && el.value) || '';
  const keys = raw.split(/[\s,;\n]+/).map(k => k.trim()).filter(Boolean);
  return [...new Set(keys)];
}

// --- Get single API key for non-parallel calls (uses first key) ---
function getApiKey() {
  const keys = getApiKeys();
  return keys[0] || '';
}

// --- Get selected AI provider ---
function getAIProvider() {
  return document.getElementById('aiProvider')?.value || 'gemini';
}

// --- Get TTS provider ---
function getTtsProvider() {
  return document.getElementById('ttsProvider')?.value || 'gemini';
}

// --- Get API key(s) for TTS ---
function getTtsApiKey() {
  if (getTtsProvider() === 'ai33pro') {
    return document.getElementById('ai33ApiKey')?.value?.trim() || '';
  }
  const ttsKey = document.getElementById('ttsApiKey')?.value?.trim();
  if (ttsKey) return ttsKey.split(/[\s,;\n]+/)[0]?.trim() || '';
  const geminiTts = document.getElementById('geminiTtsKey')?.value?.trim();
  if (geminiTts) return geminiTts;
  return getApiKey();
}

// --- Get Gemini API key for image generation (used when provider is Gemini or DeepSeek + Gemini key for TTS) ---
function getGeminiKeyForImages() {
  if (getAIProvider() === 'gemini') return getApiKey() || '';
  const geminiTts = document.getElementById('geminiTtsKey')?.value?.trim();
  return geminiTts || '';
}

function getTtsApiKeys() {
  if (getTtsProvider() === 'ai33pro') {
    const k = getTtsApiKey();
    return k ? [k] : [];
  }
  const raw = document.getElementById('ttsApiKey')?.value || document.getElementById('geminiTtsKey')?.value || '';
  if (!raw) return getAIProvider() === 'gemini' ? getApiKeys() : [];
  const keys = raw.split(/[\s,;\n]+/).map(k => k.trim()).filter(Boolean);
  return keys.length ? keys : (getAIProvider() === 'gemini' ? getApiKeys() : []);
}

function ttsKeyMissingMessage() {
  return getTtsProvider() === 'ai33pro'
    ? 'Add your AI33 Pro API key under Audio Generation → AI33 Pro API Key in Settings.'
    : 'Add a Gemini TTS API key (or your main Gemini key) under Audio Generation in Settings.';
}

/** Allowed voice IDs for AI casting (must match Settings optgroups per TTS provider). */
function getTtsVoiceOptionsForAi() {
  if (getTtsProvider() === 'ai33pro') {
    return {
      narrator: ['alloy', 'echo', 'fable', 'onyx', 'nova', 'shimmer'],
      female: ['nova', 'shimmer'],
      male: ['onyx', 'alloy'],
    };
  }
  return {
    narrator: ['Charon', 'Schedar', 'Sulafat', 'Kore', 'Puck', 'Orus'],
    female: ['Sulafat', 'Kore', 'Zephyr'],
    male: ['Charon', 'Puck', 'Orus'],
  };
}

function pickValidTtsVoice(value, allowed) {
  const v = String(value || '').trim().toLowerCase();
  const hit = allowed.find(a => a.toLowerCase() === v);
  return hit || allowed[0];
}

function applyTtsVoicesToSettingsUI(tv) {
  if (!tv || typeof tv !== 'object') return;
  const pairs = [
    ['narratorVoice', tv.narratorVoice],
    ['femaleVoice', tv.femaleVoice],
    ['maleVoice', tv.maleVoice],
  ];
  pairs.forEach(([id, val]) => {
    const el = document.getElementById(id);
    if (!el || !val) return;
    const ok = [...el.options].some(o => o.value === val);
    if (ok) {
      el.value = val;
      localStorage.setItem(id, val);
    }
  });
}

/** AI-picks narrator + female + male TTS voices from narrator tone, background, and character cast. */
async function suggestTtsVoicesFromTemplate(flowId, novelIndex) {
  const flow = getFlow(flowId);
  const novel = flow?.novels?.[novelIndex];
  if (!novel) {
    showToast('No novel template found.', 'error');
    return;
  }
  const fd = flowDomId(flowId);
  if (!getApiKey()) {
    showToast('Add your text API key in Settings (API Key field) so AI can suggest voices.', 'error');
    return;
  }
  const opts = getTtsVoiceOptionsForAi();
  const charSummary = (novel.characters || []).map(c =>
    `${safeStr(c.name)}${c.gender ? ` (${c.gender})` : ''}${c.role ? ` — ${c.role}` : ''}`
  ).join('; ') || 'Unknown cast — infer from tone only.';
  const prompt = `You are casting text-to-speech voices for a novel audio drama.

Novel template:
- Title: ${safeStr(novel.title)}
- Genre / category: ${safeStr(novel.genre || novel.category)}
- Narrator tone (primary cue for the narration voice): ${safeStr(novel.narratorTone)}
- Background: ${safeStr(novel.background)}
- Characters: ${charSummary}

TTS backend: ${getTtsProvider() === 'ai33pro' ? 'AI33 (OpenAI-compatible voices)' : 'Google Gemini preview TTS'}.

Pick exactly one voice ID from each list (copy spellings exactly):
- narratorVoice — NARRATOR lines and omniscient narration: ${opts.narrator.join(', ')}
- femaleVoice — female character dialogue: ${opts.female.join(', ')}
- maleVoice — male character dialogue: ${opts.male.join(', ')}

Return JSON only: narratorVoice, femaleVoice, maleVoice, rationale (one short sentence).`;

  const btn = document.getElementById(`suggestTtsVoicesBtn_${fd}_${novelIndex}`);
  const orig = btn ? btn.textContent : '';
  if (btn) {
    btn.disabled = true;
    btn.textContent = '…';
  }
  try {
    const raw = await callGeminiAPI(prompt);
    if (!raw || typeof raw !== 'object') throw new Error('Invalid AI response');
    const narratorVoice = pickValidTtsVoice(raw.narratorVoice, opts.narrator);
    const femaleVoice = pickValidTtsVoice(raw.femaleVoice, opts.female);
    const maleVoice = pickValidTtsVoice(raw.maleVoice, opts.male);
    const rationale = safeStr(raw.rationale || raw.note);
    novel.ttsVoices = { narratorVoice, femaleVoice, maleVoice, rationale };
    applyTtsVoicesToSettingsUI(novel.ttsVoices);
    const hint = document.getElementById(`ttsVoicesHint_${fd}_${novelIndex}`);
    if (hint) {
      hint.textContent = `Cast: ${narratorVoice} · ${femaleVoice} · ${maleVoice}${rationale ? ` — ${rationale}` : ''}`;
    }
    showToast('TTS voices set from this template (used for this novel’s audio). Settings updated.', 'success');
  } catch (e) {
    showToast(`Voice suggestion failed: ${e?.message || e}`, 'error');
  }
  if (btn) {
    btn.disabled = false;
    btn.textContent = orig || '🎚️ Suggest TTS voices from template';
  }
}

// --- Fetch with timeout ---
async function fetchWithTimeout(url, options, ms = 120000) {
  const ctrl = new AbortController();
  const id = setTimeout(() => ctrl.abort(), ms);
  try {
    const r = await fetch(url, { ...options, signal: ctrl.signal });
    return r;
  } catch (e) {
    if (e.name === 'AbortError') throw new Error('Request timed out after ' + (ms / 1000) + 's');
    throw e;
  } finally {
    clearTimeout(id);
  }
}

function extractFirstJsonValue(text) {
  const s = String(text || '');
  const fenced = s.match(/```(?:json)?\s*([\s\S]*?)```/i);
  const input = (fenced ? fenced[1] : s).trim();
  const o = input.indexOf('{');
  const a = input.indexOf('[');
  const startIdx = (o === -1) ? a : (a === -1 ? o : Math.min(o, a));
  if (startIdx < 0) return '';
  const openChar = input[startIdx];
  const closeChar = openChar === '{' ? '}' : ']';
  let depth = 0;
  let inStr = false;
  let esc = false;
  for (let i = startIdx; i < input.length; i++) {
    const ch = input[i];
    if (inStr) {
      if (esc) { esc = false; continue; }
      if (ch === '\\') { esc = true; continue; }
      if (ch === '"') inStr = false;
      continue;
    }
    if (ch === '"') { inStr = true; continue; }
    if (ch === openChar) depth++;
    if (ch === closeChar) {
      depth--;
      if (depth === 0) return input.slice(startIdx, i + 1).trim();
    }
  }
  return input.slice(startIdx).trim();
}

// --- Call DeepSeek API (OpenAI-compatible) ---
async function callDeepSeekAPI(prompt, expectJson = false) {
  const apiKey = getApiKey();
  if (!apiKey) throw new Error('No API key. Enter your DeepSeek key in the API Key(s) field.');
  const url = 'https://api.deepseek.com/v1/chat/completions';
  const callOnce = async (messages, bodyExtra = {}) => {
    const body = {
      model: 'deepseek-chat',
      messages,
      temperature: expectJson ? 0.2 : 0.85,
      max_tokens: 8192,
      ...bodyExtra,
    };
    const response = await fetchWithTimeout(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${apiKey}`,
      },
      body: JSON.stringify(body),
    }, 120000);
    if (!response.ok) {
      const err = await response.json().catch(() => ({}));
      throw new Error(err?.error?.message || err?.message || `DeepSeek API error: ${response.status}`);
    }
    const data = await response.json();
    return data.choices?.[0]?.message?.content?.trim() || '';
  };

  const baseSystem = expectJson
    ? 'Return ONLY valid JSON. No markdown, no code fences, no trailing commas, no comments, no extra text. Use double quotes for keys/strings.'
    : '';

  const messages1 = [];
  if (baseSystem) messages1.push({ role: 'system', content: baseSystem });
  messages1.push({ role: 'user', content: prompt });

  let text = await callOnce(messages1, expectJson ? { response_format: { type: 'json_object' } } : {});
  if (!text) throw new Error('No content returned from DeepSeek');
  if (!expectJson) return text;

  const candidate = extractFirstJsonValue(text);
  try {
    return JSON.parse(candidate || text);
  } catch (e) {
    const messages2 = [
      { role: 'system', content: baseSystem },
      { role: 'user', content: 'Fix the following into VALID JSON ONLY. Output ONLY the JSON object, nothing else.' },
      { role: 'user', content: text.slice(0, 60000) },
    ];
    const repaired = await callOnce(messages2, { response_format: { type: 'json_object' } });
    const cand2 = extractFirstJsonValue(repaired);
    try {
      return JSON.parse(cand2 || repaired);
    } catch (_) {
      throw new Error('Invalid JSON from DeepSeek. Please click Generate again.');
    }
  }
}

// --- Validation ---
function validateForm() {
  const keys = getApiKeys();
  const prompt = document.getElementById('masterPrompt').value.trim();

  if (!keys.length) {
    showToast('Please enter at least one API key (Gemini or DeepSeek)', 'error');
    document.getElementById('apiKey').focus();
    return false;
  }
  if (!prompt) {
    showToast('Please enter a master prompt / creative brief', 'error');
    document.getElementById('masterPrompt').focus();
    return false;
  }
  return true;
}

// --- Collect Form Data ---
function collectFormData() {
  return {
    numNovels: parseInt(document.getElementById('numNovels').value) || 3,
    masterPrompt: document.getElementById('masterPrompt').value.trim(),
    draftScript: document.getElementById('draftScript').value.trim(),
    characterSystem: document.getElementById('characterSystem').value.trim(),
    authorName: document.getElementById('authorName').value.trim(),
    collectionName: document.getElementById('collectionName')?.value || '',
    cateogoriesName: document.getElementById('categoryName')?.value || '',
    releaseDate: document.getElementById('releaseDate').value || '',
    narratorTone: document.getElementById('narratorTone').value.trim(),
    writingLanguage: document.getElementById('writingLanguage').value,
    referenceText: document.getElementById('referenceText').value.trim(),
  };
}

// --- Build Prompt ---
function buildGeminiPrompt(formData) {
  let prompt = `You are an expert novel architect and creative writing assistant. Based on the following creative brief, generate exactly ${formData.numNovels} unique and detailed novel templates.

## CREATIVE BRIEF
**Master Prompt / Idea:** ${formData.masterPrompt}
`;

  if (formData.draftScript) {
    prompt += `\n**Draft Script & Core Ideas:** ${formData.draftScript}`;
  }
  if (formData.characterSystem) {
    prompt += `\n**Character System Notes:** ${formData.characterSystem}`;
  }
  if (formData.authorName) {
    prompt += `\n**Author Name:** ${formData.authorName}`;
  }
  if (formData.releaseDate) {
    prompt += `\n**Target Release Date:** ${formData.releaseDate}`;
  }
  if (formData.narratorTone) {
    prompt += `\n**Narrator Tone & Background:** ${formData.narratorTone}`;
  }
  if (formData.writingLanguage) {
    prompt += `\n**Writing Language:** ${formData.writingLanguage}`;
  }
  if (formData.referenceText) {
    prompt += `\n**Reference Text (use as style/content inspiration):**\n${formData.referenceText.substring(0, 5000)}`;
  }
  if (formData.collectionName) {
    prompt += `\n**Collection:** ${formData.collectionName}`;
  }
  if (formData.cateogoriesName) {
    prompt += `\n**Categories:** ${formData.cateogoriesName}`;
  }

  const selectedCat = formData.cateogoriesName ? normalizeCategoryName(formData.cateogoriesName) : '';
  const categoryBinding = selectedCat
    ? `The brief suggests category "${selectedCat}". Use that exact string for "category" and "cateogories" on every novel unless it clearly conflicts with the story; if it conflicts, pick the closest allowed category.`
    : `For each novel, choose one "category" from the allowed list that best fits the story. Set "cateogories" to the same string as "category".`;

  const vietnameseNameRule = isVietnameseWritingLanguage(formData.writingLanguage)
    ? `\n- **Character naming rule (MANDATORY for Vietnamese output):** Use Sino-Vietnamese Chinese-style names for ALL characters (examples: "Lâm Hải", "Lâm Vi", "Triệu Mặc", "Tần Nguyệt"). Do NOT use Western names like Emma, James, Alex, etc.`
    : '';

  prompt += `

## AUTO-GENERATION (DEFAULT)
You invent the full content of every template field. The human will only supply a **Master Prompt / Idea** (and sometimes optional hints below). Do **not** leave placeholders, "TBD", or "N/A". Do **not** expect the user to control metadata: generate plausible **draftScript**, **characters** (character system), **authorName**, **releaseDate**, **category**, **narratorTone**, **background**, **collection**, and **cateogories** yourself. Optional lines in the brief are **hints only**—use them if helpful; if absent, derive everything from the Master Prompt alone.

## OPTIONAL HINTS (only if present above)
- **Draft Script & Core Ideas / Character System / Narrator Tone & Background / Author / Date / Category / Collection:** If provided, weave them in; otherwise ignore and create coherent values from the Master Prompt.
- **narratorTone** vs **background:** Split narrative voice/POV/mood into "narratorTone"; world/setting/atmosphere into "background". If you only have one combined idea, split it sensibly—both must be substantive.
- **category:** ${categoryBinding}
${vietnameseNameRule}

## OUTPUT REQUIREMENTS
Generate exactly ${formData.numNovels} novel templates. Each novel template MUST include ALL of the following fields:

1. **title** — A compelling, unique title for the novel
2. **synopsis** — A SINGLE short hook/tagline for listings, maximum 100 characters (including spaces). No newlines.
3. **overview** — A full-meaning overview of the novel: 2–5 sentences (about 300–900 characters). Plain prose only; no bullet lists. Explain premise, central conflict, emotional stakes, and what makes the story compelling. This is the reader-facing summary, not the draft script.
4. **draftScript** — Full draft scenario: core plot/scene outline and key ideas (you author this; optional brief hints may inspire it)
5. **characters** — An array of characters, each with: name, role (protagonist/antagonist/supporting), age, description, arc (character development summary), gender ("male" or "female" for voice casting)
6. **authorName** — ${formData.authorName ? `Prefer "${formData.authorName}" if it fits; otherwise a plausible pen name. Each novel should have a distinct author unless the brief implies a shared byline.` : 'A plausible author pen name per novel (e.g. Emma Vance, Marcus Webb). Each novel must have a different author name. Do NOT use "Unknown Author" or "Anonymous".'}
7. **releaseDate** — ${formData.releaseDate ? `Use "${formData.releaseDate}" unless a different date fits the story better.` : 'Invent a plausible YYYY-MM-DD (one shared date or varied per novel, e.g. within the next few years).'}
8. **narratorTone** — Narrative voice, POV, and tone (you invent; optional brief may hint)
9. **background** — World, setting, and story backdrop (distinct from narratorTone; you invent)
10. **writingLanguage** — "${formData.writingLanguage}"
11. **chapters** — An array of 5-10 chapter outlines, each with: chapterNumber, title, summary (SHORT). CRITICAL: For each chapter, the combined string "${'title'} — ${'summary'}" must be <= 100 characters.
12. **themes** — Array of 3–6 short, meaningful theme words or two-word phrases (e.g. redemption, first love, war trauma, family bonds, betrayal, survival). Use concrete terms only. Do NOT use vague phrases like "corruption of", "identity vs", "good vs evil", or "X of Y".
13. **genre** — Primary and secondary genres
14. **category** — MUST be exactly ONE of: ${ALLOWED_CATEGORIES.map(c => `"${c}"`).join(', ')} (no extra text)
15. **collection** — ${formData.collectionName ? `"${formData.collectionName}"` : 'Invent a short collection or series label that fits the novel (may repeat across the batch if they share a universe).'}
16. **cateogories** — ${formData.cateogoriesName ? `Same string as "category" (see optional category hint above).` : 'Same string as "category" for that novel.'}
17. **thumbnailPrompt** — One string for Gemini Image: book-cover art **inspired by draftScript, overview, themes, and background**; must describe **showing the novel title** as clear, readable typography (exact title text). Must be a single string.
18. **premium** — "yes" or "no" (whether the novel is premium content)
19. **show** — "yes" or "no" (whether to show the novel in listings)

Each novel should be DISTINCT — different plot, different character dynamics, different themes — while still being inspired by the creative brief.
CRITICAL: Every novel MUST have a non-empty synopsis (<= 100 characters) AND a non-empty overview (full meaning, roughly 300+ characters); non-empty draftScript, narratorTone, and background; at least two entries in characters; and MUST have 5-10 chapters with chapterNumber, title, and summary for each.

## OUTPUT FORMAT (CRITICAL)
You MUST return valid JSON only. No markdown code blocks, no backticks, no explanation—just the raw JSON object.
{
  "novels": [
    {
      "title": "...",
      "synopsis": "...",
      "overview": "...",
      "draftScript": "...",
      "characters": [
        { "name": "...", "role": "...", "age": "...", "description": "...", "arc": "...", "gender": "female" }
      ],
      "authorName": "...",
      "releaseDate": "...",
      "narratorTone": "...",
      "background": "...",
      "writingLanguage": "...",
      "chapters": [
        { "chapterNumber": 1, "title": "...", "summary": "..." }
      ],
      "themes": ["...", "..."],
      "genre": "...",
      "category": "...",
      "collection": "...",
      "cateogories": "...",
      "thumbnailPrompt": "...",
      "premium": "yes",
      "show": "yes"
    }
  ]
}`;

  return prompt;
}

/** Text bundle from the novel template to steer cover/thumbnail imagery (core ideas). */
function getNovelCoreIdeasForImage(novel, maxLen = 900) {
  const ov = safeStr(novel?.overview);
  const syn = safeStr(novel?.synopsis);
  const ds = safeStr(novel?.draftScript);
  const bg = safeStr(novel?.background);
  const themes = Array.isArray(novel?.themes) && novel.themes.length ? novel.themes.join(', ') : '';
  const chunks = [];
  if (ov) chunks.push(`Overview: ${ov}`);
  else if (syn) chunks.push(`Synopsis: ${syn}`);
  if (ds) chunks.push(`Core ideas / draft script: ${ds}`);
  if (bg) chunks.push(`Setting / backdrop: ${bg}`);
  if (themes) chunks.push(`Themes: ${themes}`);
  let s = chunks.join('\n');
  if (!s.trim()) s = 'Derive visuals from genre, narrator tone, and title.';
  if (s.length > maxLen) s = s.slice(0, Math.max(0, maxLen - 1)) + '…';
  return s;
}

function buildThumbnailPromptFromNovel(novel) {
  const title = safeStr(novel?.title) || 'Untitled novel';
  const genre = safeStr(novel?.genre) || safeStr(novel?.cateogories) || safeStr(novel?.category) || 'Fiction';
  const mood = safeStr(novel?.narratorTone) || 'Cinematic';
  const coreIdeas = getNovelCoreIdeasForImage(novel, 900);

  const lines = [
    'Create a SQUARE 1:1 book cover / thumbnail image.',
    `Typography (required): prominently display the novel title in clear, readable lettering. Exact title text to spell: "${title}". Use elegant, genre-appropriate fonts; strong contrast against the artwork.`,
    `Genre: ${genre}. Mood / narrative voice: ${mood}.`,
    'The illustration must be clearly inspired by the novel’s core ideas below (conflict, symbols, characters, world)—not generic unrelated stock imagery.',
    '---',
    coreIdeas,
    '---',
    'Style: cinematic or painterly book-cover quality, dramatic lighting, one strong focal subject, professional publishing look.',
    'Constraints: no watermarks, no logos. Title must be legible and spelled exactly as given. Avoid garbled text, extra limbs, or blurry faces.',
    'Output: 1 image only.',
  ];
  return lines.join('\n');
}

/** Build prompt for 3:4 portrait novel thumbnail (book-cover style) from template. */
function buildThumbnail34PromptFromNovel(novel) {
  const title = safeStr(novel?.title) || 'Untitled novel';
  const genre = safeStr(novel?.genre) || safeStr(novel?.cateogories) || safeStr(novel?.category) || 'Fiction';
  const mood = safeStr(novel?.narratorTone) || 'Cinematic';
  const coreIdeas = getNovelCoreIdeasForImage(novel, 900);
  const lines = [
    'Create a VERTICAL 3:4 portrait book cover / novel thumbnail (taller than wide).',
    `Typography (required): prominently display the novel title in clear, readable lettering. Exact title text: "${title}". Elegant, genre-appropriate fonts; high contrast.`,
    `Genre: ${genre}. Mood / narrative voice: ${mood}.`,
    'Artwork must be inspired by the novel’s core ideas below—scene, metaphor, or key visual from the story—not a generic unrelated image.',
    '---',
    coreIdeas,
    '---',
    'Style: cinematic, professional book-cover quality, high contrast, clear focal point.',
    'Composition: rule-of-thirds or centered hero subject; space for title; suited to 3:4 aspect ratio.',
    'Constraints: no watermarks, no logos. Title spelling must match exactly. Avoid garbled letters, extra limbs, blurry faces.',
    '',
    'Output: 1 image only.',
  ];
  return lines.join('\n');
}

// --- Test API (minimal call to verify key) ---
async function handleTestApi() {
  const keys = getApiKeys();
  if (!keys.length) {
    showToast('Enter an API key first', 'error');
    return;
  }
  const provider = getAIProvider();
  const btn = document.getElementById('testApiBtn');
  const origText = btn?.textContent;
  if (btn) { btn.disabled = true; btn.textContent = 'Testing...'; }
  try {
    if (provider === 'deepseek') {
      await callDeepSeekAPI('Reply with only: OK', false);
      showToast('DeepSeek API: Connection OK', 'success');
      updateApiStatusBadge();
    } else {
      const apiKey = getApiKey();
      const models = ['gemini-2.5-flash', 'gemini-2.0-flash', 'gemini-1.5-flash'];
      let lastErr = null;
      for (const model of models) {
        try {
          const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;
          const r = await fetchWithTimeout(url, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              contents: [{ parts: [{ text: 'Say OK' }] }],
              generationConfig: { maxOutputTokens: 10 },
            }),
          }, 15000);
          const data = await r.json();
          if (r.ok && data.candidates?.[0]?.content?.parts?.[0]?.text) {
            showToast(`Gemini API (${model}): OK`, 'success');
            updateApiStatusBadge();
            return;
          }
          if (!r.ok) {
            const msg = data?.error?.message || data?.message || `HTTP ${r.status}`;
            lastErr = msg;
            if (msg.includes('404') || msg.includes('not found') || msg.includes('Invalid model')) continue;
            throw new Error(msg);
          }
        } catch (e) {
          lastErr = e.message;
          if (e.message.includes('404') || e.message.includes('not found') || e.message.includes('Invalid model')) continue;
          throw e;
        }
      }
      throw new Error(lastErr || 'All models failed');
    }
  } catch (e) {
    const msg = e?.message || String(e);
    showToast('API test failed: ' + msg, 'error');
    console.error('API test error:', e);
  } finally {
    if (btn) { btn.disabled = false; btn.textContent = origText || 'Test'; }
  }
}

// --- API Call (routes to Gemini or DeepSeek) ---
async function callGeminiAPI(prompt) {
  if (getAIProvider() === 'deepseek') {
    return callDeepSeekAPI(prompt, true);
  }
  const apiKey = getApiKey();
  if (!apiKey) throw new Error('No API key. Enter your Gemini key in the API Key(s) field.');
  const models = ['gemini-2.5-flash', 'gemini-2.0-flash', 'gemini-1.5-flash'];
  let lastErr = null;
  for (const model of models) {
    try {
      const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;
      const response = await fetchWithTimeout(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          contents: [{ parts: [{ text: prompt }] }],
          generationConfig: {
            temperature: 0.9,
            topP: 0.95,
            topK: 40,
            maxOutputTokens: 16384,
            responseMimeType: 'application/json',
          },
        }),
      }, 120000);
      if (!response.ok) {
        const errData = await response.json().catch(() => ({}));
        lastErr = errData?.error?.message || errData?.message || `API error: ${response.status}`;
        if (lastErr.includes('404') || lastErr.includes('not found') || lastErr.includes('Invalid model')) continue;
        throw new Error(lastErr);
      }
      const data = await response.json();
      const text = data.candidates?.[0]?.content?.parts?.[0]?.text;
      if (!text) throw new Error('No content returned from Gemini API');
      return JSON.parse(text);
    } catch (e) {
      lastErr = e?.message || lastErr;
      if (e?.message?.includes('404') || e?.message?.includes('not found') || e?.message?.includes('Invalid model')) continue;
      throw e;
    }
  }
  throw new Error(lastErr || 'All Gemini models failed');
}

async function callGeminiAPIWithKey(prompt, apiKeyOverride = null, cfgOverride = null) {
  const apiKey = apiKeyOverride || getApiKey();
  if (!apiKey) throw new Error('No API key. Enter your Gemini key in the API Key(s) field.');
  const models = ['gemini-2.5-flash', 'gemini-2.0-flash', 'gemini-1.5-flash'];
  let lastErr = null;
  for (const model of models) {
    try {
      const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;
      const response = await fetchWithTimeout(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          contents: [{ parts: [{ text: prompt }] }],
          generationConfig: {
            temperature: 0.9,
            topP: 0.95,
            topK: 40,
            maxOutputTokens: 8192,
            responseMimeType: 'application/json',
            ...(cfgOverride || {}),
          },
        }),
      }, 120000);
      if (!response.ok) {
        const errData = await response.json().catch(() => ({}));
        lastErr = errData?.error?.message || errData?.message || `API error: ${response.status}`;
        if (lastErr.includes('404') || lastErr.includes('not found') || lastErr.includes('Invalid model')) continue;
        throw new Error(lastErr);
      }
      const data = await response.json();
      const text = data.candidates?.[0]?.content?.parts?.[0]?.text;
      if (!text) throw new Error('No content returned from Gemini API');
      return JSON.parse(text);
    } catch (e) {
      lastErr = e?.message || lastErr;
      if (e?.message?.includes('404') || e?.message?.includes('not found') || e?.message?.includes('Invalid model')) continue;
      throw e;
    }
  }
  throw new Error(lastErr || 'All Gemini models failed');
}

function buildSingleNovelPrompt(formData, novelIndex, total, avoidTitles = []) {
  const uniqSalt = `${Date.now()}_${Math.random().toString(16).slice(2)}_${novelIndex + 1}`;
  const asOne = { ...formData, numNovels: 1 };
  let p = buildGeminiPrompt(asOne);

  // Hard "lanes" to prevent parallel calls converging on the same idea.
  const lanes = [
    { genre: 'Urban Fantasy + Mystery', setting: 'modern city with hidden magic', twist: 'the "curse" is a bureaucratic spell system', vibe: 'noir, witty, fast' },
    { genre: 'Sci‑Fi Thriller', setting: 'near-future space habitat', twist: 'the antagonist is a safety protocol', vibe: 'tense, cinematic' },
    { genre: 'Romantic Drama', setting: 'small coastal town', twist: 'the love story is anchored by a legal secret', vibe: 'warm, bittersweet' },
    { genre: 'Historical Adventure', setting: '18th-century trade route', twist: 'a map is a coded confession', vibe: 'swashbuckling, vivid' },
    { genre: 'Horror', setting: 'isolated research facility', twist: 'the monster is an inherited memory', vibe: 'claustrophobic, eerie' },
    { genre: 'Cozy Mystery', setting: 'bookshop / café community', twist: 'the clues are hidden in margins', vibe: 'charming, clever' },
    { genre: 'Epic Fantasy', setting: 'frontier kingdom on the brink', twist: 'prophecy is a manufactured narrative', vibe: 'mythic, grand' },
    { genre: 'Crime / Heist', setting: 'glamorous metropolis', twist: 'the heist is to steal a person’s identity record', vibe: 'slick, witty' },
    { genre: 'YA Coming-of-Age', setting: 'boarding school / academy', twist: 'the “tests” are moral experiments', vibe: 'bright, heartfelt' },
    { genre: 'Speculative Literary', setting: 'a town with one impossible rule', twist: 'breaking the rule rewrites relationships', vibe: 'poetic, thoughtful' },
  ];
  const lane = lanes[novelIndex % lanes.length];
  const avoid = Array.isArray(avoidTitles) && avoidTitles.length
    ? avoidTitles.map(t => `"${safeStr(t).slice(0, 80)}"`).filter(Boolean).slice(0, 25).join(', ')
    : '';

  p += `\n\n## UNIQUENESS & LANE (CRITICAL)\nThis request is for Novel ${novelIndex + 1} of ${total}.\n` +
    `You MUST follow this lane:\n` +
    `- Required genre lane: ${lane.genre}\n` +
    `- Required setting: ${lane.setting}\n` +
    `- Required core twist: ${lane.twist}\n` +
    `- Required vibe: ${lane.vibe}\n` +
    `Title MUST be unique and MUST clearly reflect this lane.\n` +
    (avoid ? `Do NOT use any of these titles (duplicates): ${avoid}\n` : '') +
    `Use this uniqueness seed (do not repeat it in output): ${uniqSalt}\n`;
  return p;
}

async function generateNovelsParallel(formData) {
  const total = Math.max(1, parseInt(formData?.numNovels) || 1);
  const keys = getApiKeys();
  if (!keys.length && getAIProvider() !== 'deepseek') {
    throw new Error('No API key. Enter your Gemini key in the API Key(s) field.');
  }
  // Parallel novels: multiple templates in flight at once. Scale with key count; still allow several
  // concurrent requests on a single key (typical Gemini quota allows parallel calls).
  const keyCount = Math.max(1, keys.length || 1);
  const concurrency = Math.min(8, total, Math.max(4, keyCount * 4));
  const queue = Array.from({ length: total }, (_, i) => i);
  const out = new Array(total);
  let done = 0;

  const worker = async (workerId) => {
    while (queue.length) {
      const i = queue.shift();
      if (i == null) break;
      const prompt = buildSingleNovelPrompt(formData, i, total);
      const key = keys.length ? keys[workerId % keys.length] : null;
      const cfg = { maxOutputTokens: 6000 };
      const r = (getAIProvider() === 'deepseek')
        ? await callDeepSeekAPI(prompt, true)
        : await callGeminiAPIWithKey(prompt, key, cfg);

      const novel = Array.isArray(r?.novels) ? r.novels[0] : (r?.novel || r?.data?.novel);
      if (!novel || typeof novel !== 'object') throw new Error('Invalid novel JSON returned');
      out[i] = novel;
      done++;
      try {
        updateProgress(40 + Math.floor((done / total) * 45), `Generated ${done}/${total} novels...`);
      } catch (_) {}
      // Small jitter to reduce burst/rate-limit collisions.
      await new Promise(res => setTimeout(res, 120 + Math.floor(Math.random() * 160)));
    }
  };

  await Promise.all(Array.from({ length: concurrency }, (_, wid) => worker(wid)));

  // Retry duplicate titles once with explicit "avoidTitles".
  const normTitle = (t) => safeStr(t).trim().toLowerCase().replace(/\s+/g, ' ');
  const seen = new Map();
  const dupIndices = [];
  out.forEach((n, idx) => {
    const t = normTitle(n?.title);
    if (!t) return;
    if (!seen.has(t)) seen.set(t, idx);
    else dupIndices.push(idx);
  });
  if (dupIndices.length) {
    const avoidTitles = out.map(n => safeStr(n?.title)).filter(Boolean);
    const retryQueue = dupIndices.slice();
    const retryConcurrency = Math.min(6, retryQueue.length, Math.max(1, keys.length || 1));
    const retryWorker = async (workerId) => {
      while (retryQueue.length) {
        const i = retryQueue.shift();
        if (i == null) break;
        const prompt = buildSingleNovelPrompt(formData, i, total, avoidTitles);
        const key = keys.length ? keys[workerId % keys.length] : null;
        const cfg = { maxOutputTokens: 6500, temperature: 1.0 };
        const r = (getAIProvider() === 'deepseek')
          ? await callDeepSeekAPI(prompt, true)
          : await callGeminiAPIWithKey(prompt, key, cfg);
        const novel = Array.isArray(r?.novels) ? r.novels[0] : (r?.novel || r?.data?.novel);
        if (novel && typeof novel === 'object') out[i] = novel;
        await new Promise(res => setTimeout(res, 180 + Math.floor(Math.random() * 220)));
      }
    };
    await Promise.all(Array.from({ length: retryConcurrency }, (_, wid) => retryWorker(wid)));
  }

  return out.filter(Boolean);
}

// --- Generate Handler (per-flow; parallel runs do not block each other) ---
async function handleGenerate() {
  if (!validateForm()) return;
  ensureActiveFlow();
  const flow = getFlow(state.activeFlowId);
  if (flow?.generating) return;
  await runFlowGenerate(state.activeFlowId);
}

async function handleNewParallelRun() {
  // Parallel run flow was removed; keep compatibility by triggering normal generation.
  await handleGenerate();
}

async function runFlowGenerate(flowId) {
  const flow = getFlow(flowId);
  if (!flow || flow.generating) return;
  if (!validateForm()) return;

  flow.generating = true;
  flow.status = 'generating';
  flow.errorMessage = '';
  renderFlowTabs();
  refreshGenerateButtonState();
  setFlowStatusLine(`${flow.label}: generating templates…`);

  showProgress(true);
  updateProgress(10, 'Preparing creative brief…');

  try {
    const formData = collectFormData();
    flow.label = flowLabelFromPrompt(formData.masterPrompt);
    renderFlowTabs();
    updateProgress(25, 'Building AI prompt…');

    const provider = getAIProvider() === 'deepseek' ? 'DeepSeek' : 'Gemini';
    updateProgress(40, `Generating ${formData.numNovels} novel templates with ${provider} (parallel mode)…`);

    const novels = await generateNovelsParallel(formData);
    updateProgress(85, 'Processing results…');

    flow.novels = novels;
    normalizeNovelsForExport(flow.novels);
    stampCollectionAndCategoriesFromForm(flow.novels);
    applyAutoReviewTemplateToNovels(flow.novels);
    addHistoryRun(formData, flow.novels, flowId);
    flow.status = 'ready';
    updateProgress(95, 'Rendering results…');
    setTimeout(() => {
      showProgress(false);
      setFlowStatusLine('');
      renderFlowResults(flowId);
      renderFlowTabs();
      showToast(`Generated ${flow.novels.length} templates. Generating thumbnails…`, 'success');
      generateCoversForAllTemplates(flowId);

      const runStories = document.getElementById('autoGenerateStoriesAfterTemplates')?.checked;
      const keys = getApiKeys();
      if (runStories && keys.length && flow.novels?.length) {
        showToast('Templates ready. Generating full stories from each template…', 'info');
        void handleGenerateAllStories(flowId).catch((err) => {
          console.error('Auto full stories failed:', err);
          showToast('Full-story pass failed: ' + (err?.message || String(err)), 'error');
        });
      } else if (runStories && !keys.length) {
        showToast('Enable “after templates” full stories only works with an API key in Settings.', 'error');
      }
    }, 300);
  } catch (error) {
    console.error('Generation error:', error);
    showProgress(false);
    setFlowStatusLine('');
    flow.status = 'error';
    flow.errorMessage = error?.message || String(error);
    let msg = flow.errorMessage;
    if (msg === 'Failed to fetch' || msg.includes('NetworkError')) {
      msg = 'Network error. If using file://, run: npx serve -l 3000 and open http://localhost:3000';
    }
    showToast('Generation failed: ' + msg, 'error');
  } finally {
    flow.generating = false;
    renderFlowTabs();
    refreshGenerateButtonState();
    setFlowStatusLine(getActiveFlow()?.generating ? `${getActiveFlow().label}: generating…` : '');
  }
}

// --- Auto-generate cover + thumbnail for templates (for review UI + export) ---
function ensureCoverThumbInCard(flowId, index) {
  const flow = getFlow(flowId);
  if (!flow) return;
  const fd = flowDomId(flowId);
  const novel = flow.novels?.[index];
  if (!novel) return;
  const dataUrl =
    (typeof novel.thumbnail === 'string' && novel.thumbnail.startsWith('data:image/')) ? novel.thumbnail
      : (typeof novel.cover === 'string' && novel.cover.startsWith('data:image/')) ? novel.cover : '';
  if (!dataUrl) return;
  const panel = document.getElementById(`flowPanel_${fd}`);
  const card = panel?.querySelector(`.novel-card[data-index="${index}"]`);
  const info = card?.querySelector('.novel-card-header .novel-info');
  if (!info) return;
  if (info.querySelector(`img.novel-cover-thumb[data-index="${index}"]`)) return;
  const img = document.createElement('img');
  img.className = 'novel-cover-thumb';
  img.dataset.index = String(index);
  img.alt = `Cover ${index + 1}`;
  img.src = dataUrl;
  info.appendChild(img);
}

async function generateCoverForNovel(flowId, index) {
  const flow = getFlow(flowId);
  if (!flow) return;
  const novel = flow.novels?.[index];
  if (!novel) return;
  if (typeof novel.cover === 'string' && novel.cover.startsWith('data:image/')) {
    ensureCoverThumbInCard(flowId, index);
    return;
  }
  if (typeof callImageGenerationAPI !== 'function') return;

  const prompt = buildThumbnailPromptFromNovel(novel);

  const cover = await callImageGenerationAPI(prompt, novel, { aspectRatio: '1:1' });
  novel.cover = cover;
  novel.thumbnail = await resizeDataUrl(cover, 128, 128, 'image/png');
  ensureCoverThumbInCard(flowId, index);
}

async function generateCoversForAllTemplates(flowId) {
  const flow = getFlow(flowId);
  if (!flow) return;
  const novels = flow.novels || [];
  if (!Array.isArray(novels) || !novels.length) return;

  let indices = novels.map((_, i) => i).filter(i => !(typeof novels[i]?.cover === 'string' && novels[i].cover.startsWith('data:image/')));
  if (!indices.length) {
    novels.forEach((_, i) => ensureCoverThumbInCard(flowId, i));
    return;
  }

  const concurrency = Math.min(5, indices.length);
  const queue = indices.slice();
  let done = 0;
  showToast(`Generating ${indices.length} thumbnails for templates...`, 'info');

  const worker = async () => {
    while (queue.length) {
      const i = queue.shift();
      if (i == null) break;
      try {
        await generateCoverForNovel(flowId, i);
        done++;
        showToast(`Cover images: ${done}/${indices.length}`, 'info');
      } catch (e) {
        console.warn('Cover generation failed', i, e);
        showToast(`Cover generation failed for novel ${i + 1}: ${e?.message || 'Unknown error'}`, 'error');
      }
      await new Promise(r => setTimeout(r, 250));
    }
  };

  await Promise.all(Array.from({ length: concurrency }, () => worker()));
  showToast('Cover thumbnails ready', 'success');
}

// --- Thumbnail validation helpers (for 3:4 generation) ---
function approxEqual(a, b, tolerance = 0.03) {
  const x = Number(a);
  const y = Number(b);
  if (!Number.isFinite(x) || !Number.isFinite(y)) return false;
  return Math.abs(x - y) <= tolerance;
}

async function readImageSize(dataUrl) {
  if (typeof dataUrl !== 'string' || !dataUrl.startsWith('data:image/')) {
    throw new Error('Not an image data URL');
  }
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve({ width: img.naturalWidth || img.width, height: img.naturalHeight || img.height });
    img.onerror = () => reject(new Error('Image failed to load'));
    img.src = dataUrl;
  });
}

async function validateThumbnail34DataUrl(dataUrl) {
  const { width, height } = await readImageSize(dataUrl);
  const ratio = width && height ? (width / height) : 0;
  const expected = 3 / 4;
  const ratioOk = approxEqual(ratio, expected, 0.04);
  const minSizeOk = width >= 300 && height >= 400;
  return {
    ok: Boolean(ratioOk && minSizeOk),
    width,
    height,
    ratio,
    ratioOk,
    minSizeOk,
    expectedRatio: expected,
  };
}

function setThumbnail34StatusInCard(flowId, index, meta) {
  const fd = flowDomId(flowId);
  const panel = document.getElementById(`flowPanel_${fd}`);
  const card = panel?.querySelector(`.novel-card[data-index="${index}"]`);
  const el = card?.querySelector(`#thumbnail34Status_${fd}_${index}`);
  if (!el) return;
  if (!meta) {
    el.textContent = 'Not generated yet.';
    el.dataset.state = 'neutral';
    return;
  }
  if (meta.error) {
    el.textContent = `Check failed: ${meta.error}`;
    el.dataset.state = 'bad';
    return;
  }
  const dims = `${meta.width}×${meta.height}`;
  const ratioTxt = meta.ratio ? meta.ratio.toFixed(3) : 'N/A';
  if (meta.ok) {
    el.textContent = `✅ OK — ${dims} (ratio ${ratioTxt} ≈ 3:4)`;
    el.dataset.state = 'ok';
  } else {
    const problems = [
      meta.ratioOk ? null : `ratio ${ratioTxt} ≠ 3:4`,
      meta.minSizeOk ? null : 'too small',
    ].filter(Boolean).join(', ');
    el.textContent = `⚠️ Warning — ${dims} (${problems})`;
    el.dataset.state = 'warn';
  }
}

// --- Generate novel thumbnail (3:4 ratio) from template ---
async function generateThumbnail34ForNovel(flowId, index, options = {}) {
  const silent = Boolean(options.silent);
  const flow = getFlow(flowId);
  if (!flow) return;
  const fd = flowDomId(flowId);
  const novel = flow.novels?.[index];
  if (!novel) return;
  const btn = document.getElementById(`generateThumbnail34Btn_${fd}_${index}`);
  if (btn) { btn.disabled = true; btn.classList.add('loading'); }
  try {
    novel.thumbnail34Check = null;
    setThumbnail34StatusInCard(flowId, index, null);
    const prompt = buildThumbnail34PromptFromNovel(novel);
    const dataUrl = await callImageGenerationAPI(prompt, novel, { aspectRatio: '3:4' });
    if (dataUrl) {
      novel.thumbnail = dataUrl;
      novel.cover = dataUrl;
      const panel = document.getElementById(`flowPanel_${fd}`);
      const card = panel?.querySelector(`.novel-card[data-index="${index}"]`);
      const img = card?.querySelector(`.novel-cover-thumb[data-index="${index}"]`);
      if (img) img.src = dataUrl; else ensureCoverThumbInCard(flowId, index);
      try {
        novel.thumbnail34Check = await validateThumbnail34DataUrl(dataUrl);
      } catch (e) {
        novel.thumbnail34Check = { error: e?.message || String(e) };
      }
      setThumbnail34StatusInCard(flowId, index, novel.thumbnail34Check);
      const usedGemini = !!getGeminiKeyForImages();
      if (!silent) {
        showToast(`Thumbnail generated. Check: ${novel.thumbnail34Check?.ok ? 'OK' : 'Review'}${!usedGemini ? ' (free API may not be 3:4)' : ''}`, novel.thumbnail34Check?.ok ? 'success' : 'info');
      }
    }
  } catch (e) {
    if (!silent) showToast(`Thumbnail failed: ${e?.message || 'Unknown error'}`, 'error');
    novel.thumbnail34Check = { error: e?.message || String(e) };
    setThumbnail34StatusInCard(flowId, index, novel.thumbnail34Check);
  }
  if (btn) { btn.disabled = false; btn.classList.remove('loading'); }
}

async function generateThumbnails34ForAll() {
  const flow = ensureActiveFlow();
  const novels = flow.novels || [];
  if (!novels.length) {
    showToast('Generate novel templates first.', 'error');
    return;
  }
  const flowId = state.activeFlowId;
  const btn = document.getElementById('generateThumbnails34AllBtn');
  if (btn) { btn.disabled = true; btn.classList.add('loading'); }
  const keys = getApiKeys();
  const keyCount = Math.max(1, keys.length || 1);
  const indices = novels.map((_, i) => i);
  const queue = indices.slice();
  const concurrency = Math.min(6, queue.length, Math.max(3, keyCount * 2));
  let done = 0;
  const total = indices.length;
  showToast(`Generating ${total} novel thumbnails (3:4), ${concurrency} at a time…`, 'info');

  const worker = async () => {
    while (queue.length) {
      const i = queue.shift();
      if (i == null) break;
      try {
        await generateThumbnail34ForNovel(flowId, i, { silent: true });
      } catch (_) {}
      done++;
      showToast(`Thumbnails: ${done}/${total}`, 'info');
      await new Promise(r => setTimeout(r, 40));
    }
  };

  await Promise.all(Array.from({ length: concurrency }, () => worker()));
  if (btn) { btn.disabled = false; btn.classList.remove('loading'); }
  showToast(`Novel thumbnails (3:4) finished (${total}).`, 'success');
}

// --- Render Results (per flow / tab) ---
function renderFlowResults(flowId, novelsOpt) {
  const flow = getFlow(flowId);
  if (!flow) return;
  const novels = novelsOpt != null ? novelsOpt : flow.novels;
  ensureFlowPanelDom(flowId);
  const fd = flowDomId(flowId);
  const section = document.getElementById('resultsSection');
  const container = document.getElementById(`novelsContainer_${fd}`);
  const countEl = document.getElementById('resultsCount');

  if (!container || !section) return;
  section.classList.add('active');
  if (state.activeFlowId === flowId && countEl) {
    countEl.textContent = `${novels.length} novels generated`;
  }
  container.innerHTML = '';

  novels.forEach((novel, index) => {
    const card = createNovelCard(novel, index, flowId);
    card.style.animationDelay = `${index * 0.1}s`;
    container.appendChild(card);
  });

  attachEditSyncListeners(container, flowId);

  try {
    novels.forEach((_, i) => ensureCoverThumbInCard(flowId, i));
  } catch (_) {}

  try {
    novels.forEach((n, i) => setThumbnail34StatusInCard(flowId, i, n?.thumbnail34Check || null));
  } catch (_) {}

  const firstCard = container?.querySelector('.novel-card');
  if (firstCard) firstCard.classList.add('expanded');

  if (state.activeFlowId === flowId) {
    section.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }
}

function createNovelCard(novel, index, flowId) {
  const flow = getFlow(flowId);
  const fd = flowDomId(flowId);
  const card = document.createElement('div');
  card.className = 'novel-card';
  card.dataset.index = index;
  card.dataset.flowId = flowId;

  let charactersHtml = '';
  if (novel.characters && Array.isArray(novel.characters)) {
    charactersHtml = novel.characters.map((c, ci) => `
      <li contenteditable="true" data-novel="${index}" data-charindex="${ci}"><strong>${c.name}</strong> (${c.role}${c.age ? ', ' + c.age : ''}) — ${c.description}${c.arc ? '<br><em>Arc: ' + c.arc + '</em>' : ''}</li>
    `).join('');
    charactersHtml = `<ul>${charactersHtml}</ul>`;
  }

  let chaptersHtml = '';
  if (novel.chapters && Array.isArray(novel.chapters)) {
    chaptersHtml = novel.chapters.map((ch, ci) => `
      <div class="chapter-item" data-novel="${index}" data-chapterindex="${ci}">
        <span class="chapter-num">Ch.${ch.chapterNumber}</span>
        <span class="chapter-title editable" contenteditable="true" data-field="chapterTitle">${escapeHtml(ch.title)}</span>
      </div>
    `).join('');
  }

  const tagsDisplay = getTagsDisplayText(novel);

  const isReviewed = flow && flow.reviewedNovels.has(index);
  const coverThumb = (typeof novel?.thumbnail === 'string' && novel.thumbnail.startsWith('data:image/'))
    ? novel.thumbnail
    : (typeof novel?.cover === 'string' && novel.cover.startsWith('data:image/')) ? novel.cover : '';
  const coverHref = (typeof novel?.cover === 'string' && novel.cover.startsWith('data:image/')) ? novel.cover : coverThumb;
  card.innerHTML = `
    <div class="novel-card-header" onclick="toggleNovelCard('${flowId}',${index})">
      <div class="novel-info">
        <div class="novel-number">${index + 1}</div>
        <div class="novel-title editable" contenteditable="true" data-novel="${index}" data-field="title">${escapeHtml(novel.title || 'Untitled Novel')}</div>
        ${coverThumb ? `<a href="${coverHref}" target="_blank" rel="noopener" title="Open cover image"><img class="novel-cover-thumb" data-index="${index}" src="${coverThumb}" alt="Cover ${index + 1}"/></a>` : ''}
      </div>
      <div class="actions">
        <button class="btn btn-secondary btn-sm" onclick="event.stopPropagation(); generateThumbnail34ForNovel('${flowId}',${index})" id="generateThumbnail34Btn_${fd}_${index}" title="Generate 3:4 thumbnail from this novel template">
          <span class="spinner"></span><span class="btn-text">🖼️ Thumbnail (3:4)</span>
        </button>
        <button class="btn btn-story btn-sm" onclick="event.stopPropagation(); generateFullStory('${flowId}',${index})" id="storyBtn_${fd}_${index}">
          <span class="spinner"></span><span class="btn-text">📖 Generate Full Story</span>
        </button>
        <button class="btn btn-secondary btn-sm review-toggle ${isReviewed ? 'reviewed' : ''}" onclick="event.stopPropagation(); toggleManualReview('${flowId}',${index})" title="${isReviewed ? 'Revoke manual review' : 'Mark as passed manual review'}">
          ${isReviewed ? '✅ Passed' : '⬜ Review'}
        </button>
        <button class="btn btn-secondary btn-sm" onclick="event.stopPropagation(); downloadNovel('${flowId}',${index})" title="Download template as .txt">
          📥 Template
        </button>
        <span class="expand-icon">▼</span>
      </div>
    </div>
    <div class="novel-card-body">
      ${novel._autoReview?.template ? `
      <div class="auto-review-banner ${novel._autoReview.template.ok ? 'ok' : 'warn'}" id="autoReviewTemplate_${fd}_${index}">
        <div class="auto-review-title">🔍 Auto-review (template)</div>
        <p class="auto-review-sub">Character system synchronized (names deduped, roles/genders normalized).</p>
        ${formatAutoReviewIssuesHtml(novel._autoReview.template.issues)}
      </div>` : ''}
      <div class="edit-hint">💡 Click any text field below to edit it</div>

      <div class="novel-field">
        <div class="novel-field-label">📖 Genre</div>
        <div class="novel-field-content editable" contenteditable="true" data-novel="${index}" data-field="genre">${escapeHtml(novel.genre || 'N/A')}</div>
      </div>

      <div class="novel-field">
        <div class="novel-field-label">📂 Category</div>
        <div class="novel-field-content editable" contenteditable="true" data-novel="${index}" data-field="category">${escapeHtml(novel.category || 'N/A')}</div>
      </div>

      <div class="novel-field">
        <div class="novel-field-label">🏷️ Tags</div>
        <div class="novel-field-content editable" contenteditable="true" data-novel="${index}" data-field="tags" data-placeholder="Comma-separated tags">${escapeHtml(tagsDisplay)}</div>
      </div>

      <div class="divider"></div>

      <div class="novel-field">
        <div class="novel-field-label">📖 Overview <span class="optional-badge">full meaning</span></div>
        <div class="novel-field-content editable tall" contenteditable="true" data-novel="${index}" data-field="overview">${escapeHtml(novel.overview || novel.synopsis || 'N/A')}</div>
      </div>

      <div class="novel-field">
        <div class="novel-field-label">📝 Short tagline <span class="optional-badge">≤100 chars</span></div>
        <div class="novel-field-content editable" contenteditable="true" data-novel="${index}" data-field="synopsis">${escapeHtml(novel.synopsis || 'N/A')}</div>
      </div>

      <div class="divider"></div>

      <div class="novel-field">
        <div class="novel-field-label">🎬 Draft Script & Core Ideas</div>
        <div class="novel-field-content editable" contenteditable="true" data-novel="${index}" data-field="draftScript">${escapeHtml(novel.draftScript || 'N/A')}</div>
      </div>

      <div class="divider"></div>

      <div class="novel-field">
        <div class="novel-field-label">👥 Character System</div>
        <div class="novel-field-content">${charactersHtml || '<div contenteditable="true" class="editable">N/A</div>'}</div>
      </div>

      <div class="divider"></div>

      <div class="novel-field">
        <div class="novel-field-label">🎭 Narrator Tone</div>
        <div class="novel-field-content editable" contenteditable="true" data-novel="${index}" data-field="narratorTone">${escapeHtml(novel.narratorTone || 'N/A')}</div>
        <div class="tts-voice-suggest-row">
          <button type="button" class="btn btn-secondary btn-sm" onclick="event.stopPropagation(); suggestTtsVoicesFromTemplate('${flowId}',${index})" id="suggestTtsVoicesBtn_${fd}_${index}" title="Use AI (narrator tone + cast) to pick TTS voices; applies to this novel and updates Settings">
            🎚️ Suggest TTS voices from template
          </button>
          <span class="form-hint tts-voices-hint" id="ttsVoicesHint_${fd}_${index}">${novel.ttsVoices?.narratorVoice ? `Cast: ${escapeHtml(novel.ttsVoices.narratorVoice)} · ${escapeHtml(novel.ttsVoices.femaleVoice)} · ${escapeHtml(novel.ttsVoices.maleVoice)}` : ''}</span>
        </div>
      </div>

      <div class="novel-field">
        <div class="novel-field-label">🌍 Background / Setting</div>
        <div class="novel-field-content editable" contenteditable="true" data-novel="${index}" data-field="background">${escapeHtml(novel.background || 'N/A')}</div>
      </div>

      <div class="novel-field">
        <div class="novel-field-label">🖼️ Thumbnail / cover prompt (title on image + art from core ideas)</div>
        <div class="novel-field-content editable" contenteditable="true" data-novel="${index}" data-field="thumbnailPrompt">${escapeHtml(novel.thumbnailPrompt || 'N/A')}</div>
      </div>

      <div class="novel-field">
        <div class="novel-field-label">✅ Thumbnail (3:4) Check</div>
        <div class="novel-field-content" id="thumbnail34Status_${fd}_${index}" data-state="neutral">Not generated yet.</div>
      </div>

      <div class="divider"></div>

      <div class="novel-field">
        <div class="novel-field-label">📚 Chapter Outline</div>
        <div class="novel-field-content">
          <div class="chapter-list">${chaptersHtml || 'N/A'}</div>
        </div>
      </div>

      <div class="divider"></div>

      <div class="novel-field" style="display:flex; gap: 40px; flex-wrap: wrap;">
        <div>
          <div class="novel-field-label">✍️ Author</div>
          <div class="novel-field-content editable" contenteditable="true" data-novel="${index}" data-field="authorName">${escapeHtml(novel.authorName || 'N/A')}</div>
        </div>
        <div>
          <div class="novel-field-label">📅 Release Date</div>
          <div class="novel-field-content editable" contenteditable="true" data-novel="${index}" data-field="releaseDate">${escapeHtml(novel.releaseDate || 'N/A')}</div>
        </div>
        <div>
          <div class="novel-field-label">🌐 Language</div>
          <div class="novel-field-content editable" contenteditable="true" data-novel="${index}" data-field="writingLanguage">${escapeHtml(novel.writingLanguage || 'N/A')}</div>
        </div>
        <div>
          <div class="novel-field-label">⭐ Premium</div>
          <div class="novel-field-content editable" contenteditable="true" data-novel="${index}" data-field="premium">${escapeHtml(novel.premium || 'yes')}</div>
        </div>
        <div>
          <div class="novel-field-label">👁️ Show</div>
          <div class="novel-field-content editable" contenteditable="true" data-novel="${index}" data-field="show">${escapeHtml(novel.show || 'yes')}</div>
        </div>
      </div>

      <div class="divider"></div>

      ${!isReviewed ? '<div class="review-required-inline" id="reviewRequired_' + fd + '_' + index + '"><span class="review-required-icon">📋</span> (Optional) Mark as <strong>Passed Manual Review</strong> to track templates you approve.</div>' : ''}

      <div class="story-section" id="storySection_${fd}_${index}" style="display:none">
        <div id="autoReviewStory_${fd}_${index}" class="auto-review-banner" style="display:none" data-auto-review="story"></div>
        <div class="novel-field">
          <div class="novel-field-label">📜 Full Story <span class="editable-badge">(editable)</span></div>
          <div class="story-content editable" contenteditable="true" id="storyContent_${fd}_${index}" data-flow-id="${flowId}" data-story-index="${index}"></div>
        </div>
        <div class="story-actions">
          <button class="btn btn-secondary btn-sm" onclick="downloadStory('${flowId}',${index})">📥 Download Story .txt</button>
          <button class="btn btn-audio btn-sm" onclick="generateAudioDramaScript('${flowId}',${index})" id="audioScriptBtn_${fd}_${index}">
            <span class="spinner"></span>
            <span class="btn-text">🎙️ Generate Audio Drama Script</span>
          </button>
        </div>
        <div class="audio-script-section" id="audioScriptSection_${fd}_${index}" style="display:none">
          <div class="novel-field">
            <div class="novel-field-label">🎙️ Audio Drama Script <span class="editable-badge">(edit & listen to each segment)</span></div>
            <div class="script-segments" id="audioScriptSegments_${fd}_${index}"></div>
          </div>
          <div class="story-actions">
            <button class="btn btn-secondary btn-sm" onclick="downloadAudioScript('${flowId}',${index})">📥 Download Script .txt</button>
            <button class="btn btn-audio btn-sm" onclick="generateAllAudio('${flowId}',${index})" id="generateAllAudioBtn_${fd}_${index}" title="Uses multiple keys in parallel for faster generation">
              <span class="spinner"></span>
              <span class="btn-text">🎵 Generate Audio (parallel)</span>
            </button>
            <button class="btn btn-scene btn-sm" onclick="generateAllScenes('${flowId}',${index})" id="generateAllScenesBtn_${fd}_${index}">
              <span class="spinner"></span>
              <span class="btn-text">🖼️ Generate Scenes</span>
            </button>
          </div>
        </div>
      </div>
    </div>
  `;

  return card;
}


// --- Toggle Manual Review ---
function toggleManualReview(flowId, index) {
  const flow = getFlow(flowId);
  if (!flow) return;
  const fd = flowDomId(flowId);
  if (flow.reviewedNovels.has(index)) {
    flow.reviewedNovels.delete(index);
  } else {
    flow.reviewedNovels.add(index);
  }
  const container = document.getElementById(`novelsContainer_${fd}`);
  const card = container?.querySelector(`.novel-card[data-index="${index}"]`);
  if (card) {
    const novel = flow.novels[index];
    const isExpanded = card.classList.contains('expanded');
    const newCard = createNovelCard(novel, index, flowId);
    newCard.style.animationDelay = card.style.animationDelay;
    if (isExpanded) newCard.classList.add('expanded');
    const existingStorySection = card.querySelector(`#storySection_${fd}_${index}`);
    const existingContent = card.querySelector(`#storyContent_${fd}_${index}`);
    const existingAudioSection = card.querySelector(`#audioScriptSection_${fd}_${index}`);
    if (existingStorySection && existingContent?.textContent) {
      const newStorySection = newCard.querySelector(`#storySection_${fd}_${index}`);
      const newContent = newCard.querySelector(`#storyContent_${fd}_${index}`);
      const oldAutoReview = card.querySelector(`#autoReviewStory_${fd}_${index}`);
      if (newStorySection && newContent) {
        newContent.textContent = existingContent.textContent;
        newStorySection.style.display = 'block';
        const newAutoReview = newCard.querySelector(`#autoReviewStory_${fd}_${index}`);
        if (oldAutoReview && newAutoReview && (oldAutoReview.innerHTML || '').trim()) {
          newAutoReview.innerHTML = oldAutoReview.innerHTML;
          newAutoReview.className = oldAutoReview.className;
          newAutoReview.style.display = oldAutoReview.style.display || 'block';
        }
      }
      const btn = newCard.querySelector(`#storyBtn_${fd}_${index}`);
      if (btn) {
        btn.innerHTML = '<span class="btn-text">✅ Story Generated</span>';
        btn.disabled = true;
      }
    }
    if (existingAudioSection && flow.audioScriptSegments[index]?.length) {
      const textEls = card.querySelectorAll(`.script-segment-text[data-audio-index="${index}"]`);
      textEls.forEach(el => {
        const sIdx = parseInt(el.dataset.segmentIndex, 10);
        if (!isNaN(sIdx)) syncAudioSegmentEdit(flowId, index, sIdx, el.textContent || '');
      });
      const newAudioSection = newCard.querySelector(`#audioScriptSection_${fd}_${index}`);
      const newSegmentsContainer = newCard.querySelector(`#audioScriptSegments_${fd}_${index}`);
      if (newAudioSection && newSegmentsContainer) {
        newAudioSection.style.display = 'block';
        renderAudioScriptSegments(flowId, index, newSegmentsContainer, flow.audioScriptSegments[index]);
      }
    }
    card.replaceWith(newCard);
    attachEditSyncListeners(container, flowId);
  }
  showToast(
    flow.reviewedNovels.has(index)
      ? 'Template marked as passed manual review. You can now generate the full story.'
      : 'Manual review revoked.',
    'info'
  );

  try { ensureCoverThumbInCard(flowId, index); } catch (_) {}
}

// --- Sync edits from contenteditable back to state ---
function attachEditSyncListeners(container, flowId) {
  if (!container || !flowId) return;
  const flow = getFlow(flowId);
  if (!flow) return;

  const syncEditable = (el) => {
    const chapterItem = el.closest('.chapter-item');
    const novelIndex = parseInt(el.dataset.novel ?? chapterItem?.dataset.novel, 10);
    if (isNaN(novelIndex) || !flow.novels[novelIndex]) return;

    const field = el.dataset.field;
    const value = el.textContent?.trim() || '';

    if (field) {
      if (field === 'tags') {
        const raw = value === 'N/A' ? '' : value;
        const list = raw
          ? normalizeTagList(raw.split(/[,|;]|\n/).map(t => t.trim()).filter(Boolean))
          : [];
        flow.novels[novelIndex].tags = list;
        flow.novels[novelIndex].themes = list.slice();
        flow.novels[novelIndex].tag = list.length ? list.join(', ') : '';
      } else {
        flow.novels[novelIndex][field] = value;
      }
      if (field === 'title') {
        const titleEl = container.querySelector(`.novel-card[data-index="${novelIndex}"] .novel-card-header .novel-title`);
        if (titleEl && titleEl !== el) titleEl.textContent = value || 'Untitled Novel';
      }
      if (field === 'genre' && value.includes(' | ')) {
        const parts = value.split(' | ');
        flow.novels[novelIndex].genre = (parts[0] || '').trim();
        flow.novels[novelIndex].themes = (parts[1] || '')
          .split(/[·•]/)
          .map(t => t.trim())
          .filter(Boolean);
      }
    }

    const chapterIndex = chapterItem?.dataset.chapterindex;
    if (chapterIndex !== undefined) {
      const chIndex = parseInt(chapterIndex, 10);
      const novel = flow.novels[novelIndex];
      if (novel.chapters && novel.chapters[chIndex]) {
        const f = el.dataset.field;
        if (f === 'chapterTitle') {
          novel.chapters[chIndex].title = value;
        } else if (f === 'chapterSummary') {
          novel.chapters[chIndex].summary = value;
        } else {
          const parts = value.split(' — ');
          novel.chapters[chIndex].title = (parts[0] || '').trim();
          novel.chapters[chIndex].summary = (parts[1] || '').trim();
        }
      }
    }

    const charIndex = el.dataset.charindex;
    if (charIndex !== undefined) {
      const cIndex = parseInt(charIndex, 10);
      const novel = flow.novels[novelIndex];
      if (novel.characters && novel.characters[cIndex]) {
        const text = value;
        const match = text.match(/^(.+?)\s*\(([^)]*)\)\s*[—–-]\s*(.+)$/s);
        if (match) {
          novel.characters[cIndex].name = match[1].trim();
          const roleAge = (match[2] || '').split(',');
          novel.characters[cIndex].role = (roleAge[0] || '').trim();
          novel.characters[cIndex].age = (roleAge[1] || '').trim();
          let rest = (match[3] || '').trim();
          const arcMatch = rest.match(/\bArc:\s*(.+)$/i);
          if (arcMatch) {
            novel.characters[cIndex].arc = arcMatch[1].trim();
            rest = rest.replace(/\bArc:\s*.+$/i, '').trim();
          } else {
            novel.characters[cIndex].arc = '';
          }
          novel.characters[cIndex].description = rest;
        }
      }
    }
  };

  container.addEventListener('blur', (e) => {
    const el = e.target;
    if (el.isContentEditable && (el.dataset.storyIndex !== undefined || el.dataset.audioIndex !== undefined || el.dataset.novel !== undefined || el.closest('.chapter-item[data-novel]') || el.dataset.charindex !== undefined)) {
      const storyIndex = el.dataset.storyIndex;
      const audioIndex = el.dataset.audioIndex;
      const segmentIndex = el.dataset.segmentIndex;
      if (storyIndex !== undefined) {
        const idx = parseInt(storyIndex, 10);
        if (!isNaN(idx)) flow.stories[idx] = el.textContent || '';
        return;
      }
      if (audioIndex !== undefined && segmentIndex !== undefined) {
        const aIdx = parseInt(audioIndex, 10);
        const sIdx = parseInt(segmentIndex, 10);
        if (!isNaN(aIdx) && !isNaN(sIdx)) syncAudioSegmentEdit(flowId, aIdx, sIdx, el.textContent || '');
        return;
      }
      if (audioIndex !== undefined && segmentIndex === undefined) {
        const idx = parseInt(audioIndex, 10);
        if (!isNaN(idx)) flow.audioScripts[idx] = el.textContent || '';
        return;
      }
      if (el.dataset.novel !== undefined || el.closest('.chapter-item[data-novel]') || el.dataset.charindex !== undefined) {
        syncEditable(el);
      }
    }
  }, true);

  container.addEventListener('input', (e) => {
    const el = e.target;
    if (el.isContentEditable && el.dataset.field === 'title') {
      const novelIndex = parseInt(el.dataset.novel, 10);
      if (!isNaN(novelIndex)) {
        const headerTitle = container.querySelector(`.novel-card[data-index="${novelIndex}"] .novel-card-header .novel-title`);
        if (headerTitle && headerTitle !== el) {
          headerTitle.textContent = el.textContent?.trim() || 'Untitled Novel';
        }
      }
    }
  }, true);
}


// --- Toggle Card ---
function toggleNovelCard(flowId, index) {
  const fd = flowDomId(flowId);
  const panel = document.getElementById(`flowPanel_${fd}`);
  if (!panel) return;
  panel.querySelectorAll('.novel-card').forEach((card) => {
    if (parseInt(card.dataset.index, 10) === index) {
      card.classList.toggle('expanded');
    }
  });
}

// --- Download ---
function downloadNovel(flowId, index) {
  const flow = getFlow(flowId);
  if (!flow) return;
  const novel = flow.novels[index];
  if (!novel) return;

  const content = formatNovelTxt(flowId, novel, index);
  const blob = new Blob([content], { type: 'text/plain;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `novel_${index + 1}_${sanitizeFilename(novel.title)}.txt`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
  showToast(`Downloaded: ${novel.title}`, 'success');
}

function handleDownloadAll() {
  const flow = ensureActiveFlow();
  if (!flow.novels.length) {
    showToast('No novels to download', 'error');
    return;
  }
  const fid = state.activeFlowId;
  flow.novels.forEach((_, i) => {
    setTimeout(() => downloadNovel(fid, i), i * 300);
  });
}

function formatNovelTxt(flowId, novel, index) {
  let txt = '';
  txt += `${'='.repeat(60)}\n`;
  txt += `  NOVEL TEMPLATE #${index + 1}\n`;
  txt += `${'='.repeat(60)}\n\n`;

  txt += `TITLE: ${novel.title || 'Untitled'}\n`;
  txt += `DESCRIPTION (short): ${clampText(novel.synopsis || 'N/A', 100)}\n`;
  txt += `GENRE: ${novel.genre || 'N/A'}\n`;
  txt += `AUTHOR: ${novel.authorName || 'N/A'}\n`;
  txt += `RELEASE DATE: ${novel.releaseDate || 'N/A'}\n`;
  txt += `WRITING LANGUAGE: ${novel.writingLanguage || 'N/A'}\n`;

  if (novel.themes && novel.themes.length) {
    txt += `THEMES: ${novel.themes.join(', ')}\n`;
  }
  if (safeStr(novel.thumbnailPrompt)) {
    txt += `THUMBNAIL_PROMPT: ${safeStr(novel.thumbnailPrompt)}\n`;
  }

  txt += `\n${'-'.repeat(40)}\n`;
  txt += `OVERVIEW (full meaning)\n`;
  txt += `${'-'.repeat(40)}\n`;
  txt += `${novel.overview || novel.synopsis || 'N/A'}\n`;

  txt += `\n${'-'.repeat(40)}\n`;
  txt += `TAGLINE / SHORT SYNOPSIS\n`;
  txt += `${'-'.repeat(40)}\n`;
  txt += `${novel.synopsis || 'N/A'}\n`;

  txt += `\n${'-'.repeat(40)}\n`;
  txt += `DRAFT SCRIPT & CORE IDEAS\n`;
  txt += `${'-'.repeat(40)}\n`;
  txt += `${novel.draftScript || 'N/A'}\n`;

  txt += `\n${'-'.repeat(40)}\n`;
  txt += `NARRATOR TONE\n`;
  txt += `${'-'.repeat(40)}\n`;
  txt += `${novel.narratorTone || 'N/A'}\n`;

  txt += `\n${'-'.repeat(40)}\n`;
  txt += `BACKGROUND / SETTING\n`;
  txt += `${'-'.repeat(40)}\n`;
  txt += `${novel.background || 'N/A'}\n`;

  txt += `\n${'-'.repeat(40)}\n`;
  txt += `CHARACTER SYSTEM\n`;
  txt += `${'-'.repeat(40)}\n`;
  if (novel.characters && novel.characters.length) {
    novel.characters.forEach((c, i) => {
      txt += `\n  [Character ${i + 1}]\n`;
      txt += `  Name: ${c.name}\n`;
      txt += `  Role: ${c.role}\n`;
      if (c.age) txt += `  Age: ${c.age}\n`;
      txt += `  Description: ${c.description}\n`;
      if (c.arc) txt += `  Character Arc: ${c.arc}\n`;
    });
  } else {
    txt += 'N/A\n';
  }

  txt += `\n${'-'.repeat(40)}\n`;
  txt += `CHAPTER OUTLINE\n`;
  txt += `${'-'.repeat(40)}\n`;
  if (novel.chapters && novel.chapters.length) {
    novel.chapters.forEach(ch => {
      const limited = clampChapterLine(ch.title, ch.summary, 100);
      txt += `\n  Chapter ${ch.chapterNumber}: ${limited.title || ch.title || ''}\n`;
      if (safeStr(limited.summary)) txt += `  ${limited.summary}\n`;
    });
  } else {
    txt += 'N/A\n';
  }

  const fullStory = getFullStoryText(flowId, index);
  if (fullStory) {
    txt += `\n${'-'.repeat(40)}\n`;
    txt += `FULL STORY (BY CHAPTER)\n`;
    txt += `${'-'.repeat(40)}\n\n`;
    txt += fullStory;
    txt += '\n\n';
  }

  txt += `\n${'='.repeat(60)}\n`;
  txt += `  Generated by AI Novel Template Generator\n`;
  txt += `  Date: ${new Date().toLocaleDateString()}\n`;
  txt += `${'='.repeat(60)}\n`;

  return txt;
}

// --- Utilities ---
function escapeHtml(str) {
  if (!str) return '';
  const div = document.createElement('div');
  div.textContent = str;
  return div.innerHTML;
}

function sanitizeFilename(name) {
  if (!name) return 'untitled';
  return name.toLowerCase().replace(/[^a-z0-9]+/g, '_').replace(/^_|_$/g, '').substring(0, 50);
}

function showProgress(show) {
  const container = document.getElementById('progressContainer');
  if (!container) return;
  container.classList.toggle('active', !!show);
}

function updateProgress(percent, statusText) {
  const fill = document.getElementById('progressFill');
  const status = document.getElementById('progressStatus');
  if (fill) fill.style.width = `${percent}%`;
  if (status) status.textContent = statusText;
}

function showToast(message, type = 'info') {
  const container = document.getElementById('toastContainer');
  if (!container) return;
  const toast = document.createElement('div');
  toast.className = `toast ${type}`;
  toast.textContent = message;
  toast.style.cssText = 'max-width:360px;';
  container.appendChild(toast);
  setTimeout(() => toast.remove(), 5000);
}

// --- Collapsible Section Toggle ---
function toggleSection(sectionId) {
  const body = document.getElementById(sectionId);
  if (!body) return;
  const card = body.closest('.collapsible-card');
  if (!card) return;

  const isHidden = body.style.display === 'none';
  body.style.display = isHidden ? 'block' : 'none';
  card.classList.toggle('open', isHidden);
}

// --- Generate All Stories (parallel: multiple novels at once, keyed workers) ---
async function handleGenerateAllStories(targetFlowId) {
  const flowId = targetFlowId != null ? targetFlowId : state.activeFlowId;
  const flow = getFlow(flowId);
  if (!flow) return;
  const eligible = (flow.reviewedNovels.size ? [...flow.reviewedNovels] : (flow.novels || []).map((_, i) => i))
    .filter(i => {
      const novel = flow.novels?.[i];
      if (!flow.stories[i]) return true;
      return !storyIsCompleteForNovel(flowId, i, novel);
    });
  if (!eligible.length) {
    showToast(flow.reviewedNovels.size === 0
      ? 'All templates already have complete stories'
      : 'All reviewed templates already have complete stories',
    'info');
    return;
  }

  const keys = getApiKeys();
  if (!keys.length) {
    showToast(`Add one or more API keys in Settings to generate stories.`, 'error');
    return;
  }

  const btn = document.getElementById('generateAllStoriesBtn');
  if (!btn || btn.disabled) return;
  btn.classList.add('loading');
  btn.disabled = true;

  const queue = eligible.slice();
  const keyCount = keys.length;
  const workerCount = Math.max(1, Math.min(8, queue.length, Math.max(keyCount * 2, 4)));
  let done = 0;
  const total = eligible.length;
  const failed = [];

  showToast(`Generating ${total} stories in parallel (${workerCount} concurrent workers)...`, 'info');

  const worker = async (workerIdx) => {
    const key = keys[workerIdx % keys.length];
    while (queue.length) {
      const index = queue.shift();
      if (index == null) break;
      try {
        await generateFullStory(flowId, index, key, { silent: true });
        done++;
        showToast(`Stories: ${done}/${total}`, 'info');
        await new Promise(r => setTimeout(r, 80));
      } catch (e) {
        failed.push({ index, message: e?.message || String(e) });
      }
    }
  };

  await Promise.all(Array.from({ length: workerCount }, (_, i) => worker(i)));

  btn.classList.remove('loading');
  btn.disabled = false;
  if (failed.length) {
    showToast(`Generated ${done}/${total}. Failed: ${failed.length}. Try again to retry the remaining.`, 'error');
  } else {
    showToast(`All ${total} stories generated!`, 'success');
  }
}

// --- Fill missing data (synopsis, chapter outlines) for export table ---
function novelsNeedingMissingData() {
  const flow = ensureActiveFlow();
  const flowId = state.activeFlowId;
  const indices = [];
  (flow.novels || []).forEach((novel, i) => {
    const synopsis = safeStr(novel.synopsis);
    const overview = safeStr(novel.overview);
    const needsOverview = !overview || overview === 'N/A' || overview.length < 120;
    const needsSynopsis = !synopsis || synopsis === 'N/A' || synopsis.length < 20 || synopsis.startsWith('A story:') && synopsis.length < 80;
    const chapters = novel.chapters || [];
    const needsChapters = chapters.length < 2 || (chapters.length === 1 && safeStr(chapters[0].summary).length < 50);
    const needsStoryContent = !storyIsCompleteForNovel(flowId, i, novel);
    if (needsOverview || needsSynopsis || needsChapters || needsStoryContent) indices.push(i);
  });
  return indices;
}

async function generateMissingDataForNovel(flowId, index) {
  const flow = getFlow(flowId);
  if (!flow) return;
  const fd = flowDomId(flowId);
  const novel = flow.novels[index];
  if (!novel) return;
  const prompt = `You are a novel template assistant. This novel needs complete template data for export.

**Title:** ${novel.title || 'Untitled'}
**Genre:** ${novel.genre || 'Fiction'}
**Draft/Core idea:** ${(novel.draftScript || '').substring(0, 800)}
${novel.background ? `**Setting:** ${novel.background.substring(0, 300)}` : ''}

Return valid JSON only (no markdown, no backticks) with exactly these fields:
1. "overview" — 2–5 sentences, full meaning (about 300–800 characters). Plain prose.
2. "synopsis" — A SINGLE short hook line, maximum 100 characters (including spaces). No newlines.
3. "chapters" — An array of 5-10 chapter outlines. Each item: { "chapterNumber": 1, "title": "Short title", "summary": "Short summary" }. CRITICAL: For each chapter, the combined string "title — summary" must be <= 100 characters.

Example format:
{"overview": "Long overview...", "synopsis": "Short hook here...", "chapters": [{"chapterNumber": 1, "title": "...", "summary": "..."}, ...]}`;

  const result = await callGeminiAPI(prompt);
  if (result.overview) novel.overview = result.overview;
  if (result.synopsis) novel.synopsis = result.synopsis;
  if (result.chapters && Array.isArray(result.chapters) && result.chapters.length >= 2) {
    novel.chapters = result.chapters.map(ch => ({
      chapterNumber: ch.chapterNumber || 0,
      title: ch.title || '',
      summary: ch.summary || '',
    })).filter(ch => ch.chapterNumber >= 1).sort((a, b) => a.chapterNumber - b.chapterNumber);
  }
  normalizeNovelsForExport([novel]);
  const container = document.getElementById(`novelsContainer_${fd}`);
  const card = container?.querySelector(`.novel-card[data-index="${index}"]`);
  if (card) {
    const newCard = createNovelCard(novel, index, flowId);
    newCard.classList.add('expanded');
    card.replaceWith(newCard);
    attachEditSyncListeners(container, flowId);
    ensureCoverThumbInCard(flowId, index);
    if (flow.stories[index]) {
      const storySection = document.getElementById(`storySection_${fd}_${index}`);
      const storyContent = document.getElementById(`storyContent_${fd}_${index}`);
      if (storySection && storyContent) {
        renderStoryChapters(flowId, index, flow.stories[index]);
        storySection.style.display = 'block';
      }
      const storyBtn = document.getElementById(`storyBtn_${fd}_${index}`);
      if (storyBtn) storyBtn.innerHTML = '<span class="btn-text">✅ Story Generated</span>';
    }
    newCard.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }
}

async function handleFillMissingData() {
  const flowId = state.activeFlowId;
  const flow = ensureActiveFlow();
  const indices = novelsNeedingMissingData();
  if (!indices.length) {
    showToast('All templates already have synopsis, chapter outlines, and full chapter content.', 'success');
    return;
  }
  const keys = getApiKeys();
  if (!keys.length) {
    showToast('Add an API key in Settings to generate missing data.', 'error');
    return;
  }
  const btn = document.getElementById('fillMissingDataBtn');
  if (btn) { btn.classList.add('loading'); btn.disabled = true; }
  showToast(`Filling missing data for ${indices.length} novel(s)...`, 'info');
  let done = 0;
  let keyCursor = 0;
  for (const i of indices) {
    try {
      await generateMissingDataForNovel(flowId, i);
      // Also ensure chapter content exists for export (episodeContent).
      const k = keys[keyCursor % keys.length];
      keyCursor++;
      if (!storyIsCompleteForNovel(flowId, i, flow.novels?.[i])) {
        await generateFullStory(flowId, i, k);
      }
      done++;
      showToast(`Filled ${done}/${indices.length}`, 'info');
      await new Promise(r => setTimeout(r, 400));
    } catch (e) {
      console.warn('Fill missing data failed', i, e);
      showToast(`Failed for novel ${i + 1}: ${e?.message || 'Unknown error'}`, 'error');
    }
  }
  if (btn) { btn.classList.remove('loading'); btn.disabled = false; }
  showToast(`Done. Filled missing data for ${done} novel(s).`, 'success');
}

// --- Generate Full Story (safe to run for several indices concurrently; batch uses silent: true) ---
async function generateFullStory(flowId, index, apiKeyOverride, options = {}) {
  const silent = Boolean(options.silent);
  const flow = getFlow(flowId);
  if (!flow) return;
  const fd = flowDomId(flowId);
  const novel = flow.novels[index];
  if (!novel) return;

  const apiKey = apiKeyOverride || getApiKey();
  if (!apiKey) {
    showToast('Please enter your API key first', 'error');
    return;
  }

  const btn = document.getElementById(`storyBtn_${fd}_${index}`);
  if (btn) {
    btn.classList.add('loading');
    btn.disabled = true;
  }

  const panel = document.getElementById(`flowPanel_${fd}`);
  const card = panel?.querySelector(`.novel-card[data-index="${index}"]`);
  if (card && !card.classList.contains('expanded')) {
    card.classList.add('expanded');
  }

  if (!silent) {
    showToast(`Generating full story for "${novel.title}"... This may take a moment.`, 'info');
  }

  try {
    enforceSinoVietnameseCharacterNames(novel);
    // Build chapter details for the prompt
    let chapterDetails = '';
    if (novel.chapters && novel.chapters.length) {
      chapterDetails = novel.chapters.map(ch =>
        `Chapter ${ch.chapterNumber}: "${ch.title}" — ${ch.summary}`
      ).join('\n');
    }

    let characterDetails = '';
    if (novel.characters && novel.characters.length) {
      characterDetails = novel.characters.map(c =>
        `- ${c.name} (${c.role}${c.gender ? `, ${c.gender}` : ''}): ${c.description}${c.arc ? ' | Arc: ' + c.arc : ''}`
      ).join('\n');
    }

    const themesLine = Array.isArray(novel.themes) && novel.themes.length
      ? novel.themes.join(', ')
      : 'Not specified';

    const storyPrompt = `You are an expert novelist and creative writer. Write the FULL STORY for the following novel, expanding the TEMPLATE below into complete chapter prose. Honor the outline, characters, tone, and themes unless you must fix a clear contradiction.

## NOVEL DETAILS (from template)
**Title:** ${novel.title}
**Author (byline):** ${novel.authorName || 'Not specified'}
**Genre:** ${novel.genre || 'Fiction'}
**Category:** ${novel.category || 'Fiction'}
**Writing Language:** ${novel.writingLanguage || 'English'}
**Themes:** ${themesLine}
**Narrator Tone:** ${novel.narratorTone || 'Third-person omniscient'}
**Background/Setting:** ${novel.background || 'Not specified'}

## OVERVIEW (full meaning — follow this premise)
${novel.overview || novel.synopsis || 'Not provided'}

## SHORT TAGLINE (listings)
${novel.synopsis || 'Not provided'}

## DRAFT SCRIPT & CORE IDEAS
${novel.draftScript || 'Not provided'}

## CHARACTERS
${characterDetails || 'Not specified'}

## CHAPTER OUTLINE (follow this structure and chapter count)
${chapterDetails || 'Write 5-10 chapters'}

## OUTPUT FORMAT (CRITICAL)
You MUST output the story separated into chapters using EXACT markers like this, for EVERY chapter in the outline (same count and order as the outline above):

[CHAPTER 1]
Title: <chapter title>
<full chapter prose here>
[/CHAPTER 1]

[CHAPTER 2]
Title: <chapter title>
<full chapter prose here>
[/CHAPTER 2]

Rules:
- Do NOT output anything before the first [CHAPTER 1] marker.
- Do NOT add meta commentary, notes, or explanations.
- Chapter prose should be publish-ready, with dialogue and vivid description.
- Write in ${novel.writingLanguage || 'English'}.
- Match the number of [CHAPTER n] blocks to the chapter outline (if the outline lists 7 chapters, output 7 blocks).
- CHARACTER SYNC: Every character listed under CHARACTERS must appear in the story by name (or clear nickname established in prose), with roles and arcs consistent with the template—do not replace them with different people.
- TEMPLATE FIDELITY: The plot, stakes, and setting must follow the DRAFT SCRIPT & CORE IDEAS and CHAPTER OUTLINE; do not invent an unrelated story.
${isVietnameseWritingLanguage(novel.writingLanguage) ? '- NAME STYLE RULE: Keep all character names in Sino-Vietnamese Chinese style exactly as listed in the template (e.g., Lâm Hải, Lâm Vi). Do NOT replace them with Western names.' : ''}
`;

    const storyText = await callGeminiAPIRawWithKey(storyPrompt, apiKeyOverride);

    flow.stories[index] = storyText;

    const storySection = document.getElementById(`storySection_${fd}_${index}`);
    const storyContent = document.getElementById(`storyContent_${fd}_${index}`);
    renderStoryChapters(flowId, index, storyText);
    if (storySection) storySection.style.display = 'block';

    await runAutoReviewStoryStep(flowId, index, apiKeyOverride).catch((e) => console.warn('Auto-review story', e));

    // Update button
    if (btn) {
      updateStoryButtonReviewState(flowId, index);
      btn.classList.remove('loading');
    }

    if (!silent) {
      showToast(`Full story generated for "${novel.title}"!`, 'success');
      if (storySection) storySection.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    }

  } catch (error) {
    console.error('Story generation error:', error);
    const msg = error?.message || String(error);
    const provider = getAIProvider() === 'deepseek' ? 'DeepSeek' : 'Gemini';
    showToast(`Story failed (#${index + 1} ${novel.title || ''}): ${msg} (${provider})`, 'error');
    if (btn) {
      btn.classList.remove('loading');
      btn.disabled = false;
      const txt = btn.querySelector('.btn-text');
      if (txt) txt.textContent = '📖 Generate Full Story';
    }
  }
}

// --- Raw text API call (for story, script — routes to Gemini or DeepSeek) ---
async function callGeminiAPIRaw(prompt) {
  if (getAIProvider() === 'deepseek') {
    return callDeepSeekAPI(prompt, false);
  }
  const apiKey = getApiKey();
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;
  const response = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      contents: [{ parts: [{ text: prompt }] }],
      generationConfig: {
        temperature: 0.85,
        topP: 0.95,
        topK: 40,
        maxOutputTokens: 65536,
      },
    }),
  });
  if (!response.ok) {
    const errData = await response.json().catch(() => ({}));
    throw new Error(errData?.error?.message || `API error: ${response.status}`);
  }
  const data = await response.json();
  const text = data.candidates?.[0]?.content?.parts?.[0]?.text;
  if (!text) throw new Error('No content returned from Gemini API');
  return text;
}

// --- Raw text API call with an optional explicit key (for multi-flow parallel generation) ---
async function callGeminiAPIRawWithKey(prompt, apiKeyOverride) {
  if (getAIProvider() === 'deepseek') {
    const apiKey = apiKeyOverride || getApiKey();
    if (!apiKey) throw new Error('No API key. Enter your DeepSeek key in the API Key(s) field.');
    const url = 'https://api.deepseek.com/v1/chat/completions';
    const body = {
      model: 'deepseek-chat',
      messages: [{ role: 'user', content: prompt }],
      temperature: 0.85,
      max_tokens: 8192,
    };
    const response = await fetchWithTimeout(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${apiKey}`,
      },
      body: JSON.stringify(body),
    }, 120000);
    if (!response.ok) {
      const err = await response.json().catch(() => ({}));
      throw new Error(err?.error?.message || err?.message || `DeepSeek API error: ${response.status}`);
    }
    const data = await response.json();
    const text = data.choices?.[0]?.message?.content?.trim();
    if (!text) throw new Error('No content returned from DeepSeek');
    return text;
  }

  const apiKey = apiKeyOverride || getApiKey();
  if (!apiKey) throw new Error('No API key. Enter your Gemini key in the API Key(s) field.');
  const models = ['gemini-2.5-flash', 'gemini-2.0-flash', 'gemini-1.5-flash'];
  let lastErr = null;
  for (const model of models) {
    try {
      const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${encodeURIComponent(apiKey)}`;
      const response = await fetchWithTimeout(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          contents: [{ parts: [{ text: prompt }] }],
          generationConfig: {
            temperature: 0.85,
            topP: 0.95,
            topK: 40,
            maxOutputTokens: 65536,
          },
        }),
      }, 120000);
      if (!response.ok) {
        const errData = await response.json().catch(() => ({}));
        lastErr = errData?.error?.message || errData?.message || `API error: ${response.status}`;
        if (String(lastErr).includes('404') || String(lastErr).toLowerCase().includes('not found') || String(lastErr).includes('Invalid model')) continue;
        throw new Error(lastErr);
      }
      const data = await response.json();
      const cand = data.candidates?.[0];
      const text = cand?.content?.parts?.[0]?.text;
      if (!text) {
        const reason = cand?.finishReason || cand?.finishReasonReason || '';
        if (String(reason).toLowerCase().includes('safety')) throw new Error('Response blocked by safety filters. Try a different prompt or model.');
        if (String(reason).toLowerCase().includes('max')) throw new Error('Story too long; output was truncated. Try fewer chapters.');
        throw new Error('No content returned from Gemini API. Try again.');
      }
      return text;
    } catch (e) {
      lastErr = e?.message || lastErr;
      const m = String(lastErr || '');
      if (m.includes('404') || m.toLowerCase().includes('not found') || m.includes('Invalid model')) continue;
      throw e;
    }
  }
  throw new Error(lastErr || 'All Gemini models failed');
}

// --- Render full story into per-chapter blocks (using [CHAPTER N] markers) ---
function normalizeNewlines(s) {
  return String(s || '').replace(/\r\n/g, '\n').replace(/\r/g, '\n');
}

function parseChaptersFromMarkers(text) {
  const t = normalizeNewlines(text);
  const re = /^\[CHAPTER\s+(\d+)\]\s*\n([\s\S]*?)^\[\/CHAPTER\s+\1\]\s*$/gmi;
  const chapters = [];
  let m;
  while ((m = re.exec(t))) {
    const num = parseInt(m[1], 10);
    const body = (m[2] || '').trim();
    let title = '';
    let content = body;
    const titleMatch = body.match(/^Title:\s*(.+)\n+/i);
    if (titleMatch) {
      title = titleMatch[1].trim();
      content = body.slice(titleMatch[0].length).trim();
    }
    chapters.push({ number: num, title, content });
  }
  chapters.sort((a, b) => (a.number || 0) - (b.number || 0));
  return chapters;
}

function renderStoryChapters(flowId, index, storyText) {
  const fd = flowDomId(flowId);
  const storyContent = document.getElementById(`storyContent_${fd}_${index}`);
  if (!storyContent) return;

  const chapters = parseChaptersFromMarkers(storyText || '');
  if (!chapters.length) {
    storyContent.setAttribute('contenteditable', 'true');
    storyContent.textContent = storyText || '';
    return;
  }

  storyContent.setAttribute('contenteditable', 'false');
  storyContent.innerHTML = '';
  const wrap = document.createElement('div');
  wrap.className = 'chapters-wrap';
  chapters.forEach((ch) => {
    const block = document.createElement('div');
    block.className = 'chapter-block';
    const hdr = document.createElement('div');
    hdr.className = 'chapter-hdr';

    const title = document.createElement('div');
    title.className = 'chapter-title editable';
    title.setAttribute('contenteditable', 'true');
    title.textContent = `Chapter ${ch.number}${ch.title ? ': ' + ch.title : ''}`;

    const meta = document.createElement('div');
    meta.className = 'chapter-meta';
    meta.textContent = ch.content ? `${Math.max(1, ch.content.split(/\s+/).filter(Boolean).length)} words` : '';

    const body = document.createElement('div');
    body.className = 'chapter-body editable';
    body.setAttribute('contenteditable', 'true');
    body.textContent = ch.content || '';

    hdr.appendChild(title);
    hdr.appendChild(meta);
    block.appendChild(hdr);
    block.appendChild(body);
    wrap.appendChild(block);
  });
  storyContent.appendChild(wrap);
}

// --- Download Full Story (uses current edited content from DOM) ---
function downloadStory(flowId, index) {
  const flow = getFlow(flowId);
  if (!flow) return;
  const fd = flowDomId(flowId);
  const novel = flow.novels[index];
  const contentEl = document.getElementById(`storyContent_${fd}_${index}`);
  const storyText = contentEl?.textContent?.trim() || flow.stories[index];
  if (!novel || !storyText) {
    showToast('No story to download. Generate it first.', 'error');
    return;
  }

  let content = '';
  content += `${'='.repeat(60)}\n`;
  content += `  ${novel.title || 'Untitled'}\n`;
  content += `  by ${novel.authorName || 'Unknown Author'}\n`;
  content += `${'='.repeat(60)}\n\n`;
  content += storyText;
  content += `\n\n${'='.repeat(60)}\n`;
  content += `  Generated by AI Novel Template Generator\n`;
  content += `  Date: ${new Date().toLocaleDateString()}\n`;
  content += `${'='.repeat(60)}\n`;

  const blob = new Blob([content], { type: 'text/plain;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `story_${index + 1}_${sanitizeFilename(novel.title)}.txt`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
  showToast(`Downloaded story: ${novel.title}`, 'success');
}

// --- Generate Audio Drama Script ---
async function generateAudioDramaScript(flowId, index) {
  const flow = getFlow(flowId);
  if (!flow) return;
  const fd = flowDomId(flowId);
  const novel = flow.novels[index];
  const contentEl = document.getElementById(`storyContent_${fd}_${index}`);
  const storyText = contentEl?.textContent?.trim() || flow.stories[index];
  if (!novel || !storyText) {
    showToast('Generate the full story first.', 'error');
    return;
  }

  const apiKey = getApiKey();
  if (!apiKey) {
    showToast('Please enter your API key first', 'error');
    return;
  }

  const btn = document.getElementById(`audioScriptBtn_${fd}_${index}`);
  btn.classList.add('loading');
  btn.disabled = true;
  showToast(`Generating audio drama script for "${novel.title}"...`, 'info');

  try {
    const characterNames = (novel.characters || []).map(c => c.name).filter(Boolean);
    const characterList = (novel.characters || []).map(c => `${c.name}${c.gender ? ` (${c.gender})` : ''}`).filter(Boolean);
    const prompt = `You are an expert audio drama and radio play scriptwriter. Convert the following novel/story into an AUDIO DRAMA SCRIPT format suitable for voice actors and audio production.

## STORY TO CONVERT
${storyText.substring(0, 40000)}

## CHARACTERS (use EXACT names for dialogue labels; gender used for voice casting)
${characterList.length ? characterList.join(', ') : characterNames.join(', ') || 'Extract from the story'}

## OUTPUT FORMAT REQUIREMENTS
Create a script with:

1. **Scene headers:** [SCENE: Location/Description] or [INT. LOCATION - TIME]
2. **Narrator lines:** NARRATOR: [text]
3. **Character dialogue:** CHARACTER NAME: [dialogue]
4. **Sound effects cues:** [SFX: description] 
5. **Music cues:** [MUSIC: mood/description]
6. **Ambience:** [AMB: environment sound]

Format rules:
- ONE logical unit per line (each line = one segment for playback)
- One speaker per line, with name in UPPERCASE followed by colon
- Include [SFX], [MUSIC], [AMB] as separate lines where appropriate
- Keep prose descriptions minimal; focus on dialogue and audio cues
- Preserve emotional beats as parentheticals (e.g., (sadly), (whispering))
- Use single newlines between segments; no double newlines within the script
- Write in ${novel.writingLanguage || 'English'}
- Output ONLY the script, no meta-commentary`;

    const scriptText = await callGeminiAPIRaw(prompt);
    const segments = scriptText.split(/\r?\n/).map(s => s.trim()).filter(Boolean);
    flow.audioScriptSegments[index] = segments;
    flow.audioScripts[index] = scriptText;

    const scriptSection = document.getElementById(`audioScriptSection_${fd}_${index}`);
    const segmentsContainer = document.getElementById(`audioScriptSegments_${fd}_${index}`);
    renderAudioScriptSegments(flowId, index, segmentsContainer, segments);
    scriptSection.style.display = 'block';
    scriptSection.scrollIntoView({ behavior: 'smooth', block: 'start' });

    btn.innerHTML = '<span class="btn-text">✅ Script Generated</span>';
    btn.classList.remove('loading');
    showToast(`Audio drama script generated for "${novel.title}"!`, 'success');
  } catch (error) {
    console.error('Audio script generation error:', error);
    showToast(`Script generation failed: ${error.message}`, 'error');
    btn.classList.remove('loading');
    btn.disabled = false;
  }
}

// --- Find batch key for segment index (batch where segment is first) ---
function getBatchAudioForSegment(flowId, audioIndex, segmentIndex) {
  const flow = getFlow(flowId);
  if (!flow) return null;
  const batches = flow.generatedAudioBatches[audioIndex] || {};
  const audio = flow.generatedAudio[audioIndex] || {};
  for (const [key, indices] of Object.entries(batches)) {
    if (indices[0] === segmentIndex) return { url: audio[key], key, indices };
  }
  return null;
}

// --- Render Audio Script Segments (each editable + Listen + Generate Audio/Scene) ---
function renderAudioScriptSegments(flowId, audioIndex, container, segments) {
  if (!container) return;
  const fd = flowDomId(flowId);
  const flow = getFlow(flowId);
  if (!flow) return;
  container.innerHTML = '';
  const audioBlobs = flow.generatedAudio[audioIndex] || {};
  const sceneImages = flow.generatedScenes[audioIndex] || {};
  (segments || []).forEach((text, i) => {
    const seg = document.createElement('div');
    seg.className = 'script-segment';
    seg.dataset.audioIndex = String(audioIndex);
    seg.dataset.segmentIndex = String(i);
    seg.dataset.flowId = flowId;
    const isSceneCue = /^\[(SCENE|INT\.|EXT\.)[^\]]*\]/i.test(text);
    const hasAudio = !!audioBlobs[i];
    const batchInfo = getBatchAudioForSegment(flowId, audioIndex, i);
    const hasBatchAudio = !!batchInfo;
    const hasScene = !!sceneImages[i] && isSceneCue;
    seg.innerHTML = `
      <div class="script-segment-row">
        <span class="segment-num">${i + 1}</span>
        <div class="segment-actions">
          <button type="button" class="btn btn-icon segment-listen" onclick="listenToSegment('${flowId}',${audioIndex}, ${i})" id="listenBtn_${fd}_${audioIndex}_${i}" title="Listen (browser TTS)">🔊</button>
          <button type="button" class="btn btn-icon segment-gen-audio" onclick="generateAudioForSegment('${flowId}',${audioIndex}, ${i})" id="genAudioBtn_${fd}_${audioIndex}_${i}" title="Generate AI audio">🎵</button>
          ${isSceneCue ? `<button type="button" class="btn btn-icon segment-gen-scene" onclick="generateSceneForSegment('${flowId}',${audioIndex}, ${i})" id="genSceneBtn_${fd}_${audioIndex}_${i}" title="Generate scene image">🖼️</button>` : ''}
        </div>
        <div class="script-segment-text editable" contenteditable="true" data-flow-id="${flowId}" data-audio-index="${audioIndex}" data-segment-index="${i}">${escapeHtml(text)}</div>
      </div>
      ${hasAudio ? `<div class="segment-generated-audio"><audio controls src="${audioBlobs[i]}" id="audioPlayer_${fd}_${audioIndex}_${i}"></audio><a href="${audioBlobs[i]}" download="segment_${i + 1}.${getTtsProvider() === 'ai33pro' ? 'mp3' : 'wav'}" class="btn btn-icon">📥</a></div>` : ''}
      ${hasBatchAudio ? `<div class="segment-generated-audio batch-audio"><span class="batch-label">Segments ${batchInfo.indices[0] + 1}–${batchInfo.indices[batchInfo.indices.length - 1] + 1}</span><audio controls src="${batchInfo.url}" id="audioBatch_${fd}_${audioIndex}_${i}"></audio><a href="${batchInfo.url}" download="batch_${batchInfo.indices[0] + 1}-${batchInfo.indices[batchInfo.indices.length - 1] + 1}.${getTtsProvider() === 'ai33pro' ? 'mp3' : 'wav'}" class="btn btn-icon">📥</a></div>` : ''}
      ${hasScene ? `<div class="segment-generated-scene"><img src="${sceneImages[i]}" alt="Scene ${i + 1}"/><a href="${sceneImages[i]}" download="scene_${i + 1}.png" class="btn btn-icon">📥</a></div>` : ''}
    `;
    container.appendChild(seg);
  });
}

// --- TTS: Listen to a single segment ---
function listenToSegment(flowId, audioIndex, segmentIndex) {
  const flow = getFlow(flowId);
  if (!flow) return;
  const fd = flowDomId(flowId);
  const segments = flow.audioScriptSegments[audioIndex];
  if (!segments || !segments[segmentIndex]) return;
  const raw = segments[segmentIndex];
  const text = stripTextForTTS(raw);
  if (!text) return;

  if (state.speakingSegment) {
    speechSynthesis.cancel();
  }

  const utterance = new SpeechSynthesisUtterance(text);
  const novel = flow.novels[audioIndex];
  const lang = (novel?.writingLanguage || 'en').substring(0, 2).toLowerCase();
  const langMap = { vietnamese: 'vi-VN', english: 'en-US', japanese: 'ja-JP', korean: 'ko-KR', chinese: 'zh-CN', french: 'fr-FR', spanish: 'es-ES', german: 'de-DE', portuguese: 'pt-BR', thai: 'th-TH' };
  utterance.lang = langMap[lang] || 'en-US';
  utterance.rate = 1;
  utterance.pitch = 1;

  state.speakingSegment = { flowId, audioIndex, segmentIndex };
  const btn = document.getElementById(`listenBtn_${fd}_${audioIndex}_${segmentIndex}`);
  if (btn) btn.classList.add('playing');

  utterance.onend = utterance.onerror = () => {
    state.speakingSegment = null;
    if (btn) btn.classList.remove('playing');
  };

  speechSynthesis.speak(utterance);
}

// --- Sync segment edits back to state ---
function syncAudioSegmentEdit(flowId, audioIndex, segmentIndex, newText) {
  const flow = getFlow(flowId);
  if (!flow || !flow.audioScriptSegments[audioIndex]) return;
  const segs = flow.audioScriptSegments[audioIndex];
  if (segmentIndex >= 0 && segmentIndex < segs.length) {
    segs[segmentIndex] = newText;
    flow.audioScripts[audioIndex] = segs.join('\n');
  }
}

// --- PCM to WAV (for Gemini TTS output) ---
function pcmToWavBlob(pcmBase64, sampleRate = 24000) {
  const pcm = Uint8Array.from(atob(pcmBase64), c => c.charCodeAt(0));
  const numChannels = 1;
  const bitsPerSample = 16;
  const byteRate = sampleRate * numChannels * (bitsPerSample / 8);
  const dataSize = pcm.length;
  const buffer = new ArrayBuffer(44 + dataSize);
  const view = new DataView(buffer);
  let offset = 0;
  const write = (str) => { for (let i = 0; i < str.length; i++) view.setUint8(offset++, str.charCodeAt(i)); };
  write('RIFF');
  view.setUint32(offset, 36 + dataSize, true); offset += 4;
  write('WAVE');
  write('fmt ');
  view.setUint32(offset, 16, true); offset += 4; // chunk size
  view.setUint16(offset, 1, true); offset += 2;  // PCM
  view.setUint16(offset, numChannels, true); offset += 2;
  view.setUint32(offset, sampleRate, true); offset += 4;
  view.setUint32(offset, byteRate, true); offset += 4;
  view.setUint16(offset, numChannels * (bitsPerSample / 8), true); offset += 2;
  view.setUint16(offset, bitsPerSample, true); offset += 2;
  write('data');
  view.setUint32(offset, dataSize, true); offset += 4;
  new Uint8Array(buffer).set(pcm, 44);
  return new Blob([buffer], { type: 'audio/wav' });
}

/** Parse Gemini TTS generateContent JSON — audio may be inlineData or inline_data, any part index. */
function geminiTtsResponseToAudioBlob(data) {
  const parts = data?.candidates?.[0]?.content?.parts || [];
  for (let i = 0; i < parts.length; i++) {
    const p = parts[i];
    const id = p?.inlineData || p?.inline_data;
    if (!id?.data) continue;
    const mime = String(id.mimeType || id.mime_type || '').toLowerCase();
    const b64 = id.data;
    try {
      if (mime.includes('mpeg') || mime.includes('mp3')) {
        const bin = Uint8Array.from(atob(b64), c => c.charCodeAt(0));
        return new Blob([bin], { type: 'audio/mpeg' });
      }
      if (mime.includes('wav') || mime.includes('wave')) {
        const bin = Uint8Array.from(atob(b64), c => c.charCodeAt(0));
        return new Blob([bin], { type: 'audio/wav' });
      }
      if (mime.includes('ogg') || mime.includes('opus')) {
        const bin = Uint8Array.from(atob(b64), c => c.charCodeAt(0));
        return new Blob([bin], { type: mime.includes('ogg') ? 'audio/ogg' : 'audio/opus' });
      }
      return pcmToWavBlob(b64);
    } catch (e) {
      console.warn('TTS blob decode fallback to PCM WAV', e);
      return pcmToWavBlob(b64);
    }
  }
  return null;
}

function geminiTtsErrorDetail(data) {
  const c = data?.candidates?.[0];
  const fr = c?.finishReason;
  if (fr && fr !== 'STOP' && fr !== 'FINISH_REASON_STOP') {
    return `TTS stopped: ${fr}`;
  }
  const pt = c?.content?.parts;
  if (!pt?.length) return 'No content parts in TTS response';
  return 'No audio payload in TTS response (check model and API access)';
}

// --- Strip cue labels for TTS (narrator won't read [AMB:], [SFX:], etc.) ---
function stripTextForTTS(text) {
  if (!text?.trim()) return '';
  return text
    .replace(/\[(?:AMB|SFX|MUSIC|SCENE|INT\.|EXT\.)\s*:\s*([^\]]*)\]/gim, '$1')
    .trim();
}

// --- Parse speaker from segment (e.g. "NARRATOR: text" or "ALICE: Hello") ---
function parseSpeakerFromSegment(segment) {
  const m = segment.match(/^([A-Z][A-Z\s]*)\s*:\s*(.*)$/s);
  if (!m) return { speaker: null, text: segment };
  return { speaker: m[1].trim().toUpperCase(), text: (m[2] || '').trim() };
}

// --- Per-novel TTS overrides (from "Suggest TTS voices from template") + Settings defaults ---
function getTtsVoiceSettingsForNovel(novel) {
  const narrator = document.getElementById('narratorVoice')?.value || 'Charon';
  const female = document.getElementById('femaleVoice')?.value || 'Kore';
  const male = document.getElementById('maleVoice')?.value || 'Puck';
  const tv = novel?.ttsVoices;
  if (!tv || typeof tv !== 'object') return { narrator, female, male };
  return {
    narrator: tv.narratorVoice || narrator,
    female: tv.femaleVoice || female,
    male: tv.maleVoice || male,
  };
}

// --- Get voice for segment: narrator vs character (match gender — no male voice for female chars) ---
function getVoiceForSegment(segmentText, novel) {
  const { speaker } = parseSpeakerFromSegment(segmentText);
  const v = getTtsVoiceSettingsForNovel(novel);
  if (!speaker || speaker === 'NARRATOR') {
    return v.narrator;
  }
  const chars = novel?.characters || [];
  const speakerNorm = speaker.replace(/\s+/g, ' ').toUpperCase();
  const char = chars.find(c => {
    if (!c.name) return false;
    const nameNorm = c.name.trim().toUpperCase();
    return nameNorm === speakerNorm || speakerNorm.startsWith(nameNorm) || nameNorm.startsWith(speakerNorm);
  });
  const gender = (char?.gender || '').toLowerCase();
  if (gender === 'female') {
    return v.female;
  }
  if (gender === 'male') {
    return v.male;
  }
  return v.narrator;
}

// --- Throttle delay (ms) between TTS API calls to avoid rate limit (15 RPM free tier) ---
const TTS_THROTTLE_MS = 4200;

// --- Build batches for multi-speaker TTS (max 2 speakers per batch = 1 API call) ---
function buildAudioBatches(segments, novel) {
  const batches = [];
  let currentBatch = { indices: [], speakers: new Set(), lines: [] };
  const getSpeaker = (raw) => {
    const { speaker } = parseSpeakerFromSegment(raw);
    return speaker || 'NARRATOR';
  };
  const voiceSet = getTtsVoiceSettingsForNovel(novel);
  const getVoice = (speaker) => {
    if (!speaker || speaker === 'NARRATOR') return voiceSet.narrator;
    const chars = novel?.characters || [];
    const speakerNorm = speaker.replace(/\s+/g, ' ').toUpperCase();
    const char = chars.find(c => {
      if (!c.name) return false;
      const nameNorm = c.name.trim().toUpperCase();
      return nameNorm === speakerNorm || speakerNorm.startsWith(nameNorm) || nameNorm.startsWith(speakerNorm);
    });
    const gender = (char?.gender || '').toLowerCase();
    if (gender === 'female') return voiceSet.female;
    if (gender === 'male') return voiceSet.male;
    return voiceSet.narrator;
  };
  for (let i = 0; i < segments.length; i++) {
    const raw = segments[i];
    const text = stripTextForTTS(raw);
    if (!text) continue;
    const speaker = getSpeaker(raw);
    const wouldBeNewSpeaker = !currentBatch.speakers.has(speaker);
    const wouldExceedTwo = currentBatch.speakers.size >= 2 && wouldBeNewSpeaker;
    if (currentBatch.indices.length > 0 && wouldExceedTwo) {
      const prompt = currentBatch.lines.map(([s, t]) => `${s}: ${t}`).join('\n');
      const voiceMap = {};
      currentBatch.speakers.forEach(s => { voiceMap[s] = getVoice(s); });
      batches.push({ indices: [...currentBatch.indices], prompt, voiceMap });
      currentBatch = { indices: [], speakers: new Set(), lines: [] };
    }
    currentBatch.indices.push(i);
    currentBatch.speakers.add(speaker);
    const { text: lineText } = parseSpeakerFromSegment(raw);
    currentBatch.lines.push([speaker, stripTextForTTS(raw)]);
  }
  if (currentBatch.indices.length > 0) {
    const prompt = currentBatch.lines.map(([s, t]) => `${s}: ${t}`).join('\n');
    const voiceMap = {};
    currentBatch.speakers.forEach(s => { voiceMap[s] = getVoice(s); });
    batches.push({ indices: [...currentBatch.indices], prompt, voiceMap });
  }
  return batches;
}

// --- Call AI33 Pro TTS (OpenAI-compatible API) ---
async function callAi33TTS(text, voiceName, apiKeyOverride = null) {
  const apiKey = apiKeyOverride || getTtsApiKey();
  if (!apiKey) throw new Error('AI33 Pro API key required. Set it in Settings.');
  const baseUrl = (document.getElementById('ai33BaseUrl')?.value || 'https://api.ai33.pro/v1').replace(/\/$/, '');
  const normalizedVoice = encodeURIComponent(String(voiceName || 'alloy').trim());
  const url = `${baseUrl}/text-to-speech/${normalizedVoice}?output_format=mp3_44100_128`;
  // #region agent log
  fetch('http://127.0.0.1:7906/ingest/260818ca-86e0-4d11-8830-b6af7bbca1a1', { method: 'POST', headers: { 'Content-Type': 'application/json', 'X-Debug-Session-Id': 'db5bb1' }, body: JSON.stringify({ sessionId: 'db5bb1', hypothesisId: 'H7', location: 'app.js:callAi33TTS:pre', message: 'AI33 TTS request v2', data: { baseUrlLen: baseUrl.length, path: '/text-to-speech/{voice_id}', voice: voiceName || 'alloy', inputLen: String(text || '').length, hasKey: !!apiKey, hasOverride: !!apiKeyOverride }, timestamp: Date.now() }) }).catch(() => {});
  // #endregion
  // #region agent log
  fetch('http://127.0.0.1:7906/ingest/260818ca-86e0-4d11-8830-b6af7bbca1a1', { method: 'POST', headers: { 'Content-Type': 'application/json', 'X-Debug-Session-Id': 'db5bb1' }, body: JSON.stringify({ sessionId: 'db5bb1', hypothesisId: 'H6', location: 'app.js:callAi33TTS:keyMeta', message: 'AI33 key diagnostics', data: { keyLen: String(apiKey || '').length, hasInnerWhitespace: /\s/.test(String(apiKey || '')), startsWithSk: String(apiKey || '').toLowerCase().startsWith('sk-'), startsWithBearer: String(apiKey || '').toLowerCase().startsWith('bearer '), hasQuote: /["']/.test(String(apiKey || '')) }, timestamp: Date.now() }) }).catch(() => {});
  // #endregion
  const response = await fetchWithTimeout(url, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'xi-api-key': apiKey,
    },
    body: JSON.stringify({
      text: String(text || '').slice(0, 4096),
      model_id: 'eleven_multilingual_v2',
      with_transcript: false,
    }),
  }, 120000);
  // #region agent log
  fetch('http://127.0.0.1:7906/ingest/260818ca-86e0-4d11-8830-b6af7bbca1a1', { method: 'POST', headers: { 'Content-Type': 'application/json', 'X-Debug-Session-Id': 'db5bb1' }, body: JSON.stringify({ sessionId: 'db5bb1', hypothesisId: 'H2', location: 'app.js:callAi33TTS:response', message: 'AI33 HTTP response', data: { ok: response.ok, status: response.status, contentType: (response.headers.get('content-type') || '').slice(0, 80) }, timestamp: Date.now() }) }).catch(() => {});
  // #endregion
  // #region agent log
  fetch('http://127.0.0.1:7906/ingest/260818ca-86e0-4d11-8830-b6af7bbca1a1', { method: 'POST', headers: { 'Content-Type': 'application/json', 'X-Debug-Session-Id': 'db5bb1' }, body: JSON.stringify({ sessionId: 'db5bb1', hypothesisId: 'H2', location: 'app.js:callAi33TTS:responseHeaders', message: 'AI33 auth headers', data: { host: (() => { try { return new URL(url).host; } catch (_) { return 'invalid-url'; } })(), path: (() => { try { return new URL(url).pathname; } catch (_) { return ''; } })(), wwwAuthenticate: (response.headers.get('www-authenticate') || '').slice(0, 200), server: (response.headers.get('server') || '').slice(0, 80) }, timestamp: Date.now() }) }).catch(() => {});
  // #endregion
  if (!response.ok) {
    const ct = (response.headers.get('content-type') || '').toLowerCase();
    let errMsg = `AI33 TTS error: ${response.status}`;
    if (ct.includes('application/json')) {
      const err = await response.json().catch(() => ({}));
      errMsg = err?.error?.message || err?.message || errMsg;
    } else {
      const t = await response.text().catch(() => '');
      if (t && t.length < 800) errMsg = `${errMsg}: ${t.trim().slice(0, 400)}`;
    }
    // #region agent log
    fetch('http://127.0.0.1:7906/ingest/260818ca-86e0-4d11-8830-b6af7bbca1a1', { method: 'POST', headers: { 'Content-Type': 'application/json', 'X-Debug-Session-Id': 'db5bb1' }, body: JSON.stringify({ sessionId: 'db5bb1', hypothesisId: 'H4', location: 'app.js:callAi33TTS:error', message: 'AI33 error body', data: { status: response.status, errMsg: String(errMsg).slice(0, 500) }, timestamp: Date.now() }) }).catch(() => {});
    // #endregion
    throw new Error(errMsg);
  }
  const ct = (response.headers.get('content-type') || '').toLowerCase();
  if (ct.includes('application/json')) {
    const json = await response.json().catch(() => ({}));
    const audioUrl = json?.audio_url || json?.url || json?.data?.audio_url || json?.data?.url || json?.file_url || '';
    const audioBase64 = json?.audio_base64 || json?.base64 || json?.data?.audio_base64 || '';
    // #region agent log
    fetch('http://127.0.0.1:7906/ingest/260818ca-86e0-4d11-8830-b6af7bbca1a1', { method: 'POST', headers: { 'Content-Type': 'application/json', 'X-Debug-Session-Id': 'db5bb1' }, body: JSON.stringify({ sessionId: 'db5bb1', hypothesisId: 'H8', location: 'app.js:callAi33TTS:json200', message: 'AI33 JSON success payload', data: { hasAudioUrl: !!audioUrl, hasAudioBase64: !!audioBase64, keys: Object.keys(json || {}).slice(0, 8), dataKeys: Object.keys(json?.data || {}).slice(0, 8) }, timestamp: Date.now() }) }).catch(() => {});
    // #endregion
    if (audioBase64) {
      const bin = Uint8Array.from(atob(audioBase64), c => c.charCodeAt(0));
      return new Blob([bin], { type: 'audio/mpeg' });
    }
    if (audioUrl) {
      const audioResp = await fetchWithTimeout(audioUrl, { method: 'GET' }, 120000);
      if (!audioResp.ok) throw new Error(`AI33 audio URL fetch failed: ${audioResp.status}`);
      const buf2 = await audioResp.arrayBuffer();
      const ct2 = (audioResp.headers.get('content-type') || '').toLowerCase();
      const mime2 = ct2.includes('audio/') ? ct2.split(';')[0].trim() : 'audio/mpeg';
      // #region agent log
      fetch('http://127.0.0.1:7906/ingest/260818ca-86e0-4d11-8830-b6af7bbca1a1', { method: 'POST', headers: { 'Content-Type': 'application/json', 'X-Debug-Session-Id': 'db5bb1' }, body: JSON.stringify({ sessionId: 'db5bb1', hypothesisId: 'H8', location: 'app.js:callAi33TTS:urlAudio', message: 'AI33 fetched audio URL', data: { byteLength: buf2.byteLength, mime: mime2 }, timestamp: Date.now() }) }).catch(() => {});
      // #endregion
      return new Blob([buf2], { type: mime2 || 'audio/mpeg' });
    }
    if (json?.task_id) {
      const taskId = String(json.task_id);
      // #region agent log
      fetch('http://127.0.0.1:7906/ingest/260818ca-86e0-4d11-8830-b6af7bbca1a1', { method: 'POST', headers: { 'Content-Type': 'application/json', 'X-Debug-Session-Id': 'db5bb1' }, body: JSON.stringify({ sessionId: 'db5bb1', hypothesisId: 'H10', location: 'app.js:callAi33TTS:taskAccepted', message: 'AI33 async task accepted', data: { taskIdLen: taskId.length, hasCreditsField: Object.prototype.hasOwnProperty.call(json || {}, 'ec_remain_credits') }, timestamp: Date.now() }) }).catch(() => {});
      // #endregion
      throw new Error(`AI33 accepted async task_id (${taskId}) and did not return direct audio. Configure webhook delivery (receive_url) or use a synchronous endpoint for in-app demo playback.`);
    }
    throw new Error(`AI33 returned JSON without audio payload: ${JSON.stringify(json).slice(0, 220)}`);
  }
  const buf = await response.arrayBuffer();
  const mime = ct.includes('audio/') ? ct.split(';')[0].trim() : 'audio/mpeg';
  // #region agent log
  fetch('http://127.0.0.1:7906/ingest/260818ca-86e0-4d11-8830-b6af7bbca1a1', { method: 'POST', headers: { 'Content-Type': 'application/json', 'X-Debug-Session-Id': 'db5bb1' }, body: JSON.stringify({ sessionId: 'db5bb1', hypothesisId: 'H3', location: 'app.js:callAi33TTS:success', message: 'AI33 audio body', data: { byteLength: buf.byteLength, mime }, timestamp: Date.now() }) }).catch(() => {});
  // #endregion
  return new Blob([buf], { type: mime || 'audio/mpeg' });
}


const TTS_VOICE_DEMO_PHRASES = {
  narratorVoice: 'This is the narrator voice sample.',
  femaleVoice: 'Hello, this is the female character voice sample.',
  maleVoice: 'Hello, this is the male character voice sample.',
};

function stopTtsVoiceDemoPlayback() {
  if (!state.ttsDemoPlayback) return;
  try {
    state.ttsDemoPlayback.audio.pause();
  } catch (_) {}
  state.ttsDemoPlayback.audio.removeAttribute('src');
  URL.revokeObjectURL(state.ttsDemoPlayback.url);
  state.ttsDemoPlayback = null;
}

/** Preview the selected voice in Settings using the same TTS provider as chapter audio. */
async function playTtsVoiceDemo(selectId) {
  const keys = getTtsApiKeys();
  // #region agent log
  fetch('http://127.0.0.1:7906/ingest/260818ca-86e0-4d11-8830-b6af7bbca1a1', { method: 'POST', headers: { 'Content-Type': 'application/json', 'X-Debug-Session-Id': 'db5bb1' }, body: JSON.stringify({ sessionId: 'db5bb1', hypothesisId: 'H1', location: 'app.js:playTtsVoiceDemo:entry', message: 'demo entry', data: { provider: getTtsProvider(), keysLen: keys.length, selectId }, timestamp: Date.now() }) }).catch(() => {});
  // #endregion
  if (!keys.length) {
    showToast(ttsKeyMissingMessage(), 'error');
    return;
  }
  const voiceName = document.getElementById(selectId)?.value;
  if (!voiceName) return;
  const phrase = TTS_VOICE_DEMO_PHRASES[selectId] || 'Voice preview.';
  const btn = document.getElementById(`demoVoice_${selectId}`);
  stopTtsVoiceDemoPlayback();
  if (btn) {
    btn.disabled = true;
    btn.textContent = '…';
  }
  try {
    // #region agent log
    fetch('http://127.0.0.1:7906/ingest/260818ca-86e0-4d11-8830-b6af7bbca1a1', { method: 'POST', headers: { 'Content-Type': 'application/json', 'X-Debug-Session-Id': 'db5bb1' }, body: JSON.stringify({ sessionId: 'db5bb1', hypothesisId: 'H5', location: 'app.js:playTtsVoiceDemo:preTts', message: 'calling TTS', data: { voiceName, phraseLen: phrase.length }, timestamp: Date.now() }) }).catch(() => {});
    // #endregion
    const blob = await callGeminiTTSMultiSpeaker(`NARRATOR: ${phrase}`, { NARRATOR: voiceName });
    // #region agent log
    fetch('http://127.0.0.1:7906/ingest/260818ca-86e0-4d11-8830-b6af7bbca1a1', { method: 'POST', headers: { 'Content-Type': 'application/json', 'X-Debug-Session-Id': 'db5bb1' }, body: JSON.stringify({ sessionId: 'db5bb1', hypothesisId: 'H3', location: 'app.js:playTtsVoiceDemo:blob', message: 'blob before play', data: { size: blob?.size, type: blob?.type || '' }, timestamp: Date.now() }) }).catch(() => {});
    // #endregion
    const url = URL.createObjectURL(blob);
    const audio = new Audio(url);
    state.ttsDemoPlayback = { audio, url };
    audio.onended = () => {
      stopTtsVoiceDemoPlayback();
    };
    audio.onerror = () => {
      stopTtsVoiceDemoPlayback();
      showToast('Could not play audio preview.', 'error');
    };
    await audio.play().catch((err) => {
      // #region agent log
      fetch('http://127.0.0.1:7906/ingest/260818ca-86e0-4d11-8830-b6af7bbca1a1', { method: 'POST', headers: { 'Content-Type': 'application/json', 'X-Debug-Session-Id': 'db5bb1' }, body: JSON.stringify({ sessionId: 'db5bb1', hypothesisId: 'H3', location: 'app.js:playTtsVoiceDemo:playCatch', message: 'audio.play rejected', data: { err: String(err?.message || err) }, timestamp: Date.now() }) }).catch(() => {});
      // #endregion
      stopTtsVoiceDemoPlayback();
      showToast(`Playback failed: ${err?.message || err}`, 'error');
    });
  } catch (e) {
    // #region agent log
    fetch('http://127.0.0.1:7906/ingest/260818ca-86e0-4d11-8830-b6af7bbca1a1', { method: 'POST', headers: { 'Content-Type': 'application/json', 'X-Debug-Session-Id': 'db5bb1' }, body: JSON.stringify({ sessionId: 'db5bb1', hypothesisId: 'H2', location: 'app.js:playTtsVoiceDemo:catch', message: 'demo catch', data: { err: String(e?.message || e) }, timestamp: Date.now() }) }).catch(() => {});
    // #endregion
    stopTtsVoiceDemoPlayback();
    showToast(`Demo failed: ${e.message || e}`, 'error');
  }
  if (btn) {
    btn.disabled = false;
    btn.textContent = '▶ Demo';
  }
}

// --- Call Gemini Multi-Speaker TTS (batch = fewer API calls) ---
async function callGeminiTTSMultiSpeaker(prompt, voiceMap, apiKeyOverride = null) {
  if (getTtsProvider() === 'ai33pro') {
    const speakers = Object.keys(voiceMap);
    const voice = voiceMap[speakers[0]] || 'alloy';
    const cleanPrompt = prompt.replace(/^[A-Z\s]+:\s*/gm, '').trim();
    // #region agent log
    fetch('http://127.0.0.1:7906/ingest/260818ca-86e0-4d11-8830-b6af7bbca1a1', { method: 'POST', headers: { 'Content-Type': 'application/json', 'X-Debug-Session-Id': 'db5bb1' }, body: JSON.stringify({ sessionId: 'db5bb1', hypothesisId: 'H5', location: 'app.js:callGeminiTTSMultiSpeaker:ai33', message: 'AI33 branch', data: { firstSpeaker: speakers[0], voice, cleanLen: cleanPrompt.length }, timestamp: Date.now() }) }).catch(() => {});
    // #endregion
    return callAi33TTS(cleanPrompt, voice, apiKeyOverride);
  }
  const apiKey = apiKeyOverride || getTtsApiKey();
  if (!apiKey) throw new Error('No TTS API key. Add a key in Settings (TTS API Key or Gemini key).');
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash-preview-tts:generateContent?key=${encodeURIComponent(apiKey)}`;
  const speakers = Object.keys(voiceMap);
  if (speakers.length === 0) throw new Error('No speakers');
  if (speakers.length === 1) {
    const voiceName = voiceMap[speakers[0]];
    const cleanPrompt = prompt.replace(/^[A-Z\s]+:\s*/gm, '').trim();
    const response = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
        contents: [{ parts: [{ text: cleanPrompt }] }],
        generationConfig: {
          responseModalities: ['AUDIO'],
          speechConfig: { voiceConfig: { prebuiltVoiceConfig: { voiceName } } },
        },
      }),
    });
    if (!response.ok) {
      const err = await response.json().catch(() => ({}));
      throw new Error(err?.error?.message || `TTS API error: ${response.status}`);
    }
    const data = await response.json();
    const blob = geminiTtsResponseToAudioBlob(data);
    if (!blob) throw new Error(geminiTtsErrorDetail(data));
    return blob;
  }
  const speakerVoiceConfigs = speakers.map(speaker => ({
    speaker,
    voiceConfig: { prebuiltVoiceConfig: { voiceName: voiceMap[speaker] } },
  }));
  const response = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      contents: [{ parts: [{ text: prompt }] }],
      generationConfig: {
        responseModalities: ['AUDIO'],
        speechConfig: {
          multiSpeakerVoiceConfig: { speakerVoiceConfigs },
        },
      },
    }),
  });
  if (!response.ok) {
    const err = await response.json().catch(() => ({}));
    throw new Error(err?.error?.message || `TTS API error: ${response.status}`);
  }
  const data = await response.json();
  const blob = geminiTtsResponseToAudioBlob(data);
  if (!blob) throw new Error(geminiTtsErrorDetail(data));
  return blob;
}

// --- Call Gemini TTS API (single segment) ---
async function callGeminiTTS(text, novel, segmentRaw = null) {
  const cleaned = stripTextForTTS(segmentRaw || text);
  if (!cleaned) throw new Error('No text to speak (or segment is cue-only)');
  const voiceName = getVoiceForSegment(segmentRaw || text, novel);
  if (getTtsProvider() === 'ai33pro') {
    return callAi33TTS(cleaned, voiceName);
  }
  const apiKey = getTtsApiKey();
  if (!apiKey) throw new Error('No TTS API key. Add a key in Settings.');
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash-preview-tts:generateContent?key=${encodeURIComponent(apiKey)}`;
  const tone = novel?.narratorTone || '';
  const background = novel?.background || '';
  const styleHint = [tone, background].filter(Boolean).join('. ');
  const prompt = styleHint
    ? `Say in this style: ${styleHint}\n\n"${cleaned.replace(/"/g, '\\"')}"`
    : cleaned;
  const response = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      contents: [{ parts: [{ text: prompt }] }],
      generationConfig: {
        responseModalities: ['AUDIO'],
        speechConfig: {
          voiceConfig: { prebuiltVoiceConfig: { voiceName } },
        },
      },
    }),
  });
  if (!response.ok) {
    const err = await response.json().catch(() => ({}));
    throw new Error(err?.error?.message || `TTS API error: ${response.status}`);
  }
  const data = await response.json();
  const blob = geminiTtsResponseToAudioBlob(data);
  if (!blob) throw new Error(geminiTtsErrorDetail(data));
  return blob;
}

// --- Gemini Imagen: thumbnail/cover generation (when AI provider is Gemini and key set) ---
// options: { aspectRatio: '3:4' } for novel thumbnail ratio; { apiKeyOverride } to use a specific key (e.g. when provider is DeepSeek)
async function callGeminiImagen(prompt, options = {}) {
  const apiKey = options.apiKeyOverride || getApiKey();
  if (!apiKey) return null;
  const models = [
    'gemini-2.5-flash-image-preview',
    'gemini-2.0-flash-exp-image-generation',
    'gemini-1.5-pro', // last resort: may not support IMAGE output
  ];
  const aspectRatio = options.aspectRatio || null;
  const generationConfig = {
    responseModalities: ['IMAGE'],
  };
  if (aspectRatio) {
    generationConfig.imageConfig = { aspectRatio };
  }
  let lastErr = null;
  for (const model of models) {
    try {
      const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${encodeURIComponent(apiKey)}`;
      const response = await fetchWithTimeout(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          contents: [{ parts: [{ text: String(prompt).slice(0, 4096) }] }],
          generationConfig,
        }),
      }, 120000);
      if (!response.ok) {
        const err = await response.json().catch(() => ({}));
        const msg = err?.error?.message || err?.message || `HTTP ${response.status}`;
        lastErr = msg;
        if (msg.includes('404') || msg.includes('not found') || msg.includes('Invalid model')) continue;
        throw new Error(msg);
      }
      const data = await response.json();
      const parts = data?.candidates?.[0]?.content?.parts || [];
      const inline = parts.find(p => p?.inlineData?.data) || parts.find(p => p?.inline_data?.data);
      const blob = inline?.inlineData || inline?.inline_data;
      const b64 = blob?.data;
      const mime = blob?.mimeType || blob?.mime_type || 'image/png';
      if (!b64) throw new Error('No image returned');
      return `data:${mime};base64,${b64}`;
    } catch (e) {
      lastErr = e?.message || String(e);
      continue;
    }
  }
  throw new Error(lastErr || 'Gemini image generation failed');
}

// --- Free image API fallback (no key required) ---
async function callFreeImageAPI(prompt, novel) {
  if (state.imageGenerationDisabled) {
    throw new Error(state.imageGenerationDisabledReason || 'Image generation is disabled');
  }
  const tone = novel?.narratorTone || '';
  const background = novel?.background || '';
  const draftBit = safeStr(novel?.draftScript).slice(0, 320);
  const styleHint = [tone, background, draftBit].filter(Boolean).join('. ');
  const fullPrompt = styleHint
    ? `Book cover scene inspired by story: ${styleHint}. ${prompt}. Digital art, high quality, show title if requested in prompt.`
    : `Scene image: ${prompt}. Digital art, high quality, atmospheric.`;
  const url = 'https://t2i.mcpcore.xyz/api/free/generate';
  let response;
  try {
    response = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ prompt: fullPrompt, model: 'turbo' }),
    });
  } catch (e) {
    // On GitHub Pages this commonly fails due to CORS. Only disable free API for non-Gemini users.
    const msg = (e && (e.message || String(e))) || 'Failed to fetch';
    if (getAIProvider() !== 'gemini') {
      state.imageGenerationDisabled = true;
      state.imageGenerationDisabledReason = msg;
    }
    throw e;
  }
  if (!response.ok) throw new Error(`Image API error: ${response.status}`);
  const reader = response.body.getReader();
  const decoder = new TextDecoder();
  let imageUrl = null;
  let errMsg = null;
  while (true) {
    const { done, value } = await reader.read();
    if (done) break;
    const chunk = decoder.decode(value);
    const lines = chunk.split('\n').filter(l => l.startsWith('data: '));
    for (const line of lines) {
      try {
        const data = JSON.parse(line.slice(6));
        if (data.status === 'complete' && data.imageUrl) imageUrl = data.imageUrl;
        if (data.status === 'error') errMsg = data.message || 'Image generation failed';
      } catch (_) {}
    }
  }
  if (errMsg) throw new Error(errMsg);
  if (!imageUrl) {
    if (getAIProvider() !== 'gemini') {
      state.imageGenerationDisabled = true;
      state.imageGenerationDisabledReason = 'No image URL in response';
    }
    throw new Error('No image URL in response');
  }
  const imgResp = await fetch(imageUrl);
  if (!imgResp.ok) throw new Error('Failed to fetch generated image');
  const blob = await imgResp.blob();
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(r.result);
    r.onerror = () => reject(new Error('Failed to read image'));
    r.readAsDataURL(blob);
  });
}

// --- Unified image generation: Gemini Imagen when a Gemini key is available (Gemini or DeepSeek + Gemini key); else free API ---
// genOptions: { aspectRatio: '3:4' } for novel thumbnails. DeepSeek has no image API; we use Gemini key from Settings (Gemini key for TTS) when provider is DeepSeek.
async function callImageGenerationAPI(prompt, novel, genOptions = {}) {
  const core = getNovelCoreIdeasForImage(novel, 700);
  const fullPrompt = `Book cover with readable title typography: "${novel?.title || 'Untitled'}". Artwork inspired by the story:\n${core}\nGenre: ${novel?.genre || novel?.category || 'Fiction'}. Mood: ${novel?.narratorTone || 'cinematic'}. Professional cover art, no watermarks.`;
  const effectivePrompt = prompt && prompt.length > 50 ? prompt : fullPrompt;
  const geminiKey = getGeminiKeyForImages();
  if (geminiKey) {
    try {
      return await callGeminiImagen(effectivePrompt, {
        aspectRatio: genOptions.aspectRatio || null,
        apiKeyOverride: geminiKey,
      });
    } catch (e) {
      console.warn('Gemini Imagen failed, falling back to free API:', e?.message);
    }
  }
  if (getAIProvider() === 'deepseek' && !geminiKey) {
    throw new Error('Image generation requires a Gemini API key. In Settings, add a "Gemini Key (for TTS when using DeepSeek)" to use Gemini for thumbnails while using DeepSeek for text.');
  }
  return callFreeImageAPI(effectivePrompt, novel);
}

// --- Generate Audio for a single segment ---
async function generateAudioForSegment(flowId, audioIndex, segmentIndex) {
  const flow = getFlow(flowId);
  if (!flow) return;
  const fd = flowDomId(flowId);
  const segments = flow.audioScriptSegments[audioIndex];
  const raw = document.querySelector(`#flowPanel_${fd} .script-segment-text[data-flow-id="${flowId}"][data-audio-index="${audioIndex}"][data-segment-index="${segmentIndex}"]`)?.textContent?.trim() || segments?.[segmentIndex];
  const text = stripTextForTTS(raw);
  if (!text) {
    showToast('Segment has no speakable content (cue-only).', 'error');
    return;
  }
  const novel = flow.novels[audioIndex];
  const btn = document.getElementById(`genAudioBtn_${fd}_${audioIndex}_${segmentIndex}`);
  if (!getTtsApiKeys().length) {
    showToast(ttsKeyMissingMessage(), 'error');
    return;
  }
  if (btn) { btn.disabled = true; btn.textContent = '⏳'; }
  try {
    const blob = await callGeminiTTS(text, novel, raw);
    const url = URL.createObjectURL(blob);
    if (!flow.generatedAudio[audioIndex]) flow.generatedAudio[audioIndex] = {};
    flow.generatedAudio[audioIndex][segmentIndex] = url;
    const container = document.getElementById(`audioScriptSegments_${fd}_${audioIndex}`);
    renderAudioScriptSegments(flowId, audioIndex, container, flow.audioScriptSegments[audioIndex]);
    showToast(`Audio generated for segment ${segmentIndex + 1}`, 'success');
  } catch (e) {
    showToast(`Audio failed: ${e.message}`, 'error');
  }
  if (btn) { btn.disabled = false; btn.textContent = '🎵'; }
}

// --- Generate All Audio (batched + parallel across multiple API keys) ---
async function generateAllAudio(flowId, audioIndex) {
  const flow = getFlow(flowId);
  if (!flow) return;
  const fd = flowDomId(flowId);
  const novel = flow.novels[audioIndex];
  const segments = flow.audioScriptSegments[audioIndex];
  if (!novel || !segments?.length) {
    showToast('No script segments to generate audio from.', 'error');
    return;
  }
  const keys = getTtsApiKeys();
  if (!keys.length) {
    showToast(ttsKeyMissingMessage(), 'error');
    return;
  }
  const btn = document.getElementById(`generateAllAudioBtn_${fd}_${audioIndex}`);
  btn.classList.add('loading');
  btn.disabled = true;
  const container = document.getElementById(`audioScriptSegments_${fd}_${audioIndex}`);
  const batches = buildAudioBatches(segments, novel);
  if (!flow.generatedAudioBatches[audioIndex]) flow.generatedAudioBatches[audioIndex] = {};

  const batchPerKey = Math.ceil(batches.length / keys.length) || 1;
  const worker = async (keyIndex) => {
    const key = keys[keyIndex];
    const start = keyIndex * batchPerKey;
    const end = Math.min(start + batchPerKey, batches.length);
    let done = 0;
    for (let b = start; b < end; b++) {
      const batch = batches[b];
      if (b > start) await new Promise(r => setTimeout(r, TTS_THROTTLE_MS));
      try {
        const blob = await callGeminiTTSMultiSpeaker(batch.prompt, batch.voiceMap, key);
        const url = URL.createObjectURL(blob);
        const keyName = `batch_${batch.indices[0]}_${batch.indices[batch.indices.length - 1]}`;
        if (!flow.generatedAudio[audioIndex]) flow.generatedAudio[audioIndex] = {};
        flow.generatedAudio[audioIndex][keyName] = url;
        flow.generatedAudioBatches[audioIndex][keyName] = batch.indices;
        done += batch.indices.length;
        renderAudioScriptSegments(flowId, audioIndex, container, flow.audioScriptSegments[audioIndex]);
        showToast(`Key ${keyIndex + 1}: batch ${b + 1}/${batches.length}`, 'info');
      } catch (e) {
        showToast(`Batch ${b + 1} failed: ${e.message}`, 'error');
      }
    }
    return done;
  };

  const workerCount = Math.min(keys.length, batches.length);
  const results = await Promise.all(
    Array.from({ length: workerCount }, (_, i) => worker(i))
  );
  const totalDone = results.reduce((a, b) => a + b, 0);
  btn.classList.remove('loading');
  btn.disabled = false;
  showToast(`Generated ${totalDone} segments using ${workerCount} key(s) in parallel.`, 'success');
}

// --- Generate Scene for a single segment ---
async function generateSceneForSegment(flowId, audioIndex, segmentIndex) {
  const flow = getFlow(flowId);
  if (!flow) return;
  const fd = flowDomId(flowId);
  const segments = flow.audioScriptSegments[audioIndex];
  const text = document.querySelector(`#flowPanel_${fd} .script-segment-text[data-flow-id="${flowId}"][data-audio-index="${audioIndex}"][data-segment-index="${segmentIndex}"]`)?.textContent?.trim() || segments?.[segmentIndex];
  if (!text) return;
  const novel = flow.novels[audioIndex];
  const btn = document.getElementById(`genSceneBtn_${fd}_${audioIndex}_${segmentIndex}`);
  if (btn) { btn.disabled = true; btn.textContent = '⏳'; }
  try {
    const sceneDesc = text.replace(/^\[(SCENE|INT\.|EXT\.|SFX|AMB|MUSIC)[^\]]*\]\s*/i, '').trim() || text;
    const dataUrl = await callImageGenerationAPI(sceneDesc, novel);
    if (!flow.generatedScenes[audioIndex]) flow.generatedScenes[audioIndex] = {};
    flow.generatedScenes[audioIndex][segmentIndex] = dataUrl;
    const container = document.getElementById(`audioScriptSegments_${fd}_${audioIndex}`);
    renderAudioScriptSegments(flowId, audioIndex, container, flow.audioScriptSegments[audioIndex]);
    showToast(`Scene generated for segment ${segmentIndex + 1}`, 'success');
  } catch (e) {
    showToast(`Scene failed: ${e.message}`, 'error');
  }
  if (btn) { btn.disabled = false; btn.textContent = '🖼️'; }
}

// --- Generate All Scenes (for scene-type segments only) ---
async function generateAllScenes(flowId, audioIndex) {
  const flow = getFlow(flowId);
  if (!flow) return;
  const fd = flowDomId(flowId);
  const novel = flow.novels[audioIndex];
  const segments = flow.audioScriptSegments[audioIndex];
  if (!novel || !segments?.length) return;
  const sceneIndices = segments
    .map((t, i) => (/^\[(SCENE|INT\.|EXT\.)[^\]]*\]/i.test(t) ? i : -1))
    .filter(i => i >= 0);
  if (!sceneIndices.length) {
    showToast('No scene cues ([SCENE:...] or [INT./EXT.]) found in script.', 'error');
    return;
  }
  const apiKey = getApiKey();
  if (!apiKey) { showToast('Please enter your Gemini API key.', 'error'); return; }
  const btn = document.getElementById(`generateAllScenesBtn_${fd}_${audioIndex}`);
  btn.classList.add('loading');
  btn.disabled = true;
  const container = document.getElementById(`audioScriptSegments_${fd}_${audioIndex}`);
  let done = 0;
  for (const i of sceneIndices) {
    const text = document.querySelector(`#flowPanel_${fd} .script-segment-text[data-flow-id="${flowId}"][data-audio-index="${audioIndex}"][data-segment-index="${i}"]`)?.textContent?.trim() || segments[i];
    const sceneDesc = text.replace(/^\[(SCENE|INT\.|EXT\.)[^\]]*\]\s*/i, '').trim() || text;
    try {
      const dataUrl = await callImageGenerationAPI(sceneDesc, novel);
      if (!flow.generatedScenes[audioIndex]) flow.generatedScenes[audioIndex] = {};
      flow.generatedScenes[audioIndex][i] = dataUrl;
      renderAudioScriptSegments(flowId, audioIndex, container, flow.audioScriptSegments[audioIndex]);
      done++;
      showToast(`Scene ${done}/${sceneIndices.length}`, 'info');
    } catch (e) {
      showToast(`Scene ${i + 1} failed: ${e.message}`, 'error');
    }
  }
  btn.classList.remove('loading');
  btn.disabled = false;
  showToast(`Generated ${done}/${sceneIndices.length} scene images.`, 'success');
}

// --- Download Audio Drama Script ---
function downloadAudioScript(flowId, index) {
  const flow = getFlow(flowId);
  if (!flow) return;
  const fd = flowDomId(flowId);
  const novel = flow.novels[index];
  const textEls = document.querySelectorAll(`#flowPanel_${fd} .script-segment-text[data-audio-index="${index}"]`);
  if (textEls.length) {
    const segs = [...(flow.audioScriptSegments[index] || [])];
    textEls.forEach(el => {
      const sIdx = parseInt(el.dataset.segmentIndex, 10);
      if (!isNaN(sIdx)) segs[sIdx] = el.textContent || '';
    });
    flow.audioScriptSegments[index] = segs;
  }
  const segments = flow.audioScriptSegments[index];
  const scriptText = segments ? segments.join('\n') : flow.audioScripts[index];
  if (!novel || !scriptText) {
    showToast('No script to download. Generate it first.', 'error');
    return;
  }

  let content = '';
  content += `${'='.repeat(60)}\n`;
  content += `  AUDIO DRAMA SCRIPT: ${novel.title || 'Untitled'}\n`;
  content += `  by ${novel.authorName || 'Unknown Author'}\n`;
  content += `${'='.repeat(60)}\n\n`;
  content += scriptText;
  content += `\n\n${'='.repeat(60)}\n`;
  content += `  Generated by AI Novel Template Generator\n`;
  content += `  Date: ${new Date().toLocaleDateString()}\n`;
  content += `${'='.repeat(60)}\n`;

  const blob = new Blob([content], { type: 'text/plain;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `audio_script_${index + 1}_${sanitizeFilename(novel.title)}.txt`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
  showToast(`Downloaded script: ${novel.title}`, 'success');
}
