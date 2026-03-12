/* ============================================
   AI Novel Template Generator — Core Logic
   Gemini API Integration & Novel Generation
   ============================================ */

// --- State ---
const state = {
  apiKey: '',
  novels: [],
  stories: {},
  audioScripts: {}, // raw text for download
  audioScriptSegments: {}, // { [index]: string[] } - editable segments for each novel
  generatedAudio: {}, // { [novelIndex]: { [segmentIndex]: blobUrl } | { batch_0_4: blobUrl } }
  generatedAudioBatches: {}, // { [novelIndex]: { batchKey: [indices] } } for batch playback
  generatedScenes: {}, // { [novelIndex]: { [segmentIndex]: imageDataUrl } }
  reviewedNovels: new Set(), // indices of templates marked "passed manual review"
  isGenerating: false,
  speakingSegment: null, // { audioIndex, segmentIndex } - for stopping TTS
  imageGenerationDisabled: false,
  imageGenerationDisabledReason: '',
};

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

function addHistoryRun(formData, novels) {
  if (!Array.isArray(novels) || !novels.length) return;
  const now = new Date();
  const run = {
    id: `${now.getTime()}_${Math.random().toString(16).slice(2)}`,
    createdAt: now.toISOString(),
    provider: getAIProvider() || 'gemini',
    numNovels: novels.length,
    masterPrompt: safeStr(formData?.masterPrompt).slice(0, 220),
    novels,
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
      state.novels = run.novels;
      normalizeNovelsForExport(state.novels);
      stampCollectionAndCategoriesFromForm(state.novels);
      renderResults(state.novels);
      showToast('Loaded history run into results.', 'success');
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

  updateApiStatusBadge();

  // Event listeners
  const genBtn = document.getElementById('generateBtn');
  if (genBtn) genBtn.addEventListener('click', handleGenerate);
  document.getElementById('loadExampleBtn')?.addEventListener('click', loadExampleTemplates);
  document.getElementById('testApiBtn')?.addEventListener('click', handleTestApi);
  document.getElementById('generateAllStoriesBtn')?.addEventListener('click', handleGenerateAllStories);
  document.getElementById('fillMissingDataBtn')?.addEventListener('click', handleFillMissingData);

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
      synopsis: 'In a world where emotions manifest as physical powers, a young woman discovers she can absorb others\' pain—and their memories. When the Empathic Order recruits her to hunt a rogue empath, she must choose between duty and the truth behind the war that shattered her family.',
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
        { chapterNumber: 1, title: 'The Awakening', summary: 'Elara\'s power surfaces during an attack.' },
        { chapterNumber: 2, title: 'The Order', summary: 'Recruited; first mission to find Kael.' },
        { chapterNumber: 3, title: 'The Truth', summary: 'Discovers the Order\'s lie about the war.' }
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
      synopsis: 'A ghostwriter for a reclusive celebrity novelist uncovers a real murder tied to the author\'s past. To finish the book and stay alive, she must piece together the story from coded manuscripts and dangerous interviews.',
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
        { chapterNumber: 1, title: 'The Contract', summary: 'Maya takes the ghostwriting job.' },
        { chapterNumber: 2, title: 'The Manuscript', summary: 'Coded pages hint at a real crime.' },
        { chapterNumber: 3, title: 'The Murder', summary: 'A body turns up; the book and reality collide.' }
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
      novel.cateogories = cateogories;
      if (!novel.genre) novel.genre = cateogories;
    }
    if (formAuthor) novel.authorName = formAuthor;
    else if (!safeStr(novel.authorName)) novel.authorName = 'Unknown Author';
    if (!safeStr(novel.premium)) novel.premium = 'yes';
    if (!safeStr(novel.show)) novel.show = 'yes';
  });
}

/** Ensure every novel has data needed for export: synopsis, at least one chapter, premium, show, author, etc. */
function normalizeNovelsForExport(novels) {
  if (!Array.isArray(novels)) return;
  novels.forEach(novel => {
    if (!novel || typeof novel !== 'object') return;
    const synopsis = safeStr(novel.synopsis);
    if (!synopsis || synopsis === 'N/A') {
      novel.synopsis = safeStr(novel.draftScript) || (novel.title ? `A story: ${novel.title}.` : 'No synopsis yet.');
    }
    if (!novel.chapters || !Array.isArray(novel.chapters) || novel.chapters.length === 0) {
      novel.chapters = [
        { chapterNumber: 1, title: 'Chapter 1', summary: novel.synopsis ? novel.synopsis.substring(0, 200) + (novel.synopsis.length > 200 ? '…' : '') : 'Opening.' }
      ];
    }
    if (!safeStr(novel.authorName)) novel.authorName = safeStr(document.getElementById('authorName')?.value) || 'Unknown Author';
    if (!safeStr(novel.premium)) novel.premium = 'yes';
    if (!safeStr(novel.show)) novel.show = 'yes';
    if (!safeStr(novel.collection)) novel.collection = safeStr(document.getElementById('collectionName')?.value);
    if (!getCategoriesForExport(novel)) novel.cateogories = safeStr(document.getElementById('categoryName')?.value) || novel.genre || 'Fiction';
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
  state.novels = getExampleTemplates();
  normalizeNovelsForExport(state.novels);
  stampCollectionAndCategoriesFromForm(state.novels);
  // Pre-fill sample full stories so export shows chapter_outline and full_story columns with content
  state.stories = {};
  state.novels.forEach((_, index) => {
    const sample = getExampleFullStory(index);
    if (sample) state.stories[index] = sample;
  });
  const section = document.getElementById('resultsSection');
  const countEl = document.getElementById('resultsCount');
  const container = document.getElementById('novelsContainer');
  if (!section || !container) return;
  section.classList.add('active');
  if (countEl) countEl.textContent = `${state.novels.length} example templates`;
  container.innerHTML = '';
  state.novels.forEach((novel, index) => {
    const card = createNovelCard(novel, index);
    card.style.animationDelay = `${index * 0.1}s`;
    container.appendChild(card);
  });
  attachEditSyncListeners(container);
  // Show full story in cards so the new columns are visible in UI and in export
  state.novels.forEach((_, index) => {
    const storyText = state.stories[index];
    if (storyText) {
      const storySection = document.getElementById(`storySection_${index}`);
      const storyContent = document.getElementById(`storyContent_${index}`);
      if (storySection && storyContent) {
        renderStoryChapters(index, storyText);
        storySection.style.display = 'block';
      }
      const storyBtn = document.getElementById(`storyBtn_${index}`);
      if (storyBtn) storyBtn.innerHTML = '<span class="btn-text">✅ Story Generated</span>';
    }
  });
  const firstCard = container?.querySelector('.novel-card');
  if (firstCard) firstCard.classList.add('expanded');
  section.scrollIntoView({ behavior: 'smooth', block: 'start' });
  if (canGenerateImages()) {
    showToast('Example templates loaded. Generating thumbnails…', 'success');
    generateCoversForAllTemplates();
  } else {
    showToast('Example templates loaded. Configure image generation to create thumbnails.', 'success');
  }
}

// --- Small helpers ---
function safeStr(v) {
  if (v == null) return '';
  if (Array.isArray(v)) return v.map(safeStr).filter(Boolean).join(', ');
  if (typeof v === 'object') return '';
  return String(v).trim();
}

function clampText(s, maxChars) {
  const str = safeStr(s);
  if (!maxChars || maxChars < 1) return str;
  if (str.length <= maxChars) return str;
  return str.slice(0, Math.max(0, maxChars - 1)).trimEnd() + '…';
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

function normalizeTagList(list) {
  const out = (Array.isArray(list) ? list : [])
    .map(t => limitWords(t, 2))
    .map(t => t.replace(/[|—–-]+/g, ' ').replace(/\s+/g, ' ').trim())
    .filter(Boolean);
  const seen = new Set();
  return out.filter(t => {
    const k = t.toLowerCase();
    if (seen.has(k)) return false;
    seen.add(k);
    return true;
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
  if (Array.isArray(novel?.themes) && novel.themes.length) return normalizeTagList(novel.themes).join(', ');
  if (Array.isArray(novel?.tags) && novel.tags.length) return normalizeTagList(novel.tags).join(', ');
  const single = safeStr(novel?.tag);
  return single ? normalizeTagList([single]).join(', ') : '';
}

/** Chapter outline as plain text (for export columns): "Ch1: Title — Summary\nCh2: ..." */
function getChapterOutlineText(novel) {
  if (!novel?.chapters || !novel.chapters.length) return '';
  return novel.chapters.map(ch =>
    `Ch${ch.chapterNumber}: ${safeStr(ch.title)} — ${safeStr(ch.summary)}`
  ).join('\n');
}

/** Full story text for a novel (from DOM or state), for export and template .txt */
function getFullStoryText(index) {
  const contentEl = document.getElementById(`storyContent_${index}`);
  const fromDom = contentEl?.textContent?.trim();
  if (fromDom) return fromDom;
  return safeStr(state.stories[index]);
}

/** Parsed chapters with full content for export. Uses story with [CHAPTER N] markers when available for accurate split. */
function getFullStoryChaptersForExport(index) {
  const rawFromState = safeStr(state.stories[index]);
  const fromDom = document.getElementById(`storyContent_${index}`)?.textContent?.trim();
  const textToParse = rawFromState || fromDom || '';
  let chapters = parseChaptersFromMarkers(textToParse);
  if (!chapters.length && textToParse) {
    chapters = [{ number: 1, title: '', content: textToParse }];
  }
  return chapters;
}

/** Export column order (reformatted template: one row per chapter) */
const EXPORT_HEADERS = ['thumbnail', 'title', 'description', 'premium', 'show', 'categories', 'collection', 'author', 'tags', 'chapter_outline', 'chapter_content'];

/** Build one row per chapter for a novel. Each row has same novel metadata + one chapter_outline and chapter_content. */
function buildExportRowsForNovel(novelIndex, novel, collection, thumbPath, coverPath) {
  const templateChapters = novel?.chapters || [];
  const storyChapters = getFullStoryChaptersForExport(novelIndex);
  const outlineByNum = {};
  templateChapters.forEach(ch => {
    outlineByNum[ch.chapterNumber] = `Ch${ch.chapterNumber}: ${safeStr(ch.title)} — ${safeStr(ch.summary)}`;
  });
  const contentByNum = {};
  storyChapters.forEach(ch => {
    contentByNum[ch.number] = safeStr(ch.content);
    if (!outlineByNum[ch.number])
      outlineByNum[ch.number] = ch.title ? `Ch${ch.number}: ${ch.title}` : `Chapter ${ch.number}`;
  });
  const allNums = [...new Set([...Object.keys(outlineByNum).map(Number), ...Object.keys(contentByNum).map(Number)])].sort((a, b) => a - b);
  const rows = [];
  const description = safeStr(novel.synopsis);
  const premium = safeStr(novel.premium);
  const show = safeStr(novel.show);
  const categories = getCategoriesForExport(novel);
  const author = safeStr(novel.authorName || novel.author) || safeStr(document.getElementById('authorName')?.value) || 'Unknown Author';
  const tags = getTagsForExport(novel);
  const coll = safeStr(novel.collection) || collection;
  const title = safeStr(novel.title);
  const hasThumb = !!pickThumbnailDataUrl(novelIndex, novel);
  const thumbCell = hasThumb ? thumbPath : (safeStr(novel.thumbnailPrompt) || buildThumbnailPromptFromNovel(novel));

  if (allNums.length === 0) {
    rows.push([thumbCell, title, description, premium || 'yes', show || 'yes', categories, coll, author, tags, '', '']);
    return rows;
  }
  allNums.forEach(num => {
    rows.push([
      thumbCell,
      title,
      description,
      premium || 'yes',
      show || 'yes',
      categories,
      coll,
      author,
      tags,
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

function pickCoverDataUrl(novelIndex, novel) {
  const maybe = novel?.cover || novel?.coverImage || novel?.image || novel?.thumbnail;
  if (typeof maybe === 'string' && maybe.startsWith('data:image/')) return maybe;
  const scenes = state.generatedScenes?.[novelIndex];
  if (!scenes) return '';
  const keys = Object.keys(scenes)
    .map(k => parseInt(k, 10))
    .filter(n => !Number.isNaN(n))
    .sort((a, b) => a - b);
  if (!keys.length) return '';
  const first = scenes[keys[0]];
  return (typeof first === 'string' && first.startsWith('data:image/')) ? first : '';
}

/** Before export: ensure every novel has cover + thumbnail so thumbnail/cover columns are filled. */
async function ensureThumbnailsForExport() {
  const novels = state.novels || [];
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
  showToast(`Generating ${missing.length} missing thumbnail(s)…`, 'info');
  for (const i of missing) {
    try {
      await generateCoverForNovel(i);
    } catch (e) {
      console.warn('Thumbnail generation failed for novel', i, e);
      showToast(`Thumbnail generation failed for novel ${i + 1}: ${e?.message || 'Unknown error'}`, 'error');
    }
    await new Promise(r => setTimeout(r, 200));
  }
  showToast('Thumbnails ready for export', 'success');
}

// --- Export: CSV / XLSX ---
async function handleExportCsv() {
  try {
    if (!Array.isArray(state.novels) || !state.novels.length) {
      showToast('Nothing to export yet. Generate novel templates first.', 'error');
      return;
    }
    normalizeNovelsForExport(state.novels);
    await ensureThumbnailsForExport();
    const collection = getExportCollection();
    const lines = [EXPORT_HEADERS.map(csvEscape).join(',')];
    for (let i = 0; i < state.novels.length; i++) {
      const novel = state.novels[i] || {};
      // CSV can’t embed images. Use relative file paths so it works with the .zip package export.
      const thumbPath = `thumbnails/novel_${i + 1}.png`;
      const coverPath = `covers/novel_${i + 1}.png`;
      const rows = buildExportRowsForNovel(i, novel, collection, thumbPath, coverPath);
      rows.forEach(row => lines.push(row.map(csvEscape).join(',')));
    }
    const csv = lines.join('\r\n');
    downloadBlob(new Blob([csv], { type: 'text/csv;charset=utf-8' }), 'novels_export.csv');
    showToast('Exported CSV (use Export package .zip for image files)', 'success');
  } catch (e) {
    console.error('CSV export failed', e);
    showToast('CSV export failed: ' + (e?.message || String(e)), 'error');
  }
}

async function handleExportZipPackage() {
  try {
    if (!Array.isArray(state.novels) || !state.novels.length) {
      showToast('Nothing to export yet. Generate novel templates first.', 'error');
      return;
    }
    normalizeNovelsForExport(state.novels);
    if (!window.JSZip) {
      showToast('ZIP export library failed to load. Reload and try again.', 'error');
      return;
    }
    await ensureThumbnailsForExport();

    // Ensure we have thumbnails for each cover.
    for (let i = 0; i < state.novels.length; i++) {
      const novel = state.novels[i] || {};
      if (!novel.thumbnail && novel.cover) {
        novel.thumbnail = await resizeDataUrl(novel.cover, 128, 128, 'image/png');
      }
    }

    const zip = new JSZip();
    const thumbs = zip.folder('thumbnails');
    const covers = zip.folder('covers');

    // CSV that references the packaged file paths (one row per chapter, same format as Export CSV)
    const collection = getExportCollection();
    const csvLines = [EXPORT_HEADERS.map(csvEscape).join(',')];

    const galleryCards = [];

    for (let i = 0; i < state.novels.length; i++) {
      const novel = state.novels[i] || {};
      const coverDataUrl = pickCoverDataUrl(i, novel);
      const thumbDataUrl = novel.thumbnail || (coverDataUrl ? await resizeDataUrl(coverDataUrl, 128, 128, 'image/png') : '');

      const coverPng = coverDataUrl ? await fetch(coverDataUrl).then(r => r.blob()) : null;
      const thumbPng = thumbDataUrl ? await fetch(thumbDataUrl).then(r => r.blob()) : null;

      const coverName = `novel_${i + 1}.png`;
      const thumbName = `novel_${i + 1}.png`;
      if (coverPng) covers.file(coverName, coverPng);
      if (thumbPng) thumbs.file(thumbName, thumbPng);

      const thumbPath = `thumbnails/${thumbName}`;
      const coverPath = `covers/${coverName}`;
      const rows = buildExportRowsForNovel(i, novel, collection, thumbPath, coverPath);
      rows.forEach(row => csvLines.push(row.map(csvEscape).join(',')));

      galleryCards.push(`
        <a class="card" href="${coverPath}" target="_blank" rel="noopener">
          <img src="${thumbPath}" alt="${escapeHtml(safeStr(novel.title) || ('Novel ' + (i + 1)))}"/>
          <div class="t">${escapeHtml(safeStr(novel.title) || ('Novel ' + (i + 1)))}</div>
        </a>
      `);
    }

    zip.file('templates.csv', csvLines.join('\r\n'));
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

async function handleExportXlsx() {
  try {
    if (!Array.isArray(state.novels) || !state.novels.length) {
      showToast('Nothing to export yet. Generate novel templates first.', 'error');
      return;
    }
    normalizeNovelsForExport(state.novels);
    if (!window.ExcelJS) {
      showToast('XLSX export library failed to load. Reload and try again.', 'error');
      return;
    }
    await ensureThumbnailsForExport();
    const collection = getExportCollection();
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Novels');
    ws.columns = EXPORT_HEADERS.map((h, idx) => ({
      header: h,
      key: h,
      width: h === 'chapter_content' || h === 'description' ? 60 : h === 'chapter_outline' ? 48 : 20,
    }));
    ws.getRow(1).font = { bold: true };
    ws.getRow(1).alignment = { vertical: 'middle' };
    ws.getRow(1).height = 20;

    let rowNumber = 2;
    for (let i = 0; i < state.novels.length; i++) {
      const novel = state.novels[i] || {};
      const thumbPath = `thumbnails/novel_${i + 1}.png`;
      const coverPath = `covers/novel_${i + 1}.png`;
      const rows = buildExportRowsForNovel(i, novel, collection, thumbPath, coverPath);
      rows.forEach(row => {
        ws.addRow({
          thumbnail: row[0],
          title: row[1],
          description: row[2],
          premium: row[3],
          show: row[4],
          categories: row[5],
          collection: row[6],
          author: row[7],
          tags: row[8],
          chapter_outline: row[9],
          chapter_content: row[10],
        });
        ws.getRow(rowNumber).alignment = { vertical: 'top', wrapText: true };
        ws.getRow(rowNumber).height = 24;
        rowNumber++;
      });
    }

    ws.views = [{ state: 'frozen', ySplit: 1 }];
    const buf = await wb.xlsx.writeBuffer();
    downloadBlob(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), 'novels_export.xlsx');
    showToast('Exported XLSX', 'success');
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
    authorName: document.getElementById('authorName').value.trim() || 'Unknown Author',
    collectionName: document.getElementById('collectionName')?.value || '',
    cateogoriesName: document.getElementById('categoryName')?.value || '',
    releaseDate: document.getElementById('releaseDate').value || new Date().toISOString().split('T')[0],
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
    prompt += `\n**Collection (Tên Danh Mục):** ${formData.collectionName}`;
  }
  if (formData.cateogoriesName) {
    prompt += `\n**Cateogories (Tên Thể Loại):** ${formData.cateogoriesName}`;
  }

  prompt += `

## OUTPUT REQUIREMENTS
Generate exactly ${formData.numNovels} novel templates. Each novel template MUST include ALL of the following fields:

1. **title** — A compelling, unique title for the novel
2. **synopsis** — A 3-5 paragraph synopsis of the full story
3. **draftScript** — The core scenario, script outline, and key ideas the story conveys
4. **characters** — An array of characters, each with: name, role (protagonist/antagonist/supporting), age, description, arc (character development summary), gender ("male" or "female" for voice casting)
5. **authorName** — "${formData.authorName}"
6. **releaseDate** — "${formData.releaseDate}"
7. **narratorTone** — Description of the narrative voice, POV, and tone
8. **background** — The world/setting description and backdrop of the story
9. **writingLanguage** — "${formData.writingLanguage}"
10. **chapters** — An array of 5-10 chapter outlines, each with: chapterNumber, title, summary (SHORT). CRITICAL: For each chapter, the combined string "${'title'} — ${'summary'}" must be <= 100 characters.
11. **themes** — Array of core themes explored in the novel
12. **genre** — Primary and secondary genres
13. **category** — Reader/info category (e.g. "Young Adult", "Adult Fiction", "Children's", "Non-fiction", "Romance", "Fantasy")
14. **collection** — "${formData.collectionName || ''}"
15. **cateogories** — "${formData.cateogoriesName || ''}"
16. **thumbnailPrompt** — A ready-to-paste prompt for Gemini Image generation to create a square 1:1 thumbnail image (NO TEXT). Must be a single string.
17. **premium** — "yes" or "no" (whether the novel is premium content)
18. **show** — "yes" or "no" (whether to show the novel in listings)

Each novel should be DISTINCT — different plot, different character dynamics, different themes — while still being inspired by the creative brief.
CRITICAL: Every novel MUST have a non-empty synopsis (3-5 sentences minimum) and MUST have 5-10 chapters with chapterNumber, title, and summary for each.

## OUTPUT FORMAT (CRITICAL)
You MUST return valid JSON only. No markdown code blocks, no backticks, no explanation—just the raw JSON object.
{
  "novels": [
    {
      "title": "...",
      "synopsis": "...",
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

function buildThumbnailPromptFromNovel(novel) {
  const title = safeStr(novel?.title) || 'Untitled novel';
  const genre = safeStr(novel?.genre) || safeStr(novel?.cateogories) || safeStr(novel?.category) || 'Fiction';
  const setting = safeStr(novel?.background) || 'Not specified';
  const mood = safeStr(novel?.narratorTone) || 'Cinematic';
  const synopsis = safeStr(novel?.synopsis);
  const motifs = synopsis
    ? synopsis.split(/[.?!]/).map(s => s.trim()).filter(Boolean).slice(0, 2).join(' / ')
    : '';

  const lines = [
    'Create a SQUARE 1:1 novel thumbnail image (NO TEXT, no typography, no watermark, no logo).',
    `Title concept: "${title}"`,
    `Genre: ${genre}`,
    `Setting: ${setting}`,
    `Mood/tone: ${mood}`,
    motifs ? `Key elements: ${motifs}` : 'Key elements: 1–3 strong visual motifs from the story.',
    '',
    'Style: cinematic, professional cover-art quality, high contrast, clear focal point, simple readable silhouette at small size.',
    'Composition: centered subject, minimal clutter, dramatic lighting, sharp focus.',
    'Constraints: avoid readable text, avoid extra limbs/fingers, avoid blurry faces.',
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

// --- Generate Handler ---
async function handleGenerate() {
  if (!validateForm()) return;
  if (state.isGenerating) return;

  state.isGenerating = true;
  const btn = document.getElementById('generateBtn');
  btn.classList.add('loading');
  btn.disabled = true;

  showProgress(true);
  updateProgress(10, 'Preparing creative brief...');

  try {
    const formData = collectFormData();
    updateProgress(25, 'Building AI prompt...');

    const prompt = buildGeminiPrompt(formData);
    const provider = getAIProvider() === 'deepseek' ? 'DeepSeek' : 'Gemini';
    updateProgress(40, `Generating ${formData.numNovels} novel templates with ${provider}...`);

    const result = await callGeminiAPI(prompt);
    updateProgress(85, 'Processing results...');

    if (!result.novels || !Array.isArray(result.novels)) {
      throw new Error('Invalid response structure from AI');
    }

    state.novels = result.novels;
    normalizeNovelsForExport(state.novels);
    stampCollectionAndCategoriesFromForm(state.novels);
    addHistoryRun(formData, state.novels);
    updateProgress(95, 'Generating thumbnails...');
    setTimeout(() => {
      showProgress(false);
      renderResults(state.novels);
      showToast(`Generated ${state.novels.length} templates. Generating thumbnails...`, 'success');
      // Generate cover + thumbnail for each template (for review + export).
      generateCoversForAllTemplates();
    }, 300);

  } catch (error) {
    console.error('Generation error:', error);
    showProgress(false);
    let msg = error?.message || String(error);
    if (msg === 'Failed to fetch' || msg.includes('NetworkError')) {
      msg = 'Network error. If using file://, run: npx serve -l 3000 and open http://localhost:3000';
    }
    showToast('Generation failed: ' + msg, 'error');
  } finally {
    state.isGenerating = false;
    btn.classList.remove('loading');
    btn.disabled = false;
  }
}

// --- Auto-generate cover + thumbnail for templates (for review UI + export) ---
function ensureCoverThumbInCard(index) {
  const novel = state.novels?.[index];
  if (!novel) return;
  const dataUrl =
    (typeof novel.thumbnail === 'string' && novel.thumbnail.startsWith('data:image/')) ? novel.thumbnail
      : (typeof novel.cover === 'string' && novel.cover.startsWith('data:image/')) ? novel.cover : '';
  if (!dataUrl) return;
  const card = document.querySelector(`.novel-card[data-index="${index}"]`);
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

async function generateCoverForNovel(index) {
  const novel = state.novels?.[index];
  if (!novel) return;
  if (typeof novel.cover === 'string' && novel.cover.startsWith('data:image/')) {
    ensureCoverThumbInCard(index);
    return;
  }
  if (typeof callImageGenerationAPI !== 'function') return;

  const prompt = `Book cover illustration for a novel. No text, no typography, no watermark.
Title concept: "${novel.title || 'Untitled'}".
Genre: ${novel.genre || novel.category || 'Fiction'}.
Setting/background: ${novel.background || 'not specified'}.
Main mood/tone: ${novel.narratorTone || ''}.
Composition: centered subject, cinematic lighting, high detail, professional cover art.`;

  const cover = await callImageGenerationAPI(prompt, novel);
  novel.cover = cover;
  novel.thumbnail = await resizeDataUrl(cover, 128, 128, 'image/png');
  ensureCoverThumbInCard(index);
}

async function generateCoversForAllTemplates() {
  const novels = state.novels || [];
  if (!Array.isArray(novels) || !novels.length) return;

  const indices = novels.map((_, i) => i).filter(i => !(typeof novels[i]?.cover === 'string' && novels[i].cover.startsWith('data:image/')));
  if (!indices.length) {
    indices.forEach(i => ensureCoverThumbInCard(i));
    return;
  }

  const concurrency = Math.min(3, indices.length);
  const queue = indices.slice();
  let done = 0;
  showToast(`Generating ${indices.length} thumbnails for templates...`, 'info');

  const worker = async () => {
    while (queue.length) {
      const i = queue.shift();
      if (i == null) break;
      try {
        await generateCoverForNovel(i);
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

// --- Render Results ---
function renderResults(novels) {
  const section = document.getElementById('resultsSection');
  const container = document.getElementById('novelsContainer');
  const countEl = document.getElementById('resultsCount');

  section.classList.add('active');
  countEl.textContent = `${novels.length} novels generated`;
  container.innerHTML = '';

  novels.forEach((novel, index) => {
    const card = createNovelCard(novel, index);
    card.style.animationDelay = `${index * 0.1}s`;
    container.appendChild(card);
  });

  // Attach edit sync listeners
  attachEditSyncListeners(container);

  // Ensure thumbs persist even when cards rerender
  try {
    novels.forEach((_, i) => ensureCoverThumbInCard(i));
  } catch (_) {}

  // Expand first card so template content (synopsis, chapter outline, etc.) is visible
  const firstCard = container?.querySelector('.novel-card');
  if (firstCard) firstCard.classList.add('expanded');

  // Scroll to results
  section.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function createNovelCard(novel, index) {
  const card = document.createElement('div');
  card.className = 'novel-card';
  card.dataset.index = index;

  // Characters HTML (editable)
  let charactersHtml = '';
  if (novel.characters && Array.isArray(novel.characters)) {
    charactersHtml = novel.characters.map((c, ci) => `
      <li contenteditable="true" data-novel="${index}" data-charindex="${ci}"><strong>${c.name}</strong> (${c.role}${c.age ? ', ' + c.age : ''}) — ${c.description}${c.arc ? '<br><em>Arc: ' + c.arc + '</em>' : ''}</li>
    `).join('');
    charactersHtml = `<ul>${charactersHtml}</ul>`;
  }

  // Chapters HTML (editable) — data attributes for syncing
  let chaptersHtml = '';
  if (novel.chapters && Array.isArray(novel.chapters)) {
    chaptersHtml = novel.chapters.map((ch, ci) => `
      <div class="chapter-item" data-novel="${index}" data-chapterindex="${ci}">
        <span class="chapter-num">Ch.${ch.chapterNumber}</span>
        <span class="chapter-title editable" contenteditable="true">${escapeHtml(ch.title)} — ${escapeHtml(ch.summary)}</span>
      </div>
    `).join('');
  }

  // Themes HTML
  let themesHtml = '';
  if (novel.themes && Array.isArray(novel.themes)) {
    themesHtml = novel.themes.map(t => `<span class="theme-tag">${t}</span>`).join(' · ');
  }

  const isReviewed = state.reviewedNovels.has(index);
  const coverThumb = (typeof novel?.thumbnail === 'string' && novel.thumbnail.startsWith('data:image/'))
    ? novel.thumbnail
    : (typeof novel?.cover === 'string' && novel.cover.startsWith('data:image/')) ? novel.cover : '';
  const coverHref = (typeof novel?.cover === 'string' && novel.cover.startsWith('data:image/')) ? novel.cover : coverThumb;
  card.innerHTML = `
    <div class="novel-card-header" onclick="toggleNovelCard(${index})">
      <div class="novel-info">
        <div class="novel-number">${index + 1}</div>
        <div class="novel-title editable" contenteditable="true" data-novel="${index}" data-field="title">${escapeHtml(novel.title || 'Untitled Novel')}</div>
        ${coverThumb ? `<a href="${coverHref}" target="_blank" rel="noopener" title="Open cover image"><img class="novel-cover-thumb" data-index="${index}" src="${coverThumb}" alt="Cover ${index + 1}"/></a>` : ''}
      </div>
      <div class="actions">
        ${isReviewed
    ? '<button class="btn btn-story btn-sm" onclick="event.stopPropagation(); generateFullStory(' + index + ')" id="storyBtn_' + index + '"><span class="spinner"></span><span class="btn-text">📖 Generate Full Story</span></button>'
    : ''
}
        <button class="btn btn-secondary btn-sm review-toggle ${isReviewed ? 'reviewed' : ''}" onclick="event.stopPropagation(); toggleManualReview(${index})" title="${isReviewed ? 'Revoke manual review' : 'Mark as passed manual review'}">
          ${isReviewed ? '✅ Passed' : '⬜ Review'}
        </button>
        <button class="btn btn-secondary btn-sm" onclick="event.stopPropagation(); downloadNovel(${index})" title="Download template as .txt">
          📥 Template
        </button>
        <span class="expand-icon">▼</span>
      </div>
    </div>
    <div class="novel-card-body">
      <div class="edit-hint">💡 Click any text field below to edit it</div>

      <div class="novel-field">
        <div class="novel-field-label">📖 Genre & Themes</div>
        <div class="novel-field-content editable" contenteditable="true" data-novel="${index}" data-field="genre">${escapeHtml(novel.genre || 'N/A')}${themesHtml ? ' | ' + themesHtml : ''}</div>
      </div>

      <div class="novel-field">
        <div class="novel-field-label">📂 Category</div>
        <div class="novel-field-content editable" contenteditable="true" data-novel="${index}" data-field="category">${escapeHtml(novel.category || 'N/A')}</div>
      </div>

      <div class="divider"></div>

      <div class="novel-field">
        <div class="novel-field-label">📝 Synopsis</div>
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
      </div>

      <div class="novel-field">
        <div class="novel-field-label">🌍 Background / Setting</div>
        <div class="novel-field-content editable" contenteditable="true" data-novel="${index}" data-field="background">${escapeHtml(novel.background || 'N/A')}</div>
      </div>

      <div class="novel-field">
        <div class="novel-field-label">🖼️ Thumbnail Prompt (paste into Gemini Image)</div>
        <div class="novel-field-content editable" contenteditable="true" data-novel="${index}" data-field="thumbnailPrompt">${escapeHtml(novel.thumbnailPrompt || 'N/A')}</div>
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

      <!-- Review hint when not yet passed -->
      ${!isReviewed ? '<div class="review-required-inline" id="reviewRequired_' + index + '"><span class="review-required-icon">📋</span> Mark as <strong>Passed Manual Review</strong> in the header to enable full story generation.</div>' : ''}

      <!-- Full Story Section (populated after generation) -->
      <div class="story-section" id="storySection_${index}" style="display:none">
        <div class="novel-field">
          <div class="novel-field-label">📜 Full Story <span class="editable-badge">(editable)</span></div>
          <div class="story-content editable" contenteditable="true" id="storyContent_${index}" data-story-index="${index}"></div>
        </div>
        <div class="story-actions">
          <button class="btn btn-secondary btn-sm" onclick="downloadStory(${index})">📥 Download Story .txt</button>
          <button class="btn btn-audio btn-sm" onclick="generateAudioDramaScript(${index})" id="audioScriptBtn_${index}">
            <span class="spinner"></span>
            <span class="btn-text">🎙️ Generate Audio Drama Script</span>
          </button>
        </div>
        <!-- Audio Drama Script (populated after generation - segments with Listen + Edit) -->
        <div class="audio-script-section" id="audioScriptSection_${index}" style="display:none">
          <div class="novel-field">
            <div class="novel-field-label">🎙️ Audio Drama Script <span class="editable-badge">(edit & listen to each segment)</span></div>
            <div class="script-segments" id="audioScriptSegments_${index}"></div>
          </div>
          <div class="story-actions">
            <button class="btn btn-secondary btn-sm" onclick="downloadAudioScript(${index})">📥 Download Script .txt</button>
            <button class="btn btn-audio btn-sm" onclick="generateAllAudio(${index})" id="generateAllAudioBtn_${index}" title="Uses multiple keys in parallel for faster generation">
              <span class="spinner"></span>
              <span class="btn-text">🎵 Generate Audio (parallel)</span>
            </button>
            <button class="btn btn-scene btn-sm" onclick="generateAllScenes(${index})" id="generateAllScenesBtn_${index}">
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
function toggleManualReview(index) {
  if (state.reviewedNovels.has(index)) {
    state.reviewedNovels.delete(index);
  } else {
    state.reviewedNovels.add(index);
  }
  // Re-render the card to show/hide Generate Full Story button
  const container = document.getElementById('novelsContainer');
  const card = container.querySelector(`.novel-card[data-index="${index}"]`);
  if (card) {
    const novel = state.novels[index];
    const isExpanded = card.classList.contains('expanded');
    const newCard = createNovelCard(novel, index);
    newCard.style.animationDelay = card.style.animationDelay;
    if (isExpanded) newCard.classList.add('expanded');
    // Preserve story section if already generated
    const existingStorySection = card.querySelector(`#storySection_${index}`);
    const existingContent = card.querySelector(`#storyContent_${index}`);
    const existingAudioSection = card.querySelector(`#audioScriptSection_${index}`);
    const existingSegmentsContainer = card.querySelector(`#audioScriptSegments_${index}`);
    if (existingStorySection && existingContent?.textContent) {
      const newStorySection = newCard.querySelector(`#storySection_${index}`);
      const newContent = newCard.querySelector(`#storyContent_${index}`);
      if (newStorySection && newContent) {
        newContent.textContent = existingContent.textContent;
        newStorySection.style.display = 'block';
      }
      const btn = newCard.querySelector(`#storyBtn_${index}`);
      if (btn) {
        btn.innerHTML = '<span class="btn-text">✅ Story Generated</span>';
        btn.disabled = true;
      }
    }
    if (existingAudioSection && state.audioScriptSegments[index]?.length) {
      const textEls = card.querySelectorAll(`.script-segment-text[data-audio-index="${index}"]`);
      textEls.forEach(el => {
        const sIdx = parseInt(el.dataset.segmentIndex, 10);
        if (!isNaN(sIdx)) syncAudioSegmentEdit(index, sIdx, el.textContent || '');
      });
      const newAudioSection = newCard.querySelector(`#audioScriptSection_${index}`);
      const newSegmentsContainer = newCard.querySelector(`#audioScriptSegments_${index}`);
      if (newAudioSection && newSegmentsContainer) {
        newAudioSection.style.display = 'block';
        renderAudioScriptSegments(index, newSegmentsContainer, state.audioScriptSegments[index]);
      }
    }
    card.replaceWith(newCard);
    attachEditSyncListeners(container);
  }
  showToast(
    state.reviewedNovels.has(index)
      ? 'Template marked as passed manual review. You can now generate the full story.'
      : 'Manual review revoked.',
    'info'
  );

  // Card rerender may drop the cover thumb.
  try { ensureCoverThumbInCard(index); } catch (_) {}
}

// --- Sync edits from contenteditable back to state ---
function attachEditSyncListeners(container) {
  if (!container) return;

  const syncEditable = (el) => {
    const chapterItem = el.closest('.chapter-item');
    const novelIndex = parseInt(el.dataset.novel ?? chapterItem?.dataset.novel, 10);
    if (isNaN(novelIndex) || !state.novels[novelIndex]) return;

    const field = el.dataset.field;
    const value = el.textContent?.trim() || '';

    if (field) {
      state.novels[novelIndex][field] = value;
      if (field === 'title') {
        const titleEl = container.querySelector(`.novel-card[data-index="${novelIndex}"] .novel-card-header .novel-title`);
        if (titleEl && titleEl !== el) titleEl.textContent = value || 'Untitled Novel';
      }
      if (field === 'genre' && value.includes(' | ')) {
        const parts = value.split(' | ');
        state.novels[novelIndex].genre = (parts[0] || '').trim();
        state.novels[novelIndex].themes = (parts[1] || '')
          .split(/[·•]/)
          .map(t => t.trim())
          .filter(Boolean);
      }
    }

    const chapterIndex = chapterItem?.dataset.chapterindex;
    if (chapterIndex !== undefined) {
      const chIndex = parseInt(chapterIndex, 10);
      const novel = state.novels[novelIndex];
      if (novel.chapters && novel.chapters[chIndex]) {
        const parts = value.split(' — ');
        novel.chapters[chIndex].title = (parts[0] || '').trim();
        novel.chapters[chIndex].summary = (parts[1] || value).trim();
      }
    }

    const charIndex = el.dataset.charindex;
    if (charIndex !== undefined) {
      const cIndex = parseInt(charIndex, 10);
      const novel = state.novels[novelIndex];
      if (novel.characters && novel.characters[cIndex]) {
        // Best-effort parse: "Name (role, age) — description. Arc: arc"
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
        if (!isNaN(idx)) state.stories[idx] = el.textContent || '';
        return;
      }
      if (audioIndex !== undefined && segmentIndex !== undefined) {
        const aIdx = parseInt(audioIndex, 10);
        const sIdx = parseInt(segmentIndex, 10);
        if (!isNaN(aIdx) && !isNaN(sIdx)) syncAudioSegmentEdit(aIdx, sIdx, el.textContent || '');
        return;
      }
      if (audioIndex !== undefined && segmentIndex === undefined) {
        const idx = parseInt(audioIndex, 10);
        if (!isNaN(idx)) state.audioScripts[idx] = el.textContent || '';
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
      // Live-update header title when editing in body (if there's another instance)
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
function toggleNovelCard(index) {
  const cards = document.querySelectorAll('.novel-card');
  cards.forEach((card) => {
    if (parseInt(card.dataset.index) === index) {
      card.classList.toggle('expanded');
    }
  });
}

// --- Download ---
function downloadNovel(index) {
  const novel = state.novels[index];
  if (!novel) return;

  const content = formatNovelTxt(novel, index);
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
  if (!state.novels.length) {
    showToast('No novels to download', 'error');
    return;
  }
  state.novels.forEach((_, i) => {
    setTimeout(() => downloadNovel(i), i * 300);
  });
}

function formatNovelTxt(novel, index) {
  let txt = '';
  txt += `${'='.repeat(60)}\n`;
  txt += `  NOVEL TEMPLATE #${index + 1}\n`;
  txt += `${'='.repeat(60)}\n\n`;

  txt += `TITLE: ${novel.title || 'Untitled'}\n`;
  txt += `DESCRIPTION: ${novel.synopsis || 'N/A'}\n`;
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
  txt += `SYNOPSIS\n`;
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

  const fullStory = getFullStoryText(index);
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

// --- Generate All Stories (for all reviewed templates that don't have a story yet) ---
async function handleGenerateAllStories() {
  const eligible = [...state.reviewedNovels].filter(i => !state.stories[i]);
  if (!eligible.length) {
    showToast(state.reviewedNovels.size === 0
      ? 'Mark templates as Passed Manual Review first'
      : 'All reviewed templates already have stories',
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
  const workerCount = Math.max(1, Math.min(keys.length, queue.length));
  let done = 0;
  const total = eligible.length;
  const failed = [];

  showToast(`Generating ${total} stories in parallel (${workerCount} flows)...`, 'info');

  const worker = async (workerIdx) => {
    const key = keys[workerIdx % keys.length];
    while (queue.length) {
      const index = queue.shift();
      if (index == null) break;
      try {
        await generateFullStory(index, key);
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
  const indices = [];
  (state.novels || []).forEach((novel, i) => {
    const synopsis = safeStr(novel.synopsis);
    const needsSynopsis = !synopsis || synopsis === 'N/A' || synopsis.length < 50 || synopsis.startsWith('A story:') && synopsis.length < 80;
    const chapters = novel.chapters || [];
    const needsChapters = chapters.length < 2 || (chapters.length === 1 && safeStr(chapters[0].summary).length < 50);
    if (needsSynopsis || needsChapters) indices.push(i);
  });
  return indices;
}

async function generateMissingDataForNovel(index) {
  const novel = state.novels[index];
  if (!novel) return;
  const prompt = `You are a novel template assistant. This novel needs complete template data for export.

**Title:** ${novel.title || 'Untitled'}
**Genre:** ${novel.genre || 'Fiction'}
**Draft/Core idea:** ${(novel.draftScript || '').substring(0, 800)}
${novel.background ? `**Setting:** ${novel.background.substring(0, 300)}` : ''}

Return valid JSON only (no markdown, no backticks) with exactly these two fields:
1. "synopsis" — A 3-5 sentence (or short paragraph) description of the full story, suitable for the book description.
2. "chapters" — An array of 5-10 chapter outlines. Each item: { "chapterNumber": 1, "title": "Short title", "summary": "Short summary" }. CRITICAL: For each chapter, the combined string "title — summary" must be <= 100 characters.

Example format:
{"synopsis": "Full story description here...", "chapters": [{"chapterNumber": 1, "title": "...", "summary": "..."}, ...]}`;

  const result = await callGeminiAPI(prompt);
  if (result.synopsis) novel.synopsis = result.synopsis;
  if (result.chapters && Array.isArray(result.chapters) && result.chapters.length >= 2) {
    novel.chapters = result.chapters.map(ch => ({
      chapterNumber: ch.chapterNumber || 0,
      title: ch.title || '',
      summary: ch.summary || '',
    })).filter(ch => ch.chapterNumber >= 1).sort((a, b) => a.chapterNumber - b.chapterNumber);
  }
  normalizeNovelsForExport([novel]);
  const container = document.getElementById('novelsContainer');
  const card = container?.querySelector(`.novel-card[data-index="${index}"]`);
  if (card) {
    const newCard = createNovelCard(novel, index);
    newCard.classList.add('expanded');
    card.replaceWith(newCard);
    attachEditSyncListeners(container);
    ensureCoverThumbInCard(index);
    if (state.stories[index]) {
      const storySection = document.getElementById(`storySection_${index}`);
      const storyContent = document.getElementById(`storyContent_${index}`);
      if (storySection && storyContent) {
        renderStoryChapters(index, state.stories[index]);
        storySection.style.display = 'block';
      }
      const storyBtn = document.getElementById(`storyBtn_${index}`);
      if (storyBtn) storyBtn.innerHTML = '<span class="btn-text">✅ Story Generated</span>';
    }
    newCard.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }
}

async function handleFillMissingData() {
  const indices = novelsNeedingMissingData();
  if (!indices.length) {
    showToast('All templates already have synopsis and chapter outlines.', 'success');
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
  for (const i of indices) {
    try {
      await generateMissingDataForNovel(i);
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

// --- Generate Full Story ---
async function generateFullStory(index, apiKeyOverride) {
  const novel = state.novels[index];
  if (!novel) return;

  const apiKey = apiKeyOverride || getApiKey();
  if (!apiKey) {
    showToast('Please enter your API key first', 'error');
    return;
  }

  const btn = document.getElementById(`storyBtn_${index}`);
  if (btn) {
    btn.classList.add('loading');
    btn.disabled = true;
  }

  // Auto-expand the card to show progress
  const card = document.querySelector(`.novel-card[data-index="${index}"]`);
  if (card && !card.classList.contains('expanded')) {
    card.classList.add('expanded');
  }

  showToast(`Generating full story for "${novel.title}"... This may take a moment.`, 'info');

  try {
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
        `- ${c.name} (${c.role}): ${c.description}${c.arc ? ' | Arc: ' + c.arc : ''}`
      ).join('\n');
    }

    const storyPrompt = `You are an expert novelist and creative writer. Write the FULL STORY for the following novel.

## NOVEL DETAILS
**Title:** ${novel.title}
**Genre:** ${novel.genre || 'Fiction'}
**Category:** ${novel.category || 'Fiction'}
**Writing Language:** ${novel.writingLanguage || 'English'}
**Narrator Tone:** ${novel.narratorTone || 'Third-person omniscient'}
**Background/Setting:** ${novel.background || 'Not specified'}

## SYNOPSIS
${novel.synopsis || 'Not provided'}

## DRAFT SCRIPT & CORE IDEAS
${novel.draftScript || 'Not provided'}

## CHARACTERS
${characterDetails || 'Not specified'}

## CHAPTER OUTLINE
${chapterDetails || 'Write 5-10 chapters'}

## OUTPUT FORMAT (CRITICAL)
You MUST output the story separated into chapters using EXACT markers like this, for EVERY chapter:

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
`;

    const storyText = await callGeminiAPIRawWithKey(storyPrompt, apiKeyOverride);

    // Store the story
    state.stories[index] = storyText;

    // Display it
    const storySection = document.getElementById(`storySection_${index}`);
    const storyContent = document.getElementById(`storyContent_${index}`);
    renderStoryChapters(index, storyText);
    storySection.style.display = 'block';

    // Update button
    if (btn) {
      btn.innerHTML = '<span class="btn-text">✅ Story Generated</span>';
      btn.classList.remove('loading');
    }

    showToast(`Full story generated for "${novel.title}"!`, 'success');
    if (storySection) storySection.scrollIntoView({ behavior: 'smooth', block: 'nearest' });

  } catch (error) {
    console.error('Story generation error:', error);
    const msg = error?.message || String(error);
    showToast(`Story generation failed: ${msg}`, 'error');
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

function renderStoryChapters(index, storyText) {
  const storyContent = document.getElementById(`storyContent_${index}`);
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
function downloadStory(index) {
  const novel = state.novels[index];
  const contentEl = document.getElementById(`storyContent_${index}`);
  const storyText = contentEl?.textContent?.trim() || state.stories[index];
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
async function generateAudioDramaScript(index) {
  const novel = state.novels[index];
  const contentEl = document.getElementById(`storyContent_${index}`);
  const storyText = contentEl?.textContent?.trim() || state.stories[index];
  if (!novel || !storyText) {
    showToast('Generate the full story first.', 'error');
    return;
  }

  const apiKey = getApiKey();
  if (!apiKey) {
    showToast('Please enter your API key first', 'error');
    return;
  }

  const btn = document.getElementById(`audioScriptBtn_${index}`);
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
    state.audioScriptSegments[index] = segments;
    state.audioScripts[index] = scriptText;

    const scriptSection = document.getElementById(`audioScriptSection_${index}`);
    const segmentsContainer = document.getElementById(`audioScriptSegments_${index}`);
    renderAudioScriptSegments(index, segmentsContainer, segments);
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
function getBatchAudioForSegment(audioIndex, segmentIndex) {
  const batches = state.generatedAudioBatches[audioIndex] || {};
  const audio = state.generatedAudio[audioIndex] || {};
  for (const [key, indices] of Object.entries(batches)) {
    if (indices[0] === segmentIndex) return { url: audio[key], key, indices };
  }
  return null;
}

// --- Render Audio Script Segments (each editable + Listen + Generate Audio/Scene) ---
function renderAudioScriptSegments(audioIndex, container, segments) {
  if (!container) return;
  container.innerHTML = '';
  const audioBlobs = state.generatedAudio[audioIndex] || {};
  const sceneImages = state.generatedScenes[audioIndex] || {};
  (segments || []).forEach((text, i) => {
    const seg = document.createElement('div');
    seg.className = 'script-segment';
    seg.dataset.audioIndex = String(audioIndex);
    seg.dataset.segmentIndex = String(i);
    const isSceneCue = /^\[(SCENE|INT\.|EXT\.)[^\]]*\]/i.test(text);
    const hasAudio = !!audioBlobs[i];
    const batchInfo = getBatchAudioForSegment(audioIndex, i);
    const hasBatchAudio = !!batchInfo;
    const hasScene = !!sceneImages[i] && isSceneCue;
    seg.innerHTML = `
      <div class="script-segment-row">
        <span class="segment-num">${i + 1}</span>
        <div class="segment-actions">
          <button type="button" class="btn btn-icon segment-listen" onclick="listenToSegment(${audioIndex}, ${i})" id="listenBtn_${audioIndex}_${i}" title="Listen (browser TTS)">🔊</button>
          <button type="button" class="btn btn-icon segment-gen-audio" onclick="generateAudioForSegment(${audioIndex}, ${i})" id="genAudioBtn_${audioIndex}_${i}" title="Generate AI audio">🎵</button>
          ${isSceneCue ? `<button type="button" class="btn btn-icon segment-gen-scene" onclick="generateSceneForSegment(${audioIndex}, ${i})" id="genSceneBtn_${audioIndex}_${i}" title="Generate scene image">🖼️</button>` : ''}
        </div>
        <div class="script-segment-text editable" contenteditable="true" data-audio-index="${audioIndex}" data-segment-index="${i}">${escapeHtml(text)}</div>
      </div>
      ${hasAudio ? `<div class="segment-generated-audio"><audio controls src="${audioBlobs[i]}" id="audioPlayer_${audioIndex}_${i}"></audio><a href="${audioBlobs[i]}" download="segment_${i + 1}.${getTtsProvider() === 'ai33pro' ? 'mp3' : 'wav'}" class="btn btn-icon">📥</a></div>` : ''}
      ${hasBatchAudio ? `<div class="segment-generated-audio batch-audio"><span class="batch-label">Segments ${batchInfo.indices[0] + 1}–${batchInfo.indices[batchInfo.indices.length - 1] + 1}</span><audio controls src="${batchInfo.url}" id="audioBatch_${audioIndex}_${i}"></audio><a href="${batchInfo.url}" download="batch_${batchInfo.indices[0] + 1}-${batchInfo.indices[batchInfo.indices.length - 1] + 1}.${getTtsProvider() === 'ai33pro' ? 'mp3' : 'wav'}" class="btn btn-icon">📥</a></div>` : ''}
      ${hasScene ? `<div class="segment-generated-scene"><img src="${sceneImages[i]}" alt="Scene ${i + 1}"/><a href="${sceneImages[i]}" download="scene_${i + 1}.png" class="btn btn-icon">📥</a></div>` : ''}
    `;
    container.appendChild(seg);
  });
}

// --- TTS: Listen to a single segment ---
function listenToSegment(audioIndex, segmentIndex) {
  const segments = state.audioScriptSegments[audioIndex];
  if (!segments || !segments[segmentIndex]) return;
  const raw = segments[segmentIndex];
  const text = stripTextForTTS(raw);
  if (!text) return;

  // Stop any current speech
  if (state.speakingSegment) {
    speechSynthesis.cancel();
  }

  const utterance = new SpeechSynthesisUtterance(text);
  const novel = state.novels[audioIndex];
  const lang = (novel?.writingLanguage || 'en').substring(0, 2).toLowerCase();
  const langMap = { vietnamese: 'vi-VN', english: 'en-US', japanese: 'ja-JP', korean: 'ko-KR', chinese: 'zh-CN', french: 'fr-FR', spanish: 'es-ES', german: 'de-DE', portuguese: 'pt-BR', thai: 'th-TH' };
  utterance.lang = langMap[lang] || 'en-US';
  utterance.rate = 1;
  utterance.pitch = 1;

  state.speakingSegment = { audioIndex, segmentIndex };
  const btn = document.getElementById(`listenBtn_${audioIndex}_${segmentIndex}`);
  if (btn) btn.classList.add('playing');

  utterance.onend = utterance.onerror = () => {
    state.speakingSegment = null;
    if (btn) btn.classList.remove('playing');
  };

  speechSynthesis.speak(utterance);
}

// --- Sync segment edits back to state ---
function syncAudioSegmentEdit(audioIndex, segmentIndex, newText) {
  if (!state.audioScriptSegments[audioIndex]) return;
  const segs = state.audioScriptSegments[audioIndex];
  if (segmentIndex >= 0 && segmentIndex < segs.length) {
    segs[segmentIndex] = newText;
    state.audioScripts[audioIndex] = segs.join('\n');
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

// --- Get voice for segment: narrator vs character (match gender — no male voice for female chars) ---
function getVoiceForSegment(segmentText, novel) {
  const { speaker } = parseSpeakerFromSegment(segmentText);
  if (!speaker || speaker === 'NARRATOR') {
    return document.getElementById('narratorVoice')?.value || 'Charon';
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
    return document.getElementById('femaleVoice')?.value || 'Kore';
  }
  if (gender === 'male') {
    return document.getElementById('maleVoice')?.value || 'Puck';
  }
  return document.getElementById('narratorVoice')?.value || 'Kore';
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
  const getVoice = (speaker) => {
    if (!speaker || speaker === 'NARRATOR') return document.getElementById('narratorVoice')?.value || 'Charon';
    const chars = novel?.characters || [];
    const speakerNorm = speaker.replace(/\s+/g, ' ').toUpperCase();
    const char = chars.find(c => {
      if (!c.name) return false;
      const nameNorm = c.name.trim().toUpperCase();
      return nameNorm === speakerNorm || speakerNorm.startsWith(nameNorm) || nameNorm.startsWith(speakerNorm);
    });
    const gender = (char?.gender || '').toLowerCase();
    if (gender === 'female') return document.getElementById('femaleVoice')?.value || 'Kore';
    if (gender === 'male') return document.getElementById('maleVoice')?.value || 'Puck';
    return document.getElementById('narratorVoice')?.value || 'Charon';
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
  const url = `${baseUrl}/audio/speech`;
  const response = await fetch(url, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${apiKey}`,
    },
    body: JSON.stringify({
      model: 'tts-1',
      input: text,
      voice: voiceName || 'alloy',
      response_format: 'mp3',
    }),
  });
  if (!response.ok) {
    const err = await response.json().catch(() => ({}));
    throw new Error(err?.error?.message || `AI33 TTS error: ${response.status}`);
  }
  return await response.blob();
}

// --- Call Gemini Multi-Speaker TTS (batch = fewer API calls) ---
async function callGeminiTTSMultiSpeaker(prompt, voiceMap, apiKeyOverride = null) {
  if (getTtsProvider() === 'ai33pro') {
    const speakers = Object.keys(voiceMap);
    const voice = voiceMap[speakers[0]] || 'alloy';
    const cleanPrompt = prompt.replace(/^[A-Z\s]+:\s*/gm, '').trim();
    return callAi33TTS(cleanPrompt, voice, apiKeyOverride);
  }
  const apiKey = apiKeyOverride || getTtsApiKey();
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash-preview-tts:generateContent?key=${apiKey}`;
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
    const b64 = data.candidates?.[0]?.content?.parts?.[0]?.inlineData?.data;
    if (!b64) throw new Error('No audio in TTS response');
    return pcmToWavBlob(b64);
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
  const b64 = data.candidates?.[0]?.content?.parts?.[0]?.inlineData?.data;
  if (!b64) throw new Error('No audio in TTS response');
  return pcmToWavBlob(b64);
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
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash-preview-tts:generateContent?key=${apiKey}`;
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
  const b64 = data.candidates?.[0]?.content?.parts?.[0]?.inlineData?.data;
  if (!b64) throw new Error('No audio in TTS response');
  return pcmToWavBlob(b64);
}

// --- Gemini Imagen: thumbnail/cover generation (when AI provider is Gemini and key set) ---
async function callGeminiImagen(prompt) {
  const apiKey = getApiKey();
  if (!apiKey) return null;
  const url = `https://generativelanguage.googleapis.com/v1beta/models/imagen-3.0-generate-002:predict?key=${encodeURIComponent(apiKey)}`;
  const body = {
    instances: [{ prompt: String(prompt).slice(0, 2048) }],
    parameters: { sampleCount: 1 }
  };
  const response = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });
  const rawText = await response.text();
  if (!response.ok) {
    let errMsg = rawText;
    try {
      const errJson = JSON.parse(rawText);
      errMsg = errJson.error?.message || errJson.message || errMsg;
    } catch (_) {}
    throw new Error(`Gemini Imagen: ${errMsg}`);
  }
  let data;
  try {
    data = JSON.parse(rawText);
  } catch (_) {
    throw new Error('Invalid JSON in Imagen response');
  }
  if (data.error) {
    const msg = data.error.message || data.error.code || 'Image generation failed';
    throw new Error(String(msg));
  }
  // Response shapes: predictions[].bytesBase64Encoded, .image.bytesBase64Encoded, .image.imageBytes; or generated_images[].image.imageBytes; or structValue.fields
  const pred = (data.predictions && data.predictions[0]) || (data.generated_images && data.generated_images[0]);
  if (!pred) throw new Error('No image in Gemini Imagen response');
  let b64 = pred.bytesBase64Encoded || (pred.image && (pred.image.bytesBase64Encoded || pred.image.imageBytes));
  if (typeof b64 !== 'string' && pred.structValue?.fields?.bytesBase64Encoded) {
    const f = pred.structValue.fields.bytesBase64Encoded;
    b64 = f.stringValue ?? f.string_value;
  }
  if (typeof b64 !== 'string' && pred.generatedImage?.image?.imageBytes)
    b64 = pred.generatedImage.image.imageBytes;
  if (!b64 || typeof b64 !== 'string') throw new Error('No image bytes in Gemini Imagen response');
  return `data:image/png;base64,${b64}`;
}

// --- Free image API fallback (no key required) ---
async function callFreeImageAPI(prompt, novel) {
  if (state.imageGenerationDisabled) {
    throw new Error(state.imageGenerationDisabledReason || 'Image generation is disabled');
  }
  const tone = novel?.narratorTone || '';
  const background = novel?.background || '';
  const styleHint = [tone, background].filter(Boolean).join('. ');
  const fullPrompt = styleHint
    ? `Scene image, style: ${styleHint}. ${prompt}. Digital art, high quality, atmospheric.`
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

// --- Unified image generation: Gemini Imagen when provider is Gemini + key; else free API (DeepSeek has no image API) ---
async function callImageGenerationAPI(prompt, novel) {
  const fullPrompt = `Book cover illustration for a novel. No text, no typography, no watermark.
Title concept: "${novel?.title || 'Untitled'}".
Genre: ${novel?.genre || novel?.category || 'Fiction'}.
Setting/background: ${novel?.background || 'not specified'}.
Main mood/tone: ${novel?.narratorTone || ''}.
Composition: centered subject, cinematic lighting, high detail, professional cover art.`;
  const effectivePrompt = prompt && prompt.length > 50 ? prompt : fullPrompt;
  if (getAIProvider() === 'gemini' && getApiKey()) {
    try {
      return await callGeminiImagen(effectivePrompt);
    } catch (e) {
      console.warn('Gemini Imagen failed, falling back to free API:', e?.message);
    }
  }
  return callFreeImageAPI(effectivePrompt, novel);
}

// --- Generate Audio for a single segment ---
async function generateAudioForSegment(audioIndex, segmentIndex) {
  const segments = state.audioScriptSegments[audioIndex];
  const raw = document.querySelector(`.script-segment-text[data-audio-index="${audioIndex}"][data-segment-index="${segmentIndex}"]`)?.textContent?.trim() || segments?.[segmentIndex];
  const text = stripTextForTTS(raw);
  if (!text) {
    showToast('Segment has no speakable content (cue-only).', 'error');
    return;
  }
  const novel = state.novels[audioIndex];
  const btn = document.getElementById(`genAudioBtn_${audioIndex}_${segmentIndex}`);
  if (btn) { btn.disabled = true; btn.textContent = '⏳'; }
  try {
    const blob = await callGeminiTTS(text, novel, raw);
    const url = URL.createObjectURL(blob);
    if (!state.generatedAudio[audioIndex]) state.generatedAudio[audioIndex] = {};
    state.generatedAudio[audioIndex][segmentIndex] = url;
    const container = document.getElementById(`audioScriptSegments_${audioIndex}`);
    renderAudioScriptSegments(audioIndex, container, state.audioScriptSegments[audioIndex]);
    showToast(`Audio generated for segment ${segmentIndex + 1}`, 'success');
  } catch (e) {
    showToast(`Audio failed: ${e.message}`, 'error');
  }
  if (btn) { btn.disabled = false; btn.textContent = '🎵'; }
}

// --- Generate All Audio (batched + parallel across multiple API keys) ---
async function generateAllAudio(audioIndex) {
  const novel = state.novels[audioIndex];
  const segments = state.audioScriptSegments[audioIndex];
  if (!novel || !segments?.length) {
    showToast('No script segments to generate audio from.', 'error');
    return;
  }
  const keys = getTtsApiKeys();
  if (!keys.length) { showToast('TTS requires API key(s). Configure in Settings.', 'error'); return; }
  const btn = document.getElementById(`generateAllAudioBtn_${audioIndex}`);
  btn.classList.add('loading');
  btn.disabled = true;
  const container = document.getElementById(`audioScriptSegments_${audioIndex}`);
  const batches = buildAudioBatches(segments, novel);
  if (!state.generatedAudioBatches[audioIndex]) state.generatedAudioBatches[audioIndex] = {};

  // Distribute batches across keys (parallel workers)
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
        if (!state.generatedAudio[audioIndex]) state.generatedAudio[audioIndex] = {};
        state.generatedAudio[audioIndex][keyName] = url;
        state.generatedAudioBatches[audioIndex][keyName] = batch.indices;
        done += batch.indices.length;
        renderAudioScriptSegments(audioIndex, container, state.audioScriptSegments[audioIndex]);
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
async function generateSceneForSegment(audioIndex, segmentIndex) {
  const segments = state.audioScriptSegments[audioIndex];
  const text = document.querySelector(`.script-segment-text[data-audio-index="${audioIndex}"][data-segment-index="${segmentIndex}"]`)?.textContent?.trim() || segments?.[segmentIndex];
  if (!text) return;
  const novel = state.novels[audioIndex];
  const btn = document.getElementById(`genSceneBtn_${audioIndex}_${segmentIndex}`);
  if (btn) { btn.disabled = true; btn.textContent = '⏳'; }
  try {
    const sceneDesc = text.replace(/^\[(SCENE|INT\.|EXT\.|SFX|AMB|MUSIC)[^\]]*\]\s*/i, '').trim() || text;
    const dataUrl = await callImageGenerationAPI(sceneDesc, novel);
    if (!state.generatedScenes[audioIndex]) state.generatedScenes[audioIndex] = {};
    state.generatedScenes[audioIndex][segmentIndex] = dataUrl;
    const container = document.getElementById(`audioScriptSegments_${audioIndex}`);
    renderAudioScriptSegments(audioIndex, container, state.audioScriptSegments[audioIndex]);
    showToast(`Scene generated for segment ${segmentIndex + 1}`, 'success');
  } catch (e) {
    showToast(`Scene failed: ${e.message}`, 'error');
  }
  if (btn) { btn.disabled = false; btn.textContent = '🖼️'; }
}

// --- Generate All Scenes (for scene-type segments only) ---
async function generateAllScenes(audioIndex) {
  const novel = state.novels[audioIndex];
  const segments = state.audioScriptSegments[audioIndex];
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
  const btn = document.getElementById(`generateAllScenesBtn_${audioIndex}`);
  btn.classList.add('loading');
  btn.disabled = true;
  const container = document.getElementById(`audioScriptSegments_${audioIndex}`);
  let done = 0;
  for (const i of sceneIndices) {
    const text = document.querySelector(`.script-segment-text[data-audio-index="${audioIndex}"][data-segment-index="${i}"]`)?.textContent?.trim() || segments[i];
    const sceneDesc = text.replace(/^\[(SCENE|INT\.|EXT\.)[^\]]*\]\s*/i, '').trim() || text;
    try {
      const dataUrl = await callImageGenerationAPI(sceneDesc, novel);
      if (!state.generatedScenes[audioIndex]) state.generatedScenes[audioIndex] = {};
      state.generatedScenes[audioIndex][i] = dataUrl;
      renderAudioScriptSegments(audioIndex, container, state.audioScriptSegments[audioIndex]);
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
function downloadAudioScript(index) {
  const novel = state.novels[index];
  // Sync any unsaved edits from DOM before download
  const textEls = document.querySelectorAll(`.script-segment-text[data-audio-index="${index}"]`);
  if (textEls.length) {
    const segs = [...(state.audioScriptSegments[index] || [])];
    textEls.forEach(el => {
      const sIdx = parseInt(el.dataset.segmentIndex, 10);
      if (!isNaN(sIdx)) segs[sIdx] = el.textContent || '';
    });
    state.audioScriptSegments[index] = segs;
  }
  const segments = state.audioScriptSegments[index];
  const scriptText = segments ? segments.join('\n') : state.audioScripts[index];
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
