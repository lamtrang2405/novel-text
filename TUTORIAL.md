# How to Use AI Novel Template Generator

A step-by-step guide for **[https://lamtrang2405.github.io/novel-text/](https://lamtrang2405.github.io/novel-text/)**

---

## Overview

This tool turns one creative idea into multiple novel templates, then expands them into full stories and audio drama scripts. You can generate narrator audio and scene images.

---

## Flow Diagram

```
┌─────────────────────────────────────────────────────────────────────────────┐
│  1. API SETUP  →  2. CREATIVE BRIEF  →  3. GENERATE TEMPLATES  →  4. REVIEW │
└─────────────────────────────────────────────────────────────────────────────┘
         │                    │                        │                    │
         ▼                    ▼                        ▼                    ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│  5. FULL STORY  →  6. AUDIO DRAMA SCRIPT  →  7. AUDIO & SCENES  →  8. DOWNLOAD │
└─────────────────────────────────────────────────────────────────────────────┘
```

---

## Step 1: API Configuration

1. **Choose AI Provider**
   - **Google Gemini** — For templates, story, script, and voice (TTS). One key for everything.
   - **DeepSeek** — For templates, story, and script only. You’ll need a separate Gemini key for voice later.

2. **Get API keys**
   - **Gemini**: [Google AI Studio](https://aistudio.google.com) → Create API Key
   - **DeepSeek**: [platform.deepseek.com](https://platform.deepseek.com) → API Keys

3. **Enter your key(s)**
   - Paste into the **API Key(s)** field.
   - For DeepSeek, add a Gemini key in **Gemini API Key (for TTS)** when you use voice.

4. **Test**
   - Click **Test** to verify the connection.
   - You should see: `Gemini API (gemini-2.5-flash): OK` or `DeepSeek API: Connection OK`.

---

## Step 2: Creative Brief

**Required**

1. **Number of Novels** — 1–10 templates (e.g. 3).
2. **Writing Language** — Vietnamese, English, Japanese, etc.
3. **Master Prompt / Creative Brief** — Describe your idea, genre, themes. Example:
   > A dark fantasy series where emotions become physical powers. Themes of trauma, healing, and connection. War between the Empathic Order and the Void Syndicate.

**Optional (Template Details)**

- **Draft Script & Core Ideas** — Main plot points and scenario.
- **Character System** — Character types and relationships.
- **Author Name** — Pen name.
- **Release Date** — Target release.
- **Narrator Tone & Background** — POV, mood, atmosphere.

**Optional (Reference Material)**

- **Reference Text** — Paste samples for style or tone.
- **File Upload** — Upload a `.txt` file as reference.

---

## Step 3: Generate Novel Templates

1. Click **✨ Generate Novel Templates**.
2. Wait for progress (usually 30–120 seconds).
3. You’ll see cards for each novel with: title, synopsis, characters, chapters, themes, genre.

---

## Step 4: Review Templates

1. Read each template.
2. Use **Edit** to change text directly.
3. If satisfied, toggle **Passed Manual Review**.
4. **Generate Full Story** appears only for templates that passed review.

---

## Step 5: Generate Full Story

1. For each approved template, click **Generate Full Story**.
2. Wait (can be several minutes; output is long).
3. The full story appears in the card and is editable.
4. You can download it with **Download Story**.

---

## Step 6: Generate Audio Drama Script

1. Click **Generate Audio Drama Script** on a novel with a full story.
2. The script is split into narration, dialogue, and sound cues.
3. Edit each segment if needed.
4. Each segment has a **Listen** button for TTS preview.

---

## Step 7: Generate Audio & Scenes

**Audio (Gemini TTS)**

- **Generate All Audio** — Creates voice for all segments.
- If using DeepSeek, you need a Gemini key for TTS.
- Multiple API keys speed up generation.

**Scenes**

- **Generate Scene** — Creates an image for that segment.
- Uses your narrator tone and story background in the prompt.

---

## Step 8: Download

- **Download All** — ZIP of all novel templates.
- **Download Story** — Full story text per novel.
- **Download Script** — Audio drama script as `.txt`.
- Audio files are played in-browser; use browser tools to save if needed.

---

## Quick Reference

| Action            | Requires                |
|-------------------|-------------------------|
| Generate templates| API key + Creative brief |
| Full story        | Passed review template  |
| Audio drama script| Full story              |
| TTS audio         | Gemini API key          |
| Scene images      | Works after script       |

---

## Tips

- Start with 1–2 novels for faster testing.
- Use Vietnamese or your target language in **Writing Language** for localized output.
- Multiple keys help parallelize audio generation.
- Reference text helps keep style consistent.
