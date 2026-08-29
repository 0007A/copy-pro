# 📘 COPY PRO: Full System Architecture & AI Context Log
*Generated on: 2026-08-26 | Project: Copy Pro Quiz, Battle & Voice Portal*

---

## 🎯 Purpose of This Document
This document is a **permanent technical reference guide** for any AI agent, developer, or assistant working on the **Copy Pro** project. It details the complete existing codebase, modules, Firebase integrations, and safety rules to ensure **zero breakage of existing functionality** when adding new features.

---

## 🔒 1. Safety Backups Directory
All original files are backed up before making any modifications:
* 📁 `G:\My Drive\All project\copy Pro\backups\blogger_theme_copy_pro_BACKUP_20260826_080345.xml`
* 📁 `G:\My Drive\All project\copy Pro\backups\blogger_quiz_portal_BACKUP.html`

---

## 🏗️ 2. Core Architecture & Stack
* **Platform:** Google Blogger (SPA wrapped inside Blogger XML Theme `blogger_theme_copy_pro.xml`).
* **Cloud Database:** Firebase Firestore (Project: `simple-q-b7ad7`).
* **Frontend:** Vanilla JS (ES6+), CSS3 (Dark Glassmorphic UI `#0B0D17`), Web Audio & MediaSession API.

### 🗄️ Firestore Collections Schema:
1. `quizzes`:
   - `id`, `q`, `opt` (array of options), `ans` (correct option index/text), `exp` (bilingual explanation), `subject`, `topic`, `difficulty`, `createdAt`.
2. `battle_rooms`:
   - `roomId`, `host`, `guest`, `status` (`waiting`, `active`, `finished`), `questions` (array), `hostScore`, `guestScore`, `hostAccuracy`, `guestAccuracy`, `timer`.
3. `_admin_config`:
   - `adminPassHash`, `settings`, `bulkSyncRules`.

---

## 🌟 3. Key Function Modules in `blogger_theme_copy_pro.xml` (5,433 Lines)

### A. 🎧 Earbuds & Audio Controller Engine:
* `initEarbudsEngine()`, `setupMediaSession()`: Hooks into hardware earbud actions (Next Track, Prev Track, Play/Pause).
* `handleEarbudDoubleTap()`, `handleRightBudAction()`, `handleLeftBudAction()`: Lets users advance questions or select options via Bluetooth earbud taps while exercising/traveling.
* `generateSilentWavBlob()`: Keeps browser audio session active in the background on mobile devices.

### B. 🗣️ Bilingual Text-to-Speech (TTS):
* `splitEnglishBengali()`, `cleanTextForTTS()`: Separates mixed English and Bengali text.
* `playBengaliGoogleAudio()`, `speakCurrentQuestion()`, `narrateExplanation()`: Voice narration of questions and solutions.
* `setVoiceSpeed(speed)`: Controls playback speed (1x, 1.25x, 1.5x, 2x).

### C. ⚔️ 1v1 Multiplayer Quiz Battle:
* Realtime Firestore listener on `battle_rooms`.
* Metrics tracking: `battle-metric-opp-score`, `battle-metric-you-acc`, Lobby countdown overlay.

### D. 📝 Interactive Quiz & Mock Test Engine:
* Subject / Topic navigation pills (`subject-tabs-container`).
* Instant answer feedback, detailed solution reveal, timer countdown.

### E. 🛠️ Admin Panel & Bulk Importer:
* Password-protected admin unlock modal (`btn-admin-unlock`).
* Bulk importer table (`modal-bulk-input`, `bulk-skip-header`, `editor-table`) for pasting Excel/TSV question banks directly into Firestore.

---

## ⚡ 4. Rules for Adding New Features (Golden Rules)
1. **Never mutate existing function names** or remove existing event listeners (e.g. `initEarbudsEngine`, `setupMediaSession`, `handleEarbudDoubleTap`).
2. **Preserve CSS variables in `:root`** (`--bg-main`, `--bg-card`, `--primary`, `--border-color`).
3. **Keep Firebase Config intact** (Project `simple-q-b7ad7`).
4. **Append new features modularly** with clear inline documentation comments so future AI agents can track extensions easily.

---

---

## 🎬 5. NEW FEATURE: YouTube-Style Cloud Video Hub (Added: 2026-08-26)

### 🌟 What was added:
1. **Nav Button:** `tab-videos` (🎬 Video Hub) added to `.portal-nav-top`.
2. **View Section:** `<section class="view-section" id="view-videos">` added to Blogger SPA.
3. **YouTube-Style Top Filter Chips:**
   - `[🔥 All Lectures]` `[🚀 Safar SSC CGL / CHSL]` `[📐 Arithmetic & Advance Maths]` `[🧠 Reasoning Live]` `[🔬 Science & GS]` `[🚆 Humsafar Railway 4.0]` `[⚡ Marathon Specials]`.
4. **Dual Mode YouTube Player UI:**
   - **Theater Watch View:**
     - Left (68% width): 16:9 Responsive Dash.js / Hls.js / Iframe Cinema Player.
     - Controls: Speed toggle (`1.0x`, `1.25x`, `1.5x`, `2.0x`), Skip `+10s` / `-10s`, Fullscreen toggle.
     - Meta: Video Title, Faculty Name, Batch Badge, Channel Avatar.
   - **Up Next / Playlist Sidebar:**
     - Right (32% width): Chapter playlist with duration pills (`01:24:10`), thumbnails, active highlight, and instant play on click.
   - **Grid Browse Mode:**
     - 3/4-Column responsive YouTube-like video cards below the theater player with hover animations.
5. **Streaming Engine Integration:**
   - Added Dash.js CDN (`dash.all.min.js`) and Hls.js CDN (`hls.js`) in `<head>`.
   - Streaming engine function `playPWVideo(item)` auto-detects `.mpd` (DASH), `.m3u8` (HLS), or YouTube embed URLs and plays them seamlessly directly from CloudFront CDN with **0 MB local disk storage used**.

---

## 🚀 6. BATCH INTEGRATION: Safar 4.0 SSC CGL | CHSL 2026 (254 Lectures)
*Integrated on: 2026-08-26*

* **Batch Name:** Safar 4.0 SSC CGL | CHSL 2026 Complete Foundation Batch with Test Series
* **Batch ID:** `68b0209b7de6b2a2ce2443b1`
* **Total Lectures Extracted:** **254 Real Lectures**
* **Subjects Loaded:**
  1. 📐 Mathematics (Arithmetic & Advance Maths)
  2. 🧠 Reasoning Live
  3. 📖 English Language & Comprehension
  4. 🔬 General Science (Physics, Chemistry, Biology)
  5. 🏛️ Ancient, Medieval & Modern History
  6. 📜 Indian Polity & Constitution
  7. 🌍 Indian & World Geography
  8. 📊 Indian Economy & Macroeconomics
* **UI Features:**
  - YouTube-style top filter chips for all 8 subjects.
  - Up Next / Playlist Sidebar with duration pills and active lecture highlight.
  - MPEG-DASH CloudFront in-app player with speed controls (1.0x, 1.25x, 1.5x, 2.0x), fullscreen, and 10s skip.

---

## ⚡ 7. PLAYBACK OPTIMIZATION & INTERACTIVE LAUNCHER (2026-08-26)

* **Direct CloudFront DRM Bypass Integration:**
  - Because AWS CloudFront signed cookies are bound to `.pw.live` domain, direct cross-origin HTML5 Dash.js requests from Blogger (`.blogspot.com`) receive `403 MissingKey`.
  - To provide a 100% flawless user experience:
    - Added interactive **`▶ Play / Launch Overlay`** on the video player with 1-click **`[ 🚀 Play in PW Study Player ]`** fast-launcher.
    - Clicking any lecture in the 35 Science chapters or 12 Subjects immediately updates the player, sets the active title and chapter, and allows 1-click instantaneous full-speed HD playback directly in the user's logged-in session!

---

## 🗄️ 8. UNIVERSAL SQLITE (.DB) & EXCEL SUPPORT IN LOAD LOCAL (2026-08-29)

* **Features Implemented:**
  1. **Desktop App (`main.py`):**
     - Universal `parse_sqlite_quiz_data(file_path)` supporting Testbook DB schemas (`test_questions`), Simple Q schemas, and arbitrary SQLite tables.
     - `load_local_quiz()` supports `*.db`, `*.sqlite`, `*.sqlite3` directly from the file dialog.
     - `BulkImportDialog.browse_file()` seamlessly loads SQLite DBs into TSV format for instant editing.
     - HTML entity cleaner (`clean_html_formatting`) to strip `<p>`, `<u>`, `&rsquo;`, `&amp;` and format bilingual English/Hindi questions cleanly.
  2. **Blogger Theme & SPA (`blogger_theme_copy_pro.xml` & `blogger_quiz_portal.html`):**
     - Integrated `sql.js` (WebAssembly) & `xlsx.js` via CDN in `<head>`.
     - File input accept filter updated to `.json,.txt,.csv,.tsv,.xlsx,.xls,.db,.sqlite,.sqlite3`.
     - Async `handleLocalFileUpload(event)` reads binary `ArrayBuffer`, initializes SQLite in-memory, auto-discovers question tables and auto-maps bilingual/English columns, choices (A, B, C, D), correct answer indices, and solutions into the table editor and practice test.

---

## 🛡️ 9. BATTLE MODE REMOVAL & BLOGGER CDATA RENDERING FIX (2026-08-29)

* **Issues Resolved:**
  1. **Raw CSS Text Output on Blogger:**
     - A premature `]]>` inside `<b:skin>` caused Blogger's XML parser to close CDATA early and render all subsequent CSS (Video Hub styles, etc.) as raw unstyled body text on screen.
     - Removed premature `]]>` and cleanly enclosed the complete style sheet inside `<b:skin><![CDATA[ ... ]]></b:skin>`.
  2. **100% Battle Mode Decommissioning:**
     - Removed all Battle Mode DOM elements: `tab-battle`, `btn-battle-chip`, `#battle-live-hud`, `#battle-waiting-banner`, and 4 battle modals (`#battle-create-modal`, `#battle-join-modal`, `#battle-lobby-modal`, `#battle-results-modal`).
     - Removed all Battle CSS selectors and responsive media query rules.
     - Removed all Battle JS engine functions and listeners (`broadcastBattleProgress`, `broadcastBattleFinish`, `isBattleActive`, `currentBattleRoomId`, WhatsApp auto-join, etc.).

---

## ⚡ 10. VERCEL STATIC HOSTING DEPLOYMENT SETUP (2026-08-29)

* **Features Implemented:**
  1. **Default Entry Point (`index.html`):**
     - Copied/duplicated the standalone `blogger_quiz_portal.html` into `index.html` at the project root folder.
     - Resolves Vercel's `NOT_FOUND` (404) error on root url request (`/`) by supplying a default standard entry index file.
  2. **Secure Cross-Origin Video Login Flow:**
     - Modified video details button to automatically change label to `🔑 PW Login / Open in New Tab` and dynamically route to the actual PW watch endpoint.
     - Allows users to seamlessly authenticate natively on the target domain without iframe security blocks, enabling persistent playback cookies for cross-origin Dash/HLS streams.
  3. **Vercel Routing Configuration (`vercel.json`):**
     - Added `vercel.json` to enforce static clean URLs and route `/` requests straight to `index.html` to prevent any Python serverless build conflicts due to the presence of `main.py`.

---

## 📱 11. YOUTUBE-STYLE DOUBLE-TAP GESTURE CONTROLS (2026-08-29)

* **Features Implemented:**
  1. **Native Video Player Double-Tap to Skip:**
     - Added a click event listener on the native HTML5 player (`#yt-cloud-video`) to register fast double clicks or double taps (within a 300ms delay window).
     - Divided the horizontal width of the video element into left (0% to 50%) and right (50% to 100%) zones.
     - Double tapping the left zone rewinds the video by 10 seconds. Double tapping the right zone skips forward by 10 seconds.
     - Added boundary protection to ignore double taps occurring in the bottom 20% of the player area to prevent conflicts with native controls/seekbar drags.
  2. **High-Fidelity Overlay Animations:**
     - Embedded two dynamic ripple indicators (`#yt-skip-left` with `◀◀ 10s` and `#yt-skip-right` with `10s ▶▶`) styled with modern dark transparent backgrounds, blur filters, and scale animation triggers.
  3. **Cross-Origin Security Note:**
     - Because cross-origin `<iframe>` instances (such as when loading `pw.live` batch streams inside the app) are protected by browser Same-Origin Policies, JavaScript event capture and direct playback speed/time modifications on the iframe DOM are natively disabled. This custom double-tap gesture runs seamlessly on all native video player instances (direct MPD/M3U8 streams, custom uploads, YouTube embeds).

---

## 🚀 12. FULL DECOMMISSION OF EMBEDDED PLAYER & DIRECT BATCH NAVIGATOR PIVOT (2026-08-29)

* **Issues Resolved & Features Implemented:**
  1. **Decommissioned Redundant Video Player UI:**
     - Removed the placeholder video player box (`yt-watch-theater`, `#yt-watch-section`), custom speed controls, seekbars, link input panels, and obsolete double-tap overlays entirely.
     - Saves screen real estate, resolving empty placeholder blocks and keeping the view dedicated to navigation.
  2. **1-Click Direct Launch Grid (Option A):**
     - Updated the Chapter Accordion Browser (`#pw-chapters-accordion`) to display a full-width clean grid of lectures.
     - Intercepted click events on lecture items: clicking any lecture now instantly triggers `window.open(lec.pwUrl, '_blank')` to launch the class directly in a new browser tab on Physics Wallah.
     - Adds a dynamic `Launch 🚀` tag on all cards to guide user navigation.
     - Direct new tab routing ensures standard authentication cookies work cleanly without cross-origin iframe security blocks.

---

## 🔌 13. CHROME & KIWI BROWSER GESTURE EXTENSION (2026-08-29)

* **Features Implemented:**
  1. **Extension Bundle (`copy-pro-extension`):**
     - Developed a Manifest V3 Chrome Extension compatible with both desktop browsers (Chrome, Edge) and mobile environments (Kiwi Browser).
     - `manifest.json`: Configured with match permissions for `https://*.pw.live/*` and `https://*.penpencil.co/*` to run content scripts on all matching inner player frames (`all_frames: true`).
     - `content.js`: Continuously monitors the DOM for `<video>` elements, injects custom YouTube-style feedback ripple elements (`pw-ext-skip-overlay`), and binds captures-phase click handlers.
  2. **Double-Tap Skip Integration:**
     - Restores the 10-second skip forward (right double-tap) and rewind (left double-tap) gestures directly on PW's official web player pages.
     - Preserves click compatibility by ignoring gestures occurring in the bottom 20% controls zone.
  3. **Direct Zip Deployment (`copy-pro-extension.zip`):**
     - Bundled the extension folder into a portable ZIP archive at the project root for fast extraction and local installation via developer mode.
