# OneNote Web Clone App 

* Author: Ian Schrauth
* GitHub Handle: IHS5452
* Contact Email: ian.h.schrauth {-a-t-} gmail {-d-o-t-} com
  * Note: You may only email me if your message is related to an urgent need for this project, or offering employment in IT. Anything outside of that scope will have a very disrespectfull reply.

A lightweight, browser-based note-taking application inspired by OneNote. This app operates directly on your local file system via the **File System Access API**, allowing you to create, edit, organize, and export `.docx` notes organized by subject folders and week files without needing a server or database.

---

### Key Features

* **Direct File System Sync:** Create and open real local directory folders (`Subjects`) containing `.docx` documents (`Weeks`).
* **Natural Alphanumeric Sorting:** Automatic ordering for subjects and weeks (e.g., *Week 2* correctly sorts before *Week 10*).
* **Rich Text & Media Support:** Full text-editing capabilities with image pasting, URL embedding, text formatting, and lists.
* **Keyboard Shortcuts:**
  * **`Cmd/Ctrl + 1`**: Rename selected subject directory.
  * **`Cmd/Ctrl + 2`**: Rename selected week file.
  * **`Cmd/Ctrl + S`**: Save active file directly to your disk.
  * **`Cmd/Ctrl + E`**: Trigger document export.
* **Multi-Format Exporting:** Export your notes anytime to **DOCX**, **PDF**, **Markdown**, or **TXT**.
* **IndexedDB History:** Saves recent directory handles so you can quickly reopen recent notebooks.
* **Dark/Light Theme:** Matches system preference settings automatically.

---

### Getting Started

#### Prerequisites
A modern desktop web browser supporting the **File System Access API** (Google Chrome, Microsoft Edge, Opera, etc.).

#### Dependencies
Include these libraries in your `index.html` head section:
* [docx](https://docx.js.org/) (for OpenXML generation)
* [Mammoth.js](https://github.com/mwilliamson/mammoth.js) (for `.docx` to HTML conversion)
* [html2pdf.js](https://github.com/eKoopmans/html2pdf.js) (for PDF exporting)

```html
<!-- Required CDN Libraries -->
<script src="[https://unpkg.com/docx@8.2.2/build/index.umd.js](https://unpkg.com/docx@8.2.2/build/index.umd.js)"></script>
<script src="[https://cdnjs.cloudflare.com/ajax/libs/mammoth/1.6.0/mammoth.browser.min.js](https://cdnjs.cloudflare.com/ajax/libs/mammoth/1.6.0/mammoth.browser.min.js)"></script>
<script src="[https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js](https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js)"></script>
```
---

#### Project Structure
```
├── index.html        # App UI layout (Sidebar, Toolbar, Content Editable Canvas)
├── style.css         # Styling and Theme Variables (Light/Dark mode)
└── app.js            # Main application logic & File System Access API handlers
```
---

#### Required DOM Container Elements
Your index.html file must include the following element IDs for app.js to bind correctly:

```
<!-- Action Controls -->
<button id="btn-open-folder">Open Root Folder</button>
<button id="btn-add-subject">+ Add Subject</button>
<button id="btn-add-week">+ Add Week</button>
<select id="select-recents"></select>

<!-- Sidebar Navigation -->
<ul id="list-subjects"></ul>
<ul id="list-weeks"></ul>

<!-- Editor Structure -->
<div id="empty-state">Select a note to get started</div>
<div id="editor-content" style="display: none;">
  <div id="toolbar">
    <select id="export-format">
      <option value="docx">DOCX</option>
      <option value="pdf">PDF</option>
      <option value="md">Markdown</option>
      <option value="txt">Text</option>
    </select>
    <button id="btn-save">Export</button>
  </div>
  
  <input type="text" id="page-title" placeholder="Week Title..." />
  <span id="page-datetime"></span>
  <div id="page-body" contenteditable="true"></div>
</div>
```
---
#### How It Works
1) Open Root Folder: Click Open Subject Root Folder to pick a working directory on your local device.

2) Subject Organization: Click + Add Subject to create subdirectories inside your root folder.

3) Week Notes: Select a subject, then click + Add Week to generate .docx files.

4) Auto-Saving: Changes made inside the editor automatically compile and save back to the selected .docx file when losing focus (blur) or typing stops (1-second debounced input).
