let rootHandle = null;
let selectedSubjectHandle = null;
let selectedWeekHandle = null;
let recentsMap = new Map();

const btnOpenFolder = document.getElementById('btn-open-folder');
const btnAddSubject = document.getElementById('btn-add-subject');
const btnAddWeek = document.getElementById('btn-add-week');
const selectRecents = document.getElementById('select-recents');
const btnSave = document.getElementById('btn-save');

const listSubjects = document.getElementById('list-subjects');
const listWeeks = document.getElementById('list-weeks');
const emptyState = document.getElementById('empty-state');
const editorContent = document.getElementById('editor-content');
const toolbar = document.getElementById('toolbar');
const pageTitle = document.getElementById('page-title');
const pageDatetime = document.getElementById('page-datetime');
const pageBody = document.getElementById('page-body');

// --- IndexedDB for Persisting Handles ---
function openDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open('OneNoteWebDB', 1);
    req.onupgradeneeded = () => req.result.createObjectStore('handles');
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function saveHandleToDB(name, handle) {
  const db = await openDB();
  const tx = db.transaction('handles', 'readwrite');
  tx.objectStore('handles').put(handle, name);
}

async function loadRecentsFromDB() {
  const db = await openDB();
  const tx = db.transaction('handles', 'readonly');
  const store = tx.objectStore('handles');
  const keysReq = store.getAllKeys();
  const valuesReq = store.getAll();

  return new Promise((resolve) => {
    tx.oncomplete = () => {
      recentsMap.clear();
      const keys = keysReq.result;
      const values = valuesReq.result;
      keys.forEach((key, idx) => recentsMap.set(key, values[idx]));
      updateRecentsDropdownUI();
      resolve();
    };
  });
}

function updateRecentsDropdownUI() {
  selectRecents.innerHTML = '<option value="" disabled selected>Recent Subjects</option>';
  for (const [name] of recentsMap) {
    const opt = document.createElement('option');
    opt.value = name;
    opt.textContent = name;
    selectRecents.appendChild(opt);
  }
}

window.addEventListener('DOMContentLoaded', loadRecentsFromDB);

async function verifyPermission(fileHandle, readWrite = true) {
  const options = { mode: readWrite ? 'readwrite' : 'read' };
  if ((await fileHandle.queryPermission(options)) === 'granted') return true;
  if ((await fileHandle.requestPermission(options)) === 'granted') return true;
  return false;
}

// --- Root Directory Loading ---
btnOpenFolder.addEventListener('click', async () => {
  try {
    rootHandle = await window.showDirectoryPicker({ mode: 'readwrite' });
    await saveHandleToDB(rootHandle.name, rootHandle);
    await loadRecentsFromDB();
    selectRecents.value = rootHandle.name;
    await loadSubjects();
  } catch (err) {
    console.error("Folder loading cancelled or failed:", err);
  }
});

selectRecents.addEventListener('change', async (e) => {
  const folderName = e.target.value;
  const handle = recentsMap.get(folderName);
  if (!handle) return;

  try {
    const hasPermission = await verifyPermission(handle, true);
    if (!hasPermission) {
      alert("Permission to access this directory was denied.");
      return;
    }
    rootHandle = handle;
    await loadSubjects();
  } catch (err) {
    console.error("Failed to restore folder:", err);
    alert("Could not access recent folder. Please open it using 'Open Subject Root Folder'.");
  }
});

// --- Subject List Management ---
async function loadSubjects() {
  listSubjects.innerHTML = '';
  listWeeks.innerHTML = '';
  resetEditor();

  if (!rootHandle) return;

  for await (const entry of rootHandle.values()) {
    if (entry.kind === 'directory') {
      addSubjectUIElement(entry);
    }
  }
}

function addSubjectUIElement(entryHandle) {
  const li = document.createElement('li');
  
  const label = document.createElement('span');
  label.textContent = entryHandle.name;
  li.appendChild(label);

  const delBtn = document.createElement('span');
  delBtn.className = 'delete-icon';
  delBtn.innerHTML = '✕';
  delBtn.title = 'Delete Subject';
  delBtn.onclick = async (e) => {
    e.stopPropagation();
    if (confirm(`Are you sure you want to delete "${entryHandle.name}" and all its contents?`)) {
      try {
        await rootHandle.removeEntry(entryHandle.name, { recursive: true });
        li.remove();
        if (selectedSubjectHandle === entryHandle) {
          selectedSubjectHandle = null;
          listWeeks.innerHTML = '';
          resetEditor();
        }
      } catch (err) {
        alert(`Failed to delete subject: ${err.message}`);
      }
    }
  };
  li.appendChild(delBtn);

  li.onclick = () => selectSubject(entryHandle, li);
  listSubjects.appendChild(li);
  return li;
}

async function selectSubject(handle, element) {
  setActive(listSubjects, element);
  selectedSubjectHandle = handle;
  selectedWeekHandle = null;
  listWeeks.innerHTML = '';

  for await (const entry of handle.values()) {
    if (entry.kind === 'file' && entry.name.endsWith('.docx')) {
      addWeekUIElement(entry);
    }
  }
}

// --- Week List Management ---
function addWeekUIElement(fileHandle) {
  const li = document.createElement('li');
  
  const label = document.createElement('span');
  label.textContent = fileHandle.name.replace('.docx', '');
  li.appendChild(label);

  const delBtn = document.createElement('span');
  delBtn.className = 'delete-icon';
  delBtn.innerHTML = '✕';
  delBtn.title = 'Delete Week';
  delBtn.onclick = async (e) => {
    e.stopPropagation();
    if (confirm(`Are you sure you want to delete "${fileHandle.name}"?`)) {
      try {
        await selectedSubjectHandle.removeEntry(fileHandle.name);
        li.remove();
        if (selectedWeekHandle === fileHandle) {
          selectedWeekHandle = null;
          resetEditor();
        }
      } catch (err) {
        alert(`Failed to delete week: ${err.message}`);
      }
    }
  };
  li.appendChild(delBtn);

  li.onclick = () => loadWeek(fileHandle, li);
  listWeeks.appendChild(li);
  return li;
}

async function loadWeek(fileHandle, element) {
  setActive(listWeeks, element);
  selectedWeekHandle = fileHandle;

  const file = await fileHandle.getFile();
  const arrayBuffer = await file.arrayBuffer();
  const result = await mammoth.convertToHtml({ arrayBuffer: arrayBuffer });
  
  pageTitle.value = fileHandle.name.replace('.docx', '');
  pageDatetime.textContent = new Date(file.lastModified).toLocaleString('en-US', {
    weekday: 'long', year: 'numeric', month: 'long', day: 'numeric', hour: '2-digit', minute: '2-digit'
  });
  
  pageBody.innerHTML = result.value || '<p>Start typing here...</p>';

  emptyState.style.display = 'none';
  editorContent.style.display = 'block';
  toolbar.style.display = 'flex';
}

// --- Dynamic Creation Handlers ---
btnAddSubject.addEventListener('click', async () => {
  if (!rootHandle) return alert("Please open or select a root folder first.");

  try {
    const hasPermission = await verifyPermission(rootHandle, true);
    if (!hasPermission) return alert("Write permission was denied for the root directory.");

    const subjectName = prompt("Enter Subject Name:");
    if (!subjectName || !subjectName.trim()) return;

    const cleanName = subjectName.trim();
    const newDirHandle = await rootHandle.getDirectoryHandle(cleanName, { create: true });

    const newLi = addSubjectUIElement(newDirHandle);
    selectSubject(newDirHandle, newLi);
  } catch (err) {
    console.error("Subject creation error:", err);
    alert(`Could not create subject folder: ${err.message}`);
  }
});

btnAddWeek.addEventListener('click', async () => {
  if (!selectedSubjectHandle) return alert("Please select a Subject first.");

  try {
    const hasPermission = await verifyPermission(selectedSubjectHandle, true);
    if (!hasPermission) return alert("Write permission denied.");

    const weekName = prompt("Enter Week Name (e.g. Week 1):");
    if (!weekName || !weekName.trim()) return;

    const fileName = `${weekName.trim()}.docx`;
    const initialContent = `<!DOCTYPE html><html><body><h1>${weekName.trim()}</h1><p>Notes here...</p></body></html>`;
    const blob = htmlDocx.asBlob(initialContent);

    const newFileHandle = await selectedSubjectHandle.getFileHandle(fileName, { create: true });
    const writable = await newFileHandle.createWritable();
    await writable.write(blob);
    await writable.close();

    const newLi = addWeekUIElement(newFileHandle);
    loadWeek(newFileHandle, newLi);
  } catch (err) {
    console.error("Week creation error:", err);
    alert(`Failed to create Week file: ${err.message}`);
  }
});

// --- Formatting Actions ---
function execCmd(command) { document.execCommand(command, false, null); }
function execCmdArg(command, arg) { document.execCommand(command, false, arg); }

function insertLink() {
  const url = prompt("Enter URL:");
  if (!url) return;
  const selectedText = window.getSelection().toString();
  const customText = prompt("Enter link text:", selectedText || url);
  
  if (customText && customText !== selectedText) {
    const linkHtml = `<a href="${url}" target="_blank" style="color:#7719aa;">${customText}</a>`;
    document.execCommand('insertHTML', false, linkHtml);
  } else {
    document.execCommand('createLink', false, url);
  }
}

function insertImage() {
  const url = prompt("Enter Image URL:");
  if (url) document.execCommand('insertImage', false, url);
}

// --- Export Handlers ---
btnSave.addEventListener('click', async () => {
  if (!selectedWeekHandle) return;

  const format = document.getElementById('export-format').value;
  const title = pageTitle.value || "Untitled Week";
  const htmlBody = pageBody.innerHTML;

  try {
    if (format === 'docx') {
      const docxHtml = `<!DOCTYPE html><html><head><meta charset="utf-8"></head><body><h1>${title}</h1>${htmlBody}</body></html>`;
      const convertedBlob = htmlDocx.asBlob(docxHtml);
      const writable = await selectedWeekHandle.createWritable();
      await writable.write(convertedBlob);
      await writable.close();
      alert('Page saved back to DOCX file successfully!');
    } else if (format === 'pdf') {
      const opt = { margin: 0.5, filename: `${title}.pdf`, html2canvas: { scale: 2 } };
      const pdfWrapper = document.createElement('div');
      pdfWrapper.innerHTML = `<h1>${title}</h1><hr/><br/>${htmlBody}`;
      html2pdf().set(opt).from(pdfWrapper).save();
    } else if (format === 'md') {
      const markdownText = `# ${title}\n\n` + htmlToMarkdown(pageBody);
      downloadFile(`${title}.md`, markdownText, 'text/markdown');
    } else if (format === 'txt') {
      const plainText = `${title}\n${'='.repeat(title.length)}\n\n` + pageBody.innerText;
      downloadFile(`${title}.txt`, plainText, 'text/plain');
    }
  } catch (err) {
    console.error(err);
    alert('Saving error encountered.');
  }
});

function htmlToMarkdown(element) {
  let output = '';
  for (const node of element.childNodes) {
    if (node.nodeType === Node.TEXT_NODE) {
      output += node.textContent;
    } else if (node.nodeType === Node.ELEMENT_NODE) {
      const tag = node.tagName.toLowerCase();
      if (tag === 'p') output += htmlToMarkdown(node) + '\n\n';
      else if (tag === 'b' || tag === 'strong') output += `**${htmlToMarkdown(node)}**`;
      else if (tag === 'i' || tag === 'em') output += `*${htmlToMarkdown(node)}*`;
      else if (tag === 'u') output += `_${htmlToMarkdown(node)}_`;
      else if (tag === 'a') output += `[${htmlToMarkdown(node)}](${node.getAttribute('href')})`;
      else if (tag === 'li') output += `* ${htmlToMarkdown(node)}\n`;
      else output += htmlToMarkdown(node);
    }
  }
  return output;
}

function downloadFile(filename, content, mimeType) {
  const blob = new Blob([content], { type: mimeType });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  a.click();
  URL.revokeObjectURL(a.href);
}

function setActive(listContainer, element) {
  Array.from(listContainer.children).forEach(child => child.classList.remove('active'));
  element.classList.add('active');
}

function resetEditor() {
  emptyState.style.display = 'flex';
  editorContent.style.display = 'none';
  toolbar.style.display = 'none';
}
