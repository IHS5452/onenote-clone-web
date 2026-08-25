const docx = window.docx;

let rootHandle = null;
let selectedSubjectHandle = null;
let selectedWeekHandle = null;
let recentsMap = new Map();
let autoSaveTimer = null;

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

// --- DOM Node Formatting Helper ---
function isNodeFormatted(node, targetTag) {
  let current = node.parentElement;
  while (current && current.id !== 'page-body') {
    if (current.tagName.toUpperCase() === targetTag) return true;
    current = current.parentElement;
  }
  return false;
}

// --- Recursive Parser for Text & Styling Runs ---
function parseElementToDocxRuns(node) {
  const runs = [];
  const docxLib = window.docx || docx;

  node.childNodes.forEach(child => {
    if (child.nodeType === Node.TEXT_NODE) {
      if (child.textContent) {
        runs.push(new docxLib.TextRun({
          text: child.textContent,
          bold: isNodeFormatted(child, 'B') || isNodeFormatted(child, 'STRONG'),
          italics: isNodeFormatted(child, 'I') || isNodeFormatted(child, 'EM'),
          underline: isNodeFormatted(child, 'U') ? {} : undefined,
        }));
      }
    } else if (child.nodeType === Node.ELEMENT_NODE) {
      const tag = child.tagName.toUpperCase();
      if (tag === 'BR') {
        runs.push(new docxLib.TextRun({ break: 1 }));
      } else {
        runs.push(...parseElementToDocxRuns(child));
      }
    }
  });

  return runs;
}

// --- OpenXML Binary Saver ---
async function saveCurrentWeekToFile() {
  if (!selectedWeekHandle) return;
  try {
    const docxLib = window.docx || docx;
    const paragraphs = [];

    // Process only the canvas content (does NOT inject the week title as an H1)
    const blockElements = Array.from(pageBody.childNodes);
    
    if (blockElements.length === 0) {
      paragraphs.push(new docxLib.Paragraph({ text: "" }));
    } else {
      blockElements.forEach(node => {
        if (node.nodeType === Node.TEXT_NODE) {
          if (node.textContent.trim()) {
            paragraphs.push(new docxLib.Paragraph({
              children: [new docxLib.TextRun(node.textContent)],
            }));
          }
        } else if (node.nodeType === Node.ELEMENT_NODE) {
          const tag = node.tagName.toUpperCase();
          if (tag === 'UL' || tag === 'OL') {
            node.querySelectorAll('li').forEach(li => {
              paragraphs.push(new docxLib.Paragraph({
                children: parseElementToDocxRuns(li),
                bullet: { level: 0 }
              }));
            });
          } else if (tag === 'H1') {
            paragraphs.push(new docxLib.Paragraph({
              children: parseElementToDocxRuns(node),
              heading: docxLib.HeadingLevel.HEADING_1,
            }));
          } else if (tag === 'H2') {
            paragraphs.push(new docxLib.Paragraph({
              children: parseElementToDocxRuns(node),
              heading: docxLib.HeadingLevel.HEADING_2,
            }));
          } else {
            const runs = parseElementToDocxRuns(node);
            paragraphs.push(new docxLib.Paragraph({
              children: runs.length > 0 ? runs : [new docxLib.TextRun(node.innerText || "")]
            }));
          }
        }
      });
    }

    const doc = new docxLib.Document({
      sections: [{
        properties: {},
        children: paragraphs,
      }],
    });

    const blob = await docxLib.Packer.toBlob(doc);
    const writable = await selectedWeekHandle.createWritable();
    await writable.write(blob);
    await writable.close();
  } catch (err) {
    console.error("DOCX Binary Save Error:", err);
  }
}

// --- Root Directory Loading ---
btnOpenFolder.addEventListener('click', async (e) => {
  e.preventDefault();
  try {
    await saveCurrentWeekToFile();
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
    await saveCurrentWeekToFile();
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
    e.preventDefault();
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

  li.onclick = async (e) => {
    e.preventDefault();
    await selectSubject(entryHandle, li);
  };
  listSubjects.appendChild(li);
  return li;
}

async function selectSubject(handle, element) {
  await saveCurrentWeekToFile();
  setActive(listSubjects, element);
  selectedSubjectHandle = handle;
  selectedWeekHandle = null;
  listWeeks.innerHTML = '';
  resetEditor();

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
    e.preventDefault();
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

  li.onclick = async (e) => {
    e.preventDefault();
    await loadWeek(fileHandle, li);
  };
  listWeeks.appendChild(li);
  return li;
}

async function loadWeek(fileHandle, element) {
  await saveCurrentWeekToFile();
  setActive(listWeeks, element);
  selectedWeekHandle = fileHandle;

  try {
    const file = await fileHandle.getFile();
    const arrayBuffer = await file.arrayBuffer();
    
    const result = await mammoth.convertToHtml({ arrayBuffer: arrayBuffer });
    
    pageTitle.value = fileHandle.name.replace('.docx', '');
    pageDatetime.textContent = new Date(file.lastModified).toLocaleString('en-US', {
      weekday: 'long', year: 'numeric', month: 'long', day: 'numeric', hour: '2-digit', minute: '2-digit'
    });
    
    const contentHtml = result.value ? result.value.trim() : '';
    pageBody.innerHTML = contentHtml.length > 0 ? contentHtml : '<p><br></p>';

    emptyState.style.display = 'none';
    editorContent.style.display = 'block';
    toolbar.style.display = 'flex';
  } catch (err) {
    console.error("Mammoth DOCX parsing failed:", err);
    alert("Could not read file. Ensure it is a valid .docx document.");
  }
}

// --- Auto-Save Event Controls ---
pageBody.addEventListener('blur', saveCurrentWeekToFile);
pageTitle.addEventListener('blur', saveCurrentWeekToFile);

pageBody.addEventListener('input', () => {
  clearTimeout(autoSaveTimer);
  autoSaveTimer = setTimeout(saveCurrentWeekToFile, 1000);
});

pageTitle.addEventListener('input', () => {
  clearTimeout(autoSaveTimer);
  autoSaveTimer = setTimeout(saveCurrentWeekToFile, 1000);
});

// --- Dynamic File Handlers ---
btnAddSubject.addEventListener('click', async (e) => {
  e.preventDefault();
  if (!rootHandle) return alert("Please open or select a root folder first.");

  try {
    const hasPermission = await verifyPermission(rootHandle, true);
    if (!hasPermission) return alert("Write permission was denied.");

    const subjectName = prompt("Enter Subject Name:");
    if (!subjectName || !subjectName.trim()) return;

    await saveCurrentWeekToFile();

    const cleanName = subjectName.trim();
    const newDirHandle = await rootHandle.getDirectoryHandle(cleanName, { create: true });

    const newLi = addSubjectUIElement(newDirHandle);
    await selectSubject(newDirHandle, newLi);
  } catch (err) {
    console.error("Subject creation error:", err);
    alert(`Could not create subject folder: ${err.message}`);
  }
});

btnAddWeek.addEventListener('click', async (e) => {
  e.preventDefault();
  if (!selectedSubjectHandle) return alert("Please select a Subject first.");

  try {
    const hasPermission = await verifyPermission(selectedSubjectHandle, true);
    if (!hasPermission) return alert("Write permission denied.");

    const weekName = prompt("Enter Week Name (e.g. Week 1):");
    if (!weekName || !weekName.trim()) return;

    await saveCurrentWeekToFile();

    const fileName = `${weekName.trim()}.docx`;
    const docxLib = window.docx || docx;
    
    // Create baseline clean file
    const doc = new docxLib.Document({
      sections: [{
        children: [
          new docxLib.Paragraph({ text: "" })
        ],
      }],
    });
    
    const blob = await docxLib.Packer.toBlob(doc);

    const newFileHandle = await selectedSubjectHandle.getFileHandle(fileName, { create: true });
    const writable = await newFileHandle.createWritable();
    await writable.write(blob);
    await writable.close();

    const newLi = addWeekUIElement(newFileHandle);
    await loadWeek(newFileHandle, newLi);
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

// --- Export & Save Action ---
btnSave.addEventListener('click', async (e) => {
  e.preventDefault();
  if (!selectedWeekHandle) return;

  const format = document.getElementById('export-format').value;
  const title = pageTitle.value || "Untitled Week";
  const htmlBody = pageBody.innerHTML;

  try {
    if (format === 'docx') {
      await saveCurrentWeekToFile();
      alert('Page saved to binary .docx file successfully!');
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
