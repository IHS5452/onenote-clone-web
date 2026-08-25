const docx = window.docx;

let rootHandle = null;
let selectedSubjectHandle = null;
let selectedWeekHandle = null;
let recentsMap = new Map();
let autoSaveTimer = null;

// --- DOM Elements ---
const btnOpenFolder = document.getElementById('btn-open-folder');
const btnAddSubject = document.getElementById('btn-add-subject');
const btnAddWeek = document.getElementById('btn-add-week');
const selectRecents = document.getElementById('select-recents');
const btnSave = document.getElementById('btn-save');

const btnImportMd = document.getElementById('btn-import-md');
const inputImportMd = document.getElementById('input-import-md');

// Custom Import Modal Elements
const importModal = document.getElementById('import-modal');
const btnImportUseCurrent = document.getElementById('btn-import-use-current');
const btnImportUseNew = document.getElementById('btn-import-use-new');
const inputImportNewSubject = document.getElementById('input-import-new-subject');
const btnImportCancel = document.getElementById('btn-import-cancel');

const listSubjects = document.getElementById('list-subjects');
const listWeeks = document.getElementById('list-weeks');
const emptyState = document.getElementById('empty-state');
const editorContent = document.getElementById('editor-content');
const toolbar = document.getElementById('toolbar');
const pageTitle = document.getElementById('page-title');
const pageDatetime = document.getElementById('page-datetime');
const pageBody = document.getElementById('page-body');

// Attach paste listener to editor
pageBody.addEventListener('paste', handlePaste);

// --- Natural Alphanumeric Sorting Helper ---
function alphaNumericCompare(a, b) {
  return a.name.localeCompare(b.name, undefined, { numeric: true, sensitivity: 'base' });
}

// --- Helper: Insert HTML into Document Selection Caret ---
function insertHtmlAtCaret(html) {
  const selection = window.getSelection();
  if (!selection.rangeCount) return;

  const range = selection.getRangeAt(0);
  range.deleteContents();

  const tempDiv = document.createElement('div');
  tempDiv.innerHTML = html;

  const frag = document.createDocumentFragment();
  let node, lastNode;
  while ((node = tempDiv.firstChild)) {
    lastNode = frag.appendChild(node);
  }

  range.insertNode(frag);

  if (lastNode) {
    range.setStartAfter(lastNode);
    range.collapse(true);
    selection.removeAllRanges();
    selection.addRange(range);
  }
}

// --- Helper: Check if Raw Text is Markdown ---
function isMarkdownText(text) {
  const mdRegex = /(^#{1,6}\s+|^\s*[\*\-\+]\s+|^\s*\d+\.\s+|\*\*.*?\*\*|\*.*?\*|\[.*?\]\(.*?\)|`{1,3}.*?`{1,3})/m;
  return mdRegex.test(text);
}

// --- Auto-Formatting Paste Handler (Images, DOCX, HTML, MD) ---
async function handlePaste(e) {
  const clipboardData = e.clipboardData || window.clipboardData;
  if (!clipboardData) return;

  // 1. Process File Pastes (Images & DOCX files)
  const items = clipboardData.items;
  for (let i = 0; i < items.length; i++) {
    const item = items[i];

    // Image Paste
    if (item.type.indexOf('image') !== -1) {
      e.preventDefault();
      const file = item.getAsFile();
      const reader = new FileReader();

      reader.onload = function (event) {
        insertHtmlAtCaret(`<img src="${event.target.result}" style="max-width:100%;" />`);
        saveCurrentWeekToFile();
      };
      reader.readAsDataURL(file);
      return;
    }

    // DOCX File Paste via Mammoth
    if (item.kind === 'file' && (item.type.includes('wordprocessingml') || item.type.includes('docx'))) {
      e.preventDefault();
      const file = item.getAsFile();
      const arrayBuffer = await file.arrayBuffer();

      const options = {
        convertImage: mammoth.images.imgElement(function (image) {
          return image.read("base64").then(function (imageBuffer) {
            return { src: "data:" + image.contentType + ";base64," + imageBuffer };
          });
        })
      };

      const result = await mammoth.convertToHtml({ arrayBuffer: arrayBuffer }, options);
      if (result.value) {
        insertHtmlAtCaret(result.value);
        saveCurrentWeekToFile();
      }
      return;
    }
  }

  // 2. Process Text Formats (HTML & Markdown)
  const htmlText = clipboardData.getData('text/html');
  const plainText = clipboardData.getData('text/plain');

  // HTML Content Paste
  if (htmlText && htmlText.trim()) {
    e.preventDefault();
    insertHtmlAtCaret(htmlText);
    saveCurrentWeekToFile();
    return;
  }

  // Markdown Content Paste
  if (plainText && isMarkdownText(plainText)) {
    e.preventDefault();
    if (window.marked) {
      const parsedHtml = window.marked.parse(plainText);
      insertHtmlAtCaret(parsedHtml);
      saveCurrentWeekToFile();
      return;
    }
  }
}

// --- Modal & File Picker Sequence Fix ---
function triggerFilePicker(targetSubjectName) {
  inputImportMd.value = '';
  inputImportMd.dataset.targetSubject = targetSubjectName;
  inputImportMd.click();
}

btnImportMd.addEventListener('click', () => {
  if (!rootHandle) {
    return alert("Please open or select a root subject folder first.");
  }

  if (selectedSubjectHandle) {
    if (importModal) {
      if (inputImportNewSubject) inputImportNewSubject.value = '';
      importModal.classList.add('active');
    } else {
      triggerFilePicker(selectedSubjectHandle.name);
    }
  } else {
    const subjectName = prompt("Enter the Subject name to import these files into:");
    if (subjectName && subjectName.trim()) {
      triggerFilePicker(subjectName.trim());
    }
  }
});

if (btnImportUseCurrent) {
  btnImportUseCurrent.addEventListener('click', () => {
    importModal.classList.remove('active');
    if (selectedSubjectHandle) {
      triggerFilePicker(selectedSubjectHandle.name);
    }
  });
}

if (btnImportUseNew) {
  btnImportUseNew.addEventListener('click', () => {
    const newName = inputImportNewSubject ? inputImportNewSubject.value.trim() : '';
    if (!newName) {
      alert("Please enter a subject name.");
      return;
    }
    importModal.classList.remove('active');
    triggerFilePicker(newName);
  });
}

if (btnImportCancel) {
  btnImportCancel.addEventListener('click', () => {
    importModal.classList.remove('active');
  });
}

// --- Bulk Markdown Import & Formatting Logic ---
inputImportMd.addEventListener('change', async (e) => {
  const files = Array.from(e.target.files);
  if (!files.length) return;

  try {
    const subjectName = e.target.dataset.targetSubject;
    if (!subjectName || !subjectName.trim()) return;

    const cleanSubjectName = subjectName.trim();

    const hasPermission = await verifyPermission(rootHandle, true);
    if (!hasPermission) return alert("Write permission was denied.");

    const targetSubjectHandle = await rootHandle.getDirectoryHandle(cleanSubjectName, { create: true });

    for (let i = 0; i < files.length; i++) {
      const file = files[i];
      const originalNameWithoutExt = file.name.replace(/\.md$/i, '');
      const targetFileName = `${originalNameWithoutExt}.docx`;

      const mdContent = await file.text();

      let htmlContent = '';
      if (window.marked) {
        htmlContent = window.marked.parse(mdContent);
      } else {
        htmlContent = `<p>${mdContent.replace(/\n/g, '<br>')}</p>`;
      }

      const docxBlob = await generateDocxBlobFromHtml(htmlContent);

      const newFileHandle = await targetSubjectHandle.getFileHandle(targetFileName, { create: true });
      const writable = await newFileHandle.createWritable();
      await writable.write(docxBlob);
      await writable.close();
    }

    await loadSubjects();

    const subjectItems = Array.from(listSubjects.querySelectorAll('li'));
    const matchedLi = subjectItems.find(li => li.handle && li.handle.name === cleanSubjectName);
    if (matchedLi) {
      await selectSubject(matchedLi.handle, matchedLi);
    }

    alert(`Successfully imported and formatted ${files.length} Markdown file(s) into "${cleanSubjectName}".`);

  } catch (err) {
    console.error("Bulk MD Import Error:", err);
    alert(`Failed to import files: ${err.message}`);
  } finally {
    e.target.value = '';
    delete e.target.dataset.targetSubject;
  }
});

async function generateDocxBlobFromHtml(htmlString) {
  const tempDiv = document.createElement('div');
  tempDiv.innerHTML = htmlString;

  const originalHtml = pageBody.innerHTML;
  pageBody.innerHTML = tempDiv.innerHTML;

  const blob = await generateDocxBlob();

  pageBody.innerHTML = originalHtml;
  return blob;
}

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

// --- Convert Data URL / Base64 to Uint8Array ---
function dataURLToUint8Array(dataURL) {
  const base64 = dataURL.split(',')[1];
  const binaryString = window.atob(base64);
  const len = binaryString.length;
  const bytes = new Uint8Array(len);
  for (let i = 0; i < len; i++) {
    bytes[i] = binaryString.charCodeAt(i);
  }
  return bytes;
}

// --- Recursive Parser for Elements ---
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
      } else if (tag === 'CODE') {
        runs.push(new docxLib.TextRun({
          text: child.textContent,
          font: "Courier New",
        }));
      } else if (tag === 'IMG') {
        const src = child.src;
        if (src.startsWith('data:image')) {
          try {
            const imageBytes = dataURLToUint8Array(src);
            runs.push(new docxLib.ImageRun({
              data: imageBytes,
              transformation: { width: 300, height: 200 },
            }));
          } catch (e) {
            console.error("Failed to parse image run:", e);
          }
        } else if (src.startsWith('http')) {
          runs.push(new docxLib.TextRun({
            text: ` [Image URL: ${src}] `,
            bold: true,
          }));
        }
      } else {
        runs.push(...parseElementToDocxRuns(child));
      }
    }
  });

  return runs;
}

// --- Build OpenXML Document ---
async function generateDocxBlob() {
  const docxLib = window.docx || docx;
  const paragraphs = [];

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
        } else if (tag === 'BLOCKQUOTE') {
          paragraphs.push(new docxLib.Paragraph({
            children: parseElementToDocxRuns(node),
            indent: { left: 720 },
          }));
        } else if (tag === 'PRE') {
          paragraphs.push(new docxLib.Paragraph({
            children: parseElementToDocxRuns(node),
          }));
        } else if (tag === 'IMG') {
          const src = node.src;
          if (src.startsWith('data:image')) {
            try {
              const imageBytes = dataURLToUint8Array(src);
              paragraphs.push(new docxLib.Paragraph({
                children: [
                  new docxLib.ImageRun({
                    data: imageBytes,
                    transformation: { width: 400, height: 250 },
                  })
                ]
              }));
            } catch (err) {
              console.error("Error creating image paragraph:", err);
            }
          } else if (src.startsWith('http')) {
            paragraphs.push(new docxLib.Paragraph({
              children: [
                new docxLib.TextRun({
                  text: `[Image: ${src}]`,
                  italic: true
                })
              ]
            }));
          }
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
        } else if (tag === 'H3') {
          paragraphs.push(new docxLib.Paragraph({
            children: parseElementToDocxRuns(node),
            heading: docxLib.HeadingLevel.HEADING_3,
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

  return await docxLib.Packer.toBlob(doc);
}

// --- Local Storage Binary Saver ---
async function saveCurrentWeekToFile() {
  if (!selectedWeekHandle) return;
  try {
    const blob = await generateDocxBlob();
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

  const entries = [];
  for await (const entry of rootHandle.values()) {
    if (entry.kind === 'directory') {
      entries.push(entry);
    }
  }

  entries.sort(alphaNumericCompare);

  entries.forEach(entry => {
    addSubjectUIElement(entry);
  });
}

function addSubjectUIElement(entryHandle) {
  const li = document.createElement('li');
  li.handle = entryHandle;
  
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

  const entries = [];
  for await (const entry of handle.values()) {
    if (entry.kind === 'file' && entry.name.endsWith('.docx')) {
      entries.push(entry);
    }
  }

  entries.sort(alphaNumericCompare);

  entries.forEach(entry => {
    addWeekUIElement(entry);
  });
}

// --- Week List Management ---
function addWeekUIElement(fileHandle) {
  const li = document.createElement('li');
  li.handle = fileHandle;
  
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
    
    const options = {
      convertImage: mammoth.images.imgElement(function(image) {
        return image.read("base64").then(function(imageBuffer) {
          return {
            src: "data:" + image.contentType + ";base64," + imageBuffer
          };
        });
      })
    };

    const result = await mammoth.convertToHtml({ arrayBuffer: arrayBuffer }, options);
    
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

// --- Rename Operations ---
async function renameSelectedSubject() {
  if (!selectedSubjectHandle || !rootHandle) {
    return alert("Please select a subject to rename.");
  }

  const oldName = selectedSubjectHandle.name;
  const newName = prompt("Rename Subject:", oldName);
  if (!newName || !newName.trim() || newName.trim() === oldName) return;

  const cleanName = newName.trim();

  try {
    await saveCurrentWeekToFile();

    const newDirHandle = await rootHandle.getDirectoryHandle(cleanName, { create: true });

    for await (const entry of selectedSubjectHandle.values()) {
      if (entry.kind === 'file') {
        const oldFile = await entry.getFile();
        const newFileHandle = await newDirHandle.getFileHandle(entry.name, { create: true });
        const writable = await newFileHandle.createWritable();
        await writable.write(await oldFile.arrayBuffer());
        await writable.close();
      }
    }

    await rootHandle.removeEntry(oldName, { recursive: true });
    await loadSubjects();

    const subjectItems = Array.from(listSubjects.querySelectorAll('li'));
    const matchedLi = subjectItems.find(li => li.handle && li.handle.name === cleanName);
    if (matchedLi) {
      await selectSubject(matchedLi.handle, matchedLi);
    }
  } catch (err) {
    console.error("Failed to rename subject:", err);
    alert(`Could not rename subject folder: ${err.message}`);
  }
}

async function renameSelectedWeek() {
  if (!selectedWeekHandle || !selectedSubjectHandle) {
    return alert("Please select a week to rename.");
  }

  const oldFileName = selectedWeekHandle.name;
  const oldDisplayName = oldFileName.replace('.docx', '');
  const newName = prompt("Rename Week:", oldDisplayName);
  if (!newName || !newName.trim() || newName.trim() === oldDisplayName) return;

  const cleanName = newName.trim();
  const newFileName = `${cleanName}.docx`;

  try {
    await saveCurrentWeekToFile();

    const currentFile = await selectedWeekHandle.getFile();
    const newFileHandle = await selectedSubjectHandle.getFileHandle(newFileName, { create: true });
    
    const writable = await newFileHandle.createWritable();
    await writable.write(await currentFile.arrayBuffer());
    await writable.close();

    await selectedSubjectHandle.removeEntry(oldFileName);

    const activeSubjectLi = listSubjects.querySelector('li.active');
    await selectSubject(selectedSubjectHandle, activeSubjectLi);

    const weekItems = Array.from(listWeeks.querySelectorAll('li'));
    const matchedLi = weekItems.find(li => li.handle && li.handle.name === newFileName);
    if (matchedLi) {
      await loadWeek(matchedLi.handle, matchedLi);
    }
  } catch (err) {
    console.error("Failed to rename week:", err);
    alert(`Could not rename week file: ${err.message}`);
  }
}

// --- Auto-Save Event Controls ---
pageBody.addEventListener('blur', saveCurrentWeekToFile);

pageTitle.addEventListener('blur', async () => {
  await saveCurrentWeekToFile();
  if (selectedWeekHandle && pageTitle.value.trim()) {
    const currentName = selectedWeekHandle.name.replace('.docx', '');
    if (pageTitle.value.trim() !== currentName) {
      await renameSelectedWeek();
    }
  }
});

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
    await rootHandle.getDirectoryHandle(cleanName, { create: true });

    await loadSubjects();
    
    const subjectItems = Array.from(listSubjects.querySelectorAll('li'));
    const matchedLi = subjectItems.find(li => li.handle && li.handle.name === cleanName);
    if (matchedLi) {
      await selectSubject(matchedLi.handle, matchedLi);
    }
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

    const activeSubjectLi = listSubjects.querySelector('li.active');
    await selectSubject(selectedSubjectHandle, activeSubjectLi);

    const weekItems = Array.from(listWeeks.querySelectorAll('li'));
    const matchedLi = weekItems.find(li => li.handle && li.handle.name === fileName);
    if (matchedLi) {
      await loadWeek(matchedLi.handle, matchedLi);
    }
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

// --- Universal Download Helper ---
function downloadFile(filename, content, mimeType) {
  const blob = content instanceof Blob ? content : new Blob([content], { type: mimeType });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  a.click();
  URL.revokeObjectURL(a.href);
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
      const blob = await generateDocxBlob();
      downloadFile(`${title}.docx`, blob, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    } else if (format === 'pdf') {
      const opt = { 
        margin: 15, 
        filename: `${title}.pdf`, 
        html2canvas: { scale: 2 } 
      };
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
    alert('Export error encountered.');
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
      else if (tag === 'img') output += `![Image](${node.src})\n\n`;
      else if (tag === 'li') output += `* ${htmlToMarkdown(node)}\n`;
      else output += htmlToMarkdown(node);
    }
  }
  return output;
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

// --- Global Keyboard Shortcuts ---
window.addEventListener('keydown', async (e) => {
  const isMac = navigator.platform.toUpperCase().indexOf('MAC') >= 0;
  const modifier = isMac ? e.metaKey : e.ctrlKey;

  if (modifier && e.key === '1') {
    e.preventDefault();
    await renameSelectedSubject();
  }

  if (modifier && e.key === '2') {
    e.preventDefault();
    await renameSelectedWeek();
  }

  if (modifier && e.key.toLowerCase() === 's') {
    e.preventDefault();
    if (!selectedWeekHandle) return;
    
    await saveCurrentWeekToFile();
    
    const activeWeek = listWeeks.querySelector('li.active');
    if (activeWeek) {
      const originalText = activeWeek.querySelector('span').textContent;
      activeWeek.querySelector('span').textContent = 'Saved!';
      setTimeout(() => {
        activeWeek.querySelector('span').textContent = originalText;
      }, 1000);
    }
  }

  if (modifier && e.key.toLowerCase() === 'e') {
    e.preventDefault();
    if (!selectedWeekHandle) return;
    
    btnSave.click();
  }
});

// --- System Theme Synchronization ---
function applySystemTheme(e) {
  const isDark = e.matches;
  document.documentElement.setAttribute('data-theme', isDark ? 'dark' : 'light');
}

const colorSchemeQuery = window.matchMedia('(prefers-color-scheme: dark)');
applySystemTheme(colorSchemeQuery);
colorSchemeQuery.addEventListener('change', applySystemTheme);

// --- About Modal Actions ---
const btnAbout = document.getElementById('btn-about');
const aboutModal = document.getElementById('about-modal');
const modalClose = document.getElementById('modal-close');
const btnCopyEmail = document.getElementById('btn-copy-email');
const devEmail = document.getElementById('dev-email').textContent;

btnAbout.addEventListener('click', () => {
  aboutModal.classList.add('active');
});

modalClose.addEventListener('click', () => {
  aboutModal.classList.remove('active');
});

aboutModal.addEventListener('click', (e) => {
  if (e.target === aboutModal) {
    aboutModal.classList.remove('active');
  }
});

btnCopyEmail.addEventListener('click', async () => {
  try {
    await navigator.clipboard.writeText(devEmail);
    btnCopyEmail.textContent = 'Copied!';
    setTimeout(() => {
      btnCopyEmail.textContent = 'Copy';
    }, 2000);
  } catch (err) {
    console.error('Failed to copy email address: ', err);
  }
});
