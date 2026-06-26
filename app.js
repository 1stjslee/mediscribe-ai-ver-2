const userForm = document.getElementById('user-form');
const userName = document.getElementById('user-name');
const userIdInput = document.getElementById('user-id');
const generateIdButton = document.getElementById('generate-id');
const userMessage = document.getElementById('user-message');
const userSection = document.getElementById('user-section');
const appSection = document.getElementById('app-section');
const currentUserName = document.getElementById('current-user-name');
const currentUserId = document.getElementById('current-user-id');
const recordButton = document.getElementById('record-button');
const stopButton = document.getElementById('stop-button');
const confirmButton = document.getElementById('confirm-button');
const cancelButton = document.getElementById('cancel-button');
const signoutButton = document.getElementById('signout-button');
const syncButton = document.getElementById('sync-button');
const exportUsersButton = document.getElementById('export-users-button');
const exportNotesButton = document.getElementById('export-notes-button');
const speechStatus = document.getElementById('speech-status');
const summaryOutput = document.getElementById('summary-output');
const pendingTextArea = document.getElementById('pending-textarea');
const pendingCategories = document.getElementById('pending-categories');
const historyList = document.getElementById('history-list');
const syncStatus = document.getElementById('sync-status');
const categoryFamily = document.getElementById('category-family');
const categoryAllergies = document.getElementById('category-allergies');
const categoryTimeframe = document.getElementById('category-timeframe');
const categoryComplaints = document.getElementById('category-complaints');
const categoryOther = document.getElementById('category-other');

let pendingNote = null;
let lastSavedNote = null;
let recordTimeoutId = null;
const MAX_RECORDING_MS = 30000;

const categories = {
  family: { label: 'Family History', keywords: ['family', 'mother', 'father', 'brother', 'sister', 'hereditary', 'genetic', 'grandparent', 'uncle', 'aunt'] },
  allergies: { label: 'Allergies', keywords: ['allergy', 'allergies', 'allergic', 'sensitivity', 'penicillin', 'latex', 'peanut', 'nuts', 'dairy'] },
  timeframe: { label: 'Timeframe', keywords: ['since', 'for', 'days', 'weeks', 'months', 'years', 'today', 'yesterday', 'recent', 'last'] },
  complaints: { label: 'Main complaints', keywords: ['pain', 'ache', 'cough', 'fever', 'nausea', 'headache', 'dizzy', 'symptom', 'complaint', 'swelling', 'pressure'] },
};

let db = null;
let currentUser = null;
let recognition = null;
let isRecording = false;

function openDatabase() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open('MediScribeDB', 1);

    request.onupgradeneeded = () => {
      db = request.result;
      if (!db.objectStoreNames.contains('users')) {
        db.createObjectStore('users', { keyPath: 'userId' });
      }
      if (!db.objectStoreNames.contains('notes')) {
        const notesStore = db.createObjectStore('notes', { keyPath: 'id', autoIncrement: true });
        notesStore.createIndex('userId', 'userId', { unique: false });
        notesStore.createIndex('timestamp', 'timestamp', { unique: false });
      }
    };

    request.onsuccess = () => {
      db = request.result;
      resolve(db);
    };

    request.onerror = () => reject(request.error);
  });
}

function getObjectStore(storeName, mode = 'readonly') {
  return db.transaction(storeName, mode).objectStore(storeName);
}

function saveUser(user) {
  return new Promise((resolve, reject) => {
    const store = getObjectStore('users', 'readwrite');
    const request = store.put(user);
    request.onsuccess = () => resolve(user);
    request.onerror = () => reject(request.error);
  });
}

function loadUser(userId) {
  return new Promise((resolve, reject) => {
    const request = getObjectStore('users').get(userId);
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

function saveNote(note) {
  return new Promise((resolve, reject) => {
    const store = getObjectStore('notes', 'readwrite');
    const request = store.add(note);
    request.onsuccess = () => resolve({ ...note, id: request.result });
    request.onerror = () => reject(request.error);
  });
}

function deleteNote(noteId) {
  return new Promise((resolve, reject) => {
    const store = getObjectStore('notes', 'readwrite');
    const request = store.delete(noteId);
    request.onsuccess = () => resolve();
    request.onerror = () => reject(request.error);
  });
}

function loadNotesForUser(userId) {
  return new Promise((resolve, reject) => {
    const store = getObjectStore('notes');
    const index = store.index('userId');
    const notes = [];
    const request = index.openCursor(IDBKeyRange.only(userId), 'prev');

    request.onsuccess = event => {
      const cursor = event.target.result;
      if (cursor) {
        notes.push(cursor.value);
        cursor.continue();
      } else {
        resolve(notes);
      }
    };

    request.onerror = () => reject(request.error);
  });
}

function loadAllUsers() {
  return new Promise((resolve, reject) => {
    const store = getObjectStore('users');
    const users = [];
    const request = store.openCursor();

    request.onsuccess = event => {
      const cursor = event.target.result;
      if (cursor) {
        users.push(cursor.value);
        cursor.continue();
      } else {
        resolve(users);
      }
    };

    request.onerror = () => reject(request.error);
  });
}

function loadAllNotes() {
  return new Promise((resolve, reject) => {
    const store = getObjectStore('notes');
    const notes = [];
    const request = store.openCursor(null, 'prev');

    request.onsuccess = event => {
      const cursor = event.target.result;
      if (cursor) {
        notes.push(cursor.value);
        cursor.continue();
      } else {
        resolve(notes);
      }
    };

    request.onerror = () => reject(request.error);
  });
}

function exportToXlsx(rows, fileName, sheetName) {
  if (!window.XLSX) {
    speechStatus.textContent = 'Export library is not loaded. Please check your internet connection.';
    return;
  }

  const worksheet = XLSX.utils.json_to_sheet(rows);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
  XLSX.writeFile(workbook, fileName);
}

function exportUsers() {
  loadAllUsers().then(users => {
    if (users.length === 0) {
      speechStatus.textContent = 'No users available to export.';
      return;
    }

    const exportRows = users.map(user => ({
      userId: user.userId,
      name: user.name,
      createdAt: new Date(user.createdAt).toLocaleString(),
    }));

    exportToXlsx(exportRows, 'userid.xlsx', 'Users');
    speechStatus.textContent = 'Users exported to userid.xlsx.';
  }).catch(error => {
    speechStatus.textContent = `Unable to export users: ${error}`;
  });
}

function exportConsultations() {
  loadAllNotes().then(notes => {
    if (notes.length === 0) {
      speechStatus.textContent = 'No consultation notes available to export.';
      return;
    }

    const exportRows = notes.map(note => ({
      id: note.id,
      userId: note.userId,
      text: note.text,
      timestamp: new Date(note.timestamp).toLocaleString(),
      categories: Array.isArray(note.categories) ? note.categories.join('; ') : note.categories,
      familyHistory: note.assigned?.familyHistory?.join(' | ') || '',
      allergies: note.assigned?.allergies?.join(' | ') || '',
      timeframe: note.assigned?.timeframe?.join(' | ') || '',
      mainComplaints: note.assigned?.mainComplaints?.join(' | ') || '',
      otherKeynotes: note.assigned?.otherKeynotes?.join(' | ') || '',
    }));

    exportToXlsx(exportRows, 'userconsultation.xlsx', 'Consultations');
    speechStatus.textContent = 'Consultation notes exported to userconsultation.xlsx.';
  }).catch(error => {
    speechStatus.textContent = `Unable to export consultations: ${error}`;
  });
}

function generateUserId() {
  return `user-${Math.random().toString(36).slice(2, 10)}-${Date.now().toString().slice(-4)}`;
}

function resetPendingState() {
  pendingNote = null;
  confirmButton.disabled = true;
  cancelButton.disabled = true;
  pendingTextArea.classList.add('hidden');
  pendingCategories.classList.add('hidden');
  pendingTextArea.value = '';
}

function clearSummary() {
  summaryOutput.innerHTML = '<p>No transcription yet.</p>';
  resetPendingState();
}

function setCurrentUser(user) {
  currentUser = user;
  window.localStorage.setItem('mediScribeCurrentUser', JSON.stringify(user));
  currentUserName.textContent = user.name;
  currentUserId.textContent = user.userId;
  userSection.classList.add('hidden');
  appSection.classList.remove('hidden');
  resetPendingState();
  renderHistory();
  updateSyncStatus();
}

function splitTranscriptSegments(text) {
  return text
    .split(/[\r\n,.?!;]+/g)
    .map(segment => segment.trim())
    .filter(Boolean);
}

function matchesKeyword(text, keyword) {
  const escaped = keyword.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const regex = new RegExp(`\\b${escaped}\\b`, 'i');
  return regex.test(text);
}

function categorizeSpeech(text) {
  const segments = splitTranscriptSegments(text);
  const assigned = {
    familyHistory: [],
    allergies: [],
    timeframe: [],
    mainComplaints: [],
    otherKeynotes: [],
  };

  segments.forEach(segment => {
    const lower = segment.toLowerCase();
    let matched = false;

    if (categories.family.keywords.some(keyword => matchesKeyword(lower, keyword))) {
      assigned.familyHistory.push(segment);
      matched = true;
    }
    if (categories.allergies.keywords.some(keyword => matchesKeyword(lower, keyword))) {
      assigned.allergies.push(segment);
      matched = true;
    }
    if (categories.timeframe.keywords.some(keyword => matchesKeyword(lower, keyword))) {
      assigned.timeframe.push(segment);
      matched = true;
    }
    if (categories.complaints.keywords.some(keyword => matchesKeyword(lower, keyword))) {
      assigned.mainComplaints.push(segment);
      matched = true;
    }

    if (!matched) {
      assigned.otherKeynotes.push(segment);
    }
  });

  return assigned;
}

function renderSummary(note) {
  const formattedTime = new Date(note.timestamp).toLocaleString();
  summaryOutput.innerHTML = `
    <p><strong>${note.text}</strong></p>
    <p><small>${formattedTime}</small></p>
    <div class="note-tags">${note.categories.map(tag => `<span class="note-tag">${tag}</span>`).join('')}</div>
  `;
}

function renderHistory() {
  if (!currentUser) {
    historyList.innerHTML = '<p class="info">Login to view notes.</p>';
    return;
  }

  loadNotesForUser(currentUser.userId).then(notes => {
    categoryFamily.innerHTML = '';
    categoryAllergies.innerHTML = '';
    categoryTimeframe.innerHTML = '';
    categoryComplaints.innerHTML = '';
    categoryOther.innerHTML = '';

    if (notes.length === 0) {
      historyList.innerHTML = '<p class="info">No notes stored yet.</p>';
      return;
    }

    historyList.innerHTML = notes.map(note => {
      const noteTags = note.categories.map(tag => `<span class="note-tag">${tag}</span>`).join('');
      return `
        <article class="history-item" data-note-id="${note.id}">
          <div class="history-header">
            <time>${new Date(note.timestamp).toLocaleString()}</time>
            <button class="delete-button" data-note-id="${note.id}" type="button">Delete</button>
          </div>
          <p>${note.text}</p>
          <div class="note-tags">${noteTags}</div>
        </article>
      `;
    }).join('');

    historyList.querySelectorAll('.delete-button').forEach(button => {
      button.addEventListener('click', async event => {
        const noteId = Number(event.currentTarget.dataset.noteId);
        if (!Number.isFinite(noteId)) return;
        await deleteNote(noteId);
        if (lastSavedNote && lastSavedNote.id === noteId) {
          lastSavedNote = null;
          clearSummary();
        }
        renderHistory();
        speechStatus.textContent = 'Transcript deleted successfully.';
      });
    });

    notes.forEach(note => {
      if (note.assigned.familyHistory.length) {
        note.assigned.familyHistory.forEach(item => categoryFamily.insertAdjacentHTML('beforeend', `<p>${item}</p>`));
      }
      if (note.assigned.allergies.length) {
        note.assigned.allergies.forEach(item => categoryAllergies.insertAdjacentHTML('beforeend', `<p>${item}</p>`));
      }
      if (note.assigned.timeframe.length) {
        note.assigned.timeframe.forEach(item => categoryTimeframe.insertAdjacentHTML('beforeend', `<p>${item}</p>`));
      }
      if (note.assigned.mainComplaints.length) {
        note.assigned.mainComplaints.forEach(item => categoryComplaints.insertAdjacentHTML('beforeend', `<p>${item}</p>`));
      }
      if (note.assigned.otherKeynotes.length) {
        note.assigned.otherKeynotes.forEach(item => categoryOther.insertAdjacentHTML('beforeend', `<p>${item}</p>`));
      }
    });
  }).catch(error => {
    historyList.innerHTML = `<p class="info">Unable to load stored notes. ${error}</p>`;
  });
}

function updateSyncStatus() {
  const online = navigator.onLine;
  syncStatus.textContent = online ? 'Online sync ready' : 'Offline local only';
  syncStatus.style.background = online ? '#dcfce7' : '#e2e8f0';
  syncStatus.style.color = online ? '#15803d' : '#0f172a';
}

function setupRecognition() {
  const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
  if (!SpeechRecognition) {
    speechStatus.textContent = 'Speech recognition is not supported in this browser. Use Chrome or Edge.';
    recordButton.disabled = true;
    return;
  }

  recognition = new SpeechRecognition();
  recognition.lang = 'en-US';
  recognition.interimResults = false;
  recognition.continuous = true;

  recognition.onstart = () => {
    isRecording = true;
    speechStatus.textContent = 'Recording... speak clearly into your microphone.';
    recordButton.disabled = true;
    stopButton.disabled = false;
    confirmButton.disabled = true;
    cancelButton.disabled = true;
    if (recordTimeoutId) {
      clearTimeout(recordTimeoutId);
    }
    recordTimeoutId = window.setTimeout(() => {
      if (isRecording) {
        stopRecording();
      }
    }, MAX_RECORDING_MS);
  };

  recognition.onerror = event => {
    speechStatus.textContent = `Recording error: ${event.error}`;
    isRecording = false;
    recordButton.disabled = false;
    stopButton.disabled = true;
  };

  recognition.onend = () => {
    if (recordTimeoutId) {
      clearTimeout(recordTimeoutId);
      recordTimeoutId = null;
    }

    if (isRecording) {
      recognition.start();
      return;
    }

    speechStatus.textContent = pendingNote ? 'Recording stopped. Review transcript before saving.' : 'Recording ended.';
    recordButton.disabled = false;
    stopButton.disabled = true;
  };

  recognition.onresult = event => {
    const transcript = Array.from(event.results)
      .map(result => result[0].transcript)
      .join(' ');

    if (!transcript.trim()) {
      speechStatus.textContent = 'No speech detected. Please continue speaking or stop recording.';
      return;
    }

    const text = pendingNote ? `${pendingNote.text} ${transcript}`.trim() : transcript.trim();
    const assigned = categorizeSpeech(text);
    pendingNote = {
      userId: currentUser.userId,
      text,
      timestamp: pendingNote ? pendingNote.timestamp : Date.now(),
      categories: [...new Set([].concat(
        assigned.familyHistory.length ? categories.family.label : [],
        assigned.allergies.length ? categories.allergies.label : [],
        assigned.timeframe.length ? categories.timeframe.label : [],
        assigned.mainComplaints.length ? categories.complaints.label : [],
        assigned.otherKeynotes.length ? 'Other keynotes' : []
      ))],
      assigned,
    };

    renderPendingState();
    speechStatus.textContent = 'Speech recognized. Review and confirm save.';
    confirmButton.disabled = false;
    cancelButton.disabled = false;
  };
}

function startRecording() {
  if (!recognition) {
    return;
  }
  try {
    recognition.start();
  } catch (err) {
    speechStatus.textContent = 'Unable to start recording at this time.';
  }
}

function stopRecording() {
  if (recognition && isRecording) {
    isRecording = false;
    recognition.stop();
    stopButton.disabled = true;
    recordButton.disabled = false;
    speechStatus.textContent = 'Stopping recording...';
  }
}

function renderPendingState() {
  if (!pendingNote) {
    return;
  }
  summaryOutput.innerHTML = `
    <p><strong>${pendingNote.text}</strong></p>
    <p><small>${new Date(pendingNote.timestamp).toLocaleString()}</small></p>
    <div class="note-tags">${pendingNote.categories.map(tag => `<span class="note-tag">${tag}</span>`).join('')}</div>
  `;
  pendingTextArea.classList.remove('hidden');
  pendingCategories.classList.remove('hidden');
  pendingTextArea.value = pendingNote.text;
  updatePendingCategoriesPreview(pendingNote.text);
}

function updatePendingCategoriesPreview(text) {
  const assigned = categorizeSpeech(text);
  const lines = [];
  if (assigned.familyHistory.length) {
    lines.push(`<strong>Family History:</strong> ${assigned.familyHistory.join(' | ')}`);
  }
  if (assigned.allergies.length) {
    lines.push(`<strong>Allergies:</strong> ${assigned.allergies.join(' | ')}`);
  }
  if (assigned.timeframe.length) {
    lines.push(`<strong>Timeframe:</strong> ${assigned.timeframe.join(' | ')}`);
  }
  if (assigned.mainComplaints.length) {
    lines.push(`<strong>Main Complaints:</strong> ${assigned.mainComplaints.join(' | ')}`);
  }
  if (assigned.otherKeynotes.length) {
    lines.push(`<strong>Other Keynotes:</strong> ${assigned.otherKeynotes.join(' | ')}`);
  }
  pendingCategories.innerHTML = lines.join('<br />');
}

function savePendingNote() {
  if (!pendingNote) {
    return;
  }
  const editedText = pendingTextArea.value.trim();
  if (!editedText) {
    speechStatus.textContent = 'Cannot save empty transcription. Please edit or cancel.';
    return;
  }
  const assigned = categorizeSpeech(editedText);
  const noteToSave = {
    userId: currentUser.userId,
    text: editedText,
    timestamp: pendingNote.timestamp,
    categories: [...new Set([].concat(
      assigned.familyHistory.length ? categories.family.label : [],
      assigned.allergies.length ? categories.allergies.label : [],
      assigned.timeframe.length ? categories.timeframe.label : [],
      assigned.mainComplaints.length ? categories.complaints.label : [],
      assigned.otherKeynotes.length ? 'Other keynotes' : []
    ))],
    assigned,
  };
  saveNote(noteToSave).then(saved => {
    lastSavedNote = saved;
    speechStatus.textContent = 'Transcript saved locally.';
    renderSummary(saved);
    renderHistory();
    resetPendingState();
  }).catch(error => {
    speechStatus.textContent = `Unable to save note: ${error}`;
  });
}

function cancelPendingNote() {
  resetPendingState();
  if (lastSavedNote) {
    renderSummary(lastSavedNote);
  } else {
    clearSummary();
  }
  speechStatus.textContent = 'Pending transcription canceled.';
}

function signOut() {
  currentUser = null;
  window.localStorage.removeItem('mediScribeCurrentUser');
  userSection.classList.remove('hidden');
  appSection.classList.add('hidden');
  resetPendingState();
  clearSummary();
  historyList.innerHTML = '';
  userMessage.textContent = 'Signed out successfully.';
}

userForm.addEventListener('submit', async event => {
  event.preventDefault();
  const nameValue = userName.value.trim();
  const userIdValue = userIdInput.value.trim();
  const mode = userForm.mode.value;

  if (!nameValue) {
    userMessage.textContent = 'Please enter a name to continue.';
    return;
  }
  if (!userIdValue) {
    userMessage.textContent = 'Please enter or generate a user ID.';
    return;
  }

  try {
    const existing = await loadUser(userIdValue);
    if (mode === 'existing') {
      if (!existing) {
        userMessage.textContent = 'User not found. Switch to new user mode to create an account.';
        return;
      }
      setCurrentUser(existing);
    } else {
      if (existing) {
        userMessage.textContent = 'User ID already exists. Please choose a different ID.';
        return;
      }
      const newUser = { userId: userIdValue, name: nameValue, createdAt: Date.now() };
      await saveUser(newUser);
      setCurrentUser(newUser);
    }
    userMessage.textContent = '';
  } catch (error) {
    userMessage.textContent = `Error processing user account: ${error}`;
  }
});

generateIdButton.addEventListener('click', () => {
  userIdInput.value = generateUserId();
});

recordButton.addEventListener('click', () => {
  if (!currentUser) {
    speechStatus.textContent = 'Please login before recording.';
    return;
  }
  startRecording();
});

stopButton.addEventListener('click', () => stopRecording());

confirmButton.addEventListener('click', () => savePendingNote());
cancelButton.addEventListener('click', () => cancelPendingNote());
signoutButton.addEventListener('click', () => signOut());
exportUsersButton.addEventListener('click', () => exportUsers());
exportNotesButton.addEventListener('click', () => exportConsultations());
pendingTextArea.addEventListener('input', event => updatePendingCategoriesPreview(event.target.value));

syncButton.addEventListener('click', () => syncDataToExcel());

async function getDirectoryHandleForSync() {
  if (window.showDirectoryPicker) {
    return window.showDirectoryPicker();
  }

  speechStatus.textContent = 'Folder syncing is not supported in this browser. Use Chrome/Edge with File System Access API.';
  return null;
}

async function writeWorkbookToHandle(handle, workbook) {
  const writable = await handle.createWritable();
  const workbookArray = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
  await writable.write(new Blob([workbookArray], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
  await writable.close();
}

async function writeWorkbookToDirectory(directoryHandle, fileName, workbook) {
  const fileHandle = await directoryHandle.getFileHandle(fileName, { create: true });
  await writeWorkbookToHandle(fileHandle, workbook);
}

async function syncDataToExcel() {
  if (!window.XLSX) {
    speechStatus.textContent = 'Sync failed: Excel library not loaded.';
    return;
  }

  const directoryHandle = await getDirectoryHandleForSync();
  if (!directoryHandle) {
    return;
  }

  try {
    const users = await loadAllUsers();
    const notes = await loadAllNotes();

    const userRows = users.map(user => ({
      userId: user.userId,
      name: user.name,
      createdAt: new Date(user.createdAt).toLocaleString(),
    }));

    const noteRows = notes.map(note => ({
      id: note.id,
      userId: note.userId,
      text: note.text,
      timestamp: new Date(note.timestamp).toLocaleString(),
      categories: Array.isArray(note.categories) ? note.categories.join('; ') : note.categories,
      familyHistory: note.assigned?.familyHistory?.join(' | ') || '',
      allergies: note.assigned?.allergies?.join(' | ') || '',
      timeframe: note.assigned?.timeframe?.join(' | ') || '',
      mainComplaints: note.assigned?.mainComplaints?.join(' | ') || '',
      otherKeynotes: note.assigned?.otherKeynotes?.join(' | ') || '',
    }));

    const userWorkbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(userWorkbook, XLSX.utils.json_to_sheet(userRows), 'Users');

    const noteWorkbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(noteWorkbook, XLSX.utils.json_to_sheet(noteRows), 'Consultations');

    await writeWorkbookToDirectory(directoryHandle, 'userid.xlsx', userWorkbook);
    await writeWorkbookToDirectory(directoryHandle, 'userconsultation.xlsx', noteWorkbook);

    speechStatus.textContent = 'Sync complete: files written to selected folder.';
  } catch (error) {
    speechStatus.textContent = `Sync failed: ${error}`;
  }
}

window.addEventListener('online', updateSyncStatus);
window.addEventListener('offline', updateSyncStatus);

window.addEventListener('load', async () => {
  await openDatabase();
  setupRecognition();

  const storedUser = window.localStorage.getItem('mediScribeCurrentUser');
  if (storedUser) {
    const parsed = JSON.parse(storedUser);
    const found = await loadUser(parsed.userId);
    if (found) {
      setCurrentUser(found);
      return;
    }
  }

  updateSyncStatus();
});
