
import ExcelJS from "https://cdn.jsdelivr.net/npm/exceljs@4.4.0/+esm";

import { initializeApp } from "https://www.gstatic.com/firebasejs/12.4.0/firebase-app.js";

import {
  getAuth,
  onAuthStateChanged,
  signInWithEmailAndPassword
} from "https://www.gstatic.com/firebasejs/12.4.0/firebase-auth.js";

import {
  getDatabase,
  ref,
  get,
  set,
  onValue
} from "https://www.gstatic.com/firebasejs/12.4.0/firebase-database.js";

const firebaseConfig = {
  apiKey: "AIzaSyBVa-WUVciypPDEEV3hSY5UZcxUHYeFf1k",
  authDomain: "bitacora-477d1.firebaseapp.com",
  databaseURL: "https://bitacora-477d1-default-rtdb.firebaseio.com",
  projectId: "bitacora-477d1",
  storageBucket: "bitacora-477d1.firebasestorage.app",
  messagingSenderId: "1081946837153",
  appId: "1:1081946837153:web:56968b8c834c0d8361a0e0"
};

const firebaseApp = initializeApp(firebaseConfig);
const auth = getAuth(firebaseApp);
const database = getDatabase(firebaseApp);

let recordsRef = null;
let unsubscribeRecords = null;

"use strict";

/* =========================================
   BITÁCORA — Lógica de la aplicación
   Archivo: js/app.js

   Los registros históricos se cargan al iniciar.
   Los cambios se guardan en localStorage.
   Más adelante se conectará con Firebase.
   ========================================= */

/* 1. Elementos de la interfaz */

const elements = {
  currentDate: document.getElementById("current-date"),
  liveClock: document.getElementById("live-clock"),
  clockPeriod: document.getElementById("clock-period"),

  attendanceState: document.getElementById("attendance-state"),
  attendanceMessage: document.getElementById("attendance-message"),
  attendanceDetails: document.getElementById("attendance-details"),

  entryTime: document.getElementById("entry-time"),
  exitTime: document.getElementById("exit-time"),

  

entryButton: document.getElementById("entry-button"),
exitButton: document.getElementById("exit-button"),
  attendanceHelper: document.getElementById("attendance-helper"),

  weeklyHours: document.getElementById("weekly-hours"),
  totalHours: document.getElementById("total-hours"),

  historyList: document.getElementById("history-list"),
  recordCount: document.getElementById("record-count"),
 
 historyToggle: document.getElementById("history-toggle"),
 loginScreen: document.getElementById("login-screen"),
loginForm: document.getElementById("login-form"),
loginEmail: document.getElementById("login-email"),
loginPassword: document.getElementById("login-password"),
loginButton: document.getElementById("login-button"),
loginError: document.getElementById("login-error"),
};


/* 2. Configuración */

const CONFIG = {
  locale: "es-MX",
  storageKey: "bitacora-records-v1"
};


/* 3. Estado de la aplicación */

// status: "attended", "absent" o "pending".
// entry y exit son marcas de tiempo en milisegundos.

const state = {
  records: loadRecords(),
  currentAction: "entry",
  busy: false,
  historyExpanded: false,
  currentUser: null
};


/* 4. Funciones para fechas y horas */

function pad(value) {
  return String(value).padStart(2, "0");
}

function toLocalDateKey(date) {
  return [
    date.getFullYear(),
    pad(date.getMonth() + 1),
    pad(date.getDate())
  ].join("-");
}

function dateFromKey(key) {
  const [year, month, day] = key.split("-").map(Number);

  return new Date(year, month - 1, day);
}

function formatDate(date, options = {}) {
  return new Intl.DateTimeFormat(
    CONFIG.locale,
    options
  ).format(date);
}

function formatTime(timestamp) {
  if (timestamp == null) {
    return "Pendiente";
  }

  return formatDate(new Date(timestamp), {
    hour: "2-digit",
    minute: "2-digit",
    hour12: true
  });
}

function formatClock(timestamp) {
  const date = new Date(timestamp);

  return {
    time: formatDate(date, {
      hour: "2-digit",
      minute: "2-digit",
      hour12: true
    }),

    period: formatDate(date, {
      hour: "2-digit",
      hour12: true
    }).replace(/[0-9\s:]/g, "").toUpperCase()
  };
}

function roundToNearestHalfHour(timestamp) {
  if (timestamp == null) {
    return null;
  }

  const date = new Date(timestamp);
  const minutesOfDay = date.getHours() * 60 + date.getMinutes();

  // Redondeo a la media hora más cercana.
  // Si cae exactamente a 15 minutos del punto medio,
  // se redondea hacia abajo.
  const roundedMinutes = Math.floor((minutesOfDay + 14) / 30) * 30;

  date.setHours(0, 0, 0, 0);
  date.setMinutes(roundedMinutes);

  return date.getTime();
}

function getDurationMinutes(record) {
  if (
    !record ||
    record.status === "absent" ||
    record.status === "pending" ||
    record.entry == null ||
    record.exit == null
  ) {
    return 0;
  }

  const roundedEntry = roundToNearestHalfHour(record.entry);
  const roundedExit = roundToNearestHalfHour(record.exit);
  const difference = roundedExit - roundedEntry;

  if (difference <= 0) {
    return 0;
  }

  return Math.floor(difference / 60000);
}

function formatDuration(minutes) {
  const hours = Math.floor(minutes / 60);
  const remainingMinutes = minutes % 60;

  if (remainingMinutes === 0) {
    return `${hours} h`;
  }

  if (hours === 0) {
    return `${remainingMinutes} min`;
  }

  return `${hours} h ${remainingMinutes} min`;
}

function getCurrentWeekStart(date = new Date()) {
  const start = new Date(
    date.getFullYear(),
    date.getMonth(),
    date.getDate()
  );

  const day = start.getDay();
  const daysSinceMonday = (day + 6) % 7;

  start.setDate(start.getDate() - daysSinceMonday);

  return start;
}


/* 5. Jornadas históricas y persistencia local */

function makeTimestamp(dateKey, time) {
  if (!time) {
    return null;
  }

  const [year, month, day] = dateKey.split("-").map(Number);
  const [hour, minute] = time.split(":").map(Number);

  return new Date(
    year,
    month - 1,
    day,
    hour,
    minute,
    0,
    0
  ).getTime();
}

function makeRecord(date, entry, exit, status = "attended") {
  return {
    id: `historical-${date}`,
    date,
    entry: makeTimestamp(date, entry),
    exit: makeTimestamp(date, exit),
    status,
    source: date >= "2026-10-20" ? "recent" : "imported",
    demo: false
  };
}

function createHistoricalRecords() {
  const records = [
    // Semana 1: 18–23 de agosto
    makeRecord("2026-08-18", "10:00", "15:00"),
    makeRecord("2026-08-19", "10:00", "15:00"),
    makeRecord("2026-08-20", null, null, "absent"),
    makeRecord("2026-08-21", "10:00", "15:00"),
    makeRecord("2026-08-22", "10:00", "16:00"),
    makeRecord("2026-08-23", null, null, "absent"),

    // Semana 2: 25–30 de agosto
    makeRecord("2026-08-25", "10:00", "14:30"),
    makeRecord("2026-08-26", "10:00", "14:00"),
    makeRecord("2026-08-27", "10:00", "15:00"),
    makeRecord("2026-08-28", "10:00", "15:00"),
    makeRecord("2026-08-29", "10:00", "16:00"),
    makeRecord("2026-08-30", "10:00", "13:00"),

    // Semana 3: 1–6 de septiembre
    makeRecord("2026-09-01", "10:30", "14:30"),
    makeRecord("2026-09-02", null, null, "absent"),
    makeRecord("2026-09-03", "10:00", "14:00"),
    makeRecord("2026-09-04", "10:00", "14:00"),
    makeRecord("2026-09-05", "10:00", "15:00"),
    makeRecord("2026-09-06", "10:00", "13:00"),

    // Semana 4: 8–13 de septiembre
    makeRecord("2026-09-08", "10:00", "13:30"),
    makeRecord("2026-09-09", "10:00", "14:00"),
    makeRecord("2026-09-10", null, null, "absent"),
    makeRecord("2026-09-11", "10:00", "14:00"),
    makeRecord("2026-09-12", "10:00", "13:00"),
    makeRecord("2026-09-13", null, null, "absent"),

    // Semana 5: 15–20 de septiembre
    makeRecord("2026-09-15", "10:00", "14:00"),
    makeRecord("2026-09-16", null, null, "absent"),
    makeRecord("2026-09-17", "10:00", "14:30"),
    makeRecord("2026-09-18", "11:00", "13:30"),
    makeRecord("2026-09-19", "10:00", "14:00"),
    makeRecord("2026-09-20", "10:00", "13:00"),

    // Semana 6: 22–27 de septiembre
    makeRecord("2026-09-22", "10:00", "14:00"),
    makeRecord("2026-09-23", "10:00", "14:00"),
    makeRecord("2026-09-24", "10:00", "14:00"),
    makeRecord("2026-09-25", "10:00", "14:00"),
    makeRecord("2026-09-26", null, null, "absent"),
    makeRecord("2026-09-27", null, null, "absent"),

    // Semana 7: 29 de septiembre–4 de octubre
    makeRecord("2026-09-29", "10:00", "15:30"),
    makeRecord("2026-09-30", "10:00", "14:00"),
    makeRecord("2026-10-01", "10:00", "12:30"),
    makeRecord("2026-10-02", "11:00", "14:00"),
    makeRecord("2026-10-03", "10:00", "15:00"),
    makeRecord("2026-10-04", "10:00", "14:00"),

    // Semana 8: 6–11 de octubre
    makeRecord("2026-10-06", "10:00", "14:00"),
    makeRecord("2026-10-07", null, null, "absent"),
    makeRecord("2026-10-08", "10:30", "15:30"),
    makeRecord("2026-10-09", "10:00", "14:00"),
    makeRecord("2026-10-10", "10:30", "15:30"),
    makeRecord("2026-10-11", "10:00", "13:30"),

    // Semana 9: 13–18 de octubre
    makeRecord("2026-10-13", null, null, "absent"),
    makeRecord("2026-10-14", "11:00", "14:30"),
    makeRecord("2026-10-15", "10:30", "13:30"),
    makeRecord("2026-10-16", "10:00", "12:00"),
    makeRecord("2026-10-17", "10:30", "14:00"),
    makeRecord("2026-10-18", null, null, "absent"),

    // Semana 10: 20–25 de octubre.
    // Los días vacíos quedan pendientes.
    makeRecord("2026-10-20", "10:00", "14:00"),
    makeRecord("2026-10-21", "10:00", "14:00"),
    makeRecord("2026-10-22", "10:00", "12:00"),
    makeRecord("2026-10-23", null, null, "pending"),
    makeRecord("2026-10-24", null, null, "pending"),
    makeRecord("2026-10-25", null, null, "pending")
  ];

  return records;
}


function loadRecords() {
  try {
    const saved = localStorage.getItem(CONFIG.storageKey);

    if (saved) {
      const parsed = JSON.parse(saved);

      if (Array.isArray(parsed) && parsed.length > 0) {
        const realRecords = parsed.filter(record => !record.demo);

        if (realRecords.length > 0) {
          // Corregir la jornada del jueves 22 de octubre.
          const thursday = realRecords.find(
            record => record.date === "2026-10-22"
          );

          if (thursday) {
            thursday.entry = makeTimestamp("2026-10-22", "10:00");
            thursday.exit = makeTimestamp("2026-10-22", "12:00");
            thursday.status = "attended";
            thursday.source = "recent";
          }

          localStorage.setItem(
            CONFIG.storageKey,
            JSON.stringify(realRecords)
          );

          return realRecords;
        }
      }
    }
  } catch (error) {
    console.warn("Error al cargar registros guardados:", error);
  }

  const historical = createHistoricalRecords();

  try {
    localStorage.setItem(
      CONFIG.storageKey,
      JSON.stringify(historical)
    );
  } catch (error) {
    console.warn("Error al guardar registros:", error);
  }

  return historical;
}


function saveRecords() {
  try {
    localStorage.setItem(
      CONFIG.storageKey,
      JSON.stringify(state.records)
    );
  } catch (error) {
    console.error("Error al guardar los registros:", error);
  }
}

async function syncRecordsToFirebase() {
  if (!recordsRef || !state.currentUser) return;

  try {
    await set(recordsRef, state.records);
  } catch (error) {
    console.error("Error al sincronizar con Firebase:", error);
    alert(
      "El registro quedó guardado en este navegador, pero no se pudo sincronizar. Revisa tu conexión."
    );
  }
}


/* 6. Buscar el registro de hoy */

function getTodayRecord() {
  const todayKey = toLocalDateKey(new Date());

  return state.records.find(
    record => record.date === todayKey
  ) || null;
}

function getTodayAction() {
  const todayRecord = getTodayRecord();

  if (!todayRecord || todayRecord.entry == null) {
    return "entry";
  }

  if (todayRecord.exit == null) {
    return "exit";
  }

  return "complete";
}


/* 7. Actualizar el reloj */

function updateClock() {
  const now = new Date();
  const clock = formatClock(now.getTime());

  elements.currentDate.textContent = formatDate(now, {
    weekday: "long",
    day: "numeric",
    month: "long",
    year: "numeric"
  });

  elements.liveClock.textContent = clock.time;
  elements.clockPeriod.textContent = clock.period;
}


/* 8. Registrar entrada y salida */



function handleEntryClick() {
  if (state.busy) return;

  const todayKey = toLocalDateKey(new Date());
  const timestamp = Date.now();

  state.busy = true;

  try {
    let record = getTodayRecord();

    if (record?.entry != null) {
      alert(
        "Ya existe una entrada para hoy. Si fue un error, edita el registro en el historial."
      );
      return;
    }

    if (!record) {
      record = {
        id: `local-${todayKey}`,
        date: todayKey,
        entry: timestamp,
        exit: null,
        status: "attended",
        source: "recent",
        demo: false
      };

      state.records.push(record);
    } else {
      record.entry = timestamp;
      record.exit = null;
      record.status = "attended";
      record.source = "recent";
    }

    saveRecords();
    void syncRecordsToFirebase();

  } finally {
    state.busy = false;
    renderApp();
  }
}

function handleExitClick() {
  if (state.busy) return;

  const timestamp = Date.now();
  state.busy = true;

  try {
    const record = getTodayRecord();

    if (!record || record.entry == null) {
      alert("Primero debes registrar tu entrada.");
      return;
    }

    if (record.exit != null) {
      alert(
        "Ya existe una salida para hoy. Si fue un error, edita el registro en el historial."
      );
      return;
    }

    if (timestamp <= record.entry) {
      alert("La salida debe ser posterior a la entrada.");
      return;
    }

    record.exit = timestamp;
    record.status = "attended";
    record.source = "recent";

    saveRecords();
    void syncRecordsToFirebase();

  } finally {
    state.busy = false;
    renderApp();
  }
}




/* 9. Mostrar el estado de asistencia */


function renderAttendance() {
  const record = getTodayRecord();
  const hasEntry = record?.entry != null;
  const hasExit = record?.exit != null;

  elements.attendanceState.classList.remove(
    "is-active",
    "is-complete"
  );

  elements.attendanceDetails.hidden = !record;

  if (!hasEntry) {
    elements.attendanceMessage.textContent =
      "Aún no has registrado tu entrada.";

    elements.attendanceHelper.textContent =
      "Registra tu entrada al comenzar tu jornada.";

    elements.entryTime.textContent = "--:--";
    elements.exitTime.textContent = "Pendiente";

  } else if (!hasExit) {
    elements.attendanceState.classList.add("is-active");

    elements.attendanceMessage.textContent =
      "Tu jornada está en curso.";

    elements.attendanceHelper.textContent =
      "Registra tu salida al terminar tu jornada.";

    elements.entryTime.textContent =
      formatTime(roundToNearestHalfHour(record.entry));

    elements.exitTime.textContent = "Pendiente";

  } else {
    elements.attendanceState.classList.add("is-complete");

    elements.attendanceMessage.textContent =
      "Tu jornada de hoy ha sido registrada.";

    elements.attendanceHelper.textContent =
      `Total de hoy: ${formatDuration(getDurationMinutes(record))}.`;

    elements.entryTime.textContent =
      formatTime(roundToNearestHalfHour(record.entry));

    elements.exitTime.textContent =
      formatTime(roundToNearestHalfHour(record.exit));
  }

  // Entrada y salida tienen controles independientes.
  elements.entryButton.disabled =
    state.busy || hasEntry;

  elements.exitButton.disabled =
    state.busy || !hasEntry || hasExit;
}



/* 10. Calcular totales */

function getTotalMinutes(records) {
  return records.reduce(
    (total, record) => total + getDurationMinutes(record),
    0
  );
}

function getWeeklyRecords() {
  const weekStart = getCurrentWeekStart();
  const weekEnd = new Date(weekStart);

  weekEnd.setDate(weekEnd.getDate() + 7);

  return state.records.filter(record => {
    const recordDate = dateFromKey(record.date);

    return recordDate >= weekStart && recordDate < weekEnd;
  });
}

function renderSummary() {
  const totalMinutes = getTotalMinutes(state.records);
  const weeklyMinutes = getTotalMinutes(getWeeklyRecords());

  elements.totalHours.textContent =
    formatDuration(totalMinutes);

  elements.weeklyHours.textContent =
    formatDuration(weeklyMinutes);
}


function formatDateTimeLocal(timestamp) {
  if (timestamp == null) return "";

  const date = new Date(timestamp);

  return [
    date.getFullYear(),
    "-",
    pad(date.getMonth() + 1),
    "-",
    pad(date.getDate()),
    "T",
    pad(date.getHours()),
    ":",
    pad(date.getMinutes())
  ].join("");
}

function parseDateTimeInput(value) {
  if (value.trim() === "") {
    return { valid: true, timestamp: null };
  }

  if (!/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}$/.test(value)) {
    return { valid: false, timestamp: null };
  }

  const timestamp = new Date(value).getTime();

  return {
    valid: Number.isFinite(timestamp),
    timestamp
  };
}

function editRecord(record) {
  const dateValue = prompt(
    "Fecha de la jornada (AAAA-MM-DD):",
    record.date
  );

  if (dateValue === null) return;

  const nextDate = dateValue.trim();
  const parsedDate = new Date(`${nextDate}T00:00:00`);

  if (
    !/^\d{4}-\d{2}-\d{2}$/.test(nextDate) ||
    Number.isNaN(parsedDate.getTime()) ||
    toLocalDateKey(parsedDate) !== nextDate
  ) {
    alert("La fecha no es válida. Usa el formato AAAA-MM-DD.");
    return;
  }

  const duplicateDate = state.records.some(
    (item) => item.id !== record.id && item.date === nextDate
  );

  if (duplicateDate) {
    alert("Ya existe un registro para esa fecha.");
    return;
  }

  const statusOptions = {
  "1": { value: "attended", label: "Asistió" },
  "2": { value: "absent", label: "Falta" },
  "3": { value: "pending", label: "Pendiente" }
};

const currentStatus = {
  attended: "1",
  absent: "2",
  pending: "3"
};

const statusValue = prompt(
  "Selecciona el estado de la jornada:\n\n1. Asistió\n2. Falta\n3. Pendiente",
  currentStatus[record.status] || "1"
);

if (statusValue === null) return;

const selectedStatus = statusOptions[statusValue.trim()];

if (!selectedStatus) {
  alert("Opción no válida. Elige 1, 2 o 3.");
  return;
}

const nextStatus = selectedStatus.value;

  const entryValue = prompt(
    "Entrada (AAAA-MM-DDTHH:MM). Déjalo vacío si no hubo entrada:",
    formatDateTimeLocal(record.entry)
  );

  if (entryValue === null) return;

  const exitValue = prompt(
    "Salida (AAAA-MM-DDTHH:MM). Déjalo vacío si no hubo salida:",
    formatDateTimeLocal(record.exit)
  );

  if (exitValue === null) return;

  const parsedEntry = parseDateTimeInput(entryValue);
  const parsedExit = parseDateTimeInput(exitValue);

  if (!parsedEntry.valid || !parsedExit.valid) {
    alert("Una hora no es válida. Usa el formato AAAA-MM-DDTHH:MM.");
    return;
  }

  if (
    parsedEntry.timestamp !== null &&
    parsedExit.timestamp !== null &&
    parsedExit.timestamp <= parsedEntry.timestamp
  ) {
    alert("La salida debe ser posterior a la entrada.");
    return;
  }

  if (parsedEntry.timestamp === null && parsedExit.timestamp !== null) {
    alert("No puedes registrar una salida sin una entrada.");
    return;
  }

  if (nextStatus === "attended" && parsedEntry.timestamp === null) {
    alert("Una jornada con asistencia debe tener una entrada.");
    return;
  }

  // Aplicamos los cambios únicamente después de validar todos los campos.
  record.date = nextDate;
  record.entry = parsedEntry.timestamp;
  record.exit = parsedExit.timestamp;
  record.status = nextStatus;

  saveRecords();
  void syncRecordsToFirebase();
  renderApp();
}


function deleteRecord(record) {
  const date = dateFromKey(record.date);

  const formattedDate = formatDate(date, {
    day: "numeric",
    month: "long",
    year: "numeric"
  });

  const confirmed = confirm(
    `Vas a eliminar la jornada del ${formattedDate}.\n\n` +
    `Se eliminarán todos los datos de ese día, incluyendo ` +
    `la hora de entrada, la hora de salida y el estado de la jornada.\n\n` +
    `Esta acción no se puede deshacer desde la aplicación.\n\n` +
    `¿Deseas continuar?`
  );

  if (!confirmed) return;

  const index = state.records.findIndex(
    (item) => item.id === record.id
  );

  if (index === -1) {
    alert("No se encontró ese registro.");
    return;
  }

  state.records.splice(index, 1);

  saveRecords();
  void syncRecordsToFirebase();
  renderApp();
}


/* 11. Dibujar los historiales */

function createHistoryItem(record) {
  const item = document.createElement("article");
  item.className = "history-item";

  const date = dateFromKey(record.date);

  const dateIcon = document.createElement("div");
  dateIcon.className = "history-date-icon";

  dateIcon.textContent = [
    pad(date.getDate()),
    formatDate(date, { month: "short" })
      .replace(".", "")
      .toUpperCase()
  ].join(" ");

  const main = document.createElement("div");
  main.className = "history-main";

  const title = document.createElement("strong");

  title.textContent = formatDate(date, {
    weekday: "long",
    day: "numeric",
    month: "long"
  });

  const times = document.createElement("span");

  if (record.status === "absent") {
    times.textContent = "Sin asistencia (S/A)";

  } else if (record.status === "pending" && record.entry == null) {
    times.textContent = "Jornada pendiente";

  } else if (record.exit == null) {
    times.textContent =
      `Entrada ${formatTime(roundToNearestHalfHour(record.entry))} · Jornada en curso`;

  } else {
    times.textContent =
      `${formatTime(roundToNearestHalfHour(record.entry))} – ${formatTime(roundToNearestHalfHour(record.exit))}`;
  }

  main.append(title, times);

  const total = document.createElement("strong");
  total.className = "history-total";

  if (record.status === "absent") {
    total.textContent = "S/A";

  } else if (record.status === "pending" && record.entry == null) {
    total.textContent = "Pendiente";

  } else {
    total.textContent = formatDuration(getDurationMinutes(record));
  }


const actions = document.createElement("div");
actions.className = "history-actions";

const editButton = document.createElement("button");
editButton.type = "button";
editButton.className = "history-action-button";
editButton.textContent = "Editar";
editButton.addEventListener("click", () => {
  editRecord(record);
});

const deleteButton = document.createElement("button");
deleteButton.type = "button";
deleteButton.className =
  "history-action-button history-delete-button";
deleteButton.textContent = "Eliminar";
deleteButton.addEventListener("click", () => {
  deleteRecord(record);
});

actions.append(editButton, deleteButton);
item.append(dateIcon, main, total, actions);

return item;

}

function renderHistoryList(
  listElement,
  countElement,
  records,
  emptyTitle,
  emptyDescription
) {
  if (!listElement) {
    return;
  }

  const sortedRecords = [...records].sort(
    (a, b) => b.date.localeCompare(a.date)
  );

  listElement.replaceChildren();

  if (countElement) {
    countElement.textContent =
      `${sortedRecords.length} ${sortedRecords.length === 1 ? "registro" : "registros"}`;
  }

  if (sortedRecords.length === 0) {
    const empty = document.createElement("div");
    empty.className = "empty-state";

    const title = document.createElement("strong");
    title.textContent = emptyTitle;

    const description = document.createElement("p");
    description.textContent = emptyDescription;

    empty.append(title, description);
    listElement.append(empty);

    return;
  }

  sortedRecords.forEach(record => {
    listElement.append(createHistoryItem(record));
  });
}



function renderHistory() {
  const sortedRecords = [...state.records].sort(
    (a, b) => b.date.localeCompare(a.date)
  );

  const visibleRecords = state.historyExpanded
    ? sortedRecords
    : sortedRecords.slice(0, 6);

  renderHistoryList(
    elements.historyList,
    elements.recordCount,
    visibleRecords,
    "Tu historial aparecerá aquí",
    "Cuando registres jornadas, podrás consultarlas en esta sección."
  );

  if (elements.historyToggle) {
    const hasMore = sortedRecords.length > 6;

    elements.historyToggle.hidden = !hasMore;
    elements.historyToggle.textContent =
      state.historyExpanded ? "Ver menos" : "Ver más";

    elements.historyToggle.setAttribute(
      "aria-expanded",
      String(state.historyExpanded)
    );
  }
}



/* 12. Renderizar toda la aplicación */

function renderApp() {
  renderAttendance();
  renderSummary();
  renderHistory();
}

function handleLoginSubmit(event) {
  event.preventDefault();

  elements.loginError.textContent = "";
  elements.loginButton.disabled = true;

  signInWithEmailAndPassword(
    auth,
    elements.loginEmail.value.trim(),
    elements.loginPassword.value
  )
    .catch((error) => {
      console.error("Error al iniciar sesión:", error);

      elements.loginError.textContent =
        error.code === "auth/invalid-credential"
          ? "Correo o contraseña incorrectos."
          : "No se pudo iniciar sesión. Verifica tus datos e inténtalo de nuevo.";
    })
    .finally(() => {
      elements.loginButton.disabled = false;
    });
}

async function loadCloudRecords(user) {
  const userRecordsRef = ref(
    database,
    `users/${user.uid}/records`
  );

  recordsRef = userRecordsRef;

  try {
    const snapshot = await get(userRecordsRef);

    if (snapshot.exists()) {
      // Firebase ya tiene registros: usamos la nube como fuente principal.
      const cloudRecords = snapshot.val();

      state.records = Array.isArray(cloudRecords)
        ? cloudRecords
        : Object.values(cloudRecords);

      saveRecords();
      renderApp();
    } else {
      // Primera sincronización: conservamos y subimos el historial local.
      await set(userRecordsRef, state.records);
    }

    if (unsubscribeRecords) {
      unsubscribeRecords();
    }

    unsubscribeRecords = onValue(
      userRecordsRef,
      (updatedSnapshot) => {
        const cloudRecords = updatedSnapshot.val();

        state.records = cloudRecords
          ? Array.isArray(cloudRecords)
            ? cloudRecords
            : Object.values(cloudRecords)
          : [];

        saveRecords();
        renderApp();
      },
      (error) => {
        console.error("Error al escuchar Firebase:", error);
      }
    );
  } catch (error) {
    console.error("Error al cargar los registros:", error);

    alert(
      "Iniciaste sesión, pero no se pudieron cargar los registros de Firebase. Revisa las reglas de seguridad y la conexión."
    );
  }
}

function setupFirebaseAuth() {
  elements.loginForm.addEventListener(
    "submit",
    handleLoginSubmit
  );

  onAuthStateChanged(auth, async (user) => {
    state.currentUser = user;

    if (unsubscribeRecords) {
      unsubscribeRecords();
      unsubscribeRecords = null;
    }

    recordsRef = null;

    if (!user) {
      document.body.classList.add("auth-required");
      return;
    }

    document.body.classList.remove("auth-required");

    await loadCloudRecords(user);
  });
}


  async function exportExcel() {
  alert("¡El botón sí está ejecutando exportExcel!");


  const button = document.getElementById("export-excel-button");

  if (button) {
    button.disabled = true;
    button.textContent = "Generando Excel...";
  }

  try {
    const response = await fetch(
      "assets/Control_de_Asistencia_Bitacora.xlsx"
    );

    if (!response.ok) {
      throw new Error("No se encontró el archivo Excel de plantilla.");
    }

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(await response.arrayBuffer());

    const worksheet = workbook.getWorksheet("Control de Asistencia");

    if (!worksheet) {
      throw new Error("No se encontró la hoja Control de Asistencia.");
    }

    const startDate = new Date(2026, 7, 17);

    const addDays = (date, days) => {
      const result = new Date(date);
      result.setDate(result.getDate() + days);
      return result;
    };

    const dateKey = (date) =>
      `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;

    const formatDayMonth = (date) =>
      new Intl.DateTimeFormat("es-MX", {
        day: "numeric",
        month: "long"
      }).format(date);

    const formatExcelTime = (timestamp) => {
      if (timestamp == null) return "-";

      const rounded = roundToNearestHalfHour(timestamp);
      const date = new Date(rounded);

      return `${String(date.getHours()).padStart(2, "0")}:${String(date.getMinutes()).padStart(2, "0")} hrs`;
    };

    const copyRowFormatting = (sourceRow, targetRow) => {
      worksheet.getRow(targetRow).height =
        worksheet.getRow(sourceRow).height;

      for (let col = 1; col <= 6; col++) {
        const source = worksheet.getCell(sourceRow, col);
        const target = worksheet.getCell(targetRow, col);

        target.style = JSON.parse(JSON.stringify(source.style));
        target.numFmt = source.numFmt;
      }
    };

    // Localizar las semanas y el total general.
    let weekStarts = [];

    for (let row = 1; row <= worksheet.rowCount; row++) {
      const value = worksheet.getCell(row, 1).value;

      if (typeof value === "string" && /^Semana \d+:/i.test(value)) {
        weekStarts.push(row);
      }
    }

    let generalRow = 0;

    for (let row = 1; row <= worksheet.rowCount; row++) {
      if (
        worksheet.getCell(row, 1).value ===
        "TOTAL GENERAL ACUMULADO"
      ) {
        generalRow = row;
        break;
      }
    }

    if (weekStarts.length === 0 || generalRow === 0) {
      throw new Error("La plantilla no tiene la estructura esperada.");
    }

    // Determinar hasta qué semana generar el reporte.
    const today = new Date();
    let latestDate = new Date(today);

    for (const record of state.records) {
      if (!record.date) continue;

      const recordDate = new Date(`${record.date}T00:00:00`);

      if (!Number.isNaN(recordDate.getTime()) && recordDate > latestDate) {
        latestDate = recordDate;
      }
    }

    const daysUntilSunday = (7 - latestDate.getDay()) % 7;
    const lastSunday = addDays(latestDate, daysUntilSunday);

    let requiredWeeks = 1;

    while (
      addDays(startDate, requiredWeeks * 7 - 1) < lastSunday
    ) {
      requiredWeeks++;
    }

    // Crear las semanas que todavía no existen.
    const sourceWeekStart = weekStarts[weekStarts.length - 1];

    while (weekStarts.length < requiredWeeks) {
      const newStart = generalRow;

      worksheet.spliceRows(
        newStart,
        0,
        ...Array.from({ length: 11 }, () => [])
      );

      for (let offset = 0; offset <= 9; offset++) {
        copyRowFormatting(
          sourceWeekStart + offset,
          newStart + offset
        );
      }

      copyRowFormatting(16, newStart + 10);

      worksheet.mergeCells(newStart, 1, newStart, 6);

      worksheet.getCell(newStart, 1).value = "";
      worksheet.getRow(newStart + 1).values = [
        "DÍA",
        "HORA DE ENTRADA",
        "HORA DE SALIDA",
        "FIRMA DE ASISTENCIA",
        "OBSERVACIÓN",
        "TOTAL HORAS"
      ];

      const days = [
        "Lunes",
        "Martes",
        "Miércoles",
        "Jueves",
        "Viernes",
        "Sábado",
        "Domingo"
      ];

      for (let day = 0; day < 7; day++) {
        const row = newStart + 2 + day;

        worksheet.getCell(row, 1).value = days[day];
        worksheet.getCell(row, 2).value = "-";
        worksheet.getCell(row, 3).value = "-";
        worksheet.getCell(row, 4).value = null;
        worksheet.getCell(row, 5).value = "Sin registro en bitácora";
        worksheet.getCell(row, 6).value = 0;
      }

      worksheet.getCell(newStart + 9, 1).value = "TOTAL DE LA SEMANA";
      worksheet.getCell(newStart + 9, 6).value = {
        formula: `SUM(F${newStart + 2}:F${newStart + 8})`
      };

      generalRow += 11;
      weekStarts.push(newStart);
    }

    // Preparar los registros actuales de la aplicación.
    const recordsByDate = new Map();

    for (const record of state.records) {
      recordsByDate.set(record.date, record);
    }

    const weekdayNames = [
      "Lunes",
      "Martes",
      "Miércoles",
      "Jueves",
      "Viernes",
      "Sábado",
      "Domingo"
    ];

    const weeklyTotalCells = [];

    let totalMinutesAllWeeks = 0;

const getMinutesFromExcelCell = (cell) => {
  const value = cell.value;

  const numericValue =
    typeof value === "number"
      ? value
      : value &&
          typeof value === "object" &&
          typeof value.result === "number"
        ? value.result
        : null;

  if (numericValue === null || !Number.isFinite(numericValue)) {
    return 0;
  }

  const isTimeFormat = /h|m|s/i.test(cell.numFmt || "");

  return Math.round(numericValue * (isTimeFormat ? 1440 : 60));
};

    for (let week = 0; week < weekStarts.length; week++) {
      const firstRow = weekStarts[week];
      const weekStart = addDays(startDate, week * 7);
      const weekEnd = addDays(weekStart, 6);

      worksheet.getCell(firstRow, 1).value =
        `Semana ${week + 1}: del ${formatDayMonth(weekStart)} al ${formatDayMonth(weekEnd)}`;

      const dayRows = [];

      for (let day = 0; day < 7; day++) {
        const row = firstRow + 2 + day;
        const date = addDays(weekStart, day);
        const record = recordsByDate.get(dateKey(date));

        worksheet.getCell(row, 1).value = weekdayNames[day];
        dayRows.push(row);

        // Si no hay registro en la aplicación, conservar los
        // datos históricos de la plantilla.
        if (!record) continue;

        const entryCell = worksheet.getCell(row, 2);
        const exitCell = worksheet.getCell(row, 3);
        const signatureCell = worksheet.getCell(row, 4);
        const observationCell = worksheet.getCell(row, 5);
        const hoursCell = worksheet.getCell(row, 6);

        if (record.status === "absent") {
          entryCell.value = "-";
          exitCell.value = "-";
          signatureCell.value = null;
          observationCell.value = "Falta";
          hoursCell.value = 0;
        } else if (record.status === "pending") {
          entryCell.value = "-";
          exitCell.value = "-";
          signatureCell.value = null;
          observationCell.value = "Pendiente";
          hoursCell.value = 0;
        } else if (record.entry != null) {
          entryCell.value = formatExcelTime(record.entry);
          exitCell.value = formatExcelTime(record.exit);
          signatureCell.value = "[ FIRMA ]";

          if (record.exit == null) {
            observationCell.value = "Salida pendiente";
            hoursCell.value = 0;
          } else {
            observationCell.value = null;
            hoursCell.value = getDurationMinutes(record) / 1440;
            hoursCell.numFmt = "[h]:mm";
          }
        } else {
          entryCell.value = "-";
          exitCell.value = "-";
          signatureCell.value = null;
          observationCell.value = "Entrada pendiente";
          hoursCell.value = 0;
        }
      }

      
      const totalRow = firstRow + 9;

      worksheet.getCell(totalRow, 1).value = "TOTAL DE LA SEMANA";

      let weeklyMinutes = 0;

      for (const row of dayRows) {
        weeklyMinutes += getMinutesFromExcelCell(
          worksheet.getCell(row, 6)
        );
      }

      const weeklyTotalCell = worksheet.getCell(totalRow, 6);

      weeklyTotalCell.value = weeklyMinutes / 1440;
      weeklyTotalCell.numFmt = "[h]:mm";

      totalMinutesAllWeeks += weeklyMinutes;
    }


    // Actualizar el periodo y el total acumulado.
    worksheet.getCell(4, 6).value =
      `${formatDayMonth(startDate)} al ${formatDayMonth(lastSunday)}`;

    for (let row = 1; row <= worksheet.rowCount; row++) {
      if (
        worksheet.getCell(row, 1).value ===
        "TOTAL GENERAL ACUMULADO"
      ) {
        generalRow = row;
        break;
      }
    }

    const grandTotalCell = worksheet.getCell(generalRow, 6);

grandTotalCell.value = totalMinutesAllWeeks / 1440;
grandTotalCell.numFmt = "[h]:mm";

    workbook.calcProperties.fullCalcOnLoad = true;

    const buffer = await workbook.xlsx.writeBuffer();

    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });

    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");

    link.href = url;
    link.download = "Control_de_Asistencia_Bitacora_actualizado.xlsx";
    document.body.appendChild(link);
    link.click();
    link.remove();

    window.setTimeout(() => URL.revokeObjectURL(url), 1000);
  } catch (error) {
    console.error("Error al exportar Excel:", error);
    alert(
      "No se pudo generar el Excel. Revisa la consola del navegador para ver el error."
    );
  } finally {
    if (button) {
      button.disabled = false;
      button.textContent = "Descargar Excel";
    }
  }
}


/* 13. Inicializar */

function init() {
  elements.historyToggle?.addEventListener("click", () => {
    state.historyExpanded = !state.historyExpanded;
    renderHistory();
  });

  elements.entryButton.addEventListener(
    "click",
    handleEntryClick
  );

  elements.exitButton.addEventListener(
    "click",
    handleExitClick
  );

  document
  .getElementById("export-excel-button")
  ?.addEventListener("click", exportExcel);

  updateClock();
  renderApp();

  window.setInterval(updateClock, 1000);
}


init();
setupFirebaseAuth();