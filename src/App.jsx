import React, { useEffect, useMemo, useRef, useState } from "react";
import { createClient } from "@supabase/supabase-js";
import * as XLSX from "xlsx";

const ENV = typeof import.meta !== "undefined" && import.meta.env ? import.meta.env : {};
const SUPABASE_URL = ENV.VITE_SUPABASE_URL || "";
const SUPABASE_ANON_KEY = ENV.VITE_SUPABASE_ANON_KEY || "";
const supabase = SUPABASE_URL && SUPABASE_ANON_KEY ? createClient(SUPABASE_URL, SUPABASE_ANON_KEY) : null;

const DEVICE_ID_KEY = "aft-device-id-v7";

const COMBAT_MOS_SET = new Set([
  "11A", "11B", "11C", "11Z", "12A", "12B", "13A", "13F",
  "18A", "180A", "18B", "18C", "18D", "18E", "18F", "18Z",
  "19A", "19C", "19D", "19K", "19Z",
]);

const EVENTS = [
  {
    key: "MDL",
    name: "Maximum Deadlift",
    shortName: "Deadlift",
    fields: [
      { key: "value_1", label: "Attempt 1, lb", placeholder: "300", inputMode: "numeric" },
      { key: "value_2", label: "Attempt 2, lb", placeholder: "320", inputMode: "numeric" },
    ],
  },
  {
    key: "HRP",
    name: "Hand-Release Push-Up",
    shortName: "Push-up",
    fields: [{ key: "value_1", label: "Repetitions", placeholder: "42", inputMode: "numeric" }],
  },
  {
    key: "SDC",
    name: "Sprint-Drag-Carry",
    shortName: "SDC",
    fields: [{ key: "value_1", label: "Time, MMSS", placeholder: "215", inputMode: "numeric" }],
  },
  {
    key: "PLK",
    name: "Plank",
    shortName: "Plank",
    fields: [{ key: "value_1", label: "Time, MMSS", placeholder: "322", inputMode: "numeric" }],
  },
  {
    key: "2MR",
    name: "Two-Mile Run",
    shortName: "2MR",
    fields: [{ key: "value_1", label: "Time, MMSS", placeholder: "1732", inputMode: "numeric" }],
  },
];

const REQUIRED_ROSTER_COLUMNS = ["RN", "RANK", "NAME", "SEX", "MOS", "MSC", "AGE", "YYYYMMDD"];

const FULL_EXPORT_COLUMNS = [
  "RN", "RANK", "NAME", "SEX", "MOS", "MSC", "AGE", "YYYYMMDD",
  "AGE GROUP", "AFT GROUP", "PLATOON",
  "MDL1", "MDL2", "HRP", "SDC", "PLK", "2MR", "LAST_UPDATED",
  "MDL_GRADER", "HRP_GRADER", "SDC_GRADER", "PLK_GRADER", "TMR_GRADER",
];

const PYTHON_INPUT_COLUMNS = [
  "ROSTER", "RANK", "NAME", "SEX", "AGE", "MOS", "MSC", "DATE",
  "AGE GROUP", "AFT GROUP", "MDL1", "MDL2", "HRP", "SDC", "PLK", "2MR",
];

const EMPTY_SCORES = {
  MDL_SCORE: "",
  HRP_SCORE: "",
  SDC_SCORE: "",
  PLK_SCORE: "",
  "2MR_SCORE": "",
  TOTAL_SCORE: "",
  MDL_PF: "",
  HRP_PF: "",
  SDC_PF: "",
  PLK_PF: "",
  "2MR_PF": "",
  STATUS: "",
};

function normalizeText(value) {
  return String(value ?? "").trim();
}

function normalizeUpper(value) {
  return normalizeText(value).toUpperCase();
}

function parseAge(value) {
  const text = normalizeText(value);
  if (!text) return null;
  const number = Number(text);
  return Number.isInteger(number) && number >= 0 ? number : null;
}

function getDeviceId() {
  if (typeof window === "undefined") return "server";
  try {
    let id = window.localStorage.getItem(DEVICE_ID_KEY);
    if (!id) {
      id = `device-${Date.now()}-${Math.random().toString(16).slice(2)}`;
      window.localStorage.setItem(DEVICE_ID_KEY, id);
    }
    return id;
  } catch {
    return `device-${Date.now()}-${Math.random().toString(16).slice(2)}`;
  }
}

function sanitizeInitials(value) {
  return normalizeUpper(value).replace(/[^A-Z]/g, "").slice(0, 4);
}

function sanitizeRNInput(value) {
  return String(value ?? "").replace(/[^0-9A-Za-z-]/g, "").slice(0, 20);
}

function sanitizePerformanceValue(value) {
  return String(value ?? "").replace(/[^0-9:.]/g, "").slice(0, 12);
}

function classifyAgeGroup(ageValue) {
  const age = Number(ageValue);
  if (!Number.isFinite(age) || age < 17) return "";
  if (age <= 21) return "17-21";
  if (age <= 26) return "22-26";
  if (age <= 31) return "27-31";
  if (age <= 36) return "32-36";
  if (age <= 41) return "37-41";
  if (age <= 46) return "42-46";
  if (age <= 51) return "47-51";
  if (age <= 56) return "52-56";
  if (age <= 61) return "57-61";
  return "62+";
}

function classifyAftGroup(sexValue, mosValue) {
  const sex = normalizeUpper(sexValue);
  const mos = normalizeUpper(mosValue);
  if (sex === "M") return "M";
  if (sex === "F" && COMBAT_MOS_SET.has(mos)) return "M";
  if (sex === "F") return "F";
  return "";
}

function inferPlatoon(rnValue) {
  const first = normalizeText(rnValue)[0];
  if (first === "1") return "1ST PLATOON";
  if (first === "2") return "2ND PLATOON";
  if (first === "3") return "3RD PLATOON";
  if (first === "4") return "4TH PLATOON";
  return "UNKNOWN";
}

function toRosterRow(row) {
  const rn = normalizeText(row.RN ?? row.ROSTER ?? row.rn);
  const rank = normalizeUpper(row.RANK ?? row.rank);
  const name = normalizeText(row.NAME ?? row.name);
  const sex = normalizeUpper(row.SEX ?? row.sex);
  const mos = normalizeUpper(row.MOS ?? row.mos);
  const msc = normalizeText(row.MSC ?? row.msc);
  const age = parseAge(row.AGE ?? row.age);
  const yyyymmdd = normalizeText(row.YYYYMMDD ?? row.DATE ?? row.yyyymmdd);

  return {
    rn,
    rank,
    name,
    sex,
    mos,
    msc,
    age,
    yyyymmdd,
    age_group: classifyAgeGroup(age),
    aft_group: classifyAftGroup(sex, mos),
    platoon: inferPlatoon(rn),
  };
}

function parseCSV(text) {
  const rows = [];
  let current = [];
  let cell = "";
  let quoted = false;

  for (let i = 0; i < text.length; i += 1) {
    const char = text[i];
    const next = text[i + 1];

    if (char === '"' && quoted && next === '"') {
      cell += '"';
      i += 1;
    } else if (char === '"') {
      quoted = !quoted;
    } else if (char === "," && !quoted) {
      current.push(cell);
      cell = "";
    } else if ((char === "\n" || char === "\r") && !quoted) {
      if (char === "\r" && next === "\n") i += 1;
      current.push(cell);
      if (current.some((x) => normalizeText(x) !== "")) rows.push(current);
      current = [];
      cell = "";
    } else {
      cell += char;
    }
  }

  current.push(cell);
  if (current.some((x) => normalizeText(x) !== "")) rows.push(current);
  if (!rows.length) return [];

  const headers = rows[0].map((h) => String(h).replace(/^\uFEFF/, "").trim().toUpperCase());
  return rows.slice(1).map((row) => {
    const obj = {};
    headers.forEach((header, index) => {
      obj[header] = row[index] ?? "";
    });
    return obj;
  });
}

function csvEscape(value) {
  const text = value === undefined || value === null ? "" : String(value);
  if (/[",\n\r]/.test(text)) return `"${text.replace(/"/g, '""')}"`;
  return text;
}

function rowsToCSV(rows, columns) {
  const header = columns.map(csvEscape).join(",");
  const body = rows.map((row) => columns.map((column) => csvEscape(row[column])).join(",")).join("\n");
  return body ? `${header}\n${body}` : header;
}

function downloadCSV(filename, rows, columns) {
  if (typeof document === "undefined") return;
  const blob = new Blob([rowsToCSV(rows, columns)], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function makeExportRows(rows, columns) {
  return rows.map((row) => {
    const item = {};
    columns.forEach((column) => {
      item[column] = row && row[column] !== undefined && row[column] !== null ? row[column] : "";
    });
    return item;
  });
}

function downloadXLSX(filename, rows, columns, sheetName) {
  const safeRows = makeExportRows(rows, columns);
  const worksheet = XLSX.utils.json_to_sheet(safeRows, { header: columns });
  worksheet["!cols"] = columns.map((column) => ({ wch: Math.max(12, String(column).length + 2) }));
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, sheetName.slice(0, 31));
  XLSX.writeFile(workbook, filename);
}

function eventIsComplete(record, eventKey) {
  if (!record) return false;
  if (eventKey === "MDL") return Boolean(record.MDL1 && record.MDL2 && record.MDL_GRADER);
  if (eventKey === "HRP") return Boolean(record.HRP && record.HRP_GRADER);
  if (eventKey === "SDC") return Boolean(record.SDC && record.SDC_GRADER);
  if (eventKey === "PLK") return Boolean(record.PLK && record.PLK_GRADER);
  if (eventKey === "2MR") return Boolean(record["2MR"] && record.TMR_GRADER);
  return false;
}

function makeEmptyCombinedRow(r) {
  return {
    RN: r.rn,
    RANK: r.rank || "",
    NAME: r.name || "",
    SEX: r.sex || "",
    MOS: r.mos || "",
    MSC: r.msc || "",
    AGE: r.age ?? "",
    YYYYMMDD: r.yyyymmdd || "",
    "AGE GROUP": r.age_group || classifyAgeGroup(r.age),
    "AFT GROUP": r.aft_group || classifyAftGroup(r.sex, r.mos),
    PLATOON: r.platoon || inferPlatoon(r.rn),
    MDL1: "",
    MDL2: "",
    MDL_GRADER: "",
    HRP: "",
    HRP_GRADER: "",
    SDC: "",
    SDC_GRADER: "",
    PLK: "",
    PLK_GRADER: "",
    "2MR": "",
    TMR_GRADER: "",
    ...EMPTY_SCORES,
    LAST_UPDATED: r.updated_at || r.created_at || "",
  };
}

function buildCombinedRows(rosterRows, eventRows) {
  const byRn = new Map();

  rosterRows.forEach((r) => {
    if (!r?.rn) return;
    byRn.set(String(r.rn), makeEmptyCombinedRow(r));
  });

  eventRows.forEach((event) => {
    const row = byRn.get(String(event.rn));
    if (!row) return;
    row.LAST_UPDATED = event.updated_at || event.created_at || row.LAST_UPDATED;

    if (event.event_key === "MDL") {
      row.MDL1 = event.value_1 || "";
      row.MDL2 = event.value_2 || "";
      row.MDL_GRADER = event.grader_initials || "";
    } else if (event.event_key === "HRP") {
      row.HRP = event.value_1 || "";
      row.HRP_GRADER = event.grader_initials || "";
    } else if (event.event_key === "SDC") {
      row.SDC = event.value_1 || "";
      row.SDC_GRADER = event.grader_initials || "";
    } else if (event.event_key === "PLK") {
      row.PLK = event.value_1 || "";
      row.PLK_GRADER = event.grader_initials || "";
    } else if (event.event_key === "2MR") {
      row["2MR"] = event.value_1 || "";
      row.TMR_GRADER = event.grader_initials || "";
    }
  });

  return Array.from(byRn.values()).sort((a, b) => String(a.RN).localeCompare(String(b.RN), undefined, { numeric: true }));
}

function getScoringScale(row) {
  const sex = normalizeUpper(row.SEX);
  const mos = normalizeUpper(row.MOS);
  if (sex === "M") return "M";
  if (sex === "F" && COMBAT_MOS_SET.has(mos)) return "M";
  if (sex === "F") return "F";
  return "";
}

function valueToNumber(value) {
  const text = normalizeUpper(value);
  if (!text) return null;
  if (text === "PROFILE") return "PROFILE";
  const number = Number(text.replace(/:/g, ""));
  return Number.isFinite(number) ? number : null;
}

function scoreByThreshold(value, thresholds, higherIsBetter) {
  if (value === "PROFILE") return 60;
  if (value === null) return "";
  const sorted = [...thresholds].sort((a, b) => higherIsBetter ? b.value - a.value : a.value - b.value);
  for (const item of sorted) {
    if (higherIsBetter && value >= item.value) return item.score;
    if (!higherIsBetter && value <= item.value) return item.score;
  }
  return 0;
}

function estimateMdlScore(row, value) {
  const scale = getScoringScale(row);
  const age = Number(row.AGE);
  let min = scale === "F" ? 60 : 80;
  let max = scale === "F" ? 230 : 350;
  if (age >= 57) {
    min = scale === "F" ? 60 : 80;
    max = scale === "F" ? 170 : 250;
  }
  const thresholds = [
    { value: max, score: 100 },
    { value: Math.round((max * 0.9) / 10) * 10, score: 90 },
    { value: Math.round((max * 0.8) / 10) * 10, score: 80 },
    { value: Math.round((max * 0.7) / 10) * 10, score: 70 },
    { value: Math.max(min, Math.round((max * 0.6) / 10) * 10), score: 60 },
  ];
  return scoreByThreshold(value, thresholds, true);
}

function estimateHrpScore(row, value) {
  const scale = getScoringScale(row);
  const age = Number(row.AGE);
  let sixty = scale === "F" ? 10 : 14;
  let hundred = scale === "F" ? 50 : 60;
  if (age >= 42) {
    sixty = scale === "F" ? 10 : 11;
    hundred = scale === "F" ? 40 : 57;
  }
  if (age >= 57) {
    sixty = 10;
    hundred = scale === "F" ? 24 : 46;
  }
  const thresholds = [
    { value: hundred, score: 100 },
    { value: Math.round(sixty + (hundred - sixty) * 0.75), score: 90 },
    { value: Math.round(sixty + (hundred - sixty) * 0.5), score: 80 },
    { value: Math.round(sixty + (hundred - sixty) * 0.25), score: 70 },
    { value: sixty, score: 60 },
  ];
  return scoreByThreshold(value, thresholds, true);
}

function estimateSdcScore(row, value) {
  const scale = getScoringScale(row);
  const sixty = scale === "F" ? 315 : 232;
  const hundred = scale === "F" ? 155 : 130;
  const thresholds = [
    { value: hundred, score: 100 },
    { value: Math.round(hundred + (sixty - hundred) * 0.25), score: 90 },
    { value: Math.round(hundred + (sixty - hundred) * 0.5), score: 80 },
    { value: Math.round(hundred + (sixty - hundred) * 0.75), score: 70 },
    { value: sixty, score: 60 },
  ];
  return scoreByThreshold(value, thresholds, false);
}

function estimatePlkScore(value) {
  const thresholds = [
    { value: 340, score: 100 },
    { value: 300, score: 90 },
    { value: 230, score: 80 },
    { value: 200, score: 70 },
    { value: 120, score: 60 },
  ];
  return scoreByThreshold(value, thresholds, true);
}

function estimateTmrScore(row, value) {
  const scale = getScoringScale(row);
  const age = Number(row.AGE);
  let sixty = scale === "F" ? 2448 : 2200;
  let hundred = scale === "F" ? 1718 : 1322;
  if (age >= 57) {
    sixty = scale === "F" ? 2448 : 2336;
    hundred = scale === "F" ? 1718 : 1528;
  }
  const thresholds = [
    { value: hundred, score: 100 },
    { value: Math.round(hundred + (sixty - hundred) * 0.25), score: 90 },
    { value: Math.round(hundred + (sixty - hundred) * 0.5), score: 80 },
    { value: Math.round(hundred + (sixty - hundred) * 0.75), score: 70 },
    { value: sixty, score: 60 },
  ];
  return scoreByThreshold(value, thresholds, false);
}

function passFail(score) {
  return typeof score === "number" && score >= 60 ? "PASS" : "FAIL";
}

function calculateAftScoreForRow(row) {
  const next = { ...row, ...EMPTY_SCORES };
  next.PLATOON = row.PLATOON || inferPlatoon(row.RN);

  if (!(row.MDL1 && row.MDL2 && row.HRP && row.SDC && row.PLK && row["2MR"])) return next;

  const mdl1 = valueToNumber(row.MDL1);
  const mdl2 = valueToNumber(row.MDL2);
  const hrp = valueToNumber(row.HRP);
  const sdc = valueToNumber(row.SDC);
  const plk = valueToNumber(row.PLK);
  const tmr = valueToNumber(row["2MR"]);

  const mdlMax = mdl1 === "PROFILE" || mdl2 === "PROFILE" ? "PROFILE" : Math.max(Number(mdl1) || 0, Number(mdl2) || 0);
  const mdlScore = estimateMdlScore(row, mdlMax);
  const hrpScore = estimateHrpScore(row, hrp);
  const sdcScore = estimateSdcScore(row, sdc);
  const plkScore = estimatePlkScore(plk);
  const tmrScore = estimateTmrScore(row, tmr);

  const numericScores = [mdlScore, hrpScore, sdcScore, plkScore, tmrScore].map((score) => Number(score) || 0);
  const totalScore = numericScores.reduce((sum, score) => sum + score, 0);
  const allPass = numericScores.every((score) => score >= 60);
  const combatMaleScale = getScoringScale(row) === "M" && COMBAT_MOS_SET.has(normalizeUpper(row.MOS));

  next.MDL_SCORE = mdlScore;
  next.HRP_SCORE = hrpScore;
  next.SDC_SCORE = sdcScore;
  next.PLK_SCORE = plkScore;
  next["2MR_SCORE"] = tmrScore;
  next.TOTAL_SCORE = totalScore;
  next.MDL_PF = passFail(mdlScore);
  next.HRP_PF = passFail(hrpScore);
  next.SDC_PF = passFail(sdcScore);
  next.PLK_PF = passFail(plkScore);
  next["2MR_PF"] = passFail(tmrScore);
  next.STATUS = allPass && (!combatMaleScale || totalScore >= 350) ? "PASS" : "RETEST";
  return next;
}

function safeCalculateRows(rows) {
  const baseRows = Array.isArray(rows) ? rows.map((row) => ({ ...row, ...EMPTY_SCORES, ...row })) : [];
  try {
    return baseRows.map(calculateAftScoreForRow);
  } catch {
    return baseRows;
  }
}

function pythonReadyRows(rows) {
  return rows.map((r) => ({
    ROSTER: r.RN,
    RANK: r.RANK,
    NAME: r.NAME,
    SEX: r.SEX,
    AGE: r.AGE,
    MOS: r.MOS,
    MSC: r.MSC,
    DATE: r.YYYYMMDD,
    "AGE GROUP": r["AGE GROUP"],
    "AFT GROUP": r["AFT GROUP"],
    MDL1: r.MDL1,
    MDL2: r.MDL2,
    HRP: r.HRP,
    SDC: r.SDC,
    PLK: r.PLK,
    "2MR": r["2MR"],
  }));
}

function readEventValuesFromRecord(record, eventKey) {
  if (!record) return { value_1: "", value_2: "" };
  if (eventKey === "MDL") return { value_1: record.MDL1 || "", value_2: record.MDL2 || "" };
  if (eventKey === "HRP") return { value_1: record.HRP || "", value_2: "" };
  if (eventKey === "SDC") return { value_1: record.SDC || "", value_2: "" };
  if (eventKey === "PLK") return { value_1: record.PLK || "", value_2: "" };
  if (eventKey === "2MR") return { value_1: record["2MR"] || "", value_2: "" };
  return { value_1: "", value_2: "" };
}

function runSelfTests() {
  const parsed = parseCSV('RN,RANK,NAME,SEX,MOS,MSC,AGE,YYYYMMDD\n101,PFC,"LANDA, EDWARD",M,92Y,8A NCOA,24,20260517');
  console.assert(parsed.length === 1, "CSV parses one row");
  console.assert(parsed[0].NAME === "LANDA, EDWARD", "CSV preserves quoted comma");
  console.assert(classifyAgeGroup("24") === "22-26", "Age group works");
  console.assert(classifyAftGroup("F", "11B") === "M", "Combat MOS uses male standard");
  console.assert(sanitizeInitials("e.l1") === "EL", "Initials sanitize");
  console.assert(sanitizeRNInput(" 10 1! ") === "101", "RN sanitizes");
  console.assert(sanitizePerformanceValue("1:23abc") === "1:23", "Performance sanitizes");
  console.assert(toRosterRow({ RN: "1", NAME: "A", AGE: "abc" }).age === null, "Invalid age becomes null");
  const combined = buildCombinedRows(
    [toRosterRow({ RN: "101", RANK: "PFC", NAME: "TEST", SEX: "M", MOS: "92Y", MSC: "8A", AGE: "24", YYYYMMDD: "20260517" })],
    [
      { rn: "101", event_key: "MDL", value_1: "300", value_2: "320", grader_initials: "EL" },
      { rn: "101", event_key: "HRP", value_1: "42", grader_initials: "EL" },
      { rn: "101", event_key: "SDC", value_1: "215", grader_initials: "EL" },
      { rn: "101", event_key: "PLK", value_1: "322", grader_initials: "EL" },
      { rn: "101", event_key: "2MR", value_1: "1732", grader_initials: "EL" },
    ]
  );
  const scored = safeCalculateRows(combined);
  console.assert(scored[0].HRP === "42" && scored[0].HRP_GRADER === "EL", "Combined rows merge event data");
  console.assert(scored[0].TOTAL_SCORE !== "", "Scoring produces total score");
  console.assert(rowsToCSV([{ NAME: "LANDA, EDWARD" }], ["NAME"]).includes('"LANDA, EDWARD"'), "CSV quotes commas");
}

if (typeof window !== "undefined" && !window.__AFT_SELF_TESTS_RAN__) {
  window.__AFT_SELF_TESTS_RAN__ = true;
  runSelfTests();
}

export default function App() {
  const [deviceId] = useState(() => getDeviceId());
  const [mode, setMode] = useState("grader");
  const [rosterRows, setRosterRows] = useState([]);
  const [eventRows, setEventRows] = useState([]);
  const [selectedEventKey, setSelectedEventKey] = useState("MDL");
  const [graderInitials, setGraderInitials] = useState("");
  const [rnInput, setRnInput] = useState("");
  const [adminSelectedRN, setAdminSelectedRN] = useState("");
  const [formValues, setFormValues] = useState({ value_1: "", value_2: "" });
  const [message, setMessage] = useState(supabase ? "Connected to Supabase. Loading data." : "Missing Supabase environment variables. Add VITE_SUPABASE_URL and VITE_SUPABASE_ANON_KEY.");
  const [busy, setBusy] = useState(false);
  const [syncing, setSyncing] = useState(false);
  const [editing, setEditing] = useState(false);
  const [exportPreview, setExportPreview] = useState({ title: "", csv: "" });
  const mountedRef = useRef(false);
  const refreshTimerRef = useRef(null);
  const editingTimerRef = useRef(null);
  const loadSequenceRef = useRef(0);

  const combinedRows = useMemo(() => buildCombinedRows(rosterRows, eventRows), [rosterRows, eventRows]);
  const scoredRows = useMemo(() => safeCalculateRows(combinedRows), [combinedRows]);
  const selectedEvent = EVENTS.find((event) => event.key === selectedEventKey) || EVENTS[0];
  const selectedRecord = scoredRows.find((row) => String(row.RN).toLowerCase() === normalizeText(rnInput).toLowerCase());
  const adminSelectedRecord = scoredRows.find((row) => String(row.RN) === String(adminSelectedRN));

  const stats = useMemo(() => {
    const total = scoredRows.length;
    const complete = scoredRows.filter((row) => EVENTS.every((event) => eventIsComplete(row, event.key))).length;
    const partial = scoredRows.filter((row) => EVENTS.some((event) => eventIsComplete(row, event.key)) && !EVENTS.every((event) => eventIsComplete(row, event.key))).length;
    return { total, complete, partial, incomplete: total - complete };
  }, [scoredRows]);

  async function loadAllData(options = {}) {
    if (!supabase) return;
    const sequence = loadSequenceRef.current + 1;
    loadSequenceRef.current = sequence;
    if (!options.silent) setSyncing(true);

    try {
      const [{ data: roster, error: rosterError }, { data: events, error: eventError }] = await Promise.all([
        supabase.from("aft_roster").select("*").order("rn", { ascending: true }),
        supabase.from("aft_event_records").select("*").order("updated_at", { ascending: false }),
      ]);
      if (sequence !== loadSequenceRef.current || !mountedRef.current) return;
      if (rosterError) throw rosterError;
      if (eventError) throw eventError;
      setRosterRows(roster || []);
      setEventRows(events || []);
      if (!options.silent) setMessage(`Synced. ${roster?.length || 0} roster records, ${events?.length || 0} event records.`);
    } catch (error) {
      if (mountedRef.current && !options.silent) setMessage(`Sync failed: ${error.message}`);
    } finally {
      if (mountedRef.current && !options.silent) setSyncing(false);
    }
  }

  useEffect(() => {
    mountedRef.current = true;
    if (!supabase) {
      return () => {
        mountedRef.current = false;
      };
    }

    loadAllData();

    const scheduleRefresh = () => {
      if (refreshTimerRef.current) window.clearTimeout(refreshTimerRef.current);
      refreshTimerRef.current = window.setTimeout(() => {
        loadAllData({ silent: true });
      }, 350);
    };

    const channel = supabase
      .channel("aft-live-updates")
      .on("postgres_changes", { event: "*", schema: "public", table: "aft_roster" }, scheduleRefresh)
      .on("postgres_changes", { event: "*", schema: "public", table: "aft_event_records" }, scheduleRefresh)
      .subscribe();

    return () => {
      mountedRef.current = false;
      if (refreshTimerRef.current) window.clearTimeout(refreshTimerRef.current);
      if (editingTimerRef.current) window.clearTimeout(editingTimerRef.current);
      supabase.removeChannel(channel);
    };
  }, []);

  useEffect(() => {
    if (editing) return;
    setFormValues(readEventValuesFromRecord(selectedRecord, selectedEventKey));
  }, [editing, rnInput, selectedEventKey, selectedRecord?.RN, selectedRecord?.LAST_UPDATED]);

  useEffect(() => {
    if (!scoredRows.length) {
      if (adminSelectedRN) setAdminSelectedRN("");
      return;
    }
    if (!adminSelectedRN || !scoredRows.some((row) => row.RN === adminSelectedRN)) {
      setAdminSelectedRN(scoredRows[0].RN);
    }
  }, [scoredRows, adminSelectedRN]);

  function updateFormValue(key, value) {
    setEditing(true);
    if (editingTimerRef.current) window.clearTimeout(editingTimerRef.current);
    editingTimerRef.current = window.setTimeout(() => setEditing(false), 2500);
    setFormValues((prev) => ({ ...prev, [key]: sanitizePerformanceValue(value) }));
  }

  function upsertLocalEventRecord(payload) {
    setEventRows((current) => {
      const now = new Date().toISOString();
      const nextPayload = { ...payload, updated_at: now, created_at: now };
      const exists = current.some((row) => row.rn === payload.rn && row.event_key === payload.event_key);
      if (!exists) return [nextPayload, ...current];
      return current.map((row) => (row.rn === payload.rn && row.event_key === payload.event_key ? { ...row, ...nextPayload, created_at: row.created_at || now } : row));
    });
  }

  async function handleRosterUpload(e) {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = async () => {
      setBusy(true);
      try {
        const parsed = parseCSV(String(reader.result || ""));
        if (!parsed.length) throw new Error("CSV file is empty.");
        const headers = Object.keys(parsed[0]).map((h) => h.toUpperCase());
        const missing = REQUIRED_ROSTER_COLUMNS.filter((column) => !headers.includes(column));
        if (missing.length) throw new Error(`Missing columns: ${missing.join(", ")}`);

        const uploadRows = parsed.map(toRosterRow).filter((row) => row.rn && row.name);
        const seen = new Set();
        const dedupedRows = uploadRows.filter((row) => {
          if (seen.has(row.rn)) return false;
          seen.add(row.rn);
          return true;
        });
        if (!dedupedRows.length) throw new Error("No valid roster rows found.");

        if (!supabase) {
          setRosterRows(dedupedRows);
          setEventRows([]);
          setRnInput("");
          setAdminSelectedRN(dedupedRows[0]?.rn || "");
          setMessage(`Roster loaded locally: ${dedupedRows.length} soldiers. Add Supabase env variables for shared multi-phone database mode.`);
          return;
        }

        const { error } = await supabase.from("aft_roster").upsert(dedupedRows, { onConflict: "rn" });
        if (error) throw error;
        setMessage(`Roster uploaded to shared database: ${dedupedRows.length} soldiers.`);
        await loadAllData();
      } catch (error) {
        setMessage(`Roster upload failed: ${error.message}`);
      } finally {
        if (mountedRef.current) setBusy(false);
      }
    };
    reader.onerror = () => setMessage("Roster upload failed: file could not be read.");
    reader.readAsText(file);
    e.target.value = "";
  }

  async function submitEventRecord() {
    if (!selectedRecord) {
      setMessage("Type a valid RN first.");
      return;
    }

    const initials = sanitizeInitials(graderInitials);
    if (!initials) {
      setMessage("Enter grader initials first.");
      return;
    }

    const missing = selectedEvent.fields.filter((field) => !normalizeText(formValues[field.key]));
    if (missing.length) {
      setMessage(`Missing ${missing.map((field) => field.label).join(", ")}.`);
      return;
    }

    const payload = {
      rn: selectedRecord.RN,
      event_key: selectedEvent.key,
      value_1: normalizeText(formValues.value_1),
      value_2: selectedEvent.key === "MDL" ? normalizeText(formValues.value_2) : null,
      grader_initials: initials,
      device_id: deviceId,
    };

    setBusy(true);
    try {
      if (!supabase) {
        upsertLocalEventRecord(payload);
        setMessage(`${selectedEvent.key} saved locally for ${selectedRecord.RANK} ${selectedRecord.NAME}. Add Supabase env variables for shared multi-phone database mode.`);
      } else {
        const { error } = await supabase.from("aft_event_records").upsert(payload, { onConflict: "rn,event_key" });
        if (error) throw error;
        setMessage(`${selectedEvent.key} saved to shared database for ${selectedRecord.RANK} ${selectedRecord.NAME}.`);
        await loadAllData({ silent: true });
      }
      setEditing(false);
      if (mountedRef.current) {
        setRnInput("");
        setFormValues({ value_1: "", value_2: "" });
      }
    } catch (error) {
      setMessage(`Submit failed: ${error.message}`);
    } finally {
      if (mountedRef.current) setBusy(false);
    }
  }

  function prepareExport(filename, rows, columns) {
    if (!rows.length) {
      setMessage("No rows available to export.");
      setExportPreview({ title: filename, csv: rowsToCSV([], columns) });
      return;
    }
    const safeRows = makeExportRows(rows, columns);
    const csv = rowsToCSV(safeRows, columns);
    setExportPreview({ title: filename, csv });
    try {
      downloadCSV(filename, safeRows, columns);
      setMessage(`${filename} prepared.`);
    } catch {
      setMessage(`${filename} prepared. Copy the CSV from the preview below if download is blocked.`);
    }
  }

  function exportFullRecords() {
    prepareExport("AFT_full_web_records.csv", scoredRows, FULL_EXPORT_COLUMNS);
  }

  function exportPythonInput() {
    prepareExport("AFT_python_ready_input.csv", pythonReadyRows(scoredRows), PYTHON_INPUT_COLUMNS);
  }

  function exportFullRecordsExcel() {
    if (!scoredRows.length) {
      setMessage("No rows available to export as Excel.");
      return;
    }
    try {
      downloadXLSX("AFT_full_web_records.xlsx", scoredRows, FULL_EXPORT_COLUMNS, "AFT_FULL_RECORDS");
      setMessage("AFT_full_web_records.xlsx downloaded.");
    } catch (error) {
      setMessage(`Excel export failed: ${error.message}`);
    }
  }

  function exportPythonInputExcel() {
    const rows = pythonReadyRows(scoredRows);
    if (!rows.length) {
      setMessage("No rows available to export as Excel.");
      return;
    }
    try {
      downloadXLSX("AFT_python_ready_input.xlsx", rows, PYTHON_INPUT_COLUMNS, "AFT_PYTHON_INPUT");
      setMessage("AFT_python_ready_input.xlsx downloaded.");
    } catch (error) {
      setMessage(`Excel export failed: ${error.message}`);
    }
  }

  async function copyExportPreview() {
    if (!exportPreview.csv) return;
    try {
      await navigator.clipboard.writeText(exportPreview.csv);
      setMessage("CSV copied.");
    } catch {
      setMessage("Clipboard blocked. Select the CSV text manually.");
    }
  }

  async function clearEventForAdminSelectedRN(eventKey) {
    if (!adminSelectedRecord) {
      setMessage("Select an RN first.");
      return;
    }
    setBusy(true);
    try {
      if (!supabase) {
        setEventRows((current) => current.filter((row) => !(row.rn === adminSelectedRecord.RN && row.event_key === eventKey)));
        setMessage(`${eventKey} cleared locally for RN ${adminSelectedRecord.RN}.`);
      } else {
        const { error } = await supabase.from("aft_event_records").delete().eq("rn", adminSelectedRecord.RN).eq("event_key", eventKey);
        if (error) throw error;
        setMessage(`${eventKey} cleared for RN ${adminSelectedRecord.RN}.`);
        await loadAllData({ silent: true });
      }
    } catch (error) {
      setMessage(`Clear failed: ${error.message}`);
    } finally {
      if (mountedRef.current) setBusy(false);
    }
  }

  const eventStatus = selectedRecord
    ? EVENTS.map((event) => ({ event, complete: eventIsComplete(selectedRecord, event.key) }))
    : [];

  const isDisabled = busy || syncing;
  const uploadDisabled = isDisabled;

  return (
    <div style={styles.page}>
      <div style={styles.shell}>
        <header style={styles.header}>
          <div>
            <p style={styles.eyebrow}>WELCOME!</p>
            <h1 style={styles.title}>BLC AFT SCORING AUTOMATION SYSTEM</h1>
            <p style={styles.subtitle}>Please enter your grader initial, student RN, and records for each subject.</p>
          </div>
          <div style={supabase ? styles.onlineBadge : styles.offlineBadge}>{supabase ? "SUPABASE READY" : "ENV NOT SET"}</div>
        </header>

        <div style={styles.modeTabs}>
          <button type="button" style={{ ...styles.modeButton, ...(mode === "grader" ? styles.modeButtonActive : {}) }} onClick={() => setMode("grader")}>Grader</button>
          <button type="button" style={{ ...styles.modeButton, ...(mode === "admin" ? styles.modeButtonActive : {}) }} onClick={() => setMode("admin")}>Admin</button>
        </div>

        {mode === "grader" ? (
          <main style={styles.stack}>
            <section style={styles.card}>
              <label style={styles.label}>Grader initials</label>
              <input style={styles.bigInput} value={graderInitials} onChange={(e) => setGraderInitials(sanitizeInitials(e.target.value))} placeholder="EL" maxLength={4} autoCapitalize="characters" />
            </section>

            <section style={styles.card}>
              <label style={styles.label}>Event station</label>
              <div style={styles.eventGrid}>
                {EVENTS.map((event) => (
                  <button key={event.key} type="button" style={{ ...styles.eventButton, ...(selectedEventKey === event.key ? styles.eventButtonActive : {}) }} onClick={() => setSelectedEventKey(event.key)}>
                    <strong>{event.key}</strong><span>{event.shortName}</span>
                  </button>
                ))}
              </div>
            </section>

            <section style={styles.card}>
              <label style={styles.label}>Roster number, RN</label>
              <input style={styles.rnInput} value={rnInput} onChange={(e) => setRnInput(sanitizeRNInput(e.target.value))} placeholder="101" inputMode="numeric" />
              {selectedRecord ? (
                <div style={styles.soldierBox}>
                  <strong>{selectedRecord.RN} · {selectedRecord.RANK} · {selectedRecord.NAME}</strong>
                  <span>{selectedRecord.SEX} · age {selectedRecord.AGE} · MOS {selectedRecord.MOS} · {selectedRecord.MSC}</span>
                </div>
              ) : <p style={styles.hint}>Type a valid RN from the uploaded roster.</p>}
            </section>

            <section style={styles.card}>
              <h2 style={styles.sectionTitle}>{selectedEvent.name}</h2>
              <div style={styles.formStack}>
                {selectedEvent.fields.map((field) => (
                  <label key={field.key} style={styles.label}>{field.label}
                    <input
                      style={styles.bigInput}
                      value={formValues[field.key] || ""}
                      onChange={(e) => updateFormValue(field.key, e.target.value)}
                      onFocus={() => setEditing(true)}
                      onBlur={() => {
                        if (editingTimerRef.current) window.clearTimeout(editingTimerRef.current);
                        editingTimerRef.current = window.setTimeout(() => setEditing(false), 600);
                      }}
                      placeholder={field.placeholder}
                      inputMode={field.inputMode}
                    />
                  </label>
                ))}
              </div>
              <button type="button" style={{ ...styles.primaryButton, ...(isDisabled ? styles.disabled : {}) }} disabled={isDisabled} onClick={submitEventRecord}>{isDisabled ? "Working..." : `Submit ${selectedEvent.key}`}</button>
            </section>

            {selectedRecord ? (
              <section style={styles.card}>
                <h2 style={styles.sectionTitle}>Current RN status</h2>
                <div style={styles.statusGrid}>{eventStatus.map(({ event, complete }) => <div key={event.key} style={complete ? styles.donePill : styles.missingPill}>{event.key}: {complete ? "done" : "missing"}</div>)}</div>
              </section>
            ) : null}
          </main>
        ) : (
          <main style={styles.adminGrid}>
            <section style={styles.card}>
              <h2 style={styles.sectionTitle}>Upload master roster</h2>
              <p style={styles.hint}>CSV columns: RN, RANK, NAME, SEX, MOS, MSC, AGE, YYYYMMDD.</p>
              <label style={{ ...styles.primaryButton, ...(uploadDisabled ? styles.disabled : {}) }}>Upload CSV<input type="file" accept=".csv" onChange={handleRosterUpload} disabled={uploadDisabled} style={{ display: "none" }} /></label>
            </section>

            <section style={styles.card}>
              <h2 style={styles.sectionTitle}>Live progress</h2>
              <div style={styles.statsGrid}>
                <Stat label="Total" value={stats.total} />
                <Stat label="Complete" value={stats.complete} />
                <Stat label="Partial" value={stats.partial} />
                <Stat label="Incomplete" value={stats.incomplete} />
              </div>
              <button type="button" style={{ ...styles.secondaryButton, ...(isDisabled ? styles.disabled : {}) }} disabled={isDisabled} onClick={() => supabase ? loadAllData() : setMessage("Refresh is only needed after Supabase is configured. Local roster data is already loaded.")}>Refresh</button>
            </section>

            <section style={{ ...styles.card, gridColumn: "1 / -1" }}>
              <div style={styles.adminHeader}>
                <h2 style={styles.sectionTitle}>Export and review</h2>
                <div style={styles.actions}>
                  <button type="button" style={styles.secondaryButton} onClick={exportPythonInputExcel}>Python input Excel</button>
                  <button type="button" style={styles.secondaryButton} onClick={exportFullRecordsExcel}>Full records Excel</button>
                  <button type="button" style={styles.secondaryButton} onClick={exportPythonInput}>Python input CSV</button>
                  <button type="button" style={styles.secondaryButton} onClick={exportFullRecords}>Full records CSV</button>
                </div>
              </div>
              <div style={styles.tableWrap}>
                <table style={styles.table}>
                  <thead><tr>{["RN", "Name", "MDL", "HRP", "SDC", "PLK", "2MR", "Score / Status"].map((h) => <th key={h} style={styles.th}>{h}</th>)}</tr></thead>
                  <tbody>
                    {scoredRows.map((r) => (
                      <tr key={r.RN} onClick={() => setAdminSelectedRN(r.RN)} style={adminSelectedRN === r.RN ? styles.selectedRow : styles.clickableRow}>
                        <td style={styles.td}>{r.RN}</td>
                        <td style={styles.td}>{r.RANK} {r.NAME}</td>
                        <td style={styles.td}>{r.MDL1 || "—"}/{r.MDL2 || "—"} {r.MDL_GRADER}</td>
                        <td style={styles.td}>{r.HRP || "—"} {r.HRP_GRADER}</td>
                        <td style={styles.td}>{r.SDC || "—"} {r.SDC_GRADER}</td>
                        <td style={styles.td}>{r.PLK || "—"} {r.PLK_GRADER}</td>
                        <td style={styles.td}>{r["2MR"] || "—"} {r.TMR_GRADER}</td>
                        <td style={styles.td}>{r.TOTAL_SCORE || "—"} / {r.STATUS || "—"}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </section>

            {adminSelectedRecord ? (
              <section style={{ ...styles.card, gridColumn: "1 / -1" }}>
                <h2 style={styles.sectionTitle}>Admin correction for selected RN {adminSelectedRecord.RN}</h2>
                <div style={styles.actions}>{EVENTS.map((event) => <button key={event.key} type="button" disabled={isDisabled} style={{ ...styles.secondaryButton, ...(isDisabled ? styles.disabled : {}) }} onClick={() => clearEventForAdminSelectedRN(event.key)}>Clear {event.key}</button>)}</div>
              </section>
            ) : null}
          </main>
        )}

        {exportPreview.csv ? (
          <section style={styles.card}>
            <div style={styles.adminHeader}>
              <div><h2 style={styles.sectionTitle}>Export preview</h2><p style={styles.hint}>{exportPreview.title}</p></div>
              <button type="button" style={styles.secondaryButton} onClick={copyExportPreview}>Copy CSV</button>
            </div>
            <textarea style={styles.exportBox} value={exportPreview.csv} readOnly />
          </section>
        ) : null}

        <footer style={styles.footer}>{isDisabled ? "Working... " : ""}{message}</footer>
      </div>
    </div>
  );
}

function Stat({ label, value }) {
  return (
    <div style={styles.stat}>
      <span>{label}</span>
      <strong>{value}</strong>
    </div>
  );
}

const styles = {
  page: { minHeight: "100vh", background: "#f1f5f9", color: "#0f172a", fontFamily: "Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, Segoe UI, sans-serif", padding: 12 },
  shell: { width: "100%", maxWidth: 1180, margin: "0 auto" },
  header: { background: "white", border: "1px solid #e2e8f0", borderRadius: 24, padding: 20, display: "flex", justifyContent: "space-between", gap: 16, boxShadow: "0 10px 30px rgba(15,23,42,0.06)", flexWrap: "wrap" },
  eyebrow: { margin: 0, fontSize: 11, letterSpacing: 1.8, fontWeight: 900, color: "#64748b" },
  title: { margin: "6px 0", fontSize: "clamp(28px, 7vw, 44px)", lineHeight: 1.05, letterSpacing: -1.2 },
  subtitle: { margin: 0, color: "#475569", lineHeight: 1.5 },
  onlineBadge: { background: "#dcfce7", color: "#166534", border: "1px solid #bbf7d0", borderRadius: 999, padding: "8px 12px", fontSize: 12, fontWeight: 900, height: "fit-content", whiteSpace: "nowrap" },
  offlineBadge: { background: "#fee2e2", color: "#991b1b", border: "1px solid #fecaca", borderRadius: 999, padding: "8px 12px", fontSize: 12, fontWeight: 900, height: "fit-content", whiteSpace: "nowrap" },
  modeTabs: { display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginTop: 12 },
  modeButton: { border: "1px solid #cbd5e1", background: "white", borderRadius: 18, padding: 14, fontWeight: 900, fontSize: 16, cursor: "pointer" },
  modeButtonActive: { background: "#0f172a", color: "white", borderColor: "#0f172a" },
  stack: { display: "grid", gridTemplateColumns: "1fr", gap: 12, marginTop: 12 },
  adminGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(280px, 1fr))", gap: 12, marginTop: 12 },
  card: { background: "white", border: "1px solid #e2e8f0", borderRadius: 22, padding: 18, marginTop: 12, boxShadow: "0 8px 24px rgba(15,23,42,0.05)" },
  sectionTitle: { margin: "0 0 12px", fontSize: 20 },
  label: { display: "flex", flexDirection: "column", gap: 8, fontWeight: 900, color: "#334155" },
  hint: { color: "#64748b", lineHeight: 1.5, margin: "8px 0 0" },
  bigInput: { width: "100%", boxSizing: "border-box", border: "1px solid #cbd5e1", borderRadius: 18, padding: "18px 16px", fontSize: 26, fontWeight: 900, outline: "none" },
  rnInput: { width: "100%", boxSizing: "border-box", border: "2px solid #0f172a", borderRadius: 18, padding: "18px 16px", fontSize: 30, fontWeight: 900, outline: "none" },
  eventGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(72px, 1fr))", gap: 8 },
  eventButton: { border: "1px solid #cbd5e1", background: "white", borderRadius: 18, padding: 12, cursor: "pointer", display: "flex", flexDirection: "column", gap: 4, alignItems: "center" },
  eventButtonActive: { background: "#0f172a", color: "white", borderColor: "#0f172a" },
  soldierBox: { marginTop: 12, background: "#f8fafc", border: "1px solid #e2e8f0", borderRadius: 18, padding: 14, display: "flex", flexDirection: "column", gap: 6 },
  formStack: { display: "grid", gap: 14, marginBottom: 16 },
  primaryButton: { width: "100%", boxSizing: "border-box", border: "none", background: "#0f172a", color: "white", borderRadius: 18, padding: "18px 16px", fontWeight: 900, fontSize: 18, cursor: "pointer", textAlign: "center", display: "block" },
  secondaryButton: { border: "1px solid #cbd5e1", background: "white", color: "#0f172a", borderRadius: 16, padding: "12px 14px", fontWeight: 800, cursor: "pointer" },
  disabled: { opacity: 0.55, cursor: "not-allowed", pointerEvents: "none" },
  statusGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(90px, 1fr))", gap: 8 },
  donePill: { background: "#dcfce7", color: "#166534", border: "1px solid #bbf7d0", borderRadius: 14, padding: 10, textAlign: "center", fontWeight: 900 },
  missingPill: { background: "#fee2e2", color: "#991b1b", border: "1px solid #fecaca", borderRadius: 14, padding: 10, textAlign: "center", fontWeight: 900 },
  statsGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(120px, 1fr))", gap: 10, marginBottom: 12 },
  stat: { background: "#f8fafc", border: "1px solid #e2e8f0", borderRadius: 16, padding: 14, display: "flex", justifyContent: "space-between", gap: 10 },
  adminHeader: { display: "flex", justifyContent: "space-between", alignItems: "center", gap: 12, flexWrap: "wrap" },
  actions: { display: "flex", gap: 8, flexWrap: "wrap" },
  tableWrap: { overflowX: "auto" },
  table: { width: "100%", borderCollapse: "collapse", minWidth: 840 },
  th: { textAlign: "left", borderBottom: "1px solid #e2e8f0", padding: "10px 8px", color: "#475569", fontSize: 13 },
  td: { borderBottom: "1px solid #f1f5f9", padding: "10px 8px", fontSize: 14 },
  clickableRow: { cursor: "pointer" },
  selectedRow: { background: "#eff6ff", cursor: "pointer" },
  exportBox: { width: "100%", minHeight: 220, boxSizing: "border-box", border: "1px solid #cbd5e1", borderRadius: 16, padding: 14, fontFamily: "ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace", fontSize: 13, lineHeight: 1.5, resize: "vertical", outline: "none" },
  footer: { marginTop: 12, background: "white", border: "1px solid #e2e8f0", borderRadius: 18, padding: 14, color: "#475569", fontWeight: 700 },
};

if (typeof window !== "undefined" && !document.head.querySelector("style[data-aft-responsive]")) {
  const style = document.createElement("style");
  style.setAttribute("data-aft-responsive", "true");
  style.textContent = "body{margin:0}button,input,textarea{font:inherit}@media(max-width:760px){table{font-size:12px}}";
  document.head.appendChild(style);
}
