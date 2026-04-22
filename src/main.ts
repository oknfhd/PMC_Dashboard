import Chart from "chart.js/auto";
import "bootstrap/dist/js/bootstrap.bundle.min.js";
import * as XLSX from "xlsx";

// ============================================================================
// TYPES AND CONSTANTS
// ============================================================================

type Sex = "male" | "female" | "unknown";
type SexFilter = Sex | "all";


type Client = {
  date: string;
  sex: SexFilter;
  age: number;
  address: string;
  occupation: string;
  education: string;
  sourceFileId: string;
  sourceFileName: string;
};

type Filters = {
  // Date range filters (from/to)
  fromYear: string;
  fromMonth: string;
  toYear: string;
  toMonth: string;

  // Legacy single selection (kept for backward compatibility)
  year: string;
  month: string;

  sex: SexFilter;
  ageGroup: string;
  occupation: string;
  education: string;
  address: string;

  occupationList: string[],
  educationList: string[],
  addressList: string[]
};

const MONTH_NAMES = [
  "January", "February", "March", "April", "May", "June",
  "July", "August", "September", "October", "November", "December"
];

const AGE_GROUPS = ["0-17", "18-25", "26-40", "41-59", "60+"];

type ImportedFile = {
  id: string;
  name: string;
  size: number;
  lastModified: number;
};

const STORAGE_CLIENTS_KEY = "pmc_dashboard_clients_v2";
const STORAGE_FILTERS_KEY = "pmc_dashboard_filters_v2";
const STORAGE_FILES_KEY = "pmc_dashboard_files_v1";

// ============================================================================
// GLOBAL STATE
// ============================================================================

let allClients: Client[] = [];
let importedFiles: ImportedFile[] = [];

let filters: Filters = {
  fromYear: "all",
  fromMonth: "all",
  toYear: "all",
  toMonth: "all",
  year: "all",
  month: "all",
  sex: "all",
  ageGroup: "all",
  occupation: "all",
  education: "all",
  address: "all",

  occupationList: [],
  educationList: [],
  addressList: []
};

let sexChart: Chart;
let ageChart: Chart;
let occupationChart: Chart;
let educationChart: Chart;
let addressChart: Chart;
let monthlyChart: Chart;
let yearlyChart: Chart;

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

function setTextById(id: string, text: string): void {
  const el = document.getElementById(id);
  if (el) el.innerText = text;
}

function normalize(value: string): string {
  return value.trim().toLowerCase();
}

function capitalize(str: string): string {
  return str
    .split(" ")
    .map(s => s.charAt(0).toUpperCase() + s.slice(1))
    .join(" ");
}

function formatDate(date: Date): string {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

function formatYmd(year: number, month: number, day: number): string {
  return `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
}

function parseYmdToParts(ymd: string): { y: number; m: number; d: number } | null {
  const match = String(ymd ?? "").trim().match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return null;
  const y = Number(match[1]);
  const m = Number(match[2]);
  const d = Number(match[3]);
  if (!y || m < 1 || m > 12 || d < 1 || d > 31) return null;
  return { y, m, d };
}

function todayYmdLocal(): string {
  return formatDate(new Date());
}

function monthIndexFromToken(token: string): number | null {
  let t = String(token ?? "").toUpperCase().replace(/[^A-Z]/g, "");
  if (!t) return null;
  if (t.startsWith("SEPT")) t = "SEP";
  const key = t.slice(0, 3);
  const months: Record<string, number> = {
    JAN: 1, FEB: 2, MAR: 3, APR: 4, MAY: 5, JUN: 6,
    JUL: 7, AUG: 8, SEP: 9, OCT: 10, NOV: 11, DEC: 12,
  };
  return months[key] ?? null;
}

function normalizeTwoDigitYear(year: number): number {
  if (year >= 100) return year;
  return year >= 70 ? 1900 + year : 2000 + year;
}

function parseDateStringToYmd(input: string): string {
  const raw = String(input ?? "").replace(/\u00A0/g, " ").trim();
  if (!raw) return "";

  const iso = raw.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
  if (iso) {
    const y = Number(iso[1]);
    const m = Number(iso[2]);
    const d = Number(iso[3]);
    if (y && m >= 1 && m <= 12 && d >= 1 && d <= 31) return formatYmd(y, m, d);
  }

  const numeric = raw.match(/^(\d{1,2})\s*[\/\-]\s*(\d{1,2})\s*[\/\-]\s*(\d{2,4})$/);
  if (numeric) {
    const a = Number(numeric[1]);
    const b = Number(numeric[2]);
    let y = Number(numeric[3]);
    if (!a || !b || !y) return "";
    y = normalizeTwoDigitYear(y);

    let m = a;
    let d = b;
    if (a > 12 && b <= 12) {
      d = a;
      m = b;
    }

    const dt = new Date(Date.UTC(y, m - 1, d));
    if (dt.getUTCFullYear() !== y || (dt.getUTCMonth() + 1) !== m || dt.getUTCDate() !== d) return "";
    return formatYmd(y, m, d);
  }

  const monthMatch = raw.match(/^\s*([A-Za-z]{3,})\.?\s*(\d{1,2})\s*,?\s*(\d{2,4})\s*$/);
  if (monthMatch) {
    const m = monthIndexFromToken(monthMatch[1]);
    const d = Number(monthMatch[2]);
    let y = Number(monthMatch[3]);
    if (!m || !d || !y) return "";
    y = normalizeTwoDigitYear(y);
    const dt = new Date(Date.UTC(y, m - 1, d));
    if (dt.getUTCFullYear() !== y || (dt.getUTCMonth() + 1) !== m || dt.getUTCDate() !== d) return "";
    return formatYmd(y, m, d);
  }

  // Handles cases like "SEPT. 26,2024" after inserting spacing between letters and digits.
  const cleaned = raw
    .replace(/([A-Za-z])(\d)/g, "$1 $2")
    .replace(/(\d)([A-Za-z])/g, "$1 $2")
    .replace(/[.,]/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  const parts = cleaned.split(" ").filter(Boolean);
  if (parts.length >= 3) {
    const m = monthIndexFromToken(parts[0]);
    const d = Number(parts[1]);
    let y = Number(parts[2]);
    if (m && d && y) {
      y = normalizeTwoDigitYear(y);
      const dt = new Date(Date.UTC(y, m - 1, d));
      if (dt.getUTCFullYear() !== y || (dt.getUTCMonth() + 1) !== m || dt.getUTCDate() !== d) return "";
      return formatYmd(y, m, d);
    }
  }

  return "";
}

function computeAgeYears(dobYmd: string, refYmd: string): number {
  const dob = parseYmdToParts(dobYmd);
  const ref = parseYmdToParts(refYmd);
  if (!dob || !ref) return 0;

  let age = ref.y - dob.y;
  if (ref.m < dob.m || (ref.m === dob.m && ref.d < dob.d)) age -= 1;
  return age < 0 ? 0 : age;
}

function saveStateToStorage(): void {
  try {
    localStorage.setItem(STORAGE_CLIENTS_KEY, JSON.stringify(allClients));
    localStorage.setItem(STORAGE_FILTERS_KEY, JSON.stringify(filters));
    localStorage.setItem(STORAGE_FILES_KEY, JSON.stringify(importedFiles));
  } catch (err) {
    console.warn("[storage] failed to save state", err);
  }
}

function loadStateFromStorage(): void {
  try {
    const rawClients = localStorage.getItem(STORAGE_CLIENTS_KEY) || sessionStorage.getItem("pmc_dashboard_clients_v1");
    if (rawClients) {
      const parsed = JSON.parse(rawClients) as any[];
      if (Array.isArray(parsed)) {
        allClients = parsed
          .filter(Boolean)
          .map((c) => ({
            date: String(c.date ?? ""),
            sex: normalizeSexValue(String(c.sex ?? "")),
            age: Number(c.age ?? 0) || 0,
            address: String(c.address ?? ""),
            occupation: String(c.occupation ?? ""),
            education: String(c.education ?? ""),
            sourceFileId: String(c.sourceFileId ?? `legacy:${String(c.sourceFileName ?? "Imported Data")}`),
            sourceFileName: String(c.sourceFileName ?? "Imported Data"),
          }));
      }
    }

    const rawFilters = localStorage.getItem(STORAGE_FILTERS_KEY) || sessionStorage.getItem("pmc_dashboard_filters_v1");
    if (rawFilters) {
      const parsed = JSON.parse(rawFilters) as Partial<Filters>;
      if (parsed && typeof parsed === "object") {
        filters = {
          fromYear: parsed.fromYear ?? "all",
          fromMonth: parsed.fromMonth ?? "all",
          toYear: parsed.toYear ?? "all",
          toMonth: parsed.toMonth ?? "all",
          year: parsed.year ?? "all",
          month: parsed.month ?? "all",
          sex: parsed.sex ?? "all",
          ageGroup: parsed.ageGroup ?? "all",
          occupation: parsed.occupation ?? "all",
          education: parsed.education ?? "all",
          address: parsed.address ?? "all",

          occupationList: parsed.occupationList ?? [],
          educationList: parsed.educationList ?? [],
          addressList: parsed.addressList ?? []
        };

      }
    }

    const rawFiles = localStorage.getItem(STORAGE_FILES_KEY);
    if (rawFiles) {
      const parsed = JSON.parse(rawFiles) as ImportedFile[];
      if (Array.isArray(parsed)) importedFiles = parsed;
    } else {
      // Backfill file list from any previously stored clients (pre multi-file support)
      if (allClients.length > 0) {
        const uniqueByName = new Map<string, ImportedFile>();
        allClients.forEach((c) => {
          if (!c?.sourceFileName) return;
          if (!uniqueByName.has(c.sourceFileName)) {
            uniqueByName.set(c.sourceFileName, {
              id: c.sourceFileId || `legacy:${c.sourceFileName}`,
              name: c.sourceFileName,
              size: 0,
              lastModified: 0,
            });
          }
        });
        importedFiles = Array.from(uniqueByName.values());
      }
    }

    // Migrate any old state into localStorage for persistence across app restarts.
    saveStateToStorage();
  } catch (err) {
    console.warn("[storage] failed to load state", err);
  }
}

// ============================================================================
// EXCEL FILE IMPORT & DATA DETECTION
// ============================================================================

function parseSheetToClients(
  workbook: XLSX.WorkBook,
  sheetName: string,
  sourceFile: ImportedFile,
): Client[] {
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) return [];

  const matrix = XLSX.utils.sheet_to_json<any[]>(sheet, { header: 1, defval: "" });
  const rowObjects = matrixToRowObjects(matrix);
  if (rowObjects.length === 0) return [];

  const fixedRows = fillMergedCells(rowObjects);
  const parsed = mapToClients(fixedRows, sourceFile);

  console.info(`[import] sheet="${sheetName}" rows=${parsed.length}`);
  return parsed;
}

function buildFileId(file: Pick<File, "name" | "size" | "lastModified">): string {
  return `${file.name}::${file.size}::${file.lastModified}`;
}

async function parseFileToClients(file: File, sourceFile: ImportedFile): Promise<Client[]> {
  const data = new Uint8Array(await file.arrayBuffer());
  const workbook = XLSX.read(data, { type: "array" });

  // Import all sheets. Sheets without a recognizable header row will be skipped.
  const parsed = workbook.SheetNames.flatMap((sheetName) =>
    parseSheetToClients(workbook, sheetName, sourceFile),
  );

  console.info(`[import] file="${file.name}" total_sheets=${workbook.SheetNames.length} total_rows=${parsed.length}`);
  return parsed;
}

async function handleFiles(fileList: FileList): Promise<void> {
  const files = Array.from(fileList);
  if (files.length === 0) return;

  const newImportedFiles: ImportedFile[] = [];
  const newClients: Client[] = [];

  for (const file of files) {
    const fileId = buildFileId(file);
    const exists = importedFiles.some((f) => f.id === fileId);
    if (exists) {
      console.info(`[import] skipped duplicate file="${file.name}"`);
      continue;
    }

    const meta: ImportedFile = {
      id: fileId,
      name: file.name,
      size: file.size,
      lastModified: file.lastModified,
    };

    const parsed = await parseFileToClients(file, meta);
    newImportedFiles.push(meta);
    newClients.push(...parsed);
  }

  if (newImportedFiles.length === 0) return;

  importedFiles = [...importedFiles, ...newImportedFiles];
  allClients = [...allClients, ...newClients];

  resetFilters();
  renderYearDropdown();
  renderMonthDropdown();
  renderImportedFilesMenu();
  applyFilters();
}

function normalizeHeaderCell(value: any): string {
  return String(value ?? "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function matrixToRowObjects(matrix: any[][]): any[] {
  if (!Array.isArray(matrix) || matrix.length === 0) return [];

  const expected = [
    "date",
    "date of birth",
    "dob",
    "birthdate",
    "sex",
    "age",
    "address",
    "occupation",
    "name of couple",
    "name",
    "highest educational attainment",
    "education",
    "no.",
    "no",
  ];

  const headerIndex = matrix.findIndex((row) => {
    if (!Array.isArray(row)) return false;
    const cells = row.map(normalizeHeaderCell).filter(Boolean);
    const hits = expected.reduce((acc, key) => (cells.includes(key) ? acc + 1 : acc), 0);
    return hits >= 2;
  });

  if (headerIndex < 0) {
    console.warn("No header row detected; import will likely be empty.", { sample: matrix[0] });
    return [];
  }

  const headerRow = matrix[headerIndex] ?? [];
  const seen = new Set<string>();
  const headers = headerRow.map((cell, i) => {
    const base = normalizeHeaderCell(cell) || `__col${i + 1}`;
    let key = base;
    let suffix = 2;
    while (seen.has(key)) {
      key = `${base}_${suffix}`;
      suffix += 1;
    }
    seen.add(key);
    return key;
  });

  const out: any[] = [];
  for (let r = headerIndex + 1; r < matrix.length; r += 1) {
    const row = matrix[r];
    if (!Array.isArray(row)) continue;

    const obj: any = {};
    let hasAnyValue = false;
    for (let c = 0; c < headers.length; c += 1) {
      const v = row[c];
      obj[headers[c]] = v ?? "";
      if (String(v ?? "").trim() !== "") hasAnyValue = true;
    }

    if (hasAnyValue) out.push(obj);
  }

  return out;
}

function fillMergedCells(rows: any[]): any[] {
  let lastDate: any = null;
  let lastNo: any = null;

  return rows.map(row => {
    const cleaned: any = {};

    Object.keys(row).forEach(key => {
      const k = key.trim().toLowerCase();
      cleaned[k] = row[key];
    });

    // Preserve last date/no for merged cells
    if (cleaned["date"]) lastDate = cleaned["date"];
    if (cleaned["no."]) lastNo = cleaned["no."];

    return {
      ...cleaned,
      date: cleaned["date"] || lastDate,
      "no.": cleaned["no."] || lastNo
    };
  });
}

function mapToClients(rows: any[], sourceFile: ImportedFile): Client[] {
  return rows
    .map((r, idx) => {
      const row = normalizeRowKeys(r);

      const name = row["name of couple"] || row["name"];
      console.log(`Row ${idx}:`, { name, date: row["date"], hasName: !!name });

      if (!name) return null;

      const dateYmd = parseExcelDate(row["date"] || "");
      const dobYmd = parseExcelDate(row["date of birth"] ?? row["dob"] ?? row["birthdate"] ?? "");

      let age = Number(row["age"]) || 0;
      if ((!age || age <= 0) && dobYmd) {
        age = computeAgeYears(dobYmd, dateYmd || todayYmdLocal());
      }

      const parsed = {
        date: dateYmd,
        sex: normalizeSexValue(String(row["sex"] ?? "")),
        age,
        address: normalize(String(row["address"] ?? "")),
        occupation: normalize(String(row["occupation"] ?? "")),
        education: normalize(String(row["highest educational attainment"] ?? row["education"] ?? "")),
        sourceFileId: sourceFile.id,
        sourceFileName: sourceFile.name,
      };

      console.log(`  Parsed:`, parsed);
      return parsed;
    })
    .filter(c => c !== null && c.date) as Client[];
}

function normalizeRowKeys(row: any): any {
  const cleaned: any = {};
  Object.keys(row).forEach(key => {
    cleaned[key.trim().toLowerCase()] = row[key];
  });
  console.log("  Row keys:", Object.keys(cleaned));
  return cleaned;
}

function parseExcelDate(value: any): string {
  if (!value) {
    console.log("  parseExcelDate: empty value");
    return "";
  }

  if (value instanceof Date) {
    const result = formatDate(value);
    console.log(`  parseExcelDate: Date â†’ ${result}`);
    return result;
  }

  // Handle Excel serial numbers
  if (typeof value === "number") {
    const excelEpochUtc = Date.UTC(1899, 11, 30);
    const date = new Date(excelEpochUtc + Math.floor(value) * 86400000);
    const result = formatYmd(date.getUTCFullYear(), date.getUTCMonth() + 1, date.getUTCDate());
    console.log(`  parseExcelDate: ${value} → ${result}`);
    return result;
  }

  // Handle string dates
  if (typeof value === "string") {
    const result = parseDateStringToYmd(value);

    console.log(`  parseExcelDate: "${value}" → ${result}`);
    return result;
  }

  console.log("  parseExcelDate: unknown type");
  return "";
}

// ============================================================================
// DATA TRANSFORMATION FUNCTIONS
// ============================================================================

function getAgeGroup(age: number): string {
  if (age <= 17) return "0-17";
  if (age <= 25) return "18-25";
  if (age <= 40) return "26-40";
  if (age <= 59) return "41-59";
  return "60+";
}

function getMonthName(monthNumber: string): string {
  const index = parseInt(monthNumber) - 1;
  return index >= 0 && index < 12 ? MONTH_NAMES[index] : monthNumber;
}

function getYears(data: Client[]): number[] {
  const years = data
    .map((c) => parseYmdToParts(c.date)?.y)
    .filter((y): y is number => typeof y === "number");
  return [...new Set(years)].sort((a, b) => b - a);
}

function getMonths(data: Client[], selectedYear: string): number[] {
  let filtered = data;

  if (selectedYear !== "all") {
    filtered = data.filter((c) => {
      const parts = parseYmdToParts(c.date);
      return !!parts && parts.y.toString() === selectedYear;
    });
  }

  return [
    ...new Set(
      filtered
        .map((c) => parseYmdToParts(c.date)?.m)
        .filter((m): m is number => typeof m === "number"),
    )
  ].sort((a, b) => a - b);
}

// ============================================================================
// FILTER LOGIC
// ============================================================================

function resetFilters(): void {
  filters = {
    fromYear: "all",
    fromMonth: "all",
    toYear: "all",
    toMonth: "all",
    year: "all",
    month: "all",
    sex: "all",
    ageGroup: "all",
    occupation: "all",
    education: "all",
    address: "all",

    occupationList: [],
    educationList: [],
    addressList: []
  };
  saveStateToStorage();
}

function filterClientsByFilters(source: Client[], f: Filters): Client[] {
  let filtered = source;

  // Date range filtering using from/to year/month
  const hasFromDate = f.fromYear !== "all" || f.fromMonth !== "all";
  const hasToDate = f.toYear !== "all" || f.toMonth !== "all";

  if (hasFromDate || hasToDate) {
    filtered = filtered.filter((c) => {
      const parts = parseYmdToParts(c.date);
      if (!parts) return false;

      const dateYear = parts.y;
      const dateMonth = parts.m;

      // Check from date (inclusive)
      if (f.fromYear !== "all") {
        const fromYear = parseInt(f.fromYear);
        if (f.fromMonth !== "all") {
          const fromMonth = parseInt(f.fromMonth);
          // Compare year-month: date must be >= fromYear-fromMonth
          if (dateYear < fromYear) return false;
          if (dateYear === fromYear && dateMonth < fromMonth) return false;
        } else {
          // Only year filter: date must be >= fromYear-01
          if (dateYear < fromYear) return false;
        }
      }

      // Check to date (inclusive)
      if (f.toYear !== "all") {
        const toYear = parseInt(f.toYear);
        if (f.toMonth !== "all") {
          const toMonth = parseInt(f.toMonth);
          // Compare year-month: date must be <= toYear-toMonth
          if (dateYear > toYear) return false;
          if (dateYear === toYear && dateMonth > toMonth) return false;
        } else {
          // Only year filter: date must be <= toYear-12
          if (dateYear > toYear) return false;
        }
      }

      return true;
    });
  }

  // Legacy single year/month filtering (for backward compatibility)
  if (f.year !== "all" && !hasFromDate) {
    filtered = filtered.filter((c) => {
      const parts = parseYmdToParts(c.date);
      return !!parts && parts.y.toString() === f.year;
    });
  }

  if (f.month !== "all" && !hasFromDate && !hasToDate && f.year === "all") {
    filtered = filtered.filter((c) => {
      const parts = parseYmdToParts(c.date);
      return !!parts && parts.m.toString() === f.month;
    });
  }

  if (f.sex !== "all") {
    filtered = filtered.filter((c) => c.sex === f.sex);
  }

  if (f.ageGroup !== "all") {
    filtered = filtered.filter((c) => getAgeGroup(c.age) === f.ageGroup);
  }

  if (f.occupationList && f.occupationList.length) {
    filtered = filtered.filter(c => f.occupationList!.includes(c.occupation));
  } else if (f.occupation !== "all") {
    filtered = filtered.filter(c => c.occupation === f.occupation);
  }

  if (f.educationList && f.educationList.length) {
    filtered = filtered.filter(c => f.educationList!.includes(c.education));
  } else if (f.education !== "all") {
    filtered = filtered.filter(c => c.education === f.education);
  }

  if (f.addressList && f.addressList.length) {
    filtered = filtered.filter(c => f.addressList!.includes(c.address));
  } else if (f.address !== "all") {
    filtered = filtered.filter(c => c.address === f.address);
  }

  return filtered;
}

function applyFilters(): void {
  const filtered = filterClientsByFilters(allClients, filters);

  console.log("Active filters:", filters);
  console.log("Filtered data:", filtered);
  displayActiveFilters();
  updateUI(filtered);
  saveStateToStorage();
}

function displayActiveFilters(): void {
  const displayEl = document.getElementById("active_filters");
  if (!displayEl) return;

  const activeFilters: string[] = [];

  // Date range filters (from/to)
  const hasFromDate = filters.fromYear !== "all" || filters.fromMonth !== "all";
  const hasToDate = filters.toYear !== "all" || filters.toMonth !== "all";

  if (hasFromDate || hasToDate) {
    let dateRange = "Date: ";
    if (filters.fromYear !== "all") {
      dateRange += filters.fromYear;
      if (filters.fromMonth !== "all") {
        dateRange += "-" + getMonthName(filters.fromMonth).substring(0, 3);
      }
    } else if (filters.fromMonth !== "all") {
      dateRange += getMonthName(filters.fromMonth);
    }

    if (hasToDate) {
      dateRange += " to ";
      if (filters.toYear !== "all") {
        dateRange += filters.toYear;
        if (filters.toMonth !== "all") {
          dateRange += "-" + getMonthName(filters.toMonth).substring(0, 3);
        }
      } else if (filters.toMonth !== "all") {
        dateRange += getMonthName(filters.toMonth);
      }
    }

    activeFilters.push(dateRange);
  }

  if (filters.sex !== "all") activeFilters.push(`Sex: ${capitalize(filters.sex)}`);
  if (filters.ageGroup !== "all") activeFilters.push(`Age Group: ${filters.ageGroup}`);
  if (filters.occupation !== "all") activeFilters.push(`Occupation: ${capitalize(filters.occupation)}`);
  if (filters.education !== "all") activeFilters.push(`Education: ${capitalize(filters.education)}`);
  if (filters.address !== "all") activeFilters.push(`Address: ${capitalize(filters.address)}`);

  const displayText = activeFilters.length > 0 ? activeFilters.join(" | ") : "No filters applied";
  displayEl.innerHTML = displayText;
}

// ============================================================================
// UI UPDATES
// ============================================================================

function updateUI(data: Client[]): void {
  // Update active filter labels for display
  const hasDateRange = filters.fromYear !== "all" || filters.fromMonth !== "all" || filters.toYear !== "all" || filters.toMonth !== "all";
  if (hasDateRange) {
    let yearLabel = "";
    if (filters.fromYear !== "all" && filters.toYear !== "all" && filters.fromYear === filters.toYear) {
      yearLabel = filters.fromYear;
    } else if (filters.fromYear !== "all" || filters.toYear !== "all") {
      yearLabel = (filters.fromYear !== "all" ? filters.fromYear : "...") + "-" + (filters.toYear !== "all" ? filters.toYear : "...");
    } else {
      yearLabel = "All Years";
    }
    setTextById("active_year", yearLabel);
  } else {
    setTextById("active_year", filters.year === "all" ? "All Years" : filters.year);
  }
  setTextById("active_month", filters.month === "all" ? "Month" : getMonthName(filters.month));

  const totalClients = allClients.length;
  setTextById("total_clients", totalClients.toString());

  // Total clients in selected timeline
  const timelineClients = data.length;
  setTextById("total_clients_year", timelineClients.toString());

  // --- Average Monthly Clients ---
  const monthSet = new Set<string>();

  data.forEach(c => {
    const parts = parseYmdToParts(c.date);
    if (!parts) return;

    const key = `${parts.y}-${parts.m}`;
    monthSet.add(key);
  });

  const totalMonths = monthSet.size;

  // avoid division by zero
  const avgMonthly = totalMonths > 0
    ? (timelineClients / totalMonths)
    : 0;

  // round (no decimals or 1 decimal if you prefer)
  setTextById("total_clients_month", avgMonthly.toFixed(1));


  // 🔥 NEW: update labels dynamically
  updateDashboardLabels();

  renderCharts(data);
  renderSummary(data);
  renderImportedFilesMenu();
}

function updateDashboardLabels(): void {
  const yearLabel = document.getElementById("year_filtered");
  const monthLabel = document.getElementById("month_filtered");

  if (yearLabel) {
    yearLabel.innerText =
      filters.year === "all" ? "Year" : filters.year;
  }

  if (monthLabel) {
    monthLabel.innerText =
      filters.month === "all" ? "Month" : getMonthName(filters.month);
  }
}


function renderCharts(data: Client[]): void {
  renderSexChart(data);
  renderAgeChart(data);
  renderOccupationChart(data);
  renderEducationChart(data);
  renderAddressChart(data);
  renderMonthlyChart(data);
  renderYearlyChart(allClients);
}

// ============================================================================
// DROPDOWN RENDERING
// ============================================================================

function renderYearDropdownFrom(): void {
  const menu = document.getElementById("year_menu_from");
  if (!menu) return;
  menu.innerHTML = "";

  menu.innerHTML += `<li><a class="dropdown-item" data-year-from="all">All Years</a></li>`;

  getYears(allClients).forEach(year => {
    menu.innerHTML += `<li><a class="dropdown-item" data-year-from="${year}">${year}</a></li>`;
  });
}

function renderYearDropdownTo(): void {
  const menu = document.getElementById("year_menu_to");
  if (!menu) return;
  menu.innerHTML = "";

  menu.innerHTML += `<li><a class="dropdown-item" data-year-to="all">All Years</a></li>`;

  getYears(allClients).forEach(year => {
    menu.innerHTML += `<li><a class="dropdown-item" data-year-to="${year}">${year}</a></li>`;
  });
}

function renderMonthDropdownFrom(): void {
  const menu = document.getElementById("month_menu_from");
  if (!menu) return;
  menu.innerHTML = "";

  menu.innerHTML += `<li><a class="dropdown-item" data-month-from="all">Month</a></li>`;

  const availableMonths = getMonths(allClients, filters.fromYear);

  availableMonths.forEach(m => {
    menu.innerHTML += `<li><a class="dropdown-item" data-month-from="${m}">${MONTH_NAMES[m - 1]}</a></li>`;
  });
}

function renderMonthDropdownTo(): void {
  const menu = document.getElementById("month_menu_to");
  if (!menu) return;
  menu.innerHTML = "";

  menu.innerHTML += `<li><a class="dropdown-item" data-month-to="all">Month</a></li>`;

  const availableMonths = getMonths(allClients, filters.toYear);

  availableMonths.forEach(m => {
    menu.innerHTML += `<li><a class="dropdown-item" data-month-to="${m}">${MONTH_NAMES[m - 1]}</a></li>`;
  });
}

// Legacy functions for backward compatibility
function renderYearDropdown(): void {
  renderYearDropdownFrom();
  renderYearDropdownTo();
}

function renderMonthDropdown(): void {
  renderMonthDropdownFrom();
  renderMonthDropdownTo();
}

// ============================================================================
// CHART RENDERING
// ============================================================================

function renderSexChart(data: Client[]): void {
  const ctx = document.getElementById("chart_sex") as HTMLCanvasElement | null;
  if (!ctx) return;

  const counts = { male: 0, female: 0, unknown: 0 };

  data.forEach(c => {
    const sex = normalizeSexValue(c.sex);
    counts[sex]++;
  });

  const labels = ["Male", "Female"];
  // Use actual counts for display, but ensure chart always has data to render
  const displayValues = [counts.male, counts.female];

  // Define colors for Male and Female
  const backgroundColors = ["#36A2EB", "#FF6384"];
  const borderColors = ["#2c7ebb", "#cc4f69"];

  // Calculate opacity based on filter state - gray out the non-selected segment
  const isMaleFiltered = filters.sex === "male";
  const isFemaleFiltered = filters.sex === "female";

  const bgColors = [
    isMaleFiltered && counts.male > 0 ? backgroundColors[0] : (counts.male === 0 ? '#cccccc' : backgroundColors[0]),
    isFemaleFiltered && counts.female > 0 ? backgroundColors[1] : (counts.female === 0 ? '#cccccc' : backgroundColors[1])
  ];

  if (sexChart) sexChart.destroy();

  sexChart = new Chart(ctx, {
    type: "doughnut",
    data: {
      labels,
      datasets: [{
        data: displayValues,
        backgroundColor: bgColors,
        borderColor: borderColors,
        borderWidth: 2
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      cutout: "50%",

      plugins: {
        legend: {
          position: "right",
          labels: {
            generateLabels: function () {
              return [
                {
                  text: `Male`,
                  fontColor: '#1a1a1a',
                  fillStyle: counts.male === 0 ? '#cccccc' : backgroundColors[0],
                  strokeStyle: borderColors[0],
                  lineWidth: 2,
                  hidden: false,
                  index: 0
                },
                {
                  text: `Female`,
                  fontColor: '#1a1a1a',
                  fillStyle: counts.female === 0 ? '#cccccc' : backgroundColors[1],
                  strokeStyle: borderColors[1],
                  lineWidth: 2,
                  hidden: false,
                  index: 1
                }
              ];
            }
          }
        },
        tooltip: {
          callbacks: {
            label: function (context) {
              const total = counts.male + counts.female;
              const percentage = total > 0 ? ((context.parsed as number) / total * 100).toFixed(1) : '0.0';
              return `${context.label}: ${context.parsed} (${percentage}%)`;
            }
          }
        }
      },
      onClick: (_, elements) => {
        if (!elements.length) return;

        const index = elements[0].index;

        const selected = index === 0 ? "male" : "female";

        filters.sex = filters.sex === selected ? "all" : selected;

        applyFilters();
      }
    }
  });
}



function renderAgeChart(data: Client[]): void {
  const ctx = document.getElementById("chart_age") as HTMLCanvasElement | null;
  if (!ctx) return;

  const counts: Record<string, number> = {};

  data.forEach(c => {
    if (c.age === null || c.age === undefined || isNaN(c.age) || c.age <= 0) return;

    const group = getAgeGroup(c.age);
    counts[group] = (counts[group] || 0) + 1;
  });

  const labels = AGE_GROUPS;
  const values = labels.map(l => counts[l] || 0);

  if (ageChart) ageChart.destroy();

  ageChart = new Chart(ctx, {
    type: "bar",
    data: { labels, datasets: [{ data: values }] },
    options: {
      responsive: true,
      scales: {
        y: { ticks: { callback: (value) => Number.isInteger(value) ? value : null } }
      },
      plugins: { legend: { display: false } },
      onClick: (_, elements) => {
        if (elements.length) {
          const selected = labels[elements[0].index];
          filters.ageGroup = filters.ageGroup === selected ? "all" : selected;
          applyFilters();
        }
      }
    }
  });
}


// height
function setDynamicHeight(canvasId: string, itemCount: number) {
  const canvas = document.getElementById(canvasId) as HTMLCanvasElement;
  if (!canvas) return;

  const rowHeight = 35; // adjust (30–45 ideal)
  const minHeight = 250;

  const height = Math.max(minHeight, itemCount * rowHeight);
  canvas.style.height = `${height}px`;
}
//height --   

function renderOccupationChart(data: Client[]): void {
  const ctx = document.getElementById("chart_occupation") as HTMLCanvasElement | null;
  if (!ctx) return;

  const counts: Record<string, number> = {};

  data.forEach(c => {
    const key = normalize(c.occupation);
    if (!key) return;
    counts[key] = (counts[key] || 0) + 1;
  });


  const sorted = Object.entries(counts).sort((a, b) => b[1] - a[1]);
  const labels = sorted.map(entry => entry[0]);
  const values = sorted.map(entry => entry[1]);

  const combined = labels.map((label, i) => ({
    label,
    value: values[i]
  }));

  combined.sort((a, b) => b.value - a.value);

  const top10 = combined.slice(0, 10);
  const rest = combined.slice(10);
  const othersTotal = rest.reduce((sum, item) => sum + item.value, 0);

  if (othersTotal > 0) {
    top10.push({ label: "Others", value: othersTotal });
  }

  const finalLabels = top10.map(i => i.label);
  const finalValues = top10.map(i => i.value);
  setDynamicHeight("chart_occupation", finalLabels.length);

  if (occupationChart) occupationChart.destroy();

  occupationChart = new Chart(ctx, {
    type: "bar",
    data: {
      labels: finalLabels.map(capitalize),
      datasets: [{ data: finalValues }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      scales: {
        x: { ticks: { callback: (value) => Number.isInteger(value) ? value : null } }
      },
      plugins: { legend: { display: false } },
      indexAxis: "y",
      onClick: (_, elements) => {
        const othersList = rest.map(i => i.label);
        if (elements.length) {
          const selected = finalLabels[elements[0].index];

          // Handle "Others" differently (optional)
          if (selected === "Others") {
            filters.occupation = "all";
            filters.occupationList = othersList;
          } else {
            filters.occupationList = [];
            filters.occupation =
              filters.occupation === selected ? "all" : selected;
          }


          applyFilters();
        }
      }

    }
  });
}

function renderEducationChart(data: Client[]): void {
  const ctx = document.getElementById("chart_education") as HTMLCanvasElement | null;
  if (!ctx) return;

  const counts: Record<string, number> = {};

  data.forEach(c => {
    const key = normalize(c.education);
    if (!key) return;
    counts[key] = (counts[key] || 0) + 1;
  });

  const combined = Object.entries(counts).map(([label, value]) => ({ label, value }));
  combined.sort((a, b) => b.value - a.value);

  const top10 = combined.slice(0, 10);
  const rest = combined.slice(10);

  const othersTotal = rest.reduce((sum, item) => sum + item.value, 0);

  if (othersTotal > 0) {
    top10.push({ label: "Others", value: othersTotal });
  }

  const finalLabels = top10.map(i => i.label);
  const finalValues = top10.map(i => i.value);
  const othersList = rest.map(i => i.label);
  setDynamicHeight("chart_education", finalLabels.length);

  if (educationChart) educationChart.destroy();

  educationChart = new Chart(ctx, {
    type: "bar",
    data: {
      labels: finalLabels.map(capitalize),
      datasets: [{ data: finalValues }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      indexAxis: "y",
      plugins: { legend: { display: false } },
      scales: {
        x: { ticks: { callback: (value) => Number.isInteger(value) ? value : null } }
      },
      onClick: (_, elements) => {
        if (elements.length) {
          const selected = finalLabels[elements[0].index];

          if (selected === "Others") {
            filters.education = "all";
            filters.educationList = othersList;
          } else {
            filters.educationList = [];
            filters.education =
              filters.education === selected ? "all" : selected;
          }

          applyFilters();
        }
      }
    }
  });
}

function renderAddressChart(data: Client[]): void {
  const ctx = document.getElementById("chart_address") as HTMLCanvasElement | null;
  if (!ctx) return;

  const counts: Record<string, number> = {};

  data.forEach(c => {
    const key = normalize(c.address);
    if (!key) return;
    counts[key] = (counts[key] || 0) + 1;
  });

  const combined = Object.entries(counts).map(([label, value]) => ({ label, value }));
  combined.sort((a, b) => b.value - a.value);

  const top10 = combined.slice(0, 10);
  const rest = combined.slice(10);

  const othersTotal = rest.reduce((sum, item) => sum + item.value, 0);

  if (othersTotal > 0) {
    top10.push({ label: "Others", value: othersTotal });
  }

  const finalLabels = top10.map(i => i.label);
  const finalValues = top10.map(i => i.value);
  const othersList = rest.map(i => i.label);
  setDynamicHeight("chart_address", finalLabels.length);

  if (addressChart) addressChart.destroy();

  addressChart = new Chart(ctx, {
    type: "bar",
    data: {
      labels: finalLabels.map(capitalize),
      datasets: [{ data: finalValues }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      indexAxis: "y",
      plugins: { legend: { display: false } },
      scales: {
        x: { ticks: { callback: (value) => Number.isInteger(value) ? value : null } }
      },
      onClick: (_, elements) => {
        if (elements.length) {
          const selected = finalLabels[elements[0].index];

          if (selected === "Others") {
            filters.address = "all";
            filters.addressList = othersList;
          } else {
            filters.addressList = [];
            filters.address =
              filters.address === selected ? "all" : selected;
          }

          applyFilters();
        }
      }
    }
  });
}


function renderMonthlyChart(data: Client[]): void {
  const ctx = document.getElementById("dash_monthly") as HTMLCanvasElement | null;
  if (!ctx) return;

  // Check if date range filter is active (from/to year/month)
  const hasDateRange = filters.fromYear !== "all" || filters.fromMonth !== "all" ||
    filters.toYear !== "all" || filters.toMonth !== "all";


  // Count clients by month within the selected date range
  const monthlyCounts: Record<string, number> = {};

  data.forEach(c => {
    const parts = parseYmdToParts(c.date);
    if (!parts) return;

    const yearMonth = `${parts.y}-${String(parts.m).padStart(2, '0')}`;
    monthlyCounts[yearMonth] = (monthlyCounts[yearMonth] || 0) + 1;
  });

  // Sort by year-month and create labels/values
  const sortedKeys = Object.keys(monthlyCounts).sort();
  const labels = sortedKeys.map(key => {
    const [year, month] = key.split('-');
    return `${MONTH_NAMES[parseInt(month) - 1]} ${year}`;
  });
  const values = sortedKeys.map(key => monthlyCounts[key]);

  if (monthlyChart) monthlyChart.destroy();

  monthlyChart = new Chart(ctx, {
    type: "line",
    data: {
      labels,
      datasets: [{ data: values }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: {
        y: {
          ticks: {
            callback: (value) => Number.isInteger(value) ? value : null
          }
        }
      }
    }
  });
}




function renderYearlyChart(data: Client[]): void {
  const ctx = document.getElementById("dash_yearly") as HTMLCanvasElement | null;
  if (!ctx) return;

  const counts: Record<string, number> = {};
  data.forEach(c => {
    const parts = parseYmdToParts(c.date);
    if (!parts) return;
    const year = parts.y.toString();
    counts[year] = (counts[year] || 0) + 1;
  });

  const labels = Object.keys(counts).sort();
  const values = labels.map(l => counts[l]);

  if (yearlyChart) yearlyChart.destroy();

  yearlyChart = new Chart(ctx, {
    type: "line",
    data: { labels, datasets: [{ data: values }] },
    options: {
      responsive: true,
      scales: {
        y: { ticks: { callback: (value) => Number.isInteger(value) ? value : null } }
      },
      plugins: { legend: { display: false } }
    }
  });
}

// ============================================================================
// EVENT LISTENERS
// ============================================================================

document.addEventListener("click", (e) => {
  const target = e.target as HTMLElement;

  // From Year selection
  const yearFromEl = target.closest("[data-year-from]") as HTMLElement | null;
  if (yearFromEl?.dataset.yearFrom) {
    filters.fromYear = yearFromEl.dataset.yearFrom;
    setTextById("year_dropdown_from", filters.fromYear === "all" ? "Year" : filters.fromYear);

    // Reset from month when year changes
    filters.fromMonth = "all";
    setTextById("month_dropdown_from", "Month");

    renderMonthDropdownFrom();
    applyFilters();
  }

  // From Month selection
  const monthFromEl = target.closest("[data-month-from]") as HTMLElement | null;
  if (monthFromEl?.dataset.monthFrom) {
    filters.fromMonth = monthFromEl.dataset.monthFrom;
    setTextById("month_dropdown_from", filters.fromMonth === "all" ? "Month" : monthFromEl.innerText);
    applyFilters();
  }

  // To Year selection
  const yearToEl = target.closest("[data-year-to]") as HTMLElement | null;
  if (yearToEl?.dataset.yearTo) {
    filters.toYear = yearToEl.dataset.yearTo;
    setTextById("year_dropdown", filters.toYear === "all" ? "Year" : filters.toYear);

    // Reset to month when year changes
    filters.toMonth = "all";
    setTextById("month_dropdown", "Month");

    renderMonthDropdownTo();
    applyFilters();
  }

  // To Month selection
  const monthToEl = target.closest("[data-month-to]") as HTMLElement | null;
  if (monthToEl?.dataset.monthTo) {
    filters.toMonth = monthToEl.dataset.monthTo;
    setTextById("month_dropdown", filters.toMonth === "all" ? "Month" : monthToEl.innerText);
    applyFilters();
  }

  // Reset filter button
if (target.id === "reset_filter") {
  resetFilters();

  if (allClients.length > 0) {
    const years = getYears(allClients);

    if (years.length > 0) {
      const earliestYear = years[years.length - 1];
      const latestYear = years[0];

      filters.fromYear = earliestYear.toString();
      filters.toYear = latestYear.toString();

      // reset months
      filters.fromMonth = "all";
      filters.toMonth = "all";

      // ✅ Update UI labels
      setTextById("year_dropdown_from", filters.fromYear);
      setTextById("month_dropdown_from", "Month");

      setTextById("year_dropdown", filters.toYear);
      setTextById("month_dropdown", "Month");
    }
  } else {
    // fallback if no data
    setTextById("year_dropdown_from", "Year");
    setTextById("month_dropdown_from", "Month");
    setTextById("year_dropdown", "Year");
    setTextById("month_dropdown", "Month");
  }

  renderYearDropdown();
  renderMonthDropdown();
  applyFilters();
}


  const deleteEl = target.closest("[data-delete-file]") as HTMLElement | null;
  if (deleteEl?.dataset.deleteFile) {
    e.preventDefault();
    e.stopPropagation();
    deleteImportedFile(deleteEl.dataset.deleteFile);
    return;
  }
});

// ============================================================================
// INITIALIZATION
// ============================================================================

function init(): void {
  loadStateFromStorage();

  // Set default date range: from earliest to latest date
  if (allClients.length > 0 && filters.fromYear === "all" && filters.toYear === "all") {
    const dates = allClients
      .map(c => parseYmdToParts(c.date))
      .filter(Boolean) as { y: number; m: number; d: number }[];

    if (dates.length > 0) {
      // Sort ascending
      dates.sort((a, b) =>
        a.y !== b.y ? a.y - b.y :
          a.m !== b.m ? a.m - b.m :
            a.d - b.d
      );

      const earliest = dates[0];
      const latest = dates[dates.length - 1];

      // ✅ Set FULL range
      filters.fromYear = earliest.y.toString();
      filters.fromMonth = earliest.m.toString();

      filters.toYear = latest.y.toString();
      filters.toMonth = latest.m.toString();

      // ✅ Update UI labels
      setTextById("year_dropdown_from", filters.fromYear);
      setTextById("month_dropdown_from", MONTH_NAMES[earliest.m - 1]);

      setTextById("year_dropdown", filters.toYear);
      setTextById("month_dropdown", MONTH_NAMES[latest.m - 1]);
    }
  }


  renderYearDropdown();
  renderMonthDropdown();
  updateDateDropdownLabels();
  renderImportedFilesMenu();
  applyFilters();
}

function updateDateDropdownLabels(): void {
  // FROM
  setTextById(
    "year_dropdown_from",
    filters.fromYear === "all" ? "Year" : filters.fromYear
  );

  setTextById(
    "month_dropdown_from",
    filters.fromMonth === "all"
      ? "Month"
      : MONTH_NAMES[parseInt(filters.fromMonth) - 1]
  );

  // TO
  setTextById(
    "year_dropdown",
    filters.toYear === "all" ? "Year" : filters.toYear
  );

  setTextById(
    "month_dropdown",
    filters.toMonth === "all"
      ? "Month"
      : MONTH_NAMES[parseInt(filters.toMonth) - 1]
  );
}


window.addEventListener("DOMContentLoaded", () => {
  const importBtn = document.getElementById("import-btn") as HTMLButtonElement | null;

  if (importBtn) {
    importBtn.addEventListener("click", () => {
      const input = document.createElement("input");
      input.type = "file";
      input.accept = ".xlsx, .xls, .csv";
      input.multiple = true;

      input.onchange = (e: any) => {
        const fileList = e.target.files as FileList | null;
        if (fileList && fileList.length > 0) void handleFiles(fileList);
      };

      input.click();
    });
  }

  init();
});

// ============================================================================
// SUMMARY PAGE RENDERING
// ============================================================================

function normalizeSexValue(value: string): Sex {
  const v = normalize(value);

  if (v === "m" || v === "male") return "male";
  if (v === "f" || v === "female" || v === "") return "female";
  return "unknown";
}




function getAllCounts(values: string[]): Array<[string, number]> {
  const counts: Record<string, number> = {};

  values
    .map(v => normalize(String(v ?? "")))
    .filter(Boolean)
    .forEach(v => {
      counts[v] = (counts[v] || 0) + 1;
    });

  return Object.entries(counts)
    .sort((a, b) => (b[1] - a[1]) || a[0].localeCompare(b[0]));
}


function renderCountList(
  listEl: HTMLUListElement | null,
  counts: Array<[string, number]>,
): void {
  if (!listEl) return;
  if (counts.length === 0) {
    listEl.innerHTML = `<li class="list_data">No data</li>`;
    return;
  }

  listEl.innerHTML = counts
    .map(([key, count]) => `<li class="list_name"><span>${capitalize(key)}:</span> <span class="list_data">${count}</span></li>`)
    .join("");
}

function renderSummary(data: Client[]): void {
  const hasSummary =
    !!document.getElementById("female_count") ||
    !!document.getElementById("age_1") ||
    !!document.querySelector(".occupation_list ul") ||
    !!document.querySelector(".education_list ul") ||
    !!document.querySelector(".address_list ul");

  if (!hasSummary) return;

  // Sex
  let female = 0;
  let male = 0;
  data.forEach(c => {
    const s = normalizeSexValue(c.sex);
    if (s === "female") female += 1;
    if (s === "male") male += 1;
  });

  setTextById("female_count", female.toString());
  setTextById("male_count", male.toString());

  // Age groups
  const ageCounts: Record<string, number> = {};
  AGE_GROUPS.forEach(g => { ageCounts[g] = 0; });
  data.forEach(c => {
    if (c.age === null || c.age === undefined || isNaN(c.age) || c.age <= 0) return;
    const group = getAgeGroup(c.age);
    ageCounts[group] = (ageCounts[group] || 0) + 1;
  });

  setTextById("age_1", String(ageCounts["0-17"] ?? 0));
  setTextById("age_2", String(ageCounts["18-25"] ?? 0));
  setTextById("age_3", String(ageCounts["26-40"] ?? 0));
  setTextById("age_4", String(ageCounts["41-59"] ?? 0));
  setTextById("age_5", String(ageCounts["60+"] ?? 0));

  // Top counts lists
  const occupationList = document.querySelector(".occupation_list ul") as HTMLUListElement | null;
  const educationList = document.querySelector(".education_list ul") as HTMLUListElement | null;
  const addressList = document.querySelector(".address_list ul") as HTMLUListElement | null;

  renderCountList(occupationList, getAllCounts(data.map(d => d.occupation)));
  renderCountList(educationList, getAllCounts(data.map(d => d.education)));
  renderCountList(addressList, getAllCounts(data.map(d => d.address)));

}

// ============================================================================
// IMPORTED FILES DROPDOWN
// ============================================================================

function formatBytes(bytes: number): string {
  if (!bytes || bytes < 0) return "0 B";
  const units = ["B", "KB", "MB", "GB"];
  let value = bytes;
  let unitIndex = 0;
  while (value >= 1024 && unitIndex < units.length - 1) {
    value /= 1024;
    unitIndex += 1;
  }
  return `${value.toFixed(unitIndex === 0 ? 0 : 1)} ${units[unitIndex]}`;
}

function renderImportedFilesMenu(): void {
  const menu = document.getElementById("files_menu");
  const dropdown = document.getElementById("files_dropdown");
  if (!menu || !dropdown) return;

  dropdown.innerText = `Import... (${importedFiles.length})`;
  menu.innerHTML = "";

  if (importedFiles.length === 0) {
    menu.innerHTML = `<span class="dropdown-item-text text-secondary">No imported files</span>`;
    return;
  }

  importedFiles.forEach((f) => {
    const meta = f.size ? `(${formatBytes(f.size)})` : "";
    menu.innerHTML += `<div class="dropdown-item d-flex justify-content-between align-items-center gap-2">
      <div class="text-truncate">
        <div class="fw-semibold text-truncate">${f.name}</div>
        <div class="small text-secondary">${meta}</div>
      </div>
      <button type="button" class="btn btn-sm btn-outline-danger delete_btn" data-delete-file="${f.id}"><i class="bi bi-trash3 text-danger"></i></button>
    </div>`;
  });
}

function deleteImportedFile(fileId: string): void {
  const before = importedFiles.length;
  importedFiles = importedFiles.filter((f) => f.id !== fileId);

  if (importedFiles.length === before) return;

  allClients = allClients.filter((c) => c.sourceFileId !== fileId);

  // If the active filters point to values that no longer exist, keep them but UI will show 0.
  renderYearDropdown();
  renderMonthDropdown();
  renderImportedFilesMenu();
  applyFilters();
}

//print
const printBtn = document.getElementById("print_mode_btn");

printBtn?.addEventListener("click", () => {
  document.body.classList.add("print-mode");

  // give charts time to resize
  // setTimeout(() => {
  //   window.print();

  //   // restore UI after printing
  //   setTimeout(() => {
  //     document.body.classList.remove("print-mode");
  //   }, 500);
  // }, 300);
});
