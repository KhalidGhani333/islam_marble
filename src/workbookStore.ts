import ExcelJS from 'exceljs';

export type WorkbookSettings = {
  company_name: string;
  address: string;
  contact_number: string;
  logo_path: string;
  gate_pass_prefix: string;
  default_status: 'Returnable' | 'Non-Returnable';
};

export type WorkbookPass = {
  id: number;
  gate_pass_number: string;
  date: string;
  person_name: string;
  mobile_number: string;
  company_name: string;
  taken_by: string;
  authorized_by: string;
  status: 'Returnable' | 'Non-Returnable';
  items: Array<{
    id: number;
    description: string;
    pieces: number;
    feet: number;
    rate: number;
    amount: number;
  }>;
  total_items: number;
  total_pcs: number;
  total_feet: number;
  total_amount: number;
};

type LocalWorkbookFileHandle = {
  name: string;
  getFile(): Promise<File>;
  createWritable(): Promise<{
    write(data: Blob | BufferSource): Promise<void>;
    close(): Promise<void>;
    abort?(): Promise<void>;
  }>;
  queryPermission?(options: { mode: 'readwrite' }): Promise<PermissionState>;
  requestPermission?(options: { mode: 'readwrite' }): Promise<PermissionState>;
};

type WorkbookPickerWindow = Window & {
  showOpenFilePicker?: (options: object) => Promise<LocalWorkbookFileHandle[]>;
  showSaveFilePicker?: (options: object) => Promise<LocalWorkbookFileHandle>;
};

const PASS_HEADERS = [
  'id', 'gate_pass_number', 'date', 'person_name', 'mobile_number', 'company_name',
  'taken_by', 'authorized_by', 'status', 'items_json', 'created_at', 'updated_at',
];
const SETTINGS_HEADERS = [
  'id', 'company_name', 'address', 'contact_number', 'logo_path',
  'gate_pass_prefix', 'default_status',
];
const DEFAULT_SETTINGS: WorkbookSettings = {
  company_name: 'ISLAM MARBLE',
  address: '',
  contact_number: '',
  logo_path: '',
  gate_pass_prefix: 'GP-',
  default_status: 'Returnable',
};

let workbook: ExcelJS.Workbook | null = null;
let fileHandle: LocalWorkbookFileHandle | null = null;
let fileName = '';

function getWorksheet(name: string, headers: string[]) {
  if (!workbook) throw new Error('Open or create an Excel workbook first.');
  const sheet = workbook.getWorksheet(name) || workbook.addWorksheet(name);
  if (sheet.rowCount === 0) {
    sheet.addRow(headers);
    sheet.getRow(1).font = { bold: true };
    sheet.views = [{ state: 'frozen', ySplit: 1 }];
  } else {
    const actualHeaders = headers.map((_, index) => String(sheet.getRow(1).getCell(index + 1).value ?? '').trim());
    if (actualHeaders.some((header, index) => header !== headers[index])) {
      throw new Error(`The "${name}" worksheet has a different layout. Use an Excel workbook created by this app.`);
    }
  }
  return sheet;
}

function readCell(value: ExcelJS.CellValue): string {
  if (value === null || value === undefined) return '';
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  if (typeof value === 'object') {
    if ('text' in value && typeof value.text === 'string') return value.text.trim();
    if ('result' in value) return String(value.result ?? '').trim();
    return '';
  }
  return String(value).trim();
}

function readSettings(): WorkbookSettings {
  const sheet = getWorksheet('Settings', SETTINGS_HEADERS);
  const row = sheet.getRow(2);
  return {
    company_name: readCell(row.getCell(2).value) || DEFAULT_SETTINGS.company_name,
    address: readCell(row.getCell(3).value),
    contact_number: readCell(row.getCell(4).value),
    logo_path: readCell(row.getCell(5).value),
    gate_pass_prefix: readCell(row.getCell(6).value) || DEFAULT_SETTINGS.gate_pass_prefix,
    default_status: readCell(row.getCell(7).value) === 'Non-Returnable' ? 'Non-Returnable' : 'Returnable',
  };
}

function summarizePass(pass: Omit<WorkbookPass, 'total_items' | 'total_pcs' | 'total_feet' | 'total_amount'>): WorkbookPass {
  const items = pass.items.map((item, index) => ({
    ...item,
    id: index + 1,
    amount: Math.round((Number(item.feet || 0) * Number(item.rate || 0) + Number.EPSILON) * 100) / 100,
  }));
  return {
    ...pass,
    items,
    total_items: items.length,
    total_pcs: items.reduce((total, item) => total + item.pieces, 0),
    total_feet: items.reduce((total, item) => total + item.feet, 0),
    total_amount: items.reduce((total, item) => total + item.amount, 0),
  };
}

function readPasses(): WorkbookPass[] {
  const sheet = getWorksheet('Gate Passes', PASS_HEADERS);
  const passes: WorkbookPass[] = [];
  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    const id = Number(row.getCell(1).value);
    if (!id) return;

    let rawItems: Array<Record<string, unknown>> = [];
    try {
      const savedItems = readCell(row.getCell(10).value);
      rawItems = savedItems ? JSON.parse(savedItems) : [];
    } catch (error) {
      console.error(`Invalid items_json in Excel row ${rowNumber}:`, error);
      throw new Error(`Invalid material data in the Excel file, row ${rowNumber}.`);
    }

    passes.push(summarizePass({
      id,
      gate_pass_number: readCell(row.getCell(2).value),
      date: readCell(row.getCell(3).value),
      person_name: readCell(row.getCell(4).value),
      mobile_number: readCell(row.getCell(5).value),
      company_name: readCell(row.getCell(6).value) || DEFAULT_SETTINGS.company_name,
      taken_by: readCell(row.getCell(7).value),
      authorized_by: readCell(row.getCell(8).value),
      status: readCell(row.getCell(9).value) === 'Non-Returnable' ? 'Non-Returnable' : 'Returnable',
      items: rawItems.map((item, index) => ({
        id: index + 1,
        description: String(item.description ?? item.item_serial_number ?? item.purpose ?? ''),
        pieces: Number(item.pieces ?? (item.unit === 'Pcs' ? item.quantity : 0) ?? 0),
        feet: Number(item.feet ?? (item.unit === 'Feet' ? item.quantity : 0) ?? 0),
        rate: Number(item.rate || 0),
        amount: Number(item.amount || 0),
      })),
    }));
  });
  return passes;
}

async function ensureWritePermission() {
  if (!fileHandle) throw new Error('Select or create an Excel file before saving.');
  const currentPermission = await fileHandle.queryPermission?.({ mode: 'readwrite' });
  const permission = currentPermission === 'granted'
    ? currentPermission
    : await fileHandle.requestPermission?.({ mode: 'readwrite' });
  if (permission && permission !== 'granted') {
    throw new Error('Excel file write permission was not granted.');
  }
}

async function persistWorkbook() {
  if (!workbook || !fileHandle) throw new Error('Open or create an Excel workbook first.');
  await ensureWritePermission();
  const bytes = await workbook.xlsx.writeBuffer();
  const writable = await fileHandle.createWritable();
  try {
    await writable.write(bytes);
    await writable.close();
  } catch (error) {
    await writable.abort?.();
    throw error;
  }
}

async function savePassRows(passes: WorkbookPass[]) {
  const sheet = getWorksheet('Gate Passes', PASS_HEADERS);
  if (sheet.rowCount > 1) sheet.spliceRows(2, sheet.rowCount - 1);
  for (const pass of passes) {
    sheet.addRow([
      pass.id,
      pass.gate_pass_number,
      pass.date,
      pass.person_name,
      pass.mobile_number,
      pass.company_name,
      pass.taken_by,
      pass.authorized_by,
      pass.status,
      JSON.stringify(pass.items),
      new Date().toISOString(),
      new Date().toISOString(),
    ]);
  }
  await persistWorkbook();
}

function workbookAsBlob(buffer: ExcelJS.Buffer) {
  return new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
}

export function getWorkbookName() {
  return fileName;
}

export async function openWorkbookFile() {
  const pickerWindow = window as WorkbookPickerWindow;
  if (!pickerWindow.showOpenFilePicker) {
    throw new Error('Opening and saving Excel files directly requires Chrome or Edge on HTTPS or localhost.');
  }
  const [selectedFile] = await pickerWindow.showOpenFilePicker({
    multiple: false,
    types: [{
      description: 'Excel workbook',
      accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] },
    }],
  });
  const selectedWorkbook = new ExcelJS.Workbook();
  await selectedWorkbook.xlsx.load(await (await selectedFile.getFile()).arrayBuffer());
  workbook = selectedWorkbook;
  fileHandle = selectedFile;
  fileName = selectedFile.name;
  const settings = readSettings();
  const passes = readPasses();
  return { fileName, settings, passes };
}

export async function createWorkbookFile() {
  const pickerWindow = window as WorkbookPickerWindow;
  if (!pickerWindow.showSaveFilePicker) {
    throw new Error('Creating and saving Excel files directly requires Chrome or Edge on HTTPS or localhost.');
  }
  const selectedFile = await pickerWindow.showSaveFilePicker({
    suggestedName: 'Islam-Marble-Gate-Passes.xlsx',
    types: [{
      description: 'Excel workbook',
      accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] },
    }],
  });
  workbook = new ExcelJS.Workbook();
  workbook.creator = 'Islam Marble Gate Pass';
  workbook.created = new Date();
  fileHandle = selectedFile;
  fileName = selectedFile.name;
  getWorksheet('Gate Passes', PASS_HEADERS);
  getWorksheet('Settings', SETTINGS_HEADERS);
  await persistWorkbook();
  return { fileName, settings: { ...DEFAULT_SETTINGS }, passes: [] as WorkbookPass[] };
}

export async function saveGatePass(payload: Omit<WorkbookPass, 'total_items' | 'total_pcs' | 'total_feet' | 'total_amount'>, editingId: number | null) {
  if (!workbook || !fileHandle) throw new Error('Select or create an Excel file before saving.');
  const settings = readSettings();
  const existingPasses = readPasses();
  const existing = editingId ? existingPasses.find((pass) => pass.id === editingId) : undefined;
  if (editingId && !existing) throw new Error('Gate pass not found in the selected Excel file.');

  const id = existing?.id ?? Math.max(0, ...existingPasses.map((pass) => pass.id)) + 1;
  const prefix = settings.gate_pass_prefix || 'GP-';
  const passNumberPattern = new RegExp(`^${prefix.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}(\\d+)$`);
  const nextNumber = existingPasses.reduce((maximum, pass) => {
    const match = pass.gate_pass_number.match(passNumberPattern);
    return match ? Math.max(maximum, Number(match[1])) : maximum;
  }, 0) + 1;
  const pass = summarizePass({
    ...payload,
    id,
    gate_pass_number: existing?.gate_pass_number || payload.gate_pass_number || `${prefix}${String(nextNumber).padStart(4, '0')}`,
    company_name: payload.company_name || settings.company_name,
    items: payload.items,
  });
  const nextPasses = existing
    ? existingPasses.map((item) => item.id === existing.id ? pass : item)
    : [...existingPasses, pass];
  await savePassRows(nextPasses);
  return pass;
}

export async function deleteGatePass(id: number) {
  const passes = readPasses();
  if (!passes.some((pass) => pass.id === id)) throw new Error('Gate pass not found in the selected Excel file.');
  await savePassRows(passes.filter((pass) => pass.id !== id));
}

export async function duplicateGatePass(id: number) {
  const original = readPasses().find((pass) => pass.id === id);
  if (!original) throw new Error('Gate pass not found in the selected Excel file.');
  const { id: _id, gate_pass_number: _number, total_items: _count, total_pcs: _pcs, total_feet: _feet, total_amount: _amount, ...copy } = original;
  return saveGatePass(copy, null);
}

export async function saveWorkbookSettings(settings: WorkbookSettings) {
  const sheet = getWorksheet('Settings', SETTINGS_HEADERS);
  sheet.getRow(2).values = [
    1,
    settings.company_name,
    settings.address,
    settings.contact_number,
    settings.logo_path,
    settings.gate_pass_prefix,
    settings.default_status,
  ];
  await persistWorkbook();
}

export async function downloadWorkbookCopy() {
  if (!workbook) throw new Error('Open or create an Excel workbook first.');
  const blob = workbookAsBlob(await workbook.xlsx.writeBuffer());
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = fileName || 'Islam-Marble-Gate-Passes.xlsx';
  link.click();
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
}

export async function downloadReportWorkbook(passes: WorkbookPass[]) {
  const report = new ExcelJS.Workbook();
  report.creator = 'Islam Marble';
  report.created = new Date();

  const registerSheet = report.addWorksheet('Gate Pass Register');
  registerSheet.columns = [
    { header: 'Gate Pass No.', key: 'gate_pass_number', width: 18 },
    { header: 'Date', key: 'date', width: 14 },
    { header: 'Person Name', key: 'person_name', width: 22 },
    { header: 'Mobile Number', key: 'mobile_number', width: 18 },
    { header: 'Company Name', key: 'company_name', width: 20 },
    { header: 'Total Items', key: 'total_items', width: 12 },
    { header: 'Total Pcs', key: 'total_pcs', width: 12 },
    { header: 'Total Feet', key: 'total_feet', width: 12 },
    { header: 'Total Amount', key: 'total_amount', width: 16 },
    { header: 'Description Summary', key: 'description_summary', width: 28 },
    { header: 'Taken By', key: 'taken_by', width: 18 },
    { header: 'Status', key: 'status', width: 16 },
    { header: 'Authorized By', key: 'authorized_by', width: 18 },
  ];
  const itemsSheet = report.addWorksheet('Gate Pass Items');
  itemsSheet.columns = [
    { header: 'Gate Pass No.', key: 'gate_pass_number', width: 18 },
    { header: 'S. No.', key: 'line_number', width: 10 },
    { header: 'Description', key: 'description', width: 30 },
    { header: 'Pieces', key: 'pieces', width: 14 },
    { header: 'Feet', key: 'feet', width: 14 },
    { header: 'Rate / Ft', key: 'rate', width: 14 },
    { header: 'Amount', key: 'amount', width: 16 },
  ];

  for (const pass of passes) {
    registerSheet.addRow({
      ...pass,
      description_summary: pass.items.map((item) => item.description).join('; '),
    });
    pass.items.forEach((item, index) => itemsSheet.addRow({
      gate_pass_number: pass.gate_pass_number,
      line_number: index + 1,
      description: item.description,
      pieces: item.pieces,
      feet: item.feet,
      rate: item.rate,
      amount: item.amount,
    }));
  }
  for (const sheet of [registerSheet, itemsSheet]) {
    sheet.views = [{ state: 'frozen', ySplit: 1 }];
    sheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
    sheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F2937' } };
  }

  const summarySheet = report.addWorksheet('Summary');
  summarySheet.columns = [{ header: 'Metric', width: 24 }, { header: 'Value', width: 18 }];
  summarySheet.addRows([
    ['Total Gate Passes', passes.length],
    ['Total Material Items', passes.reduce((total, pass) => total + pass.total_items, 0)],
    ['Total Pieces', passes.reduce((total, pass) => total + pass.total_pcs, 0)],
    ['Total Feet', passes.reduce((total, pass) => total + pass.total_feet, 0)],
    ['Total Amount', passes.reduce((total, pass) => total + pass.total_amount, 0)],
  ]);
  const blob = workbookAsBlob(await report.xlsx.writeBuffer());
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = 'islam-marble-gate-pass-report.xlsx';
  link.click();
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
}
