import { google } from 'googleapis';
import { createHmac, timingSafeEqual } from 'node:crypto';

const SPREADSHEET_ID = process.env.GOOGLE_SHEET_ID || '1jdVg33v88AQlnObha0kxecQW3hah3kGmnUdhfA9ipGU';
const PASS_SHEET = 'Gate Passes';
const SETTINGS_SHEET = 'Settings';
const PASS_HEADERS = [
  'id', 'gate_pass_number', 'date', 'person_name', 'mobile_number', 'company_name',
  'taken_by', 'authorized_by', 'status', 'items_json', 'created_at', 'updated_at',
];
const SETTINGS_HEADERS = [
  'id', 'company_name', 'address', 'contact_number', 'logo_path',
  'gate_pass_prefix', 'default_status', 'updated_at',
];
const DEFAULT_SETTINGS = {
  id: 1,
  company_name: 'ISLAM MARBLE',
  address: '',
  contact_number: '',
  logo_path: '',
  gate_pass_prefix: 'GP-',
  default_status: 'Returnable',
};
const COOKIE_NAME = 'islam_marble_session';
const SESSION_DURATION_SECONDS = 7 * 24 * 60 * 60;

let sheetsPromise;
let initPromise;
let writeQueue = Promise.resolve();

function getSheets() {
  if (!sheetsPromise) {
    const { GOOGLE_SERVICE_ACCOUNT_EMAIL, GOOGLE_PRIVATE_KEY } = process.env;
    if (!GOOGLE_SERVICE_ACCOUNT_EMAIL || !GOOGLE_PRIVATE_KEY) {
      throw new Error('Google Sheets credentials are missing. Configure GOOGLE_SERVICE_ACCOUNT_EMAIL and GOOGLE_PRIVATE_KEY in Vercel.');
    }
    const auth = new google.auth.GoogleAuth({
      credentials: {
        client_email: GOOGLE_SERVICE_ACCOUNT_EMAIL,
        private_key: GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n'),
      },
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });
    sheetsPromise = google.sheets({ version: 'v4', auth });
  }
  return sheetsPromise;
}

function columnName(index) {
  let name = '';
  let value = index + 1;
  while (value > 0) {
    const remainder = (value - 1) % 26;
    name = String.fromCharCode(65 + remainder) + name;
    value = Math.floor((value - 1) / 26);
  }
  return name;
}

async function initializeSheets() {
  if (!initPromise) {
    initPromise = (async () => {
      const sheets = getSheets();
      const metadata = await sheets.spreadsheets.get({
        spreadsheetId: SPREADSHEET_ID,
        fields: 'sheets.properties(sheetId,title)',
      });
      const properties = new Map((metadata.data.sheets || []).map((sheet) => [
        sheet.properties.title,
        sheet.properties.sheetId,
      ]));
      const missing = [PASS_SHEET, SETTINGS_SHEET].filter((title) => !properties.has(title));
      if (missing.length) {
        await sheets.spreadsheets.batchUpdate({
          spreadsheetId: SPREADSHEET_ID,
          requestBody: {
            requests: missing.map((title) => ({ addSheet: { properties: { title } } })),
          },
        });
      }

      const headerRanges = [`'${PASS_SHEET}'!A1:Z1`, `'${SETTINGS_SHEET}'!A1:Z1`];
      const headerResult = await sheets.spreadsheets.values.batchGet({
        spreadsheetId: SPREADSHEET_ID,
        ranges: headerRanges,
      });
      const writes = [];
      for (const [index, response] of (headerResult.data.valueRanges || []).entries()) {
        const headers = index === 0 ? PASS_HEADERS : SETTINGS_HEADERS;
        const actual = (response.values?.[0] || []).map((header) => String(header).trim());
        const missingHeaders = headers.filter((header) => !actual.includes(header));
        if (actual.length === 0) {
          writes.push({ range: `'${index === 0 ? PASS_SHEET : SETTINGS_SHEET}'!A1:${columnName(headers.length - 1)}1`, values: [headers] });
        } else if (missingHeaders.length) {
          writes.push({
            range: `'${index === 0 ? PASS_SHEET : SETTINGS_SHEET}'!${columnName(actual.length)}1:${columnName(actual.length + missingHeaders.length - 1)}1`,
            values: [missingHeaders],
          });
        }
      }
      if (writes.length) {
        await sheets.spreadsheets.values.batchUpdate({
          spreadsheetId: SPREADSHEET_ID,
          requestBody: { valueInputOption: 'RAW', data: writes },
        });
      }
    })().catch((error) => {
      initPromise = undefined;
      throw error;
    });
  }
  return initPromise;
}

async function getSheetRows(title) {
  await initializeSheets();
  const sheets = getSheets();
  const result = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${title}'!A:Z`,
    valueRenderOption: 'UNFORMATTED_VALUE',
  });
  return result.data.values || [];
}

function recordsFromRows(rows, headers) {
  const headerRow = rows[0] || [];
  return rows.slice(1).map((values, index) => {
    const record = {};
    for (const key of headers) {
      const columnIndex = headerRow.indexOf(key);
      record[key] = columnIndex >= 0 ? values[columnIndex] ?? '' : '';
    }
    record._rowNumber = index + 2;
    record._values = values;
    return record;
  });
}

function normalizeItems(rawItems, rowNumber) {
  try {
    const items = typeof rawItems === 'string' ? JSON.parse(rawItems || '[]') : [];
    if (!Array.isArray(items)) throw new Error('Items are not an array.');
    return items.map((item, index) => {
      const feet = Number(item.feet ?? (item.unit === 'Feet' ? item.quantity : 0) ?? 0);
      const rate = Number(item.rate || 0);
      return {
        id: index + 1,
        description: String(item.description ?? item.item_serial_number ?? item.purpose ?? ''),
        pieces: Number(item.pieces ?? (item.unit === 'Pcs' ? item.quantity : 0) ?? 0),
        feet,
        rate,
        amount: Math.round((feet * rate + Number.EPSILON) * 100) / 100,
      };
    });
  } catch (error) {
    console.error(`Invalid item data in Google Sheet row ${rowNumber}:`, error);
    throw new Error(`Material data in Google Sheet row ${rowNumber} is invalid.`);
  }
}

function mapPass(record) {
  const items = normalizeItems(record.items_json, record._rowNumber);
  return {
    id: Number(record.id),
    gate_pass_number: String(record.gate_pass_number || ''),
    date: String(record.date || ''),
    person_name: String(record.person_name || ''),
    mobile_number: String(record.mobile_number || ''),
    company_name: String(record.company_name || DEFAULT_SETTINGS.company_name),
    taken_by: String(record.taken_by || ''),
    authorized_by: String(record.authorized_by || ''),
    status: record.status === 'Non-Returnable' ? 'Non-Returnable' : 'Returnable',
    items,
    total_items: items.length,
    total_pcs: items.reduce((total, item) => total + item.pieces, 0),
    total_feet: items.reduce((total, item) => total + item.feet, 0),
    total_amount: items.reduce((total, item) => total + item.amount, 0),
  };
}

async function listPasses() {
  const rows = await getSheetRows(PASS_SHEET);
  return recordsFromRows(rows, PASS_HEADERS)
    .filter((record) => Number(record.id))
    .map(mapPass)
    .sort((a, b) => b.date.localeCompare(a.date) || b.id - a.id);
}

async function getSettings() {
  const rows = await getSheetRows(SETTINGS_SHEET);
  const records = recordsFromRows(rows, SETTINGS_HEADERS);
  const record = records.find((item) => Number(item.id) === 1);
  if (!record) return { ...DEFAULT_SETTINGS };
  return {
    ...DEFAULT_SETTINGS,
    company_name: String(record.company_name || DEFAULT_SETTINGS.company_name),
    address: String(record.address || ''),
    contact_number: String(record.contact_number || ''),
    logo_path: String(record.logo_path || ''),
    gate_pass_prefix: String(record.gate_pass_prefix || DEFAULT_SETTINGS.gate_pass_prefix),
    default_status: record.default_status === 'Non-Returnable' ? 'Non-Returnable' : 'Returnable',
  };
}

function runSerialized(callback) {
  const result = writeQueue.then(callback, callback);
  writeQueue = result.catch(() => {});
  return result;
}

async function writeRecord(title, headers, record, rowNumber) {
  const rows = await getSheetRows(title);
  const currentHeaders = rows[0] || [];
  const values = [...(rowNumber ? rows[rowNumber - 1] || [] : [])];
  for (const [key, value] of Object.entries(record)) {
    const index = currentHeaders.indexOf(key);
    if (index >= 0) values[index] = value;
  }
  const lastColumn = columnName(Math.max(currentHeaders.length, headers.length) - 1);
  const targetRow = rowNumber || Math.max(2, rows.length + 1);
  const sheets = getSheets();
  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${title}'!A${targetRow}:${lastColumn}${targetRow}`,
    valueInputOption: 'RAW',
    requestBody: { values: [values] },
  });
}

async function savePass(payload, editingId = null) {
  const settings = await getSettings();
  const passes = await listPasses();
  const existing = editingId ? passes.find((pass) => pass.id === Number(editingId)) : undefined;
  if (editingId && !existing) throw new Error('Gate pass not found in the Google Sheet.');
  if (!payload.person_name || !payload.mobile_number) throw new Error('Person name and mobile number are required.');

  const items = (payload.items || []).map((item, index) => {
    const description = String(item.description || '').trim();
    const pieces = Number(item.pieces || 0);
    const feet = Number(item.feet || 0);
    const rate = Number(item.rate || 0);
    if (!description) throw new Error(`Description is required for material row ${index + 1}.`);
    if (![pieces, feet, rate].every((value) => Number.isFinite(value) && value >= 0)) {
      throw new Error(`Pieces, feet, and rate must be non-negative numbers for material row ${index + 1}.`);
    }
    return { description, pieces, feet, rate, amount: Math.round((feet * rate + Number.EPSILON) * 100) / 100 };
  });
  if (!items.length) throw new Error('At least one material item is required.');

  return runSerialized(async () => {
    const latestPasses = await listPasses();
    const latestExisting = editingId ? latestPasses.find((pass) => pass.id === Number(editingId)) : undefined;
    if (editingId && !latestExisting) throw new Error('Gate pass not found in the Google Sheet.');
    const id = latestExisting?.id ?? Math.max(0, ...latestPasses.map((pass) => pass.id)) + 1;
    const prefix = settings.gate_pass_prefix || 'GP-';
    const numberPattern = new RegExp(`^${prefix.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}(\\d+)$`);
    const maxNumber = latestPasses.reduce((maximum, pass) => {
      const match = pass.gate_pass_number.match(numberPattern);
      return match ? Math.max(maximum, Number(match[1])) : maximum;
    }, 0);
    const pass = {
      id,
      gate_pass_number: latestExisting?.gate_pass_number || payload.gate_pass_number || `${prefix}${String(maxNumber + 1).padStart(4, '0')}`,
      date: String(payload.date || new Date().toISOString().slice(0, 10)),
      person_name: String(payload.person_name).trim(),
      mobile_number: String(payload.mobile_number).trim(),
      company_name: String(payload.company_name || settings.company_name),
      taken_by: String(payload.taken_by || ''),
      authorized_by: String(payload.authorized_by || ''),
      status: payload.status === 'Non-Returnable' ? 'Non-Returnable' : 'Returnable',
      items,
      created_at: new Date().toISOString(),
      updated_at: new Date().toISOString(),
    };
    const previousRecord = latestExisting
      ? recordsFromRows(await getSheetRows(PASS_SHEET), PASS_HEADERS).find((record) => Number(record.id) === id)
      : undefined;
    const rowNumber = previousRecord?._rowNumber;
    await writeRecord(PASS_SHEET, PASS_HEADERS, {
      ...pass,
      items_json: JSON.stringify(items),
    }, rowNumber);
    return mapPass({ ...pass, items_json: JSON.stringify(items), _rowNumber: rowNumber || 0 });
  });
}

async function deletePass(id) {
  return runSerialized(async () => {
    const rows = await getSheetRows(PASS_SHEET);
    const record = recordsFromRows(rows, PASS_HEADERS).find((item) => Number(item.id) === Number(id));
    if (!record) throw new Error('Gate pass not found in the Google Sheet.');
    const lastColumn = columnName(Math.max(rows[0]?.length || 0, PASS_HEADERS.length) - 1);
    await getSheets().spreadsheets.values.clear({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${PASS_SHEET}'!A${record._rowNumber}:${lastColumn}${record._rowNumber}`,
    });
  });
}

async function saveSettings(payload) {
  return runSerialized(async () => {
    const rows = await getSheetRows(SETTINGS_SHEET);
    const record = recordsFromRows(rows, SETTINGS_HEADERS).find((item) => Number(item.id) === 1);
    const settings = {
      id: 1,
      company_name: String(payload.company_name || DEFAULT_SETTINGS.company_name),
      address: String(payload.address || ''),
      contact_number: String(payload.contact_number || ''),
      logo_path: String(payload.logo_path || ''),
      gate_pass_prefix: String(payload.gate_pass_prefix || DEFAULT_SETTINGS.gate_pass_prefix),
      default_status: payload.default_status === 'Non-Returnable' ? 'Non-Returnable' : 'Returnable',
      updated_at: new Date().toISOString(),
    };
    await writeRecord(SETTINGS_SHEET, SETTINGS_HEADERS, settings, record?._rowNumber);
    return settings;
  });
}

function getAppPassword() {
  const password = process.env.APP_PASSWORD;
  if (!password || password.length < 12) {
    throw new Error('Set APP_PASSWORD in Vercel Environment Variables (use at least 12 characters).');
  }
  return password;
}

function createSessionToken(password, expiresAt) {
  return createHmac('sha256', password).update(String(expiresAt)).digest('hex');
}

function sessionCookie(password) {
  const expiresAt = Math.floor(Date.now() / 1000) + SESSION_DURATION_SECONDS;
  const token = `${expiresAt}.${createSessionToken(password, expiresAt)}`;
  const secure = process.env.NODE_ENV === 'production' ? '; Secure' : '';
  return `${COOKIE_NAME}=${token}; HttpOnly; SameSite=Strict; Path=/; Max-Age=${SESSION_DURATION_SECONDS}${secure}`;
}

function isAuthenticated(req) {
  try {
    const password = getAppPassword();
    const cookieHeader = String(req.headers.cookie || '');
    const cookie = cookieHeader.split(';').map((part) => part.trim()).find((part) => part.startsWith(`${COOKIE_NAME}=`));
    if (!cookie) return false;
    const [expiresAtText, signature] = decodeURIComponent(cookie.slice(COOKIE_NAME.length + 1)).split('.');
    const expiresAt = Number(expiresAtText);
    if (!Number.isSafeInteger(expiresAt) || expiresAt <= Math.floor(Date.now() / 1000) || !signature) return false;
    const expected = createSessionToken(password, expiresAt);
    const suppliedBuffer = Buffer.from(signature);
    const expectedBuffer = Buffer.from(expected);
    return suppliedBuffer.length === expectedBuffer.length && timingSafeEqual(suppliedBuffer, expectedBuffer);
  } catch {
    return false;
  }
}

function verifyPassword(value) {
  const password = getAppPassword();
  const supplied = Buffer.from(String(value || ''));
  const expected = Buffer.from(password);
  return supplied.length === expected.length && timingSafeEqual(supplied, expected);
}

export {
  deletePass,
  getSettings,
  isAuthenticated,
  listPasses,
  savePass,
  saveSettings,
  sessionCookie,
  verifyPassword,
};
