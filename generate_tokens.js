const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const ExcelJS = require('exceljs');

const csvHeaders = [
  '*Order Number',
  'Backer No.',
  '*Reward Name',
  '*Item Quantity',
  '*Full Name',
  'Phone Number',
  '*Country/Region Code',
  'State/Province/Region',
  '*City',
  '*Address1',
  'Address2',
  'Zip Code',
  'Email',
  'Total Payment',
  'Match Status',
  'Changed Fields',
  'Changed Values',
  'Submitted At',
  'Token Link'
];

const headerMap = {
  '*Order Number': 'Order Number',
  'Backer No.': 'Backer No.',
  '*Reward Name': 'Reward Name',
  '*Item Quantity': 'Item Quantity',
  '*Full Name': 'Full Name',
  'Phone Number': 'Phone Number',
  '*Country/Region Code': 'Country/Region Code',
  'State/Province/Region': 'State/Province/Region',
  '*City': 'City',
  '*Address1': 'Address1',
  'Address2': 'Address2',
  'Zip Code': 'Zip Code',
  'Email': 'Email',
  'Total Payment': 'Total Payment',
  'Match Status': 'Match Status',
  'Changed Fields': 'Changed Fields',
  'Changed Values': 'Changed Values',
  'Submitted At': 'Submitted At',
  'Token Link': 'Token Link'
};

function loadEnvFile() {
  const envPath = path.join(process.cwd(), '.env.local');
  if (!fs.existsSync(envPath)) return;

  const content = fs.readFileSync(envPath, 'utf8');
  content.split(/\r?\n/).forEach((line) => {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith('#')) return;
    const [key, ...rest] = trimmed.split('=');
    if (!key) return;
    if (process.env[key]) return;
    process.env[key] = rest.join('=').trim();
  });
}

function normalizeRow(row) {
  const normalized = {};
  Object.keys(row).forEach((key) => {
    const mappedKey = headerMap[key] || key;
    normalized[mappedKey] = row[key];
  });

  return normalized;
}

function normalizeCsvRow(row) {
  const normalized = {};
  csvHeaders.forEach((h) => {
    normalized[h] = row[h] ?? '';
  });
  return normalized;
}

function encodeShareLink(link) {
  const base64 = Buffer.from(link)
    .toString('base64')
    .replace(/\+/g, '-')
    .replace(/\//g, '_')
    .replace(/=+$/g, '');
  return `u!${base64}`;
}

async function getGraphAccessToken() {
  const tenantId = process.env.TENANT_ID;
  const clientId = process.env.CLIENT_ID;
  const clientSecret = process.env.CLIENT_SECRET;

  if (!tenantId || !clientId || !clientSecret) {
    throw new Error('Missing Graph credentials');
  }

  const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id: clientId,
    client_secret: clientSecret,
    scope: 'https://graph.microsoft.com/.default',
    grant_type: 'client_credentials'
  });

  const res = await fetch(tokenUrl, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body
  });

  if (!res.ok) {
    const errorText = await res.text();
    throw new Error(`Token request failed: ${errorText}`);
  }

  const data = await res.json();
  return data.access_token;
}

async function downloadContent(shareLink) {
  if (!shareLink) throw new Error('Missing SHARE_LINK');

  const accessToken = await getGraphAccessToken();
  const shareId = encodeShareLink(shareLink);
  const url = `https://graph.microsoft.com/v1.0/shares/${shareId}/driveItem/content`;

  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${accessToken}` }
  });

  if (!res.ok) {
    const errorText = await res.text();
    throw new Error(`Failed to download XLSX: ${errorText}`);
  }

  const arrayBuffer = await res.arrayBuffer();
  return Buffer.from(arrayBuffer);
}

async function uploadContent(shareLink, content) {
  if (!shareLink) throw new Error('Missing SHARE_LINK');

  const accessToken = await getGraphAccessToken();
  const shareId = encodeShareLink(shareLink);
  const url = `https://graph.microsoft.com/v1.0/shares/${shareId}/driveItem/content`;

  const res = await fetch(url, {
    method: 'PUT',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    },
    body: content
  });

  if (!res.ok) {
    const errorText = await res.text();
    throw new Error(`Failed to upload XLSX: ${errorText}`);
  }
}

function normalizeCellValue(value) {
  if (value == null) return '';
  if (typeof value === 'object') {
    if (typeof value.text === 'string') return value.text;
    if (Array.isArray(value.richText)) {
      return value.richText.map((part) => part.text || '').join('');
    }
    if (value.formula && value.result != null) return String(value.result);
    if (value.hyperlink) return value.text || value.hyperlink;
  }
  return String(value);
}

async function readWorkbook(buffer) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  const worksheet = workbook.worksheets[0];
  if (!worksheet) return { rows: [], sheetName: 'Sheet1' };

  const headerRow = worksheet.getRow(1).values.slice(1);
  const rows = [];
  for (let rowIndex = 2; rowIndex <= worksheet.rowCount; rowIndex += 1) {
    const row = worksheet.getRow(rowIndex);
    if (row.actualCellCount === 0) continue;
    const obj = {};
    headerRow.forEach((header, index) => {
      if (!header) return;
      const cellValue = row.getCell(index + 1).value;
      obj[header] = normalizeCellValue(cellValue);
    });
    rows.push(obj);
  }

  return { rows, sheetName: worksheet.name || 'Sheet1' };
}

async function writeWorkbook(rows, sheetName) {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet(sheetName || 'Sheet1');
  worksheet.addRow(csvHeaders);
  rows.forEach((row) => {
    const values = csvHeaders.map((header) => row[header] ?? '');
    worksheet.addRow(values);
  });
  const buffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(buffer);
}

function buildTokenLink(token) {
  return `https://gift-list-tool.vercel.app/?token=${token}`;
}

async function main() {
  loadEnvFile();

  const shareLink = process.env.SHARE_LINK;
  if (!shareLink) throw new Error('Missing SHARE_LINK');

  const workbookBuffer = await downloadContent(shareLink);
  const { rows, sheetName } = await readWorkbook(workbookBuffer);

  const giftListWithToken = rows.map((row) => {
    const token = crypto.randomBytes(16).toString('hex');
    const expire = Date.now() + 7 * 24 * 60 * 60 * 1000;
    const tokenLink = buildTokenLink(token);
    row['Token Link'] = tokenLink;

    const normalized = normalizeRow(row);
    return { ...normalized, token, tokenLink, expire };
  });

  const normalizedRows = rows.map((row) => normalizeCsvRow(row));
  const workbookOutput = await writeWorkbook(normalizedRows, sheetName);
  await uploadContent(shareLink, workbookOutput);

  fs.writeFileSync('gift_list_token.json', JSON.stringify(giftListWithToken, null, 2));
  console.log('gift_list.xlsx updated and gift_list_token.json generated.');
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
