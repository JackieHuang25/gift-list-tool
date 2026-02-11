// pages/api/getAddress.js
import ExcelJS from 'exceljs';

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

const csvHeaderMap = {
  'Order Number': '*Order Number',
  'Backer No.': 'Backer No.',
  'Reward Name': '*Reward Name',
  'Item Quantity': '*Item Quantity',
  'Full Name': '*Full Name',
  'Phone Number': 'Phone Number',
  'Country/Region Code': '*Country/Region Code',
  'State/Province/Region': 'State/Province/Region',
  'City': '*City',
  'Address1': '*Address1',
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

async function downloadWorkbookContent() {
  const shareLink = process.env.SHARE_LINK;
  if (!shareLink) throw new Error('Missing SHARE_LINK');

  const accessToken = await getGraphAccessToken();
  const shareId = encodeShareLink(shareLink);
  const url = `https://graph.microsoft.com/v1.0/shares/${shareId}/driveItem/content`;

  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${accessToken}` }
  });

  if (!res.ok) {
    const errorText = await res.text();
    throw new Error(`Failed to download XLSX content: ${errorText}`);
  }

  const arrayBuffer = await res.arrayBuffer();
  return Buffer.from(arrayBuffer);
}

function buildTokenLink(token) {
  return `https://gift-list-tool.vercel.app/?token=${token}`;
}

function matchesToken(rowToken, token) {
  if (!rowToken) return false;
  const tokenLink = buildTokenLink(token);
  return rowToken === tokenLink || rowToken.endsWith(`?token=${token}`) || rowToken === token;
}

function normalizeCsvRow(row) {
  const normalized = {};
  csvHeaders.forEach((h) => {
    normalized[h] = row[h] ?? '';
  });
  return normalized;
}

function csvRowToStandard(row) {
  const standard = {};
  Object.keys(csvHeaderMap).forEach((standardKey) => {
    const csvKey = csvHeaderMap[standardKey];
    standard[standardKey] = row[csvKey] ?? '';
  });
  return standard;
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
  if (!worksheet) return [];

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

  return rows;
}

export default async function handler(req, res) {
  const { token } = req.query;

  if (!token) {
    return res.status(400).json({ error: 'Missing token' });
  }

  let rows = [];
  try {
    const workbookBuffer = await downloadWorkbookContent();
    rows = await readWorkbook(workbookBuffer);
  } catch (err) {
    console.error('Failed to read gift_list.xlsx:', err);
    return res.status(500).json({ error: 'Failed to read gift_list.xlsx' });
  }

  const normalizedRows = rows.map((row) => normalizeCsvRow(row));
  const tokenHeader = csvHeaderMap['Token Link'];
  let match = null;
  for (let i = normalizedRows.length - 1; i >= 0; i -= 1) {
    if (matchesToken(normalizedRows[i][tokenHeader], token)) {
      match = normalizedRows[i];
      break;
    }
  }
  if (!match) return res.status(404).json({ error: 'Invalid token' });

  const record = csvRowToStandard(match);
  if (record.expire && record.expire < Date.now()) {
    return res.status(403).json({ error: 'Token expired' });
  }

  return res.status(200).json(record);
}
