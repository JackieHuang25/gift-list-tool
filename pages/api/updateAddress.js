// pages/api/updateAddress.js
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

const legacyHeaderMap = {
  '*Order Number': ['Order Number'],
  'Backer No.': ['Double Check No'],
  '*Reward Name': ['Reward Name'],
  '*Item Quantity': ['Item Quantity'],
  '*Full Name': ['Full Name'],
  '*Country/Region Code': ['Country/Region Code'],
  '*City': ['City'],
  '*Address1': ['Address1']
};

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
  if (!worksheet) {
    return { sheetName: 'Sheet1', rows: [] };
  }

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

  return { sheetName: worksheet.name || 'Sheet1', rows };
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

function normalizeCsvRow(row) {
  const normalized = {};
  csvHeaders.forEach((h) => {
    if (row[h] !== undefined) {
      normalized[h] = row[h];
      return;
    }

    const legacyKeys = legacyHeaderMap[h] || [];
    for (const legacyKey of legacyKeys) {
      if (row[legacyKey] !== undefined) {
        normalized[h] = row[legacyKey];
        return;
      }
    }

    normalized[h] = '';
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

function buildUpdatedRows(result, normalizedRows) {
  const updatedRow = {};
  csvHeaders.forEach((h) => {
    updatedRow[h] = '';
  });

  Object.keys(result.original).forEach((k) => {
    const header = csvHeaderMap[k];
    if (header) updatedRow[header] = result.original[k] ?? '';
  });
  Object.keys(result.submitted).forEach((k) => {
    const header = csvHeaderMap[k];
    if (header) updatedRow[header] = result.submitted[k];
  });

  const changedValues = result.comparison.changedValues || result.comparison.diffs.map((k) => {
    const fromVal = result.original[k] ?? '';
    const toVal = result.submitted[k] ?? '';
    return `${k}: ${fromVal} -> ${toVal}`;
  });

  updatedRow[csvHeaderMap['Match Status']] = result.comparison.matchStatus;
  updatedRow[csvHeaderMap['Changed Fields']] = result.comparison.diffs.join(',');
  updatedRow[csvHeaderMap['Changed Values']] = changedValues.join(' | ');
  updatedRow[csvHeaderMap['Submitted At']] = result.submittedAt;
  updatedRow[csvHeaderMap['Token Link']] = result.tokenLink;

  const orderHeader = csvHeaderMap['Order Number'];
  const backerHeader = csvHeaderMap['Backer No.'];
  const origIndex = normalizedRows.findIndex(
    (r) => r[orderHeader] === result.identifiers.OrderNumber &&
      r[backerHeader] === result.identifiers.BackerNo
  );

  if (origIndex === -1) {
    normalizedRows.push(updatedRow);
  } else {
    normalizedRows.splice(origIndex + 1, 0, updatedRow);
  }

  return normalizedRows;
}

function buildTokenLink(token) {
  return `https://gift-list-tool.vercel.app/?token=${token}`;
}

function matchesToken(rowToken, token) {
  if (!rowToken) return false;
  const tokenLink = buildTokenLink(token);
  return rowToken === tokenLink || rowToken.endsWith(`?token=${token}`) || rowToken === token;
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

async function downloadContent(shareLink, missingMessage) {
  if (!shareLink) throw new Error(missingMessage);

  const accessToken = await getGraphAccessToken();
  const shareId = encodeShareLink(shareLink);
  const url = `https://graph.microsoft.com/v1.0/shares/${shareId}/driveItem/content`;

  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${accessToken}` }
  });

  if (!res.ok) {
    const errorText = await res.text();
    throw new Error(`Failed to download content: ${errorText}`);
  }

  const arrayBuffer = await res.arrayBuffer();
  return Buffer.from(arrayBuffer);
}

async function uploadContent(shareLink, content, contentType, missingMessage) {
  if (!shareLink) throw new Error(missingMessage);

  const accessToken = await getGraphAccessToken();
  const shareId = encodeShareLink(shareLink);
  const url = `https://graph.microsoft.com/v1.0/shares/${shareId}/driveItem/content`;

  const res = await fetch(url, {
    method: 'PUT',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': contentType
    },
    body: content
  });

  if (!res.ok) {
    const errorText = await res.text();
    throw new Error(`Failed to upload content: ${errorText}`);
  }
}

export default async function handler(req, res) {
  const editableFields = [
    'Full Name',
    'Phone Number',
    'Country/Region Code',
    'State/Province/Region',
    'City',
    'Address1',
    'Address2',
    'Zip Code',
    'Email'
  ];
  const maxSubmissions = 3;

  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  const { token, address } = req.body || {};
  if (!token || !address) {
    return res.status(400).json({ error: 'Missing token or address' });
  }

  const tokenLink = buildTokenLink(token);
  let rows = [];
  let normalizedRows = [];
  let original = null;
  let sheetName = 'Sheet1';
  try {
    const workbookBuffer = await downloadContent(
      process.env.SHARE_LINK,
      'Missing SHARE_LINK'
    );
    const parsed = await readWorkbook(workbookBuffer);
    rows = parsed.rows;
    sheetName = parsed.sheetName;
    normalizedRows = rows.map((row) => normalizeCsvRow(row));

    const tokenHeader = csvHeaderMap['Token Link'];
    const origIndex = normalizedRows.findIndex((row) =>
      matchesToken(row[tokenHeader], token)
    );

    if (origIndex === -1) {
      return res.status(404).json({ error: 'Invalid token' });
    }

    normalizedRows[origIndex][tokenHeader] = tokenLink;
    original = csvRowToStandard(normalizedRows[origIndex]);

    const orderHeader = csvHeaderMap['Order Number'];
    const backerHeader = csvHeaderMap['Backer No.'];
    const matchingRows = normalizedRows.filter(
      (r) => r[orderHeader] === original['Order Number'] &&
        r[backerHeader] === original['Backer No.']
    ).length;
    const submissionCount = Math.max(0, matchingRows - 1);

    if (submissionCount >= maxSubmissions) {
      return res.status(429).json({
        error: `Submission limit reached (max ${maxSubmissions}).`
      });
    }
  } catch (err) {
    console.error('Failed to read gift_list.xlsx:', err);
    return res.status(500).json({ error: 'Failed to read gift_list.xlsx' });
  }

  if (!original) {
    return res.status(404).json({ error: 'Invalid token' });
  }

  if (original.expire && original.expire < Date.now()) {
    return res.status(403).json({ error: 'Token expired' });
  }

  const fields = editableFields.filter((k) => k in original);

  const diffs = fields.filter((k) => {
    const originalVal = original[k] ?? '';
    const newVal = address[k] ?? '';
    return String(originalVal) !== String(newVal);
  });

  const matchStatus = diffs.length === 0 ? 'MATCH' : 'MISMATCH';
  const changedValues = diffs.map((k) => {
    const fromVal = original[k] ?? '';
    const toVal = address[k] ?? '';
    return `${k}: ${fromVal} -> ${toVal}`;
  });

  const result = {
    token,
    tokenLink,
    submittedAt: new Date().toISOString(),
    identifiers: {
      OrderNumber: original['Order Number'],
      BackerNo: original['Backer No.'],
      Email: original['Email']
    },
    comparison: { matchStatus, diffs, changedValues },
    original: { ...original },
    submitted: { ...address }
  };

  const tokenHeader = csvHeaderMap['Token Link'];
  const origIndex = normalizedRows.findIndex((row) =>
    matchesToken(row[tokenHeader], token)
  );
  if (origIndex === -1) {
    return res.status(404).json({ error: 'Invalid token' });
  }

  fields.forEach((k) => {
    if (k in address) {
      const csvKey = csvHeaderMap[k];
      if (csvKey) normalizedRows[origIndex][csvKey] = address[k];
    }
  });
  normalizedRows[origIndex][tokenHeader] = tokenLink;

  try {
    const updatedRows = buildUpdatedRows(result, normalizedRows);
    const workbookBuffer = await writeWorkbook(updatedRows, sheetName);
    await uploadContent(
      process.env.SHARE_LINK,
      workbookBuffer,
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'Missing SHARE_LINK'
    );
  } catch (err) {
    console.error('Failed to update gift_list.xlsx:', err);
    return res.status(500).json({ error: 'Failed to update gift_list.xlsx' });
  }

  return res.status(200).json({ matchStatus, diffs, changedValues });
}