// pages/api/updateAddress.js
import fs from 'fs';
import path from 'path';
import { parse } from 'json2csv';
import csvParser from 'csv-parser';
import { Readable } from 'stream';

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
  'Submitted At'
];

const csvHeaderMap = {
  'Order Number': '*Order Number',
  'Double Check No': 'Backer No.',
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
  'Submitted At': 'Submitted At'
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

function readCsvContent(csvContent) {
  return new Promise((resolve, reject) => {
    const rows = [];
    let headers = [];
    Readable.from([csvContent])
      .pipe(csvParser())
      .on('headers', (h) => {
        headers = h;
      })
      .on('data', (data) => rows.push(data))
      .on('end', () => resolve({ headers, rows }))
      .on('error', reject);
  });
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

function buildUpdatedCsv(result, rows) {
  const normalizedRows = rows.map((row) => normalizeCsvRow(row));

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

  const changedValues = result.comparison.diffs.map((k) => {
    const fromVal = result.original[k] ?? '';
    const toVal = result.submitted[k] ?? '';
    return `${k}: ${fromVal} -> ${toVal}`;
  });

  updatedRow[csvHeaderMap['Match Status']] = result.comparison.matchStatus;
  updatedRow[csvHeaderMap['Changed Fields']] = result.comparison.diffs.join(',');
  updatedRow[csvHeaderMap['Changed Values']] = changedValues.join(' | ');
  updatedRow[csvHeaderMap['Submitted At']] = result.submittedAt;

  const orderHeader = csvHeaderMap['Order Number'];
  const backerHeader = csvHeaderMap['Double Check No'];
  const origIndex = normalizedRows.findIndex(
    (r) => r[orderHeader] === result.identifiers.OrderNumber &&
      r[backerHeader] === result.identifiers.DoubleCheckNo
  );

  if (origIndex === -1) {
    normalizedRows.push(updatedRow);
  } else {
    normalizedRows.splice(origIndex + 1, 0, updatedRow);
  }

  return parse(normalizedRows, { fields: csvHeaders });
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

  return await res.text();
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
    'Double Check No',
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

  let giftList = [];
  try {
    const tokenShareLink = process.env.TOKEN_SHARE_LINK;
    const tokenContent = await downloadContent(
      tokenShareLink,
      'Missing TOKEN_SHARE_LINK'
    );
    giftList = JSON.parse(tokenContent);
  } catch (err) {
    console.error('Failed to read gift_list_token.json:', err);
    return res.status(500).json({ error: 'Failed to read gift_list_token.json' });
  }

  const index = giftList.findIndex((r) => r.token === token);
  if (index === -1) {
    return res.status(404).json({ error: 'Invalid token' });
  }
  if (giftList[index].expire < Date.now()) {
    return res.status(403).json({ error: 'Token expired' });
  }

  const original = giftList[index];
  const fields = editableFields.filter((k) => k in original);

  const diffs = fields.filter((k) => {
    const originalVal = original[k] ?? '';
    const newVal = address[k] ?? '';
    return String(originalVal) !== String(newVal);
  });

  const matchStatus = diffs.length === 0 ? 'MATCH' : 'MISMATCH';

  const result = {
    token,
    submittedAt: new Date().toISOString(),
    identifiers: {
      OrderNumber: original['Order Number'],
      DoubleCheckNo: original['Double Check No'],
      Email: original['Email']
    },
    comparison: { matchStatus, diffs },
    original: { ...original },
    submitted: { ...address }
  };

  let rows = [];
  try {
    const csvContent = await downloadContent(
      process.env.SHARE_LINK,
      'Missing SHARE_LINK'
    );
    const parsed = await readCsvContent(csvContent);
    rows = parsed.rows;

    const normalizedRows = rows.map((row) => normalizeCsvRow(row));
    const orderHeader = csvHeaderMap['Order Number'];
    const backerHeader = csvHeaderMap['Double Check No'];
    const matchingRows = normalizedRows.filter(
      (r) => r[orderHeader] === result.identifiers.OrderNumber &&
        r[backerHeader] === result.identifiers.DoubleCheckNo
    ).length;
    const submissionCount = Math.max(0, matchingRows - 1);

    if (submissionCount >= maxSubmissions) {
      return res.status(429).json({
        error: `Submission limit reached (max ${maxSubmissions}).`
      });
    }
  } catch (err) {
    console.error('Failed to read gift_list.csv:', err);
    return res.status(500).json({ error: 'Failed to read gift_list.csv' });
  }

  const updated = { ...original };
  fields.forEach((k) => {
    if (k in address) updated[k] = address[k];
  });

  giftList[index] = updated;
  try {
    await uploadContent(
      process.env.TOKEN_SHARE_LINK,
      JSON.stringify(giftList, null, 2),
      'application/json',
      'Missing TOKEN_SHARE_LINK'
    );
  } catch (err) {
    console.error('Failed to update gift_list_token.json:', err);
    return res.status(500).json({ error: 'Failed to update gift_list_token.json' });
  }

  try {
    // Save the submitted row right under the original row in gift_list.csv.
    const csvOutput = buildUpdatedCsv(result, rows);
    await uploadContent(
      process.env.SHARE_LINK,
      csvOutput,
      'text/csv',
      'Missing SHARE_LINK'
    );
  } catch (err) {
    console.error('Failed to update gift_list.csv:', err);
    return res.status(500).json({ error: 'Failed to update gift_list.csv' });
  }

  return res.status(200).json({ matchStatus, diffs });
}
