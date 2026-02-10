// pages/api/getAddress.js
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

async function downloadTokenContent() {
  const shareLink = process.env.TOKEN_SHARE_LINK;
  if (!shareLink) throw new Error('Missing TOKEN_SHARE_LINK');

  const accessToken = await getGraphAccessToken();
  const shareId = encodeShareLink(shareLink);
  const url = `https://graph.microsoft.com/v1.0/shares/${shareId}/driveItem/content`;

  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${accessToken}` }
  });

  if (!res.ok) {
    const errorText = await res.text();
    throw new Error(`Failed to download token content: ${errorText}`);
  }

  return await res.text();
}

export default async function handler(req, res) {
  const { token } = req.query;

  let giftList = [];
  try {
    const tokenContent = await downloadTokenContent();
    giftList = JSON.parse(tokenContent);
  } catch (err) {
    console.error('Failed to read gift_list_token.json:', err);
    return res.status(500).json({ error: 'Failed to read gift_list_token.json' });
  }

  const record = giftList.find((r) => r.token === token);
  if (!record) return res.status(404).json({ error: 'Invalid token' });
  if (record.expire < Date.now()) return res.status(403).json({ error: 'Token expired' });

  return res.status(200).json(record);
}
