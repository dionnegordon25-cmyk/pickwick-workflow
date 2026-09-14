// netlify/functions/get-signed-doc.js
//
// Serves the signed T&Cs PDF for a given envelope. DocuSign's document
// endpoint requires a valid Bearer access token on every request — a plain
// link straight to that URL fails with PARTNER_AUTHENTICATION_FAILED because
// the browser has no way to attach that token. This function does the JWT
// auth server-side (same as the other DocuSign functions), fetches the PDF
// bytes, and returns them to the browser as a normal downloadable file.
//
// Called as: GET /.netlify/functions/get-signed-doc?envelopeId=xxx

const fetch = (...args) => import('node-fetch').then(({ default: f }) => f(...args));
const jwt = require('jsonwebtoken');

const {
  DOCUSIGN_INTEGRATION_KEY,
  DOCUSIGN_USER_ID,
  DOCUSIGN_ACCOUNT_ID,
  DOCUSIGN_BASE_URI,
  FIREBASE_URL,
  FIREBASE_SECRET
} = process.env;

const DOCUSIGN_AUTH_SERVER = DOCUSIGN_BASE_URI && DOCUSIGN_BASE_URI.includes('demo')
  ? 'account-d.docusign.com'
  : 'account.docusign.com';

async function getPrivateKey() {
  const res = await fetch(`${FIREBASE_URL}/config/docusignPrivateKey.json?auth=${FIREBASE_SECRET}`);
  const raw = await res.json();
  if (!raw) throw new Error('DocuSign private key not found in Firebase at /config/docusignPrivateKey');
  return normalisePemKey(raw);
}

function normalisePemKey(raw) {
  let key = raw.replace(/\\n/g, '\n').trim();
  const match = key.match(/-----BEGIN ([^-]+)-----([\s\S]*?)-----END \1-----/);
  if (match) {
    const label = match[1].trim();
    const body = match[2].replace(/\s+/g, '');
    const wrapped = body.match(/.{1,64}/g).join('\n');
    key = `-----BEGIN ${label}-----\n${wrapped}\n-----END ${label}-----\n`;
  }
  return key;
}

async function getAccessToken() {
  const privateKey = await getPrivateKey();
  const now = Math.floor(Date.now() / 1000);
  const token = jwt.sign(
    {
      iss: DOCUSIGN_INTEGRATION_KEY,
      sub: DOCUSIGN_USER_ID,
      aud: DOCUSIGN_AUTH_SERVER,
      iat: now,
      exp: now + 3600,
      scope: 'signature impersonation'
    },
    privateKey,
    { algorithm: 'RS256' }
  );

  const res = await fetch(`https://${DOCUSIGN_AUTH_SERVER}/oauth/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: `grant_type=urn:ietf:params:oauth:grant-type:jwt-bearer&assertion=${token}`
  });

  const data = await res.json();
  if (!res.ok) throw new Error(`DocuSign auth failed: ${JSON.stringify(data)}`);
  return data.access_token;
}

exports.handler = async (event) => {
  if (event.httpMethod !== 'GET') {
    return { statusCode: 405, body: 'Method not allowed' };
  }

  const envelopeId = event.queryStringParameters && event.queryStringParameters.envelopeId;
  if (!envelopeId) {
    return { statusCode: 400, body: 'Missing envelopeId' };
  }

  try {
    const accessToken = await getAccessToken();

    const res = await fetch(
      `${DOCUSIGN_BASE_URI}/restapi/v2.1/accounts/${DOCUSIGN_ACCOUNT_ID}/envelopes/${envelopeId}/documents/combined`,
      { headers: { Authorization: `Bearer ${accessToken}` } }
    );

    if (!res.ok) {
      const text = await res.text();
      throw new Error(`DocuSign document fetch failed (${res.status}): ${text}`);
    }

    const arrayBuffer = await res.arrayBuffer();
    const base64 = Buffer.from(arrayBuffer).toString('base64');

    return {
      statusCode: 200,
      headers: {
        'Content-Type': 'application/pdf',
        'Content-Disposition': `inline; filename="landlord-tcs-${envelopeId}.pdf"`
      },
      body: base64,
      isBase64Encoded: true
    };
  } catch (err) {
    console.error('get-signed-doc error:', err);
    return { statusCode: 500, body: JSON.stringify({ error: err.message }) };
  }
};
