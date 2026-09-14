// netlify/functions/send-tcs-envelope.js
//
// Creates and sends a DocuSign envelope for the Landlord T&Cs document.
// Called from the "Send T&Cs" button on the Offer Accepted checklist.
//
// AUTH: uses JWT Grant (server-to-server, no landlord/staff login needed) —
// matches the "Service Integration" app you set up in DocuSign.
// https://developers.docusign.com/platform/auth/jwt/
//
// ENV VARS REQUIRED (Netlify > Site configuration > Environment variables):
//   DOCUSIGN_INTEGRATION_KEY   - Integration Key from the app's General Info page
//   DOCUSIGN_USER_ID           - User ID, from Admin > Apps and Keys > My Account Information
//   DOCUSIGN_ACCOUNT_ID        - API Account ID, same page
//   DOCUSIGN_BASE_URI          - Account Base URI, same page
//   DOCUSIGN_TEMPLATE_ID       - the Landlord T&Cs DocuSign Template ID
//   FIREBASE_URL, FIREBASE_SECRET - already in use elsewhere in the CRM
//
// NOTE: the RSA private key is deliberately NOT an environment variable.
// AWS Lambda caps the combined size of all env vars at 4KB, and this key
// alone is large enough to push every function over that limit. Instead,
// it's stored in Firebase (same trust boundary as everything else in the
// CRM, secured by FIREBASE_SECRET) and fetched at runtime — see
// getPrivateKey() below. Stored at: /config/docusignPrivateKey
//
// This assumes a DocuSign Template exists for the Landlord T&Cs with a signer
// role called "Landlord" and the signature tab already placed on it. Using a
// Template keeps the document wording editable in DocuSign directly, rather
// than needing a code change every time the wording changes.

const fetch = (...args) => import('node-fetch').then(({ default: f }) => f(...args));
const jwt = require('jsonwebtoken');

const {
  DOCUSIGN_INTEGRATION_KEY,
  DOCUSIGN_USER_ID,
  DOCUSIGN_ACCOUNT_ID,
  DOCUSIGN_BASE_URI,
  DOCUSIGN_TEMPLATE_ID,
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

// Rebuilds a proper PEM key regardless of how its line breaks survived being
// pasted into Firebase's single-line value field (common failure mode: real
// newlines get stripped, or turned into literal "\n" text). Reconstructs the
// standard -----BEGIN...-----/body/-----END...----- structure from scratch.
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
  if (event.httpMethod !== 'POST') {
    return { statusCode: 405, body: 'Method not allowed' };
  }

  try {
    const { propertyId, landlordName, landlordEmail, propertyAddress } = JSON.parse(event.body);

    if (!propertyId || !landlordName || !landlordEmail) {
      return { statusCode: 400, body: 'Missing propertyId, landlordName, or landlordEmail' };
    }
    if (!DOCUSIGN_TEMPLATE_ID) {
      return { statusCode: 400, body: JSON.stringify({ error: 'DOCUSIGN_TEMPLATE_ID not set yet — create the Landlord T&Cs template in DocuSign first' }) };
    }

    const accessToken = await getAccessToken();

    const envelopeDefinition = {
      templateId: DOCUSIGN_TEMPLATE_ID,
      templateRoles: [
        {
          roleName: 'Landlord',
          name: landlordName,
          email: landlordEmail,
          tabs: {
            textTabs: [
              { tabLabel: 'PropertyAddress', value: propertyAddress || '' }
            ]
          }
        }
      ],
      status: 'sent'
    };

    const sendRes = await fetch(
      `${DOCUSIGN_BASE_URI}/restapi/v2.1/accounts/${DOCUSIGN_ACCOUNT_ID}/envelopes`,
      {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${accessToken}`,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify(envelopeDefinition)
      }
    );

    const envelope = await sendRes.json();
    if (!sendRes.ok) throw new Error(`Envelope send failed: ${JSON.stringify(envelope)}`);

    // Record sent status against the PROPERTY record (matches the CRM's schema —
    // properties/{propertyId}/checklist/tcs) — T&Cs, PEP check, and the questionnaire
    // all live on the property now, not on any one tenancy offer, so the checklist
    // shows "Sent — awaiting signature" immediately.
    await fetch(
      `${FIREBASE_URL}/properties/${propertyId}/checklist/tcs.json?auth=${FIREBASE_SECRET}`,
      {
        method: 'PATCH',
        body: JSON.stringify({
          status: 'sent',
          envelopeId: envelope.envelopeId,
          sentDate: new Date().toISOString().slice(0, 10)
        })
      }
    );

    return {
      statusCode: 200,
      body: JSON.stringify({ success: true, envelopeId: envelope.envelopeId })
    };
  } catch (err) {
    console.error('send-tcs-envelope error:', err);
    return { statusCode: 500, body: JSON.stringify({ error: err.message }) };
  }
};
