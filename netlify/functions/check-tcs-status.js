// netlify/functions/check-tcs-status.js
//
// Manual fallback: lets staff click "Check status" on an offer to poll DocuSign
// directly, in case the Connect webhook hasn't fired yet (e.g. Connect isn't
// set up yet, or there's a delay). Called as:
//   GET /.netlify/functions/check-tcs-status?envelopeId=xxx
//
// Uses the same JWT auth as send-tcs-envelope.js.

const fetch = (...args) => import('node-fetch').then(({ default: f }) => f(...args));
const jwt = require('jsonwebtoken');

const {
  DOCUSIGN_INTEGRATION_KEY,
  DOCUSIGN_USER_ID,
  DOCUSIGN_ACCOUNT_ID,
  DOCUSIGN_BASE_URI,
  DOCUSIGN_PRIVATE_KEY,
  FIREBASE_URL,
  FIREBASE_SECRET
} = process.env;

const DOCUSIGN_AUTH_SERVER = DOCUSIGN_BASE_URI && DOCUSIGN_BASE_URI.includes('demo')
  ? 'account-d.docusign.com'
  : 'account.docusign.com';

async function getAccessToken() {
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
    DOCUSIGN_PRIVATE_KEY,
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
    return { statusCode: 400, body: JSON.stringify({ error: 'Missing envelopeId' }) };
  }

  try {
    const accessToken = await getAccessToken();

    const res = await fetch(
      `${DOCUSIGN_BASE_URI}/restapi/v2.1/accounts/${DOCUSIGN_ACCOUNT_ID}/envelopes/${envelopeId}`,
      { headers: { Authorization: `Bearer ${accessToken}` } }
    );
    const envelope = await res.json();
    if (!res.ok) throw new Error(`Envelope lookup failed: ${JSON.stringify(envelope)}`);

    if (envelope.status === 'completed') {
      const lookupRes = await fetch(
        `${FIREBASE_URL}/properties.json?orderBy="checklist/tcs/envelopeId"&equalTo="${envelopeId}"&auth=${FIREBASE_SECRET}`
      );
      const matches = await lookupRes.json();
      const propertyId = matches ? Object.keys(matches)[0] : null;
      if (propertyId) {
        await fetch(
          `${FIREBASE_URL}/properties/${propertyId}/checklist/tcs.json?auth=${FIREBASE_SECRET}`,
          {
            method: 'PATCH',
            body: JSON.stringify({
              status: 'done',
              signedDate: new Date().toISOString().slice(0, 10),
              signedDocUrl: `${DOCUSIGN_BASE_URI}/restapi/v2.1/accounts/${DOCUSIGN_ACCOUNT_ID}/envelopes/${envelopeId}/documents/combined`
            })
          }
        );
      }
    }

    return {
      statusCode: 200,
      body: JSON.stringify({ status: envelope.status, envelopeId })
    };
  } catch (err) {
    console.error('check-tcs-status error:', err);
    return { statusCode: 500, body: JSON.stringify({ error: err.message }) };
  }
};
