// netlify/functions/docusign-webhook.js
//
// Receives status updates from DocuSign Connect whenever an envelope changes
// status: sent, delivered, viewed, completed, declined, voided. This is what
// lets the Offer Accepted checklist and Dashboard "outstanding signatures"
// strip update automatically, without anyone checking DocuSign manually.
//
// SETUP REQUIRED IN DOCUSIGN (once API access + template are ready):
//   Admin > Connect > Add Configuration
//   - URL to publish to: https://<your-netlify-site>/.netlify/functions/docusign-webhook
//   - Envelope events to track: Sent, Delivered, Completed, Declined, Voided
//   - Include documents: Yes
//
// ENV VARS REQUIRED (same as send-tcs-envelope.js):
//   FIREBASE_URL, FIREBASE_SECRET, DOCUSIGN_BASE_URI, DOCUSIGN_ACCOUNT_ID
//   DOCUSIGN_HMAC_KEY - optional but recommended: verifies the payload really
//     came from DocuSign (set up under the Connect configuration's HMAC settings).

const fetch = (...args) => import('node-fetch').then(({ default: f }) => f(...args));
const crypto = require('crypto');

const { FIREBASE_URL, FIREBASE_SECRET, DOCUSIGN_HMAC_KEY, DOCUSIGN_BASE_URI, DOCUSIGN_ACCOUNT_ID } = process.env;

function verifyHmac(rawBody, signatureHeader) {
  if (!DOCUSIGN_HMAC_KEY) return true; // skip verification until HMAC is configured
  const computed = crypto
    .createHmac('sha256', DOCUSIGN_HMAC_KEY)
    .update(rawBody)
    .digest('base64');
  return computed === signatureHeader;
}

exports.handler = async (event) => {
  if (event.httpMethod !== 'POST') {
    return { statusCode: 405, body: 'Method not allowed' };
  }

  const signature = event.headers['x-docusign-signature-1'];
  if (!verifyHmac(event.body, signature)) {
    console.warn('DocuSign webhook: HMAC verification failed — rejecting');
    return { statusCode: 401, body: 'Invalid signature' };
  }

  try {
    const payload = JSON.parse(event.body);

    // Confirm the exact field paths against your Connect configuration once
    // it's live — the shape depends on which Connect format (legacy JSON vs
    // aggregate) you select when setting it up.
    const envelopeId = payload.envelopeId || (payload.data && payload.data.envelopeId);
    const status = payload.status || (payload.data && payload.data.envelopeSummary && payload.data.envelopeSummary.status);

    if (!envelopeId) {
      return { statusCode: 400, body: 'No envelopeId in payload' };
    }

    // Find which PROPERTY this envelope belongs to (checklist.tcs.envelopeId) —
    // T&Cs live on the property record now, not on any one tenancy offer.
    const lookupRes = await fetch(
      `${FIREBASE_URL}/properties.json?orderBy="checklist/tcs/envelopeId"&equalTo="${envelopeId}"&auth=${FIREBASE_SECRET}`
    );
    const matches = await lookupRes.json();
    const propertyId = matches ? Object.keys(matches)[0] : null;

    if (!propertyId) {
      console.warn(`No property found for envelopeId ${envelopeId}`);
      return { statusCode: 200, body: 'No matching property — ignored' };
    }

    const update = {
      status: status === 'completed' ? 'done' : status,
      lastUpdated: new Date().toISOString()
    };

    if (status === 'completed') {
      update.signedDate = new Date().toISOString().slice(0, 10);
      // Store the DocuSign documents endpoint rather than duplicating the
      // signed PDF into Firebase — avoids managing a second copy of a signed
      // legal document. Staff open this link straight from the checklist.
      update.signedDocUrl =
        `${DOCUSIGN_BASE_URI}/restapi/v2.1/accounts/${DOCUSIGN_ACCOUNT_ID}/envelopes/${envelopeId}/documents/combined`;
    }

    await fetch(
      `${FIREBASE_URL}/properties/${propertyId}/checklist/tcs.json?auth=${FIREBASE_SECRET}`,
      { method: 'PATCH', body: JSON.stringify(update) }
    );

    return { statusCode: 200, body: 'OK' };
  } catch (err) {
    console.error('docusign-webhook error:', err);
    return { statusCode: 500, body: JSON.stringify({ error: err.message }) };
  }
};
