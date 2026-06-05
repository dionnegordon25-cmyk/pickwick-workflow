const fetch = (...args) => import('node-fetch').then(({default: f}) => f(...args));

const TENANT_ID    = 'f22402ce-b358-43c7-91f9-b90742bf68e4';
const MAILBOX      = 'maintenance@pickwickestates.com';
const FIREBASE_URL = process.env.FIREBASE_URL;

async function getAccessToken() {
  const body = new URLSearchParams({
    grant_type:    'client_credentials',
    client_id:     process.env.AZURE_CLIENT_ID,
    client_secret: process.env.AZURE_CLIENT_SECRET,
    scope:         'https://graph.microsoft.com/.default',
  });
  const res = await fetch(`https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`, { method: 'POST', body });
  const data = await res.json();
  if (!data.access_token) throw new Error('Token error: ' + JSON.stringify(data));
  return data.access_token;
}

exports.handler = async function(event) {
  try {
    const subRes = await fetch(FIREBASE_URL + '/subscription.json');
    const sub = await subRes.json();
    const token = await getAccessToken();
    const expiry = new Date(Date.now() + 3 * 24 * 60 * 60 * 1000).toISOString();
    let result, action;

    if (sub && sub.id) {
      const res = await fetch(`https://graph.microsoft.com/v1.0/subscriptions/${sub.id}`, {
        method: 'PATCH',
        headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
        body: JSON.stringify({ expirationDateTime: expiry }),
      });
      result = await res.json();
      action = result.id ? 'renewed' : 'failed';
    }

    if (!result || !result.id) {
      // Create fresh if renewal failed
      const res = await fetch('https://graph.microsoft.com/v1.0/subscriptions', {
        method: 'POST',
        headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
        body: JSON.stringify({
          changeType: 'created',
          notificationUrl: process.env.SITE_URL + '/.netlify/functions/email-handler',
          resource: `users/${MAILBOX}/messages`,
          expirationDateTime: expiry,
          clientState: 'dg-lettings-secret-2026',
        }),
      });
      result = await res.json();
      action = 'recreated';
    }

    if (!result.id) throw new Error('Failed: ' + JSON.stringify(result));

    await fetch(FIREBASE_URL + '/subscription.json', {
      method: 'PUT',
      body: JSON.stringify({ id: result.id, expiry: result.expirationDateTime, lastRenewed: new Date().toISOString() }),
    });

    return { statusCode: 200, body: JSON.stringify({ message
