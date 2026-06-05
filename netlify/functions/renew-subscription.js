// @netlify schedule: 0 6 */2 * *
// =====================================================
// DG LETTINGS — RENEW GRAPH API WEBHOOK SUBSCRIPTION
// Runs every 2 days via scheduler to keep inbox watcher alive
// Microsoft subscriptions expire every 3 days
// =====================================================

const fetch = (...args) => import('node-fetch').then(({default: f}) => f(...args));

const CLIENT_ID     = process.env.AZURE_CLIENT_ID;
const CLIENT_SECRET = process.env.AZURE_CLIENT_SECRET;
const REFRESH_TOKEN = process.env.OUTLOOK_REFRESH_TOKEN;
const FIREBASE_URL  = process.env.FIREBASE_URL;

async function getAccessToken() {
  const body = new URLSearchParams({
    grant_type:    'refresh_token',
    client_id:     CLIENT_ID,
    client_secret: CLIENT_SECRET,
    refresh_token: REFRESH_TOKEN,
    scope:         'https://graph.microsoft.com/Mail.ReadWrite https://graph.microsoft.com/Mail.Read offline_access',
  });
  const res = await fetch('https://login.microsoftonline.com/f22402ce-b358-43c7-91f9-b90742bf68e4/oauth2/v2.0/token', { method: 'POST', body });
  const data = await res.json();
  if (!data.access_token) throw new Error('Token error: ' + JSON.stringify(data));
  return data.access_token;
}

exports.handler = async function() {
  try {
    // Get subscription ID from Firebase
    const subRes = await fetch(FIREBASE_URL + '/subscription.json');
    const subData = await subRes.json();

    if (!subData || !subData.id) {
      return { statusCode: 400, body: JSON.stringify({ error: 'No subscription found. Run create-subscription first.' }) };
    }

    const token = await getAccessToken();
    const newExpiry = new Date(Date.now() + 3 * 24 * 60 * 60 * 1000).toISOString();

    const res = await fetch('https://graph.microsoft.com/v1.0/subscriptions/' + subData.id, {
      method: 'PATCH',
      headers: {
        'Authorization': 'Bearer ' + token,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({ expirationDateTime: newExpiry }),
    });

    const data = await res.json();

    if (data.id) {
      // Update expiry in Firebase
      await fetch(FIREBASE_URL + '/subscription.json', {
        method: 'PUT',
        body: JSON.stringify({ id: data.id, expiry: data.expirationDateTime }),
      });
      return { statusCode: 200, body: JSON.stringify({ message: 'Subscription renewed', expiry: data.expirationDateTime }) };
    } else {
      // Subscription may have expired - recreate it
      return { statusCode: 400, body: JSON.stringify({ error: 'Renewal failed - run create-subscription', data }) };
    }
  } catch(e) {
    return { statusCode: 500, body: JSON.stringify({ error: e.message }) };
  }
};
