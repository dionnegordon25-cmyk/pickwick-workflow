// =====================================================
// DG LETTINGS — CREATE GRAPH API WEBHOOK SUBSCRIPTION
// Registers inbox watcher on maintenance@pickwickestates.com
// Run once to set up, then renew-subscription keeps it alive
// =====================================================

const fetch = (...args) => import('node-fetch').then(({default: f}) => f(...args));

const CLIENT_ID     = process.env.AZURE_CLIENT_ID;
const CLIENT_SECRET = process.env.AZURE_CLIENT_SECRET;
const REFRESH_TOKEN = process.env.OUTLOOK_REFRESH_TOKEN;
const SITE_URL      = process.env.SITE_URL; // e.g. https://pickwick-maintenance1.netlify.app

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

exports.handler = async function(event) {
  try {
    const token = await getAccessToken();

    // Set expiry to 3 days from now (max allowed by Microsoft)
    const expiry = new Date(Date.now() + 3 * 24 * 60 * 60 * 1000).toISOString();

    const subscription = {
      changeType: 'created',
      notificationUrl: SITE_URL + '/.netlify/functions/email-handler',
      resource: 'me/mailFolders/inbox/messages',
      expirationDateTime: expiry,
      clientState: 'dg-lettings-secret-2026',
    };

    const res = await fetch('https://graph.microsoft.com/v1.0/subscriptions', {
      method: 'POST',
      headers: {
        'Authorization': 'Bearer ' + token,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(subscription),
    });

    const data = await res.json();

    if (data.id) {
      // Save subscription ID to Firebase for renewal
      const FIREBASE_URL = process.env.FIREBASE_URL;
      await fetch(FIREBASE_URL + '/subscription.json', {
        method: 'PUT',
        body: JSON.stringify({ id: data.id, expiry: data.expirationDateTime }),
      });

      return {
        statusCode: 200,
        body: JSON.stringify({ message: 'Subscription created', id: data.id, expiry: data.expirationDateTime }),
      };
    } else {
      return { statusCode: 400, body: JSON.stringify({ error: data }) };
    }
  } catch(e) {
    return { statusCode: 500, body: JSON.stringify({ error: e.message }) };
  }
};
