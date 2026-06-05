const fetch = (...args) => import('node-fetch').then(({default: f}) => f(...args));
const FIREBASE_URL = process.env.FIREBASE_URL;

exports.handler = async function(event) {
  try {
    const subRes = await fetch(FIREBASE_URL + '/subscription.json');
    const sub = await subRes.json();
    const actRes = await fetch(FIREBASE_URL + '/activityLog.json');
    const actData = await actRes.json();
    const today = new Date().toLocaleDateString('en-GB');
    let draftsToday = 0, emailsProcessed = 0, lastActivity = null;
    if (actData) {
      const entries = Object.values(actData);
      const todayEntries = entries.filter(e => e.date === today);
      draftsToday = todayEntries.filter(e => e.action === 'Draft created').length;
      emailsProcessed = todayEntries.filter(e => e.action === 'Email processed').length;
      const sorted = entries.sort((a, b) => (b.timestamp || 0) - (a.timestamp || 0));
      if (sorted.length) lastActivity = sorted[0].label || sorted[0].action || '—';
    }
    const hasSubscription = sub && sub.id;
    const expiry = sub && sub.expiry ? new Date(sub.expiry) : null;
    const now = new Date();
    const isExpired = expiry ? expiry < now : true;
    const expiresIn = expiry ? Math.floor((expiry - now) / (1000 * 60 * 60)) : 0;
    if (!hasSubscription || isExpired) {
      return { statusCode: 200, headers: { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' }, body: JSON.stringify({ active: false, error: isExpired ? 'Webhook subscription expired — please renew' : 'No subscription found' }) };
    }
    return { statusCode: 200, headers: { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' }, body: JSON.stringify({ active: true, draftsToday, emailsProcessed, lastActivity: lastActivity || 'No activity today', subscriptionId: sub.id, expiry: expiry ? expiry.toLocaleDateString('en-GB') : '—', expiresIn, expiryWarning: expiresIn < 24 }) };
  } catch(e) {
    return { statusCode: 200, headers: { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' }, body: JSON.stringify({ active: false, error: e.message }) };
  }
};
