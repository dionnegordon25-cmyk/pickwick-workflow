// =====================================================
// DG LETTINGS — INBOUND EMAIL HANDLER
// Triggered by Microsoft Graph webhook when email arrives
// Intent-driven — works regardless of sender type
// Firebase auth via Database Secret (server-side only)
// =====================================================

const fetch = (...args) => import('node-fetch').then(({default: f}) => f(...args));

const TENANT_ID    = 'f22402ce-b358-43c7-91f9-b90742bf68e4';
const MAILBOX      = 'maintenance@pickwickestates.com';
const FIREBASE_URL = process.env.FIREBASE_URL;
const FIREBASE_SECRET = process.env.FIREBASE_SECRET; // Database secret for server auth
const FROM_NAME    = 'Dionne \u2014 Pickwick Estates';
const FROM_EMAIL   = process.env.OUTLOOK_EMAIL || MAILBOX;

// App-only auth for Microsoft Graph
async function getAccessToken() {
  const body = new URLSearchParams({
    grant_type:    'client_credentials',
    client_id:     process.env.AZURE_CLIENT_ID,
    client_secret: process.env.AZURE_CLIENT_SECRET,
    scope:         'https://graph.microsoft.com/.default',
  });
  const res = await fetch(
    `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`,
    { method: 'POST', body }
  );
  const data = await res.json();
  if (!data.access_token) throw new Error('Token error: ' + JSON.stringify(data));
  return data.access_token;
}

// Load Firebase data — authenticated via Database Secret
async function loadFirebase(path) {
  const auth = FIREBASE_SECRET ? `?auth=${FIREBASE_SECRET}` : '';
  const res = await fetch(FIREBASE_URL + '/' + path + '.json' + auth);
  if (!res.ok) {
    console.error(`Firebase load failed for ${path}: ${res.status}`);
    return [];
  }
  const data = await res.json();
  if (!data) return [];
  // Firebase returns object with keys — convert to array
  return typeof data === 'object' && !Array.isArray(data) ? Object.values(data) : data;
}

// Write to Firebase — authenticated
async function writeFirebase(path, body, method = 'POST') {
  const auth = FIREBASE_SECRET ? `?auth=${FIREBASE_SECRET}` : '';
  return fetch(FIREBASE_URL + '/' + path + '.json' + auth, {
    method,
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });
}

// Read full email from Outlook
async function getEmail(token, messageId) {
  const res = await fetch(
    `https://graph.microsoft.com/v1.0/users/${MAILBOX}/messages/${messageId}`,
    { headers: { 'Authorization': 'Bearer ' + token } }
  );
  return await res.json();
}

// Create draft in Outlook mailbox
async function createDraft(token, { to, subject, body }) {
  const res = await fetch(
    `https://graph.microsoft.com/v1.0/users/${MAILBOX}/messages`,
    {
      method: 'POST',
      headers: {
        'Authorization': 'Bearer ' + token,
        'Content-Type':  'application/json',
      },
      body: JSON.stringify({
        subject,
        body:         { contentType: 'Text', content: body },
        toRecipients: to ? [{ emailAddress: { address: to } }] : [],
        isDraft:      true,
      }),
    }
  );
  return await res.json();
}

// Identify who sent the email
function identifySender(fromEmail, fromName, properties, contractors) {
  const emailLower = (fromEmail || '').toLowerCase();
  const nameLower  = (fromName  || '').toLowerCase();

  const contractor = contractors.find(c =>
    (c.email || '').toLowerCase() === emailLower ||
    nameLower.includes((c.name || '').toLowerCase().split(' ')[0])
  );
  if (contractor) return { type: 'contractor', match: contractor };

  const llProp = properties.find(p =>
    (p.llEmail || '').toLowerCase() === emailLower ||
    (p.llName  || '').toLowerCase().includes(nameLower.split(' ')[0])
  );
  if (llProp) return { type: 'landlord', match: llProp };

  const tenantProp = properties.find(p =>
    (p.tenantEmail || '').toLowerCase() === emailLower ||
    (p.tenant      || '').toLowerCase().includes(nameLower.split(' ')[0])
  );
  if (tenantProp) return { type: 'tenant', match: tenantProp };

  return { type: 'unknown', match: null };
}

// Find property mentioned in email text
function findPropertyInEmail(text, properties) {
  const textLower = text.toLowerCase();
  return properties.find(p => {
    const parts = p.address.toLowerCase().split(',')[0].split(' ');
    return parts.length >= 2 && textLower.includes(parts.slice(0, 2).join(' '));
  }) || null;
}

// Detect intent from subject + body
function detectIntent(subject, body) {
  const text = (subject + ' ' + body).toLowerCase();
  if (text.includes('boiler') || text.includes('heating') || text.includes('hot water'))  return 'boiler';
  if (text.includes('leak')   || text.includes('water damage') || text.includes('damp'))  return 'leak';
  if (text.includes('electric') || text.includes('power') || text.includes('fuse'))       return 'electrical';
  if (text.includes('lock')   || text.includes('key') || text.includes('door'))           return 'access';
  if (text.includes('quote')  || text.includes('estimate') || text.includes('cost'))      return 'quote';
  if (text.includes('invoice') || text.includes('payment') || text.includes('bill'))      return 'invoice';
  if (text.includes('complete') || text.includes('finished') || text.includes('done'))    return 'complete';
  if (text.includes('approve') || text.includes('authorise') || text.includes('go ahead')) return 'approval';
  if (text.includes('deposit') || text.includes('refund'))                                return 'deposit';
  if (text.includes('notice')  || text.includes('leaving') || text.includes('vacate'))    return 'notice';
  return 'general';
}

// Generate draft emails based on intent
async function generateDrafts(emailData, senderInfo, property, intent) {
  const { fromEmail, fromName, subject, body } = emailData;
  const nl = '\n';
  const drafts = [];

  if (!property) {
    drafts.push({
      to: FROM_EMAIL,
      subject: '[ACTION NEEDED] Unmatched email \u2014 ' + subject,
      body: 'An email was received that could not be automatically matched to a property or contact.' + nl + nl +
        'From: ' + fromName + ' <' + fromEmail + '>' + nl +
        'Subject: ' + subject + nl + nl +
        'Please review and respond manually.' + nl + nl +
        'Original message:' + nl + body.substring(0, 500),
      label: 'Unmatched email \u2014 manual review needed'
    });
    return drafts;
  }

  const llFirst     = property.llName ? property.llName.split(' ')[0]   : 'Landlord';
  const tenFirst    = property.tenant ? property.tenant.split(' ')[0]   : 'Resident';
  const senderFirst = fromName        ? fromName.split(' ')[0]          : 'there';
  const addr        = property.address;
  const isTenant    = senderInfo.type === 'tenant';
  const isLandlord  = senderInfo.type === 'landlord';
  const isMaintenanceIntent = ['boiler','leak','electrical','access','general'].includes(intent);

  if (isMaintenanceIntent) {
    // 1. Acknowledge sender
    drafts.push({
      to: fromEmail,
      subject: 'Re: ' + subject,
      body: 'Dear ' + senderFirst + ',' + nl + nl +
        'Thank you for getting in touch regarding ' + addr + '.' + nl + nl +
        'We have received your message and will look into this as a matter of priority.' + nl + nl +
        (isTenant
          ? 'We will arrange for a contractor to attend and will be in touch shortly to confirm a date and time. If the matter is urgent please do not hesitate to call us directly.'
          : 'We will arrange for a contractor to attend and will keep you updated on progress.') + nl + nl +
        'Kind regards,' + nl + FROM_NAME,
      label: 'Acknowledge sender \u2014 ' + senderInfo.type
    });

    // 2. Notify landlord
    if (!isLandlord && property.llEmail) {
      drafts.push({
        to: property.llEmail,
        subject: 'Maintenance issue reported \u2014 ' + addr,
        body: 'Dear ' + llFirst + ',' + nl + nl +
          'I am writing to advise that a ' + intent + ' issue has been reported at ' + addr + '.' + nl + nl +
          'Reported by: ' + fromName + nl +
          'Details: ' + body.substring(0, 300) + (body.length > 300 ? '...' : '') + nl + nl +
          'We are arranging for a contractor to attend. I will keep you updated and will seek your approval before any significant works are instructed.' + nl + nl +
          'Kind regards,' + nl + FROM_NAME,
        label: 'Notify landlord \u2014 maintenance reported'
      });
    }

    // 3. Notify tenant
    if (!isTenant && property.tenantEmail) {
      drafts.push({
        to: property.tenantEmail,
        subject: 'Maintenance update \u2014 ' + addr,
        body: 'Dear ' + tenFirst + ',' + nl + nl +
          'I am writing to advise that we have been notified of a ' + intent + ' issue at your property and are arranging for a contractor to attend.' + nl + nl +
          'We will be in touch shortly to confirm a date and time.' + nl + nl +
          'Kind regards,' + nl + FROM_NAME,
        label: 'Notify tenant \u2014 contractor being arranged'
      });
    }

    // 4. Internal action draft
    drafts.push({
      to: '',
      subject: 'INSTRUCT CONTRACTOR \u2014 ' + intent.toUpperCase() + ' \u2014 ' + addr,
      body: 'ACTION REQUIRED \u2014 please instruct appropriate contractor.' + nl + nl +
        'Property:    ' + addr + nl +
        'Issue type:  ' + intent + nl +
        'Reported by: ' + fromName + ' (' + senderInfo.type + ')' + nl +
        'Tenant:      ' + (property.tenant  || 'Not recorded') + nl +
        'Landlord:    ' + (property.llName  || 'Not recorded') + nl + nl +
        'Original message:' + nl + body.substring(0, 400) + (body.length > 400 ? '...' : '') + nl + nl +
        'Kind regards,' + nl + FROM_NAME,
      label: 'Internal \u2014 instruct contractor'
    });
  }

  else if (intent === 'quote') {
    if (property.llEmail) {
      drafts.push({
        to: property.llEmail,
        subject: 'Quote received \u2014 ' + addr,
        body: 'Dear ' + llFirst + ',' + nl + nl +
          'Please find below a quote received from ' + fromName + ' for works at ' + addr + ':' + nl + nl +
          body.substring(0, 500) + (body.length > 500 ? '...' : '') + nl + nl +
          'Please could you confirm whether you are happy to proceed?' + nl + nl +
          'Kind regards,' + nl + FROM_NAME,
        label: 'Forward quote to landlord'
      });
    }
    drafts.push({
      to: fromEmail,
      subject: 'Re: ' + subject,
      body: 'Hi ' + senderFirst + ',' + nl + nl +
        'Thank you for sending over the quote for ' + addr + '.' + nl + nl +
        'I have forwarded this to the landlord for approval and will be in touch as soon as I hear back.' + nl + nl +
        'Kind regards,' + nl + FROM_NAME,
      label: 'Acknowledge quote received'
    });
  }

  else if (intent === 'invoice') {
    if (property.llEmail) {
      drafts.push({
        to: property.llEmail,
        subject: 'Invoice received \u2014 ' + addr,
        body: 'Dear ' + llFirst + ',' + nl + nl +
          'Please find below an invoice received from ' + fromName + ' for works at ' + addr + ':' + nl + nl +
          body.substring(0, 400) + nl + nl +
          'Please could you confirm you are happy for this to be processed?' + nl + nl +
          'Kind regards,' + nl + FROM_NAME,
        label: 'Forward invoice to landlord'
      });
    }
    drafts.push({
      to: fromEmail,
      subject: 'Re: ' + subject,
      body: 'Hi ' + senderFirst + ',' + nl + nl +
        'Thank you for sending over the invoice for ' + addr + '.' + nl + nl +
        'I have passed this to the landlord for approval and will be in touch shortly.' + nl + nl +
        'Kind regards,' + nl + FROM_NAME,
      label: 'Acknowledge invoice received'
    });
  }

  else if (intent === 'complete') {
    if (property.llEmail) {
      drafts.push({
        to: property.llEmail,
        subject: 'Works completed \u2014 ' + addr,
        body: 'Dear ' + llFirst + ',' + nl + nl +
          fromName + ' has confirmed works are now complete at ' + addr + '.' + nl + nl +
          'Please let me know if you require any further information.' + nl + nl +
          'Kind regards,' + nl + FROM_NAME,
        label: 'Notify landlord \u2014 works complete'
      });
    }
    if (property.tenantEmail) {
      drafts.push({
        to: property.tenantEmail,
        subject: 'Maintenance update \u2014 works complete \u2014 ' + addr,
        body: 'Dear ' + tenFirst + ',' + nl + nl +
          'The contractor has confirmed that works are now complete at your property.' + nl + nl +
          'If you have any concerns please do not hesitate to get in touch.' + nl + nl +
          'Kind regards,' + nl + FROM_NAME,
        label: 'Notify tenant \u2014 works complete'
      });
    }
  }

  else if (intent === 'approval') {
    drafts.push({
      to: '',
      subject: 'Works approved \u2014 proceed \u2014 ' + addr,
      body: 'Hi,' + nl + nl +
        'The landlord has approved the works at ' + addr + '. Please could you confirm your availability?' + nl + nl +
        'Tenant: ' + (property.tenant || 'Not recorded') + nl +
        'Please arrange access directly with the tenant.' + nl + nl +
        'Kind regards,' + nl + FROM_NAME,
      label: 'Instruct contractor \u2014 works approved'
    });
    if (property.tenantEmail) {
      drafts.push({
        to: property.tenantEmail,
        subject: 'Maintenance update \u2014 ' + addr,
        body: 'Dear ' + tenFirst + ',' + nl + nl +
          'Works have been approved for your property. A contractor will be in touch shortly to arrange access.' + nl + nl +
          'Kind regards,' + nl + FROM_NAME,
        label: 'Notify tenant \u2014 works approved'
      });
    }
  }

  else if (intent === 'deposit') {
    drafts.push({
      to: fromEmail,
      subject: 'Re: ' + subject,
      body: 'Dear ' + senderFirst + ',' + nl + nl +
        'Thank you for your message regarding the deposit for ' + addr + '.' + nl + nl +
        'I will look into this and come back to you shortly.' + nl + nl +
        'Kind regards,' + nl + FROM_NAME,
      label: 'Acknowledge deposit query'
    });
  }

  else if (intent === 'notice') {
    drafts.push({
      to: fromEmail,
      subject: 'Re: ' + subject,
      body: 'Dear ' + senderFirst + ',' + nl + nl +
        'Thank you for your message regarding ' + addr + '.' + nl + nl +
        'I will review the details and come back to you shortly.' + nl + nl +
        'Kind regards,' + nl + FROM_NAME,
      label: 'Acknowledge notice / tenancy query'
    });
    if (!isLandlord && property.llEmail) {
      drafts.push({
        to: property.llEmail,
        subject: 'Tenancy update \u2014 ' + addr,
        body: 'Dear ' + llFirst + ',' + nl + nl +
          'We have received correspondence regarding ' + addr + ' which may require your attention.' + nl + nl +
          'Details: ' + body.substring(0, 300) + nl + nl +
          'Please let me know if you would like to discuss.' + nl + nl +
          'Kind regards,' + nl + FROM_NAME,
        label: 'Notify landlord \u2014 tenancy correspondence'
      });
    }
  }

  return drafts;
}

// Auto-log maintenance job to Firebase
async function logJob(property, subject, fromName, intent) {
  if (!property) return;
  await writeFirebase('jobs', {
    id:           Date.now(),
    prop:         property.address,
    tenant:       property.tenant || '',
    issue:        subject,
    status:       'Awaiting contractor',
    pill:         'pill-amber',
    contractor:   '',
    contractorId: '',
    access:       'Contact tenant to arrange',
    quote:        0,
    invoice:      0,
    notes:        'Auto-logged from email by ' + fromName,
    dateLogged:   new Date().toISOString().split('T')[0],
    lastActionDate: new Date().toISOString().split('T')[0],
    category:     intent,
  });
}

// Update automation stats
async function updateAutomationStats(newDrafts) {
  try {
    const auth = FIREBASE_SECRET ? `?auth=${FIREBASE_SECRET}` : '';
    const res = await fetch(FIREBASE_URL + '/automationStats.json' + auth);
    const current = (await res.json()) || {};
    const today = new Date().toLocaleDateString('en-GB');
    await writeFirebase('automationStats', {
      lastRun:         new Date().toISOString(),
      lastRunDate:     today,
      draftsToday:     (current.lastRunDate === today ? (current.draftsToday || 0) : 0) + newDrafts,
      draftsTotal:     (current.draftsTotal || 0) + newDrafts,
      emailsProcessed: (current.emailsProcessed || 0) + 1,
      status:          'active',
    }, 'PUT');
    console.log('Stats updated \u2014 drafts today:', newDrafts);
  } catch(e) {
    console.log('Stats update failed (non-critical):', e.message);
  }
}

// ── MAIN HANDLER ─────────────────────────────────────────────────
exports.handler = async function(event) {

  // Webhook validation handshake
  const params = new URLSearchParams(event.rawQuery || '');
  const validationToken = params.get('validationToken');
  if (validationToken) {
    console.log('Validation handshake');
    return { statusCode: 200, headers: { 'Content-Type': 'text/plain' }, body: validationToken };
  }

  if (!event.body) return { statusCode: 200, body: 'No body' };

  try {
    const notification = JSON.parse(event.body);
    const values = notification.value || [];
    console.log('Notifications:', values.length);
    if (!values.length) return { statusCode: 200, body: 'No notifications' };

    const token = await getAccessToken();
    const [properties, contractors] = await Promise.all([
      loadFirebase('properties'),
      loadFirebase('contractors'),
    ]);
    console.log('Firebase loaded \u2014 properties:', properties.length, 'contractors:', contractors.length);

    const results = [];
    const processedIds = new Set();

    for (const notif of values) {
      if (notif.clientState !== 'dg-lettings-secret-2026') { console.log('Bad clientState'); continue; }
      const messageId = notif.resourceData && notif.resourceData.id;
      if (!messageId) continue;
      if (processedIds.has(messageId)) { console.log('Duplicate skip'); continue; }
      processedIds.add(messageId);

      const email = await getEmail(token, messageId);
      if (!email || !email.from) continue;

      const fromEmail = email.from.emailAddress.address;
      const fromName  = email.from.emailAddress.name || fromEmail;
      const subject   = email.subject || '(No subject)';
      const body      = email.body ? email.body.content.replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim() : '';

      if (fromEmail.toLowerCase() === MAILBOX.toLowerCase()) { console.log('Own email skip'); continue; }

      const senderInfo = identifySender(fromEmail, fromName, properties, contractors);
      let property = senderInfo.match && senderInfo.type !== 'contractor' ? senderInfo.match : null;
      if (!property) property = findPropertyInEmail(subject + ' ' + body, properties);

      const intent = detectIntent(subject, body);
      console.log('Email:', subject, '| Sender:', senderInfo.type, '| Intent:', intent, '| Property:', property ? property.address : 'none');

      const drafts = await generateDrafts({ fromEmail, fromName, subject, body }, senderInfo, property, intent);

      for (const draft of drafts) {
        const result = await createDraft(token, draft);
        console.log('Draft created:', result.id ? result.id.substring(0,20) : JSON.stringify(result.error));
      }

      if (property && ['boiler','leak','electrical','access','general'].includes(intent)) {
        await logJob(property, subject, fromName, intent);
      }

      if (drafts.length > 0) await updateAutomationStats(drafts.length);

      results.push({ from: fromName, type: senderInfo.type, property: property ? property.address : 'unmatched', intent, draftsCreated: drafts.length });
    }

    return { statusCode: 200, body: JSON.stringify({ processed: results.length, results }) };

  } catch(e) {
    console.error('Handler error:', e.message);
    return { statusCode: 200, body: JSON.stringify({ error: e.message }) };
  }
};
