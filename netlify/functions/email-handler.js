// =====================================================
// DG LETTINGS — INBOUND EMAIL HANDLER
// Triggered by Microsoft Graph webhook when email arrives
// Uses app-only auth (client credentials) — matches create-subscription.js
// =====================================================

const fetch = (...args) => import('node-fetch').then(({default: f}) => f(...args));

const TENANT_ID     = 'f22402ce-b358-43c7-91f9-b90742bf68e4';
const MAILBOX       = 'maintenance@pickwickestates.com';
const FIREBASE_URL  = process.env.FIREBASE_URL;
const FROM_NAME     = 'Dionne — Pickwick Estates';
const FROM_EMAIL    = process.env.OUTLOOK_EMAIL || MAILBOX;

// App-only auth — matches create-subscription.js
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

// Load Firebase data
async function loadFirebase(path) {
  const res = await fetch(FIREBASE_URL + '/' + path + '.json');
  const data = await res.json();
  return data ? Object.values(data) : [];
}

// Read full email — using mailbox path not 'me'
async function getEmail(token, messageId) {
  const res = await fetch(
    `https://graph.microsoft.com/v1.0/users/${MAILBOX}/messages/${messageId}`,
    { headers: { 'Authorization': 'Bearer ' + token } }
  );
  return await res.json();
}

// Create draft in the mailbox — using mailbox path not 'me'
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

// Identify sender type
function identifySender(fromEmail, fromName, properties, contractors) {
  const emailLower = (fromEmail || '').toLowerCase();
  const nameLower  = (fromName || '').toLowerCase();

  const contractor = contractors.find(c =>
    (c.email || '').toLowerCase() === emailLower ||
    nameLower.includes((c.name || '').toLowerCase().split(' ')[0])
  );
  if (contractor) return { type: 'contractor', match: contractor };

  const llProp = properties.find(p =>
    (p.llEmail || '').toLowerCase() === emailLower ||
    (p.llName || '').toLowerCase().includes(nameLower.split(' ')[0])
  );
  if (llProp) return { type: 'landlord', match: llProp };

  const tenantProp = properties.find(p =>
    (p.tenantEmail || '').toLowerCase() === emailLower ||
    (p.tenant || '').toLowerCase().includes(nameLower.split(' ')[0])
  );
  if (tenantProp) return { type: 'tenant', match: tenantProp };

  return { type: 'unknown', match: null };
}

// Find property mentioned in email
function findPropertyInEmail(text, properties) {
  const textLower = text.toLowerCase();
  return properties.find(p => {
    const parts = p.address.toLowerCase().split(',')[0].split(' ');
    return parts.length >= 2 && textLower.includes(parts.slice(0, 2).join(' '));
  }) || null;
}

// Detect intent
function detectIntent(subject, body) {
  const text = (subject + ' ' + body).toLowerCase();
  if (text.includes('boiler') || text.includes('heating') || text.includes('hot water')) return 'boiler';
  if (text.includes('leak') || text.includes('water damage') || text.includes('damp'))    return 'leak';
  if (text.includes('electric') || text.includes('power') || text.includes('fuse'))       return 'electrical';
  if (text.includes('lock') || text.includes('key') || text.includes('door'))             return 'access';
  if (text.includes('quote') || text.includes('estimate') || text.includes('cost'))       return 'quote';
  if (text.includes('invoice') || text.includes('payment') || text.includes('bill'))      return 'invoice';
  if (text.includes('complete') || text.includes('finished') || text.includes('done'))    return 'complete';
  if (text.includes('approve') || text.includes('authorise') || text.includes('go ahead')) return 'approval';
  if (text.includes('deposit') || text.includes('refund'))                                return 'deposit';
  if (text.includes('notice') || text.includes('leaving') || text.includes('vacate'))     return 'notice';
  return 'general';
}

// Generate draft responses
async function generateDrafts(emailData, senderInfo, property, intent) {
  const { fromEmail, fromName, subject, body } = emailData;
  const nl = '\n';
  const drafts = [];

  if (senderInfo.type === 'tenant' && property) {
    const ll = property.llName ? 'Dear ' + property.llName.split(' ')[0] + ',' : 'Dear Landlord,';

    drafts.push({
      to: fromEmail,
      subject: 'Re: ' + subject,
      body: 'Dear ' + fromName.split(' ')[0] + ',' + nl + nl +
        'Thank you for getting in touch regarding ' + property.address + '.' + nl + nl +
        'I have received your message and will look into this as a matter of priority. I will be in touch shortly with an update.' + nl + nl +
        'If the matter is urgent, please do not hesitate to call us directly.' + nl + nl +
        'Kind regards,' + nl + FROM_NAME,
      label: 'Reply to tenant — acknowledgement'
    });

    drafts.push({
      to: property.llEmail || '',
      subject: 'Maintenance issue reported — ' + property.address,
      body: ll + nl + nl +
        'I am writing to advise that your tenant ' + property.tenant + ' at ' + property.address + ' has been in touch regarding the following:' + nl + nl +
        '"' + subject + '"' + nl + nl +
        body.substring(0, 300) + (body.length > 300 ? '...' : '') + nl + nl +
        'Please could you confirm whether you would like us to arrange for a contractor to attend, or whether you wish to handle this directly?' + nl + nl +
        'Kind regards,' + nl + FROM_NAME,
      label: 'Inform landlord of tenant issue'
    });

  } else if (senderInfo.type === 'contractor' && property) {
    const ll = property.llName ? 'Dear ' + property.llName.split(' ')[0] + ',' : 'Dear Landlord,';

    if (intent === 'quote') {
      drafts.push({
        to: property.llEmail || '',
        subject: 'Quote received — ' + property.address,
        body: ll + nl + nl +
          'Please find below a quote received from ' + fromName + ' for works at ' + property.address + ':' + nl + nl +
          body.substring(0, 500) + (body.length > 500 ? '...' : '') + nl + nl +
          'Please could you confirm whether you are happy to proceed with these works?' + nl + nl +
          'Kind regards,' + nl + FROM_NAME,
        label: 'Forward contractor quote to landlord'
      });
      drafts.push({
        to: fromEmail,
        subject: 'Re: ' + subject,
        body: 'Hi ' + fromName.split(' ')[0] + ',' + nl + nl +
          'Thank you for sending over the quote for ' + property.address + '.' + nl + nl +
          'I have forwarded this to the landlord for approval and will be in touch as soon as I hear back.' + nl + nl +
          'Kind regards,' + nl + FROM_NAME,
        label: 'Acknowledge quote receipt to contractor'
      });
    } else if (intent === 'complete') {
      drafts.push({
        to: property.llEmail || '',
        subject: 'Works completed — ' + property.address,
        body: ll + nl + nl +
          'I am writing to advise that ' + fromName + ' has confirmed works are now complete at ' + property.address + '.' + nl + nl +
          'Please let me know if you require any further information.' + nl + nl +
          'Kind regards,' + nl + FROM_NAME,
        label: 'Notify landlord works complete'
      });
    } else if (intent === 'invoice') {
      drafts.push({
        to: property.llEmail || '',
        subject: 'Invoice received — ' + property.address,
        body: ll + nl + nl +
          'Please find below an invoice received from ' + fromName + ' for works at ' + property.address + ':' + nl + nl +
          body.substring(0, 400) + nl + nl +
          'Please could you confirm you are happy for this to be processed?' + nl + nl +
          'Kind regards,' + nl + FROM_NAME,
        label: 'Forward invoice to landlord'
      });
    }

  } else if (senderInfo.type === 'landlord' && property) {
    if (intent === 'approval') {
      drafts.push({
        to: '',
        subject: 'Works approved — please proceed — ' + property.address,
        body: 'Hi,' + nl + nl +
          'I am writing to confirm that the landlord has approved the works at ' + property.address + '.' + nl + nl +
          'Please could you confirm your availability to attend and provide a proposed date?' + nl + nl +
          'Tenant contact: ' + property.tenant + nl +
          'Please arrange access directly with the tenant.' + nl + nl +
          'Kind regards,' + nl + FROM_NAME,
        label: 'Instruct contractor — works approved'
      });
      drafts.push({
        to: property.tenantEmail || '',
        subject: 'Maintenance update — ' + property.address,
        body: 'Dear ' + property.tenant.split(' ')[0] + ',' + nl + nl +
          'I am writing to advise that works have been approved for your property at ' + property.address + '.' + nl + nl +
          'A contractor will be in touch shortly to arrange a convenient time to attend.' + nl + nl +
          'Kind regards,' + nl + FROM_NAME,
        label: 'Notify tenant — works approved'
      });
    }
  }

  // Unknown sender — flag for manual review
  if (senderInfo.type === 'unknown' || !property) {
    drafts.push({
      to: FROM_EMAIL,
      subject: '[ACTION NEEDED] Unmatched email — ' + subject,
      body: 'An email was received that could not be automatically matched to a property or contact.' + nl + nl +
        'From: ' + fromName + ' <' + fromEmail + '>' + nl +
        'Subject: ' + subject + nl + nl +
        'Please review and respond manually.' + nl + nl +
        'Original message:' + nl + body.substring(0, 500),
      label: 'Unmatched email — manual review needed'
    });
  }

  return drafts;
}

// Log job to Firebase
async function logJob(property, subject, fromName, intent) {
  if (!property) return;
  await fetch(FIREBASE_URL + '/jobs.json', {
    method: 'POST',
    body: JSON.stringify({
      id:           Date.now(),
      prop:         property.address,
      tenant:       property.tenant,
      issue:        subject,
      status:       'New — awaiting response',
      contractor:   '',
      contractorId: '',
      access:       'Contact tenant to arrange',
      quote:        0,
      invoice:      0,
      notes:        'Auto-logged from email by ' + fromName,
      date:         new Date().toLocaleDateString('en-GB'),
      intent:       intent,
    }),
  });
}

exports.handler = async function(event) {
  // Validation handshake — Microsoft sends this when subscription is first created
  const params = new URLSearchParams(event.rawQuery || '');
  const validationToken = params.get('validationToken');
  if (validationToken) {
    console.log('Validation token received — confirming subscription');
    return {
      statusCode: 200,
      headers: { 'Content-Type': 'text/plain' },
      body: validationToken,
    };
  }

  if (!event.body) return { statusCode: 200, body: 'No body' };

  try {
    const notification = JSON.parse(event.body);
    const values = notification.value || [];
    console.log('Notifications received:', values.length);

    if (!values.length) return { statusCode: 200, body: 'No notifications' };

    const token = await getAccessToken();
    console.log('Token acquired');

    const [properties, contractors] = await Promise.all([
      loadFirebase('properties'),
      loadFirebase('contractors'),
    ]);
    console.log('Firebase loaded — properties:', properties.length, 'contractors:', contractors.length);

    const results = [];

    for (const notif of values) {
      if (notif.clientState !== 'dg-lettings-secret-2026') {
        console.log('Invalid clientState — skipping');
        continue;
      }

      const messageId = notif.resourceData && notif.resourceData.id;
      if (!messageId) { console.log('No messageId'); continue; }

      const email = await getEmail(token, messageId);
      console.log('Email fetched:', email.subject);

      if (!email || !email.from) { console.log('No from address'); continue; }

      const fromEmail = email.from.emailAddress.address;
      const fromName  = email.from.emailAddress.name || fromEmail;
      const subject   = email.subject || '(No subject)';
      const body      = email.body ? email.body.content.replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim() : '';

      if (fromEmail.toLowerCase() === MAILBOX.toLowerCase()) { console.log('Skipping own email'); continue; }

      const senderInfo = identifySender(fromEmail, fromName, properties, contractors);
      console.log('Sender identified as:', senderInfo.type);

      let property = senderInfo.match && senderInfo.type !== 'contractor' ? senderInfo.match : null;
      if (!property) property = findPropertyInEmail(subject + ' ' + body, properties);
      console.log('Property matched:', property ? property.address : 'none');

      const intent = detectIntent(subject, body);
      console.log('Intent:', intent);

      const drafts = await generateDrafts({ fromEmail, fromName, subject, body }, senderInfo, property, intent);
      console.log('Drafts to create:', drafts.length);

      for (const draft of drafts) {
        const result = await createDraft(token, draft);
        console.log('Draft created:', result.id || result.error);
      }

      if (senderInfo.type === 'tenant' && property && !['quote','invoice','approval','complete','deposit','notice'].includes(intent)) {
        await logJob(property, subject, fromName, intent);
        console.log('Job logged to Firebase');
      }

      results.push({
        from:          fromName,
        type:          senderInfo.type,
        property:      property ? property.address : 'unmatched',
        intent,
        draftsCreated: drafts.length,
      });
    }

    return {
      statusCode: 200,
      body: JSON.stringify({ processed: results.length, results }),
    };

  } catch(e) {
    console.error('Email handler error:', e.message);
    return { statusCode: 200, body: JSON.stringify({ error: e.message }) };
  }
};
