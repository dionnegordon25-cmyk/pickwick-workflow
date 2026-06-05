// =====================================================
// DG LETTINGS — INBOUND EMAIL HANDLER v3
// Intent-driven — works regardless of sender type
// Deduplication via Firebase processed IDs
// =====================================================

const fetch = (...args) => import('node-fetch').then(({default: f}) => f(...args));

const TENANT_ID    = 'f22402ce-b358-43c7-91f9-b90742bf68e4';
const MAILBOX      = 'maintenance@pickwickestates.com';
const FIREBASE_URL = process.env.FIREBASE_URL;
const FROM_NAME    = 'Dionne — Pickwick Estates';
const FROM_EMAIL   = process.env.OUTLOOK_EMAIL || MAILBOX;

// App-only auth
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

// Check if message already processed (Firebase-persisted deduplication)
async function isAlreadyProcessed(messageId) {
  const key = messageId.replace(/[^a-zA-Z0-9]/g, '').substring(0, 60);
  const res = await fetch(FIREBASE_URL + '/processedMessages/' + key + '.json');
  const data = await res.json();
  return !!data;
}

// Mark message as processed in Firebase (expires after 24 hours via TTL value)
async function markProcessed(messageId) {
  const key = messageId.replace(/[^a-zA-Z0-9]/g, '').substring(0, 60);
  await fetch(FIREBASE_URL + '/processedMessages/' + key + '.json', {
    method: 'PUT',
    body: JSON.stringify({ processedAt: Date.now(), expires: Date.now() + 86400000 }),
  });
}

// Read full email
async function getEmail(token, messageId) {
  const res = await fetch(
    `https://graph.microsoft.com/v1.0/users/${MAILBOX}/messages/${messageId}`,
    { headers: { 'Authorization': 'Bearer ' + token } }
  );
  return await res.json();
}

// Create draft in mailbox
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

// ── INTENT-DRIVEN DRAFT GENERATION ──────────────────────────────────────
async function generateDrafts(emailData, senderInfo, property, intent) {
  const { fromEmail, fromName, subject, body } = emailData;
  const nl = '\n';
  const drafts = [];

  // No property matched — flag for manual review
  if (!property) {
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

    // 1. Acknowledge the sender
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
      label: 'Acknowledge sender — ' + senderInfo.type
    });

    // 2. Notify landlord (only if they are NOT the sender AND we have their email)
    if (!isLandlord && property.llEmail && property.llEmail.trim()) {
      drafts.push({
        to: property.llEmail,
        subject: 'Maintenance issue reported — ' + addr,
        body: 'Dear ' + llFirst + ',' + nl + nl +
          'I am writing to advise that a ' + intent + ' issue has been reported at ' + addr + '.' + nl + nl +
          'Reported by: ' + fromName + nl +
          'Details: ' + body.substring(0, 300) + (body.length > 300 ? '...' : '') + nl + nl +
          'We are arranging for a contractor to attend. I will keep you updated on progress and will seek your approval before any significant works are instructed.' + nl + nl +
          'Kind regards,' + nl + FROM_NAME,
        label: 'Notify landlord — maintenance reported'
      });
    } else if (!isLandlord && (!property.llEmail || !property.llEmail.trim())) {
      // No landlord email — create internal flag draft instead
      drafts.push({
        to: FROM_EMAIL,
        subject: '[ACTION NEEDED] No landlord email — ' + addr,
        body: 'A maintenance issue has been reported at ' + addr + ' but no landlord email address is recorded.' + nl + nl +
          'Issue: ' + subject + nl +
          'Reported by: ' + fromName + ' (' + senderInfo.type + ')' + nl + nl +
          'Please add the landlord email to the property record and notify them manually.' + nl + nl +
          'Kind regards,' + nl + FROM_NAME,
        label: 'Internal — missing landlord email'
      });
    }

    // 3. Notify tenant (only if NOT the sender AND we have their email)
    if (!isTenant && property.tenantEmail && property.tenantEmail.trim()) {
      drafts.push({
        to: property.tenantEmail,
        subject: 'Maintenance update — ' + addr,
        body: 'Dear ' + tenFirst + ',' + nl + nl +
          'I am writing to advise that we have been notified of a ' + intent + ' issue at your property and are arranging for a contractor to attend.' + nl + nl +
          'We will be in touch shortly to confirm a date and time. If the matter is urgent please do not hesitate to contact us directly.' + nl + nl +
          'Kind regards,' + nl + FROM_NAME,
        label: 'Notify tenant — contractor being arranged'
      });
    }

    // 4. Internal contractor instruction draft
    drafts.push({
      to: '',
      subject: 'INSTRUCT CONTRACTOR — ' + intent.toUpperCase() + ' — ' + addr,
      body: 'ACTION REQUIRED — please instruct appropriate contractor.' + nl + nl +
        'Property:    ' + addr + nl +
        'Issue type:  ' + intent + nl +
        'Reported by: ' + fromName + ' (' + senderInfo.type + ')' + nl +
        'Tenant:      ' + (property.tenant   || 'Not recorded') + nl +
        'Landlord:    ' + (property.llName   || 'Not recorded') + nl +
        'LL email:    ' + (property.llEmail  || 'NOT ON RECORD — add to property') + nl + nl +
        'Original message:' + nl +
        body.substring(0, 400) + (body.length > 400 ? '...' : '') + nl + nl +
        'Kind regards,' + nl + FROM_NAME,
      label: 'Internal — instruct contractor'
    });
  }

  // ── QUOTE ─────────────────────────────────────────────────────────────
  else if (intent === 'quote') {
    if (property.llEmail && property.llEmail.trim()) {
      drafts.push({
        to: property.llEmail,
        subject: 'Quote received — ' + addr,
        body: 'Dear ' + llFirst + ',' + nl + nl +
          'Please find below a quote received from ' + fromName + ' for works at ' + addr + ':' + nl + nl +
          body.substring(0, 500) + (body.length > 500 ? '...' : '') + nl + nl +
          'Please could you confirm whether you are happy to proceed with these works?' + nl + nl +
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

  // ── INVOICE ───────────────────────────────────────────────────────────
  else if (intent === 'invoice') {
    if (property.llEmail && property.llEmail.trim()) {
      drafts.push({
        to: property.llEmail,
        subject: 'Invoice received — ' + addr,
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

  // ── WORKS COMPLETE ────────────────────────────────────────────────────
  else if (intent === 'complete') {
    if (property.llEmail && property.llEmail.trim()) {
      drafts.push({
        to: property.llEmail,
        subject: 'Works completed — ' + addr,
        body: 'Dear ' + llFirst + ',' + nl + nl +
          'I am writing to advise that ' + fromName + ' has confirmed works are now complete at ' + addr + '.' + nl + nl +
          'Please let me know if you require any further information.' + nl + nl +
          'Kind regards,' + nl + FROM_NAME,
        label: 'Notify landlord — works complete'
      });
    }
    if (property.tenantEmail && property.tenantEmail.trim()) {
      drafts.push({
        to: property.tenantEmail,
        subject: 'Maintenance update — works complete — ' + addr,
        body: 'Dear ' + tenFirst + ',' + nl + nl +
          'I am writing to confirm that the contractor has advised that works are now complete at your property.' + nl + nl +
          'If you have any concerns about the works carried out please do not hesitate to get in touch.' + nl + nl +
          'Kind regards,' + nl + FROM_NAME,
        label: 'Notify tenant — works complete'
      });
    }
  }

  // ── APPROVAL ──────────────────────────────────────────────────────────
  else if (intent === 'approval') {
    drafts.push({
      to: '',
      subject: 'Works approved — proceed — ' + addr,
      body: 'Hi,' + nl + nl +
        'The landlord has approved the works at ' + addr + '. Please could you confirm your availability and proposed date?' + nl + nl +
        'Tenant contact: ' + (property.tenant || 'Not recorded') + nl +
        'Please arrange access directly with the tenant.' + nl + nl +
        'Kind regards,' + nl + FROM_NAME,
      label: 'Instruct contractor — works approved'
    });
    if (property.tenantEmail && property.tenantEmail.trim()) {
      drafts.push({
        to: property.tenantEmail,
        subject: 'Maintenance update — ' + addr,
        body: 'Dear ' + tenFirst + ',' + nl + nl +
          'I am writing to advise that works have been approved for your property. A contractor will be in touch shortly to arrange a convenient time to attend.' + nl + nl +
          'Kind regards,' + nl + FROM_NAME,
        label: 'Notify tenant — works approved'
      });
    }
  }

  // ── DEPOSIT ───────────────────────────────────────────────────────────
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

  // ── NOTICE ────────────────────────────────────────────────────────────
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
    if (!isLandlord && property.llEmail && property.llEmail.trim()) {
      drafts.push({
        to: property.llEmail,
        subject: 'Tenancy update — ' + addr,
        body: 'Dear ' + llFirst + ',' + nl + nl +
          'I am writing to advise that we have received correspondence regarding ' + addr + ' which may require your attention.' + nl + nl +
          'Details: ' + body.substring(0, 300) + nl + nl +
          'Please let me know if you would like to discuss.' + nl + nl +
          'Kind regards,' + nl + FROM_NAME,
        label: 'Notify landlord — tenancy correspondence'
      });
    }
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

// Update automation stats in Firebase — keeps CRM strip in sync
async function updateAutomationStats(newDrafts) {
  try {
    const res = await fetch(FIREBASE_URL + '/automationStats.json');
    const current = await res.json() || {};
    const today = new Date().toLocaleDateString('en-GB');
    const updated = {
      lastRun:          new Date().toISOString(),
      lastRunDate:      today,
      draftsToday:      (current.lastRunDate === today ? (current.draftsToday || 0) : 0) + newDrafts,
      draftsTotal:      (current.draftsTotal || 0) + newDrafts,
      emailsProcessed:  (current.emailsProcessed || 0) + 1,
      status:           'active',
    };
    await fetch(FIREBASE_URL + '/automationStats.json', {
      method: 'PUT',
      body: JSON.stringify(updated),
    });
    console.log('Automation stats updated — drafts today:', updated.draftsToday);
  } catch(e) {
    console.log('Stats update failed (non-critical):', e.message);
  }
}

// ── MAIN HANDLER ─────────────────────────────────────────────────────────
exports.handler = async function(event) {

  // Validation handshake
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

      // Firebase-persisted deduplication — works across separate webhook calls
      const alreadyDone = await isAlreadyProcessed(messageId);
      if (alreadyDone) {
        console.log('Already processed — skipping duplicate');
        continue;
      }
      await markProcessed(messageId);

      const email = await getEmail(token, messageId);
      if (!email || !email.from) { console.log('No from address'); continue; }

      const fromEmail = email.from.emailAddress.address;
      const fromName  = email.from.emailAddress.name || fromEmail;
      const subject   = email.subject || '(No subject)';
      const body      = email.body ? email.body.content.replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim() : '';

      console.log('Email fetched:', subject);

      // Skip emails sent from the mailbox itself
      if (fromEmail.toLowerCase() === MAILBOX.toLowerCase()) {
        console.log('Skipping own email');
        continue;
      }

      const senderInfo = identifySender(fromEmail, fromName, properties, contractors);
      console.log('Sender identified as:', senderInfo.type);

      let property = senderInfo.match && senderInfo.type !== 'contractor' ? senderInfo.match : null;
      if (!property) property = findPropertyInEmail(subject + ' ' + body, properties);
      console.log('Property matched:', property ? property.address : 'none');
      console.log('Landlord email:', property ? (property.llEmail || 'MISSING') : 'n/a');

      const intent = detectIntent(subject, body);
      console.log('Intent:', intent);

      const drafts = await generateDrafts({ fromEmail, fromName, subject, body }, senderInfo, property, intent);
      console.log('Drafts to create:', drafts.length);

      for (const draft of drafts) {
        const result = await createDraft(token, draft);
        console.log('Draft created:', draft.label, '— to:', draft.to || '(blank)', '— id:', result.id || JSON.stringify(result.error));
      }

      // Auto-log maintenance jobs to Firebase
      const maintenanceIntents = ['boiler','leak','electrical','access','general'];
      if (property && maintenanceIntents.includes(intent)) {
        await logJob(property, subject, fromName, intent);
        console.log('Job logged to Firebase');
      }

      // Update automation stats
      if (drafts.length > 0) {
        await updateAutomationStats(drafts.length);
      }

      results.push({
        from:          fromName,
        type:          senderInfo.type,
        property:      property ? property.address : 'unmatched',
        llEmail:       property ? (property.llEmail || 'MISSING') : 'n/a',
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
