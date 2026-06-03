// =====================================================
// DG LETTINGS — DAILY OUTLOOK DRAFT CREATOR
// Netlify Scheduled Function — runs every morning 8am
// Uses OAuth refresh token — no password stored
// Creates drafts in Outlook for anything outstanding
// Nothing sends without Dionne's approval
// =====================================================

const fetch = (...args) => import('node-fetch').then(({default: f}) => f(...args));

const CLIENT_ID      = process.env.AZURE_CLIENT_ID;
const CLIENT_SECRET  = process.env.AZURE_CLIENT_SECRET;
const REFRESH_TOKEN  = process.env.OUTLOOK_REFRESH_TOKEN;
const OUTLOOK_EMAIL  = process.env.OUTLOOK_EMAIL;
const FIREBASE_URL   = process.env.FIREBASE_URL;
const FROM_NAME      = 'Dionne — Pickwick Estates';

// ── GET ACCESS TOKEN FROM REFRESH TOKEN ──
async function getAccessToken() {
  const body = new URLSearchParams({
    grant_type:    'refresh_token',
    client_id:     CLIENT_ID,
    client_secret: CLIENT_SECRET,
    refresh_token: REFRESH_TOKEN,
    scope:         'https://graph.microsoft.com/Mail.ReadWrite offline_access',
  });
  const res = await fetch(
    'https://login.microsoftonline.com/f22402ce-b358-43c7-91f9-b90742bf68e4/oauth2/v2.0/token',
    { method: 'POST', body }
  );
  const data = await res.json();
  if(!data.access_token) throw new Error('Token error: ' + JSON.stringify(data));
  return data.access_token;
}

// ── CREATE OUTLOOK DRAFT ──
async function createDraft(token, { to, subject, body }) {
  const res = await fetch('https://graph.microsoft.com/v1.0/me/messages', {
    method: 'POST',
    headers: {
      'Authorization': 'Bearer ' + token,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      subject,
      body: { contentType: 'Text', content: body },
      toRecipients: to ? [{ emailAddress: { address: to } }] : [],
      isDraft: true,
    }),
  });
  if(!res.ok) throw new Error('Draft error: ' + await res.text());
  return await res.json();
}

// ── LOAD FROM FIREBASE ──
async function loadFirebase(path) {
  const res = await fetch(FIREBASE_URL + '/' + path + '.json');
  const data = await res.json();
  return data ? Object.values(data) : [];
}

// ── DATE HELPERS ──
function daysUntil(dateStr) {
  if(!dateStr) return null;
  const d = new Date(dateStr);
  const today = new Date(); today.setHours(0,0,0,0);
  return Math.round((d - today) / 86400000);
}

function parseUKDate(str) {
  if(!str) return null;
  const p = str.split('/');
  return p.length === 3 ? new Date(p[2], p[1]-1, p[0]) : null;
}

function nextAnniversaryDays(startDateStr) {
  const start = parseUKDate(startDateStr);
  if(!start) return null;
  const today = new Date(); today.setHours(0,0,0,0);
  let next = new Date(today.getFullYear(), start.getMonth(), start.getDate());
  if(next < today) next = new Date(today.getFullYear()+1, start.getMonth(), start.getDate());
  return Math.round((next - today) / 86400000);
}

function gbp(n) { return '£' + Number(n).toLocaleString('en-GB'); }

// ── EMAIL TEMPLATES ──
function gscEmail(p) {
  const ll = p.llName ? 'Dear ' + p.llName.split(' ')[0] + ',' : 'Dear Landlord,';
  return {
    to: p.llEmail || '',
    subject: 'Gas Safety Certificate renewal required — ' + p.address,
    body: ll + '\n\nI hope this finds you well.\n\nI am writing to advise that the Gas Safety Certificate (GSC) for ' + p.address + ' is due for renewal within the next 30 days (expiry: ' + p.gsc + ').\n\nA valid GSC is a legal requirement under the Gas Safety (Installation and Use) Regulations 1998 and must be renewed annually. Failure to hold a current certificate is a criminal offence and may invalidate your landlord insurance.\n\nPlease let me know how you would like to proceed. We are happy to arrange an engineer on your behalf if required.\n\nKind regards,\n' + FROM_NAME,
  };
}

function eicrEmail(p) {
  const ll = p.llName ? 'Dear ' + p.llName.split(' ')[0] + ',' : 'Dear Landlord,';
  return {
    to: p.llEmail || '',
    subject: 'EICR renewal required — ' + p.address,
    body: ll + '\n\nI am writing to advise that the Electrical Installation Condition Report (EICR) for ' + p.address + ' is due for renewal (expiry: ' + p.eicr + ').\n\nUnder the Electrical Safety Standards in the Private Rented Sector (England) Regulations 2020, a valid EICR is required every five years.\n\nPlease could you confirm how you would like to arrange this?\n\nKind regards,\n' + FROM_NAME,
  };
}

function epcEmail(p) {
  const ll = p.llName ? 'Dear ' + p.llName.split(' ')[0] + ',' : 'Dear Landlord,';
  return {
    to: p.llEmail || '',
    subject: 'EPC renewal required — ' + p.address,
    body: ll + '\n\nI wanted to draw your attention to the Energy Performance Certificate (EPC) for ' + p.address + ', which is due for renewal (expiry: ' + p.epc + ').\n\nA valid EPC is a legal requirement and must meet a minimum rating of E under current MEES regulations.\n\nKind regards,\n' + FROM_NAME,
  };
}

function rentReviewEmail(p, years) {
  const ll = p.llName ? 'Dear ' + p.llName.split(' ')[0] + ',' : 'Dear Landlord,';
  return {
    to: p.llEmail || '',
    subject: 'Tenancy anniversary & rent review — ' + p.address,
    body: ll + '\n\nI hope you are well.\n\nI am writing as the tenancy at ' + p.address + ' is approaching its ' + years + '-year anniversary.\n\nCurrent rent: ' + gbp(p.rent) + ' per month\nTenant: ' + p.tenant + '\nTenancy start: ' + p.startDate + '\n\nThis is a good opportunity to review the rent in line with current market conditions. Any increase would require a Section 13 notice with the required notice period.\n\nPlease let me know how you would like to proceed.\n\nKind regards,\n' + FROM_NAME,
  };
}

function depositEmail(p) {
  const ll = p.llName ? 'Dear ' + p.llName.split(' ')[0] + ',' : 'Dear Landlord,';
  return {
    to: p.llEmail || '',
    subject: 'Deposit — action required — ' + p.address,
    body: ll + '\n\nI am writing regarding the tenancy deposit for ' + p.address + '.\n\nDeposit: ' + gbp(p.deposit) + '\nDeposit ID: ' + p.depositId + '\nStatus: ' + p.depositStatus + '\n\nCould you please confirm whether you are happy to proceed with repayment in full, or whether you wish to raise any deductions?\n\nKind regards,\n' + FROM_NAME,
  };
}

function contractorEmail(job, contractor) {
  return {
    to: contractor.email || '',
    subject: 'Chasing: ' + job.prop + ' — ' + job.issue,
    body: 'Hi ' + contractor.name + ',\n\nI am following up on the outstanding maintenance job below as we have not yet received a response.\n\nProperty: ' + job.prop + '\nIssue: ' + job.issue + '\nTenant: ' + job.tenant + '\nAccess: ' + job.access + '\n\nCould you please confirm your availability and provide a quote at your earliest convenience?\n\nKind regards,\n' + FROM_NAME,
  };
}

function smokeCoEmail(p) {
  const ll = p.llName ? 'Dear ' + p.llName.split(' ')[0] + ',' : 'Dear Landlord,';
  return {
    to: p.llEmail || '',
    subject: 'Smoke & CO alarm check required — ' + p.address,
    body: ll + '\n\nI am writing to advise that the smoke and carbon monoxide alarm check for ' + p.address + ' is due.\n\nUnder the Smoke and Carbon Monoxide Alarm (Amendment) Regulations 2022, alarms must be tested at the start of each tenancy and kept in working order.\n\nPlease confirm this has been carried out or let us know if you need us to arrange a check.\n\nKind regards,\n' + FROM_NAME,
  };
}

// ── MAIN HANDLER ──
exports.handler = async function(event, context) {
  console.log('DG Lettings daily drafts — starting ' + new Date().toISOString());
  const results = [];
  const errors = [];

  try {
    const token = await getAccessToken();
    console.log('Access token obtained');

    const [properties, jobs, contractors] = await Promise.all([
      loadFirebase('properties'),
      loadFirebase('jobs'),
      loadFirebase('contractors'),
    ]);
    console.log('Loaded ' + properties.length + ' properties, ' + jobs.length + ' jobs');

    const today = new Date();
    let draftCount = 0;

    for(const p of properties) {

      // GSC expiring ≤30 days
      const gscDays = daysUntil(p.gsc);
      if(gscDays !== null && gscDays >= 0 && gscDays <= 30) {
        try { await createDraft(token, gscEmail(p)); draftCount++; results.push('GSC: ' + p.address + ' (' + gscDays + 'd)'); }
        catch(e) { errors.push('GSC ' + p.address + ': ' + e.message); }
      }

      // EICR expiring ≤30 days
      const eicrDays = daysUntil(p.eicr);
      if(eicrDays !== null && eicrDays >= 0 && eicrDays <= 30) {
        try { await createDraft(token, eicrEmail(p)); draftCount++; results.push('EICR: ' + p.address + ' (' + eicrDays + 'd)'); }
        catch(e) { errors.push('EICR ' + p.address + ': ' + e.message); }
      }

      // EPC expiring ≤30 days
      const epcDays = daysUntil(p.epc);
      if(epcDays !== null && epcDays >= 0 && epcDays <= 30) {
        try { await createDraft(token, epcEmail(p)); draftCount++; results.push('EPC: ' + p.address + ' (' + epcDays + 'd)'); }
        catch(e) { errors.push('EPC ' + p.address + ': ' + e.message); }
      }

      // APT anniversary ≤30 days
      const annivDays = nextAnniversaryDays(p.startDate);
      if(annivDays !== null && annivDays >= 0 && annivDays <= 30) {
        const start = parseUKDate(p.startDate);
        const years = start ? today.getFullYear() - start.getFullYear() : '?';
        try { await createDraft(token, rentReviewEmail(p, years)); draftCount++; results.push('Anniversary: ' + p.address + ' (' + annivDays + 'd)'); }
        catch(e) { errors.push('Anniversary ' + p.address + ': ' + e.message); }
      }

      // Deposit disputes
      const s = (p.depositStatus || '').toLowerCase();
      if(s.includes('dispute') || s.includes('claim') || s.includes('awaiting landlord')) {
        try { await createDraft(token, depositEmail(p)); draftCount++; results.push('Deposit: ' + p.address); }
        catch(e) { errors.push('Deposit ' + p.address + ': ' + e.message); }
      }

      // Smoke/CO not recorded
      if(!p.smokeAlarm && !p.coDetector && p.llEmail) {
        try { await createDraft(token, smokeCoEmail(p)); draftCount++; results.push('Smoke/CO: ' + p.address); }
        catch(e) { errors.push('Smoke/CO ' + p.address + ': ' + e.message); }
      }
    }

    // Contractor chasers — awaiting contractor, no quote
    for(const job of jobs) {
      if((job.status || '').toLowerCase().includes('awaiting contractor') && !job.quote) {
        const c = contractors.find(x => x.id === job.contractorId);
        if(c && c.email) {
          try { await createDraft(token, contractorEmail(job, c)); draftCount++; results.push('Contractor chase: ' + job.prop + ' → ' + c.name); }
          catch(e) { errors.push('Contractor ' + job.prop + ': ' + e.message); }
        }
      }
    }

    console.log('Complete: ' + draftCount + ' drafts created');
    return {
      statusCode: 200,
      body: JSON.stringify({
        message: draftCount + ' drafts created in Outlook Drafts folder',
        drafts: results,
        errors: errors,
        timestamp: new Date().toISOString(),
      }),
    };

  } catch(err) {
    console.error('Fatal:', err.message);
    return { statusCode: 500, body: JSON.stringify({ error: err.message }) };
  }
};
