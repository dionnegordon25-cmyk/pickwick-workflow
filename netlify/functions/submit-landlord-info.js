// netlify/functions/submit-landlord-info.js
//
// Receives the submission from the public landlord-info.html form.
//
// The DOB is stored under a SEPARATE Firebase node (/pepChecks/{propertyId})
// rather than on the main property record. Properties load in bulk for every
// member of staff every time the CRM opens — keeping DOB out of that bulk
// load means it only ever gets fetched when someone actually opens the
// "Review & confirm" step for that one property, not sitting in everyone's
// browser memory the rest of the time. Staff can also clear it once the
// check is done (see clearPepDobProp in the CRM).
//
// Called as: POST /.netlify/functions/submit-landlord-info

const fetch = (...args) => import('node-fetch').then(({ default: f }) => f(...args));

const { FIREBASE_URL, FIREBASE_SECRET } = process.env;

exports.handler = async (event) => {
  if (event.httpMethod !== 'POST') {
    return { statusCode: 405, body: 'Method not allowed' };
  }

  try {
    const body = JSON.parse(event.body);
    const { propertyId, name, dob } = body;

    if (!propertyId || !name || !dob) {
      return { statusCode: 400, body: JSON.stringify({ error: 'Missing propertyId, name, or dob' }) };
    }

    const today = new Date().toISOString().slice(0, 10);

    // 1. PEP details go to their own node, not the property record.
    await fetch(`${FIREBASE_URL}/pepChecks/${propertyId}.json?auth=${FIREBASE_SECRET}`, {
      method: 'PUT',
      body: JSON.stringify({ name, dob, submittedDate: today })
    });

    // 2. Mark the PEP checklist item as submitted (awaiting staff review) —
    // PATCH so any existing requestedDate is preserved, not overwritten.
    await fetch(`${FIREBASE_URL}/properties/${propertyId}/checklist/pep.json?auth=${FIREBASE_SECRET}`, {
      method: 'PATCH',
      body: JSON.stringify({ status: 'submitted', submittedDate: today })
    });

    // 3. Questionnaire answers are not sensitive — go straight onto the
    // property record as usual, marked complete immediately.
    await fetch(`${FIREBASE_URL}/properties/${propertyId}/checklist/questionnaire.json?auth=${FIREBASE_SECRET}`, {
      method: 'PUT',
      body: JSON.stringify({
        status: 'done',
        date: today,
        avgBills: body.avgBills || '',
        providers: body.providers || '',
        meterLocations: body.meterLocations || '',
        stopcockLocation: body.stopcockLocation || '',
        binDay: body.binDay || '',
        binLocation: body.binLocation || '',
        instructions: body.instructions || ''
      })
    });

    return { statusCode: 200, body: JSON.stringify({ success: true }) };
  } catch (err) {
    console.error('submit-landlord-info error:', err);
    return { statusCode: 500, body: JSON.stringify({ error: err.message }) };
  }
};
