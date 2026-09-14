// netlify/functions/get-landlord-context.js
//
// Called by the public landlord-info.html page to show which property the
// form is for. Deliberately returns only the address — never rent, deposit,
// tenant details, or anything else from the property record — since this
// endpoint has no login and is reachable by anyone with the link.
//
// Called as: GET /.netlify/functions/get-landlord-context?property=xxx

const fetch = (...args) => import('node-fetch').then(({ default: f }) => f(...args));

const { FIREBASE_URL, FIREBASE_SECRET } = process.env;

exports.handler = async (event) => {
  if (event.httpMethod !== 'GET') {
    return { statusCode: 405, body: 'Method not allowed' };
  }

  const propertyId = event.queryStringParameters && event.queryStringParameters.property;
  if (!propertyId) {
    return { statusCode: 400, body: JSON.stringify({ error: 'Missing property id' }) };
  }

  try {
    const res = await fetch(
      `${FIREBASE_URL}/properties/${propertyId}.json?auth=${FIREBASE_SECRET}`
    );
    const property = await res.json();

    if (!property || !property.address) {
      return { statusCode: 404, body: JSON.stringify({ error: 'Property not found' }) };
    }

    return {
      statusCode: 200,
      body: JSON.stringify({ address: property.address })
    };
  } catch (err) {
    console.error('get-landlord-context error:', err);
    return { statusCode: 500, body: JSON.stringify({ error: err.message }) };
  }
};
