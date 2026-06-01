// ── GET-TOKEN FUNCTION ──
// Exchanges OAuth auth code for refresh token
// Called by auth.html after Microsoft login redirect

exports.handler = async function(event) {
  if(event.httpMethod !== 'POST'){
    return { statusCode: 405, body: 'Method not allowed' };
  }

  const fetch = (...args) => import('node-fetch').then(({default: f}) => f(...args));

  try {
    const { code, redirect_uri } = JSON.parse(event.body);

    const body = new URLSearchParams({
      grant_type:    'authorization_code',
      client_id:     process.env.AZURE_CLIENT_ID,
      client_secret: process.env.AZURE_CLIENT_SECRET,
      code:          code,
      redirect_uri:  redirect_uri,
      scope:         'https://graph.microsoft.com/Mail.ReadWrite offline_access',
    });

    const res = await fetch(
      'https://login.microsoftonline.com/f22402ce-b358-43c7-91f9-b90742bf68e4/oauth2/v2.0/token',
      { method: 'POST', body }
    );

    const data = await res.json();

    if(data.refresh_token){
      // Return only the refresh token — never expose access token in response
      return {
        statusCode: 200,
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ refresh_token: data.refresh_token }),
      };
    } else {
      return {
        statusCode: 400,
        body: JSON.stringify({ error: data }),
      };
    }
  } catch(e) {
    return {
      statusCode: 500,
      body: JSON.stringify({ error: e.message }),
    };
  }
};
