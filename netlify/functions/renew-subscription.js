const fetch = (...args) => import('node-fetch').then(({default: f}) => f(...args));
const TENANT_ID = 'f22402ce-b358-43c7-91f9-b90742bf68e4';
const MAILBOX = 'maintenance@pickwickestates.com';
async function getToken() {
  const res = await fetch('https://login.microsoftonline.com/'+TENANT_ID+'/oauth2/v2.0/token',{method:'POST',body:new URLSearchParams({grant_type:'client_credentials',client_id:process.env.AZURE_CLIENT_ID,client_secret:process.env.AZURE_CLIENT_SECRET,scope:'https://graph.microsoft.com/.default'})});
  const data = await res.json();
  if (!data.access_token) throw new Error(JSON.stringify(data));
  return data.access_token;
}
exports.handler = async function() {
  try {
    const token = await getToken();
    const expiry = new Date(Date.now()+3*24*60*60*1000).toISOString();
    const FBURL = process.env.FIREBASE_URL;
    const subSnap = await fetch(FBURL+'/subscription.json');
    const sub = await subSnap.json();
    let result;
    if (sub && sub.id) {
      const r = await fetch('https://graph.microsoft.com/v1.0/subscriptions/'+sub.id,{method:'PATCH',headers:{'Authorization':'Bearer '+token,'Content-Type':'application/json'},body:JSON.stringify({expirationDateTime:expiry})});
      result = await r.json();
    }
    if (!result || !result.id) {
      const r = await fetch('https://graph.microsoft.com/v1.0/subscriptions',{method:'POST',headers:{'Authorization':'Bearer '+token,'Content-Type':'application/json'},body:JSON.stringify({changeType:'created',notificationUrl:process.env.SITE_URL+'/.netlify/functions/email-handler',resource:'users/'+MAILBOX+'/messages',expirationDateTime:expiry,clientState:'dg-lettings-secret-2026'})});
      result = await r.json();
    }
    if (!result.id) throw new Error(JSON.stringify(result));
    await fetch(FBURL+'/subscription.json',{method:'PUT',body:JSON.stringify({id:result.id,expiry:result.expirationDateTime})});
    return {statusCode:200,body:JSON.stringify({ok:true,id:result.id,expiry:result.expirationDateTime})};
  } catch(e) {
    return {statusCode:500,body:JSON.stringify({error:e.message})};
  }
};
