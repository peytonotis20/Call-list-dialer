// Twilio Function: /token  (Visibility: Public)
//
// The dialer page asks this for a short-lived calling token. It only answers
// when the request carries your DIALER_PIN, so nobody else can place calls
// on your account. On first use it also creates the "Call List Dialer"
// TwiML App that points Twilio at the /voice function in this same service.
//
// Environment variables (Settings > Environment Variables):
//   API_KEY        - API key SID (starts with SK)
//   API_SECRET     - that API key's secret
//   DIALER_PIN     - the passcode you type into the dialer
//   CALLER_ID      - optional; your Twilio number in +1XXXXXXXXXX form.
//                    Leave it out to use the first number on your account.
//   ALLOWED_ORIGIN - optional; defaults to https://peytonotis20.github.io
// Also keep "Add my Twilio Credentials (ACCOUNT_SID) and (AUTH_TOKEN) to ENV" ticked.

let appSid = null;
let callerId = null;

exports.handler = async function (context, event, callback) {
  const response = new Twilio.Response();
  response.appendHeader('Access-Control-Allow-Origin', context.ALLOWED_ORIGIN || 'https://peytonotis20.github.io');
  response.appendHeader('Content-Type', 'application/json');
  const fail = (code, error) => {
    response.setStatusCode(code);
    response.setBody({ error });
    callback(null, response);
  };

  if (!context.DIALER_PIN) return fail(500, 'DIALER_PIN is not set in this service\'s environment variables.');
  if (String(event.pin || '') !== context.DIALER_PIN) return fail(401, 'Wrong PIN.');
  if (!context.API_KEY || !context.API_SECRET) return fail(500, 'API_KEY and API_SECRET are not set in this service\'s environment variables.');

  try {
    const client = context.getTwilioClient();
    if (!appSid) {
      const voiceUrl = `https://${context.DOMAIN_NAME}/voice`;
      const apps = await client.applications.list({ friendlyName: 'Call List Dialer', limit: 20 });
      const app = apps.find(a => a.voiceUrl === voiceUrl)
        || await client.applications.create({ friendlyName: 'Call List Dialer', voiceUrl, voiceMethod: 'POST' });
      appSid = app.sid;
    }
    if (!callerId) {
      const first = (await client.incomingPhoneNumbers.list({ limit: 1 }))[0];
      callerId = context.CALLER_ID || (first && first.phoneNumber) || null;
    }
    if (!callerId) return fail(500, 'This Twilio account has no phone number yet. Buy one, then try again.');

    const AccessToken = Twilio.jwt.AccessToken;
    const token = new AccessToken(context.ACCOUNT_SID, context.API_KEY, context.API_SECRET, {
      identity: 'call-list-dialer',
      ttl: 3600,
    });
    token.addGrant(new AccessToken.VoiceGrant({ outgoingApplicationSid: appSid, incomingAllow: false }));
    response.setBody({ token: token.toJwt(), callerId });
    return callback(null, response);
  } catch (err) {
    return fail(500, err.message);
  }
};
