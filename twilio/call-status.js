// Twilio Function: /call-status  (Visibility: Public)
//
// After a call ends without being answered, the dialer asks this what
// happened: "busy", "no-answer", "failed" or "canceled" (you hung up while it
// rang). It needs the same DIALER_PIN as /token.

exports.handler = async function (context, event, callback) {
  const response = new Twilio.Response();
  response.appendHeader('Access-Control-Allow-Origin', context.ALLOWED_ORIGIN || 'https://peytonotis20.github.io');
  response.appendHeader('Content-Type', 'application/json');
  const fail = (code, error) => {
    response.setStatusCode(code);
    response.setBody({ error });
    callback(null, response);
  };

  if (!context.DIALER_PIN || String(event.pin || '') !== context.DIALER_PIN) return fail(401, 'Wrong PIN.');
  const sid = String(event.sid || '');
  if (!/^CA[0-9a-f]{32}$/.test(sid)) return fail(400, 'Missing or invalid call ID.');

  try {
    // The browser's call is the parent; the call to the contact is its child.
    const children = await context.getTwilioClient().calls.list({ parentCallSid: sid, limit: 5 });
    const child = children[0];
    response.setBody(child
      ? { status: child.status, duration: Number(child.duration || 0) }
      : { status: 'none' });
    return callback(null, response);
  } catch (err) {
    return fail(500, err.message);
  }
};
