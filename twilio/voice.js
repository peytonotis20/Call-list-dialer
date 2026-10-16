// Twilio Function: /voice  (Visibility: Protected)
//
// Twilio runs this when the dialer starts a call. It dials the number the
// dialer asked for, showing your Twilio number as caller ID. "Protected"
// means only Twilio itself can run it.
//
// answerOnBridge keeps the call "ringing" in the browser until the other
// person (or their voicemail) picks up, which is how the dialer can tell an
// answered call from a no-answer.

let callerId = null;

exports.handler = async function (context, event, callback) {
  const twiml = new Twilio.twiml.VoiceResponse();
  const to = String(event.To || '').trim();

  // US and Canada numbers only, and never 900 premium-rate numbers.
  if (!/^\+1[2-9]\d{2}[2-9]\d{6}$/.test(to) || to.startsWith('+1900')) {
    twiml.say('This number cannot be dialed from the call list dialer.');
    return callback(null, twiml);
  }

  if (!callerId) {
    const first = (await context.getTwilioClient().incomingPhoneNumbers.list({ limit: 1 }))[0];
    callerId = context.CALLER_ID || (first && first.phoneNumber);
  }

  const timeout = Math.min(60, Math.max(10, parseInt(event.Timeout, 10) || 25));
  const dial = twiml.dial({ callerId, answerOnBridge: true, timeout });
  dial.number(to);
  return callback(null, twiml);
};
