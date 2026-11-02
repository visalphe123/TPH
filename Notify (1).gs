/**
 * Daily follow-up notification (Telegram and/or Email).
 *
 * Set these in Apps Script: Project Settings -> Script Properties
 *   TELEGRAM_TOKEN    (from @BotFather)       optional
 *   TELEGRAM_CHAT_ID  (your chat or group id) optional
 *   NOTIFY_EMAIL      (e.g. you@gmail.com)    optional
 * Configure at least one channel. Both can be used together.
 */

function props_() { return PropertiesService.getScriptProperties(); }

function dueList_() {
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return getAll()
    .filter(c => c['Follow Up'] && c.Stage !== 'Done' && c['Follow Up'] <= today)
    .sort((a, b) => a['Follow Up'].localeCompare(b['Follow Up']))
    .map(c => Object.assign({}, c, { over: c['Follow Up'] < today }));
}

function buildMessage_(list, today) {
  const over = list.filter(c => c.over);
  const now = list.filter(c => !c.over);
  const line = c => '• ' + c.Customer +
    (c.Company ? ' (' + c.Company + ')' : '') +
    (c.Phone ? ' ☎ ' + c.Phone : '') +
    '\n   ' + c.Stage + (c.Project ? ' · ' + c.Project : '') +
    ' · Follow up: ' + c['Follow Up'];
  let t = '📋 Follow up ' + today + '\n';
  if (now.length) t += '\n🔔 ថ្ងៃនេះ (' + now.length + ')\n' + now.map(line).join('\n');
  if (over.length) t += '\n\n⚠️ ហួសកំណត់ (' + over.length + ')\n' + over.map(line).join('\n');
  return t.length > 3900 ? t.slice(0, 3900) + '\n...' : t;
}

function sendTelegram_(text) {
  const token = props_().getProperty('TELEGRAM_TOKEN');
  const chat = props_().getProperty('TELEGRAM_CHAT_ID');
  if (!token || !chat) return false;
  const res = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage', {
    method: 'post',
    payload: { chat_id: chat, text: text },
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) throw new Error('Telegram: ' + res.getContentText());
  return true;
}

function sendEmail_(subject, text) {
  const to = props_().getProperty('NOTIFY_EMAIL');
  if (!to) return false;
  MailApp.sendEmail(to, subject, text);
  return true;
}

/** Runs every morning via trigger. */
function dailyFollowUp(force) {
  const list = dueList_();
  if (!list.length && force !== true) return;
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const text = list.length ? buildMessage_(list, today) : '✅ មិនមានអតិថិជនត្រូវតាមដានថ្ងៃនេះទេ (' + today + ')';
  const sent = [];
  if (sendTelegram_(text)) sent.push('Telegram');
  if (sendEmail_('Follow up ' + today + ' (' + list.length + ')', text)) sent.push('Email');
  if (!sent.length) throw new Error('No channel configured. Set Script Properties first.');
  Logger.log('Sent via: ' + sent.join(', '));
}

/** Run once from the editor to test your settings. */
function testFollowUp() { dailyFollowUp(true); }

/** Run once from the editor to schedule the 8:00 daily reminder. */
function installDailyTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'dailyFollowUp')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('dailyFollowUp').timeBased().everyDays(1).atHour(8).create();
}
