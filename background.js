/*
  Function defs.
*/
async function getBase64Invite(message_id) {
  // Get base64 information from message.
  // Routine from LWR.

  let raw = await browser.messages.getRaw(message_id);
  raw = raw.replace(/\r/g, "");
  let matches = raw.match(/(?=Content-Type: text\/calendar;)(?:[\s\S]*?\n\n)(?<base64>[\s\S]*?)\n\n/);
  if (matches === null) {
    return false;
  }
  return matches['groups'].base64;
}

function base64ToText(base64String) {
  // Convert base64 to plain text.
  // Routine by LWR.

  const binaryString = atob(base64String);
  const codePoints = Array(binaryString.length);
  for (let index = 0; index < binaryString.length; index++)
    codePoints[index] = binaryString.codePointAt(index);
  const uint8CodePoints = Uint8Array.from(codePoints);
  return new TextDecoder().decode(uint8CodePoints);
}

function getICalValue(key, txt) {
  // Returns value to key from iCalender plain text.

  // Find initial index of where to look for value of key and 
  // then slice text accordingly.
  txt = txt.slice(txt.indexOf(key) + key.length);

  // Toggle through keys and compile info into "value" variable.
  if (key === 'DTSTART;' || key === 'DTEND;') {
    // Split text between identifiers ":" and "\n".
    let T = txt.slice(txt.indexOf(':')+1, txt.indexOf('\n')-1);

    // Assemble return value: YYYY-MM-DD, hh:mm (Time zone)
    value = T.slice(0,4) + '-' + T.slice(4,6) + '-' + T.slice(6,8);
    value = value + ', ' + T.slice(9,11) + ':' + T.slice(11,13);
    return value + ' (' + txt.slice(txt.indexOf('=')+1, txt.indexOf(':')) + ')'

  } else if (key === 'X-MICROSOFT-SKYPETEAMSMEETINGURL:') {
    let T = txt.slice(0, txt.indexOf('\nX-MICROSOFT-SCHEDULINGSERVICEUPDATEURL:')-1);
    return T.replace(/\s/g, '')

  } else {
    return ''
  }
}

async function copyStringToClipboard (str) {
  // Copies str to clipboard.
  // Requires "clipboardWrite" permission.

  try {
    await navigator.clipboard.writeText(str);
  } catch (err) {
    console.error('Failed to copy: ', err);
  }
}

function getTimesURL(ics) {
  // Parses iCalender data given as plain text and extracts
  // start time, end time and meeting url info.
  // Not the most robust routine, but it should work until MS changes its 
  // iCalender structure.
  let start_time = 'Start: ' + getICalValue('DTSTART;', ics);
  let end_time = 'End: ' + getICalValue('DTEND;', ics);
  let _url = getICalValue('X-MICROSOFT-SKYPETEAMSMEETINGURL:', ics);
  let url = 'URL: ' + _url;

  // Collect and return all info.
  return start_time + '\n' + end_time + '\n' + url;
}

async function isDarkMode() {
  // Check OS and browser theme and compare so to most 
  // likely detect either dark or bright theme.
  if (window.matchMedia && !!window.matchMedia('(prefers-color-scheme: dark)').matches) {
    os_theme = 'dark';
  } else {
    os_theme = 'bright';
  }
  let theme = await browser.theme.getCurrent();
  if (theme.colors === null) {
    brwsr_theme = 'default';
  } else {
    brwsr_theme = theme.colors;
  }

  if (os_theme === 'dark' && brwsr_theme === 'default') {
    return true;
  } else {
    return false;
  }
}

async function updateIcon() {
  if (await isDarkMode()) {
    browser.messageDisplayAction.setIcon({path: 'images/calendar_white.png'});
  } else {
    browser.messageDisplayAction.setIcon({path: 'images/calendar_black.png'});
  }
}


/*
  Script.
*/

// Add listener on displaying email message.
browser.messageDisplay.onMessageDisplayed.addListener(async (tab, message) => {
  if (await getBase64Invite(message.id) === false) {
    browser.messageDisplayAction.disable(tab.id);
  } else {
    updateIcon();
    browser.messageDisplayAction.enable(tab.id);
  }
});

// Get email message, extract invite info and copy it to clipboard.
browser.messageDisplayAction.onClicked.addListener(async (tab) => {
  const message = await browser.messageDisplay.getDisplayedMessage(tab.id);
  let base64invite = await getBase64Invite(message.id);

  // Convert base64 to iCalender = plain text.
  let txt = base64ToText(base64invite);

  // Extract date/time/url info from text.
  let info = getTimesURL(txt);

  // Copy info to clipboard.
  copyStringToClipboard(info);
});