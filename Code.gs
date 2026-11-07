function doGet() {
  return HtmlService.createHtmlOutput(getHtmlContent())
    .setTitle("Planning Semaine")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ----------------------------------------------------
// COORDONNÉES GPS & MÉTÉO (OPEN-METEO)
// ----------------------------------------------------
function getMeteoPrevisions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const paramsSheet = ss.getSheetByName('Paramètres');
  let lat = 48.86;
  let lon = 2.50;

  if (paramsSheet) {
    const data = paramsSheet.getDataRange().getValues();
    for (let r = 0; r < data.length; r++) {
      if (data[r][0] === 'Latitude' && data[r][1]) lat = data[r][1];
      if (data[r][0] === 'Longitude' && data[r][1]) lon = data[r][1];
    }
  }

  const url = 'https://api.open-meteo.com/v1/forecast?latitude=' + lat +
              '&longitude=' + lon +
              '&daily=weathercode,temperature_2m_max,temperature_2m_min' +
              '&hourly=precipitation_probability,weathercode' +
              '&timezone=Europe%2FParis';

  try {
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    return JSON.parse(res.getContentText());
  } catch (e) {
    Logger.log('Erreur météo : ' + e.message);
    return null;
  }
}

function getMeteoWeb(lundiIso) {
  const json = getMeteoPrevisions();
  if (!json || !json.daily || !json.daily.time) return {};

  const result = {};
  const dTimes = json.daily.time;
  const dCodes = json.daily.weathercode;
  const tMax = json.daily.temperature_2m_max;
  const tMin = json.daily.temperature_2m_min;

  function codeToIcon(c) {
    if (c === 0) return '☀️';
    if (c >= 1 && c <= 2) return '🌤️';
    if (c === 3) return '☁️';
    if ((c >= 51 && c <= 67) || (c >= 80 && c <= 82)) return '🌧️';
    if (c >= 71 && c <= 77) return '❄️';
    if (c >= 95) return '⛈';
    return '⛅';
  }

  for (let i = 0; i < dTimes.length; i++) {
    result[dTimes[i]] = {
      icon: codeToIcon(dCodes[i]),
      min: Math.round(tMin[i]),
      max: Math.round(tMax[i])
    };
  }
  return result;
}

function analyserPluieJour(json, dateStrIso) {
  if (!json || !json.hourly || !json.hourly.time) {
    return { hasRain: false, maxProb: 0 };
  }

  const times = json.hourly.time;
  const probs = json.hourly.precipitation_probability;
  const codes = json.hourly.weathercode;

  let maxProb = 0;
  let hasRain = false;

  for (let i = 0; i < times.length; i++) {
    if (times[i].startsWith(dateStrIso)) {
      const hour = parseInt(times[i].substring(11, 13), 10);
      if ((hour >= 7 && hour <= 9) || (hour >= 16 && hour <= 19)) {
        const prob = probs[i] || 0;
        const code = codes[i] || 0;
        if (prob > maxProb) maxProb = prob;
        if (prob >= 40 || (code >= 51 && code <= 67) || (code >= 80 && code <= 82)) {
          hasRain = true;
        }
      }
    }
  }

  return { hasRain: hasRain, maxProb: maxProb };
}

// ----------------------------------------------------
// GESTION DU CALENDRIER GOOGLE & OUTLOOK O365
// ----------------------------------------------------
function getTargetCalendar(familleNameOrId) {
  if (familleNameOrId) {
    const cleanName = String(familleNameOrId).trim();
    let famCal = CalendarApp.getCalendarById(cleanName);
    if (!famCal) {
      const match = CalendarApp.getCalendarsByName(cleanName);
      if (match.length > 0) famCal = match[0];
    }
    if (famCal) return famCal;
  }
  return CalendarApp.getDefaultCalendar();
}

function isPlanningEvent(title) {
  if (!title) return false;
  return title.startsWith("École :") ||
         title.startsWith("Musique Mag") ||
         title.startsWith("Musique (Maman)") ||
         title.startsWith("🌴 Congé") ||
         title.startsWith("📝 ");
}

function cleanWeekEvents(cal, startDate, endDate) {
  const events = cal.getEvents(startDate, endDate);
  events.forEach(ev => {
    if (isPlanningEvent(ev.getTitle()) || ev.getTag('planningEcole') === 'true') {
      ev.deleteEvent();
    }
  });

  const defaultCal = CalendarApp.getDefaultCalendar();
  if (defaultCal.getId() !== cal.getId()) {
    const defEvents = defaultCal.getEvents(startDate, endDate);
    defEvents.forEach(ev => {
      if (isPlanningEvent(ev.getTitle()) || ev.getTag('planningEcole') === 'true') {
        ev.deleteEvent();
      }
    });
  }
}

function sendO365Invite(emailPro, title, start, end, uid, method) {
  if (!emailPro || !String(emailPro).includes('@')) return;
  const icsMethod = method || 'REQUEST';
  const isCancel = (icsMethod === 'CANCEL');
  const organizerEmail = Session.getEffectiveUser().getEmail() || emailPro;

  function toIcsDate(d) {
    return d.toISOString().replace(/[-:]/g, '').split('.')[0] + 'Z';
  }

  const ics = [
    'BEGIN:VCALENDAR',
    'VERSION:2.0',
    'PRODID:-//Google Apps Script//Planning Ecole//FR',
    'METHOD:' + icsMethod,
    'BEGIN:VEVENT',
    'UID:' + uid + '@planning-ecole.fr',
    'SEQUENCE:' + (isCancel ? '1' : '0'),
    'STATUS:' + (isCancel ? 'CANCELLED' : 'CONFIRMED'),
    'DTSTAMP:' + toIcsDate(new Date()),
    'DTSTART:' + toIcsDate(start),
    'DTEND:' + toIcsDate(end),
    'SUMMARY:' + (isCancel ? 'Annulé : ' + title : title),
    'CLASS:PRIVATE',
    'X-MICROSOFT-CDO-BUSYSTATUS:' + (isCancel ? 'FREE' : 'OOF'),
    'X-MICROSOFT-CDO-INTENDEDSTATUS:' + (isCancel ? 'FREE' : 'OOF'),
    'X-MICROSOFT-CDO-RESPONSE-REQUESTED:FALSE',
    'TRANSP:OPAQUE',
    'ORGANIZER;CN=Planning Ecole:mailto:' + organizerEmail,
    'ATTENDEE;ROLE=REQ-PARTICIPANT;PARTSTAT=ACCEPTED;RSVP=FALSE;CN=Ludo:mailto:' + String(emailPro).trim(),
    'END:VEVENT',
    'END:VCALENDAR'
  ].join('\r\n');

  MailApp.sendEmail({
    to: String(emailPro).trim(),
    subject: (isCancel ? "[Annulé] " : "") + title,
    body: isCancel ? "Créneau annulé automatiquement." : "Créneau synchronisé automatiquement.",
    htmlBody: "<p>" + (isCancel ? "Créneau supprimé du planning." : "Créneau personnel synchronisé automatiquement (Privé - Absent du bureau).") + "</p>",
    attachments: [{
      fileName: "invite.ics",
      mimeType: "text/calendar; charset=UTF-8; method=" + icsMethod,
      content: ics
    }]
  });
}

// ----------------------------------------------------
// CHARGEMENT & ENREGISTREMENT PLANNING
// ----------------------------------------------------
function formatSheetDate(val) {
  if (val instanceof Date) {
    const y = val.getFullYear();
    const m = String(val.getMonth() + 1).padStart(2, '0');
    const d = String(val.getDate()).padStart(2, '0');
    return y + '-' + m + '-' + d;
  }
  return String(val).trim();
}

function chargerPlanningSemaine(lundiDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Historique_Semaines');
  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();
  const joursNoms = ['Lundi', 'Mardi', 'Mercredi', 'Jeudi', 'Vendredi'];
  let semaineData = [
    { matin: '', soir: '', tt: '', musique: 'Non', conge: '', note: '' },
    { matin: '', soir: '', tt: '', musique: '', conge: '', note: '' },
    { matin: '', soir: '', tt: '', musique: '', conge: '', note: '' },
    { matin: '', soir: '', tt: '', musique: '', conge: '', note: '' },
    { matin: '', soir: '', tt: '', musique: 'Non', conge: '', note: '' }
  ];

  let trouve = false;
  for (let i = 1; i < data.length; i++) {
    const rowDate = formatSheetDate(data[i][0]);
    if (rowDate === String(lundiDate).trim()) {
      trouve = true;
      const jourNom = data[i][1];
      const idx = joursNoms.indexOf(jourNom);
      if (idx !== -1) {
        let mVal = data[i][2] || '';
        if (mVal === 'Papa') mVal = 'Ludo';
        if (mVal === 'Maman') mVal = 'Mag';

        let sVal = data[i][3] || '';
        if (sVal === 'Papa') sVal = 'Ludo';
        if (sVal === 'Maman') sVal = 'Mag';

        let ttVal = data[i][4] || '';
        if (ttVal === 'Papa') ttVal = 'Ludo';
        if (ttVal === 'Maman') ttVal = 'Mag';

        semaineData[idx] = {
          matin: mVal,
          soir: sVal,
          tt: ttVal,
          musique: data[i][5] || (idx === 0 || idx === 4 ? 'Non' : ''),
          conge: data[i][6] || '',
          note: data[i][7] || ''
        };
      }
    }
  }

  return trouve ? semaineData : null;
}

function enregistrerPlanning(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Historique_Semaines');
  if (!sheet) {
    sheet = ss.insertSheet('Historique_Semaines');
    sheet.appendRow(['Date Lundi', 'Jour', 'Matin', 'Soir', 'TT', 'Musique Mag', 'Congés', 'Notes']);
  }

  // 1. Récupérer l'état précédent pour savoir quels créneaux O365 annuler ou créer
  const ancienPlanning = chargerPlanningSemaine(data.lundiDate);

  // 2. Nettoyer les anciennes lignes de cette semaine dans Google Sheets
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = values.length - 1; i >= 0; i--) {
      if (formatSheetDate(values[i][0]) === String(data.lundiDate).trim()) {
        sheet.deleteRow(i + 2);
      }
    }
  }

  // 3. Écrire les nouvelles lignes
  data.jours.forEach(function(j) {
    sheet.appendRow([
      data.lundiDate,
      j.jour,
      j.matin || '',
      j.soir || '',
      j.tt || '',
      j.musique || '',
      j.conge || '',
      j.note || ''
    ]);
  });

  const paramsSheet = ss.getSheetByName('Paramètres');
  const familleNameOrId = paramsSheet ? paramsSheet.getRange('B6').getValue() : 'Famille';
  const emailProLudo = paramsSheet ? paramsSheet.getRange('B7').getValue() : '';

  const targetCalendar = getTargetCalendar(familleNameOrId);
  const parts = data.lundiDate.split('-').map(Number);
  const baseDate = new Date(parts[0], parts[1] - 1, parts[2]);

  const weekStart = new Date(baseDate); weekStart.setHours(0, 0, 0, 0);
  const weekEnd = new Date(baseDate); weekEnd.setDate(weekEnd.getDate() + 5); weekEnd.setHours(23, 59, 59, 999);
  
  // 4. Purger tous les événements de la semaine sur Google Agenda
  cleanWeekEvents(targetCalendar, weekStart, weekEnd);

  // 5. Recréer les événements Google Agenda et synchroniser les deltas sur O365
  data.jours.forEach(function(j, index) {
    const eventDate = new Date(baseDate);
    eventDate.setDate(baseDate.getDate() + index);
    const dateTag = data.lundiDate + '-j' + index;
    const descNote = j.note ? ('📝 Note : ' + j.note) : '';

    const oldJ = ancienPlanning ? ancienPlanning[index] : { matin: '', soir: '', conge: '' };

    // --- CONGÉS ---
    const isLudoCongeNew = (j.conge === 'Ludo' || j.conge === 'Les 2');
    const isLudoCongeOld = (oldJ.conge === 'Ludo' || oldJ.conge === 'Les 2');
    const startConge = new Date(eventDate); startConge.setHours(8, 0, 0);
    const endConge = new Date(eventDate); endConge.setHours(19, 0, 0);

    if (j.conge) {
      const evConge = targetCalendar.createAllDayEvent('🌴 Congé (' + j.conge + ')', eventDate, {
        description: descNote
      });
      evConge.setTag('planningEcole', 'true');
    }
    if (isLudoCongeNew && !isLudoCongeOld) {
      sendO365Invite(emailProLudo, '🌴 Congé', startConge, endConge, 'ecole-conge-' + dateTag, 'REQUEST');
    } else if (!isLudoCongeNew && isLudoCongeOld) {
      sendO365Invite(emailProLudo, '🌴 Congé', startConge, endConge, 'ecole-conge-' + dateTag, 'CANCEL');
    }

    // --- DÉPOSE MATIN ---
    const startM = new Date(eventDate); startM.setHours(8, 30, 0);
    const endM = new Date(eventDate); endM.setHours(9, 0, 0);

    if (j.matin) {
      const evM = targetCalendar.createEvent('École : Dépose (' + j.matin + ')', startM, endM, {
        description: descNote
      });
      evM.setTag('planningEcole', 'true');
    }
    if (j.matin === 'Ludo' && oldJ.matin !== 'Ludo') {
      sendO365Invite(emailProLudo, 'École : Dépose', startM, endM, 'ecole-matin-' + dateTag, 'REQUEST');
    } else if (j.matin !== 'Ludo' && oldJ.matin === 'Ludo') {
      sendO365Invite(emailProLudo, 'École : Dépose', startM, endM, 'ecole-matin-' + dateTag, 'CANCEL');
    }

    // --- RÉCUPÉRATION SOIR ---
    const startS = new Date(eventDate); startS.setHours(17, 45, 0);
    const endS = new Date(eventDate); endS.setHours(18, 30, 0);

    if (j.soir) {
      const evS = targetCalendar.createEvent('École : Récupération (' + j.soir + ')', startS, endS, {
        description: descNote
      });
      evS.setTag('planningEcole', 'true');
    }
    if (j.soir === 'Ludo' && oldJ.soir !== 'Ludo') {
      sendO365Invite(emailProLudo, 'École : Récupération', startS, endS, 'ecole-soir-' + dateTag, 'REQUEST');
    } else if (j.soir !== 'Ludo' && oldJ.soir === 'Ludo') {
      sendO365Invite(emailProLudo, 'École : Récupération', startS, endS, 'ecole-soir-' + dateTag, 'CANCEL');
    }

    // --- MUSIQUE MAG ---
    if (j.musique === 'Oui') {
      const startMus = new Date(eventDate); startMus.setHours(17, 30, 0);
      const endMus = new Date(eventDate); endMus.setHours(18, 30, 0);
      const evMus = targetCalendar.createEvent('Musique Mag', startMus, endMus, {
        description: descNote
      });
      evMus.setTag('planningEcole', 'true');
    }

    // --- NOTE SEULE (si aucun créneau mais une note est présente) ---
    if (j.note && !j.matin && !j.soir && !j.conge) {
      const evNote = targetCalendar.createAllDayEvent('📝 ' + j.note, eventDate);
      evNote.setTag('planningEcole', 'true');
    }
  });

  return "OK";
}

// ----------------------------------------------------
// RAPPORT DU DIMANCHE & ALERTE VEILLE
// ----------------------------------------------------
function envoyerRapportDimanche() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Historique_Semaines');
  const paramsSheet = ss.getSheetByName('Paramètres');
  let meteoRefSheet = ss.getSheetByName('Meteo_Reference');
  if (!meteoRefSheet) {
    meteoRefSheet = ss.insertSheet('Meteo_Reference');
    meteoRefSheet.appendRow(['Date', 'Jour', 'Prevu_Dimanche', 'MaxProb']);
  }

  const emailLudo = paramsSheet ? paramsSheet.getRange('B2').getValue() : '';
  const emailMag = paramsSheet ? paramsSheet.getRange('B3').getValue() : '';
  const destinataires = [emailLudo, emailMag].filter(Boolean).join(',');
  if (!destinataires) return;

  const today = new Date();
  const dayOfWeek = today.getDay();
  const daysUntilMonday = (dayOfWeek === 0) ? 1 : ((8 - dayOfWeek) % 7 || 7);
  const nextMonday = new Date(today);
  nextMonday.setDate(today.getDate() + daysUntilMonday);

  const lundiStr = formatSheetDate(nextMonday);
  const semaineData = chargerPlanningSemaine(lundiStr);
  if (!semaineData) return;

  const joursNoms = ['Lundi', 'Mardi', 'Mercredi', 'Jeudi', 'Vendredi'];
  const meteoJson = getMeteoPrevisions();
  const baseLundi = new Date(nextMonday);

  meteoRefSheet.clearContents();
  meteoRefSheet.appendRow(['Date', 'Jour', 'Prevu_Dimanche', 'MaxProb']);

  let body = 'Bonjour à tous les deux,\n\nVoici le récapitulatif de la semaine du ' + lundiStr + ' :\n\n';

  semaineData.forEach(function(j, idx) {
    const jourNom = joursNoms[idx];
    const curDate = new Date(baseLundi);
    curDate.setDate(baseLundi.getDate() + idx);
    const curIso = curDate.toISOString().split('T')[0];

    let infoMeteoStr = 'Non disponible';
    let prevuDimanche = 'Sec';
    let maxProb = 0;

    if (meteoJson) {
      const resMeteo = analyserPluieJour(meteoJson, curIso);
      maxProb = resMeteo.maxProb;
      if (resMeteo.hasRain) {
        prevuDimanche = 'Pluie';
        infoMeteoStr = '🌧️ Risque de pluie (' + maxProb + '%) — Prévoir parapluie/ciré';
      } else {
        prevuDimanche = 'Sec';
        infoMeteoStr = '☀️ Sec (proba ' + maxProb + '%)';
      }
    }

    meteoRefSheet.appendRow([curIso, jourNom, prevuDimanche, maxProb]);

    let ligne = '• ' + jourNom + ' (' + curIso + ') :\n' +
                '  - Matin : ' + (j.matin || '—') + '\n' +
                '  - Soir : ' + (j.soir || '—') + '\n' +
                '  - TT : ' + (j.tt || 'Aucun') + '\n' +
                '  - Météo école : ' + infoMeteoStr;

    if (j.conge) ligne += '\n  - 🌴 Congés : ' + j.conge;
    if (jourNom === 'Lundi' || jourNom === 'Vendredi') {
      ligne += '\n  - 🎵 Musique Mag : ' + (j.musique || 'Non');
    }
    if (j.note) ligne += '\n  - 📝 Note : ' + j.note;

    body += ligne + '\n\n';
  });

  body += 'Bonne semaine !';
  MailApp.sendEmail(destinataires, '[Planning] Récapitulatif de la semaine du ' + lundiStr, body);
}

function verifierChangementMeteoVeille() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const meteoRefSheet = ss.getSheetByName('Meteo_Reference');
  const paramsSheet = ss.getSheetByName('Paramètres');
  if (!meteoRefSheet) return;

  const emailLudo = paramsSheet ? paramsSheet.getRange('B2').getValue() : '';
  const emailMag = paramsSheet ? paramsSheet.getRange('B3').getValue() : '';
  const destinataires = [emailLudo, emailMag].filter(Boolean).join(',');
  if (!destinataires) return;

  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  const tomorrowIso = tomorrow.toISOString().split('T')[0];

  const refData = meteoRefSheet.getDataRange().getValues();
  let prevuDimanche = null;
  let jourNom = '';

  for (let i = 1; i < refData.length; i++) {
    let d = formatSheetDate(refData[i][0]);
    if (d === tomorrowIso) {
      jourNom = refData[i][1];
      prevuDimanche = refData[i][2];
      break;
    }
  }

  if (!prevuDimanche) return;

  const meteoJson = getMeteoPrevisions();
  if (!meteoJson) return;

  const resDemain = analyserPluieJour(meteoJson, tomorrowIso);
  const etatActuel = resDemain.hasRain ? 'Pluie' : 'Sec';

  if (prevuDimanche === 'Sec' && etatActuel === 'Pluie') {
    const sujet = '⚠️ [Alerte Météo] Pluie annoncée demain (' + jourNom + ')';
    const message = 'Bonjour,\n\n' +
      'Changement météo par rapport au récapitulatif de dimanche :\n' +
      'De la pluie est désormais prévue demain ' + jourNom + ' (' + tomorrowIso + ') sur les créneaux école.\n' +
      '• Probabilité de pluie : ' + resDemain.maxProb + '%\n' +
      '• Pensez à prévoir ciré, parapluie et bottes pour le trajet.\n\n' +
      'Bonne fin de journée !';

    MailApp.sendEmail(destinataires, sujet, message);
  } else if (prevuDimanche === 'Pluie' && etatActuel === 'Sec') {
    const sujet = '☀️ [Météo] Amélioration pour demain (' + jourNom + ')';
    const message = 'Bonjour,\n\n' +
      'Bonne nouvelle par rapport à dimanche :\n' +
      'La pluie prévue demain ' + jourNom + ' (' + tomorrowIso + ') s\'est dissipée.\n' +
      'Temps sec annoncé sur les créneaux de dépose et de récupération (probabilité : ' + resDemain.maxProb + '%).\n\n' +
      'Bonne fin de journée !';

    MailApp.sendEmail(destinataires, sujet, message);
  }
}

// ----------------------------------------------------
// INTERFACE MOBILE RESPONSIVE AUTOSIZING + CONGÉS & NOTES
// ----------------------------------------------------
function getHtmlContent() {
  return '<!DOCTYPE html>\n' +
  '<html lang="fr">\n' +
  '<head>\n' +
  '  <meta charset="UTF-8">\n' +
  '  <meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no, viewport-fit=cover">\n' +
  '  <meta name="mobile-web-app-capable" content="yes">\n' +
  '  <meta name="apple-mobile-web-app-capable" content="yes">\n' +
  '  <meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">\n' +
  '  <style>\n' +
  '    :root {\n' +
  '      --bg: #0b1120;\n' +
  '      --card: #1e293b;\n' +
  '      --border: #334155;\n' +
  '      --txt: #f8fafc;\n' +
  '      --muted: #64748b;\n' +
  '      --ludo: #38bdf8;\n' +
  '      --mag: #f472b6;\n' +
  '      --both: #a855f7;\n' +
  '      --conge: #f59e0b;\n' +
  '      --accent: #4f46e5;\n' +
  '      --success: #10b981;\n' +
  '      --error: #ef4444;\n' +
  '    }\n' +
  '    * { box-sizing: border-box; -webkit-tap-highlight-color: transparent; margin: 0; padding: 0; }\n' +
  '    html, body {\n' +
  '      height: 100%;\n' +
  '      height: 100dvh;\n' +
  '      overflow: hidden;\n' +
  '    }\n' +
  '    body {\n' +
  '      background: var(--bg);\n' +
  '      color: var(--txt);\n' +
  '      font-family: -apple-system, BlinkMacSystemFont, "SF Pro Text", "Segoe UI", Roboto, sans-serif;\n' +
  '      padding: max(6px, env(safe-area-inset-top)) max(8px, env(safe-area-inset-right)) max(10px, env(safe-area-inset-bottom)) max(8px, env(safe-area-inset-left));\n' +
  '      display: flex;\n' +
  '      justify-content: center;\n' +
  '      align-items: stretch;\n' +
  '    }\n' +
  '    .app {\n' +
  '      width: 100%;\n' +
  '      max-width: clamp(340px, 96vw, 560px);\n' +
  '      height: 100%;\n' +
  '      display: flex;\n' +
  '      flex-direction: column;\n' +
  '      justify-content: space-between;\n' +
  '      gap: 5px;\n' +
  '    }\n' +
  '    .top-bar {\n' +
  '      display: flex;\n' +
  '      align-items: center;\n' +
  '      justify-content: space-between;\n' +
  '      background: var(--card);\n' +
  '      border: 1px solid var(--border);\n' +
  '      border-radius: 12px;\n' +
  '      padding: 5px 12px;\n' +
  '      flex-shrink: 0;\n' +
  '    }\n' +
  '    .nav-btn {\n' +
  '      background: #334155;\n' +
  '      border: none;\n' +
  '      color: var(--txt);\n' +
  '      width: 32px;\n' +
  '      height: 32px;\n' +
  '      border-radius: 8px;\n' +
  '      font-weight: bold;\n' +
  '      font-size: 18px;\n' +
  '      cursor: pointer;\n' +
  '      display: flex;\n' +
  '      align-items: center;\n' +
  '      justify-content: center;\n' +
  '    }\n' +
  '    .week-title {\n' +
  '      font-size: clamp(13px, 3.8vw, 15px);\n' +
  '      font-weight: 700;\n' +
  '    }\n' +
  '    .col-headers {\n' +
  '      display: grid;\n' +
  '      grid-template-columns: minmax(56px, 1.15fr) 1fr 1fr 1.38fr;\n' +
  '      padding: 0 6px;\n' +
  '      font-size: clamp(9.5px, 2.5vw, 11px);\n' +
  '      font-weight: 700;\n' +
  '      color: var(--muted);\n' +
  '      text-transform: uppercase;\n' +
  '      text-align: center;\n' +
  '      gap: 4px;\n' +
  '      flex-shrink: 0;\n' +
  '    }\n' +
  '    .col-headers span:first-child { text-align: left; }\n' +
  '    #daysList {\n' +
  '      flex: 1;\n' +
  '      display: flex;\n' +
  '      flex-direction: column;\n' +
  '      justify-content: space-between;\n' +
  '      gap: 5px;\n' +
  '      min-height: 0;\n' +
  '    }\n' +
  '    .day-row {\n' +
  '      flex: 1;\n' +
  '      background: var(--card);\n' +
  '      border: 1px solid var(--border);\n' +
  '      border-radius: 10px;\n' +
  '      padding: 4px 8px;\n' +
  '      display: flex;\n' +
  '      flex-direction: column;\n' +
  '      justify-content: center;\n' +
  '      gap: 3px;\n' +
  '      min-height: 0;\n' +
  '    }\n' +
  '    .day-main-grid {\n' +
  '      display: grid;\n' +
  '      grid-template-columns: minmax(56px, 1.15fr) 1fr 1fr 1.38fr;\n' +
  '      align-items: center;\n' +
  '      gap: 4px;\n' +
  '      width: 100%;\n' +
  '    }\n' +
  '    .day-lbl { display: flex; flex-direction: column; min-width: 0; overflow: hidden; }\n' +
  '    .d-name-line { display: flex; align-items: baseline; gap: 3px; white-space: nowrap; }\n' +
  '    .d-name { font-size: clamp(12px, 3.4vw, 14px); font-weight: 700; line-height: 1.1; }\n' +
  '    .d-date { font-size: clamp(9px, 2.5vw, 11px); color: var(--muted); }\n' +
  '    .d-meteo { font-size: clamp(8.5px, 2.3vw, 10px); color: #94a3b8; margin-top: 1px; white-space: nowrap; }\n' +
  '    .grp {\n' +
  '      display: flex;\n' +
  '      background: #0f172a;\n' +
  '      padding: 2px;\n' +
  '      border-radius: 7px;\n' +
  '      justify-content: center;\n' +
  '      gap: 2px;\n' +
  '      width: 100%;\n' +
  '    }\n' +
  '    .btn {\n' +
  '      border: none;\n' +
  '      background: transparent;\n' +
  '      color: var(--muted);\n' +
  '      font-size: clamp(11px, 3vw, 13px);\n' +
  '      font-weight: 800;\n' +
  '      padding: clamp(5px, 1vh, 8px) 0;\n' +
  '      flex: 1;\n' +
  '      text-align: center;\n' +
  '      border-radius: 5px;\n' +
  '      cursor: pointer;\n' +
  '    }\n' +
  '    .btn.act-l { background: var(--ludo); color: #0b1120; }\n' +
  '    .btn.act-m { background: var(--mag); color: #0b1120; }\n' +
  '    .btn.act-2 { background: var(--both); color: #fff; }\n' +
  '    .extra-row {\n' +
  '      padding-top: 3px;\n' +
  '      border-top: 1px dashed #263349;\n' +
  '      display: flex;\n' +
  '      align-items: center;\n' +
  '      justify-content: space-between;\n' +
  '      gap: 4px;\n' +
  '    }\n' +
  '    .chip-btn {\n' +
  '      border: 1px solid var(--border);\n' +
  '      background: #0f172a;\n' +
  '      color: var(--muted);\n' +
  '      font-size: clamp(9.5px, 2.5vw, 11px);\n' +
  '      font-weight: 700;\n' +
  '      padding: 2px 7px;\n' +
  '      border-radius: 6px;\n' +
  '      cursor: pointer;\n' +
  '      white-space: nowrap;\n' +
  '      overflow: hidden;\n' +
  '      text-overflow: ellipsis;\n' +
  '    }\n' +
  '    .chip-btn.conge-l { background: var(--ludo); color: #0b1120; border-color: var(--ludo); }\n' +
  '    .chip-btn.conge-m { background: var(--mag); color: #0b1120; border-color: var(--mag); }\n' +
  '    .chip-btn.conge-2 { background: var(--conge); color: #0b1120; border-color: var(--conge); }\n' +
  '    .chip-btn.mus-active { background: var(--mag); color: #0b1120; border-color: var(--mag); }\n' +
  '    .chip-btn.note-active { background: #334155; color: #38bdf8; border-color: #38bdf8; max-width: 140px; }\n' +
  '    .submit-btn {\n' +
  '      background: var(--accent);\n' +
  '      color: white;\n' +
  '      border: none;\n' +
  '      border-radius: 12px;\n' +
  '      padding: clamp(10px, 1.6vh, 14px);\n' +
  '      font-size: clamp(14px, 3.8vw, 16px);\n' +
  '      font-weight: 700;\n' +
  '      cursor: pointer;\n' +
  '      flex-shrink: 0;\n' +
  '    }\n' +
  '    .submit-btn:disabled { background: #334155; cursor: not-allowed; }\n' +
  '    .submit-btn.success { background: var(--success) !important; color: #fff; }\n' +
  '    .submit-btn.error { background: var(--error) !important; color: #fff; }\n' +
  '    .modal-bg {\n' +
  '      position: fixed; inset: 0; background: rgba(0,0,0,0.75);\n' +
  '      display: none; align-items: center; justify-content: center; padding: 16px; z-index: 99;\n' +
  '    }\n' +
  '    .modal-box {\n' +
  '      background: var(--card); border: 1px solid var(--border); border-radius: 14px;\n' +
  '      padding: 16px; width: 100%; max-width: 340px; display: flex; flex-direction: column; gap: 10px;\n' +
  '    }\n' +
  '    .modal-input {\n' +
  '      width: 100%; background: #0f172a; border: 1px solid var(--border); color: #fff;\n' +
  '      border-radius: 8px; padding: 10px; font-size: 14px;\n' +
  '    }\n' +
  '  </style>\n' +
  '</head>\n' +
  '<body>\n' +
  '  <div class="app">\n' +
  '    <div class="top-bar">\n' +
  '      <button class="nav-btn" onclick="changeWeek(-7)">‹</button>\n' +
  '      <span class="week-title" id="weekLabel"></span>\n' +
  '      <button class="nav-btn" onclick="changeWeek(7)">›</button>\n' +
  '    </div>\n' +
  '    <div class="col-headers">\n' +
  '      <span>Jour</span>\n' +
  '      <span>Matin</span>\n' +
  '      <span>Soir</span>\n' +
  '      <span>TT</span>\n' +
  '    </div>\n' +
  '    <div id="daysList"></div>\n' +
  '    <button class="submit-btn" id="submitBtn" onclick="save()">Synchroniser l\'agenda</button>\n' +
  '  </div>\n' +
  '\n' +
  '  <div class="modal-bg" id="noteModal">\n' +
  '    <div class="modal-box">\n' +
  '      <div style="font-weight:700; font-size:14px;" id="noteModalTitle">Note</div>\n' +
  '      <input type="text" class="modal-input" id="noteInput" placeholder="Ex: Piscine, Grève, Rdv pédiatre...">\n' +
  '      <div style="display:flex; gap:8px; justify-content:flex-end;">\n' +
  '        <button class="chip-btn" onclick="clearNote()" style="padding:6px 10px;">Effacer</button>\n' +
  '        <button class="chip-btn conge-l" onclick="closeNoteModal(true)" style="padding:6px 14px;">Valider</button>\n' +
  '      </div>\n' +
  '    </div>\n' +
  '  </div>\n' +
  '\n' +
  '  <script>\n' +
  '    const jours = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi"];\n' +
  '    let currentMonday = getMonday(new Date());\n' +
  '    let state = getDefaultState();\n' +
  '    let meteoData = {};\n' +
  '    let editingNoteIdx = -1;\n' +
  '\n' +
  '    function getDefaultState() {\n' +
  '      return [\n' +
  '        { matin: "", soir: "", tt: "", musique: "Non", conge: "", note: "" },\n' +
  '        { matin: "", soir: "", tt: "", musique: "", conge: "", note: "" },\n' +
  '        { matin: "", soir: "", tt: "", musique: "", conge: "", note: "" },\n' +
  '        { matin: "", soir: "", tt: "", musique: "", conge: "", note: "" },\n' +
  '        { matin: "", soir: "", tt: "", musique: "Non", conge: "", note: "" }\n' +
  '      ];\n' +
  '    }\n' +
  '\n' +
  '    function getMonday(d) {\n' +
  '      const date = new Date(d);\n' +
  '      const day = date.getDay();\n' +
  '      const diff = date.getDate() - day + (day === 0 ? -6 : 1);\n' +
  '      date.setDate(diff);\n' +
  '      date.setHours(0,0,0,0);\n' +
  '      return date;\n' +
  '    }\n' +
  '\n' +
  '    function getIsoForDay(baseDate, offsetDays) {\n' +
  '      const d = new Date(baseDate);\n' +
  '      d.setDate(baseDate.getDate() + offsetDays);\n' +
  '      const y = d.getFullYear();\n' +
  '      const m = String(d.getMonth() + 1).padStart(2, "0");\n' +
  '      const day = String(d.getDate()).padStart(2, "0");\n' +
  '      return y + "-" + m + "-" + day;\n' +
  '    }\n' +
  '\n' +
  '    function updateWeekLabel() {\n' +
  '      const end = new Date(currentMonday);\n' +
  '      end.setDate(currentMonday.getDate() + 4);\n' +
  '      const opt = { day: "numeric", month: "short" };\n' +
  '      document.getElementById("weekLabel").textContent = \n' +
  '        currentMonday.toLocaleDateString("fr-FR", opt) + " — " + end.toLocaleDateString("fr-FR", opt);\n' +
  '    }\n' +
  '\n' +
  '    function render() {\n' +
  '      const c = document.getElementById("daysList");\n' +
  '      c.innerHTML = "";\n' +
  '      jours.forEach((nom, i) => {\n' +
  '        const d = new Date(currentMonday);\n' +
  '        d.setDate(currentMonday.getDate() + i);\n' +
  '        const dateStr = d.toLocaleDateString("fr-FR", { day: "numeric", month: "numeric" });\n' +
  '        const iso = getIsoForDay(currentMonday, i);\n' +
  '        const hasMusique = (nom === "Lundi" || nom === "Vendredi");\n' +
  '\n' +
  '        const mInfo = meteoData[iso];\n' +
  '        const meteoStr = mInfo ? (mInfo.icon + " " + mInfo.min + "° | " + mInfo.max + "°") : "";\n' +
  '\n' +
  '        const cg = state[i].conge || "";\n' +
  '        const cgClass = cg === "Ludo" ? "conge-l" : (cg === "Mag" ? "conge-m" : (cg === "Les 2" ? "conge-2" : ""));\n' +
  '        const cgLabel = cg ? ("🌴 " + cg) : "🌴 Congé";\n' +
  '\n' +
  '        const nt = state[i].note || "";\n' +
  '        const ntLabel = nt ? ("📝 " + nt) : "📝 Note";\n' +
  '\n' +
  '        const row = document.createElement("div");\n' +
  '        row.className = "day-row";\n' +
  '        let html = \'\' +\n' +
  '          \'<div class="day-main-grid">\' +\n' +
  '            \'<div class="day-lbl">\' +\n' +
  '              \'<div class="d-name-line">\' +\n' +
  '                \'<span class="d-name">\' + nom.substring(0,3) + \'.</span>\' +\n' +
  '                \'<span class="d-date">\' + dateStr + \'</span>\' +\n' +
  '              \'</div>\' +\n' +
  '              (meteoStr ? \'<span class="d-meteo">\' + meteoStr + \'</span>\' : \'\') +\n' +
  '            \'</div>\' +\n' +
  '            \'<div class="grp">\' +\n' +
  '              \'<button type="button" class="btn \' + (state[i].matin === "Ludo" ? "act-l" : "") + \'" onclick="setVal(\' + i + \', \\\'matin\\\', \\\'Ludo\\\')">L</button>\' +\n' +
  '              \'<button type="button" class="btn \' + (state[i].matin === "Mag" ? "act-m" : "") + \'" onclick="setVal(\' + i + \', \\\'matin\\\', \\\'Mag\\\')">M</button>\' +\n' +
  '            \'</div>\' +\n' +
  '            \'<div class="grp">\' +\n' +
  '              \'<button type="button" class="btn \' + (state[i].soir === "Ludo" ? "act-l" : "") + \'" onclick="setVal(\' + i + \', \\\'soir\\\', \\\'Ludo\\\')">L</button>\' +\n' +
  '              \'<button type="button" class="btn \' + (state[i].soir === "Mag" ? "act-m" : "") + \'" onclick="setVal(\' + i + \', \\\'soir\\\', \\\'Mag\\\')">M</button>\' +\n' +
  '            \'</div>\' +\n' +
  '            \'<div class="grp">\' +\n' +
  '              \'<button type="button" class="btn \' + (state[i].tt === "Ludo" ? "act-l" : "") + \'" onclick="setVal(\' + i + \', \\\'tt\\\', \\\'Ludo\\\')">L</button>\' +\n' +
  '              \'<button type="button" class="btn \' + (state[i].tt === "Mag" ? "act-m" : "") + \'" onclick="setVal(\' + i + \', \\\'tt\\\', \\\'Mag\\\')">M</button>\' +\n' +
  '              \'<button type="button" class="btn \' + (state[i].tt === "Les 2" ? "act-2" : "") + \'" onclick="setVal(\' + i + \', \\\'tt\\\', \\\'Les 2\\\')">2</button>\' +\n' +
  '            \'</div>\' +\n' +
  '          \'</div>\' +\n' +
  '          \'<div class="extra-row">\' +\n' +
  '            \'<button type="button" class="chip-btn \' + cgClass + \'" onclick="cycleConge(\' + i + \')">\' + cgLabel + \'</button>\' +\n' +
  '            (hasMusique ? \'<button type="button" class="chip-btn \' + (state[i].musique === "Oui" ? "mus-active" : "") + \'" onclick="toggleMusique(\' + i + \')">🎵 Musique : \' + (state[i].musique === "Oui" ? "Oui" : "Non") + \'</button>\' : \'\') +\n' +
  '            \'<button type="button" class="chip-btn \' + (nt ? "note-active" : "") + \'" onclick="openNoteModal(\' + i + \')">\' + ntLabel + \'</button>\' +\n' +
  '          \'</div>\';\n' +
  '        row.innerHTML = html;\n' +
  '        c.appendChild(row);\n' +
  '      });\n' +
  '    }\n' +
  '\n' +
  '    function setVal(idx, k, val) {\n' +
  '      state[idx][k] = (state[idx][k] === val) ? "" : val;\n' +
  '      render();\n' +
  '    }\n' +
  '\n' +
  '    function cycleConge(idx) {\n' +
  '      const order = ["", "Ludo", "Mag", "Les 2"];\n' +
  '      const cur = state[idx].conge || "";\n' +
  '      const next = order[(order.indexOf(cur) + 1) % order.length];\n' +
  '      state[idx].conge = next;\n' +
  '      render();\n' +
  '    }\n' +
  '\n' +
  '    function toggleMusique(idx) {\n' +
  '      state[idx].musique = (state[idx].musique === "Oui") ? "Non" : "Oui";\n' +
  '      render();\n' +
  '    }\n' +
  '\n' +
  '    function openNoteModal(idx) {\n' +
  '      editingNoteIdx = idx;\n' +
  '      document.getElementById("noteModalTitle").textContent = "Note — " + jours[idx];\n' +
  '      document.getElementById("noteInput").value = state[idx].note || "";\n' +
  '      document.getElementById("noteModal").style.display = "flex";\n' +
  '      setTimeout(() => document.getElementById("noteInput").focus(), 50);\n' +
  '    }\n' +
  '\n' +
  '    function clearNote() {\n' +
  '      document.getElementById("noteInput").value = "";\n' +
  '      closeNoteModal(true);\n' +
  '    }\n' +
  '\n' +
  '    function closeNoteModal(saveNote) {\n' +
  '      if (saveNote && editingNoteIdx !== -1) {\n' +
  '        state[editingNoteIdx].note = document.getElementById("noteInput").value.trim();\n' +
  '      }\n' +
  '      document.getElementById("noteModal").style.display = "none";\n' +
  '      editingNoteIdx = -1;\n' +
  '      render();\n' +
  '    }\n' +
  '\n' +
  '    function changeWeek(offset) {\n' +
  '      currentMonday.setDate(currentMonday.getDate() + offset);\n' +
  '      updateWeekLabel();\n' +
  '      loadCurrentWeekData();\n' +
  '    }\n' +
  '\n' +
  '    function loadCurrentWeekData() {\n' +
  '      const iso = getIsoForDay(currentMonday, 0);\n' +
  '      google.script.run\n' +
  '        .withSuccessHandler(function(res) {\n' +
  '          state = res ? res : getDefaultState();\n' +
  '          render();\n' +
  '        })\n' +
  '        .chargerPlanningSemaine(iso);\n' +
  '\n' +
  '      google.script.run\n' +
  '        .withSuccessHandler(function(mRes) {\n' +
  '          if (mRes) {\n' +
  '            meteoData = mRes;\n' +
  '            render();\n' +
  '          }\n' +
  '        })\n' +
  '        .getMeteoWeb(iso);\n' +
  '    }\n' +
  '\n' +
  '    function save() {\n' +
  '      const btn = document.getElementById("submitBtn");\n' +
  '      btn.disabled = true;\n' +
  '      btn.className = "submit-btn";\n' +
  '      btn.textContent = "Synchronisation en cours...";\n' +
  '\n' +
  '      const payload = {\n' +
  '        lundiDate: getIsoForDay(currentMonday, 0),\n' +
  '        jours: state.map((s, idx) => ({ jour: jours[idx], ...s }))\n' +
  '      };\n' +
  '      google.script.run\n' +
  '        .withSuccessHandler(function() {\n' +
  '          btn.className = "submit-btn success";\n' +
  '          btn.textContent = "✓ Synchronisé !";\n' +
  '          setTimeout(function() {\n' +
  '            btn.className = "submit-btn";\n' +
  '            btn.textContent = "Synchroniser l\'agenda";\n' +
  '            btn.disabled = false;\n' +
  '          }, 3000);\n' +
  '        })\n' +
  '        .withFailureHandler(function(err) {\n' +
  '          btn.className = "submit-btn error";\n' +
  '          btn.textContent = "Erreur : " + err.message;\n' +
  '          setTimeout(function() {\n' +
  '            btn.className = "submit-btn";\n' +
  '            btn.textContent = "Synchroniser l\'agenda";\n' +
  '            btn.disabled = false;\n' +
  '          }, 3500);\n' +
  '        })\n' +
  '        .enregistrerPlanning(payload);\n' +
  '    }\n' +
  '\n' +
  '    updateWeekLabel();\n' +
  '    loadCurrentWeekData();\n' +
  '  </script>\n' +
  '</body>\n' +
  '</html>';
}
