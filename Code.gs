const CONFIG = Object.freeze({
  prospectsSheetName: 'Prospects',
  status: {
    new: 'Nouveau',
    contacted: 'Contacté',
  },
  emailSubject: 'ImageUp - Retouche photographique immobilière en 24h',
  senderName: 'L’équipe ImageUp (by L4win Store)',
  beforeImageFileId: 'PASTE_BEFORE_IMAGE_FILE_ID',
  afterImageFileId: 'PASTE_AFTER_IMAGE_FILE_ID',
});

const STATUS_COLUMN_INDEX = 3; // Column "Statut" (1-indexed).

/**
 * Parcourt l’onglet Prospects et envoie un e-mail à chaque prospect "Nouveau" disposant d’une adresse mail.
 * Une fois l’e-mail expédié, le statut est mis à jour en "Contacté" pour éviter les doublons.
 */
function sendNewProspectEmails() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(5000)) {
    Logger.log('Un autre envoi est déjà en cours. Nouvelle exécution annulée.');
    return;
  }

  try {
    const sheet = getProspectSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return; // Aucun prospect à traiter (uniquement l’en-tête).
    }

    const dataRange = sheet.getRange(2, 1, lastRow - 1, STATUS_COLUMN_INDEX);
    const values = dataRange.getValues();
    const emailQueue = [];

    const normalizedNewStatus = CONFIG.status.new.toLowerCase();

    values.forEach((row, index) => {
      const rowNumber = index + 2; // Décalage pour tenir compte de l’en-tête.
      const name = normalizeValue_(row[0]);
      const email = normalizeValue_(row[1]);
      const status = normalizeValue_(row[2]).toLowerCase();

      if (status === normalizedNewStatus && name && email) {
        emailQueue.push({ rowNumber, name, email });
      }
    });

    if (emailQueue.length === 0) {
      return; // Aucun e-mail à envoyer.
    }

    const inlineImages = getInlineImages_();

    emailQueue.forEach(({ rowNumber, name, email }) => {
      try {
        const bodies = buildEmailBodies_(name);
        GmailApp.sendEmail(email, CONFIG.emailSubject, bodies.textBody, {
          name: CONFIG.senderName,
          htmlBody: bodies.htmlBody,
          inlineImages: {
            beforeImage: inlineImages.before.copyBlob(),
            afterImage: inlineImages.after.copyBlob(),
          },
        });

        sheet.getRange(rowNumber, STATUS_COLUMN_INDEX).setValue(CONFIG.status.contacted);
      } catch (error) {
        Logger.log(`Impossible d’envoyer l’e-mail pour la ligne ${rowNumber} (${name}) : ${error}`);
      }
    });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Crée un déclencheur installable "on edit" pour lancer automatiquement l’envoi dès qu’une ligne est modifiée.
 * À exécuter une seule fois depuis l’éditeur Apps Script.
 */
function createProspectEditTrigger() {
  const handlerFunction = 'sendNewProspectEmails';
  const triggers = ScriptApp.getProjectTriggers();
  const alreadyExists = triggers.some(trigger => trigger.getHandlerFunction() === handlerFunction);

  if (alreadyExists) {
    Logger.log('Le déclencheur est déjà configuré.');
    return;
  }

  ScriptApp.newTrigger(handlerFunction)
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();

  Logger.log('Déclencheur "on edit" créé : la fonction sendNewProspectEmails sera appelée automatiquement.');
}

/**
 * Prépare les versions texte et HTML de l’e-mail.
 * @param {string} prospectName Nom du prospect (colonne Nom).
 * @returns {{textBody: string, htmlBody: string}}
 */
function buildEmailBodies_(prospectName) {
  const safeNameHtml = escapeHtml_(prospectName);

  const htmlBody = [
    '<p>Bonjour,</p>',
    `<p>J’ai découvert l’engagement de <strong>${safeNameHtml}</strong> pour des projets immobiliers durables, et je salue la qualité de vos réalisations. Pour refléter cet engagement jusque dans vos supports de vente, il est crucial de disposer de visuels percutants et authentiques. À ce titre, permettez-moi de vous présenter ImageUp, notre service de retouche photographique immobilière ultra-rapide et fiable, qui pourrait grandement servir votre communication.</p>`,
    '<p>Ce que nous proposons :</p>',
    '<ul>',
    '<li><strong>Service express</strong> : vos photos retouchées en 24h chrono, pour soutenir la rapidité de vos mises en vente et promotions.</li>',
    '<li><strong>Prix concurrentiels</strong> : une offre économique adaptée aux promoteurs indépendants, vous assurant un excellent rapport qualité-prix.</li>',
    '<li><strong>Retouche fidèle au bien</strong> : chaque image est optimisée (lumière, couleurs, netteté) tout en conservant l’aspect réel du bien, afin de rester transparent vis-à-vis des acquéreurs.</li>',
    '</ul>',
    '<div style="margin:20px 0;text-align:center;">',
    '<div style="display:inline-block;margin:0 12px;text-align:center;">',
    '<img src="cid:beforeImage" alt="Visuel avant retouche" style="max-width:260px;width:100%;height:auto;border-radius:6px;box-shadow:0 2px 6px rgba(0,0,0,0.15);">',
    '<div style="margin-top:8px;font-size:12px;color:#555;">Avant retouche</div>',
    '</div>',
    '<div style="display:inline-block;margin:0 12px;text-align:center;">',
    '<img src="cid:afterImage" alt="Visuel après retouche" style="max-width:260px;width:100%;height:auto;border-radius:6px;box-shadow:0 2px 6px rgba(0,0,0,0.15);">',
    '<div style="margin-top:8px;font-size:12px;color:#555;">Après retouche</div>',
    '</div>',
    '</div>',
    '<p>Voici un avant/après illustre comment une photo peut être améliorée tout en restant honnête sur l’état du bien. Cette qualité de retouche rapide vous intéresserait-elle pour vos programmes en cours ? Je vous propose volontiers un essai gratuit : envoyez-moi une de vos photos, et vous jugerez du résultat par vous-même.</p>',
    '<p>Je reste bien sûr à votre écoute pour toute question, dans l’espoir d’une collaboration à venir.</p>',
    `<p>Cordialement,<br>${escapeHtml_(CONFIG.senderName)}</p>`,
    '<p>— mail envoyé par Jahwin Schmitt-Flore<br>SIREN : 941 338 659<br>📧 <a href="mailto:schmittdarwin42@gmail.com">schmittdarwin42@gmail.com</a> | 📞 <a href="tel:+33769229493">07 69 22 94 93</a></p>',
  ].join('\n');

  const textBody = [
    'Bonjour,',
    '',
    `J'ai découvert l'engagement de ${prospectName} pour des projets immobiliers durables, et je salue la qualité de vos réalisations. Pour refléter cet engagement jusque dans vos supports de vente, il est crucial de disposer de visuels percutants et authentiques. À ce titre, permettez-moi de vous présenter ImageUp, notre service de retouche photographique immobilière ultra-rapide et fiable, qui pourrait grandement servir votre communication.`,
    '',
    'Ce que nous proposons :',
    '- Service express : vos photos retouchées en 24h chrono, pour soutenir la rapidité de vos mises en vente et promotions.',
    '- Prix concurrentiels : une offre économique adaptée aux promoteurs indépendants, vous assurant un excellent rapport qualité-prix.',
    "- Retouche fidèle au bien : chaque image est optimisée (lumière, couleurs, netteté) tout en conservant l'aspect réel du bien, afin de rester transparent vis-à-vis des acquéreurs.",
    '',
    'Avant/Après : voir l’illustration jointe dans le message.',
    "Voici un avant/après illustre comment une photo peut être améliorée tout en restant honnête sur l'état du bien. Cette qualité de retouche rapide vous intéresserait-elle pour vos programmes en cours ? Je vous propose volontiers un essai gratuit : envoyez-moi une de vos photos, et vous jugerez du résultat par vous-même.",
    '',
    'Je reste bien sûr à votre écoute pour toute question, dans l’espoir d’une collaboration à venir.',
    '',
    'Cordialement,',
    "L’équipe ImageUp (by L4win Store)",
    '',
    '— mail envoyé par Jahwin Schmitt-Flore',
    'SIREN : 941 338 659',
    'schmittdarwin42@gmail.com | 07 69 22 94 93',
  ].join('\n');

  return { textBody, htmlBody };
}

/**
 * Récupère les images "avant" et "après" depuis Google Drive pour les insérer en tant qu’images en ligne.
 * @returns {{before: GoogleAppsScript.Base.Blob, after: GoogleAppsScript.Base.Blob}}
 */
function getInlineImages_() {
  const beforeId = CONFIG.beforeImageFileId;
  const afterId = CONFIG.afterImageFileId;

  if (!beforeId || beforeId === 'PASTE_BEFORE_IMAGE_FILE_ID') {
    throw new Error('Configurer CONFIG.beforeImageFileId avec l’ID Drive de la photo "avant".');
  }

  if (!afterId || afterId === 'PASTE_AFTER_IMAGE_FILE_ID') {
    throw new Error('Configurer CONFIG.afterImageFileId avec l’ID Drive de la photo "après".');
  }

  return {
    before: DriveApp.getFileById(beforeId).getBlob(),
    after: DriveApp.getFileById(afterId).getBlob(),
  };
}

/**
 * Retourne la feuille Prospects, ou lève une erreur si elle est introuvable.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getProspectSheet_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.prospectsSheetName);
  if (!sheet) {
    throw new Error(`Onglet "${CONFIG.prospectsSheetName}" introuvable.`);
  }
  return sheet;
}

/**
 * Supprime les espaces superflus et convertit en chaîne.
 * @param {*} value
 * @returns {string}
 */
function normalizeValue_(value) {
  if (value === null || value === undefined) {
    return '';
  }
  return String(value).trim();
}

/**
 * Échappe les caractères HTML sensibles.
 * @param {string} value
 * @returns {string}
 */
function escapeHtml_(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
