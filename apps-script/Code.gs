/**
 * LA MÉMOIRE — ÉCLAIRAGE MODERNE SARAN
 * Un seul fichier, un seul projet Apps Script. Fonction à lancer : passe
 *
 * Étape 1 (ce fichier) — TES PO ENVOYÉS AUX FOURNISSEURS
 *   Lit les courriels envoyés « Bon de Commande d'Achat 408107 v1 » / « Purchase Order 408100 v1 »
 *   et le PDF joint (Eclairage_Moderne_Saran_PO_408107_v1.pdf) :
 *     - fournisseur (« Acheté à »), date commandée, date requise ;
 *     - chaque ligne : produit, description, quantité, montant,
 *       et son TAG = la commande du client (ex. « Tag:119698-CAMPBELL DANIEL ») ;
 *     - sous-total, TPS, TVQ, total.
 *   Un PO corrigé (v2, v3) remplace la version précédente du même numéro.
 *
 * RÈGLE : rien n'est deviné. Si la somme des lignes ne donne pas le sous-total, ou si les taxes
 * ne balancent pas, le PO n'est PAS inscrit : il est noté « À VÉRIFIER » dans le journal.
 *
 * ESSAI = true  -> écrit SEULEMENT dans le journal (rien dans Drive, rien dans un Sheet).
 *                  Montre aussi le texte lu dans les 2 premiers PDF, pour ajuster la lecture.
 * ESSAI = false -> crée le Sheet « La mémoire — Éclairage » dans le dossier (s'il n'existe pas),
 *                  inscrit les PO (onglets « Commandes fournisseurs » et « Lignes PO »),
 *                  et range les PDF dans La mémoire › Fournisseurs › <fournisseur>.
 *
 * Version 9.1 — 3 nov. 2026 : (1) chaque ligne n'est jugée que sur SA phrase (le texte après son code s'arrête au
 *   bout de la ligne) : le courriel SPI du 3 nov. (PO 408072) avait mis en BO le LEDG9 à cause de la phrase
 *   suivante ; (2) les numéros du fournisseur écrits dans notre description (« STD 70161 ») servent aussi à
 *   reconnaître la ligne ; (3) Tag avec un NOM seulement (« TAG: DUQUETTE », « tag : Lambert + stock ») :
 *   relié à la commande en cours de ce client s'il n'y en a qu'une, sinon le nom est inscrit sans numéro.
 * Version 9 — 3 nov. 2026 : STATUTS AUTOMATIQUES (AUTO_STATUTS) — quand un fournisseur écrit, les lignes du PO
 *   avancent toutes seules : avis d'expédition → « Expédié » + date prévue ; BO / ETA → « BO » + date ;
 *   article annulé / discontinué → « Annulé » ; confirmation → « Confirmé » ; prêt à ramasser → « Prêt à ramasser ».
 *   Chaque ligne est jugée à part, d'après ce qui est écrit juste après SON code (un même courriel peut dire
 *   BO pour une ligne et annulé pour une autre, ex. Eglo PO 408082). Le PDF joint du fournisseur est lu aussi
 *   (Canarm, Kuzco… mettent les détails dans le PDF). Un statut n'avance jamais à reculons et ne touche pas
 *   une ligne Reçu / Livré / Annulé. Chaque changement est inscrit au fil « Notes » du CRM (Par = « Script »).
 *   Aussi : le Tag écrit dans la description (« TAG: GESTION ADG - PAQUET SAILLANT 119722 ») est lu ;
 *   LECTEUR passe à 6 (les PO du dernier mois sont relus une fois).
 * Version 8 — 27 oct. 2026 (après l'essai des courriels) : EN VRAI ; « ETA » ne prend plus « retaining » ni « details » ;
 *   nouveaux types « Prêt à ramasser » et « Expédition prévue » ; les relevés de compte sont ignorés.
 * Version 7 — 27 oct. 2026 : ÉTAPE 2 (MODE ESSAI à part) — les courriels des fournisseurs (confirmations, BO/ETA,
 *   expéditions, annulations) reliés à tes PO par leur numéro, dans l'onglet « Courriels PO ». Rien n'est deviné :
 *   aucun statut n'est changé, le CRM montre ce que le fournisseur a dit et c'est vous qui confirmez.
 * Version 6 — 27 oct. 2026 : le suivi fait dans le CRM (Statut, Date prévue, Note de suivi) est GARDÉ quand un PO est relu
 *   ou remplacé par une nouvelle version ; un seul déclencheur (installer : chaque heure, seulement les 30 derniers jours) ;
 *   « À vérifier » se vide tout seul quand le PO finit par être lu.
 * Version 5 — 27 oct. 2026 : lignes SANS PRIX (prix à confirmer, SNOC, Lit Up), codes Canarm « #… », en-tête de page 2,
 *   et chaque PO déjà inscrit est RELU une fois avec la nouvelle lecture (colonne « Lu avec »).
 * Version 4 — 27 oct. 2026 : EN VRAI, depuis le 1er juin 2026. Onglet « À vérifier » (pas relu à chaque passe).
 * Version 3 — 27 oct. 2026 (après le 2e essai) : codes collés au texte (DALS), retours Standard Pro,
 *   annulations de BO signalées clairement.
 * Version 2 — 27 oct. 2026 (après le 1er essai) : bon fournisseur, notes de ligne (« TAG: STOCK », « add on… »),
 *   PO de 2 pages, retours (RMA) inscrits en négatif avec le PO d'origine, PO en anglais.
 * Aucun déclencheur pour l'instant : on le lance à la main.
 */

const ESSAI = false;   // essais du 27 oct. 2026 validés (25 PO sur 28 lus, les 3 autres signalés)
const DOSSIER_MEMOIRE = '1HE-T772uLZAWsFKQd5MOH5wBaFGDr0KO';   // La mémoire — Éclairage
const DEPUIS = '2026/06/01';                                   // début de l'année financière (choisi le 27 oct. 2026)
const MOI = 'service@eclairagesaran.com';
const NOM_SHEET = 'La mémoire — Éclairage';
const LECTEUR = 7;   // version de la lecture des PO : un PO lu avec une version plus vieille est relu une fois
const AUTO_STATUTS = true;   // v9 : false = le script ne change plus aucun statut (le CRM montre seulement les courriels)
const PDF_PAR_PASSE = 25;    // v9 : PDF des fournisseurs lus au plus par passe (chaque lecture prend quelques secondes)

const ENTETE_PO = ['PO', 'Version', 'Type', 'Fournisseur', 'Envoyé à', 'Date commandée', 'Date requise',
  'Sous-total', 'Total', 'Lignes', 'PDF', 'Lien courriel', 'Inscrit le', 'Lu avec'];
const ENTETE_LIGNES = ['PO', 'Ligne', 'Produit', 'Description', 'Qté', 'Unité', 'Montant',
  'Commande client', 'Client', 'Réquisition', 'Statut', 'Note', 'Date prévue', 'Note suivi', 'Modifié le', 'Par'];
// Colonnes K (Statut) et M à P : écrites par le CRM (suivi). Le script les garde toujours.

// ───────────────────────── Point d'entrée ─────────────────────────

function passe() {
  const debut = Date.now();
  poEnvoyes_(debut);
  if (Date.now() - debut < 4 * 60 * 1000) courrielsFournisseurs_(debut);
  else Logger.log('Courriels des fournisseurs : remis à la prochaine passe (manque de temps).');
}

// ───────────────────────── Étape 2 : ce que les fournisseurs répondent ─────────────────────────
// COURRIELS_ESSAI = true : regarde les 14 derniers jours et écrit SEULEMENT dans le journal.
// COURRIELS_ESSAI = false : rattrape depuis COURRIELS_DEPUIS une semaine à la fois, puis suit les 2 derniers jours à chaque passe.
const COURRIELS_ESSAI = false;   // essai du 27 oct. 2026 validé (79 courriels reliés sur 14 jours)
const COURRIELS_DEPUIS = '2026/08/15';
const ENTETE_COURRIELS = ['Clé', 'PO', 'Date', 'De', 'Courriel', 'Sujet', 'Type', 'ETA (texte)', 'Extrait', 'Lien', 'Lu le'];

function courrielsFournisseurs_(debut) {
  const props = PropertiesService.getScriptProperties();
  const ss = sheet_();
  const shPO = ss.getSheetByName('Commandes fournisseurs');
  if (!shPO || shPO.getLastRow() < 2) return;
  const poConnus = new Set(shPO.getRange(2, 1, shPO.getLastRow() - 1, 1).getDisplayValues().map(r => r[0]));
  // Domaines des fournisseurs (à qui on envoie les PO) : seuls leurs PDF sont lus pour trouver le no de PO.
  const domaines = new Set(shPO.getRange(2, 5, shPO.getLastRow() - 1, 1).getDisplayValues().map(r => (String(r[0]).split('@')[1] || '').toLowerCase()).filter(Boolean));
  const shC = COURRIELS_ESSAI ? null : onglet_(ss, 'Courriels PO', ENTETE_COURRIELS);
  const vus = new Set(shC && shC.getLastRow() > 1 ? shC.getRange(2, 1, shC.getLastRow() - 1, 1).getDisplayValues().map(r => String(r[0]).split('|')[0]) : []);
  const stats = { messages: 0, relies: 0, pdf: 0 };
  const neuves = [];
  const auto = AUTO_STATUTS && !COURRIELS_ESSAI ? moteurStatuts_(ss) : null;   // v9
  const tropLong = () => Date.now() - debut > 5 * 60 * 1000;
  const jour = d => Utilities.formatDate(d, 'America/Toronto', 'yyyy/MM/dd');
  const base = ' -in:sent -from:' + MOI + ' -category:promotions -category:social -in:spam -in:trash';

  const traiterFils = requete => {
    for (const fil of GmailApp.search(requete + base, 0, 300)) {
      if (tropLong()) return false;
      for (const m of fil.getMessages()) {
        if (/eclairagesaran\.com/i.test(m.getFrom())) continue;
        if (vus.has(m.getId())) continue;
        stats.messages++;
        const sujet = m.getSubject(), corps = m.getPlainBody().replace(/\r/g, '');
        let pos = [...new Set((sujet + '\n' + corps).match(/\b40[4-9]\d{3}\b/g) || [])].filter(p => poConnus.has(p));
        // Pas de no de PO dans le courriel (ex. DVI « Order Acknowledgement 305244 ») : on le cherche dans le PDF joint.
        const domaine = (String(m.getFrom()).match(/@([\w.-]+)/) || [])[1] || '';
        // PDF joint du fournisseur (jamais notre propre PO renvoyé en pièce jointe).
        const pdfJoint = () => m.getAttachments({ includeInlineImages: false }).find(a => /\.pdf$/i.test(a.getName()) && !/Eclairage_Moderne_Saran_PO_/i.test(a.getName()));
        let textePdf = '';
        if (!pos.length && stats.pdf < PDF_PAR_PASSE && [...domaines].some(d => domaine.toLowerCase().endsWith(d.split('.').slice(-2).join('.')))) {
          const pdf = pdfJoint();
          if (pdf) { stats.pdf++; textePdf = lirePdf_(pdf.copyBlob()); pos = [...new Set(String(textePdf).match(/\b40[4-9]\d{3}\b/g) || [])].filter(p => poConnus.has(p)); }
        }
        vus.add(m.getId());
        if (!pos.length) continue;
        const texte = sujet + ' ' + corps.split(/\n(?:Le .{5,80}a écrit|On .{5,80}wrote|-{2,}\s*Original|De ?:|From:|_{8,})/)[0];
        const type = typeCourriel_(texte, m);
        if (type === 'Relevé') continue;                                   // relevé de compte : pas une nouvelle sur une commande
        // v9 : le PDF du fournisseur est lu aussi quand le no de PO est dans le courriel (les détails des lignes y sont souvent).
        if (auto && !textePdf && type !== 'Facture' && stats.pdf < PDF_PAR_PASSE) { const pdf = pdfJoint(); if (pdf) { stats.pdf++; textePdf = lirePdf_(pdf.copyBlob()); } }
        const eta = ((texte.match(/\bETA\b[^.\n]{0,70}/i) || texte.match(/\b(?:back ?order\w*|en souffrance|disponib\w*|dispo\b|exp[ée]di(?:era|er|é)\b|will ship|ship(?:ping)? date|ships?\b|pr[êe]te? pour la cueillette|ready for pick ?up|jusqu'?(?:au|à)\b|available\b|delay\w*|retard\w*)[^.\n]{0,70}/i) || [''])[0]).replace(/\s+/g, ' ').trim().slice(0, 90);
        const extrait = corps.split(/\n(?:Le .{5,80}a écrit|On .{5,80}wrote|-{2,}\s*Original|De ?:|From:|_{8,})/)[0].replace(/\[[^\]]*\]|<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 300);
        const de = String(m.getFrom()).replace(/<.*>/, '').replace(/"/g, '').trim() || m.getFrom();
        pos.forEach(po => {
          const ligne = [m.getId() + '|' + po, po, Utilities.formatDate(m.getDate(), 'America/Toronto', 'yyyy-MM-dd'), de, (String(m.getFrom()).match(/[\w.+-]+@[\w.-]+/) || [''])[0].toLowerCase(),
            sujet.slice(0, 120), type, eta, extrait, 'https://mail.google.com/mail/?authuser=' + MOI + '#all/' + fil.getId(), Utilities.formatDate(new Date(), 'America/Toronto', 'yyyy-MM-dd HH:mm')];
          neuves.push(ligne); stats.relies++;
          Logger.log((COURRIELS_ESSAI ? '[essai] ' : '') + 'PO ' + po + ' | ' + ligne[2] + ' | ' + de + ' | ' + type + (eta ? ' | « ' + eta + ' »' : '') + ' | ' + sujet.slice(0, 70));
          if (auto) auto.appliquer(po, texte + '\n' + textePdf, type, ligne[2], de);   // v9 : les lignes avancent toutes seules
        });
      }
    }
    return true;
  };

  if (COURRIELS_ESSAI) traiterFils('newer_than:14d');
  else if (props.getProperty('courrielsFini')) traiterFils('newer_than:2d');
  else {
    // Rattrapage une semaine à la fois ; on n'avance que si la semaine est finie.
    let cur = props.getProperty('courrielsDepuis') || COURRIELS_DEPUIS;
    while (!tropLong()) {
      const d = new Date(cur.replace(/\//g, '-') + 'T12:00:00'), fin = new Date(d.getTime() + 7 * 864e5);
      if (d > new Date()) { props.setProperty('courrielsFini', jour(new Date())); break; }
      if (!traiterFils('after:' + cur + ' before:' + jour(fin))) break;
      cur = jour(fin); props.setProperty('courrielsDepuis', cur);
    }
  }
  if (shC && neuves.length) shC.getRange(shC.getLastRow() + 1, 1, neuves.length, ENTETE_COURRIELS.length).setValues(neuves);
  if (auto) auto.ecrire();
  Logger.log((COURRIELS_ESSAI ? 'MODE ESSAI (courriels) — rien écrit. ' : '') + 'Courriels des fournisseurs lus : ' + stats.messages + ' | reliés à un PO : ' + stats.relies + ' | PDF lus pour trouver le PO : ' + stats.pdf +
    (!COURRIELS_ESSAI && !props.getProperty('courrielsFini') ? ' | rattrapage rendu au ' + (props.getProperty('courrielsDepuis') || COURRIELS_DEPUIS) : ''));
}

// Ce que dit le courriel, par mots clés. Le CRM l'affiche ; personne ne change de statut tout seul.
function typeCourriel_(texte, m) {
  const sujet = m ? m.getSubject() : '';
  if (/statement|relev[ée] de compte|[ée]tat de compte/i.test(sujet)) return 'Relevé';
  // Le sujet d'abord (courriels automatiques des fournisseurs)
  if (/shipping notice|confirmation d.exp[ée]dition|ship notification|packing slip|has been shipped/i.test(sujet)) return 'Expédition';
  if (/acknowledg|confirmation de (la )?commande|order confirmation/i.test(sujet) && !/cancel|annul|back ?order|\bETA\b/i.test(sujet)) return 'Confirmation';
  // Phrases toutes faites (« back-ordered items will be indicated… ») : ignorées
  const t = String(texte).replace(/[^.\n]*(will be indicated|seront indiqu)[^.\n]*[.\n]/gi, ' ');
  if (/\bcancel|\bannul/i.test(t)) return 'Annulation';
  if (/\bback ?orders?\b|\bbackordered\b|\bB\/O\b|en souffrance|\brupture\b|\bETA\b|\bdelay(s|ed)?\b|\bretard|disponible (le|en|vers)|dispo (le|en)/i.test(t)) return 'BO / ETA';
  if (/pr[êe]te? pour (la )?(cueillette|ramassage)|ready for pick ?up|\bpick ?up\b.{0,20}\bready/i.test(t)) return 'Prêt à ramasser';
  if (/\bshipped\b|\bexp[ée]di[ée]e?s?\b|tracking|suivi de colis|waybill|bon de livraison/i.test(t)) return 'Expédition';
  if (/will ship|\bship (it )?(today|tomorrow|this week|asap|as soon|on |monday|tuesday|wednesday|thursday|friday)|exp[ée]diera|sera exp[ée]di|exp[ée]dier (la m[êe]me journ[ée]e|aujourd|demain)/i.test(t)) return 'Expédition prévue';
  if (/\binvoice\b|\bfacture\b/i.test(t)) return 'Facture';
  if (/acknowledg|confirmation de (la )?commande|order confirmation|voici la confirmation|\bconfirm/i.test(t)) return 'Confirmation';
  return 'Autre';
}

// ───────────────────────── v9 : statuts automatiques ─────────────────────────
// Le statut d'une ligne avance seulement (Commandé → Confirmé → BO → Expédié → Prêt à ramasser) ;
// Reçu, Livré et Annulé ne sont jamais touchés. Une ligne modifiée à la main dans le CRM APRÈS la date
// du courriel n'est pas touchée non plus (la personne en sait plus que le courriel).
const ENTETE_NOTES = ['id', 'Date', 'Par', 'PO', 'Ligne', 'Produit', 'Commande client', 'Client', 'Texte', 'Fournisseur'];
const RANG_STATUT = { 'Commandé': 0, 'Confirmé': 1, 'BO': 2, 'Expédié': 3, 'Prêt à ramasser': 4 };

function moteurStatuts_(ss) {
  const shL = ss.getSheetByName('Lignes PO');
  if (!shL || shL.getLastRow() < 2) return null;
  const vals = shL.getRange(2, 1, shL.getLastRow() - 1, ENTETE_LIGNES.length).getDisplayValues();
  const parPO = {};
  vals.forEach((r, i) => { (parPO[r[0]] = parPO[r[0]] || []).push(i); });
  const changees = {}, applique = {}, notes = [];
  const maintenant = () => Utilities.formatDate(new Date(), 'America/Toronto', 'yyyy-MM-dd HH:mm');

  function appliquer(po, texteBrut, type, dateMsg, de) {
    const idx = parPO[po] || [];
    if (!idx.length) return;
    const t = normaliser_(texteBrut);
    const parLigne = analyserLignes_(t, idx.map(i => vals[i][2]), dateMsg, idx.map(i => vals[i][3]));
    const mentionnees = parLigne.some(a => a.vu);
    const dateLivraison = datePrevue_(t, dateMsg);
    idx.forEach((i, k) => {
      const a = parLigne[k];
      let statut = null, date = '';
      if (a.statut) { statut = a.statut; date = a.date || (a.statut === 'Expédié' ? dateLivraison : ''); }
      else if (type === 'Expédition') { if (!mentionnees || a.vu) { statut = 'Expédié'; date = dateLivraison; } }
      else if (type === 'Prêt à ramasser') { if (!mentionnees || a.vu) statut = 'Prêt à ramasser'; }
      else if (type === 'Confirmation') statut = 'Confirmé';
      else if (type === 'BO / ETA' && idx.length === 1) { statut = 'BO'; date = dateDansTexte_(t, dateMsg); }
      else if (type === 'Annulation' && idx.length === 1) statut = 'Annulé';
      if (statut) changer(i, statut, date, dateMsg, de);
    });
  }

  function changer(i, statut, date, dateMsg, de) {
    const r = vals[i];
    const actuel = r[10] || 'Commandé';
    if (/^(Reçu|Livré|Annulé|Retour)/.test(actuel)) return;
    if (statut === 'Annulé' || statut === 'BO') { if (!/^(Commandé|Confirmé|BO)$/.test(actuel)) return; }
    else if (!(RANG_STATUT[statut] > (RANG_STATUT[actuel] || 0))) return;
    const mod = String(r[14]).slice(0, 10);
    if (/^\d{4}-\d{2}-\d{2}$/.test(mod) && mod > dateMsg && !/^Script/.test(r[15])) return;     // modifié à la main après le courriel
    if (applique[i] && applique[i] > dateMsg) return;                                          // un courriel plus récent a déjà décidé
    if (statut === actuel && (!date || date === r[12])) return;                                // rien de nouveau
    const avant = actuel + (r[12] ? ' (prévu ' + r[12] + ')' : '');
    r[10] = statut; if (date) r[12] = date;
    r[13] = 'auto : courriel du ' + dateMsg + ' (' + String(de).slice(0, 40) + ')';
    r[14] = maintenant(); r[15] = 'Script';
    changees[i] = true; applique[i] = dateMsg;
    notes.push([Utilities.getUuid().slice(0, 12), Utilities.formatDate(new Date(), 'America/Toronto', 'yyyy-MM-dd HH:mm:ss'), 'Script', r[0], r[1], r[2], r[7], r[8],
      'Statut → ' + statut + (date ? ' · prévu le ' + date : '') + ' (avant : ' + avant + ') · d\'après le courriel du ' + dateMsg + ' de ' + String(de).slice(0, 40), '']);
  }

  function ecrire() {
    const lignes = Object.keys(changees).map(Number);
    if (lignes.length) {
      // Avant d'écrire : la rangée est-elle toujours la même ligne ? (le CRM ou une autre passe a pu bouger les rangées)
      const ids = shL.getRange(2, 1, shL.getLastRow() - 1, 3).getDisplayValues();
      lignes.forEach(i => {
        const r = vals[i], x = ids[i];
        if (!x || x[0] !== r[0] || x[1] !== r[1] || x[2] !== r[2]) { Logger.log('Statut auto non écrit (rangée déplacée) : PO ' + r[0] + ' ' + r[2]); return; }
        shL.getRange(i + 2, 11).setNumberFormat('@').setValue(r[10]);
        shL.getRange(i + 2, 13, 1, 4).setNumberFormat('@').setValues([[r[12], r[13], r[14], r[15]]]);
      });
    }
    if (notes.length) { const shN = onglet_(ss, 'Notes', ENTETE_NOTES); shN.getRange(shN.getLastRow() + 1, 1, notes.length, ENTETE_NOTES.length).setNumberFormat('@').setValues(notes); }
    Logger.log('Statuts automatiques : ' + lignes.length + ' ligne(s) mise(s) à jour.');
    notes.forEach(n => Logger.log('   PO ' + n[3] + ' ' + n[5] + ' — ' + n[8]));
  }
  return { appliquer, ecrire };
}

function normaliser_(s) { return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/[ \t\u00a0]+/g, ' ').replace(/ *\r?\n[\s]*/g, '\n'); }

// Les façons dont un fournisseur écrit un de nos codes : « EG204922A » → « 204922a », « KUZVL18224-BK-UNV » → « vl18224-bk-unv »…
function variantesCode_(code) {
  const c = String(code || '').toLowerCase().replace(/^[#*]+/, '');
  const v = new Set([c]);
  const sans = c.replace(/^[a-z]{2,3}/, '');
  if (sans !== c && sans.length >= 6 && /\d/.test(sans)) v.add(sans);
  return [...v].filter(x => x.length >= 4);
}

// Pour chaque ligne : le code est-il dans le texte (vu), et que dit le texte juste après LUI (jusqu'au code suivant) ?
function analyserLignes_(t, codes, dateMsg, descs) {
  const pos = codes.map((code, k) => {
    let best = -1, lg = 0;
    // v9.1 : aussi le no du fournisseur écrit dans notre description (« STD 70161 », « STD67093 »)
    const nos = (String((descs || [])[k] || '').match(/(?:^|[^0-9])(\d{5,6})(?![0-9])/g) || []).map(x => x.replace(/^[^0-9]/, '')).filter(x => !/^20\d\d/.test(x) || x.length > 5);
    variantesCode_(code).concat(nos).forEach(v => {
      let i = t.indexOf(v);
      while (i >= 0) {
        const avant = i === 0 ? ' ' : t[i - 1], apres = t[i + v.length] || ' ';
        if (!/[a-z0-9]/.test(avant) && !/[a-z0-9]/.test(apres)) { if (best < 0 || i < best) { best = i; lg = v.length; } break; }
        i = t.indexOf(v, i + 1);
      }
    });
    return { i: best, fin: best + lg };
  });
  return pos.map((p, k) => {
    if (p.i < 0) return { vu: false };
    const prochain = pos.filter((q, j) => j !== k && q.i > p.i).map(q => q.i);
    let fen = t.slice(p.fin, Math.min(p.fin + 180, prochain.length ? Math.min(...prochain) : Infinity));
    // v9.1 : on s'arrête au bout de SA ligne (ou à « ligne 16 », « line 3 »…) ; si le code est seul sur sa ligne
    // (tableau de PDF), on prend la ligne suivante.
    const lignesFen = fen.split('\n');
    fen = lignesFen[0].replace(/[\s\-–:)(]+/g, '').length < 3 && lignesFen.length > 1 ? lignesFen[0] + ' ' + lignesFen[1] : lignesFen[0];
    fen = fen.split(/\b(?:ligne|line|item|article)\s*#?\s*\d/)[0];
    const date = dateDansTexte_(fen, dateMsg);
    if (/cancel|annul|discontinu|no longer (available|made|produced|offered)|n.est plus (disponible|fabrique|offert)|obsolete/.test(fen)) return { vu: true, statut: 'Annulé' };
    const motBO = /back ?order(?:ed)?|\bb\/o\b|rupture|en souffrance|not (?:yet )?(?:available|in stock)|pas (?:encore )?(?:disponible|en stock)|out of stock|\beta\b|delay|retard/;
    const zeroBO = /(?:back ?order(?:ed)?|b\/o)\s*(?:qty|quantity|quantite)?\s*:?\s*0(?:\.0+)?\b/;
    if ((motBO.test(fen) && !zeroBO.test(fen)) || (date && /not confirmed|non confirme|approx|environ|\bmid-|\bmi-|early|late|debut|fin |vers |around/.test(fen))) return { vu: true, statut: 'BO', date };
    if (/shipped|expedie|ship date|tracking/.test(fen)) return { vu: true, statut: 'Expédié', date: '' };
    return { vu: true };
  });
}

const MOIS_ = { jan: 1, janv: 1, janvier: 1, january: 1, feb: 2, fev: 2, fevr: 2, fevrier: 2, february: 2, mar: 3, mars: 3, march: 3, apr: 4, avr: 4, avril: 4, april: 4,
  may: 5, mai: 5, jun: 6, juin: 6, june: 6, jul: 7, juil: 7, juillet: 7, july: 7, aug: 8, aou: 8, aout: 8, august: 8, sep: 9, sept: 9, septembre: 9, september: 9,
  oct: 10, octobre: 10, october: 10, nov: 11, novembre: 11, november: 11, dec: 12, decembre: 12, december: 12 };

// Une date dans un bout de texte déjà normalisé. « mid-December » → le 15, « early/début » → le 5, « late/fin » → le 25,
// un mois seul → le 1er. L'année : celle du courriel, ou la suivante si la date tombe plus d'un mois avant le courriel.
function dateDansTexte_(s, dateMsg) {
  const p2 = n => ('0' + n).slice(-2), ok = (y, m, d) => m >= 1 && m <= 12 && d >= 1 && d <= 31 ? y + '-' + p2(m) + '-' + p2(d) : '';
  const base = String(dateMsg || '').slice(0, 10) || Utilities.formatDate(new Date(), 'America/Toronto', 'yyyy-MM-dd'), y0 = Number(base.slice(0, 4));
  const avecAnnee = (m, d) => { let r = ok(y0, m, d); if (r && r < base && (new Date(base) - new Date(r)) > 30 * 864e5) r = ok(y0 + 1, m, d); return r; };
  let m = s.match(/\b(20\d{2})-(\d{1,2})-(\d{1,2})\b/); if (m) return ok(+m[1], +m[2], +m[3]);
  m = s.match(/\b(\d{1,2})\/(\d{1,2})\/(20\d{2})\b/); if (m) return ok(+m[3], +m[1], +m[2]);                 // M/J/AAAA (fournisseurs américains)
  m = s.match(/\b(mid|mi|early|debut|late|fin|end of)[\s-]+(?:de |d['’] ?)?([a-z]{3,9})\b/);
  if (m && MOIS_[m[2]]) return avecAnnee(MOIS_[m[2]], /mi/.test(m[1]) ? 15 : /early|debut/.test(m[1]) ? 5 : 25);
  m = s.match(/\b([a-z]{3,9})\.?\s+(\d{1,2})(?:st|nd|rd|th)?\b/); if (m && MOIS_[m[1]]) return avecAnnee(MOIS_[m[1]], +m[2]);
  m = s.match(/\b(\d{1,2})(?:er)?\s+([a-z]{3,9})\b/); if (m && MOIS_[m[2]]) return avecAnnee(MOIS_[m[2]], +m[1]);
  m = s.match(/\b(?:en|in|for|pour)\s+([a-z]{3,9})\b/); if (m && MOIS_[m[1]] && m[1].length > 3) return avecAnnee(MOIS_[m[1]], 1);
  return '';
}

// La date de livraison annoncée dans tout le courriel : seulement juste après un mot de livraison
// (pour ne pas prendre la date d'expédition : « Estimated Delivery date: 11/9/2026 … Ship Date: 11/2/2026 »).
function datePrevue_(t, dateMsg) {
  const z = t.match(/(estimated delivery|delivery date|expected delivery|\beta\b|livraison prevue|date prevue|prevu le|arrivee prevue)[\s\S]{0,45}/);
  return z ? dateDansTexte_(z[0], dateMsg) : '';
}

// ───────────────────────── Étape 1 : PO envoyés ─────────────────────────

function poEnvoyes_(debut) {
  const stats = { lus: 0, inscrits: 0, remplaces: 0, deja: 0, aVerifier: 0, textesMontres: 0 };

  // 1) Les courriels de PO envoyés, une version par numéro (la plus haute ; à version égale, la plus récente).
  const parPO = {};
  const props = PropertiesService.getScriptProperties();
  const fenetre = !ESSAI && props.getProperty('rattrapageFini') ? 'newer_than:30d' : 'after:' + DEPUIS;   // une fois tout inscrit : seulement le dernier mois
  chercher_('in:sent from:' + MOI + ' subject:("Bon de Commande d\'Achat" OR "Purchase Order") has:attachment ' + fenetre)
    .forEach(fil => fil.getMessages().forEach(m => {
      if (m.getFrom().toLowerCase().indexOf(MOI) < 0) return;
      const sujet = m.getSubject();
      if (/^\s*(re|tr|fwd?|fw)\s*:/i.test(sujet)) return;                   // réponses et transferts : pas un envoi de PO
      const s = sujet.match(/(?:Bon de Commande d'Achat|Purchase Order)\s+(\d{6})\s+v(\d+)/i);
      if (!s) return;
      const pdf = m.getAttachments({ includeInlineImages: false })
        .filter(a => /\.pdf$/i.test(a.getName()) || /pdf/i.test(a.getContentType()))
        .find(a => new RegExp('PO_' + s[1]).test(a.getName())) || null;
      if (!pdf) return;
      const v = Number(s[2]), ex = parPO[s[1]];
      if (!ex || v > ex.version || (v === ex.version && m.getDate() > ex.msg.getDate()))
        parPO[s[1]] = { po: s[1], version: v, msg: m, pdf: pdf, fil: fil };
    }));

  // 2) Ce qui est déjà inscrit (en vrai seulement).
  const ss = ESSAI ? null : sheet_();
  const shPO = ss ? onglet_(ss, 'Commandes fournisseurs', ENTETE_PO) : null;
  const shL = ss ? onglet_(ss, 'Lignes PO', ENTETE_LIGNES) : null;
  const shV = ss ? onglet_(ss, 'À vérifier', ['PO', 'Version', 'Envoyé à', 'Date', 'Raison', 'Lien courriel', 'Noté le', 'Lu avec']) : null;
  const lignesV = shV && shV.getLastRow() > 1 ? shV.getRange(2, 1, shV.getLastRow() - 1, 8).getDisplayValues() : [];
  const dejaNotes = new Set(lignesV.filter(r => Number(r[7]) >= LECTEUR).map(r => r[0] + '|' + r[1]));
  const ligneV = {}; lignesV.forEach((r, i) => { ligneV[r[0] + '|' + r[1]] = i + 2; });
  const aRetirerV = [];
  const inscrits = {};
  if (shPO && shPO.getLastRow() > 1)
    shPO.getRange(2, 1, shPO.getLastRow() - 1, ENTETE_PO.length).getDisplayValues().forEach((r, i) => {
      inscrits[r[0]] = { version: Number(r[1]), ligne: i + 2, lecteur: Number(r[ENTETE_PO.length - 1]) || 0 }; });

  // 3) Lecture des PDF, du plus vieux au plus récent.
  const liste = Object.values(parPO).sort((a, b) => a.msg.getDate() - b.msg.getDate());
  for (const c of liste) {
    if (Date.now() - debut > 5 * 60 * 1000) { stats.reste = liste.length - liste.indexOf(c); Logger.log('Presque 6 minutes : je m\'arrête ici. Relance « passe » pour continuer.'); break; }
    const deja = inscrits[c.po];
    if (deja && (deja.version > c.version || (deja.version === c.version && deja.lecteur >= LECTEUR))) { stats.deja++; continue; }
    if (dejaNotes.has(c.po + '|' + c.version)) { stats.aVerifier++; continue; }        // déjà dans l'onglet « À vérifier »

    const texte = lirePdf_(c.pdf.copyBlob());
    stats.lus++;
    if (ESSAI && stats.textesMontres < 2) {
      stats.textesMontres++;
      Logger.log('===== Texte lu dans ' + c.pdf.getName() + ' =====\n' + String(texte).slice(0, 2500));
    }
    const lu = lirePO_(texte);
    const envoyeA = (c.msg.getTo().match(/[\w.+-]+@[\w.-]+/) || [''])[0].toLowerCase();
    const dateCmd = Utilities.formatDate(c.msg.getDate(), 'America/Toronto', 'yyyy-MM-dd');

    if (lu.erreur) {
      stats.aVerifier++;
      Logger.log('À VÉRIFIER — PO ' + c.po + ' v' + c.version + ' (' + envoyeA + ', ' + dateCmd + ') : ' + lu.erreur + ' — pas inscrit.');
      const lv = [c.po, c.version, envoyeA, dateCmd, lu.erreur, 'https://mail.google.com/mail/?authuser=' + MOI + '#all/' + c.fil.getId(),
        Utilities.formatDate(new Date(), 'America/Toronto', 'yyyy-MM-dd HH:mm'), LECTEUR];
      if (shV) { const k = ligneV[c.po + '|' + c.version]; k ? shV.getRange(k, 1, 1, lv.length).setValues([lv]) : shV.appendRow(lv); }
      if (ESSAI && stats.textesMontres < 12) { stats.textesMontres++; Logger.log('   texte lu : ' + String(texte).replace(/\s+/g, ' ').slice(0, 1500)); }
      continue;
    }

    if (ESSAI && !lu.fournisseur && stats.textesMontres < 12) { stats.textesMontres++; Logger.log('   (fournisseur non lu) texte : ' + String(texte).slice(0, 900)); }
    Logger.log((ESSAI ? '[essai] ' : '') + (deja ? 'REMPLACE v' + deja.version + ' → ' : '') +
      (lu.type === 'Retour' ? 'RETOUR (RMA) ' : '') + 'PO ' + c.po + ' v' + c.version + ' | ' + dateCmd + ' | ' + (lu.fournisseur || '?') + ' (envoyé à ' + envoyeA + ')' +
      (lu.dateRequise ? ' | requis le ' + lu.dateRequise : '') + ' | ' + lu.lignes.length + ' ligne(s) | ' +
      (lu.st === null ? 'PRIX À CONFIRMER (aucun montant sur le PO)' : lu.st.toFixed(2) + ' $ avant taxes, ' + lu.total.toFixed(2) + ' $ total'));
    lu.lignes.forEach(l => Logger.log('      ' + l.no + '. ' + l.produit + ' — ' + l.desc + ' × ' + l.qte + ' ' + l.unite + ' — ' +
      (l.sansPrix ? 'prix à confirmer' : l.montant.toFixed(2) + ' $') + ' — ' + (l.cmdClient ? 'client ' + l.cmdClient + ' ' + l.client : 'aucun tag') + (l.note ? '  [note : ' + l.note + ']' : '')));

    if (ESSAI) { deja ? stats.remplaces++ : stats.inscrits++; continue; }
    Object.keys(ligneV).filter(k => k.split('|')[0] === c.po).forEach(k => { aRetirerV.push(ligneV[k]); delete ligneV[k]; });   // lu maintenant : plus « À vérifier »

    // En vrai : PDF dans La mémoire › Fournisseurs › <fournisseur>, puis les lignes du Sheet.
    const nomF = (lu.fournisseur || envoyeA.split('@')[1] || 'Inconnu').replace(/[\/\\]/g, '-').slice(0, 60);
    const dossier = sousDossier_(sousDossier_(DriveApp.getFolderById(DOSSIER_MEMOIRE), 'Fournisseurs'), nomF);
    const nomPdf = 'PO ' + c.po + ' v' + c.version + ' (' + dateCmd + ').pdf';
    const it = dossier.getFilesByName(nomPdf);
    const pdfId = (it.hasNext() ? it.next() : dossier.createFile(c.pdf.copyBlob()).setName(nomPdf)).getId();
    const lienCourriel = 'https://mail.google.com/mail/?authuser=' + MOI + '#all/' + c.fil.getId();
    const lignePO = [c.po, c.version, lu.type, lu.fournisseur, envoyeA, dateCmd, lu.dateRequise, lu.st === null ? '' : lu.st.toFixed(2), lu.total === null ? '' : lu.total.toFixed(2),
      lu.lignes.length, pdfId, lienCourriel, Utilities.formatDate(new Date(), 'America/Toronto', 'yyyy-MM-dd HH:mm'), LECTEUR];

    const suivi = {};                                                      // suivi du CRM de l'ancienne version : gardé
    if (deja) {
      shPO.getRange(deja.ligne, 1, 1, ENTETE_PO.length).setValues([lignePO]);
      if (shL.getLastRow() > 1) shL.getRange(2, 1, shL.getLastRow() - 1, ENTETE_LIGNES.length).getValues().forEach(r => {
        if (String(r[0]) === c.po) suivi[r[1] + '|' + r[2]] = [r[10], r[12], r[13], r[14], r[15]]; });
      retirerLignes_(shL, c.po);                                           // les lignes de l'ancienne version
      stats.remplaces++;
    } else {
      shPO.appendRow(lignePO);
      inscrits[c.po] = { version: c.version, ligne: shPO.getLastRow(), lecteur: LECTEUR };
      stats.inscrits++;
    }
    // v9.1 : Tag avec un nom seulement → la commande en cours de ce client, s'il n'y en a qu'une.
    lu.lignes.filter(l => l.tagNom && !l.cmdClient).forEach(l => {
      const t = relierParNom_(shL, l.tagNom);
      if (t) { l.cmdClient = t.cmd; l.client = t.client; l.note = ('Tag « ' + l.tagNom + ' » relié à la commande ' + t.cmd + (l.note ? ' — ' + l.note : '')).slice(0, 150); }
      else { l.client = l.tagNom; l.note = ('Tag « ' + l.tagNom + ' » sans no de commande' + (l.note ? ' — ' + l.note : '')).slice(0, 150); }
      Logger.log('      Tag par nom « ' + l.tagNom + ' » (' + l.produit + ') → ' + (t ? 'commande ' + t.cmd + ' ' + t.client : 'aucune commande unique : nom inscrit seul'));
    });
    const lignes = lu.lignes.map(l => {
      const g = suivi[l.no + '|' + l.produit];
      return [c.po, l.no, l.produit, l.desc, l.qte, l.unite, l.sansPrix ? '' : l.montant.toFixed(2), l.cmdClient, l.client, l.requisition,
        g && g[0] ? g[0] : (lu.type === 'Retour' ? 'Retour demandé' : 'Commandé'), l.note, g ? g[1] : '', g ? g[2] : '', g ? g[3] : '', g ? g[4] : ''];
    });
    if (lignes.length) shL.getRange(shL.getLastRow() + 1, 1, lignes.length, ENTETE_LIGNES.length).setValues(lignes);
  }

  if (shV && aRetirerV.length) aRetirerV.sort((x, y) => y - x).forEach(n => shV.deleteRow(n));
  if (!ESSAI && !stats.reste && !props.getProperty('rattrapageFini')) props.setProperty('rattrapageFini', Utilities.formatDate(new Date(), 'America/Toronto', 'yyyy-MM-dd'));
  Logger.log('---------------------------------');
  Logger.log((ESSAI ? 'MODE ESSAI — rien écrit. ' : '') + 'PO trouvés : ' + Object.keys(parPO).length + ' | PDF lus : ' + stats.lus +
    ' | ' + (ESSAI ? 'À inscrire : ' : 'Inscrits : ') + stats.inscrits + ' | Nouvelles versions / relus : ' + stats.remplaces +
    ' | Déjà inscrits : ' + stats.deja + ' | À vérifier : ' + stats.aVerifier + (stats.reste ? ' | IL EN RESTE ' + stats.reste + ' : relance « passe ».' : ' | Terminé, tout est à jour.'));
  if (ss) Logger.log('Sheet « ' + NOM_SHEET + ' » — ID à coller dans les Réglages du CRM : ' + ss.getId());
}

// ───────────────────────── Lecture d'un PO ─────────────────────────

/**
 * Texte d'un PO -> { type, fournisseur, dateRequise, lignes[], st, total } ou { erreur }.
 * Une ligne de produit finit toujours par « quantité unité prix montant » (ex. « 30.00 CH 43.00 1,290.00 »).
 * Juste avant : « no de ligne, code produit, description ». Juste après : le no d'article, puis
 * « Tag:119521-NOM DU CLIENT » et « Réquisition:534284- » (souvent sur la ligne suivante).
 * Les lignes sans prix (« 7 OUPS FORGOT ONE! », « 8 TAG: STOCK », « 15 ****add on 21 octobre ») sont des notes :
 * elles vont dans la colonne Note de la ligne de produit qui suit.
 */
function lirePO_(texte) {
  if (!texte) return { erreur: 'PDF illisible' };
  const brut = String(texte).replace(/[  ]/g, ' ').replace(/\r/g, '');
  const num = x => Number(String(x).replace(/,/g, ''));
  const type = /DEMANDE DE RETOUR|Retour de Marchandise|RMA REQUEST/i.test(brut) ? 'Retour' : 'Commande';

  // Fournisseur : après « Acheté à », la première ligne qui n'est pas le titre du document ni vide.
  let fournisseur = '';
  const lignesTxt = brut.split('\n').map(l => l.trim()).filter(Boolean);
  const iA = lignesTxt.findIndex(l => /Achet[ée] [àa]|Purchased from|Vendor/i.test(l));
  if (iA >= 0) {
    const titre = /^(Achet|Exp[ée]di|Bon de COMMANDE|Purchase Order|Retour de Marchandise|DEMANDE DE RETOUR|Original|Revision|Copie|,)/i;
    fournisseur = (lignesTxt.slice(iA + 1, iA + 10).find(l => !titre.test(l) && /[A-Z]{3}/.test(l)) || '').replace(/\s*,\s*$/, '').replace(/\s{2,}.*$/, '').trim().slice(0, 70);
  }

  // Pages : on enlève le pied de page et l'en-tête répétés entre deux pages (« LINDA 2026-08-28 09:02:02 eForm … Prix Montant »).
  let plat = brut.replace(/\s+/g, ' ');
  plat = plat.replace(/\S+ 20\d{2}-\d{2}-\d{2} \d{2}:\d{2}:\d{2} eForm.*?(?:Prix|Price) (?:Montant|Amount)(?=.*(?:Sous-? ?Total|Sub-? ?Total))/gi, ' ');

  const dr = plat.match(/(20\d{2}-\d{2}-\d{2})\s+\d{6}\b/);
  const dateRequise = dr ? dr[1] : '';

  const iDeb = plat.search(/(?:Prix|Price)\s+(?:Montant|Amount)/i);
  const iFin = plat.search(/Sous-?\s?Total|Sub-?\s?Total/i);
  if (iDeb < 0 || iFin < 0 || iFin < iDeb) return { erreur: 'tableau des lignes introuvable' };
  const zone = plat.slice(iDeb, iFin).replace(/^(?:Prix|Price)\s+(?:Montant|Amount)/i, ' ')
    // pied et en-tête d'une page suivante restés au milieu des lignes
    .replace(/\S+\s+20\d{2}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2}\s+eForm.*?Page:\s*\d+\/\d+/g, ' ')
    .replace(/Eclairage Moderne Saran Inc\..*?(?:Prix|Price)\s+(?:Montant|Amount)/g, ' ') + ' ';

  // Chaque « quantité unité prix montant » ferme une ligne de produit.
  const reMontants = /\s(-?\d+(?:\.\d+)?)\s+([A-Z]{1,4})\s+(-?[\d,]*\d\.\d{2,4})\s+(-?[\d,]*\d\.\d{2})(?=\s)/g;
  const fins = [];
  let m;
  while ((m = reMontants.exec(zone))) fins.push({ m, debut: m.index, fin: reMontants.lastIndex });
  // Lignes SANS PRIX (prix à confirmer) : « … 200.00 529570 » = quantité puis no d'article, sans unité ni montant.
  const reSansPrix = /\s(-?\d+\.\d{2})\s+(\d{5,7})(?=\s)/g;
  while ((m = reSansPrix.exec(zone))) {
    const d0 = m.index, f0 = m.index + m[1].length + 1;
    if (fins.some(f => d0 < f.fin && f0 > f.debut - 1)) continue;          // fait partie d'une ligne avec prix
    if (!/\.00$/.test(m[1])) continue;                                    // « 82.20 » : un prix seul, pas une quantité
    fins.push({ m: [m[0], m[1], '', '', '0.00'], debut: d0, fin: f0, sansPrix: true });
  }
  fins.sort((a, b) => a.debut - b.debut);
  if (!fins.length) return { erreur: /CANCEL|ANNUL/i.test(zone) ? 'PO sans montants — c\'est une ANNULATION (ex. « please cancel this back order ») : à traiter à la main pour l\'instant' : 'aucune ligne de produit lue' };

  // Début d'une ligne : « no code » où le code ressemble à un code de produit (pas de minuscules, pas « *** » ni « TAG: »).
  // Code : commence par une majuscule ou un chiffre (un seul « * » permis devant), plus de majuscules/chiffres que de minuscules
  // (« DALSLEDDOWNG4-CC-WHMulti » oui, « Light » non), pas de « : » ni de virgule.
  const estCode = c => /^[*#]?[A-Z0-9][A-Za-z0-9\-\/.+#()*]*$/.test(c) && !/^\d+$/.test(c) &&
    (c.match(/[A-Z0-9]/g) || []).length > (c.match(/[a-z]/g) || []).length;
  const lignes = [], traines = [];
  let prec = 0, curseur = 0, noteSuivante = '';
  for (const f of fins) {
    const seg = zone.slice(curseur, f.debut);
    const cands = [];
    const reDebut = /(?:^|\s)(\d{1,3})\s+(\S+)/g;
    let c;
    while ((c = reDebut.exec(seg))) { const n = Number(c[1]); if (n > prec && n <= prec + 40 && estCode(c[2])) cands.push({ n, i: c.index, code: c[2], apres: reDebut.lastIndex }); }
    // Le vrai début : le DERNIER candidat dont le code contient un chiffre (« 8 D85139A » plutôt que « 7 OUPS ») ;
    // si aucun code n'a de chiffre (« 5 BALLAST 8 PIED »), le PREMIER.
    const avecChiffre = cands.filter(x => /\d/.test(x.code));
    const d = avecChiffre.length ? avecChiffre[avecChiffre.length - 1] : cands[0];
    if (!d) return { erreur: 'début de la ligne introuvable avant « ' + f.m[0].trim() + ' »' };
    // Avant le début : la fin de la ligne précédente (no d'article, Tag, Réquisition, suite de description), puis les notes numérotées.
    const avant = seg.slice(0, d.i);
    const iNote = avant.search(new RegExp('(?:^|\\s)(?:' + Array.from({ length: Math.max(0, d.n - prec - 1) }, (_, k) => prec + 1 + k).join('|') + ')\\s'));
    traines.push(d.n - prec > 1 && iNote >= 0 ? avant.slice(0, iNote) : avant);
    const note = d.n - prec > 1 && iNote >= 0 ? avant.slice(iNote).trim() : '';
    lignes.push({ no: d.n, produit: d.code, desc: seg.slice(d.apres).trim().slice(0, 120), qte: num(f.m[1]), unite: f.m[2], montant: num(f.m[4]),
      cmdClient: '', client: '', requisition: '', note: note.slice(0, 150), sansPrix: !!f.sansPrix });
    prec = d.n; curseur = f.fin;
  }
  traines.push(zone.slice(curseur));
  // La traîne qui suit chaque ligne (traines[k+1]) : son Tag et sa Réquisition.
  lignes.forEach((l, k) => {
    const t = traines[k + 1] || '';
    const tag = t.match(/Tag\s*:\s*(\d{4,})\s*-?\s*(.*?)(?=\s+R[ée]quisition|\s*$)/i);
    const req = t.match(/R[ée]quisition\s*:\s*(\d+)/i);
    if (tag) { l.cmdClient = tag[1]; l.client = tag[2].trim().slice(0, 80); }
    if (req) l.requisition = req[1];
    const reste = t.replace(/Tag\s*:\s*\d{4,}.*?(?=\s+R[ée]quisition|\s*$)/i, ' ').replace(/R[ée]quisition\s*:\s*\d*-?/i, ' ').replace(/^\s*\d{5,7}\b/, ' ').replace(/\s+/g, ' ').trim();
    if (reste) l.desc = (l.desc + ' ' + reste).slice(0, 160);
    // v9 : Tag écrit dans la description ou la note, nom AVANT le numéro (« TAG: GESTION ADG - PAQUET SAILLANT 119722 »).
    if (!l.cmdClient) {
      const t2 = (l.desc + ' ' + l.note).match(/TAG\s*:?\s*([^\d:]{3,80}?)[\s\-–]*\b(\d{6})\b/i);
      if (t2 && !/^\s*(stock|stk|showroom)\b/i.test(t2[1])) { l.cmdClient = t2[2]; l.client = t2[1].replace(/[\s\-–]+$/, '').trim().slice(0, 80); }
      // v9.1 : Tag avec un NOM seulement (« TAG: DUQUETTE », « tag : Lambert + stock ») : on garde le nom ; poEnvoyes_ cherche la commande.
      else if (!t2) {
        const t3 = (l.desc + ' ' + l.note).match(/TAG\s*:\s*([A-Za-zÀ-ÿ'&.\- ]{3,60})/i);
        const nom = t3 ? t3[1].replace(/\+.*$/, '').replace(/\b(stock|stk|showroom|per|par)\b.*$/i, '').replace(/[\s\-–.]+$/, '').trim() : '';
        if (nom.length >= 3) l.tagNom = nom;
      }
    }
    const orig = (l.desc + ' ' + l.note).match(/PO\s?#?\s?(\d{6})\b/i);           // retour : le PO d'origine
    if (orig) l.note = ('PO d\'origine ' + orig[1] + (l.note ? ' — ' + l.note : '')).slice(0, 150);
  });

  // Montants : un sous-total S tel que TPS = 5 % de S, TVQ = 9,975 % de S et S + TPS + TVQ = un total du PO (retour : en négatif).
  const pied = plat.slice(iFin);
  const argent = (pied.match(/-?\d{1,3}(?:,\d{3})*\.\d{2}(?!\d)|-?\d+\.\d{2}(?!\d)/g) || []).map(num);
  const abs = argent.map(Math.abs);
  const present = v => abs.some(a => Math.abs(a - v) < 0.015);
  let st = null, total = null;
  abs.forEach(s => {
    if (s <= 0) return;
    const tps = Math.round(s * 5) / 100, tvq = Math.round(s * 9.975) / 100;
    const t = abs.find(a => Math.abs(a - (s + tps + tvq)) <= 0.02);
    if (present(tps) && present(tvq) && t != null && (st === null || s > st)) { st = s; total = t; }
  });
  if (st === null && !lignes.some(l => !l.sansPrix) && !argent.some(a => a !== 0)) {
    if (/CANCEL|ANNUL/i.test(zone)) return { erreur: 'PO sans montants — c\'est une ANNULATION (ex. « please cancel this back order ») : à traiter à la main pour l\'instant' };
    return { type, fournisseur, dateRequise, lignes, st: null, total: null };   // PO sans aucun prix : prix à confirmer
  }
  if (st === null) return { erreur: 'sous-total et taxes ne balancent pas (transport ? fournisseur hors taxes ?)' };
  const somme = Math.round(lignes.reduce((a, l) => a + l.montant, 0) * 100) / 100;
  if (Math.abs(Math.abs(somme) - st) > 0.02) return { erreur: 'la somme des lignes (' + somme.toFixed(2) + ') ne donne pas le sous-total (' + st.toFixed(2) + ')' };
  if (somme < 0) { st = -st; total = -total; }

  return { type, fournisseur, dateRequise, lignes, st, total };
}

// v9.1 : les commandes clients en cours dont le nom contient ce nom (mot entier). Une seule → on la prend.
let INDEX_CLIENTS_ = null;
function relierParNom_(shL, nom) {
  if (!shL) return null;
  if (!INDEX_CLIENTS_) {
    INDEX_CLIENTS_ = {};
    if (shL.getLastRow() > 1) shL.getRange(2, 1, shL.getLastRow() - 1, 11).getDisplayValues().forEach(r => {
      const cmd = r[7], client = r[8], statut = r[10] || 'Commandé';
      if (!cmd || !client || /stock|showroom|\bstk\b/i.test(client)) return;
      const c = INDEX_CLIENTS_[cmd] || (INDEX_CLIENTS_[cmd] = { client, ouvert: false });
      if (!/^(Reçu|Livré|Annulé)/.test(statut)) c.ouvert = true;
    });
  }
  const n = normaliser_(nom).replace(/[^a-z0-9& ]/g, ' ').trim();
  if (n.length < 3) return null;
  const re = new RegExp('(^|[^a-z0-9])' + n.replace(/\s+/g, '[^a-z0-9]+') + '($|[^a-z0-9])');
  const trouve = Object.keys(INDEX_CLIENTS_).filter(cmd => INDEX_CLIENTS_[cmd].ouvert && re.test(normaliser_(INDEX_CLIENTS_[cmd].client)));
  return trouve.length === 1 ? { cmd: trouve[0], client: INDEX_CLIENTS_[trouve[0]].client } : null;
}

// ───────────────────────── Outils ─────────────────────────

// Texte d'un PDF : Google Drive le convertit en Doc le temps de le lire, puis le Doc est jeté.
// (Repris du CRM de Saran Électrique, où il a fait ses preuves.)
function lirePdf_(blob) {
  const jeton = ScriptApp.getOAuthToken();
  const limite = '-------lamemoire' + Date.now();
  const meta = JSON.stringify({ name: 'tmp lecture PO', mimeType: 'application/vnd.google-apps.document' });
  const corps = Utilities.newBlob(
    '--' + limite + '\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n' + meta +
    '\r\n--' + limite + '\r\nContent-Type: application/pdf\r\n\r\n').getBytes()
    .concat(blob.getBytes())
    .concat(Utilities.newBlob('\r\n--' + limite + '--').getBytes());
  let rep;
  for (let essai = 0; essai < 3; essai++) {                 // Google répond parfois « Internal Error » : on réessaie
    rep = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&ocrLanguage=fr', {
      method: 'post', contentType: 'multipart/related; boundary=' + limite, payload: corps,
      headers: { Authorization: 'Bearer ' + jeton }, muteHttpExceptions: true });
    if (rep.getResponseCode() < 500) break;
    Utilities.sleep(2000);
  }
  if (rep.getResponseCode() !== 200) { Logger.log('Lecture PDF impossible : ' + rep.getContentText().slice(0, 200)); return ''; }
  const id = JSON.parse(rep.getContentText()).id;
  const texte = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + id + '/export?mimeType=text/plain',
    { headers: { Authorization: 'Bearer ' + jeton } }).getContentText();
  DriveApp.getFileById(id).setTrashed(true);
  return texte;
}

function chercher_(requete) {
  const tout = [];
  for (let debut = 0; debut < 3000; debut += 100) {
    const lot = GmailApp.search(requete + ' -in:spam -in:trash', debut, 100);
    tout.push(...lot);
    if (lot.length < 100) break;
  }
  return tout;
}

// Le Sheet « La mémoire — Éclairage » dans le dossier : trouvé, sinon créé (en vrai seulement).
function sheet_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('sheetMemoire');
  if (id) { try { return SpreadsheetApp.openById(id); } catch (e) {} }
  const dossier = DriveApp.getFolderById(DOSSIER_MEMOIRE);
  const it = dossier.getFilesByName(NOM_SHEET);
  let ss;
  if (it.hasNext()) ss = SpreadsheetApp.openById(it.next().getId());
  else {
    ss = SpreadsheetApp.create(NOM_SHEET);
    DriveApp.getFileById(ss.getId()).moveTo(dossier);
    Logger.log('Sheet créé : ' + NOM_SHEET);
  }
  props.setProperty('sheetMemoire', ss.getId());
  return ss;
}

function onglet_(ss, nom, entete) {
  let sh = ss.getSheetByName(nom);
  if (!sh) {
    const vide = ss.getSheets().length === 1 && ss.getSheets()[0].getLastRow() === 0 && /^(Feuille|Sheet)\s?1$/i.test(ss.getSheets()[0].getName());
    sh = vide ? ss.getSheets()[0].setName(nom) : ss.insertSheet(nom);
    sh.getRange(1, 1, 1, entete.length).setValues([entete]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('A:A').setNumberFormat('@');                      // numéros gardés comme texte
  } else if (sh.getLastColumn() < entete.length) {
    sh.getRange(1, 1, 1, entete.length).setValues([entete]).setFontWeight('bold');   // colonnes ajoutées dans une nouvelle version
  }
  return sh;
}

function retirerLignes_(sh, po) {
  if (sh.getLastRow() < 2) return;
  const v = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getDisplayValues().map(r => r[0]);
  const i0 = v.indexOf(po), i1 = v.lastIndexOf(po);
  if (i0 < 0) return;
  if (v.slice(i0, i1 + 1).every(x => x === po)) { sh.deleteRows(i0 + 2, i1 - i0 + 1); return; }   // lignes d'un PO = un bloc
  for (let i = v.length - 1; i >= 0; i--) if (v[i] === po) sh.deleteRow(i + 2);
}

function sousDossier_(parent, nom) {
  const it = parent.getFoldersByName(nom);
  return it.hasNext() ? it.next() : parent.createFolder(nom);
}

// À lancer UNE fois : « passe » roule toute seule chaque heure (un seul déclencheur pour tout le projet).
function installer() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'passe').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('passe').timeBased().everyHours(1).create();
  Logger.log('Installé : « passe » roule chaque heure. Déclencheurs dans ce projet : ' + ScriptApp.getProjectTriggers().length + '.');
}