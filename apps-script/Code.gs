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
const LECTEUR = 5;   // version de la lecture des PO : un PO lu avec une version plus vieille est relu une fois

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
        if (!pos.length && stats.pdf < 15 && [...domaines].some(d => domaine.toLowerCase().endsWith(d.split('.').slice(-2).join('.')))) {
          const pdf = m.getAttachments({ includeInlineImages: false }).find(a => /\.pdf$/i.test(a.getName()));
          if (pdf) { stats.pdf++; const t = lirePdf_(pdf.copyBlob()); pos = [...new Set(String(t).match(/\b40[4-9]\d{3}\b/g) || [])].filter(p => poConnus.has(p)); }
        }
        vus.add(m.getId());
        if (!pos.length) continue;
        const texte = sujet + ' ' + corps.split(/\n(?:Le .{5,80}a écrit|On .{5,80}wrote|-{2,}\s*Original|De ?:|From:|_{8,})/)[0];
        const type = typeCourriel_(texte, m);
        if (type === 'Relevé') continue;                                   // relevé de compte : pas une nouvelle sur une commande
        const eta = ((texte.match(/\bETA\b[^.\n]{0,70}/i) || texte.match(/\b(?:back ?order\w*|en souffrance|disponib\w*|dispo\b|exp[ée]di(?:era|er|é)\b|will ship|ship(?:ping)? date|ships?\b|pr[êe]te? pour la cueillette|ready for pick ?up|jusqu'?(?:au|à)\b|available\b|delay\w*|retard\w*)[^.\n]{0,70}/i) || [''])[0]).replace(/\s+/g, ' ').trim().slice(0, 90);
        const extrait = corps.split(/\n(?:Le .{5,80}a écrit|On .{5,80}wrote|-{2,}\s*Original|De ?:|From:|_{8,})/)[0].replace(/\[[^\]]*\]|<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 300);
        const de = String(m.getFrom()).replace(/<.*>/, '').replace(/"/g, '').trim() || m.getFrom();
        pos.forEach(po => {
          const ligne = [m.getId() + '|' + po, po, Utilities.formatDate(m.getDate(), 'America/Toronto', 'yyyy-MM-dd'), de, (String(m.getFrom()).match(/[\w.+-]+@[\w.-]+/) || [''])[0].toLowerCase(),
            sujet.slice(0, 120), type, eta, extrait, 'https://mail.google.com/mail/?authuser=' + MOI + '#all/' + fil.getId(), Utilities.formatDate(new Date(), 'America/Toronto', 'yyyy-MM-dd HH:mm')];
          neuves.push(ligne); stats.relies++;
          Logger.log((COURRIELS_ESSAI ? '[essai] ' : '') + 'PO ' + po + ' | ' + ligne[2] + ' | ' + de + ' | ' + type + (eta ? ' | « ' + eta + ' »' : '') + ' | ' + sujet.slice(0, 70));
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