/**
 * Au Temps du Pain — réception des commandes et e-mails automatiques.
 * Script lié au Google Sheet Commandes, privé (Extensions > Apps Script). Voir apps-script/README.md.
 * Les réglages (onglet "contact") sont lus dans le Sheet catalogue, public en lecture pour le site.
 *
 * - doPost         : reçoit la commande du site, l'enregistre, envoie la confirmation au client
 *                    (HTML + texte) et l'alerte "Nouvelle commande" au gestionnaire (ADMIN_EMAIL).
 * - onEditPaiement : déclencheur installable "Lors de la modification". Prévient le client quand
 *                    Paiement passe à Payé / Remboursé, ou Statut commande à Prête / Remise / Annulée.
 *                    Chaque notification n'est envoyée qu'une fois (colonne "Notifications").
 * - doGet          : version navigateur d'un e-mail, pour la page commande.html du site (jeton signé).
 * - Envoi : via Brevo si la propriété BREVO_API_KEY existe, sinon via Gmail (MailApp).
 */

const BRAND = "Au Temps du Pain";
const SITE_URL = "https://autantdupain.eremy.fr";
const ADMIN_PREFIX = "[ATDP · Commande]";
const SHEET_ORDERS = "commandes";
const SHEET_CONTACT = "contact";
// Sheet catalogue (Produits, Farines, contact), public en lecture pour le site.
const CATALOGUE_SPREADSHEET_ID = "1n7WGoY4z922pAjxOYDWSpoD3d9VRYOhBBu9G-tnf-FY";
const HEADERS = ["Référence", "Date", "Prénom", "Nom", "E-mail", "Téléphone", "Jour de remise", "Créneau",
  "Produits", "Total (€)", "Mode de paiement", "Paiement", "Transaction", "Statut commande", "Mail paiement envoyé",
  "Notifications"];

// Clé interne -> libellé écrit dans la colonne Notifications.
const EVENTS = { commande: "commande", paye: "payé", rembourse: "remboursé", prete: "prête", remise: "remise", annulee: "annulée" };
const PAY_EVENTS = ["paye", "rembourse"];
const STATUS_EVENTS = ["prete", "remise", "annulee"];

/* ================= Points d'entrée ================= */

function doPost(e) {
  const o = JSON.parse(e.postData.contents);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let sh, row, paiement = "En attente";
  try {
    sh = ordersSheet_();
    if (findRow_(sh, o.ref)) return json_({ ok: true, duplicate: true });

    let transaction = "";
    if (o.pay === "paypal" && o.paypalOrderId) {
      const v = verifyPayPal_(o.paypalOrderId, o.total, o.test);
      paiement = v.ok ? "Payé" : "À vérifier";
      transaction = o.paypalOrderId + (o.test ? " (sandbox)" : "");
    }
    const rec = {
      "Référence": o.ref, "Date": new Date(), "Prénom": o.prenom, "Nom": o.nom, "E-mail": o.email, "Téléphone": o.tel,
      "Jour de remise": o.jourLabel, "Créneau": o.creneau,
      "Produits": o.items.map(i => i.qty + " × " + i.nom + " — " + eur_(i.qty * i.prix)).join("\n"),
      "Total (€)": o.total, "Mode de paiement": payLabel_(o.pay), "Paiement": paiement, "Transaction": transaction,
      "Statut commande": "Confirmée"
    };
    const head = headers_(sh);
    row = sh.getLastRow() + 1;
    // Texte brut, sinon Sheets convertit "lundi 6 octobre 2026" en date.
    sh.getRange(row, head.indexOf("Jour de remise") + 1).setNumberFormat("@");
    sh.getRange(row, 1, 1, head.length).setValues([head.map(h => h in rec ? rec[h] : "")]);
    SpreadsheetApp.flush();
    token_(o.ref); // crée VIEW_SECRET sous verrou au premier appel
  } finally {
    lock.releaseLock();
  }

  const d = rowData_(sh, row);
  safely_("confirmation client", () => {
    sendClientMail_(d, "commande");
    markSent_(sh, row, "commande");
    // Paiement PayPal vérifié : la confirmation l'indique déjà, pas de second e-mail "Payé".
    if (paiement === "Payé") markSent_(sh, row, "paye");
  });
  safely_("alerte gestionnaire", () => sendAdminMail_(d, sh, row));
  return json_({ ok: true, paiement: paiement });
}

/** Déclencheur installable "Lors de la modification" (garder ce nom : le déclencheur y est rattaché). */
function onEditPaiement(e) {
  const sh = e.range.getSheet();
  if (sh.getName() !== SHEET_ORDERS) return;
  ensureHeaders_(sh);
  const head = headers_(sh);
  const cPay = head.indexOf("Paiement") + 1, cStat = head.indexOf("Statut commande") + 1;
  const c0 = e.range.getColumn(), c1 = c0 + e.range.getNumColumns() - 1;
  const watchPay = cPay >= c0 && cPay <= c1, watchStat = cStat >= c0 && cStat <= c1;
  if (!watchPay && !watchStat) return;

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    // Lit les cellules plutôt que e.value (absent en cas de copier-coller ou de plage multiple).
    const r0 = Math.max(2, e.range.getRow()), r1 = e.range.getLastRow();
    for (let r = r0; r <= r1; r++) {
      const d = rowData_(sh, r);
      if (!d["Référence"]) continue;
      const due = [];
      if (watchPay) due.push(norm_(d["Paiement"]));
      if (watchStat) due.push(norm_(d["Statut commande"]));
      due.filter(k => PAY_EVENTS.includes(k) || STATUS_EVENTS.includes(k)).forEach(k => {
        if (alreadySent_(d, k)) { console.log("Ligne " + r + " : e-mail " + EVENTS[k] + " déjà envoyé"); return; }
        sendClientMail_(d, k);
        markSent_(sh, r, k);
        console.log("Ligne " + r + " : e-mail " + EVENTS[k] + " envoyé");
      });
    }
  } finally {
    lock.releaseLock();
  }
}

/** Rattrapage manuel : envoie les notifications dues et pas encore envoyées, sur toutes les lignes. */
function envoyerPaiementsEnAttente() {
  const sh = ordersSheet_();
  onEditPaiement({ range: sh.getRange(2, 1, Math.max(1, sh.getLastRow() - 1), sh.getLastColumn()) });
}

/** Version navigateur d'un e-mail client : ?ref=CMD-…&type=commande&t=jeton */
function doGet(e) {
  const p = (e && e.parameter) || {};
  const sh = ordersSheet_();
  const row = p.ref ? findRow_(sh, p.ref) : 0;
  if (!row || !p.t || p.t !== token_(p.ref)) return json_({ ok: false });
  const type = EVENTS[p.type] ? p.type : "commande";
  const m = clientMessage_(rowData_(sh, row), type);
  return json_({ ok: true, subject: m.subject, html: renderHtml_(m, { browser: true }) });
}

/* ================= Contenu des e-mails client ================= */

function sendClientMail_(d, type) {
  const m = clientMessage_(d, type);
  m.viewUrl = viewUrl_(d["Référence"], type);
  sendMail_({ to: d["E-mail"], replyTo: shopEmail_(), subject: m.subject, html: renderHtml_(m), text: renderText_(m) });
}

function clientMessage_(d, type) {
  const ref = d["Référence"], prenom = d["Prénom"], total = eur_(d["Total (€)"]);
  const mode = d["Mode de paiement"], pay = norm_(d["Paiement"]), stat = norm_(d["Statut commande"]);
  const paid = pay === "paye";
  const remise = jour_(d["Jour de remise"]) + (d["Créneau"] ? " · " + d["Créneau"] : "");
  const steps = { recue: true, payee: paid, prete: stat === "prete" || stat === "remise", remise: stat === "remise" };
  const adresse = contacts_().ADRESSE_REMISE;
  const deadline = "Besoin de modifier ou d’annuler ? Répondez simplement à cet e-mail jusqu’à l’avant-veille du jour de remise, 18 h — le temps de rafraîchir le levain et de laisser pousser la pâte.";
  const base = { ref: ref, items: items_(d), total: total, cta: { label: "Voir ma commande" } };

  switch (type) {
    case "commande": {
      const intro = paid ? ["Votre paiement " + mode + " est confirmé. Je prépare votre pain pour le jour choisi et vous écris dès que la commande est prête."]
        : mode === "Wero" ? ["Je vous écris dès réception de votre virement Wero, puis quand votre commande sera prête."]
        : mode === "PayPal" ? ["Votre paiement PayPal est en cours de vérification. Je vous écris dès qu’il est confirmé, puis quand votre commande sera prête."]
        : ["Le règlement se fait à la remise, en espèces ou par virement. Je vous écris quand votre commande sera prête."];
      const info = [["Remise en main propre", remise], ["Référence", ref], ["Paiement", paid ? "Payé · " + mode : payPending_(mode)]];
      const wero = contacts_().WERO_ID;
      if (mode === "Wero" && !paid && wero) info.push(["Virement Wero", wero + " · référence " + ref]);
      return Object.assign(base, {
        subject: "Votre commande " + ref + " est confirmée" + (paid ? " et payée" : ""),
        preheader: "Remise " + remise + " · " + total,
        title: "Merci " + prenom + ", votre commande est bien reçue", intro: intro, steps: steps, info: info, note: deadline
      });
    }
    case "paye":
      return Object.assign(base, {
        subject: "Paiement reçu · commande " + ref,
        preheader: total + " reçus · remise " + remise,
        title: "Paiement bien reçu, merci " + prenom,
        intro: ["J’ai bien reçu votre paiement. Je vous écris dès que votre commande est prête."],
        steps: Object.assign(steps, { payee: true }),
        info: [["Montant", total], ["Mode de paiement", mode]].concat(d["Transaction"] ? [["Transaction", d["Transaction"]]] : [])
          .concat([["Remise en main propre", remise], ["Référence", ref]]),
        note: deadline
      });
    case "prete":
      return Object.assign(base, {
        subject: "Votre commande " + ref + " est prête",
        preheader: "À récupérer " + remise,
        title: "Votre pain est prêt, " + prenom + " !",
        intro: ["Votre commande vous attend " + remise + "."].concat(paid ? [] : ["Montant à régler à la remise : " + total + " (espèces ou virement)."]),
        steps: Object.assign(steps, { prete: true }),
        info: [["Remise en main propre", remise]].concat(adresse ? [["Adresse", adresse]] : []).concat([["Référence", ref]]),
        note: "Un empêchement ? Répondez à cet e-mail pour convenir d’un autre moment."
      });
    case "remise":
      return Object.assign(base, {
        subject: "Merci pour votre commande " + ref,
        preheader: "Bonne dégustation !",
        title: "Merci " + prenom + ", bonne dégustation !",
        intro: ["Votre commande vous a bien été remise. Pour garder la mie moelleuse, conservez le pain dans un sac en tissu ou un linge, jamais au réfrigérateur."],
        steps: { recue: true, payee: true, prete: true, remise: true },
        info: [["Référence", ref], ["Remise", remise]],
        cta: { label: "Commander à nouveau", url: SITE_URL },
        note: "Un avis, une suggestion ? Répondez à cet e-mail, je lis tout."
      });
    case "annulee":
      return Object.assign(base, {
        subject: "Commande " + ref + " annulée",
        preheader: paid ? "Remboursement de " + total + " sous 14 jours" : "Aucun paiement à prévoir",
        title: "Votre commande a été annulée",
        intro: ["Bonjour " + prenom + ", votre commande " + ref + " prévue " + remise + " est annulée."].concat(
          paid ? ["Comme vous l’aviez déjà réglée, vous serez remboursé(e) de " + total + " par le même moyen de paiement, sous 14 jours au plus. Un e-mail vous le confirmera."]
            : pay === "rembourse" ? ["Le remboursement a déjà été effectué."]
            : ["Aucun paiement n’avait été enregistré : vous n’avez rien à faire."]),
        steps: null,
        info: [["Référence", ref], ["Montant", total]],
        cta: { label: "Voir le site", url: SITE_URL },
        note: "Une question ? Répondez simplement à cet e-mail."
      });
    case "rembourse":
      return Object.assign(base, {
        subject: "Remboursement effectué · commande " + ref,
        preheader: total + " remboursés",
        title: "Votre remboursement est en route",
        intro: ["Bonjour " + prenom + ", j’ai procédé au remboursement de votre commande " + ref + " " + refundWay_(mode) + ".",
          "Selon votre banque, le crédit peut mettre quelques jours à apparaître."],
        steps: null,
        info: [["Montant remboursé", total], ["Moyen", mode], ["Référence", ref]],
        cta: { label: "Voir le site", url: SITE_URL },
        note: "Une question ? Répondez simplement à cet e-mail."
      });
  }
  throw new Error("Type d'e-mail inconnu : " + type);
}

function payPending_(mode) {
  return mode === "Wero" ? "Wero — en attente de réception"
    : mode === "PayPal" ? "PayPal — en cours de vérification"
    : "À régler en main propre, à la remise";
}
function refundWay_(mode) {
  return mode === "PayPal" ? "sur votre compte PayPal" : mode === "Wero" ? "par virement Wero" : "selon le moyen convenu ensemble";
}

/* ================= Alerte gestionnaire ================= */

function sendAdminMail_(d, sh, row) {
  const ref = d["Référence"], total = eur_(d["Total (€)"]), client = d["Prénom"] + " " + d["Nom"];
  const paid = norm_(d["Paiement"]) === "paye";
  const payBadge = paid ? "Payé · " + d["Mode de paiement"]
    : norm_(d["Paiement"]) === "a verifier" ? "À vérifier · " + d["Mode de paiement"]
    : d["Mode de paiement"] === "En main propre" ? "À régler à la remise" : "En attente · " + d["Mode de paiement"];
  const remise = jour_(d["Jour de remise"]) + (d["Créneau"] ? " · " + d["Créneau"] : "");
  const placed = d["Date"] instanceof Date ? Utilities.formatDate(d["Date"], Session.getScriptTimeZone(), "dd/MM/yyyy 'à' HH'h'mm") : String(d["Date"]);
  const sheetUrl = SpreadsheetApp.getActive().getUrl() + "#gid=" + sh.getSheetId() + "&range=A" + row;
  const replyUrl = "mailto:" + d["E-mail"] + "?subject=" + encodeURIComponent("Votre commande " + ref);
  const items = items_(d);
  const subject = ADMIN_PREFIX + " " + ref + " · " + total + " · " + remise + " · " + payBadge;

  const rows = [["Client", "<b>" + esc_(client) + "</b>"],
    ["E-mail", '<a href="mailto:' + esc_(d["E-mail"]) + '" style="color:' + C.accent + '">' + esc_(d["E-mail"]) + "</a>"]]
    .concat(d["Téléphone"] ? [["Téléphone", '<a href="tel:' + esc_(String(d["Téléphone"]).replace(/[^\d+]/g, "")) + '" style="color:' + C.accent + '">' + esc_(d["Téléphone"]) + "</a>"]] : [])
    .concat([["Remise", esc_(remise)], ["Passée le", esc_(placed)]])
    .concat(d["Transaction"] ? [["Transaction", esc_(d["Transaction"])]] : []);

  const html = doc_(subject, client + " · " + items.length + " article" + (items.length > 1 ? "s" : ""),
    '<tr><td style="background:' + C.ink + ';border-radius:8px 8px 0 0;padding:16px 24px;font:700 11px/1.4 ' + F.sans + ';letter-spacing:.14em;text-transform:uppercase;color:' + C.wheat300 + '">Nouvelle commande · ' + BRAND + "</td></tr>"
    + '<tr><td class="card" style="background:' + C.card + ";border:1px solid " + C.line + ';border-top:0;border-radius:0 0 8px 8px;padding:28px">'
    + '<table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr>'
    + '<td style="font:400 24px/1.2 ' + F.display + ";color:" + C.ink + '">' + esc_(ref) + "</td>"
    + '<td align="right" style="font:700 22px/1.2 ' + F.sans + ";color:" + C.ink + '">' + esc_(total) + "</td></tr></table>"
    + '<p style="margin:12px 0 20px">' + badge_(payBadge, paid ? C.sunken : C.soft, paid ? C.ink2 : C.accentDark) + " " + badge_("Remise " + remise, C.sunken, C.ink2) + "</p>"
    + '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="font:400 14px/1.6 ' + F.sans + ";color:" + C.ink + ';margin-bottom:20px">'
    + rows.map(r => '<tr><td width="110" valign="top" style="color:' + C.muted + ';padding:3px 0">' + r[0] + '</td><td style="padding:3px 0">' + r[1] + "</td></tr>").join("")
    + "</table>"
    + '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="font:400 14px/1.5 ' + F.sans + ";color:" + C.ink + ";background:" + C.bg + ';border-radius:6px">'
    + items.map((it, i) => '<tr><td style="padding:' + (i ? 4 : 12) + "px 16px " + (i === items.length - 1 ? 12 : 4) + 'px">' + esc_(it.label)
      + (it.sub ? ' <span style="color:' + C.muted + '">(' + esc_(it.sub) + ")</span>" : "") + '</td><td align="right" style="padding:' + (i ? 4 : 12) + "px 16px " + (i === items.length - 1 ? 12 : 4) + 'px;white-space:nowrap">' + esc_(it.price) + "</td></tr>").join("")
    + "</table>"
    + '<table role="presentation" cellpadding="0" cellspacing="0" style="margin-top:24px"><tr>'
    + '<td style="background:' + C.ink + ';border-radius:2px"><a href="' + esc_(sheetUrl) + '" style="display:inline-block;padding:12px 20px;font:700 14px ' + F.sans + ";color:" + C.card + ';text-decoration:none">Ouvrir le Google Sheet</a></td><td width="10"></td>'
    + '<td style="border:1px solid ' + C.ink + ';border-radius:2px"><a href="' + esc_(replyUrl) + '" style="display:inline-block;padding:11px 18px;font:700 14px ' + F.sans + ";color:" + C.ink + ';text-decoration:none">Répondre au client</a></td>'
    + "</tr></table></td></tr>");

  const text = "Nouvelle commande " + ref + " — " + total + "\n" + payBadge + "\n\n"
    + "Client : " + client + "\nE-mail : " + d["E-mail"] + (d["Téléphone"] ? "\nTéléphone : " + d["Téléphone"] : "")
    + "\nRemise : " + remise + "\nPassée le : " + placed + (d["Transaction"] ? "\nTransaction : " + d["Transaction"] : "") + "\n\n"
    + items.map(it => it.label + (it.sub ? " (" + it.sub + ")" : "") + (it.price ? " — " + it.price : "")).join("\n")
    + "\n\nGoogle Sheet : " + sheetUrl;

  // Répondre à cette alerte écrit directement au client.
  sendMail_({ to: adminEmail_(), replyTo: d["E-mail"], subject: subject, html: html, text: text });
}

/* ================= Gabarit HTML ================= */

const C = { bg: "#F4EDE0", card: "#FBF7EF", line: "#D8CBB3", soft: "#F1E3C3", sunken: "#E9DFCC", ink: "#2B1E16", ink2: "#4A3526",
  muted: "#7A6452", accent: "#8B4A24", accentDark: "#733B1B", wheat: "#B38A3E", wheat300: "#D9BC82", ring: "#BFAE91" };
const F = { display: "'Libre Caslon Display',Georgia,'Times New Roman',serif", serif: "'Libre Caslon Text',Georgia,serif",
  sans: "Karla,'Helvetica Neue',Arial,sans-serif" };

function renderHtml_(m, opts) {
  const browser = opts && opts.browser;
  const ctaUrl = m.cta.url || (browser ? "" : m.viewUrl);
  return doc_(m.subject, m.preheader,
    (!browser && m.viewUrl ? '<tr><td align="center" style="padding:0 0 12px;font:400 12px/1.5 ' + F.sans + ";color:" + C.muted + '">Cet e-mail s’affiche mal ? <a href="' + esc_(m.viewUrl) + '" style="color:' + C.accent + '">Voir dans le navigateur</a></td></tr>' : "")
    + '<tr><td align="center" style="padding:24px 24px 20px"><a href="' + SITE_URL + '" style="text-decoration:none">'
    + '<div style="font:400 30px/1 ' + F.display + ";color:" + C.ink + '">' + BRAND + "</div>"
    + '<div style="font:700 11px/1 ' + F.sans + ";letter-spacing:.16em;text-transform:uppercase;color:" + C.wheat + ';margin-top:8px">Pain au levain naturel</div></a></td></tr>'
    + '<tr><td class="card" style="background:' + C.card + ";border:1px solid " + C.line + ';border-radius:8px;padding:32px 32px 28px">'
    + (m.steps ? steps_(m.steps) : "")
    + '<h1 class="h1" style="margin:0 0 12px;font:400 26px/1.2 ' + F.display + ";color:" + C.ink + '">' + esc_(m.title) + "</h1>"
    + m.intro.map(p => '<p style="margin:0 0 16px;font:400 15px/1.6 ' + F.sans + ";color:" + C.ink2 + '">' + esc_(p) + "</p>").join("")
    + '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:' + C.soft + ';border-radius:6px;margin:8px 0 24px"><tr><td style="padding:16px 18px;font:400 14px/1.7 ' + F.sans + ";color:" + C.ink + '">'
    + m.info.map(r => "<b>" + esc_(r[0]) + "</b> · " + esc_(r[1])).join("<br>") + "</td></tr></table>"
    + itemsHtml_(m.items, m.total)
    + (ctaUrl ? '<table role="presentation" cellpadding="0" cellspacing="0" style="margin:28px 0 8px"><tr><td style="background:' + C.accent + ';border-radius:2px"><a href="' + esc_(ctaUrl) + '" style="display:inline-block;padding:13px 24px;font:700 15px ' + F.sans + ";color:" + C.card + ';text-decoration:none">' + esc_(m.cta.label) + "</a></td></tr></table>" : "")
    + (m.note ? '<p style="margin:24px 0 0;padding-top:20px;border-top:1px solid ' + C.line + ";font:italic 400 14px/1.6 " + F.serif + ";color:" + C.ink2 + '">' + esc_(m.note) + "</p>" : "")
    + "</td></tr>"
    + '<tr><td align="center" style="padding:24px 24px 0;font:400 12px/1.7 ' + F.sans + ";color:" + C.muted + '">'
    + BRAND + ' · <a href="mailto:' + esc_(shopEmail_()) + '" style="color:' + C.muted + '">' + esc_(shopEmail_()) + "</a><br>"
    + '<a href="' + SITE_URL + '" style="color:' + C.muted + '">' + SITE_URL.replace("https://", "") + "</a> · CGV et confidentialité sur le site<br>"
    + "Vous recevez cet e-mail car vous avez passé commande sur notre site.</td></tr>",
    browser);
}

function doc_(title, preheader, rows, browser) {
  return '<!doctype html><html lang="fr"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">'
    + '<meta name="color-scheme" content="light"><meta name="supported-color-schemes" content="light">'
    + (browser ? '<base target="_blank">' : "") + "<title>" + esc_(title) + "</title>"
    + '<link href="https://fonts.googleapis.com/css2?family=Karla:wght@400;500;700&family=Libre+Caslon+Display&family=Libre+Caslon+Text:ital,wght@0,400;1,400&display=swap" rel="stylesheet">'
    + "<style>body{margin:0;padding:0}@media (max-width:620px){.card{padding:24px 18px !important}.h1{font-size:22px !important}}</style></head>"
    + '<body style="margin:0;padding:0;background:' + C.bg + '">'
    + '<div style="display:none;max-height:0;overflow:hidden;opacity:0;color:transparent">' + esc_(preheader) + "&#8203;&nbsp;".repeat(40) + "</div>"
    + '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:' + C.bg + '"><tr><td align="center" style="padding:16px 12px 32px">'
    + '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:600px">' + rows + "</table>"
    + "</td></tr></table></body></html>";
}

function steps_(s) {
  const list = [["recue", "Reçue"], ["payee", "Payée"], ["prete", "Prête"], ["remise", "Remise"]];
  return '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:24px"><tr>'
    + list.map((x, i) => {
      const done = s[x[0]];
      const dot = done
        ? '<div style="width:22px;height:22px;line-height:22px;border-radius:50%;background:' + C.accent + ";color:" + C.card + ';margin:0 auto 6px;text-align:center">&#10003;</div>'
        : '<div style="width:22px;height:22px;line-height:19px;border-radius:50%;border:1.5px solid ' + C.ring + ';margin:0 auto 6px;text-align:center;box-sizing:border-box">' + (i + 1) + "</div>";
      return '<td align="center" width="25%" style="font:' + (done ? 700 : 500) + " 11px/1.2 " + F.sans + ";color:" + (done ? C.accent : C.muted) + '">' + dot + x[1] + "</td>";
    }).join("") + "</tr></table>";
}

function itemsHtml_(items, total) {
  if (!items.length) return "";
  const td = 'style="padding:12px 0;border-bottom:1px dashed ' + C.line + '"';
  return '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="font:400 14px/1.5 ' + F.sans + ";color:" + C.ink + '">'
    + '<tr><td colspan="2" style="padding:0 0 8px;font:700 11px/1 ' + F.sans + ";letter-spacing:.12em;text-transform:uppercase;color:" + C.muted + ";border-bottom:1px solid " + C.line + '">Votre commande</td></tr>'
    + items.map(it => "<tr><td " + td + ">" + esc_(it.label) + (it.sub ? '<br><span style="color:' + C.muted + ';font-size:13px">' + esc_(it.sub) + "</span>" : "")
      + '</td><td align="right" valign="top" ' + td.replace('style="', 'style="white-space:nowrap;') + ">" + esc_(it.price) + "</td></tr>").join("")
    + '<tr><td style="padding:14px 0 0;font:700 16px ' + F.sans + '">Total TTC</td><td align="right" style="padding:14px 0 0;font:700 16px ' + F.sans + ';white-space:nowrap">' + esc_(total) + "</td></tr></table>";
}

function badge_(label, bg, fg) {
  return '<span style="display:inline-block;margin:0 4px 6px 0;padding:4px 10px;border-radius:2px;background:' + bg + ";color:" + fg + ";font:700 11px/1.4 " + F.sans + ';letter-spacing:.1em;text-transform:uppercase">' + esc_(label) + "</span>";
}

/** Version texte (multipart) : même contenu que le HTML, pour les messageries sans HTML. */
function renderText_(m) {
  const ctaUrl = m.cta.url || m.viewUrl;
  return [BRAND.toUpperCase(), "", m.title, ""].concat(m.intro.map(p => p + "\n"))
    .concat(m.info.map(r => r[0] + " : " + r[1])).concat([""])
    .concat(m.items.length ? ["VOTRE COMMANDE"].concat(m.items.map(it => it.label + (it.sub ? " (" + it.sub + ")" : "") + (it.price ? " — " + it.price : "")), ["Total TTC : " + m.total, ""]) : [])
    .concat(ctaUrl ? [m.cta.label + " : " + ctaUrl, ""] : [])
    .concat(m.note ? [m.note, ""] : [])
    .concat(["—", BRAND + " · " + shopEmail_(), SITE_URL]).join("\n");
}

/** Lignes "2 × Pain (tranché) — 7,80 €" du Sheet -> { label, sub, price }. */
function items_(d) {
  return String(d["Produits"] || "").split("\n").map(l => l.trim()).filter(Boolean).map(l => {
    const m = l.match(/^(\d+)\s*×\s*(.+?)\s*—\s*([\d\s.,]+\s*€)$/);
    if (!m) return { label: l, sub: "", price: "" };
    const n = m[2].match(/^(.*?)\s*\((.+)\)$/);
    return { label: m[1] + " × " + (n ? n[1] : m[2]), sub: n ? n[2] : "", price: m[3] };
  });
}

/* ================= Envoi ================= */

/** Brevo (API transactionnelle) si BREVO_API_KEY est défini, sinon Gmail. */
function sendMail_(o) {
  const p = PropertiesService.getScriptProperties();
  const key = p.getProperty("BREVO_API_KEY");
  if (!key) {
    MailApp.sendEmail({ to: o.to, replyTo: o.replyTo, name: BRAND, subject: o.subject, body: o.text, htmlBody: o.html });
    return;
  }
  const res = UrlFetchApp.fetch("https://api.brevo.com/v3/smtp/email", {
    method: "post", contentType: "application/json", muteHttpExceptions: true,
    headers: { "api-key": key, accept: "application/json" },
    payload: JSON.stringify({
      sender: { name: BRAND, email: p.getProperty("BREVO_SENDER") || "commandes@autantdupain.eremy.fr" },
      to: [{ email: o.to }], replyTo: { email: o.replyTo },
      subject: o.subject, htmlContent: o.html, textContent: o.text
    })
  });
  if (res.getResponseCode() >= 300) throw new Error("Brevo " + res.getResponseCode() + " : " + res.getContentText());
}

function safely_(label, fn) {
  try { fn(); } catch (err) { console.error("Échec " + label + " — " + err.message); }
}

/* ================= Suivi des notifications ================= */

function alreadySent_(d, key) {
  if (key === "paye" && d["Mail paiement envoyé"]) return true;
  return String(d["Notifications"] || "").split("·").map(s => norm_(s.trim().split(" ")[0])).includes(key);
}

function markSent_(sh, row, key) {
  const head = headers_(sh);
  const cNotif = head.indexOf("Notifications") + 1;
  const cell = sh.getRange(row, cNotif);
  const stamp = EVENTS[key] + " " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM HH:mm");
  cell.setValue(cell.getValue() ? cell.getValue() + " · " + stamp : stamp);
  if (key === "paye") sh.getRange(row, head.indexOf("Mail paiement envoyé") + 1).setValue(new Date());
}

/* ================= Version navigateur ================= */

function viewUrl_(ref, type) {
  return SITE_URL + "/commande.html?ref=" + encodeURIComponent(ref) + "&type=" + type + "&t=" + token_(ref);
}

/** Jeton HMAC de la référence : empêche d'afficher une commande en devinant sa référence. */
function token_(ref) {
  const p = PropertiesService.getScriptProperties();
  let secret = p.getProperty("VIEW_SECRET");
  if (!secret) { secret = Utilities.getUuid() + Utilities.getUuid(); p.setProperty("VIEW_SECRET", secret); }
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(String(ref), secret)).replace(/=+$/, "").slice(0, 24);
}

/* ================= PayPal : vérification serveur de la transaction ================= */
function verifyPayPal_(orderId, total, test) {
  const p = PropertiesService.getScriptProperties();
  const id = p.getProperty(test ? "PAYPAL_SANDBOX_CLIENT_ID" : "PAYPAL_CLIENT_ID");
  const secret = p.getProperty(test ? "PAYPAL_SANDBOX_SECRET" : "PAYPAL_SECRET");
  if (!id || !secret) return { ok: false };
  const api = test ? "https://api-m.sandbox.paypal.com" : "https://api-m.paypal.com";
  const tok = JSON.parse(UrlFetchApp.fetch(api + "/v1/oauth2/token", {
    method: "post", payload: { grant_type: "client_credentials" },
    headers: { Authorization: "Basic " + Utilities.base64Encode(id + ":" + secret) }
  }).getContentText()).access_token;
  const order = JSON.parse(UrlFetchApp.fetch(api + "/v2/checkout/orders/" + orderId, {
    headers: { Authorization: "Bearer " + tok }, muteHttpExceptions: true
  }).getContentText());
  const amount = order.purchase_units && order.purchase_units[0].amount;
  return { ok: order.status === "COMPLETED" && amount && Math.abs(Number(amount.value) - Number(total)) < 0.01 };
}

/* ================= Utilitaires ================= */
function ordersSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_ORDERS);
  if (!sh) {
    sh = ss.insertSheet(SHEET_ORDERS);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    const col = h => HEADERS.indexOf(h) + 1;
    sh.getRange(2, col("Paiement"), 999).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(["En attente", "À vérifier", "Payé", "Remboursé"]).build());
    sh.getRange(2, col("Statut commande"), 999).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(["Confirmée", "Prête", "Remise", "Annulée"]).build());
  }
  ensureHeaders_(sh);
  return sh;
}

/**
 * Ajoute la colonne "Notifications" aux Sheets existants. Les lignes déjà présentes sont marquées
 * avec leurs statuts actuels, pour ne pas renvoyer d'e-mails sur d'anciennes commandes.
 */
function ensureHeaders_(sh) {
  const head = headers_(sh);
  if (head.includes("Notifications")) return;
  const c = head.length + 1;
  sh.getRange(1, c).setValue("Notifications");
  const n = sh.getLastRow() - 1;
  if (n < 1) return;
  const cPay = head.indexOf("Paiement"), cStat = head.indexOf("Statut commande");
  const values = sh.getRange(2, 1, n, head.length).getValues().map(r => {
    const keys = ["commande", norm_(r[cPay]), norm_(r[cStat])].filter(k => EVENTS[k]);
    return [keys.map(k => EVENTS[k]).join(" · ") + " (avant mise à jour)"];
  });
  sh.getRange(2, c, n, 1).setValues(values);
}

function headers_(sh) {
  return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h).trim());
}
function findRow_(sh, ref) {
  const refs = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
  for (let i = 1; i < refs.length; i++) if (refs[i][0] === ref) return i + 1;
  return 0;
}
/** Ligne -> objet { en-tête: valeur }, retrouvé par nom de colonne. */
function rowData_(sh, row) {
  const head = headers_(sh);
  const v = sh.getRange(row, 1, 1, head.length).getValues()[0];
  return Object.fromEntries(head.map((h, i) => [h, v[i]]));
}

let CONTACTS_;
/** Onglet "contact" du Sheet catalogue, lu une seule fois par exécution. */
function contacts_() {
  if (CONTACTS_) return CONTACTS_;
  CONTACTS_ = {};
  const sh = SpreadsheetApp.openById(CATALOGUE_SPREADSHEET_ID).getSheetByName(SHEET_CONTACT);
  if (sh) sh.getDataRange().getValues().forEach(r => {
    const k = String(r[0]).trim().toUpperCase(), v = String(r[1]).trim();
    if (k && v) CONTACTS_[k] = v;
  });
  return CONTACTS_;
}
function shopEmail_() { return contacts_().CONTACT_EMAIL || Session.getEffectiveUser().getEmail(); }
function adminEmail_() { return contacts_().ADMIN_EMAIL || shopEmail_(); }

/** Jour de remise lisible : texte tel qu'envoyé par le site, ou date convertie par Sheets (anciennes lignes). */
function jour_(v) {
  if (!(v instanceof Date)) return String(v);
  const J = ["dimanche", "lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi"];
  const M = ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre"];
  return J[v.getDay()] + " " + v.getDate() + " " + M[v.getMonth()] + " " + v.getFullYear();
}
/** "Payé" / "Annulée" / "À vérifier" -> "paye" / "annulee" / "a verifier". */
function norm_(s) { return String(s || "").trim().toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, ""); }
function esc_(s) { return String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;"); }
function payLabel_(p) { return p === "paypal" ? "PayPal" : p === "wero" ? "Wero" : "En main propre"; }
function eur_(n) { return Number(n).toFixed(2).replace(".", ",") + " €"; }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
