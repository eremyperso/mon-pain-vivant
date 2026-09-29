/**
 * Au Temps du Pain — réception des commandes et e-mails automatiques.
 * Script lié au Google Sheet (Extensions > Apps Script). Voir apps-script/README.md.
 *
 * - doPost      : reçoit la commande du site, l'enregistre dans l'onglet "commandes",
 *                 envoie la validation de commande (client + boutique) et, si le paiement
 *                 PayPal est vérifié, la confirmation de paiement.
 * - Envoi des e-mails : via Brevo si la propriété BREVO_API_KEY existe, sinon via Gmail (MailApp).
 * - onEditPaiement (déclencheur "Lors de la modification") : quand vous passez la colonne
 *                 "Paiement" à "Payé" (Wero, main propre), envoie la confirmation de paiement.
 */

const BRAND = "Au Temps du Pain";
const SHEET_ORDERS = "commandes";
const SHEET_CONTACT = "contact";
const HEADERS = ["Référence", "Date", "Prénom", "Nom", "E-mail", "Téléphone", "Jour de remise", "Créneau",
  "Produits", "Total (€)", "Mode de paiement", "Paiement", "Transaction", "Statut commande", "Mail paiement envoyé"];
const COL = Object.fromEntries(HEADERS.map((h, i) => [h, i + 1]));

function doPost(e) {
  const o = JSON.parse(e.postData.contents);
  const sh = ordersSheet_();
  if (findRow_(sh, o.ref)) return json_({ ok: true, duplicate: true });

  let paiement = "En attente", transaction = "";
  if (o.pay === "paypal" && o.paypalOrderId) {
    const v = verifyPayPal_(o.paypalOrderId, o.total, o.test);
    paiement = v.ok ? "Payé" : "À vérifier";
    transaction = o.paypalOrderId + (o.test ? " (sandbox)" : "");
  }
  const lignes = o.items.map(i => i.qty + " × " + i.nom + " — " + eur_(i.qty * i.prix)).join("\n");
  sh.getRange(sh.getLastRow() + 1, COL["Jour de remise"]).setNumberFormat("@");
  sh.appendRow([o.ref, new Date(), o.prenom, o.nom, o.email, o.tel, o.jourLabel, o.creneau,
    lignes, o.total, payLabel_(o.pay), paiement, transaction, "Confirmée", ""]);
  const row = sh.getLastRow();

  sendOrderMail_(rowData_(sh, row));
  if (paiement === "Payé") sendPaymentMail_(sh, row);
  return json_({ ok: true, paiement: paiement });
}

/** Déclencheur installable "Lors de la modification" sur cette fonction. */
function onEditPaiement(e) {
  const sh = e.range.getSheet();
  if (sh.getName() !== SHEET_ORDERS) return;
  // Colonnes retrouvées par leur en-tête : résiste à un déplacement de colonnes.
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const cPay = head.indexOf("Paiement") + 1, cSent = head.indexOf("Mail paiement envoyé") + 1;
  if (!cPay || !cSent) { console.warn("En-têtes Paiement / Mail paiement envoyé introuvables"); return; }
  const c0 = e.range.getColumn(), c1 = c0 + e.range.getNumColumns() - 1;
  if (cPay < c0 || cPay > c1) return;
  // Lit les cellules plutôt que e.value (absent en cas de copier-coller ou de plage multiple).
  const r0 = Math.max(2, e.range.getRow()), r1 = e.range.getLastRow();
  for (let r = r0; r <= r1; r++) {
    const pay = String(sh.getRange(r, cPay).getValue()).trim().toLowerCase();
    if (pay !== "payé" && pay !== "paye") continue;
    if (sh.getRange(r, cSent).getValue()) { console.log("Ligne " + r + " : mail de paiement déjà envoyé, ignoré"); continue; }
    try {
      sendPaymentMail_(sh, r);
      console.log("Ligne " + r + " : mail de paiement envoyé");
    } catch (err) {
      console.error("Ligne " + r + " : échec d'envoi — " + err.message);
      throw err;
    }
  }
}

/** Rattrapage manuel : envoie les confirmations de paiement en attente (Paiement = Payé, pas encore envoyées). */
function envoyerPaiementsEnAttente() {
  const sh = ordersSheet_();
  onEditPaiement({ range: sh.getRange(2, 1, Math.max(1, sh.getLastRow() - 1), sh.getLastColumn()) });
}

/* ---------- E-mails ---------- */
function sendOrderMail_(d) {
  const shop = contact_("CONTACT_EMAIL") || Session.getEffectiveUser().getEmail();
  const pay = d["Paiement"] === "Payé" ? "Payé (" + d["Mode de paiement"] + ")"
    : d["Mode de paiement"] === "Wero" ? "Wero — en cours de vérification"
    : d["Mode de paiement"] === "PayPal" ? "PayPal — en cours de vérification"
    : "À régler en main propre, à la remise";
  const body = "Bonjour " + d["Prénom"] + ",\n\nNous avons bien reçu votre commande " + d["Référence"] + ".\n\n"
    + d["Produits"] + "\n\nTotal : " + eur_(d["Total (€)"]) + "\n"
    + "Remise en main propre : " + jour_(d["Jour de remise"]) + (d["Créneau"] ? " · " + d["Créneau"] : "") + "\n"
    + "Paiement : " + pay + "\n"
    + "Statut de la commande : " + d["Statut commande"] + "\n\n"
    + "Modifiable ou annulable jusqu’à l’avant-veille du jour de remise, 18 h (le temps de rafraîchir le levain et de laisser pousser la pâte), en répondant à cet e-mail.\n\n" + BRAND;
  sendMail_(d["E-mail"], shop, BRAND + " — Commande " + d["Référence"] + " validée", body);
}

function sendPaymentMail_(sh, row) {
  const d = rowData_(sh, row);
  const shop = contact_("CONTACT_EMAIL") || Session.getEffectiveUser().getEmail();
  const body = "Bonjour " + d["Prénom"] + ",\n\nVotre paiement pour la commande " + d["Référence"] + " est confirmé.\n\n"
    + "Montant : " + eur_(d["Total (€)"]) + "\n"
    + "Mode de paiement : " + d["Mode de paiement"] + "\n"
    + (d["Transaction"] ? "Transaction : " + d["Transaction"] + "\n" : "")
    + "Statut de la commande : " + d["Statut commande"] + " — payée\n"
    + "Remise en main propre : " + jour_(d["Jour de remise"]) + "\n\nÀ bientôt,\n" + BRAND;
  sendMail_(d["E-mail"], shop, BRAND + " — Paiement confirmé · " + d["Référence"], body);
  const cSent = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].indexOf("Mail paiement envoyé") + 1;
  sh.getRange(row, cSent || COL["Mail paiement envoyé"]).setValue(new Date());
}

/** Brevo (API transactionnelle) si BREVO_API_KEY est défini, sinon Gmail. */
function sendMail_(to, shop, subject, text) {
  const p = PropertiesService.getScriptProperties();
  const key = p.getProperty("BREVO_API_KEY");
  if (!key) {
    MailApp.sendEmail({ to: to, bcc: shop, replyTo: shop, name: BRAND, subject: subject, body: text });
    return;
  }
  const res = UrlFetchApp.fetch("https://api.brevo.com/v3/smtp/email", {
    method: "post", contentType: "application/json", muteHttpExceptions: true,
    headers: { "api-key": key, accept: "application/json" },
    payload: JSON.stringify({
      sender: { name: BRAND, email: p.getProperty("BREVO_SENDER") || "commandes@autantdupain.eremy.fr" },
      to: [{ email: to }], bcc: shop ? [{ email: shop }] : undefined, replyTo: { email: shop },
      subject: subject, textContent: text
    })
  });
  if (res.getResponseCode() >= 300) throw new Error("Brevo " + res.getResponseCode() + " : " + res.getContentText());
}

/* ---------- PayPal : vérification serveur de la transaction ---------- */
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

/* ---------- Utilitaires ---------- */
function ordersSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_ORDERS);
  if (!sh) {
    sh = ss.insertSheet(SHEET_ORDERS);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(2, COL["Paiement"], 999).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(["En attente", "À vérifier", "Payé", "Remboursé"]).build());
    sh.getRange(2, COL["Statut commande"], 999).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(["Confirmée", "Prête", "Remise", "Annulée"]).build());
  }
  return sh;
}
function findRow_(sh, ref) {
  const refs = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
  for (let i = 1; i < refs.length; i++) if (refs[i][0] === ref) return i + 1;
  return 0;
}
function rowData_(sh, row) {
  const v = sh.getRange(row, 1, 1, HEADERS.length).getValues()[0];
  return Object.fromEntries(HEADERS.map((h, i) => [h, v[i]]));
}
function contact_(key) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_CONTACT);
  if (!sh) return "";
  const r = sh.getDataRange().getValues().find(r => String(r[0]).trim().toUpperCase() === key);
  return r ? String(r[1]).trim() : "";
}
/** Jour de remise lisible : texte tel qu'envoyé par le site, ou date convertie par Sheets (anciennes lignes). */
function jour_(v) {
  if (!(v instanceof Date)) return String(v);
  const J = ["dimanche", "lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi"];
  const M = ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre"];
  return J[v.getDay()] + " " + v.getDate() + " " + M[v.getMonth()] + " " + v.getFullYear();
}
function payLabel_(p) { return p === "paypal" ? "PayPal" : p === "wero" ? "Wero" : "En main propre"; }
function eur_(n) { return Number(n).toFixed(2).replace(".", ",") + " €"; }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
