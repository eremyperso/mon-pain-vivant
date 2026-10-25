/**
 * Au Temps du Pain — lettre d'information : campagnes programmées depuis le Google Sheet Commandes.
 * Même projet Apps Script que Code.gs (constantes et utilitaires partagés). Voir apps-script/README.md.
 *
 * - Onglet "abonnes"   : une ligne par e-mail, alimentée par la case « Recevoir la liste des pains de la
 *                        semaine » du site. Statut Inscrit / Désinscrit. Les colonnes de profil (commandes,
 *                        pains achetés) sont recalculées depuis l'onglet commandes avant chaque envoi.
 * - Onglet "campagnes" : l'agenda. Une ligne = un e-mail (contenu, liste visée, date d'envoi). Seules les
 *                        lignes "Programmée" partent, à l'heure prévue (déclencheur toutes les 15 min).
 * - Onglet "envois"    : journal campagne × destinataire ; personne ne reçoit deux fois la même campagne.
 * - Désinscription     : lien signé dans chaque e-mail (desinscription.html) + en-tête List-Unsubscribe (Brevo).
 */

const SHEET_SUBS = "abonnes";
const SHEET_CAMPAIGNS = "campagnes";
const SHEET_SENDS = "envois";
const SUB_HEADERS = ["E-mail", "Prénom", "Statut", "Source", "Inscrit le", "Désinscrit le", "Commandes", "Dernière commande", "Pains achetés"];
const CAMP_HEADERS = ["ID", "Statut", "Envoi prévu", "Liste", "Objet", "Pré-en-tête", "Titre", "Texte", "Produit vedette",
  "Texte vedette", "Produits associés", "Bouton", "Lien bouton", "Note", "Envoyés", "Rapport"];
const SEND_HEADERS = ["Campagne", "E-mail", "Date", "Résultat"];
const CAMP_STATUS = ["Brouillon", "Programmée", "En cours", "Envoyée", "Annulée"];
const SENT = "envoyé";
const NEWS_TIME_BUDGET_MS = 4.5 * 60 * 1000; // Apps Script coupe une exécution à 6 min
const NEWS_DAILY_RESERVE = 20;               // e-mails gardés chaque jour pour les commandes (quota Gmail)
const NEWS_BREVO_MAX_DAY = 250;              // offre gratuite Brevo : 300/jour, commandes comprises

/* ================= Menu du Google Sheet ================= */

function onOpen() {
  SpreadsheetApp.getUi().createMenu("Lettre")
    .addItem("Envoyer un test de la campagne sélectionnée", "envoyerTestCampagne")
    .addItem("Actualiser les abonnés", "actualiserAbonnes")
    .addItem("Créer l’agenda « Le Norvégien »", "creerAgendaNorvegien")
    .addSeparator()
    .addItem("Activer l’envoi automatique", "installerEnvoiAutomatique")
    .addToUi();
}

/** Crée (une seule fois) le déclencheur qui envoie les campagnes programmées, toutes les 15 minutes. */
function installerEnvoiAutomatique() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === "envoyerCampagnesProgrammees")
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("envoyerCampagnesProgrammees").timeBased().everyMinutes(15).create();
  campaignsSheet_(); subsSheet_(); sendsSheet_();
  ui_("Envoi automatique activé : les campagnes « Programmée » partent dans le quart d’heure qui suit leur date d’envoi.");
}

/* ================= Abonnés ================= */

/** Inscription (case cochée à la commande) ou réinscription d'un abonné désinscrit. Appelé sous verrou. */
function subscribe_(email, prenom, source) {
  email = String(email || "").trim().toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return;
  const sh = subsSheet_(), head = headers_(sh), col = h => head.indexOf(h) + 1;
  const row = findEmail_(sh, email);
  if (!row) {
    const rec = { "E-mail": email, "Prénom": prenom || "", "Statut": "Inscrit", "Source": source, "Inscrit le": new Date() };
    sh.appendRow(head.map(h => h in rec ? rec[h] : ""));
    return;
  }
  if (norm_(sh.getRange(row, col("Statut")).getValue()) === "inscrit") return;
  sh.getRange(row, col("Statut")).setValue("Inscrit");
  sh.getRange(row, col("Source")).setValue(source);
  sh.getRange(row, col("Inscrit le")).setValue(new Date());
  sh.getRange(row, col("Désinscrit le")).setValue("");
}

function unsubscribe_(email) {
  const sh = subsSheet_(), head = headers_(sh), row = findEmail_(sh, email);
  if (!row) return;
  sh.getRange(row, head.indexOf("Statut") + 1).setValue("Désinscrit");
  sh.getRange(row, head.indexOf("Désinscrit le") + 1).setValue(new Date());
}

/** Lien signé des e-mails : ?action=desinscription|reinscription&e=…&t=… (GET depuis desinscription.html, POST one-click). */
function newsletterAction_(p) {
  const email = String(p.e || "").trim().toLowerCase();
  if (!email || p.t !== unsubToken_(email)) return { ok: false };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (p.action === "reinscription") subscribe_(email, "", "réinscription (lien de l’e-mail)");
    else unsubscribe_(email);
  } finally {
    lock.releaseLock();
  }
  return { ok: true, action: p.action, email: email };
}

/** Menu : recalcule Commandes / Dernière commande / Pains achetés de chaque abonné depuis l'onglet commandes. */
function actualiserAbonnes() {
  const subs = subscribers_();
  ui_(subs.length + " abonné(s), dont " + subs.filter(s => s.inscrit).length + " inscrit(s). Profils mis à jour depuis les commandes.");
}

/** Abonnés avec leur profil d'achat ; réécrit au passage les colonnes de profil du Sheet. */
function subscribers_() {
  const sh = subsSheet_(), n = sh.getLastRow() - 1;
  if (n < 1) return [];
  const head = headers_(sh), c = h => head.indexOf(h);
  const rows = sh.getRange(2, 1, n, head.length).getValues();
  const profiles = profiles_();
  const subs = rows.map(r => {
    const email = String(r[c("E-mail")]).trim().toLowerCase();
    return { email: email, prenom: String(r[c("Prénom")]).trim(), inscrit: norm_(r[c("Statut")]) === "inscrit", profile: profiles[email] || emptyProfile_() };
  });
  // Seules les colonnes de profil sont réécrites : une inscription ou désinscription simultanée n'est pas écrasée.
  const cN = c("Commandes");
  sh.getRange(2, cN + 1, n, 3).setValues(subs.map(s => [s.profile.n, s.profile.last || "", s.profile.products.join(", ")]));
  return subs.filter(s => s.email);
}

function emptyProfile_() { return { n: 0, last: null, products: [], lastProducts: [], tags: [] }; }

/** Historique d'achat par e-mail, tiré de l'onglet commandes (commandes annulées exclues). */
function profiles_() {
  const sh = ordersSheet_(), n = sh.getLastRow() - 1, out = {};
  if (n < 1) return out;
  const head = headers_(sh), c = h => head.indexOf(h);
  sh.getRange(2, 1, n, head.length).getValues().forEach(r => {
    const email = String(r[c("E-mail")]).trim().toLowerCase();
    if (!email || norm_(r[c("Statut commande")]) === "annulee") return;
    const p = out[email] || (out[email] = emptyProfile_());
    const names = items_({ "Produits": r[c("Produits")] }).map(it => it.label.replace(/^\d+\s*×\s*/, ""));
    const date = r[c("Date")] instanceof Date ? r[c("Date")] : null;
    p.n++;
    if (date && (!p.last || date > p.last)) { p.last = date; p.lastProducts = names; }
    names.forEach(x => { if (!p.products.includes(x)) p.products.push(x); });
  });
  Object.values(out).forEach(p => {
    p.tags = [...new Set(p.products.map(product_).filter(Boolean).flatMap(x => x.tags))];
  });
  return out;
}

/* ================= Listes de diffusion ================= */

/**
 * Colonne Liste : critères séparés par « ; », tous obligatoires (ET). Plusieurs valeurs : « | » (OU).
 *   tous · clients · commandes:3 · tag:Graines|Seigle · a-commande:Le norvégien · pas-commande:Le norvégien
 *   inactifs:30 (aucune commande depuis 30 jours) · pas-recu:NORV-01-avant-premiere (n'a pas reçu cette campagne)
 *   email:marie@exemple.fr|paul@exemple.fr (ces abonnés seulement) · test (uniquement l'adresse gestionnaire)
 */
function segment_(spec) {
  const tests = String(spec || "").split(";").map(s => s.trim()).filter(Boolean).map(s => {
    const i = s.indexOf(":"), key = norm_(i < 0 ? s : s.slice(0, i)), arg = i < 0 ? "" : s.slice(i + 1).trim();
    const list = arg.split("|").map(x => x.trim()).filter(Boolean);
    const bought = pr => pr.products.some(x => list.map(slug_).includes(slug_(x)));
    switch (key) {
      case "tous": return () => true;
      case "clients": return pr => pr.n > 0;
      case "commandes": return pr => pr.n >= (Number(arg) || 1);
      case "tag": return pr => pr.tags.some(t => list.map(norm_).includes(norm_(t)));
      case "a-commande": return bought;
      case "pas-commande": return pr => !bought(pr);
      case "inactifs": return pr => !pr.last || (Date.now() - pr.last) / 864e5 >= (Number(arg) || 30);
      case "email": return (pr, email) => list.map(x => x.toLowerCase()).includes(email);
      case "pas-recu": {
        const got = new Set(sends_().filter(r => list.includes(String(r[0])) && r[3] === SENT).map(r => String(r[1])));
        return (pr, email) => !got.has(email);
      }
    }
    throw new Error("critère de liste inconnu : « " + s + " »");
  });
  if (!tests.length) throw new Error("colonne Liste vide");
  return (pr, email) => tests.every(t => t(pr, email));
}

function recipients_(camp) {
  if (norm_(camp["Liste"]) === "test") return [{ email: adminEmail_(), prenom: "", profile: emptyProfile_() }];
  const match = segment_(camp["Liste"]);
  return subscribers_().filter(s => s.inscrit && match(s.profile, s.email));
}

/* ================= Envoi programmé ================= */

/** Déclencheur : envoie les campagnes "Programmée" dont l'heure est passée et reprend celles "En cours". */
function envoyerCampagnesProgrammees() {
  // Verrou distinct de celui des commandes : un envoi de plusieurs minutes ne bloque pas le site.
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(1000)) return;
  try {
    const start = Date.now(), sh = campaignsSheet_(), now = new Date();
    for (let r = 2; r <= sh.getLastRow(); r++) {
      const d = rowData_(sh, r), stat = norm_(d["Statut"]);
      const due = stat === "en cours" || (stat === "programmee" && d["Envoi prévu"] instanceof Date && d["Envoi prévu"] <= now);
      if (due && !sendCampaign_(sh, r, start)) break; // temps ou quota épuisé : reprise au prochain passage
    }
  } finally {
    lock.releaseLock();
  }
}

/** Envoie une campagne aux destinataires qui ne l'ont pas encore reçue. Renvoie false s'il faut s'arrêter là. */
function sendCampaign_(sh, row, start) {
  const camp = rowData_(sh, row), id = String(camp["ID"]).trim();
  const set = (h, v) => sh.getRange(row, headers_(sh).indexOf(h) + 1).setValue(v);
  const stamp = () => Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM HH:mm");
  let todo;
  try {
    if (!id) throw new Error("ID manquant");
    renderCampaign_(camp, {}); // valide produits et contenu avant tout envoi
    const done = sentSet_(id);
    todo = recipients_(camp).filter(s => !done.has(s.email));
  } catch (err) {
    set("Statut", "Brouillon");
    set("Rapport", "Erreur " + stamp() + " : " + err.message + ". Corrigez puis repassez en « Programmée ».");
    safely_("alerte campagne", () => sendMail_({ to: adminEmail_(), replyTo: shopEmail_(), subject: ADMIN_PREFIX + " Campagne " + id + " non envoyée",
      text: "La campagne " + id + " n’a pas été envoyée : " + err.message + "\n\n" + SpreadsheetApp.getActive().getUrl(),
      html: "<p>La campagne <b>" + esc_(id) + "</b> n’a pas été envoyée : " + esc_(err.message) + "</p><p><a href=\"" + SpreadsheetApp.getActive().getUrl() + "\">Ouvrir le Google Sheet</a></p>" }));
    return true;
  }

  set("Statut", "En cours");
  SpreadsheetApp.flush();
  let budget = dailyBudget_(), failed = 0, streak = 0, stop = "";
  for (const s of todo) {
    if (Date.now() - start > NEWS_TIME_BUDGET_MS) { stop = "temps d’exécution atteint"; break; }
    if (budget <= 0) { stop = "quota d’e-mails du jour atteint"; break; }
    try {
      sendCampaignTo_(camp, s);
      logSend_(id, s.email, SENT);
      budget--; streak = 0;
    } catch (err) {
      logSend_(id, s.email, "échec : " + err.message);
      failed++;
      if (++streak >= 3) { stop = "3 échecs de suite (" + err.message + ")"; break; }
    }
  }
  const total = sentSet_(id).size;
  set("Envoyés", total);
  if (stop) {
    set("Rapport", "En cours " + stamp() + " : " + total + " envoyé(s), pause (" + stop + "), reprise automatique.");
    return false;
  }
  set("Statut", "Envoyée");
  set("Rapport", "Terminée " + stamp() + " : " + total + " envoyé(s)" + (failed ? ", " + failed + " échec(s) (voir onglet envois)" : "") + ".");
  return true;
}

/** prefix : "[TEST] " pour un envoi de test, par le même chemin que l'envoi réel (en-têtes compris). */
function sendCampaignTo_(camp, s, prefix) {
  const id = String(camp["ID"]).trim();
  const m = renderCampaign_(camp, varsFor_(s), { unsub: unsubUrl_(s.email), view: campaignViewUrl_(id) });
  // En-têtes de désinscription en un clic (Gmail, Yahoo) : pris en compte par Brevo, ignorés par Gmail/MailApp.
  const hook = contacts_().ORDER_WEBHOOK_URL;
  const unsub = ["<mailto:" + shopEmail_() + "?subject=" + encodeURIComponent("Désinscription " + s.email) + ">"];
  if (hook) unsub.unshift("<" + hook + "?action=desinscription&e=" + encodeURIComponent(s.email) + "&t=" + unsubToken_(s.email) + ">");
  sendMail_({ to: s.email, replyTo: shopEmail_(), subject: (prefix || "") + m.subject, html: m.html, text: m.text, tags: ["lettre", id],
    headers: Object.assign({ "List-Unsubscribe": unsub.join(", ") }, hook ? { "List-Unsubscribe-Post": "List-Unsubscribe=One-Click" } : {}) });
}

/** E-mails encore envoyables aujourd'hui, en gardant de la marge pour les e-mails de commande. */
function dailyBudget_() {
  const p = PropertiesService.getScriptProperties();
  if (!p.getProperty("BREVO_API_KEY")) return MailApp.getRemainingDailyQuota() - NEWS_DAILY_RESERVE;
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd");
  const sentToday = sends_().filter(r => r[3] === SENT && r[2] instanceof Date
    && Utilities.formatDate(r[2], Session.getScriptTimeZone(), "yyyyMMdd") === today).length;
  return (Number(p.getProperty("NEWSLETTER_MAX_JOUR")) || NEWS_BREVO_MAX_DAY) - sentToday;
}

/* ================= Test depuis le menu ================= */

/** Envoie la campagne de la ligne sélectionnée à l'adresse gestionnaire, personnalisée avec le profil du premier destinataire. */
function envoyerTestCampagne() {
  const sh = SpreadsheetApp.getActiveSheet();
  if (sh.getName() !== SHEET_CAMPAIGNS || sh.getActiveRange().getRow() < 2) return ui_("Sélectionnez une ligne de l’onglet « " + SHEET_CAMPAIGNS + " ».");
  const camp = rowData_(sh, sh.getActiveRange().getRow()), id = String(camp["ID"]).trim();
  try {
    const list = recipients_(camp);
    const sample = list[0] || { email: adminEmail_(), prenom: "", profile: emptyProfile_() };
    sendCampaignTo_(camp, Object.assign({}, sample, { email: adminEmail_() }), "[TEST] ");
    ui_("Test envoyé à " + adminEmail_() + ".\n\nListe « " + camp["Liste"] + " » : " + list.length + " destinataire(s) aujourd’hui."
      + (list[0] ? "\nAperçu personnalisé comme pour " + (sample.prenom || sample.email) + "." : ""));
  } catch (err) {
    ui_("Campagne " + id + " : " + err.message);
  }
}

/* ================= Contenu ================= */

/** Variables utilisables dans les textes : {{prenom}}, {{prenom|à vous}} (valeur par défaut après « | »)… */
function varsFor_(s) {
  const r = nextRemise_(new Date()), pr = s.profile || emptyProfile_();
  const pain = pr.lastProducts.concat(pr.products).map(product_).find(p => p && !/epicerie/.test(norm_(p.categorie)));
  return {
    prenom: s.prenom, bonjour: "Bonjour" + (s.prenom ? " " + s.prenom : ""),
    dernier_pain: pain ? pain.nom : "", nb_commandes: pr.n || "",
    prochaine_remise: jour_(r.remise).replace(/ \d{4}$/, ""),
    date_limite: jour_(r.limite).replace(/ \d{4}$/, "") + ", 18 h"
  };
}

/** Prochain lundi ou jeudi encore commandable (jusqu'à l'avant-veille 18 h), comme sur le site. */
function nextRemise_(now) {
  for (let i = 0; i < 14; i++) {
    const day = new Date(now.getFullYear(), now.getMonth(), now.getDate() + i);
    const limite = new Date(day.getFullYear(), day.getMonth(), day.getDate() - 2, 18);
    if ((day.getDay() === 1 || day.getDay() === 4) && limite > now) return { remise: day, limite: limite };
  }
}

function fill_(s, vars) {
  return String(s == null ? "" : s).replace(/\{\{\s*([a-z_]+)\s*(?:\|([^}]*))?\}\}/gi, (m, k, fb) => {
    const v = vars[k.toLowerCase()];
    return v !== undefined && v !== "" ? String(v) : (fb || "").trim();
  });
}

/** Ligne de l'onglet campagnes -> { subject, html, text }. Lève une erreur si un produit est introuvable ou épuisé. */
function renderCampaign_(camp, vars, links) {
  links = links || {};
  const id = String(camp["ID"]).trim();
  const utm = u => u + (u.includes("?") ? "&" : "?") + "utm_source=lettre&utm_medium=email&utm_campaign=" + encodeURIComponent(id);
  const productUrl = p => utm(SITE_URL + "/?produit=" + p.slug);
  const heroName = String(camp["Produit vedette"] || "").trim();
  const hero = heroName ? product_(heroName) : null;
  if (heroName && !hero) throw new Error("produit vedette « " + heroName + " » absent de l’onglet Produits");
  if (hero && hero.stock === "out") throw new Error("produit vedette « " + hero.nom + " » épuisé (stock = out)");
  const names = String(camp["Produits associés"] || "").split(/[;\n]/).map(x => x.trim()).filter(Boolean);
  const missing = names.filter(x => !product_(x));
  if (missing.length) throw new Error("produit(s) associé(s) absent(s) de l’onglet Produits : " + missing.join(", "));
  const extras = names.map(product_).filter(p => p.stock !== "out" && p !== hero);

  vars = Object.assign({ produit: hero ? hero.nom : "", prix: hero ? eur_(hero.prix) : "" }, vars);
  const t = s => fill_(s, vars);
  const paras = s => t(s).split(/\r?\n/).map(x => x.trim()).filter(Boolean);
  const subject = t(camp["Objet"]);
  if (!subject) throw new Error("Objet vide");
  const m = {
    subject: subject, preheader: t(camp["Pré-en-tête"]), title: t(camp["Titre"]), intro: paras(camp["Texte"]),
    hero: hero, heroText: paras(camp["Texte vedette"] || (hero ? hero.description : "")), extras: extras,
    cta: { label: t(camp["Bouton"]) || (hero ? "Commander " + hero.nom : "Voir les pains de la semaine"),
      url: String(camp["Lien bouton"] || "").trim() ? utm(String(camp["Lien bouton"]).trim()) : hero ? productUrl(hero) : utm(SITE_URL) },
    remise: "Prochaine remise " + vars.prochaine_remise + " · commande jusqu’au " + vars.date_limite,
    note: t(camp["Note"]), productUrl: productUrl, siteUrl: utm(SITE_URL), links: links
  };
  return { subject: subject, html: campaignHtml_(m), text: campaignText_(m) };
}

function campaignHtml_(m) {
  const P = (s, style) => '<p style="margin:0 0 16px;font:400 15px/1.6 ' + F.sans + ";color:" + C.ink2 + ";" + (style || "") + '">' + esc_(s) + "</p>";
  const button = (label, url) => '<table role="presentation" cellpadding="0" cellspacing="0" style="margin:8px 0 4px"><tr><td style="background:' + C.accent + ';border-radius:2px"><a href="' + esc_(url) + '" style="display:inline-block;padding:13px 24px;font:700 15px ' + F.sans + ";color:" + C.card + ';text-decoration:none">' + esc_(label) + "</a></td></tr></table>";
  const img = (p, w) => p.photo ? '<a href="' + esc_(m.productUrl(p)) + '"><img src="' + esc_(p.photo + (p.photo.includes("googleusercontent.com/d/") ? "=w" + w * 2 : "")) + '" width="' + w + '" alt="' + esc_(p.nom) + '" style="display:block;width:100%;max-width:' + w + 'px;height:auto;border:0;border-radius:6px"></a>' : "";
  const h = m.hero;

  const heroBlock = h ? '<tr><td class="card" style="background:' + C.card + ";border:1px solid " + C.line + ';border-top:0;padding:0 32px 28px">'
    + img(h, 536)
    + '<p style="margin:20px 0 8px">' + (h.nouveau ? badge_("Nouveauté", C.accent, C.card) : "") + h.tags.slice(0, 3).map(x => badge_(x, C.sunken, C.ink2)).join("") + "</p>"
    + '<table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr>'
    + '<td style="font:400 26px/1.2 ' + F.display + ";color:" + C.ink + '"><a href="' + esc_(m.productUrl(h)) + '" style="color:' + C.ink + ';text-decoration:none">' + esc_(h.nom) + "</a></td>"
    + '<td align="right" valign="bottom" style="font:700 20px/1.2 ' + F.sans + ";color:" + C.ink + ';white-space:nowrap">' + esc_(eur_(h.prix)) + (h.poids ? '<br><span style="font:400 13px ' + F.sans + ";color:" + C.muted + '">' + esc_(h.poids) + "</span>" : "") + "</td></tr></table>"
    + '<div style="height:12px"></div>' + m.heroText.map(x => P(x)).join("")
    + button(m.cta.label, m.cta.url) + "</td></tr>" : "";

  const cell = p => '<td class="col" width="50%" valign="top" style="padding:0 8px 16px">' + img(p, 250)
    + '<div style="font:400 18px/1.25 ' + F.display + ";color:" + C.ink + ';margin:10px 0 4px"><a href="' + esc_(m.productUrl(p)) + '" style="color:' + C.ink + ';text-decoration:none">' + esc_(p.nom) + "</a></div>"
    + '<div style="font:400 14px/1.5 ' + F.sans + ";color:" + C.muted + '">' + esc_(eur_(p.prix) + (p.poids ? " · " + p.poids : "")) + ' · <a href="' + esc_(m.productUrl(p)) + '" style="color:' + C.accent + ';font-weight:700">Voir</a></div></td>';
  const pairs = [];
  for (let i = 0; i < m.extras.length; i += 2) pairs.push(m.extras.slice(i, i + 2));
  const extrasBlock = m.extras.length ? '<tr><td class="card" style="background:' + C.card + ";border:1px solid " + C.line + ';border-top:0;padding:24px 24px 8px">'
    + '<div style="padding:0 8px 14px;font:700 11px/1 ' + F.sans + ";letter-spacing:.12em;text-transform:uppercase;color:" + C.muted + '">Aussi dans la fournée</div>'
    + '<table role="presentation" width="100%" cellpadding="0" cellspacing="0">'
    + pairs.map(pr => "<tr>" + cell(pr[0]) + (pr[1] ? cell(pr[1]) : '<td class="col" width="50%"></td>') + "</tr>").join("") + "</table></td></tr>" : "";

  const promise = [["Commandez en ligne", "jusqu’à l’avant-veille, 18 h"], ["Remise en main propre", "le lundi ou le jeudi"], ["100 % levain", "farine, eau, sel, levain"]];

  return doc_(m.subject, m.preheader,
    (m.links.view ? '<tr><td align="center" style="padding:0 0 12px;font:400 12px/1.5 ' + F.sans + ";color:" + C.muted + '">Cet e-mail s’affiche mal ? <a href="' + esc_(m.links.view) + '" style="color:' + C.accent + '">Voir dans le navigateur</a></td></tr>' : "")
    + '<tr><td align="center" style="padding:24px 24px 20px"><a href="' + esc_(m.siteUrl) + '" style="text-decoration:none">'
    + '<div style="font:400 30px/1 ' + F.display + ";color:" + C.ink + '">' + BRAND + "</div>"
    + '<div style="font:700 11px/1 ' + F.sans + ";letter-spacing:.16em;text-transform:uppercase;color:" + C.wheat + ';margin-top:8px">La lettre du fournil</div></a></td></tr>'
    + '<tr><td class="card" style="background:' + C.card + ";border:1px solid " + C.line + ";border-radius:8px 8px 0 0;" + (h ? "border-bottom:0;" : "") + 'padding:32px 32px 12px">'
    + (m.title ? '<h1 class="h1" style="margin:0 0 16px;font:400 28px/1.2 ' + F.display + ";color:" + C.ink + '">' + esc_(m.title) + "</h1>" : "")
    + m.intro.map(x => P(x)).join("")
    + (h ? "" : button(m.cta.label, m.cta.url) + '<div style="height:16px"></div>')
    + "</td></tr>"
    + heroBlock + extrasBlock
    + '<tr><td class="card" style="background:' + C.soft + ";border:1px solid " + C.line + ';border-top:0;border-radius:0 0 8px 8px;padding:20px 24px">'
    + '<table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr>'
    + promise.map(x => '<td class="col" width="33%" valign="top" style="padding:4px 8px;font:400 13px/1.5 ' + F.sans + ";color:" + C.ink2 + '"><b style="color:' + C.ink + '">' + x[0] + "</b><br>" + x[1] + "</td>").join("")
    + "</tr></table>"
    + '<p style="margin:14px 8px 0;font:700 14px/1.5 ' + F.sans + ";color:" + C.accentDark + '">' + esc_(m.remise) + "</p>"
    + (m.note ? '<p style="margin:14px 8px 0;font:italic 400 14px/1.6 ' + F.serif + ";color:" + C.ink2 + '">' + esc_(m.note) + "</p>" : "")
    + "</td></tr>"
    + '<tr><td align="center" style="padding:24px 24px 0;font:400 12px/1.7 ' + F.sans + ";color:" + C.muted + '">'
    + BRAND + ' · <a href="mailto:' + esc_(shopEmail_()) + '" style="color:' + C.muted + '">' + esc_(shopEmail_()) + "</a><br>"
    + '<a href="' + esc_(m.siteUrl) + '" style="color:' + C.muted + '">' + SITE_URL.replace("https://", "") + "</a> · CGV et confidentialité sur le site<br>"
    + "Vous recevez cette lettre car vous avez demandé la liste des pains de la semaine lors d’une commande."
    + (m.links.unsub ? '<br><a href="' + esc_(m.links.unsub) + '" style="color:' + C.muted + '">Se désinscrire</a>' : "")
    + "</td></tr>",
    !m.links.unsub);
}

function campaignText_(m) {
  const h = m.hero;
  return [BRAND.toUpperCase() + " · LA LETTRE DU FOURNIL", ""].concat(m.title ? [m.title, ""] : [])
    .concat(m.intro.map(x => x + "\n"))
    .concat(h ? [(h.nouveau ? "NOUVEAUTÉ · " : "") + h.nom + " — " + eur_(h.prix) + (h.poids ? " · " + h.poids : "")].concat(m.heroText, [""]) : [])
    .concat([m.cta.label + " : " + m.cta.url, ""])
    .concat(m.extras.length ? ["AUSSI DANS LA FOURNÉE"].concat(m.extras.map(p => p.nom + " — " + eur_(p.prix) + " : " + m.productUrl(p)), [""]) : [])
    .concat([m.remise, ""]).concat(m.note ? [m.note, ""] : [])
    .concat(["—", BRAND + " · " + shopEmail_(), SITE_URL])
    .concat(m.links.unsub ? ["Se désinscrire : " + m.links.unsub] : []).join("\n");
}

/** Version navigateur d'une campagne (lettre.html), sans personnalisation : ?action=lettre&id=ID&t=jeton */
function campaignView_(p) {
  const id = String(p.id || "");
  if (!id || p.t !== token_("lettre:" + id)) return { ok: false };
  const sh = campaignsSheet_(), ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues().map(r => String(r[0]).trim());
  const row = ids.indexOf(id) + 1;
  if (row < 2) return { ok: false };
  const m = renderCampaign_(rowData_(sh, row), varsFor_({ prenom: "" }));
  return { ok: true, subject: m.subject, html: m.html };
}

/* ================= Catalogue ================= */

let CATALOGUE_;
/** Onglet Produits du Sheet catalogue (gid 0), lu comme sur le site. */
function catalogue_() {
  if (CATALOGUE_) return CATALOGUE_;
  const ss = SpreadsheetApp.openById(CATALOGUE_SPREADSHEET_ID);
  const sh = ss.getSheets().find(s => s.getSheetId() === 0) || ss.getSheets()[0];
  const v = sh.getDataRange().getValues(), head = v[0].map(h => String(h).trim().toLowerCase());
  CATALOGUE_ = v.slice(1).map(r => Object.fromEntries(head.map((h, i) => [h, r[i]]))).filter(r => String(r.nom || "").trim()).map(r => ({
    nom: String(r.nom).trim(), slug: slug_(r.nom), categorie: String(r.categorie || ""),
    prix: Number(String(r.prix).replace(",", ".")) || 0, poids: String(r.poids || "").trim(),
    stock: ["in", "low", "out"].includes(String(r.stock).trim()) ? String(r.stock).trim() : "in",
    description: String(r.description || "").trim(), tags: String(r.tags || "").split(/[;,]/).map(x => x.trim()).filter(Boolean),
    photo: driveImg_(String(r.photo || "").trim()), nouveau: isNew_(r.nouveau)
  }));
  return CATALOGUE_;
}
function product_(name) { const s = slug_(name); return catalogue_().find(p => p.slug === s) || null; }

/** Colonne "nouveau" : « oui », ou une date de fin (31/12/2026 ou 2026-12-31) après laquelle le badge disparaît. */
function isNew_(v) {
  if (v instanceof Date) return v >= new Date(new Date().setHours(0, 0, 0, 0));
  const s = norm_(v);
  if (["oui", "x", "1", "true"].includes(s)) return true;
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/) || s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (!m) return false;
  const end = s.includes("/") ? new Date(m[3], m[2] - 1, m[1]) : new Date(m[1], m[2] - 1, m[3]);
  return end >= new Date(new Date().setHours(0, 0, 0, 0));
}
function driveImg_(url) { const m = url.match(/\/d\/([a-zA-Z0-9_-]+)/); return m ? "https://lh3.googleusercontent.com/d/" + m[1] : url; }
function slug_(s) { return norm_(s).replace(/[^a-z0-9]+/g, "-").replace(/(^-|-$)/g, ""); }

/* ================= Agenda « Le Norvégien » ================= */

/** Menu : ajoute les 4 e-mails du lancement du Norvégien à l'onglet campagnes, en Brouillon (rien ne part sans vous). */
function creerAgendaNorvegien() {
  const sh = campaignsSheet_(), head = headers_(sh);
  const ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues().map(r => String(r[0]).trim());
  const P = "Le norvégien";
  const rows = [
    { "ID": "NORV-01-avant-premiere", "Envoi prévu": new Date(2026, 9, 28, 18), "Liste": "tag:Graines|Seigle",
      "Objet": "En avant-première : un nouveau pain pour vous", "Pré-en-tête": "Le Norvégien : seigle torréfié, petit épeautre, miel et graines.",
      "Titre": "Un nouveau pain sort du four, et je vous le réserve", "Produit vedette": P,
      "Texte": "{{bonjour}},\nVous avez déjà goûté « {{dernier_pain|nos pains aux graines}} » : je voulais que vous soyez parmi les premiers à découvrir le Norvégien, avant de l’annoncer à tout le monde.\nIl est déjà disponible sur le site, pour une remise {{prochaine_remise}}.",
      "Texte vedette": "Farine de meule T80, petit épeautre et seigle T130 torréfié artisanalement, une touche de miel et des graines généreuses. Des notes de noisette et de caramel, une mie dense et parfumée : un pain rustique et gourmand, à partager.",
      "Produits associés": "Mélange 4 graines ; Tourte de seigle ; Confiture de Figues", "Bouton": "Commander le Norvégien",
      "Note": "Une question sur ce pain, une allergie ? Répondez simplement à cet e-mail, je lis tout." },
    { "ID": "NORV-02-lancement", "Envoi prévu": new Date(2026, 10, 1, 9), "Liste": "tous ; pas-commande:" + P + " ; pas-recu:NORV-01-avant-premiere",
      "Objet": "Nouveau au fournil : le Norvégien", "Pré-en-tête": "Seigle torréfié, petit épeautre, miel et graines · remise {{prochaine_remise}}",
      "Titre": "Le Norvégien est arrivé", "Produit vedette": P,
      "Texte": "{{bonjour}},\nUn nouveau pain rejoint la fournée : le Norvégien, un pain de caractère aux saveurs nordiques.\nIl est disponible dès maintenant sur le site, pour une remise en main propre {{prochaine_remise}}.",
      "Produits associés": "Campagne au levain ; Mélange 4 graines ; La miche tradition ; Confiture de Figues", "Bouton": "Découvrir le Norvégien",
      "Note": "Les pains sont cuits à la commande : commandez jusqu’à l’avant-veille, 18 h (samedi pour le lundi, mardi pour le jeudi)." },
    { "ID": "NORV-03-relance", "Envoi prévu": new Date(2026, 10, 8, 9), "Liste": "tous ; pas-commande:" + P,
      "Objet": "Avez-vous goûté le Norvégien ?", "Pré-en-tête": "Remise {{prochaine_remise}} · commande jusqu’au {{date_limite}}",
      "Titre": "Le Norvégien vous attend", "Produit vedette": P,
      "Texte": "{{bonjour}},\nLe Norvégien a rejoint la fournée la semaine dernière. Si vous ne l’avez pas encore goûté : seigle torréfié, petit épeautre, une touche de miel et beaucoup de graines.\nPour le recevoir {{prochaine_remise}}, commandez avant le {{date_limite}}.",
      "Produits associés": "La miche tradition ; Campagne au levain", "Bouton": "Commander le Norvégien",
      "Note": "Vous préférez un autre pain ? Toute la fournée de la semaine est sur le site." },
    { "ID": "NORV-04-avis", "Envoi prévu": new Date(2026, 10, 15, 10), "Liste": "a-commande:" + P,
      "Objet": "Alors, ce Norvégien ?", "Pré-en-tête": "Votre avis m’aide à faire évoluer la fournée",
      "Titre": "Merci d’avoir goûté le Norvégien", "Produit vedette": P,
      "Texte": "{{bonjour}},\nVous avez été parmi les premiers à commander le Norvégien : merci !\nPlus de graines, moins de miel, une croûte plus cuite ? Répondez simplement à cet e-mail : vos retours m’aident à faire évoluer la recette.",
      "Texte vedette": "Pour garder la mie moelleuse, conservez-le dans un sac en tissu ou un linge, jamais au réfrigérateur.",
      "Produits associés": "Confiture de Figues", "Bouton": "Le commander à nouveau",
      "Note": "Un avis, une suggestion ? Répondez à cet e-mail, je lis tout." }
  ].filter(r => !ids.includes(r.ID));
  if (!rows.length) return ui_("L’agenda « Le Norvégien » est déjà dans l’onglet campagnes.");
  const r0 = sh.getLastRow() + 1;
  sh.getRange(r0, 1, rows.length, head.length).setValues(rows.map(r => head.map(h => h === "Statut" ? "Brouillon" : h in r ? r[h] : "")));
  ui_(rows.length + " campagne(s) ajoutée(s) en Brouillon. Vérifiez dates et textes, envoyez-vous un test, puis passez chaque ligne en « Programmée ».");
}

/* ================= Onglets et journal ================= */

function subsSheet_() {
  return sheet_(SHEET_SUBS, SUB_HEADERS, sh => {
    const c = h => SUB_HEADERS.indexOf(h) + 1;
    sh.getRange(2, c("Statut"), 999).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(["Inscrit", "Désinscrit"]).build());
  });
}
function campaignsSheet_() {
  return sheet_(SHEET_CAMPAIGNS, CAMP_HEADERS, sh => {
    const c = h => CAMP_HEADERS.indexOf(h) + 1;
    sh.getRange(2, c("Statut"), 999).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(CAMP_STATUS).build());
    sh.getRange(2, c("Envoi prévu"), 999).setNumberFormat("dd/MM/yyyy HH:mm");
    sh.getRange(2, 1, 999, CAMP_HEADERS.length).setWrap(true).setVerticalAlignment("top");
  });
}
function sendsSheet_() { return sheet_(SHEET_SENDS, SEND_HEADERS); }

function sheet_(name, headers, init) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    if (init) init(sh);
  }
  return sh;
}

function sends_() {
  const sh = sendsSheet_(), n = sh.getLastRow() - 1;
  return n < 1 ? [] : sh.getRange(2, 1, n, SEND_HEADERS.length).getValues();
}
function sentSet_(id) { return new Set(sends_().filter(r => String(r[0]) === id && r[3] === SENT).map(r => String(r[1]))); }
function logSend_(id, email, result) { sendsSheet_().appendRow([id, email, new Date(), result]); }

function findEmail_(sh, email) {
  const v = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
  for (let i = 1; i < v.length; i++) if (String(v[i][0]).trim().toLowerCase() === email) return i + 1;
  return 0;
}

function unsubToken_(email) { return token_("desinscription:" + String(email).trim().toLowerCase()); }
function unsubUrl_(email) { return SITE_URL + "/desinscription.html?e=" + encodeURIComponent(email) + "&t=" + unsubToken_(email); }
function campaignViewUrl_(id) { return SITE_URL + "/lettre.html?c=" + encodeURIComponent(id) + "&t=" + token_("lettre:" + id); }

/** Message dans le Sheet quand la fonction est lancée depuis le menu ; journal sinon (déclencheur, éditeur). */
function ui_(msg) {
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { console.log(msg); }
}
