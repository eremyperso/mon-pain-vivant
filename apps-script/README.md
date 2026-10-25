# Commandes et e-mails automatiques (Google Apps Script)

Le script [Code.gs](Code.gs) est lié au Google Sheet **Commandes**, privé et non partagé. Il y enregistre chaque commande (onglet `commandes`, créé automatiquement) et envoie les e-mails. Les réglages (`CONTACT_EMAIL`…) sont lus dans l'onglet `contact` du Sheet catalogue (`CATALOGUE_SPREADSHEET_ID`), public en lecture pour le site.

Les e-mails sont envoyés en HTML aux couleurs du site, avec une version texte de secours pour les messageries qui n'affichent pas le HTML. Chaque e-mail client contient un lien « Voir dans le navigateur » qui ouvre [commande.html](../commande.html) sur le site.

## Installation (une seule fois)

1. Ouvrez le Google Sheet **Commandes**, puis **Extensions > Apps Script**. Collez le contenu de `Code.gs` et enregistrez.
2. **Déployer > Nouveau déploiement**, type **Application Web** :
   - Exécuter en tant que : **Moi**
   - Qui a accès : **Tout le monde**

   Autorisez l'accès, puis copiez l'URL qui se termine par `/exec`.
3. Dans l'onglet `contact` du Sheet **catalogue**, ajoutez ces lignes (colonne A = clé, colonne B = valeur) :
   - `ORDER_WEBHOOK_URL` : l'URL `/exec` copiée à l'étape 2
   - `CONTACT_EMAIL` : l'adresse affichée aux clients ; leurs réponses aux e-mails arrivent ici
   - `ADMIN_EMAIL` : l'adresse qui reçoit l'alerte « Nouvelle commande », par exemple un alias dédié comme `gestion@autantdupain.eremy.fr`. L'alias doit exister chez l'hébergeur du domaine et rediriger vers votre boîte. Cet onglet étant public, préférez la propriété de script `ADMIN_EMAIL` (étape 5) pour une adresse personnelle : elle est prioritaire. Si aucune des deux n'est définie, l'alerte part vers `CONTACT_EMAIL`.
   - `ADRESSE_REMISE` (facultatif) : l'adresse de remise, rappelée dans l'e-mail « Commande prête »
   - `WERO_ID` (déjà utilisé par le site) : rappelé dans l'e-mail de commande tant que le virement Wero n'est pas reçu
4. Menu **Déclencheurs** (icône horloge) > **Ajouter un déclencheur** :
   - Fonction : `onEditPaiement`
   - Source : **Depuis la feuille de calcul**
   - Type : **Lors de la modification**
5. Facultatif, pour que les paiements PayPal soient vérifiés automatiquement : **Paramètres du projet > Propriétés du script**. Ajoutez `PAYPAL_CLIENT_ID` et `PAYPAL_SECRET`, et pour les tests `PAYPAL_SANDBOX_CLIENT_ID` et `PAYPAL_SANDBOX_SECRET`. Au même endroit, `ADMIN_EMAIL` reçoit l'alerte « Nouvelle commande » sans que l'adresse soit visible dans le Sheet public.

Chaque modification de `Code.gs` doit être redéployée : **Gérer les déploiements > Modifier > Nouvelle version**. L'URL reste la même.

La propriété de script `VIEW_SECRET` est créée automatiquement à la première commande : elle signe les liens « Voir dans le navigateur ». Si vous la supprimez, les liens des e-mails déjà envoyés ne fonctionneront plus.

## Fonctionnement

### À la commande

| Mode de paiement | E-mail client | Alerte gestionnaire |
|---|---|---|
| PayPal vérifié | « Commande confirmée et payée » (un seul e-mail) | immédiate |
| PayPal non vérifié, Wero, en main propre | « Commande confirmée », paiement en attente | immédiate |

L'alerte gestionnaire a un objet préfixé `[ATDP · Commande]` (pratique pour un filtre ou un libellé), un bouton vers la ligne du Google Sheet, et « Répondre » écrit directement au client.

### Quand vous modifiez le Google Sheet

| Colonne | Valeur | E-mail envoyé au client |
|---|---|---|
| **Paiement** | `Payé` | Paiement reçu |
| **Paiement** | `Remboursé` | Remboursement effectué |
| **Statut commande** | `Prête` | Commande prête (avec le montant à régler s'il n'est pas encore payé) |
| **Statut commande** | `Remise` | Remerciement et conseils de conservation |
| **Statut commande** | `Annulée` | Annulation, avec l'annonce du remboursement si la commande était payée |

Chaque e-mail n'est envoyé qu'une fois par commande. La colonne **Notifications** garde la trace des envois (par exemple `commande 01/10 14:32 · payé 02/10 09:10`). Pour renvoyer un e-mail, effacez sa mention dans cette colonne (et, pour `Payé`, la colonne **Mail paiement envoyé**) puis ressaisissez la valeur.

Lors de la mise à jour, la colonne **Notifications** est ajoutée automatiquement et les commandes existantes sont marquées « (avant mise à jour) » : elles ne recevront pas d'e-mail pour leur statut actuel.

La fonction `envoyerPaiementsEnAttente`, à lancer à la main depuis l'éditeur, envoie les e-mails dus qui n'ont pas encore été envoyés (par exemple après une panne).

## Lettre d'information (campagnes programmées)

Le fichier [Newsletter.gs](Newsletter.gs) s'ajoute au même projet Apps Script (**+ > Script**, nommé `Newsletter`). Il envoie « la liste des pains de la semaine » aux clients qui ont coché la case lors d'une commande, à la date prévue, avec le contenu écrit dans le Google Sheet **Commandes**.

### Installation

1. Collez `Newsletter.gs` dans un nouveau fichier du projet, mettez à jour `Code.gs`, puis **redéployez** (Gérer les déploiements > Modifier > Nouvelle version).
2. Rechargez le Google Sheet **Commandes** : un menu **Lettre** apparaît. Lancez **Activer l'envoi automatique** une fois (autorisez l'accès). Les onglets `abonnes`, `campagnes` et `envois` sont créés.
3. Facultatif : dans l'onglet **Produits** du Sheet catalogue, ajoutez une colonne `nouveau`. Mettez `oui`, ou une date de fin (`06/12/2026`) : le produit porte alors le badge « Nouveau » sur le site et dans les lettres, et passe en tête de la grille.

### Les onglets

| Onglet | Rôle |
|---|---|
| `abonnes` | Une ligne par e-mail. Remplie automatiquement quand la case est cochée à la commande. **Statut** `Inscrit` / `Désinscrit`. Les colonnes Commandes, Dernière commande et Pains achetés sont recalculées depuis l'onglet `commandes` avant chaque envoi. |
| `campagnes` | **L'agenda** : une ligne = un e-mail. Seules les lignes `Programmée` partent, dans le quart d'heure qui suit **Envoi prévu**. Les colonnes **Envoyés** et **Rapport** sont remplies par le script. |
| `envois` | Journal campagne × destinataire. Personne ne reçoit deux fois la même campagne, même si l'envoi est repris. |

### Écrire une campagne

| Colonne | Contenu |
|---|---|
| ID | Identifiant unique, repris dans les statistiques Google Analytics (`utm_campaign`) |
| Statut | `Brouillon` → `Programmée` → (`En cours`) → `Envoyée`, ou `Annulée` |
| Envoi prévu | Date et heure, par exemple `01/11/2026 09:00` |
| Liste | Les destinataires (voir ci-dessous) |
| Objet, Pré-en-tête, Titre | Le pré-en-tête est la ligne grise affichée après l'objet dans la boîte de réception |
| Texte | Un paragraphe par ligne (Alt+Entrée dans la cellule) |
| Produit vedette | Nom exact d'un produit de l'onglet Produits : photo, prix, poids et badge « Nouveauté » sont repris du catalogue |
| Texte vedette | Facultatif : remplace la description du catalogue sous la photo |
| Produits associés | Noms séparés par `;`. Les produits épuisés sont retirés automatiquement |
| Bouton, Lien bouton | Facultatifs : par défaut « Commander … » vers la fiche du produit vedette |
| Note | Phrase en italique en bas de la lettre |

Variables utilisables dans tous les textes : `{{bonjour}}` (« Bonjour Marie » ou « Bonjour »), `{{prenom}}`, `{{dernier_pain}}` (dernier pain commandé), `{{nb_commandes}}`, `{{prochaine_remise}}` (« jeudi 5 novembre »), `{{date_limite}}` (« mardi 3 novembre, 18 h »), `{{produit}}` et `{{prix}}` (produit vedette). Ajoutez une valeur par défaut après `|` : `{{dernier_pain|nos pains aux graines}}`.

**Listes** : critères séparés par `;`, tous obligatoires. `|` sépare plusieurs valeurs possibles.

| Critère | Destinataires |
|---|---|
| `tous` | Tous les inscrits |
| `clients` / `commandes:3` | Au moins 1 / 3 commandes |
| `tag:Graines\|Seigle` | Ont acheté un produit portant l'un de ces tags |
| `a-commande:Le norvégien` / `pas-commande:Le norvégien` | Ont / n'ont pas commandé ce produit |
| `inactifs:30` | Aucune commande depuis 30 jours |
| `pas-recu:NORV-01-avant-premiere` | N'ont pas reçu cette campagne |
| `email:marie@exemple.fr` | Uniquement ces abonnés (pratique pour un essai réel) |
| `test` | Uniquement l'adresse gestionnaire |

### Avant chaque envoi

Sélectionnez la ligne dans `campagnes`, puis **Lettre > Envoyer un test de la campagne sélectionnée**. Le test arrive à l'adresse gestionnaire, personnalisé comme pour le premier destinataire, et la fenêtre indique combien de personnes la liste touche aujourd'hui. Passez ensuite la ligne en `Programmée`.

Si un produit est introuvable ou que le produit vedette est épuisé, rien n'est envoyé : la ligne repasse en `Brouillon`, le **Rapport** donne la raison et vous recevez une alerte.

### Quotas

Sans Brevo, Gmail permet environ 100 e-mails par jour : le script en garde 20 pour les commandes. Avec Brevo, la lettre s'arrête à 250 e-mails par jour (propriété de script `NEWSLETTER_MAX_JOUR` pour changer ce plafond). Au-delà, la campagne reste `En cours` et reprend toute seule le lendemain.

### Désinscription

Chaque lettre contient un lien « Se désinscrire » ([desinscription.html](../desinscription.html), lien signé : impossible de désinscrire quelqu'un d'autre) qui agit immédiatement et propose de se réinscrire en cas d'erreur. Avec Brevo, les en-têtes `List-Unsubscribe` ajoutent aussi le bouton « Se désabonner » de Gmail. Un client qui recoche la case lors d'une commande est réinscrit. Les lignes désinscrites restent dans `abonnes` : ne les effacez pas, elles gardent la trace de la demande.

« Voir dans le navigateur » ouvre [lettre.html](../lettre.html), sans personnalisation.
