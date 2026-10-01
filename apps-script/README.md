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
   - `ADMIN_EMAIL` : l'adresse qui reçoit l'alerte « Nouvelle commande », par exemple un alias dédié comme `gestion@autantdupain.eremy.fr`. L'alias doit exister chez l'hébergeur du domaine et rediriger vers votre boîte. Si la ligne est absente, l'alerte part vers `CONTACT_EMAIL`.
   - `ADRESSE_REMISE` (facultatif) : l'adresse de remise, rappelée dans l'e-mail « Commande prête »
   - `WERO_ID` (déjà utilisé par le site) : rappelé dans l'e-mail de commande tant que le virement Wero n'est pas reçu
4. Menu **Déclencheurs** (icône horloge) > **Ajouter un déclencheur** :
   - Fonction : `onEditPaiement`
   - Source : **Depuis la feuille de calcul**
   - Type : **Lors de la modification**
5. Facultatif, pour que les paiements PayPal soient vérifiés automatiquement : **Paramètres du projet > Propriétés du script**. Ajoutez `PAYPAL_CLIENT_ID` et `PAYPAL_SECRET`, et pour les tests `PAYPAL_SANDBOX_CLIENT_ID` et `PAYPAL_SANDBOX_SECRET`.

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
