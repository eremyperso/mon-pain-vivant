# Commandes et e-mails automatiques (Google Apps Script)

Le script [Code.gs](Code.gs) enregistre chaque commande dans l'onglet `commandes` du Google Sheet et envoie les e-mails depuis votre compte Google.

## Installation (une seule fois)

1. Ouvrez le Google Sheet, puis **Extensions > Apps Script**. Collez le contenu de `Code.gs` et enregistrez.
2. **Déployer > Nouveau déploiement**, type **Application Web** :
   - Exécuter en tant que : **Moi**
   - Qui a accès : **Tout le monde**

   Autorisez l'accès, puis copiez l'URL qui se termine par `/exec`.
3. Dans l'onglet `contact` du Sheet, ajoutez ces lignes (colonne A = clé, colonne B = valeur) :
   - `ORDER_WEBHOOK_URL` : l'URL `/exec` copiée à l'étape 2
   - `CONTACT_EMAIL` : l'adresse qui reçoit une copie de chaque e-mail
4. Menu **Déclencheurs** (icône horloge) > **Ajouter un déclencheur** :
   - Fonction : `onEditPaiement`
   - Source : **Depuis la feuille de calcul**
   - Type : **Lors de la modification**
5. Facultatif, pour que les paiements PayPal soient vérifiés automatiquement : **Paramètres du projet > Propriétés du script**. Ajoutez `PAYPAL_CLIENT_ID` et `PAYPAL_SECRET`, et pour les tests `PAYPAL_SANDBOX_CLIENT_ID` et `PAYPAL_SANDBOX_SECRET`.

Chaque modification de `Code.gs` doit être redéployée : **Gérer les déploiements > Modifier > Nouvelle version**. L'URL reste la même.

## Fonctionnement

| Mode de paiement | E-mail « Commande validée » | E-mail « Paiement confirmé » |
|---|---|---|
| PayPal | immédiat | immédiat, si la transaction est vérifiée auprès de PayPal |
| Wero | immédiat | quand vous passez la colonne **Paiement** à `Payé` |
| En main propre | immédiat | quand vous passez la colonne **Paiement** à `Payé` |

Les colonnes **Paiement** et **Statut commande** proposent une liste de valeurs pour le suivi.
