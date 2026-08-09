# Facturation Ground & Rail — CEVA

Application de facturation (opérations import/export/transfert/mise à terre,
tarification Sympos, factures PDF/Excel, tableau de bord) pour un client
unique : CEVA Ground and Rail.

Elle est responsive (utilisable sur mobile, tablette, ordinateur) et les
données sont partagées en ligne entre tous les utilisateurs via Supabase.

## 1. Créer la base de données (5 minutes, gratuit)

1. Allez sur https://supabase.com et créez un compte gratuit.
2. Cliquez sur **New project**, donnez-lui un nom, choisissez une région
   proche (ex: Europe) et un mot de passe (à conserver, pas besoin de le
   redonner ensuite).
3. Une fois le projet créé, allez dans **SQL Editor** (menu de gauche) >
   **New query**, collez le contenu du fichier `supabase.sql` fourni dans ce
   projet, puis cliquez **Run**. Cela crée la table qui stockera toutes les
   données (opérations, factures, tarifs, paramètres).
4. Allez dans **Project Settings > API**. Notez deux valeurs :
   - **Project URL** (ex: `https://xxxxx.supabase.co`)
   - **anon public key** (une longue chaîne de caractères)

## 2. Déployer le site (5 minutes, gratuit)

Le plus simple est **Netlify** :

1. Allez sur https://app.netlify.com et créez un compte gratuit.
2. Cliquez **Add new site > Deploy manually**, puis faites glisser le
   dossier de ce projet tel quel — Netlify propose aussi de connecter un
   dépôt GitHub si vous préférez un déploiement automatique à chaque
   modification (dans ce cas, poussez ce dossier sur un nouveau repo GitHub
   puis "Import from Git" dans Netlify).
3. Une fois le site créé, allez dans **Site configuration > Environment
   variables** et ajoutez :
   - `VITE_SUPABASE_URL` = (l'URL notée à l'étape 1)
   - `VITE_SUPABASE_ANON_KEY` = (la clé notée à l'étape 1)
   - (optionnel) `VITE_APP_PASSCODE` = un code d'accès simple si vous voulez
     un minimum de protection sur le lien
4. Si vous avez déployé via Git, allez dans **Deploys > Trigger deploy** pour
   relancer un build avec les nouvelles variables. Si vous avez glissé le
   dossier manuellement, il faut reconstruire le site localement d'abord
   (voir "Construire localement" ci-dessous) puis redéposer le dossier
   `dist/`, car Netlify ne "build" pas un dépôt déposé en glisser-déposer —
   **le déploiement via GitHub est donc recommandé** pour que Netlify exécute
   `npm install` et `npm run build` automatiquement avec vos variables.

Alternative équivalente : **Vercel** (https://vercel.com), même principe
(import du projet, ajout des variables d'environnement dans Settings >
Environment Variables, redeploy).

## 3. Construire et tester localement (optionnel)

Nécessite Node.js installé sur votre ordinateur.

```bash
cp .env.example .env
# éditez .env avec vos vraies valeurs Supabase
npm install
npm run dev
```

Ouvrez l'adresse affichée (ex: http://localhost:5173).

Pour générer la version de production (dossier `dist/`) :

```bash
npm run build
```

## Sécurité — à savoir

Cette version n'a **pas de système de comptes/connexion**. Toute personne
disposant du lien de votre site peut consulter et modifier les données
(opérations, factures, tarifs), car la clé Supabase utilisée est publique
par nature (elle est visible dans le code du navigateur). Le code d'accès
optionnel (`VITE_APP_PASSCODE`) n'est qu'un frein léger, pas une vraie
sécurité. Pour un usage en production avec plusieurs comptes utilisateurs et
des permissions, il faudrait ajouter l'authentification Supabase — possible
dans une prochaine itération si besoin.

## Structure du projet

- `src/App.jsx` — toute l'application (mêmes fonctionnalités que la version
  précédente : opérations, facturation, tarifs Sympos, tableau de bord,
  export Excel/PDF).
- `src/storage.js` — connexion à Supabase et fonctions de lecture/écriture.
- `supabase.sql` — schéma de la base de données à exécuter une fois.
