# Configuration

JEDI utilise les **Script Properties** de Google Apps Script pour isoler la configuration d'une instance du code source. Les valeurs ci-dessous sont des exemples de noms uniquement : ne pas publier les valeurs réelles.

## Propriétés obligatoires

| Propriété | Usage |
| --- | --- |
| `SG_DRIVE_FOLDER_ID` | Dossier racine Drive du secrétariat général |
| `DOC_WORKFLOW_WEBHOOK_URL` | Webhook Google Chat utilisé pour les documents en relecture |
| `FIREBASE_PROJECT_ID` | Projet Firebase / Google Cloud |
| `FIREBASE_CLIENT_EMAIL` | Compte de service Firestore |
| `FIREBASE_PRIVATE_KEY` | Clé privée du compte de service |

## Propriétés optionnelles

| Propriété | Usage |
| --- | --- |
| `FIRESTORE_DATABASE_ID` | Base Firestore, `(default)` par défaut |
| `FIREBASE_TOKEN_URI` | Endpoint OAuth2, `https://oauth2.googleapis.com/token` par défaut |
| `STUDIES_DRIVE_ROOT_FOLDER_ID` | Dossier Drive racine des études |
| `DSI_CHAT_WEBHOOK_URL` | Webhook direct pour le support DSI |
| `DSI_CHAT_WEBHOOK_KEY` / `DSI_CHAT_WEBHOOK_TOKEN` | Alternative au webhook DSI direct |
| `SG_INTERVENANT_CHAT_WEBHOOK_URL` | Webhook des notifications intervenants |
| `SG_INTERVENANT_CHAT_WEBHOOK_KEY` / `SG_INTERVENANT_CHAT_WEBHOOK_TOKEN` | Alternative au webhook intervenants |

## Mise en place

Dans Apps Script : **Project Settings > Script Properties > Add script property**. Les propriétés peuvent aussi être définies par les fonctions d'initialisation prévues dans le code, notamment `setFirestoreSettings`.

Le compte qui exécute le script doit avoir accès aux dossiers Drive, aux modèles, au tableur et au projet Firebase configurés. Utiliser un projet de test et des comptes de service distincts pour le développement.

## Contrôle avant publication

- Rechercher les URLs contenant `key=`, `token=` ou `Bearer` dans le dépôt.
- Vérifier qu'aucune clé privée, export JSON de compte de service, donnée adhérent ou donnée client n'est suivie par Git.
- Révoquer et régénérer immédiatement toute clé déjà publiée dans l'historique Git.
- Ne jamais ajouter les valeurs de Script Properties à un fichier de configuration versionné.