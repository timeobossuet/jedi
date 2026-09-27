# JEDI

> ERP métier pour piloter les études, les documents et la vie administrative d'une Junior-Entreprise.

![Google Apps Script](https://img.shields.io/badge/Google%20Apps%20Script-4285F4?logo=google&logoColor=white)
![JavaScript](https://img.shields.io/badge/JavaScript-ES5%2B-F7DF1E?logo=javascript&logoColor=111111)
![Licence MIT](https://img.shields.io/badge/licence-MIT-2ea44f)

JEDI centralise les flux opérationnels de **Junior INSA Services** dans Google Workspace : suivi des études, génération de livrables, échanges Google Chat et administration de l'association.

## Fonctionnalités

- **Gestion des études** : suivi des informations, intervenants, documents et synchronisation Drive / Firestore.
- **Génération documentaire** : création de conventions, devis, bons de commande, comptes rendus et documents génériques à partir de modèles Google Docs et Slides.
- **Secrétariat général** : membres, réunions, votes, licences, RGPD et configuration métier.
- **Support opérationnel** : signalement d'incidents et notifications Google Chat.
- **Interfaces web** : tableaux de bord HTML intégrés aux Google Apps Script.

## Architecture

Le projet est volontairement organisé par domaine métier. Cette structure correspond aux fichiers déployés dans Google Apps Script et évite de casser les chemins des templates HTML.

```text
generateurDocuments/  Génération et remplissage des documents
gestionEtudes/        Suivi des études, Drive, Firestore et support
secretaireGeneral/    Administration, réunions, votes et RGPD
logos/                Identité visuelle
docs/                 Documentation de configuration et d'exploitation
```

Les fichiers JavaScript et HTML sont chargés dans un ou plusieurs projets Google Apps Script selon le déploiement retenu. Le code s'appuie sur les services natifs Google Apps Script : Drive, Sheets, Docs, Slides, Gmail, Cache et Properties Service.

## Installation

1. Créer un projet Google Apps Script lié au Google Drive de l'organisation.
2. Importer les dossiers métier et leurs fichiers dans le projet Apps Script.
3. Activer les services avancés utilisés par l'instance, notamment Drive API si la génération documentaire l'exige.
4. Configurer les Script Properties décrites dans [docs/configuration.md](docs/configuration.md).
5. Créer les modèles Google Docs / Slides et les dossiers Drive attendus par l'organisation.
6. Déployer l'application web ou les fonctions nécessaires depuis Apps Script.

## Configuration et sécurité

Les secrets, tokens, clés Firebase, webhooks et identifiants de ressources sont des **Script Properties**. Ils ne doivent jamais être commités dans Git.

La procédure complète, les noms des propriétés et les droits Google Workspace sont documentés dans [docs/configuration.md](docs/configuration.md). Avant une première publication, vérifier également l'historique Git et révoquer toute clé qui aurait déjà été exposée.

## Qualité et limites connues

Le dépôt contient du code Apps Script exécuté dans l'environnement Google, qui n'est pas entièrement simulable localement. La validation locale couvre donc la syntaxe JavaScript ; les scénarios métier doivent être vérifiés dans un environnement Apps Script de test avec des données non sensibles.

## Contexte

JEDI a été conçu pour répondre aux besoins réels d'une Junior-Entreprise et réduire les tâches répétitives de gestion, de production documentaire et de coordination.

## Licence

Distribué sous licence [MIT](LICENSE).
