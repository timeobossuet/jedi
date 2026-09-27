# JEDI

![Google Apps Script](https://img.shields.io/badge/Google%20Apps%20Script-4285F4?logo=google&logoColor=white)
![JavaScript](https://img.shields.io/badge/JavaScript-ES5%2B-F7DF1E?logo=javascript&logoColor=111111)
![Licence MIT](https://img.shields.io/badge/licence-MIT-2ea44f)

> A Google Workspace document and operations layer for French Junior-Entreprises.

## Public repository scope

This public repository contains the **document generation features and Google Workspace integrations** of JEDI. It is the open technical showcase of the Apps Script layer used to generate and manage documents with Google Docs, Google Slides, Google Drive, Gmail, Google Chat and Firestore.

The complete JEDI ERP is not included in this repository. The full product also contains:

- an operational dashboard;
- two React Native mobile applications;
- a web application for desktop use;
- additional workflows and features for the complete ERP experience.

Access to the complete source code is available on request. Please contact me through my GitHub profile to discuss the project, a demonstration or the private codebase.

## About JEDI

JEDI is an ERP designed for **French Junior-Entreprises**. It reflects their specific vocabulary, processes and operational needs, including client studies, project documents, members, meetings, approvals and JEH tracking.

The user interface was designed with accessibility and onboarding in mind. Even users who are new to Junior-Entreprises should be able to understand the workflow, follow the lifecycle of a mission and see how JEHs are planned and used throughout a project.

The goal is not only to automate administration, but also to make the operational cycle understandable: from the first study request to document production, review, validation and completion.

## What is included here

- **Document generation**: conventions, quotations, purchase orders, meeting reports and generic documents based on Google Docs and Google Slides templates.
- **Study workflows**: document statuses, versions, Drive folders and Firestore synchronization.
- **Google Workspace automation**: Drive, Docs, Slides, Gmail, Google Chat and Apps Script Properties.
- **Operational support**: Google Chat notifications and support workflows.

## Repository structure

```text
generateurDocuments/  Document generation and Google Workspace workflows
gestionEtudes/        Study data, Drive, Firestore and support integrations
secretaireGeneral/    Shared administrative Apps Script features
logos/                Project visual identity
docs/                 Configuration and operational documentation
```

The source is organized by business domain because the files are deployed to Google Apps Script projects and rely on shared global functions. The repository does not contain the React Native applications or the complete dashboard codebase.

## Installation

1. Create a Google Apps Script project in the organization's Google Workspace environment.
2. Import the relevant source folders into the Apps Script project.
3. Enable the advanced services required by the deployment, including Drive API when Office template conversion is used.
4. Configure the Script Properties described in [docs/configuration.md](docs/configuration.md).
5. Create the required Google Docs and Google Slides templates and Drive folders.
6. Deploy the Apps Script functions or web application required by the instance.

## Configuration and security

Secrets, tokens, Firebase keys, webhooks and environment-specific resource identifiers must be stored as **Script Properties**. They must never be committed to Git.

The configuration names, required permissions and deployment notes are documented in [docs/configuration.md](docs/configuration.md). Before publishing or deploying a copy, review the Git history and revoke any credential that may have been exposed.

## Quality and limitations

This repository contains Google Apps Script code that runs inside Google's runtime and cannot be fully simulated locally. Local validation covers JavaScript syntax; business workflows should be tested in a dedicated Apps Script environment with non-sensitive data.

## Project context

JEDI was created to reduce repetitive administrative work while making the operating model of a French Junior-Entreprise easier to understand and follow.

## License

The public code is distributed under the [MIT License](LICENSE).
