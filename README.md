# Get-ShareAudit.ps1

<div align="center">

```
 ██████╗ ███████╗████████╗      ███████╗██╗  ██╗ █████╗ ██████╗ ███████╗
██╔════╝ ██╔════╝╚══██╔══╝      ██╔════╝██║  ██║██╔══██╗██╔══██╗██╔════╝
██║  ███╗█████╗     ██║   █████╗███████╗███████║███████║██████╔╝█████╗  
██║   ██║██╔══╝     ██║   ╚════╝╚════██║██╔══██║██╔══██║██╔══██╗██╔══╝  
╚██████╔╝███████╗   ██║         ███████║██║  ██║██║  ██║██║  ██║███████╗
 ╚═════╝ ╚══════╝   ╚═╝         ╚══════╝╚═╝  ╚═╝╚═╝  ╚═╝╚═╝  ╚═╝╚══════╝
             Audit Partages SharePoint & OneDrive — Microsoft 365
```

[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue?logo=powershell&logoColor=white)](https://docs.microsoft.com/powershell/)
[![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20Linux%20%7C%20macOS-informational)](https://github.com/PowerShell/PowerShell)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![PnP.PowerShell](https://img.shields.io/badge/PnP.PowerShell-2.0%2B-orange)](https://pnp.github.io/powershell/)
[![Microsoft Graph](https://img.shields.io/badge/Microsoft%20Graph-v1.0-00a1f1?logo=microsoft)](https://docs.microsoft.com/graph/)
[![Made with ❤️](https://img.shields.io/badge/Made%20with%20%E2%9D%A4%EF%B8%8F%20by-Ayi%20NEDJIMI-red)](https://ayinedjimi-consultants.fr)

**Audit automatisé et exhaustif de tous les partages SharePoint Online & OneDrive for Business d'un tenant Microsoft 365 — détection des liens anonymes, export CSV/HTML, prêt pour la production.**

[Guide complet FR](https://ayinedjimi-consultants.fr/articles/audit-partages-sharepoint-onedrive-powershell-anonymes) • [Signaler un bug](https://github.com/ayinedjimi/sharepoint-onedrive-share-audit/issues) • [Documentation Microsoft](https://docs.microsoft.com/sharepoint/)

</div>

---

## 🇫🇷 Documentation Française

### Présentation

`Get-ShareAudit.ps1` est un script PowerShell **prêt pour la production** qui permet d'auditer de manière exhaustive l'ensemble des liens de partage d'un tenant **Microsoft 365** : SharePoint Online, OneDrive for Business, et tous les sous-sites.

Conçu pour les équipes sécurité, les RSSI et les auditeurs conformité (RGPD, NIS 2, ISO 27001), il détecte en priorité les **liens anonymes** ("Anyone links") — c'est-à-dire les fichiers accessibles **sans aucune authentification** — qui constituent le risque de fuite de données le plus critique dans un environnement M365.

### Fonctionnalités clés

| Fonctionnalité | Détail |
|---|---|
| 🔍 **Énumération complète** | Tous les sites SharePoint + OneDrive, récursif |
| 🔴 **Détection liens anonymes** | Anyone links mis en évidence et comptabilisés |
| 📊 **Export CSV** | Séparateur `;`, encodage UTF-8, prêt pour Excel |
| 📄 **Rapport HTML** | Tableau interactif avec tri, filtre et code couleur |
| 🔁 **Retry automatique** | Back-off exponentiel sur throttling HTTP 429 |
| 📈 **Barre de progression** | Write-Progress sur chaque site analysé |
| 🔀 **Double moteur** | PnP.PowerShell (préféré) + Microsoft.Graph (fallback) |
| 🏁 **Résumé final** | Top 10 des items les plus partagés en console |

### Pré-requis

#### Environnement PowerShell

```powershell
# Option 1 : PnP.PowerShell (recommandé)
Install-Module PnP.PowerShell -MinimumVersion 2.0 -Scope CurrentUser

# Option 2 : Microsoft.Graph (fallback automatique)
Install-Module Microsoft.Graph -Scope CurrentUser
```

**Versions compatibles :** PowerShell 5.1 (Windows) · PowerShell 7.x (cross-platform)

#### Enregistrement de l'application Azure AD

1. Accéder au [portail Azure](https://portal.azure.com) → **Azure Active Directory** → **Inscriptions d'applications**
2. Cliquer **Nouvelle inscription** → nommer l'app (ex : `M365-ShareAudit`)
3. Ajouter les **autorisations d'application** (Application permissions, pas déléguées) :

| API | Permission | Justification |
|-----|-----------|---------------|
| Microsoft Graph | `Sites.Read.All` | Lire tous les sites SharePoint |
| Microsoft Graph | `Files.Read.All` | Lire tous les fichiers et permissions |
| Microsoft Graph | `User.Read.All` | Énumérer les utilisateurs pour OneDrive |
| SharePoint | `Sites.FullControl.All` | Requis pour PnP.PowerShell app-only |

4. Cliquer **Accorder le consentement administrateur**
5. Créer un **secret client** (Certificates & secrets) → noter la valeur

#### Rôles requis (selon mode)

- **SharePoint Administrator** ou **Global Administrator** (PnP app-only)
- **Global Reader** suffisant en mode Graph lecture seule

### Installation

```bash
# Cloner le dépôt
git clone https://github.com/ayinedjimi/sharepoint-onedrive-share-audit.git
cd sharepoint-onedrive-share-audit

# Débloquer le script (Windows)
Unblock-File .\Get-ShareAudit.ps1
```

### Utilisation

#### Audit basique (SharePoint uniquement)

```powershell
.\Get-ShareAudit.ps1 `
    -TenantId "contoso.onmicrosoft.com" `
    -ClientId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -ClientSecret "votre_secret_ici"
```

#### Audit complet avec OneDrive, liens anonymes uniquement

```powershell
.\Get-ShareAudit.ps1 `
    -TenantId "contoso.onmicrosoft.com" `
    -ClientId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -ClientSecret "votre_secret_ici" `
    -IncludeOneDrive `
    -AnonymousOnly `
    -OutputPath "C:\Audits\M365\$(Get-Date -Format 'yyyyMMdd')" `
    -Verbose
```

#### Exemple de sortie console

```
════════════════════════════════════════════════════════════════
  Get-ShareAudit.ps1 v2.0.0 — Audit Partages M365
  ayinedjimi-consultants.fr — Cybersécurité Microsoft 365
════════════════════════════════════════════════════════════════

[2026-05-21 09:12:03] [INFO]    Vérification des modules disponibles...
[2026-05-21 09:12:04] [INFO]    Module PnP.PowerShell détecté. Connexion en mode application...
[2026-05-21 09:12:06] [SUCCESS] Connecté via PnP.PowerShell à https://contoso-admin.sharepoint.com
[2026-05-21 09:12:08] [SUCCESS] 47 site(s) SharePoint trouvé(s).
[2026-05-21 09:12:09] [SUCCESS] 312 OneDrive trouvé(s).
```

#### Captures simulées — Rapport HTML

```
┌─────────────────────────────────────────────────────────────────┐
│  Rapport Audit Partages Microsoft 365                            │
│  Généré le 21/05/2026 09:45:12 — Durée : 33.2 min               │
├────────────────┬────────────┬─────────────┬──────────────────────┤
│ Sites analysés │ Éléments   │ Liens total │ Liens ANONYMES       │
│      359       │  48 291    │    1 847    │    ████ 23 ████       │
└────────────────┴────────────┴─────────────┴──────────────────────┘

Tous les partages détectés
[Filtrer...                                                       ]
┌──────────────────────┬──────────┬──────────────────┬───────────┐
│ Site                 │ Type     │ Élément          │ Partage   │
├──────────────────────┼──────────┼──────────────────┼───────────┤
│ Equipe-Marketing     │SharePoint│ brief-Q3-2026.pdf│[Anyone]   │ ← ROUGE
│ John Smith - OD      │OneDrive  │ contrat-NDA.docx │[Anyone]   │ ← ROUGE
│ Intranet-RH          │SharePoint│ organigramme.xlsx│[Company]  │
│ Projet-Alpha         │SharePoint│ spec-technique.md│[Specific] │
└──────────────────────┴──────────┴──────────────────┴───────────┘
```

### Paramètres de référence

| Paramètre | Type | Obligatoire | Description |
|-----------|------|-------------|-------------|
| `-TenantId` | String | Oui | Domaine ou GUID du tenant Azure AD |
| `-ClientId` | String | Oui | Application (client) ID — format GUID |
| `-ClientSecret` | String | Oui | Secret client de l'application Azure AD |
| `-OutputPath` | String | Non | Dossier de sortie (défaut : répertoire courant) |
| `-IncludeOneDrive` | Switch | Non | Inclure les MySites OneDrive de tous les users |
| `-AnonymousOnly` | Switch | Non | Rapporter uniquement les liens "Anyone" |
| `-Verbose` | Switch | Non | Afficher les messages de diagnostic détaillés |

### Révoquer les liens anonymes identifiés

```powershell
# Exemple de remédiation via PnP.PowerShell
Connect-PnPOnline -Url "https://contoso.sharepoint.com/sites/Marketing" `
                  -ClientId $ClientId -ClientSecret $ClientSecret

# Lister les liens de partage d'un fichier
$links = Get-PnPFileSharingLink -FileUrl "/sites/Marketing/docs/brief-Q3.pdf"

# Supprimer tous les liens anonymes
$links | Where-Object { $_.Link.Scope -like "*anonymous*" } | ForEach-Object {
    Remove-PnPFileSharingLink -FileUrl "/sites/Marketing/docs/brief-Q3.pdf" -Id $_.Id
}
```

### Avertissement légal

> ⚠️ **Ce script est destiné exclusivement aux administrateurs et auditeurs disposant des autorisations légales et contractuelles pour accéder au tenant Microsoft 365 concerné.**
>
> Toute utilisation sur un tenant tiers sans autorisation explicite est **illégale** et peut engager votre responsabilité pénale. Consultez votre équipe juridique avant tout déploiement.

---

## 🇬🇧 English Documentation

### Overview

`Get-ShareAudit.ps1` is a **production-ready** PowerShell script that performs an exhaustive audit of all sharing links in a **Microsoft 365 tenant**: SharePoint Online, OneDrive for Business, and all sub-sites.

Designed for security teams, CISOs, and compliance auditors (GDPR, NIS 2, ISO 27001), it primarily detects **anonymous links** ("Anyone links") — files accessible **without any authentication** — which represent the most critical data leak risk in an M365 environment.

### Key Features

| Feature | Details |
|---|---|
| 🔍 **Full enumeration** | All SharePoint sites + OneDrive, recursive |
| 🔴 **Anonymous link detection** | Anyone links highlighted and counted |
| 📊 **CSV export** | Semicolon separator, UTF-8 encoding, Excel-ready |
| 📄 **HTML report** | Interactive table with sort, filter, and color coding |
| 🔁 **Auto-retry** | Exponential back-off on HTTP 429 throttling |
| 📈 **Progress bar** | Write-Progress for each site analyzed |
| 🔀 **Dual engine** | PnP.PowerShell (preferred) + Microsoft.Graph (fallback) |
| 🏁 **Final summary** | Top 10 most-shared items in console |

### Prerequisites

#### PowerShell Modules

```powershell
# Option 1: PnP.PowerShell (recommended)
Install-Module PnP.PowerShell -MinimumVersion 2.0 -Scope CurrentUser

# Option 2: Microsoft.Graph (automatic fallback)
Install-Module Microsoft.Graph -Scope CurrentUser
```

**Compatible versions:** PowerShell 5.1 (Windows) · PowerShell 7.x (cross-platform)

#### Azure AD App Registration

1. Go to [Azure Portal](https://portal.azure.com) → **Azure Active Directory** → **App registrations**
2. Click **New registration** → name the app (e.g., `M365-ShareAudit`)
3. Add **Application permissions** (not delegated):

| API | Permission | Purpose |
|-----|-----------|---------|
| Microsoft Graph | `Sites.Read.All` | Read all SharePoint sites |
| Microsoft Graph | `Files.Read.All` | Read all files and permissions |
| Microsoft Graph | `User.Read.All` | Enumerate users for OneDrive |
| SharePoint | `Sites.FullControl.All` | Required for PnP.PowerShell app-only |

4. Click **Grant admin consent**
5. Create a **client secret** → copy the value

#### Required Roles

- **SharePoint Administrator** or **Global Administrator** (PnP app-only mode)
- **Global Reader** sufficient in Graph read-only mode

### Installation

```bash
# Clone the repository
git clone https://github.com/ayinedjimi/sharepoint-onedrive-share-audit.git
cd sharepoint-onedrive-share-audit

# Unblock the script (Windows)
Unblock-File .\Get-ShareAudit.ps1
```

### Usage

#### Basic audit (SharePoint only)

```powershell
.\Get-ShareAudit.ps1 `
    -TenantId "contoso.onmicrosoft.com" `
    -ClientId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -ClientSecret "your_secret_here"
```

#### Full audit with OneDrive, anonymous links only

```powershell
.\Get-ShareAudit.ps1 `
    -TenantId "contoso.onmicrosoft.com" `
    -ClientId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -ClientSecret "your_secret_here" `
    -IncludeOneDrive `
    -AnonymousOnly `
    -OutputPath "C:\Audits\M365\$(Get-Date -Format 'yyyyMMdd')" `
    -Verbose
```

### Parameters Reference

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `-TenantId` | String | Yes | Azure AD tenant domain or GUID |
| `-ClientId` | String | Yes | Application (client) ID — GUID format |
| `-ClientSecret` | String | Yes | Azure AD app client secret |
| `-OutputPath` | String | No | Output folder (default: current directory) |
| `-IncludeOneDrive` | Switch | No | Include all users' OneDrive MySites |
| `-AnonymousOnly` | Switch | No | Report only "Anyone" anonymous links |
| `-Verbose` | Switch | No | Show detailed diagnostic messages |

### Revoking Anonymous Links

```powershell
# Example remediation via PnP.PowerShell
Connect-PnPOnline -Url "https://contoso.sharepoint.com/sites/Marketing" `
                  -ClientId $ClientId -ClientSecret $ClientSecret

# Get sharing links for a file
$links = Get-PnPFileSharingLink -FileUrl "/sites/Marketing/docs/brief-Q3.pdf"

# Remove all anonymous links
$links | Where-Object { $_.Link.Scope -like "*anonymous*" } | ForEach-Object {
    Remove-PnPFileSharingLink -FileUrl "/sites/Marketing/docs/brief-Q3.pdf" -Id $_.Id
}
```

### Legal Notice

> ⚠️ **This script is intended exclusively for administrators and auditors with the legal and contractual authorizations to access the Microsoft 365 tenant in question.**
>
> Any unauthorized use on a third-party tenant is **illegal** and may engage your criminal liability. Consult your legal team before any deployment.

---

## License

```
MIT License

Copyright (c) 2026 Ayi NEDJIMI

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
```

---

<div align="center">

Made with ❤️ by **[Ayi NEDJIMI](https://ayinedjimi-consultants.fr)** · Cybersécurité Microsoft 365

**[ayinedjimi-consultants.fr](https://ayinedjimi-consultants.fr)**

</div>
