# Get-ShareAudit.ps1

<div align="center">

```
 ███████╗██╗  ██╗ █████╗ ██████╗ ███████╗ █████╗ ██╗   ██╗██████╗ ██╗████████╗
 ██╔════╝██║  ██║██╔══██╗██╔══██╗██╔════╝██╔══██╗██║   ██║██╔══██╗██║╚══██╔══╝
 ███████╗███████║███████║██████╔╝█████╗  ███████║██║   ██║██║  ██║██║   ██║   
 ╚════██║██╔══██║██╔══██║██╔══██╗██╔══╝  ██╔══██║██║   ██║██║  ██║██║   ██║   
 ███████║██║  ██║██║  ██║██║  ██║███████╗██║  ██║╚██████╔╝██████╔╝██║   ██║   
 ╚══════╝╚═╝  ╚═╝╚═╝  ╚═╝╚═╝  ╚═╝╚══════╝╚═╝  ╚═╝ ╚═════╝ ╚═════╝ ╚═╝   ╚═╝   

          Audit Partages SharePoint & OneDrive — Microsoft 365
```

[![PowerShell 5.1+](https://img.shields.io/badge/PowerShell-5.1%2B-blue?style=flat-square&logo=powershell&logoColor=white)](https://docs.microsoft.com/powershell/)
[![Windows | Linux | macOS](https://img.shields.io/badge/Platform-Windows%20%7C%20Linux%20%7C%20macOS-informational?style=flat-square)](https://github.com/PowerShell/PowerShell)
[![License: MIT](https://img.shields.io/badge/License-MIT-green?style=flat-square)](LICENSE)
[![Version](https://img.shields.io/badge/Version-2.1.0-brightgreen?style=flat-square)](https://github.com/ayinedjimi/sharepoint-onedrive-share-audit/releases)
[![PnP.PowerShell 2.0+](https://img.shields.io/badge/PnP.PowerShell-2.0%2B-orange?style=flat-square)](https://pnp.github.io/powershell/)
[![Microsoft Graph](https://img.shields.io/badge/Microsoft%20Graph-v1.0-00a1f1?style=flat-square&logo=microsoft)](https://docs.microsoft.com/graph/)
[![Bugs Fixed: 10](https://img.shields.io/badge/Bugs%20Fixed-10-red?style=flat-square)](https://github.com/ayinedjimi/sharepoint-onedrive-share-audit/blob/main/CHANGELOG.md)
[![Blog](https://img.shields.io/badge/Blog-ayinedjimi--consultants.fr-purple?style=flat-square)](https://ayinedjimi-consultants.fr)

<br/>

**Audit automatisé et exhaustif de tous les partages SharePoint Online & OneDrive for Business d'un tenant Microsoft 365.**  
Détection des liens anonymes, export CSV/HTML interactif, prêt pour la production en PS 5.1+.

<br/>

[📖 Guide complet FR](https://ayinedjimi-consultants.fr/articles/audit-partages-sharepoint-onedrive-powershell-anonymes) &nbsp;|&nbsp; [🐛 Signaler un bug](https://github.com/ayinedjimi/sharepoint-onedrive-share-audit/issues) &nbsp;|&nbsp; [🔐 Audit M365 complet](https://ayinedjimi-consultants.fr/articles/audit-securite-microsoft-365-guide) &nbsp;|&nbsp; [📚 Documentation](https://docs.microsoft.com/sharepoint/)

</div>

---

<details>
<summary>📋 Table des matières / Table of Contents</summary>

**🇫🇷 Français**
- [Présentation](#-présentation)
- [Architecture & Flux](#-architecture--flux)
- [Fonctionnalités](#-fonctionnalités-clés)
- [Prérequis](#-prérequis)
- [Installation](#-installation)
- [Utilisation](#-utilisation)
- [Format des sorties](#-format-des-sorties)
- [Révoquer les liens anonymes](#-révoquer-les-liens-anonymes)
- [Dépannage](#-dépannage)
- [Ressources complémentaires](#-ressources-complémentaires)

**🇬🇧 English**
- [Overview](#-overview)
- [Architecture & Flow](#-architecture--flow)
- [Features](#-key-features)
- [Prerequisites](#-prerequisites)
- [Installation](#-installation-1)
- [Usage](#-usage)
- [Output Format](#-output-format)
- [Revoking Anonymous Links](#-revoking-anonymous-links)
- [Troubleshooting](#-troubleshooting)
- [Further Reading](#-further-reading)

**📋 Général**
- [Changelog](#-changelog)
- [Contribuer / Contributing](#-contribuer--contributing)
- [License](#-license)

</details>

---

## 🇫🇷 Documentation Française

### 🔍 Présentation

`Get-ShareAudit.ps1` est un script PowerShell **prêt pour la production** qui réalise un audit exhaustif de l'ensemble des liens de partage d'un tenant **Microsoft 365** : SharePoint Online, OneDrive for Business, et tous les sous-sites.

Conçu pour les **équipes sécurité**, les **RSSI** et les **auditeurs conformité** (RGPD, NIS 2, ISO 27001), il cible en priorité les **liens anonymes** — fichiers accessibles **sans authentification** — qui représentent le risque de fuite de données le plus critique dans un environnement M365.

> 💡 Ce script a été développé à partir des méthodologies décrites dans notre [guide d'audit de sécurité Microsoft 365](https://ayinedjimi-consultants.fr/articles/audit-securite-microsoft-365-guide).

---

### 🏗️ Architecture & Flux

```
┌─────────────────────────────────────────────────────────────────────────┐
│                    Get-ShareAudit.ps1 — Flux d'exécution                │
└─────────────────────────────────────────────────────────────────────────┘

  Paramètres                Module disponible ?
  TenantId ──────────────── PnP.PowerShell 2.0+ ──────────────── SharePoint API
  ClientId                      │                                        │
  ClientSecret              Non disponible ?                             │
                             └── Microsoft.Graph ────────────── Graph API v1.0
                                                                         │
                              ┌──────────────────────────────────────────┘
                              │
                              ▼
                    ┌─────────────────┐
                    │ Enumération     │
                    │ • SharePoint    │◄─── Tous les sites du tenant
                    │ • OneDrive      │◄─── MySites (si -IncludeOneDrive)
                    └────────┬────────┘
                             │ Pour chaque site
                             ▼
                    ┌─────────────────┐
                    │ Analyse items   │◄─── CAML Query paginée
                    │ • Fichiers      │     (do-while + ListItemCollectionPosition)
                    │ • Dossiers      │
                    │ • Permissions   │◄─── REST API SharePoint
                    └────────┬────────┘     /_api/.../SharingLinks
                             │
                    ┌────────▼────────┐
                    │ Filtrage        │
                    │ • AnonymousOnly │◄─── -AnonymousOnly switch
                    │ • LinkKind 4/5  │     (Anyone View/Edit)
                    └────────┬────────┘
                             │
               ┌─────────────┼─────────────┐
               ▼             ▼             ▼
          CSV Export    HTML Report    Console Summary
          (RFC 4180)    (interactif)   (Top 10 items)
```

---

### ✅ Fonctionnalités clés

| Fonctionnalité | Détail |
|---|---|
| 🔍 **Énumération complète** | Tous les sites SharePoint + OneDrive for Business, récursif |
| 🔴 **Détection liens anonymes** | Anyone View (LinkKind 4) et Anyone Edit (LinkKind 5) mis en évidence |
| 📊 **Export CSV** | Format RFC 4180, séparateur `,`, UTF-8, prêt pour Excel et Power BI |
| 📄 **Rapport HTML** | Tableau interactif : tri par colonne, filtre temps réel, code couleur |
| 🔁 **Retry automatique** | Back-off exponentiel (1s → 2s → 4s → 8s) sur HTTP 429 throttling |
| ⚡ **Fast-fail** | Arrêt immédiat sur HTTP 401/403 (pas de retry inutile) |
| 📈 **Barre de progression** | `Write-Progress` sur chaque site : sites traités / total |
| 🔀 **Double moteur** | PnP.PowerShell 2.0+ (préféré) → Microsoft.Graph (fallback auto) |
| 📑 **Pagination CAML** | Boucle `do-while` sur `ListItemCollectionPosition` — aucun item manqué |
| 🏁 **Résumé final** | Top 10 des items les plus exposés affichés en console |
| 🛡️ **XSS-safe** | `[System.Web.HttpUtility]::HtmlEncode()` systématique dans le rapport |
| 🔧 **PS 5.1 compatible** | Aucun opérateur PS 7+ (`??`, `?:`, `?.`) — tourne sur Windows 10+ natif |

---

### 📋 Prérequis

#### Modules PowerShell

```powershell
# Option 1 : PnP.PowerShell (recommandé)
Install-Module PnP.PowerShell -MinimumVersion 2.0 -Scope CurrentUser -Force

# Option 2 : Microsoft.Graph (fallback automatique si PnP absent)
Install-Module Microsoft.Graph -Scope CurrentUser -Force

# Vérifier les versions installées
Get-Module PnP.PowerShell -ListAvailable | Select Name, Version
Get-Module Microsoft.Graph -ListAvailable | Select Name, Version
```

> ⚠️ Le script **ne requiert pas** `#Requires -Modules` — il détecte dynamiquement le module disponible au runtime et bascule automatiquement sur le fallback.

**Versions compatibles :**
| OS | PowerShell | Support |
|---|---|---|
| Windows 10/11 | PS 5.1 (natif) | ✅ Complet |
| Windows Server 2019/2022 | PS 5.1 | ✅ Complet |
| Windows 10/11 | PS 7.x | ✅ Complet |
| Linux/macOS | PS 7.x | ✅ Via Graph fallback |

#### Enregistrement de l'application Azure AD

1. Accéder au [portail Azure](https://portal.azure.com) → **Azure Active Directory** → **Inscriptions d'applications**
2. Cliquer **Nouvelle inscription** → nommer l'app (ex : `M365-ShareAudit`)
3. Ajouter les **autorisations d'application** (Application permissions — **pas** déléguées) :

| API | Permission | Justification |
|-----|-----------|---------------|
| Microsoft Graph | `Sites.Read.All` | Lire tous les sites SharePoint |
| Microsoft Graph | `Files.Read.All` | Lire tous les fichiers et permissions |
| Microsoft Graph | `User.Read.All` | Énumérer les utilisateurs pour OneDrive |
| SharePoint | `Sites.FullControl.All` | Requis pour PnP.PowerShell app-only |

4. Cliquer **Accorder le consentement administrateur** (Global Admin requis)
5. **Certificates & Secrets** → Créer un secret client (noter la valeur immédiatement)

> 🔐 **Bonne pratique sécurité :** Préférez un **certificat client** à un secret texte en production, et limitez la durée de vie à 90 jours max. Voir notre [guide de configuration Entra ID](https://ayinedjimi-consultants.fr/articles/entra-id-azure-ad-securite-configuration).

#### Rôles Azure AD requis

| Rôle | Suffisant pour |
|---|---|
| SharePoint Administrator | Mode PnP app-only complet |
| Global Administrator | Tout (y compris consentement admin) |
| Global Reader | Mode Graph lecture seule uniquement |

---

### 📦 Installation

```bash
# Cloner le dépôt
git clone https://github.com/ayinedjimi/sharepoint-onedrive-share-audit.git
cd sharepoint-onedrive-share-audit

# Débloquer le script (Windows — contourne la restriction d'exécution de zone internet)
Unblock-File .\Get-ShareAudit.ps1

# Vérifier la signature du script (optionnel)
Get-AuthenticodeSignature .\Get-ShareAudit.ps1
```

Ou téléchargement direct :
```powershell
Invoke-WebRequest -Uri "https://raw.githubusercontent.com/ayinedjimi/sharepoint-onedrive-share-audit/main/Get-ShareAudit.ps1" -OutFile "Get-ShareAudit.ps1"
Unblock-File .\Get-ShareAudit.ps1
```

---

### 🚀 Utilisation

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

#### Mode diagnostic verbose (debug)

```powershell
$VerbosePreference = 'Continue'
.\Get-ShareAudit.ps1 `
    -TenantId "contoso.onmicrosoft.com" `
    -ClientId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -ClientSecret "votre_secret_ici" `
    -Verbose
```

#### Exemple de sortie console

```
════════════════════════════════════════════════════════════════
  Get-ShareAudit.ps1 v2.1.0 — Audit Partages M365
  ayinedjimi-consultants.fr — Cybersécurité Microsoft 365
════════════════════════════════════════════════════════════════

[2026-05-21 09:12:03] [INFO]    Vérification des modules disponibles...
[2026-05-21 09:12:04] [INFO]    Module PnP.PowerShell 2.3.0 détecté.
[2026-05-21 09:12:04] [INFO]    Connexion en mode application (app-only)...
[2026-05-21 09:12:06] [SUCCESS] Connecté via PnP.PowerShell à https://contoso-admin.sharepoint.com
[2026-05-21 09:12:08] [SUCCESS] 47 site(s) SharePoint trouvé(s).
[2026-05-21 09:12:09] [SUCCESS] 312 OneDrive trouvé(s) (filtrage client-side).
[2026-05-21 09:12:10] [INFO]    Analyse en cours... (359 sites au total)
...
[2026-05-21 09:45:12] [SUCCESS] Analyse terminée en 33m 09s.
[2026-05-21 09:45:13] [INFO]    Export CSV : C:\Audits\M365\20260521\ShareAudit_20260521.csv
[2026-05-21 09:45:14] [INFO]    Rapport HTML : C:\Audits\M365\20260521\ShareAudit_20260521.html

════════ TOP 10 — ITEMS LES PLUS EXPOSÉS ════════
  1. brief-Q3-2026.pdf          (Equipe-Marketing)   [ANYONE] 1 lien
  2. contrat-NDA-fournisseur.docx (John.Smith-OD)    [ANYONE] 1 lien
  ...
═════════════════════════════════════════════════

[2026-05-21 09:45:14] [INFO]    Résumé : 359 sites, 48 291 items, 1 847 liens, 23 ANONYMES
```

---

### 📊 Format des sorties

#### Fichier CSV

| Colonne | Description | Exemple |
|---------|-------------|---------|
| `SiteUrl` | URL du site SharePoint/OneDrive | `https://contoso.sharepoint.com/sites/Marketing` |
| `SiteType` | Type de site | `SharePoint` / `OneDrive` |
| `ItemType` | Type d'élément | `File` / `Folder` |
| `ItemName` | Nom du fichier/dossier | `brief-Q3-2026.pdf` |
| `ItemUrl` | URL relative de l'élément | `/sites/Marketing/Shared%20Documents/brief.pdf` |
| `ShareType` | Type de partage | `Anyone` / `Company` / `Specific` / `Inherited` |
| `ShareLink` | URL du lien de partage | `https://contoso.sharepoint.com/:b:/s/...` |
| `CreatedBy` | Créateur du lien | `john.smith@contoso.com` |
| `ExpirationDate` | Date d'expiration (si définie) | `2026-12-31` |
| `LastModified` | Dernière modification | `2026-05-15T14:23:01` |

#### Rapport HTML

Le rapport HTML inclut :
- **En-tête statistiques** : sites analysés, items scannés, liens totaux, liens anonymes
- **Tableau filtrable** : tri par n'importe quelle colonne, filtre temps réel
- **Code couleur** : rouge = Anyone, orange = Company, vert = Specific
- **Section ressources** : liens vers guides de remédiation

---

### 🔧 Paramètres de référence

| Paramètre | Type | Obligatoire | Description |
|-----------|------|:-----------:|-------------|
| `-TenantId` | String | ✅ | Domaine (`contoso.onmicrosoft.com`) ou GUID du tenant |
| `-ClientId` | String | ✅ | Application (client) ID — format GUID |
| `-ClientSecret` | String | ✅ | Secret client de l'application Azure AD |
| `-OutputPath` | String | ❌ | Dossier de sortie (défaut : répertoire courant) |
| `-IncludeOneDrive` | Switch | ❌ | Inclure les MySites OneDrive de tous les utilisateurs |
| `-AnonymousOnly` | Switch | ❌ | Rapporter uniquement les liens "Anyone" (anonymes) |
| `-Verbose` | Switch | ❌ | Afficher les messages DEBUG détaillés |

---

### 🛡️ Révoquer les liens anonymes identifiés

Les liens de partage SharePoint sont gérés via l'**API REST SharePoint** (via `Invoke-PnPSPRestMethod`). La cmdlet `Get-PnPFileSharingLink` n'existe plus dans PnP.PowerShell 2.0+.

```powershell
# 1. Connexion au site concerné
Connect-PnPOnline -Url "https://contoso.sharepoint.com/sites/Marketing" `
                  -ClientId $ClientId `
                  -ClientSecret $ClientSecret `
                  -TenantId $TenantId

# 2. Récupérer tous les liens d'un fichier via l'API REST
$fileUrl = "/sites/Marketing/Shared Documents/brief-Q3.pdf"
$restEndpoint = "/_api/web/GetFileByServerRelativeUrl('{0}')/ListItemAllFields/SharingLinks" `
    -f [uri]::EscapeDataString($fileUrl)

$links = Invoke-PnPSPRestMethod -Url $restEndpoint -Method Get

# 3. Identifier et supprimer les liens anonymes
# LinkKind 4 = AnonymousView | LinkKind 5 = AnonymousEdit
$anonymousLinks = $links.value | Where-Object { $_.LinkKind -eq 4 -or $_.LinkKind -eq 5 }

foreach ($link in $anonymousLinks) {
    Write-Host "Suppression du lien anonyme : $($link.Url)" -ForegroundColor Yellow
    # Implémentation de la suppression via Graph ou API REST admin
    # Voir : https://ayinedjimi-consultants.fr/articles/audit-securite-microsoft-365-guide
}

Write-Host "$($anonymousLinks.Count) lien(s) anonyme(s) identifié(s) sur ce fichier."
```

> 📖 Procédures de remédiation détaillées : [Guide d'audit sécurité Microsoft 365](https://ayinedjimi-consultants.fr/articles/audit-securite-microsoft-365-guide)

---

### 🔬 Dépannage

<details>
<summary><strong>❌ Erreur : "No module PnP.PowerShell or Microsoft.Graph found"</strong></summary>

Aucun module compatible n'est installé. Installez-en au moins un :
```powershell
Install-Module PnP.PowerShell -MinimumVersion 2.0 -Scope CurrentUser -Force
# OU
Install-Module Microsoft.Graph -Scope CurrentUser -Force
```
</details>

<details>
<summary><strong>❌ Erreur : HTTP 401 Unauthorized lors de la connexion</strong></summary>

1. Vérifiez que le `ClientSecret` est correct et non expiré (Azure Portal → App registrations → Certificates & secrets)
2. Vérifiez que le **consentement administrateur** a été accordé pour toutes les permissions
3. Pour PnP.PowerShell, vérifiez que `Sites.FullControl.All` est accordé côté SharePoint (pas seulement Graph)
</details>

<details>
<summary><strong>❌ Erreur : HTTP 429 Too Many Requests (throttling)</strong></summary>

Le script intègre un back-off exponentiel automatique (1s → 2s → 4s → 8s, max 3 tentatives). Si le problème persiste :
- Réduisez la charge en auditant un sous-ensemble de sites
- Planifiez l'exécution en dehors des heures de bureau
- Voir la [politique de throttling SharePoint](https://docs.microsoft.com/sharepoint/dev/general-development/how-to-avoid-getting-throttled-or-blocked-in-sharepoint-online)
</details>

<details>
<summary><strong>❌ Erreur : HTTP 403 Access Denied sur certains sites</strong></summary>

Certains sites peuvent avoir des paramètres d'accès restrictifs bloquant même l'accès app-only. Le script les journalise et continue. Vérifiez que `Sites.FullControl.All` est accordé pour toutes les collections de sites.
</details>

<details>
<summary><strong>⚠️ OneDrive : aucun résultat alors que -IncludeOneDrive est passé</strong></summary>

Le filtrage OneDrive utilise `Where-Object { $_.Url -like '*-my.sharepoint.com/personal/*' }` côté client. Si votre tenant a un domaine MySite personnalisé, modifiez le pattern en conséquence. Exemple : `*-my.sharepoint.com/*`.
</details>

<details>
<summary><strong>⚠️ Rapport HTML corrompu (caractères spéciaux)</strong></summary>

Le script encode systématiquement le contenu avec `[System.Web.HttpUtility]::HtmlEncode()`. Si vous voyez des caractères étranges, vérifiez que le fichier HTML est ouvert en UTF-8 dans votre navigateur.
</details>

<details>
<summary><strong>⏱️ Performances : temps estimé selon la taille du tenant</strong></summary>

| Sites SharePoint | OneDrive inclus | Temps estimé |
|:-:|:-:|:-:|
| < 10 | Non | 2-5 min |
| 10-50 | Non | 5-20 min |
| 50-200 | Non | 20-60 min |
| 50-200 | Oui (+300 OD) | 1-3 heures |
| 200+ | Oui | 3+ heures |

*Estimations pour un tenant avec ~100 items par bibliothèque en moyenne.*
</details>

---

### 📚 Ressources complémentaires

Ces articles du blog [ayinedjimi-consultants.fr](https://ayinedjimi-consultants.fr) approfondissent les thématiques couvertes par ce script :

| Article | Sujet | Pertinence |
|---------|-------|:-----------:|
| [🔐 Audit complet de sécurité Microsoft 365](https://ayinedjimi-consultants.fr/articles/audit-securite-microsoft-365-guide) | Méthodologie d'audit M365 de A à Z, remédiation | ⭐⭐⭐ |
| [🔎 Forensique M365 — Unified Audit Log](https://ayinedjimi-consultants.fr/articles/forensique-microsoft-365-unified-audit-log) | Investigation et analyse des logs d'audit M365 | ⭐⭐⭐ |
| [🛡️ Entra ID / Azure AD — Sécurisation](https://ayinedjimi-consultants.fr/articles/entra-id-azure-ad-securite-configuration) | Hardening Entra ID, certificats vs secrets | ⭐⭐⭐ |
| [📋 Identity Governance — Cycle de vie des comptes](https://ayinedjimi-consultants.fr/articles/identity-governance-iga-cycle-vie-comptes) | Gestion des identités, revues d'accès | ⭐⭐ |
| [🔍 Threat Hunting M365 avec Sentinel](https://ayinedjimi-consultants.fr/articles/threat-hunting-microsoft-365-sentinel) | Détection des menaces dans les environnements M365 | ⭐⭐ |
| [📖 Guide complet — Audit partages SP/OD](https://ayinedjimi-consultants.fr/articles/audit-partages-sharepoint-onedrive-powershell-anonymes) | Article de référence accompagnant ce script | ⭐⭐⭐ |

---

## 🇬🇧 English Documentation

### 🔍 Overview

`Get-ShareAudit.ps1` is a **production-ready** PowerShell script performing an exhaustive audit of all sharing links across a **Microsoft 365 tenant**: SharePoint Online, OneDrive for Business, and all sub-sites.

Built for **security teams**, **CISOs**, and **compliance auditors** (GDPR, NIS 2, ISO 27001), it focuses on detecting **anonymous links** — files accessible **without any authentication** — the most critical data exfiltration risk in an M365 environment.

> 💡 This script implements the methodology described in our [Microsoft 365 security audit guide](https://ayinedjimi-consultants.fr/articles/audit-securite-microsoft-365-guide) (French).

---

### 🏗️ Architecture & Flow

```
┌─────────────────────────────────────────────────────────────────────────┐
│                   Get-ShareAudit.ps1 — Execution Flow                   │
└─────────────────────────────────────────────────────────────────────────┘

  Parameters                Module available?
  TenantId ──────────────── PnP.PowerShell 2.0+ ──────────────── SharePoint API
  ClientId                      │                                        │
  ClientSecret              Not available?                               │
                             └── Microsoft.Graph ────────────── Graph API v1.0
                                                                         │
                              ┌──────────────────────────────────────────┘
                              │
                              ▼
                    ┌─────────────────┐
                    │ Enumeration     │
                    │ • SharePoint    │◄─── All tenant sites
                    │ • OneDrive      │◄─── MySites (if -IncludeOneDrive)
                    └────────┬────────┘
                             │ Per site
                             ▼
                    ┌─────────────────┐
                    │ Item Analysis   │◄─── Paginated CAML Query
                    │ • Files         │     (do-while + ListItemCollectionPosition)
                    │ • Folders       │
                    │ • Permissions   │◄─── SharePoint REST API
                    └────────┬────────┘     /_api/.../SharingLinks
                             │
                    ┌────────▼────────┐
                    │ Filtering       │
                    │ • AnonymousOnly │◄─── -AnonymousOnly switch
                    │ • LinkKind 4/5  │     (Anyone View/Edit)
                    └────────┬────────┘
                             │
               ┌─────────────┼─────────────┐
               ▼             ▼             ▼
          CSV Export    HTML Report    Console Summary
          (RFC 4180)  (interactive)   (Top 10 items)
```

---

### ✅ Key Features

| Feature | Details |
|---|---|
| 🔍 **Full enumeration** | All SharePoint sites + OneDrive for Business, recursive |
| 🔴 **Anonymous link detection** | Anyone View (LinkKind 4) and Anyone Edit (LinkKind 5) highlighted |
| 📊 **CSV export** | RFC 4180 format, comma separator, UTF-8, Excel & Power BI ready |
| 📄 **HTML report** | Interactive: sortable columns, real-time filter, color coding |
| 🔁 **Auto-retry** | Exponential back-off (1s → 2s → 4s → 8s) on HTTP 429 throttling |
| ⚡ **Fast-fail** | Immediate stop on HTTP 401/403 (no useless retries) |
| 📈 **Progress bar** | `Write-Progress` per site: sites processed / total |
| 🔀 **Dual engine** | PnP.PowerShell 2.0+ (preferred) → Microsoft.Graph (auto-fallback) |
| 📑 **CAML pagination** | `do-while` loop on `ListItemCollectionPosition` — no items missed |
| 🏁 **Final summary** | Top 10 most exposed items displayed in console |
| 🛡️ **XSS-safe** | `[System.Web.HttpUtility]::HtmlEncode()` throughout the report |
| 🔧 **PS 5.1 compatible** | No PS 7+ operators (`??`, `?:`, `?.`) — runs natively on Windows 10+ |

---

### 📋 Prerequisites

#### PowerShell Modules

```powershell
# Option 1: PnP.PowerShell (recommended)
Install-Module PnP.PowerShell -MinimumVersion 2.0 -Scope CurrentUser -Force

# Option 2: Microsoft.Graph (automatic fallback if PnP not available)
Install-Module Microsoft.Graph -Scope CurrentUser -Force

# Verify installed versions
Get-Module PnP.PowerShell -ListAvailable | Select Name, Version
Get-Module Microsoft.Graph -ListAvailable | Select Name, Version
```

> ⚠️ The script does **not** use `#Requires -Modules` — it detects available modules dynamically at runtime and switches to the fallback automatically.

**Compatibility matrix:**
| OS | PowerShell | Support |
|---|---|---|
| Windows 10/11 | PS 5.1 (built-in) | ✅ Full |
| Windows Server 2019/2022 | PS 5.1 | ✅ Full |
| Windows 10/11 | PS 7.x | ✅ Full |
| Linux/macOS | PS 7.x | ✅ Via Graph fallback |

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

4. Click **Grant admin consent** (requires Global Admin)
5. **Certificates & Secrets** → Create a client secret (copy the value immediately)

> 🔐 **Security tip:** Use a **client certificate** instead of a plain-text secret in production, with a max 90-day lifetime. See our [Entra ID security guide](https://ayinedjimi-consultants.fr/articles/entra-id-azure-ad-securite-configuration) (French).

#### Required Azure AD Roles

| Role | Sufficient for |
|---|---|
| SharePoint Administrator | Full PnP app-only mode |
| Global Administrator | Everything (including admin consent) |
| Global Reader | Graph read-only mode only |

---

### 📦 Installation

```bash
# Clone the repository
git clone https://github.com/ayinedjimi/sharepoint-onedrive-share-audit.git
cd sharepoint-onedrive-share-audit

# Unblock the script (Windows — bypasses internet zone execution restriction)
Unblock-File .\Get-ShareAudit.ps1
```

Or direct download:
```powershell
Invoke-WebRequest -Uri "https://raw.githubusercontent.com/ayinedjimi/sharepoint-onedrive-share-audit/main/Get-ShareAudit.ps1" -OutFile "Get-ShareAudit.ps1"
Unblock-File .\Get-ShareAudit.ps1
```

---

### 🚀 Usage

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

#### Verbose/debug mode

```powershell
$VerbosePreference = 'Continue'
.\Get-ShareAudit.ps1 `
    -TenantId "contoso.onmicrosoft.com" `
    -ClientId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -ClientSecret "your_secret_here" `
    -Verbose
```

---

### 📊 Output Format

#### CSV File

| Column | Description | Example |
|--------|-------------|---------|
| `SiteUrl` | SharePoint/OneDrive site URL | `https://contoso.sharepoint.com/sites/Marketing` |
| `SiteType` | Site type | `SharePoint` / `OneDrive` |
| `ItemType` | Element type | `File` / `Folder` |
| `ItemName` | File/folder name | `brief-Q3-2026.pdf` |
| `ItemUrl` | Relative element URL | `/sites/Marketing/Shared%20Documents/brief.pdf` |
| `ShareType` | Sharing type | `Anyone` / `Company` / `Specific` / `Inherited` |
| `ShareLink` | Sharing link URL | `https://contoso.sharepoint.com/:b:/s/...` |
| `CreatedBy` | Link creator | `john.smith@contoso.com` |
| `ExpirationDate` | Expiration date (if set) | `2026-12-31` |
| `LastModified` | Last modification | `2026-05-15T14:23:01` |

---

### 📋 Parameters Reference

| Parameter | Type | Required | Description |
|-----------|------|:--------:|-------------|
| `-TenantId` | String | ✅ | Tenant domain (`contoso.onmicrosoft.com`) or GUID |
| `-ClientId` | String | ✅ | Application (client) ID — GUID format |
| `-ClientSecret` | String | ✅ | Azure AD app client secret value |
| `-OutputPath` | String | ❌ | Output folder (default: current directory) |
| `-IncludeOneDrive` | Switch | ❌ | Include all users' OneDrive MySites |
| `-AnonymousOnly` | Switch | ❌ | Report only "Anyone" anonymous links |
| `-Verbose` | Switch | ❌ | Show detailed DEBUG messages |

---

### 🛡️ Revoking Anonymous Links

Sharing links are managed via the **SharePoint REST API** (via `Invoke-PnPSPRestMethod`). The former `Get-PnPFileSharingLink` cmdlet was removed from PnP.PowerShell 2.0+:

```powershell
# Connect to the affected site
Connect-PnPOnline -Url "https://contoso.sharepoint.com/sites/Marketing" `
                  -ClientId $ClientId `
                  -ClientSecret $ClientSecret `
                  -TenantId $TenantId

# Get sharing links via SharePoint REST API
$fileUrl = "/sites/Marketing/Shared Documents/brief-Q3.pdf"
$restEndpoint = "/_api/web/GetFileByServerRelativeUrl('{0}')/ListItemAllFields/SharingLinks" `
    -f [uri]::EscapeDataString($fileUrl)

$links = Invoke-PnPSPRestMethod -Url $restEndpoint -Method Get

# Identify anonymous links (LinkKind 4 = AnonymousView, 5 = AnonymousEdit)
$anonymousLinks = $links.value | Where-Object { $_.LinkKind -eq 4 -or $_.LinkKind -eq 5 }
Write-Host "Found $($anonymousLinks.Count) anonymous link(s) on this file."
```

> 📖 See our [M365 security audit guide](https://ayinedjimi-consultants.fr/articles/audit-securite-microsoft-365-guide) for detailed remediation procedures (French).

---

### 🔬 Troubleshooting

<details>
<summary><strong>❌ "No module PnP.PowerShell or Microsoft.Graph found"</strong></summary>

No compatible module is installed. Install at least one:
```powershell
Install-Module PnP.PowerShell -MinimumVersion 2.0 -Scope CurrentUser -Force
# OR
Install-Module Microsoft.Graph -Scope CurrentUser -Force
```
</details>

<details>
<summary><strong>❌ HTTP 401 Unauthorized on connect</strong></summary>

1. Verify the `ClientSecret` is correct and not expired (Azure Portal → App registrations → Certificates & secrets)
2. Verify **admin consent** was granted for all permissions
3. For PnP.PowerShell, ensure `Sites.FullControl.All` is granted on the SharePoint side (not just Graph)
</details>

<details>
<summary><strong>❌ HTTP 429 Too Many Requests (throttling)</strong></summary>

The script includes automatic exponential back-off (1s → 2s → 4s → 8s, max 3 attempts). If the issue persists, schedule execution outside business hours or audit a subset of sites.
</details>

<details>
<summary><strong>⏱️ Performance estimates by tenant size</strong></summary>

| SharePoint Sites | OneDrive included | Estimated time |
|:-:|:-:|:-:|
| < 10 | No | 2-5 min |
| 10-50 | No | 5-20 min |
| 50-200 | No | 20-60 min |
| 50-200 | Yes (+300 OD) | 1-3 hours |
| 200+ | Yes | 3+ hours |

*Estimates for a tenant averaging ~100 items per library.*
</details>

---

### 📚 Further Reading

These articles from [ayinedjimi-consultants.fr](https://ayinedjimi-consultants.fr) (in French) cover topics related to this script:

| Article | Topic | Relevance |
|---------|-------|:---------:|
| [🔐 Microsoft 365 Security Audit Guide](https://ayinedjimi-consultants.fr/articles/audit-securite-microsoft-365-guide) | Complete M365 audit methodology, remediation | ⭐⭐⭐ |
| [🔎 M365 Forensics — Unified Audit Log](https://ayinedjimi-consultants.fr/articles/forensique-microsoft-365-unified-audit-log) | Investigation and audit log analysis | ⭐⭐⭐ |
| [🛡️ Entra ID / Azure AD Security](https://ayinedjimi-consultants.fr/articles/entra-id-azure-ad-securite-configuration) | Hardening Entra ID, certificates vs secrets | ⭐⭐⭐ |
| [📋 Identity Governance — Account Lifecycle](https://ayinedjimi-consultants.fr/articles/identity-governance-iga-cycle-vie-comptes) | Enterprise identity and access management | ⭐⭐ |
| [🔍 Threat Hunting M365 with Sentinel](https://ayinedjimi-consultants.fr/articles/threat-hunting-microsoft-365-sentinel) | Threat detection in M365 environments | ⭐⭐ |
| [📖 SharePoint/OneDrive Share Audit Guide](https://ayinedjimi-consultants.fr/articles/audit-partages-sharepoint-onedrive-powershell-anonymes) | Reference article for this script | ⭐⭐⭐ |

---

## 📋 Changelog

### v2.1.0 — 2026-05-21

**10 bugs corrigés / 10 bugs fixed :**

| # | Severity | Bug | Fix |
|---|:--------:|-----|-----|
| 1 | 🔴 CRITICAL | `#Requires -Modules` blocking Graph fallback | Removed — dynamic runtime detection |
| 2 | 🔴 CRITICAL | `??` null-coalescing operator (PS 7+ only) | Replaced with `Get-ValueOrDefault` PS 5.1 helper |
| 3 | 🔴 CRITICAL | Ternary `? :` operator (PS 7+ only) | Replaced with explicit `if/else` |
| 4 | 🔴 CRITICAL | `Get-PnPFileSharingLink` doesn't exist in PnP 2.0+ | Replaced with `Invoke-PnPSPRestMethod` REST API |
| 5 | 🟠 HIGH | `Connect-MgGraph -Scopes` invalid for app-only | Removed — `ClientSecretCredential` doesn't take `-Scopes` |
| 6 | 🟠 HIGH | `Get-PnPTenantSite -Filter` invalid CAML syntax | Client-side filter with `Where-Object` |
| 7 | 🟠 HIGH | No CAML pagination on large lists | `do-while` loop with `ListItemCollectionPosition` |
| 8 | 🟠 HIGH | Graph site lookup URL double-encoding | URI correctly built with `hostname:path` |
| 9 | 🟡 MEDIUM | `[datetime]` cast throws on invalid format | Replaced with `[datetime]::TryParse()` |
| 10 | 🟡 MEDIUM | DEBUG log operator precedence bug | Explicit parentheses added |

**Other improvements:**
- CSV separator: `,` (RFC 4180 standard) instead of `;`
- HTTP 401/403 fast-fail in `Invoke-WithRetry` (avoids useless retries)
- Systematic HTML encoding in report (`[System.Web.HttpUtility]::HtmlEncode`)
- Links to reference articles in `.NOTES` and HTML report resources section

---

## 🤝 Contribuer / Contributing

Les contributions sont les bienvenues / Contributions are welcome:

1. **Fork** le dépôt / Fork the repository
2. Créez une branche / Create a branch: `git checkout -b feature/ma-fonctionnalite`
3. Commitez vos changements / Commit your changes: `git commit -m 'feat: description'`
4. Poussez vers la branche / Push: `git push origin feature/ma-fonctionnalite`
5. Ouvrez une **Pull Request**

**Tests requis / Required tests:**
- Testé sur PS 5.1 (Windows 10/11) / Tested on PS 5.1
- Testé avec PnP.PowerShell 2.x / Tested with PnP.PowerShell 2.x
- Testé avec Microsoft.Graph / Tested with Microsoft.Graph

**Rapport de bug / Bug report:**
Utilisez le [template de bug](.github/ISSUE_TEMPLATE/bug_report.md) / Use the [bug report template](.github/ISSUE_TEMPLATE/bug_report.md)

---

## ⚖️ Avertissement légal / Legal Notice

> ⚠️ **FR : Ce script est destiné exclusivement aux administrateurs et auditeurs disposant des autorisations légales et contractuelles pour accéder au tenant Microsoft 365 concerné. Toute utilisation sur un tenant tiers sans autorisation explicite est illégale et peut engager votre responsabilité pénale.**
>
> ⚠️ **EN: This script is intended exclusively for administrators and auditors with the legal and contractual authorizations to access the Microsoft 365 tenant in question. Any unauthorized use on a third-party tenant is illegal and may engage your criminal liability.**

---

## 📄 License

```
MIT License — Copyright (c) 2026 Ayi NEDJIMI

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
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

---

<div align="center">

<br/>

Made with ❤️ by **[Ayi NEDJIMI](https://ayinedjimi-consultants.fr)** — Cybersécurité & Cloud Microsoft 365

[![Blog](https://img.shields.io/badge/Blog-ayinedjimi--consultants.fr-purple?style=for-the-badge)](https://ayinedjimi-consultants.fr)
[![Guide SharePoint](https://img.shields.io/badge/Guide%20SharePoint-Audit%20Partages-blue?style=for-the-badge)](https://ayinedjimi-consultants.fr/articles/audit-partages-sharepoint-onedrive-powershell-anonymes)
[![Audit M365](https://img.shields.io/badge/Audit%20M365-Guide%20Complet-red?style=for-the-badge)](https://ayinedjimi-consultants.fr/articles/audit-securite-microsoft-365-guide)

<br/>

⭐ **Si ce script vous a été utile, n'hésitez pas à star le dépôt / Star this repo if it helped you!** ⭐

</div>
