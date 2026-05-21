#Requires -Version 5.1
# NOTE : PnP.PowerShell ou Microsoft.Graph requis — détection automatique au démarrage.
# Pour installer : Install-Module PnP.PowerShell -MinimumVersion 2.0 -Scope CurrentUser
#                  Install-Module Microsoft.Graph -Scope CurrentUser

<#
.SYNOPSIS
    Audit complet des partages SharePoint Online et OneDrive for Business dans un tenant Microsoft 365.

.DESCRIPTION
    Get-ShareAudit.ps1 se connecte à Microsoft 365 via PnP.PowerShell (avec fallback vers
    Microsoft.Graph) et réalise un inventaire exhaustif de tous les liens de partage :
      - SharePoint Online : tous les sites et sous-sites
      - OneDrive for Business : les MySites de chaque utilisateur

    Pour chaque lien de partage détecté, le script extrait :
      - Le type de partage (Anyone / Company / Specific)
      - Le créateur du lien et sa date de création
      - La date d'expiration (si définie)
      - L'URL du lien et l'URL de l'élément partagé

    Les partages "Anyone" (liens anonymes, sans authentification) sont mis en évidence
    car ils représentent le risque de fuite de données le plus élevé.

    Exports disponibles : console colorée, CSV et rapport HTML interactif.

.PARAMETER TenantId
    Domaine du tenant Microsoft 365 (ex: contoso.onmicrosoft.com).
    IMPORTANT : Fournir le domaine, pas le GUID, pour la résolution des URLs SharePoint.

.PARAMETER ClientId
    Application ID (Client ID) de l'application Azure AD enregistrée (format GUID).

.PARAMETER ClientSecret
    Secret client de l'application Azure AD.

.PARAMETER OutputPath
    Dossier de destination pour les exports CSV et HTML. Par défaut : répertoire courant.

.PARAMETER IncludeOneDrive
    Switch : inclure les OneDrive for Business de tous les utilisateurs.

.PARAMETER AnonymousOnly
    Switch : ne rapporter que les liens "Anyone" (partages anonymes).

.EXAMPLE
    .\Get-ShareAudit.ps1 -TenantId "contoso.onmicrosoft.com" -ClientId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -ClientSecret "votre_secret"

.EXAMPLE
    .\Get-ShareAudit.ps1 -TenantId "contoso.onmicrosoft.com" -ClientId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -ClientSecret "votre_secret" -IncludeOneDrive -AnonymousOnly -OutputPath "C:\Audits\M365"

.NOTES
    Auteur      : Ayi NEDJIMI — ayinedjimi-consultants.fr
    Version     : 2.1.0
    Création    : 2026-05-21
    Licence     : MIT

    AVERTISSEMENT : Ce script doit uniquement être utilisé sur des tenants Microsoft 365
    pour lesquels vous disposez des autorisations légales et contractuelles appropriées.

    Permissions requises sur l'application Azure AD :
      - Sites.Read.All       (Application)
      - Files.Read.All       (Application)
      - User.Read.All        (Application)
      - SharePoint > Full Control (si PnP.PowerShell en mode app-only)

    Throttling : Le script gère automatiquement les erreurs HTTP 429 avec back-off exponentiel.

    Ressources complémentaires :
      - Guide complet        : https://ayinedjimi-consultants.fr/articles/audit-partages-sharepoint-onedrive-powershell-anonymes
      - Audit Microsoft 365  : https://ayinedjimi-consultants.fr/articles/audit-securite-microsoft-365-guide
      - Forensique M365      : https://ayinedjimi-consultants.fr/articles/forensique-microsoft-365-unified-audit-log
      - Gouvernance Entra ID : https://ayinedjimi-consultants.fr/articles/entra-id-azure-ad-securite-configuration
      - Identity Governance  : https://ayinedjimi-consultants.fr/articles/identity-governance-iga-cycle-vie-comptes

.LINK
    https://github.com/ayinedjimi/sharepoint-onedrive-share-audit
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true, HelpMessage = "Domaine du tenant (ex: contoso.onmicrosoft.com)")]
    [ValidateNotNullOrEmpty()]
    [string]$TenantId,

    [Parameter(Mandatory = $true, HelpMessage = "Client ID (GUID) de l'application Azure AD")]
    [ValidatePattern('^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$')]
    [string]$ClientId,

    [Parameter(Mandatory = $true, HelpMessage = "Secret client de l'application Azure AD")]
    [string]$ClientSecret,

    [Parameter(Mandatory = $false, HelpMessage = "Dossier de sortie pour les exports")]
    [string]$OutputPath = (Get-Location).Path,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeOneDrive,

    [Parameter(Mandatory = $false)]
    [switch]$AnonymousOnly
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 0 : CONSTANTES ET INITIALISATION
# ─────────────────────────────────────────────────────────────────────────────

$Script:Version        = "2.1.0"
$Script:StartTime      = Get-Date
$Script:LogEntries     = [System.Collections.Generic.List[string]]::new()
$Script:SharingResults = [System.Collections.Generic.List[PSObject]]::new()
$Script:MaxRetries     = 5
$Script:InitialDelay   = 2  # secondes (back-off exponentiel)
$Script:UsePnP         = $false

# Statistiques globales
$Script:Stats = @{
    TotalSites         = 0
    TotalItems         = 0
    TotalLinks         = 0
    TotalAnonymous     = 0
    TotalExpired       = 0
    Errors             = 0
}

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 1 : FONCTIONS UTILITAIRES
# ─────────────────────────────────────────────────────────────────────────────

function Write-AuditLog {
    <#
    .SYNOPSIS Journalise un message avec horodatage et niveau de sévérité.
    #>
    param(
        [string]$Message,
        [ValidateSet('INFO','WARN','ERROR','SUCCESS','DEBUG')]
        [string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $entry     = "[$timestamp] [$Level] $Message"
    $Script:LogEntries.Add($entry)

    $color = switch ($Level) {
        'INFO'    { 'Cyan' }
        'WARN'    { 'Yellow' }
        'ERROR'   { 'Red' }
        'SUCCESS' { 'Green' }
        'DEBUG'   { 'Gray' }
    }

    # BUG FIX : précédence de -not corrigée (was: -not $Var -eq 'Value')
    if ($Level -eq 'DEBUG' -and ($VerbosePreference -ne 'Continue')) { return }

    Write-Host $entry -ForegroundColor $color
}

function Invoke-WithRetry {
    <#
    .SYNOPSIS Exécute un ScriptBlock avec retry automatique sur throttling HTTP 429.
    .DESCRIPTION
        Gère les codes 429 (Too Many Requests) et 503 (Service Unavailable) avec
        un back-off exponentiel : 2s, 4s, 8s, 16s, 32s maximum.
        Les erreurs non-throttling (auth, permission) ne sont PAS retentées.
    #>
    param(
        [Parameter(Mandatory)]
        [scriptblock]$ScriptBlock,
        [string]$OperationName = "Opération"
    )

    $attempt = 0
    $delay   = $Script:InitialDelay

    while ($attempt -lt $Script:MaxRetries) {
        try {
            return (& $ScriptBlock)
        }
        catch {
            $attempt++
            $statusCode = 0

            # Extraction du code HTTP si disponible
            if ($_.Exception.Response) {
                $statusCode = [int]$_.Exception.Response.StatusCode
            }
            elseif ($_.Exception.Message -match '429|503') {
                $statusCode = if ($_.Exception.Message -match '429') { 429 } else { 503 }
            }

            # Erreurs non-récupérables : on n'attend pas, on abandonne immédiatement
            if ($statusCode -in 401, 403) {
                Write-AuditLog "Erreur d'autorisation ($statusCode) pour '$OperationName' — vérifiez les permissions Azure AD." 'ERROR'
                $Script:Stats.Errors++
                return $null
            }

            # Throttling : on attend et on réessaie
            if ($statusCode -in 429, 503) {
                Write-AuditLog "Throttling ($statusCode) pour '$OperationName'. Tentative $attempt/$Script:MaxRetries. Attente ${delay}s..." 'WARN'
                Start-Sleep -Seconds $delay
                $delay = [Math]::Min($delay * 2, 60)
                continue
            }

            # Dépassement du nombre de tentatives
            if ($attempt -ge $Script:MaxRetries) {
                Write-AuditLog "ÉCHEC définitif de '$OperationName' après $attempt tentatives : $_" 'ERROR'
                $Script:Stats.Errors++
                return $null
            }

            Write-AuditLog "Erreur transitoire pour '$OperationName' (tentative $attempt) : $_" 'WARN'
            Start-Sleep -Seconds $delay
            $delay = [Math]::Min($delay * 2, 60)
        }
    }

    Write-AuditLog "Nombre maximum de tentatives atteint pour '$OperationName'." 'ERROR'
    $Script:Stats.Errors++
    return $null
}

function Test-ModuleAvailable {
    <#
    .SYNOPSIS Vérifie si un module PowerShell est disponible.
    #>
    param([string]$ModuleName, [string]$MinVersion = "")

    try {
        if ($MinVersion) {
            $mod = Get-Module -Name $ModuleName -ListAvailable |
                   Where-Object { [version]$_.Version -ge [version]$MinVersion } |
                   Sort-Object Version -Descending |
                   Select-Object -First 1
        }
        else {
            $mod = Get-Module -Name $ModuleName -ListAvailable |
                   Sort-Object Version -Descending | Select-Object -First 1
        }
        return ($null -ne $mod)
    }
    catch { return $false }
}

function Get-TenantAdminUrl {
    <#
    .SYNOPSIS Construit l'URL du site d'administration SharePoint.
    #>
    param([string]$Tenant)

    if ($Tenant -match '^[0-9a-fA-F]{8}-') {
        throw "Fournissez le domaine du tenant (ex: contoso.onmicrosoft.com) plutôt que le GUID. Requis pour construire les URLs SharePoint."
    }

    $domain = $Tenant -replace '\.onmicrosoft\.com$', '' -replace '@', ''
    return "https://$domain-admin.sharepoint.com"
}

function Get-SharePointRootUrl {
    param([string]$Tenant)
    $domain = $Tenant -replace '\.onmicrosoft\.com$', '' -replace '@', ''
    return "https://$domain.sharepoint.com"
}

# BUG FIX : helper PS 5.1 pour remplacer l'opérateur ?? (PS 7+ uniquement)
function Get-ValueOrDefault {
    param($Value, $Default)
    if ($null -ne $Value -and $Value -ne '') { return $Value }
    return $Default
}

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 2 : CONNEXION ET AUTHENTIFICATION
# ─────────────────────────────────────────────────────────────────────────────

function Initialize-Connection {
    <#
    .SYNOPSIS
        Initialise la connexion à Microsoft 365 via PnP.PowerShell (préféré)
        ou Microsoft.Graph (fallback).
    .NOTES
        BUG FIX : #Requires -Modules supprimé — la détection est faite ici à l'exécution,
        pas au parsing, ce qui permet le fallback si PnP absent.
    #>

    Write-AuditLog "Vérification des modules disponibles..." 'INFO'

    # ── Tentative 1 : PnP.PowerShell ──────────────────────────────────────────
    if (Test-ModuleAvailable -ModuleName "PnP.PowerShell" -MinVersion "2.0") {
        try {
            Write-AuditLog "Module PnP.PowerShell détecté. Connexion en mode application..." 'INFO'
            Import-Module PnP.PowerShell -MinimumVersion 2.0 -Force -ErrorAction Stop

            $adminUrl = Get-TenantAdminUrl -Tenant $TenantId
            Connect-PnPOnline -Url $adminUrl `
                              -ClientId $ClientId `
                              -ClientSecret $ClientSecret `
                              -ErrorAction Stop

            $Script:UsePnP = $true
            Write-AuditLog "Connecté via PnP.PowerShell à $adminUrl" 'SUCCESS'
            return
        }
        catch {
            Write-AuditLog "Échec connexion PnP.PowerShell : $_. Bascule vers Microsoft.Graph..." 'WARN'
        }
    }

    # ── Tentative 2 : Microsoft.Graph ─────────────────────────────────────────
    if (Test-ModuleAvailable -ModuleName "Microsoft.Graph") {
        try {
            Write-AuditLog "Module Microsoft.Graph détecté. Connexion via client credentials..." 'INFO'
            Import-Module Microsoft.Graph.Authentication -Force -ErrorAction Stop

            $secureSecret = ConvertTo-SecureString $ClientSecret -AsPlainText -Force
            $credential   = New-Object System.Management.Automation.PSCredential($ClientId, $secureSecret)

            # BUG FIX : -Scopes supprimé — invalide en app-only (client credentials).
            # Les permissions viennent de l'app registration Azure AD (Sites.Read.All, etc.)
            Connect-MgGraph -TenantId $TenantId `
                            -ClientSecretCredential $credential `
                            -ErrorAction Stop

            $Script:UsePnP = $false
            Write-AuditLog "Connecté via Microsoft.Graph (app-only)" 'SUCCESS'
            return
        }
        catch {
            Write-AuditLog "Échec connexion Microsoft.Graph : $_" 'ERROR'
        }
    }

    # ── Aucun module disponible ────────────────────────────────────────────────
    throw @"
Aucun module compatible trouvé. Installez l'un des suivants :
  Install-Module PnP.PowerShell -MinimumVersion 2.0 -Scope CurrentUser
  Install-Module Microsoft.Graph -Scope CurrentUser

Documentation :
  https://ayinedjimi-consultants.fr/articles/audit-partages-sharepoint-onedrive-powershell-anonymes
"@
}

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 3 : RÉCUPÉRATION DES SITES SHAREPOINT
# ─────────────────────────────────────────────────────────────────────────────

function Get-AllSharePointSites {
    <#
    .SYNOPSIS Retourne la liste de tous les sites SharePoint Online du tenant.
    #>

    Write-AuditLog "Récupération de tous les sites SharePoint Online..." 'INFO'
    $sites = [System.Collections.Generic.List[PSObject]]::new()

    if ($Script:UsePnP) {
        $rawSites = Invoke-WithRetry -OperationName "Get-PnPTenantSite" -ScriptBlock {
            Get-PnPTenantSite -IncludeOneDriveSites:$false -ErrorAction Stop
        }

        if ($null -eq $rawSites) { return $sites }

        foreach ($s in $rawSites) {
            $sites.Add([PSCustomObject]@{
                Url         = $s.Url
                Title       = $s.Title
                Template    = $s.Template
                StorageUsed = $s.StorageUsageCurrent
                Type        = "SharePoint"
            })
        }
    }
    else {
        # Graph : /sites?$search=* énumère tous les sites
        $nextLink = "https://graph.microsoft.com/v1.0/sites?`$search=*&`$top=200"

        while ($nextLink) {
            $response = Invoke-WithRetry -OperationName "Graph:ListSites" -ScriptBlock {
                Invoke-MgGraphRequest -Uri $nextLink -Method GET -ErrorAction Stop
            }

            if ($null -eq $response) { break }

            foreach ($s in $response.value) {
                # BUG FIX : opérateur ternaire ? : remplacé (PS 7+ uniquement)
                $tpl = if ($s.root) { "Root" } else { "Site" }

                $sites.Add([PSCustomObject]@{
                    Url         = $s.webUrl
                    Title       = $s.displayName
                    Template    = $tpl
                    StorageUsed = 0
                    Type        = "SharePoint"
                    GraphId     = $s.id
                })
            }
            $nextLink = $response.'@odata.nextLink'
        }
    }

    Write-AuditLog "$($sites.Count) site(s) SharePoint trouvé(s)." 'SUCCESS'
    return $sites
}

function Get-AllOneDriveSites {
    <#
    .SYNOPSIS Retourne la liste des MySites OneDrive for Business de tous les utilisateurs.
    .NOTES
        Nécessite User.Read.All + Sites.Read.All en mode Application.
    #>

    Write-AuditLog "Récupération des OneDrive for Business (MySites)..." 'INFO'
    $drives = [System.Collections.Generic.List[PSObject]]::new()

    if ($Script:UsePnP) {
        # BUG FIX : filtrage côté client — la syntaxe -Filter PnP ne supporte pas -like sur URL
        $rawSites = Invoke-WithRetry -OperationName "Get-PnPTenantSite (OneDrive)" -ScriptBlock {
            Get-PnPTenantSite -IncludeOneDriveSites -ErrorAction Stop
        }

        if ($null -eq $rawSites) { return $drives }

        foreach ($s in $rawSites | Where-Object { $_.Url -like '*-my.sharepoint.com/personal/*' }) {
            $drives.Add([PSCustomObject]@{
                Url         = $s.Url
                Title       = $s.Title
                Template    = "SPSPERS"
                StorageUsed = $s.StorageUsageCurrent
                Type        = "OneDrive"
            })
        }
    }
    else {
        $usersLink = "https://graph.microsoft.com/v1.0/users?`$select=id,displayName,userPrincipalName&`$top=999"

        while ($usersLink) {
            $resp = Invoke-WithRetry -OperationName "Graph:ListUsers" -ScriptBlock {
                Invoke-MgGraphRequest -Uri $usersLink -Method GET -ErrorAction Stop
            }
            if ($null -eq $resp) { break }

            foreach ($user in $resp.value) {
                $driveUri = "https://graph.microsoft.com/v1.0/users/$($user.id)/drive"
                $drive = Invoke-WithRetry -OperationName "Graph:GetDrive:$($user.userPrincipalName)" -ScriptBlock {
                    Invoke-MgGraphRequest -Uri $driveUri -Method GET -ErrorAction SilentlyContinue
                }
                if ($drive -and $drive.webUrl) {
                    $storageUsed = 0
                    if ($drive.quota -and $drive.quota.used) {
                        $storageUsed = [Math]::Round($drive.quota.used / 1MB, 2)
                    }

                    $drives.Add([PSCustomObject]@{
                        Url         = $drive.webUrl
                        Title       = "$($user.displayName) - OneDrive"
                        Template    = "SPSPERS"
                        StorageUsed = $storageUsed
                        Type        = "OneDrive"
                        GraphId     = $drive.id
                        OwnerId     = $user.id
                    })
                }
            }
            $usersLink = $resp.'@odata.nextLink'
        }
    }

    Write-AuditLog "$($drives.Count) OneDrive trouvé(s)." 'SUCCESS'
    return $drives
}

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 4 : ANALYSE DES PARTAGES PAR SITE
# ─────────────────────────────────────────────────────────────────────────────

function Get-SharingsForSite {
    <#
    .SYNOPSIS Énumère récursivement tous les éléments d'un site et leurs liens de partage.
    #>
    param(
        [Parameter(Mandatory)]
        [PSObject]$Site
    )

    Write-Verbose "  → Analyse du site : $($Site.Url)"

    if ($Script:UsePnP) {
        return Get-SharingsViaPnP -Site $Site
    }
    else {
        return Get-SharingsViaGraph -Site $Site
    }
}

function Get-SharingsViaPnP {
    param([PSObject]$Site)

    $results = [System.Collections.Generic.List[PSObject]]::new()

    try {
        # Connexion au site cible
        Invoke-WithRetry -OperationName "Connect-PnPOnline:$($Site.Url)" -ScriptBlock {
            Connect-PnPOnline -Url $Site.Url `
                              -ClientId $ClientId `
                              -ClientSecret $ClientSecret `
                              -ErrorAction Stop
        } | Out-Null

        # Récupération de toutes les bibliothèques de documents
        $lists = Invoke-WithRetry -OperationName "Get-PnPList:$($Site.Url)" -ScriptBlock {
            Get-PnPList -ErrorAction Stop |
                Where-Object { $_.BaseType -eq "DocumentLibrary" -and -not $_.Hidden }
        }

        if ($null -eq $lists) { return $results }

        foreach ($list in $lists) {
            Write-Verbose "    Bibliothèque : $($list.Title) ($($list.ItemCount) éléments)"

            # BUG FIX : pagination CAML — on boucle jusqu'à épuisement des items
            # sans pagination, seuls les 500 premiers items étaient analysés
            $position = $null
            do {
                $pageItems = Invoke-WithRetry -OperationName "Get-PnPListItem:$($list.Title)" -ScriptBlock {
                    $query = "<View Scope='RecursiveAll'><RowLimit>500</RowLimit></View>"
                    $params = @{
                        List   = $list
                        Query  = $query
                        Fields = @("FileRef","FileLeafRef","Editor","Modified")
                        ErrorAction = 'Stop'
                    }
                    if ($position) { $params['ListItemCollectionPosition'] = $position }
                    Get-PnPListItem @params
                }

                if ($null -eq $pageItems -or $pageItems.Count -eq 0) { break }

                $Script:Stats.TotalItems += $pageItems.Count
                $position = $pageItems.ListItemCollectionPosition

                foreach ($item in $pageItems) {
                    $itemUrl = $item["FileRef"]
                    if (-not $itemUrl) { continue }

                    # BUG FIX : Get-PnPFileSharingLink n'existe pas.
                    # Utilisation de l'API REST PnP pour récupérer les SharingLinks.
                    $restUrl   = "/_api/web/GetFileByServerRelativeUrl('{0}')/ListItemAllFields/SharingLinks" -f [uri]::EscapeDataString($itemUrl)
                    $sharingInfo = Invoke-WithRetry -OperationName "REST:SharingLinks:$itemUrl" -ScriptBlock {
                        Invoke-PnPSPRestMethod -Url $restUrl -Method Get -ErrorAction SilentlyContinue
                    }

                    if ($null -eq $sharingInfo -or $null -eq $sharingInfo.value) { continue }

                    foreach ($link in $sharingInfo.value) {
                        $linkType  = if ($link.link) { $link.link.type } else { $link.type }
                        $linkScope = if ($link.link) { $link.link.scope } else { $link.scope }
                        $linkUrl   = if ($link.link) { $link.link.webUrl } else { $link.url }

                        $shareType = switch -Wildcard ($linkScope) {
                            "*anonymous*"    { "Anyone" }
                            "*organization*" { "Company" }
                            default          { "Specific" }
                        }

                        if ($AnonymousOnly -and $shareType -ne "Anyone") { continue }

                        $isExpired    = $false
                        $expDateTime  = $null
                        if ($link.expirationDateTime) {
                            $parsedDate = [datetime]::MinValue
                            if ([datetime]::TryParse($link.expirationDateTime, [ref]$parsedDate)) {
                                $expDateTime = $parsedDate
                                if ((Get-Date) -gt $parsedDate) {
                                    $isExpired = $true
                                    $Script:Stats.TotalExpired++
                                }
                            }
                        }

                        # BUG FIX : remplacement de ?? par Get-ValueOrDefault (PS 5.1 compat.)
                        $hasPassword = Get-ValueOrDefault -Value $link.requiresPassword -Default $false
                        $createdBy   = Get-ValueOrDefault -Value $link.createdBy.user.displayName -Default "N/A"

                        $entry = [PSCustomObject]@{
                            SiteUrl          = $Site.Url
                            SiteTitle        = $Site.Title
                            SiteType         = $Site.Type
                            ItemPath         = $itemUrl
                            ItemName         = $item["FileLeafRef"]
                            ShareType        = $shareType
                            ShareUrl         = $linkUrl
                            CreatedBy        = $createdBy
                            CreatedDate      = $link.createdDateTime
                            ExpirationDate   = $expDateTime
                            IsExpired        = $isExpired
                            IsAnonymous      = ($shareType -eq "Anyone")
                            HasPassword      = $hasPassword
                            Roles            = (Get-ValueOrDefault -Value ($link.roles -join "; ") -Default "")
                            AuditTimestamp   = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
                        }

                        $results.Add($entry)
                        $Script:Stats.TotalLinks++
                        if ($shareType -eq "Anyone") { $Script:Stats.TotalAnonymous++ }
                    }
                }

            } while ($null -ne $position)
        }
    }
    catch {
        Write-AuditLog "Erreur lors de l'analyse de $($Site.Url) : $_" 'ERROR'
        $Script:Stats.Errors++
    }

    return $results
}

function Get-SharingsViaGraph {
    param([PSObject]$Site)

    $results = [System.Collections.Generic.List[PSObject]]::new()

    try {
        $siteId = $Site.GraphId

        if (-not $siteId) {
            # BUG FIX : résolution par URL via le format hostname:/path (pas $search avec URL encodée)
            try {
                $uri = [uri]$Site.Url
                $hostname = $uri.Host
                $path     = $uri.AbsolutePath.TrimEnd('/')
                if (-not $path -or $path -eq '/') { $path = '/sites/root' }
                $lookupUri = "https://graph.microsoft.com/v1.0/sites/$($hostname):$($path)"

                $lookup = Invoke-WithRetry -OperationName "Graph:LookupSite:$($Site.Url)" -ScriptBlock {
                    Invoke-MgGraphRequest -Uri $lookupUri -Method GET -ErrorAction Stop
                }
                $siteId = $lookup.id
            }
            catch {
                Write-AuditLog "Impossible de résoudre le site Graph pour $($Site.Url) : $_" 'WARN'
                return $results
            }
        }

        if (-not $siteId) {
            Write-AuditLog "ID de site introuvable pour $($Site.Url)" 'WARN'
            return $results
        }

        # Drives du site
        $drivesResp = Invoke-WithRetry -OperationName "Graph:GetDrives:$siteId" -ScriptBlock {
            Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/sites/$siteId/drives" -Method GET -ErrorAction Stop
        }

        if ($null -eq $drivesResp) { return $results }

        foreach ($drive in $drivesResp.value) {
            Enumerate-GraphDriveItems -DriveId $drive.id -Site $Site -Results $results
        }
    }
    catch {
        Write-AuditLog "Erreur Graph pour $($Site.Url) : $_" 'ERROR'
        $Script:Stats.Errors++
    }

    return $results
}

function Enumerate-GraphDriveItems {
    param(
        [string]$DriveId,
        [PSObject]$Site,
        [System.Collections.Generic.List[PSObject]]$Results,
        [string]$FolderPath = "root"
    )

    $itemsUri = "https://graph.microsoft.com/v1.0/drives/$DriveId/items/$FolderPath/children?`$top=200"

    while ($itemsUri) {
        $resp = Invoke-WithRetry -OperationName "Graph:ListItems:$FolderPath" -ScriptBlock {
            Invoke-MgGraphRequest -Uri $itemsUri -Method GET -ErrorAction Stop
        }
        if ($null -eq $resp) { break }

        foreach ($item in $resp.value) {
            $Script:Stats.TotalItems++

            # Récursion dans les dossiers
            if ($item.folder) {
                Enumerate-GraphDriveItems -DriveId $DriveId -Site $Site -Results $Results -FolderPath $item.id
                continue
            }

            # Récupération des permissions de l'item
            $permsUri = "https://graph.microsoft.com/v1.0/drives/$DriveId/items/$($item.id)/permissions"
            $permsResp = Invoke-WithRetry -OperationName "Graph:GetPerms:$($item.name)" -ScriptBlock {
                Invoke-MgGraphRequest -Uri $permsUri -Method GET -ErrorAction SilentlyContinue
            }

            if ($null -eq $permsResp) { continue }

            foreach ($perm in $permsResp.value) {
                # On ne s'intéresse qu'aux liens de partage (sharing links)
                if (-not $perm.link) { continue }

                $shareType = switch ($perm.link.scope) {
                    "anonymous"    { "Anyone" }
                    "organization" { "Company" }
                    "users"        { "Specific" }
                    default        { "Unknown" }
                }

                if ($AnonymousOnly -and $shareType -ne "Anyone") { continue }

                $isExpired   = $false
                $expDateTime = $null
                if ($perm.expirationDateTime) {
                    $parsedDate = [datetime]::MinValue
                    # BUG FIX : [datetime]::TryParse au lieu de cast direct (évite crash sur date malformée)
                    if ([datetime]::TryParse($perm.expirationDateTime, [ref]$parsedDate)) {
                        $expDateTime = $parsedDate
                        if ((Get-Date) -gt $parsedDate) {
                            $isExpired = $true
                            $Script:Stats.TotalExpired++
                        }
                    }
                }

                # BUG FIX : remplacement de ?? par Get-ValueOrDefault (PS 5.1 compat.)
                $createdBy   = Get-ValueOrDefault -Value $perm.grantedToV2.user.displayName -Default "N/A"
                $hasPassword = Get-ValueOrDefault -Value $perm.link.preventsDownload -Default $false

                $entry = [PSCustomObject]@{
                    SiteUrl          = $Site.Url
                    SiteTitle        = $Site.Title
                    SiteType         = $Site.Type
                    ItemPath         = $item.webUrl
                    ItemName         = $item.name
                    ShareType        = $shareType
                    ShareUrl         = $perm.link.webUrl
                    CreatedBy        = $createdBy
                    CreatedDate      = $perm.createdDateTime
                    ExpirationDate   = $expDateTime
                    IsExpired        = $isExpired
                    IsAnonymous      = ($shareType -eq "Anyone")
                    HasPassword      = $hasPassword
                    Roles            = (Get-ValueOrDefault -Value ($perm.roles -join "; ") -Default "")
                    AuditTimestamp   = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
                }

                $Results.Add($entry)
                $Script:Stats.TotalLinks++
                if ($shareType -eq "Anyone") { $Script:Stats.TotalAnonymous++ }
            }
        }
        $itemsUri = $resp.'@odata.nextLink'
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 5 : EXPORTS (CSV + HTML)
# ─────────────────────────────────────────────────────────────────────────────

function Export-ToCSV {
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $csvPath   = Join-Path $OutputPath "ShareAudit_$timestamp.csv"

    Write-AuditLog "Export CSV : $csvPath" 'INFO'
    $Script:SharingResults | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8 -Delimiter ","
    Write-AuditLog "CSV exporté : $($Script:SharingResults.Count) entrées" 'SUCCESS'
    return $csvPath
}

function Export-ToHTML {
    $timestamp  = Get-Date -Format 'yyyyMMdd_HHmmss'
    $htmlPath   = Join-Path $OutputPath "ShareAudit_$timestamp.html"
    $reportDate = Get-Date -Format 'dd/MM/yyyy HH:mm:ss'
    $duration   = [Math]::Round(((Get-Date) - $Script:StartTime).TotalMinutes, 1)

    $top10 = $Script:SharingResults |
        Group-Object ItemPath |
        Sort-Object Count -Descending |
        Select-Object -First 10 |
        ForEach-Object {
            $item = $_.Group[0]
            "<tr><td>$([System.Web.HttpUtility]::HtmlEncode($item.ItemName))</td><td>$([System.Web.HttpUtility]::HtmlEncode($item.SiteTitle))</td><td>$($_.Count)</td><td><a href='$($item.ItemPath)' target='_blank'>Ouvrir</a></td></tr>"
        }

    $tableRows = $Script:SharingResults | ForEach-Object {
        $rowClass = if ($_.IsAnonymous) { "danger" } elseif ($_.SiteType -eq "OneDrive") { "warning" } else { "" }
        $badge = switch ($_.ShareType) {
            "Anyone"   { "<span class='badge badge-danger'>Anyone</span>" }
            "Company"  { "<span class='badge badge-info'>Company</span>" }
            "Specific" { "<span class='badge badge-success'>Specific</span>" }
            default    { "<span class='badge badge-secondary'>$($_.ShareType)</span>" }
        }
        $expired = if ($_.IsExpired) { "<span class='badge badge-warning'>Expiré</span>" } else { "" }

        # BUG FIX : TryParse au lieu de cast direct pour les dates d'affichage HTML
        $createdStr = "N/A"
        if ($_.CreatedDate) {
            $d = [datetime]::MinValue
            if ([datetime]::TryParse($_.CreatedDate, [ref]$d)) { $createdStr = $d.ToString('dd/MM/yyyy') }
        }
        $expiresStr = if ($_.ExpirationDate) { $_.ExpirationDate.ToString('dd/MM/yyyy') } else { "Jamais" }

        "<tr class='$rowClass'>
            <td>$([System.Web.HttpUtility]::HtmlEncode($_.SiteTitle))</td>
            <td>$($_.SiteType)</td>
            <td title='$($_.ItemPath)'>$([System.Web.HttpUtility]::HtmlEncode($_.ItemName))</td>
            <td>$badge $expired</td>
            <td>$([System.Web.HttpUtility]::HtmlEncode($_.CreatedBy))</td>
            <td>$createdStr</td>
            <td>$expiresStr</td>
            <td><a href='$($_.ShareUrl)' target='_blank'>Lien</a></td>
        </tr>"
    }

    $html = @"
<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Rapport Audit Partages M365 — $reportDate</title>
<style>
  :root { --red:#dc3545;--yellow:#ffc107;--green:#28a745;--blue:#0078d4;--gray:#6c757d; }
  body { font-family:'Segoe UI',sans-serif; background:#f8f9fa; color:#212529; margin:0; }
  header { background:linear-gradient(135deg,#0078d4,#004578); color:#fff; padding:2rem; }
  header h1 { margin:0; font-size:1.8rem; }
  header p  { margin:.5rem 0 0; opacity:.8; }
  .container { max-width:1400px; margin:0 auto; padding:1.5rem; }
  .stats-grid { display:grid; grid-template-columns:repeat(auto-fit,minmax(180px,1fr)); gap:1rem; margin:1.5rem 0; }
  .stat-card { background:#fff; border-radius:8px; padding:1.2rem; text-align:center; box-shadow:0 2px 4px rgba(0,0,0,.08); border-top:4px solid var(--blue); }
  .stat-card.danger  { border-top-color:var(--red); }
  .stat-card.warning { border-top-color:var(--yellow); }
  .stat-card.success { border-top-color:var(--green); }
  .stat-value { font-size:2.5rem; font-weight:700; }
  .stat-label { color:var(--gray); font-size:.85rem; margin-top:.3rem; }
  table { width:100%; border-collapse:collapse; background:#fff; border-radius:8px; overflow:hidden; box-shadow:0 2px 4px rgba(0,0,0,.08); }
  th { background:#0078d4; color:#fff; padding:.75rem 1rem; text-align:left; cursor:pointer; user-select:none; white-space:nowrap; }
  th:hover { background:#005a9e; }
  td { padding:.6rem 1rem; border-bottom:1px solid #dee2e6; font-size:.88rem; }
  tr.danger td { background:#fff5f5; }
  tr.warning td { background:#fffdf0; }
  tr:hover td { background:#e9f3ff; }
  .badge { display:inline-block; padding:.2em .5em; border-radius:4px; font-size:.78rem; font-weight:600; color:#fff; }
  .badge-danger { background:var(--red); }
  .badge-info { background:var(--blue); }
  .badge-success { background:var(--green); }
  .badge-warning { background:var(--yellow); color:#212529; }
  .badge-secondary { background:var(--gray); }
  #filterInput { width:100%; padding:.5rem .75rem; margin-bottom:1rem; border:1px solid #ced4da; border-radius:4px; font-size:.95rem; }
  .section-title { font-size:1.2rem; font-weight:600; color:#004578; margin:2rem 0 1rem; border-left:4px solid #0078d4; padding-left:.75rem; }
  .resources { background:#fff; border-radius:8px; padding:1.5rem; box-shadow:0 2px 4px rgba(0,0,0,.08); margin-top:2rem; }
  .resources h3 { color:#004578; margin-top:0; }
  .resources a { color:#0078d4; text-decoration:none; display:block; margin:.3rem 0; }
  .resources a:hover { text-decoration:underline; }
  footer { text-align:center; padding:1.5rem; color:var(--gray); font-size:.8rem; border-top:1px solid #dee2e6; margin-top:2rem; }
  a { color:#0078d4; }
</style>
</head>
<body>
<header>
  <h1>Rapport Audit Partages Microsoft 365</h1>
  <p>Généré le $reportDate — Durée : ${duration} min | Outil : Get-ShareAudit.ps1 v$($Script:Version)</p>
</header>
<div class="container">

  <div class="stats-grid">
    <div class="stat-card">
      <div class="stat-value">$($Script:Stats.TotalSites)</div>
      <div class="stat-label">Sites analysés</div>
    </div>
    <div class="stat-card">
      <div class="stat-value">$($Script:Stats.TotalItems)</div>
      <div class="stat-label">Éléments scannés</div>
    </div>
    <div class="stat-card">
      <div class="stat-value">$($Script:Stats.TotalLinks)</div>
      <div class="stat-label">Liens de partage</div>
    </div>
    <div class="stat-card danger">
      <div class="stat-value" style="color:var(--red)">$($Script:Stats.TotalAnonymous)</div>
      <div class="stat-label">Liens anonymes (Anyone)</div>
    </div>
    <div class="stat-card warning">
      <div class="stat-value" style="color:var(--yellow)">$($Script:Stats.TotalExpired)</div>
      <div class="stat-label">Liens expirés</div>
    </div>
    <div class="stat-card success">
      <div class="stat-value" style="color:$(if ($Script:Stats.Errors -gt 0) {'var(--red)'} else {'var(--green)'})">$($Script:Stats.Errors)</div>
      <div class="stat-label">Erreurs</div>
    </div>
  </div>

  <div class="section-title">Top 10 — Éléments les plus partagés</div>
  <table>
    <thead><tr><th>Fichier / Dossier</th><th>Site</th><th>Nb liens</th><th>Accès</th></tr></thead>
    <tbody>$($top10 -join "`n")</tbody>
  </table>

  <div class="section-title">Tous les partages détectés</div>
  <input id="filterInput" type="text" placeholder="Filtrer par nom, site, type..." oninput="filterTable()">
  <table id="mainTable" data-sort-dir="">
    <thead>
      <tr>
        <th onclick="sortTable(0)">Site ↕</th>
        <th onclick="sortTable(1)">Type ↕</th>
        <th onclick="sortTable(2)">Élément ↕</th>
        <th onclick="sortTable(3)">Partage ↕</th>
        <th onclick="sortTable(4)">Créé par ↕</th>
        <th onclick="sortTable(5)">Date création ↕</th>
        <th onclick="sortTable(6)">Expiration ↕</th>
        <th>Lien</th>
      </tr>
    </thead>
    <tbody id="tableBody">$($tableRows -join "`n")</tbody>
  </table>

  <div class="resources">
    <h3>Ressources complémentaires</h3>
    <a href="https://ayinedjimi-consultants.fr/articles/audit-partages-sharepoint-onedrive-powershell-anonymes" target="_blank">Guide complet : Audit des partages SharePoint & OneDrive avec PowerShell</a>
    <a href="https://ayinedjimi-consultants.fr/articles/audit-securite-microsoft-365-guide" target="_blank">Audit de sécurité Microsoft 365 — méthodologie complète</a>
    <a href="https://ayinedjimi-consultants.fr/articles/forensique-microsoft-365-unified-audit-log" target="_blank">Forensique Microsoft 365 — Unified Audit Log</a>
    <a href="https://ayinedjimi-consultants.fr/articles/entra-id-azure-ad-securite-configuration" target="_blank">Sécurisation Entra ID / Azure AD</a>
    <a href="https://ayinedjimi-consultants.fr/articles/identity-governance-iga-cycle-vie-comptes" target="_blank">Identity Governance — cycle de vie des comptes</a>
    <a href="https://github.com/ayinedjimi/sharepoint-onedrive-share-audit" target="_blank">Repo GitHub — sharepoint-onedrive-share-audit</a>
  </div>

</div>
<footer>
  Généré par <strong>Get-ShareAudit.ps1 v$($Script:Version)</strong> —
  <a href="https://ayinedjimi-consultants.fr/articles/audit-partages-sharepoint-onedrive-powershell-anonymes">
    Guide complet sur ayinedjimi-consultants.fr
  </a> |
  <a href="https://github.com/ayinedjimi/sharepoint-onedrive-share-audit">GitHub</a>
</footer>
<script>
function filterTable() {
  const q = document.getElementById('filterInput').value.toLowerCase();
  document.querySelectorAll('#tableBody tr').forEach(function(row) {
    row.style.display = row.textContent.toLowerCase().indexOf(q) >= 0 ? '' : 'none';
  });
}
function sortTable(col) {
  var table = document.getElementById('mainTable');
  var rows  = Array.prototype.slice.call(table.querySelectorAll('tbody tr'));
  var asc   = table.dataset.sortDir !== 'asc';
  table.dataset.sortDir = asc ? 'asc' : 'desc';
  rows.sort(function(a, b) {
    var x = (a.cells[col] ? a.cells[col].textContent.trim() : '');
    var y = (b.cells[col] ? b.cells[col].textContent.trim() : '');
    return asc ? x.localeCompare(y, 'fr') : y.localeCompare(x, 'fr');
  });
  var tbody = table.querySelector('tbody');
  rows.forEach(function(r) { tbody.appendChild(r); });
}
</script>
</body>
</html>
"@

    $html | Out-File -FilePath $htmlPath -Encoding UTF8
    Write-AuditLog "Rapport HTML exporté : $htmlPath" 'SUCCESS'
    return $htmlPath
}

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 6 : RÉSUMÉ CONSOLE
# ─────────────────────────────────────────────────────────────────────────────

function Show-Summary {
    param([string]$CsvPath, [string]$HtmlPath)

    $duration = [Math]::Round(((Get-Date) - $Script:StartTime).TotalMinutes, 1)

    Write-Host ""
    Write-Host "════════════════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "  RÉSUMÉ DE L'AUDIT — Get-ShareAudit.ps1 v$Script:Version" -ForegroundColor Cyan
    Write-Host "════════════════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  Durée totale        : $duration minutes"         -ForegroundColor White
    Write-Host "  Sites analysés      : $($Script:Stats.TotalSites)"    -ForegroundColor White
    Write-Host "  Éléments scannés    : $($Script:Stats.TotalItems)"    -ForegroundColor White
    Write-Host "  Liens de partage    : $($Script:Stats.TotalLinks)"    -ForegroundColor White
    Write-Host "  Liens ANONYMES      : $($Script:Stats.TotalAnonymous)" -ForegroundColor $(if ($Script:Stats.TotalAnonymous -gt 0) { 'Red' } else { 'Green' })
    Write-Host "  Liens expirés       : $($Script:Stats.TotalExpired)"  -ForegroundColor Yellow
    Write-Host "  Erreurs             : $($Script:Stats.Errors)"        -ForegroundColor $(if ($Script:Stats.Errors -gt 0) { 'Red' } else { 'Green' })
    Write-Host ""

    if ($CsvPath)  { Write-Host "  Export CSV   : $CsvPath"  -ForegroundColor Cyan }
    if ($HtmlPath) { Write-Host "  Rapport HTML : $HtmlPath" -ForegroundColor Cyan }
    Write-Host ""

    $top10 = $Script:SharingResults |
        Group-Object ItemPath |
        Sort-Object Count -Descending |
        Select-Object -First 10

    if ($top10) {
        Write-Host "  TOP 10 — ÉLÉMENTS LES PLUS PARTAGÉS :" -ForegroundColor Yellow
        $top10 | ForEach-Object {
            $item = $_.Group[0]
            Write-Host ("    {0} liens — {1} [{2}]" -f $_.Count.ToString().PadLeft(4), $item.ItemName, $item.SiteTitle) -ForegroundColor White
        }
        Write-Host ""
    }

    if ($Script:Stats.TotalAnonymous -gt 0) {
        Write-Host "  ATTENTION : $($Script:Stats.TotalAnonymous) lien(s) anonyme(s) détecté(s) !" -ForegroundColor Red
        Write-Host "  Ces liens sont accessibles SANS authentification." -ForegroundColor Red
        Write-Host "  Guide de remédiation : https://ayinedjimi-consultants.fr/articles/audit-partages-sharepoint-onedrive-powershell-anonymes" -ForegroundColor Yellow
        Write-Host ""
    }

    Write-Host "  Documentation : https://ayinedjimi-consultants.fr/articles/audit-partages-sharepoint-onedrive-powershell-anonymes" -ForegroundColor Gray
    Write-Host "  GitHub        : https://github.com/ayinedjimi/sharepoint-onedrive-share-audit" -ForegroundColor Gray
    Write-Host ""
    Write-Host "════════════════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 7 : POINT D'ENTRÉE PRINCIPAL
# ─────────────────────────────────────────────────────────────────────────────

function Main {
    Write-Host ""
    Write-Host "════════════════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "  Get-ShareAudit.ps1 v$Script:Version — Audit Partages M365"   -ForegroundColor Cyan
    Write-Host "  ayinedjimi-consultants.fr — Cybersécurité Microsoft 365"      -ForegroundColor Gray
    Write-Host "  https://github.com/ayinedjimi/sharepoint-onedrive-share-audit" -ForegroundColor Gray
    Write-Host "════════════════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host ""

    if (-not (Test-Path $OutputPath)) {
        Write-AuditLog "Création du dossier de sortie : $OutputPath" 'INFO'
        New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
    }

    Initialize-Connection

    $allSites = [System.Collections.Generic.List[PSObject]]::new()

    $spSites = Get-AllSharePointSites
    foreach ($s in $spSites) { $allSites.Add($s) }

    if ($IncludeOneDrive) {
        $odSites = Get-AllOneDriveSites
        foreach ($s in $odSites) { $allSites.Add($s) }
    }

    $Script:Stats.TotalSites = $allSites.Count
    Write-AuditLog "Total sites à analyser : $($allSites.Count)" 'INFO'

    $siteIndex = 0
    foreach ($site in $allSites) {
        $siteIndex++
        $pct = [Math]::Round(($siteIndex / [Math]::Max($allSites.Count, 1)) * 100)

        Write-Progress -Activity "Audit des partages M365" `
                       -Status "Site $siteIndex/$($allSites.Count) : $($site.Title)" `
                       -PercentComplete $pct

        $siteResults = Get-SharingsForSite -Site $site
        foreach ($r in $siteResults) { $Script:SharingResults.Add($r) }
    }

    Write-Progress -Activity "Audit des partages M365" -Completed

    $csvPath  = Export-ToCSV
    $htmlPath = Export-ToHTML

    Show-Summary -CsvPath $csvPath -HtmlPath $htmlPath

    if ($Script:UsePnP) {
        Disconnect-PnPOnline -ErrorAction SilentlyContinue
    }
    else {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
    }

    Write-AuditLog "Audit terminé." 'SUCCESS'
}

Main
