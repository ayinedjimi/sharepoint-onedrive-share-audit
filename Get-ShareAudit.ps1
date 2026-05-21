#Requires -Version 5.1
#Requires -Modules @{ModuleName="PnP.PowerShell"; ModuleVersion="2.0"}

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
    Identifiant du tenant Azure AD (GUID ou domaine : contoso.onmicrosoft.com).

.PARAMETER ClientId
    Application ID (Client ID) de l'application Azure AD enregistrée.

.PARAMETER ClientSecret
    Secret client de l'application Azure AD (SecureString recommandé).

.PARAMETER OutputPath
    Dossier de destination pour les exports CSV et HTML. Par défaut : répertoire courant.

.PARAMETER IncludeOneDrive
    Switch : inclure les OneDrive for Business de tous les utilisateurs.

.PARAMETER AnonymousOnly
    Switch : ne rapporter que les liens "Anyone" (partages anonymes).

.PARAMETER Verbose
    Switch natif PowerShell : affiche les messages de diagnostic détaillés.

.EXAMPLE
    .\Get-ShareAudit.ps1 -TenantId "contoso.onmicrosoft.com" -ClientId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -ClientSecret "votre_secret"

.EXAMPLE
    .\Get-ShareAudit.ps1 -TenantId "contoso.onmicrosoft.com" -ClientId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -ClientSecret "votre_secret" -IncludeOneDrive -AnonymousOnly -OutputPath "C:\Audits\M365"

.NOTES
    Auteur      : Ayi NEDJIMI — ayinedjimi-consultants.fr
    Version     : 2.0.0
    Création    : 2026-05-21
    Licence     : MIT

    AVERTISSEMENT : Ce script doit uniquement être utilisé sur des tenants Microsoft 365
    pour lesquels vous disposez des autorisations légales et contractuelles appropriées.

    Permissions requises sur l'application Azure AD :
      - Sites.Read.All (Application)
      - Files.Read.All (Application)
      - User.Read.All  (Application)
      - SharePoint > Full Control (si PnP.PowerShell en mode app-only)

    Throttling : Le script gère automatiquement les erreurs HTTP 429 avec back-off exponentiel.

    Article de référence :
    https://ayinedjimi-consultants.fr/articles/audit-partages-sharepoint-onedrive-powershell-anonymes

.LINK
    https://github.com/ayinedjimi/sharepoint-onedrive-share-audit
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true, HelpMessage = "ID ou domaine du tenant Azure AD")]
    [ValidateNotNullOrEmpty()]
    [string]$TenantId,

    [Parameter(Mandatory = $true, HelpMessage = "Client ID de l'application Azure AD")]
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

$Script:Version        = "2.0.0"
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

    # Les messages DEBUG ne s'affichent qu'en mode Verbose
    if ($Level -eq 'DEBUG' -and -not $VerbosePreference -eq 'Continue') { return }

    Write-Host $entry -ForegroundColor $color
}

function Invoke-WithRetry {
    <#
    .SYNOPSIS Exécute un ScriptBlock avec retry automatique sur throttling HTTP 429.
    .DESCRIPTION
        Gère les codes 429 (Too Many Requests) et 503 (Service Unavailable) avec
        un back-off exponentiel : 2s, 4s, 8s, 16s, 32s maximum.
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
            return & $ScriptBlock
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

            # Throttling : on attend et on réessaie
            if ($statusCode -in 429, 503) {
                Write-AuditLog "Throttling détecté ($statusCode) pour '$OperationName'. Tentative $attempt/$Script:MaxRetries. Attente ${delay}s..." 'WARN'
                Start-Sleep -Seconds $delay
                $delay = [Math]::Min($delay * 2, 60)  # cap à 60s
                continue
            }

            # Autre erreur : on abandonne après MaxRetries tentatives
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
}

function Test-ModuleAvailable {
    <#
    .SYNOPSIS Vérifie si un module PowerShell est disponible et le charge si possible.
    #>
    param([string]$ModuleName, [string]$MinVersion = "")

    try {
        if ($MinVersion) {
            $mod = Get-Module -Name $ModuleName -ListAvailable |
                   Where-Object { [version]$_.Version -ge [version]$MinVersion } |
                   Sort-Object Version -Descending |
                   Select-Object -First 1
        } else {
            $mod = Get-Module -Name $ModuleName -ListAvailable |
                   Sort-Object Version -Descending | Select-Object -First 1
        }
        return ($null -ne $mod)
    }
    catch { return $false }
}

function Get-TenantAdminUrl {
    <#
    .SYNOPSIS Construit l'URL du site d'administration SharePoint à partir du TenantId ou du domaine.
    #>
    param([string]$Tenant)

    # Si c'est un GUID, on ne peut pas construire l'URL directement
    if ($Tenant -match '^[0-9a-fA-F]{8}-') {
        throw "Pour construire l'URL admin SharePoint, fournissez le domaine du tenant (ex: contoso.onmicrosoft.com) plutôt que le GUID."
    }

    $domain = $Tenant -replace '\.onmicrosoft\.com$', '' -replace '@', ''
    return "https://$domain-admin.sharepoint.com"
}

function Get-SharePointRootUrl {
    param([string]$Tenant)
    $domain = $Tenant -replace '\.onmicrosoft\.com$', '' -replace '@', ''
    return "https://$domain.sharepoint.com"
}

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 2 : CONNEXION ET AUTHENTIFICATION
# ─────────────────────────────────────────────────────────────────────────────

function Initialize-Connection {
    <#
    .SYNOPSIS
        Initialise la connexion à Microsoft 365 via PnP.PowerShell (préféré)
        ou Microsoft.Graph (fallback).
    .DESCRIPTION
        PnP.PowerShell offre des cmdlets natives SharePoint plus riches.
        Microsoft.Graph couvre les mêmes données via l'API REST mais nécessite
        plus d'appels. Le script détecte automatiquement la disponibilité.
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

            Connect-MgGraph -TenantId $TenantId `
                            -ClientSecretCredential $credential `
                            -Scopes "Sites.Read.All","Files.Read.All","User.Read.All" `
                            -ErrorAction Stop

            $Script:UsePnP = $false
            Write-AuditLog "Connecté via Microsoft.Graph" 'SUCCESS'
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
        # PnP : Get-PnPTenantSite retourne tous les sites du tenant
        $rawSites = Invoke-WithRetry -OperationName "Get-PnPTenantSite" -ScriptBlock {
            Get-PnPTenantSite -IncludeOneDriveSites:$false -ErrorAction Stop
        }

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
                $sites.Add([PSCustomObject]@{
                    Url         = $s.webUrl
                    Title       = $s.displayName
                    Template    = $s.root ? "Root" : "Site"
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
        $rawSites = Invoke-WithRetry -OperationName "Get-PnPTenantSite (OneDrive)" -ScriptBlock {
            Get-PnPTenantSite -IncludeOneDriveSites -Filter "Url -like '-my.sharepoint.com/personal/'" -ErrorAction Stop
        }

        foreach ($s in $rawSites) {
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
        # Graph : énumère les utilisateurs puis leurs drives
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
                    $drives.Add([PSCustomObject]@{
                        Url         = $drive.webUrl
                        Title       = "$($user.displayName) - OneDrive"
                        Template    = "SPSPERS"
                        StorageUsed = [Math]::Round($drive.quota.used / 1MB, 2)
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
    .PARAMETER Site PSObject retourné par Get-AllSharePointSites ou Get-AllOneDriveSites.
    #>
    param(
        [Parameter(Mandatory)]
        [PSObject]$Site
    )

    Write-Verbose "  → Analyse du site : $($Site.Url)"
    $localResults = [System.Collections.Generic.List[PSObject]]::new()

    if ($Script:UsePnP) {
        $localResults = Get-SharingsViaPnP -Site $Site
    }
    else {
        $localResults = Get-SharingsViaGraph -Site $Site
    }

    return $localResults
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
        }

        # Récupération de toutes les bibliothèques de documents
        $lists = Invoke-WithRetry -OperationName "Get-PnPList:$($Site.Url)" -ScriptBlock {
            Get-PnPList -ErrorAction Stop |
                Where-Object { $_.BaseType -eq "DocumentLibrary" -and -not $_.Hidden }
        }

        foreach ($list in $lists) {
            Write-Verbose "    Bibliothèque : $($list.Title) ($($list.ItemCount) éléments)"

            # Récupération des éléments avec champs de partage
            $camlQuery = "<View Scope='RecursiveAll'><RowLimit>500</RowLimit></View>"
            $items = Invoke-WithRetry -OperationName "Get-PnPListItem:$($list.Title)" -ScriptBlock {
                Get-PnPListItem -List $list -Query $camlQuery `
                                -Fields "FileRef","FileLeafRef","Editor","Modified","SMTotalFileCount" `
                                -ErrorAction Stop
            }

            if ($null -eq $items) { continue }
            $Script:Stats.TotalItems += $items.Count

            foreach ($item in $items) {
                # Récupération des informations de partage via l'API REST PnP
                $itemUrl      = $item["FileRef"]
                $sharingInfo  = Invoke-WithRetry -OperationName "Get-SharingInfo:$itemUrl" -ScriptBlock {
                    Get-PnPFileSharingLink -FileUrl $itemUrl -ErrorAction SilentlyContinue
                }

                if ($null -eq $sharingInfo) { continue }

                foreach ($link in $sharingInfo) {
                    $shareType = switch -Wildcard ($link.Link.Scope) {
                        "*anonymous*" { "Anyone" }
                        "*organization*" { "Company" }
                        default { "Specific" }
                    }

                    # Filtre si -AnonymousOnly
                    if ($AnonymousOnly -and $shareType -ne "Anyone") { continue }

                    $isExpired = $false
                    if ($link.ExpirationDateTime -and (Get-Date) -gt $link.ExpirationDateTime) {
                        $isExpired = $true
                        $Script:Stats.TotalExpired++
                    }

                    $entry = [PSCustomObject]@{
                        SiteUrl          = $Site.Url
                        SiteTitle        = $Site.Title
                        SiteType         = $Site.Type
                        ItemPath         = $itemUrl
                        ItemName         = $item["FileLeafRef"]
                        ShareType        = $shareType
                        ShareUrl         = $link.Link.WebUrl
                        CreatedBy        = $link.GrantedToIdentitiesV2.User.DisplayName -join "; "
                        CreatedDate      = $link.CreatedDateTime
                        ExpirationDate   = $link.ExpirationDateTime
                        IsExpired        = $isExpired
                        IsAnonymous      = ($shareType -eq "Anyone")
                        HasPassword      = $link.Link.PreventedBySensitivityLabel ?? $false
                        Roles            = ($link.Roles -join "; ")
                        AuditTimestamp   = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
                    }

                    $results.Add($entry)
                    $Script:Stats.TotalLinks++
                    if ($shareType -eq "Anyone") { $Script:Stats.TotalAnonymous++ }
                }
            }
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
        # Résolution de l'ID de site Graph
        $siteId = $Site.GraphId
        if (-not $siteId) {
            # On tente de résoudre via l'URL
            $encoded = [uri]::EscapeDataString($Site.Url)
            $lookup  = Invoke-WithRetry -OperationName "Graph:LookupSite" -ScriptBlock {
                Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/sites?`$search=$encoded" -Method GET -ErrorAction Stop
            }
            $siteId = $lookup.value[0].id
        }

        if (-not $siteId) {
            Write-AuditLog "Impossible de résoudre le site Graph pour $($Site.Url)" 'WARN'
            return $results
        }

        # Drives du site
        $drivesResp = Invoke-WithRetry -OperationName "Graph:GetDrives:$siteId" -ScriptBlock {
            Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/sites/$siteId/drives" -Method GET -ErrorAction Stop
        }

        foreach ($drive in $drivesResp.value) {
            $results = Enumerate-GraphDriveItems -DriveId $drive.id -Site $Site -Results $results
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
                $Results = Enumerate-GraphDriveItems -DriveId $DriveId -Site $Site -Results $Results -FolderPath $item.id
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

                $isExpired = $false
                if ($perm.expirationDateTime -and (Get-Date) -gt [datetime]$perm.expirationDateTime) {
                    $isExpired = $true
                    $Script:Stats.TotalExpired++
                }

                $entry = [PSCustomObject]@{
                    SiteUrl          = $Site.Url
                    SiteTitle        = $Site.Title
                    SiteType         = $Site.Type
                    ItemPath         = $item.webUrl
                    ItemName         = $item.name
                    ShareType        = $shareType
                    ShareUrl         = $perm.link.webUrl
                    CreatedBy        = ($perm.grantedToV2.user.displayName ?? "N/A")
                    CreatedDate      = $perm.createdDateTime
                    ExpirationDate   = $perm.expirationDateTime
                    IsExpired        = $isExpired
                    IsAnonymous      = ($shareType -eq "Anyone")
                    HasPassword      = ($null -ne $perm.link.preventsDownload)
                    Roles            = ($perm.roles -join "; ")
                    AuditTimestamp   = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
                }

                $Results.Add($entry)
                $Script:Stats.TotalLinks++
                if ($shareType -eq "Anyone") { $Script:Stats.TotalAnonymous++ }
            }
        }
        $itemsUri = $resp.'@odata.nextLink'
    }

    return $Results
}

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 5 : EXPORTS (CSV + HTML)
# ─────────────────────────────────────────────────────────────────────────────

function Export-ToCSV {
    <#
    .SYNOPSIS Exporte les résultats d'audit dans un fichier CSV.
    #>

    $timestamp  = Get-Date -Format 'yyyyMMdd_HHmmss'
    $csvPath    = Join-Path $OutputPath "ShareAudit_$timestamp.csv"

    Write-AuditLog "Export CSV : $csvPath" 'INFO'

    $Script:SharingResults |
        Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8 -Delimiter ";"

    Write-AuditLog "CSV exporté : $($Script:SharingResults.Count) entrées" 'SUCCESS'
    return $csvPath
}

function Export-ToHTML {
    <#
    .SYNOPSIS Génère un rapport HTML interactif avec tri, filtre et mise en couleur des risques.
    #>

    $timestamp   = Get-Date -Format 'yyyyMMdd_HHmmss'
    $htmlPath    = Join-Path $OutputPath "ShareAudit_$timestamp.html"
    $reportDate  = Get-Date -Format 'dd/MM/yyyy HH:mm:ss'
    $duration    = [Math]::Round(((Get-Date) - $Script:StartTime).TotalMinutes, 1)

    # Top 10 des items les plus partagés
    $top10 = $Script:SharingResults |
        Group-Object ItemPath |
        Sort-Object Count -Descending |
        Select-Object -First 10 |
        ForEach-Object {
            $item = $_.Group[0]
            "<tr><td>$($item.ItemName)</td><td>$($item.SiteTitle)</td><td>$($_.Count)</td><td><a href='$($item.ItemPath)' target='_blank'>Ouvrir</a></td></tr>"
        }

    # Lignes du tableau principal
    $tableRows = $Script:SharingResults | ForEach-Object {
        $rowClass = if ($_.IsAnonymous) { "danger" } elseif ($_.SiteType -eq "OneDrive") { "warning" } else { "" }
        $badge    = switch ($_.ShareType) {
            "Anyone"   { "<span class='badge badge-danger'>Anyone</span>" }
            "Company"  { "<span class='badge badge-info'>Company</span>" }
            "Specific" { "<span class='badge badge-success'>Specific</span>" }
            default    { "<span class='badge badge-secondary'>$($_.ShareType)</span>" }
        }
        $expired  = if ($_.IsExpired) { "<span class='badge badge-warning'>Expiré</span>" } else { "" }

        "<tr class='$rowClass'>
            <td>$($_.SiteTitle)</td>
            <td>$($_.SiteType)</td>
            <td title='$($_.ItemPath)'>$($_.ItemName)</td>
            <td>$badge $expired</td>
            <td>$($_.CreatedBy)</td>
            <td>$(if ($_.CreatedDate) { ([datetime]$_.CreatedDate).ToString('dd/MM/yyyy') } else { 'N/A' })</td>
            <td>$(if ($_.ExpirationDate) { ([datetime]$_.ExpirationDate).ToString('dd/MM/yyyy') } else { 'Jamais' })</td>
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
  :root { --red:#dc3545;--yellow:#ffc107;--green:#28a745;--blue:#007bff;--gray:#6c757d; }
  body { font-family: 'Segoe UI', sans-serif; background:#f8f9fa; color:#212529; margin:0; }
  header { background:linear-gradient(135deg,#0078d4,#004578); color:#fff; padding:2rem; }
  header h1 { margin:0; font-size:1.8rem; }
  header p  { margin:.5rem 0 0; opacity:.8; }
  .container { max-width:1400px; margin:0 auto; padding:1.5rem; }
  .stats-grid { display:grid; grid-template-columns:repeat(auto-fit,minmax(180px,1fr)); gap:1rem; margin:1.5rem 0; }
  .stat-card { background:#fff; border-radius:8px; padding:1.2rem; text-align:center; box-shadow:0 2px 4px rgba(0,0,0,.08); border-top:4px solid var(--blue); }
  .stat-card.danger { border-top-color:var(--red); }
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
      <div class="stat-value">$($Script:Stats.Errors)</div>
      <div class="stat-label">Erreurs</div>
    </div>
  </div>

  <div class="section-title">Top 10 — Éléments les plus partagés</div>
  <table>
    <thead><tr><th>Fichier / Dossier</th><th>Site</th><th>Nb liens</th><th>Accès</th></tr></thead>
    <tbody>$($top10 -join "`n")</tbody>
  </table>

  <div class="section-title">Tous les partages détectés</div>
  <input id="filterInput" type="text" placeholder="Filtrer par nom, site, type..." onkeyup="filterTable()">
  <table id="mainTable">
    <thead>
      <tr>
        <th onclick="sortTable(0)">Site</th>
        <th onclick="sortTable(1)">Type</th>
        <th onclick="sortTable(2)">Élément</th>
        <th onclick="sortTable(3)">Partage</th>
        <th onclick="sortTable(4)">Créé par</th>
        <th onclick="sortTable(5)">Date création</th>
        <th onclick="sortTable(6)">Expiration</th>
        <th>Lien</th>
      </tr>
    </thead>
    <tbody id="tableBody">$($tableRows -join "`n")</tbody>
  </table>

</div>
<footer>
  Généré par <strong>Get-ShareAudit.ps1 v$($Script:Version)</strong> —
  <a href="https://ayinedjimi-consultants.fr/articles/audit-partages-sharepoint-onedrive-powershell-anonymes">
    Guide complet sur ayinedjimi-consultants.fr
  </a>
</footer>
<script>
function filterTable() {
  const q = document.getElementById('filterInput').value.toLowerCase();
  document.querySelectorAll('#tableBody tr').forEach(row => {
    row.style.display = row.textContent.toLowerCase().includes(q) ? '' : 'none';
  });
}
function sortTable(col) {
  const table = document.getElementById('mainTable');
  const rows  = Array.from(table.querySelectorAll('tbody tr'));
  const asc   = table.dataset.sortDir !== 'asc';
  table.dataset.sortDir = asc ? 'asc' : 'desc';
  rows.sort((a, b) => {
    const x = a.cells[col]?.textContent.trim() ?? '';
    const y = b.cells[col]?.textContent.trim() ?? '';
    return asc ? x.localeCompare(y) : y.localeCompare(x);
  });
  rows.forEach(r => table.querySelector('tbody').appendChild(r));
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
    Write-Host "  Durée totale        : $duration minutes" -ForegroundColor White
    Write-Host "  Sites analysés      : $($Script:Stats.TotalSites)" -ForegroundColor White
    Write-Host "  Éléments scannés    : $($Script:Stats.TotalItems)" -ForegroundColor White
    Write-Host "  Liens de partage    : $($Script:Stats.TotalLinks)" -ForegroundColor White
    Write-Host "  Liens ANONYMES      : $($Script:Stats.TotalAnonymous)" -ForegroundColor $(if ($Script:Stats.TotalAnonymous -gt 0) { 'Red' } else { 'Green' })
    Write-Host "  Liens expirés       : $($Script:Stats.TotalExpired)" -ForegroundColor Yellow
    Write-Host "  Erreurs             : $($Script:Stats.Errors)" -ForegroundColor $(if ($Script:Stats.Errors -gt 0) { 'Red' } else { 'Green' })
    Write-Host ""

    if ($CsvPath)  { Write-Host "  Export CSV  : $CsvPath"  -ForegroundColor Cyan }
    if ($HtmlPath) { Write-Host "  Rapport HTML: $HtmlPath" -ForegroundColor Cyan }
    Write-Host ""

    # Top 10 dans la console
    $top10 = $Script:SharingResults |
        Group-Object ItemPath |
        Sort-Object Count -Descending |
        Select-Object -First 10

    if ($top10) {
        Write-Host "  TOP 10 — ÉLÉMENTS LES PLUS PARTAGÉS :" -ForegroundColor Yellow
        $top10 | ForEach-Object {
            $item = $_.Group[0]
            Write-Host "    $($_.Count.ToString().PadLeft(4)) liens — $($item.ItemName) [$($item.SiteTitle)]" -ForegroundColor White
        }
        Write-Host ""
    }

    if ($Script:Stats.TotalAnonymous -gt 0) {
        Write-Host "  ⚠  ATTENTION : $($Script:Stats.TotalAnonymous) lien(s) anonyme(s) détecté(s) !" -ForegroundColor Red
        Write-Host "     Ces liens sont accessibles SANS authentification." -ForegroundColor Red
        Write-Host "     Révoquez-les immédiatement ou appliquez une politique de gouvernance." -ForegroundColor Red
        Write-Host ""
    }

    Write-Host "════════════════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 7 : POINT D'ENTRÉE PRINCIPAL
# ─────────────────────────────────────────────────────────────────────────────

function Main {
    Write-Host ""
    Write-Host "════════════════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "  Get-ShareAudit.ps1 v$Script:Version — Audit Partages M365" -ForegroundColor Cyan
    Write-Host "  ayinedjimi-consultants.fr — Cybersécurité Microsoft 365" -ForegroundColor Gray
    Write-Host "════════════════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host ""

    # Validation du dossier de sortie
    if (-not (Test-Path $OutputPath)) {
        Write-AuditLog "Création du dossier de sortie : $OutputPath" 'INFO'
        New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
    }

    # Connexion
    Initialize-Connection

    # Collecte des sites
    $allSites = [System.Collections.Generic.List[PSObject]]::new()

    $spSites = Get-AllSharePointSites
    $spSites | ForEach-Object { $allSites.Add($_) }

    if ($IncludeOneDrive) {
        $odSites = Get-AllOneDriveSites
        $odSites | ForEach-Object { $allSites.Add($_) }
    }

    $Script:Stats.TotalSites = $allSites.Count
    Write-AuditLog "Total sites à analyser : $($allSites.Count)" 'INFO'

    # Analyse de chaque site avec barre de progression
    $siteIndex = 0
    foreach ($site in $allSites) {
        $siteIndex++
        $pct = [Math]::Round(($siteIndex / $allSites.Count) * 100)

        Write-Progress -Activity "Audit des partages M365" `
                       -Status "Site $siteIndex/$($allSites.Count) : $($site.Title)" `
                       -PercentComplete $pct `
                       -CurrentOperation "Analyse en cours..."

        $siteResults = Get-SharingsForSite -Site $site
        $siteResults | ForEach-Object { $Script:SharingResults.Add($_) }
    }

    Write-Progress -Activity "Audit des partages M365" -Completed

    # Exports
    $csvPath  = Export-ToCSV
    $htmlPath = Export-ToHTML

    # Résumé
    Show-Summary -CsvPath $csvPath -HtmlPath $htmlPath

    # Déconnexion propre
    if ($Script:UsePnP) {
        Disconnect-PnPOnline -ErrorAction SilentlyContinue
    }
    else {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
    }

    Write-AuditLog "Audit terminé." 'SUCCESS'
}

# Lancement
Main
