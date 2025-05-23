<#
.SYNOPSIS
    Hilfsfunktionen zur Extraktion von Tenant-Informationen aus SharePoint URLs.

.DESCRIPTION
    Diese Skriptdatei enthält Funktionen zum Parsen von SharePoint Online URLs,
    um Tenant-Namen, Admin URLs und andere tenant-spezifische Informationen zu extrahieren.

.NOTES
    Dateiname: Get-TenantInfo.ps1
    Autor: Andreas Hepp
    Version: 1.0
#>

function Get-TenantInfoFromUrl {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Url
    )

    Write-Host "[DEBUG] Extrahiere Tenant-Informationen aus URL: $Url" -ForegroundColor Cyan
    
    # URL säubern und normalisieren
    $Url = $Url.Trim().TrimEnd('/')
    
    # Prüfen ob URL gültige SharePoint Online URL ist
    if (-not ($Url -like "https://*.sharepoint.com*")) {
        Write-Host "[ERROR] Die URL ist keine gültige SharePoint Online URL: $Url" -ForegroundColor Red
        return $null
    }
    
    # Standard-Ergebnisobjekt
    $result = [PSCustomObject]@{
        TenantName = $null
        TenantId = $null
        AdminUrl = $null
        RootUrl = $null
        IsAdminSite = $false
        AuthorityUrl = $null
    }
    
    # Extrahiere Tenant-Namen
    if ($Url -match "https://([^.]+)-admin\.sharepoint\.com") {
        # Admin-URL
        $result.TenantName = $matches[1]
        $result.IsAdminSite = $true
    }
    elseif ($Url -match "https://([^.]+)\.sharepoint\.com") {
        # Standard SharePoint URL
        $result.TenantName = $matches[1]
    }
    else {
        # Unbekanntes Format
        Write-Host "[ERROR] Konnte Tenant-Namen nicht aus der URL extrahieren: $Url" -ForegroundColor Red
        return $null
    }
    
    # Konstruiere die verschiedenen URLs basierend auf dem Tenant-Namen
    $result.RootUrl = "https://$($result.TenantName).sharepoint.com"
    $result.AdminUrl = "https://$($result.TenantName)-admin.sharepoint.com"
    
    # Für die modernen authentifizierungsmethoden, setze die Authority URL 
    # (tenant-spezifischer Endpunkt für Azure AD)
    $result.AuthorityUrl = "https://login.microsoftonline.com/$($result.TenantName).onmicrosoft.com"
    
    # Tenant ID ist in diesem Fall nicht verfügbar ohne Azure AD Abfrage
    # Bei Bedarf kann hier eine Funktion ergänzt werden, um die TenantId zu ermitteln
    $result.TenantId = "$($result.TenantName).onmicrosoft.com"
    
    Write-Host "[DEBUG] Erfolgreich extrahierte Tenant-Infos:" -ForegroundColor Green
    Write-Host "[DEBUG] - Tenant: $($result.TenantName)" -ForegroundColor Green
    Write-Host "[DEBUG] - Admin URL: $($result.AdminUrl)" -ForegroundColor Green
    Write-Host "[DEBUG] - Root URL: $($result.RootUrl)" -ForegroundColor Green
    Write-Host "[DEBUG] - Authority: $($result.AuthorityUrl)" -ForegroundColor Green
    
    return $result
}

function Get-SPOSiteUrl {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Url,
        
        [Parameter(Mandatory = $false)]
        [switch]$AsAdminUrl
    )
    
    # Hole Tenant-Infos
    $tenantInfo = Get-TenantInfoFromUrl -Url $Url
    
    if ($null -eq $tenantInfo) {
        return $null
    }
    
    if ($AsAdminUrl) {
        return $tenantInfo.AdminUrl
    }
    else {
        # Falls es eine Admin-URL ist, gib die Root-URL zurück
        if ($tenantInfo.IsAdminSite) {
            return $tenantInfo.RootUrl
        }
        # Andernfalls gib die ursprüngliche URL zurück
        return $Url
    }
}

function Get-SPOAdminCenterUrl {
    param (
        [Parameter(Mandatory = $true)]
        [string]$TenantName
    )
    
    # Bereinige Tenant-Namen
    $TenantName = $TenantName.Replace(".onmicrosoft.com", "").Replace("-admin", "")
    
    # Konstruiere Admin-URL
    $adminUrl = "https://$TenantName-admin.sharepoint.com"
    
    return $adminUrl
}

function Parse-SPOUrl {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Url
    )
    
    $urlParts = @{
        TenantName = $null
        SiteName = $null
        IsAdminSite = $false
        IsSubsite = $false
        SubsitePath = $null
        FullSitePath = $null
    }
    
    # Säubere die URL
    $Url = $Url.Trim().TrimEnd('/')
    
    # Admin Site
    if ($Url -match "https://([^.]+)-admin\.sharepoint\.com(/.+)?") {
        $urlParts.TenantName = $matches[1]
        $urlParts.IsAdminSite = $true
        $urlParts.FullSitePath = $Url
    }
    # Root-Site
    elseif ($Url -match "https://([^.]+)\.sharepoint\.com/?$") {
        $urlParts.TenantName = $matches[1]
        $urlParts.SiteName = ""  # Root-Site
        $urlParts.FullSitePath = $Url
    }
    # Site Collection
    elseif ($Url -match "https://([^.]+)\.sharepoint\.com/sites/([^/]+)(/.*)?") {
        $urlParts.TenantName = $matches[1]
        $urlParts.SiteName = $matches[2]
        
        if ($matches.Count > 3 -and $matches[3]) {
            $urlParts.IsSubsite = $true
            $urlParts.SubsitePath = $matches[3]
        }
        
        $urlParts.FullSitePath = $Url
    }
    # Teams/Gruppen Site Collection
    elseif ($Url -match "https://([^.]+)\.sharepoint\.com/teams/([^/]+)(/.*)?") {
        $urlParts.TenantName = $matches[1]
        $urlParts.SiteName = $matches[2]
        
        if ($matches.Count > 3 -and $matches[3]) {
            $urlParts.IsSubsite = $true
            $urlParts.SubsitePath = $matches[3]
        }
        
        $urlParts.FullSitePath = $Url
    }
    else {
        Write-Host "[ERROR] Konnte URL nicht parsen: $Url" -ForegroundColor Red
        return $null
    }
    
    return $urlParts
}

# Exportiere alle Funktionen
Export-ModuleMember -Function Get-TenantInfoFromUrl, Get-SPOSiteUrl, Get-SPOAdminCenterUrl, Parse-SPOUrl
