#Requires -Version 5.1

<#
.SYNOPSIS
    SharePoint Online Management Tool mit WPF GUI
.DESCRIPTION
    Ein umfassendes Tool zur Verwaltung von SharePoint Online-Umgebungen mit
    integrierter WPF-Benutzeroberfläche, Modulprüfung und vielen SPO-Funktionen.
.NOTES
    Name: easySPO
    Version: 0.0.1
    Author: EASYit
    Datum: 26.04.2025
#>

#region Initialisierung und Einstellungen
# Fehlerbehandlung aktivieren
$ErrorActionPreference = "Stop"
$DebugPreference = "Continue"

# Protokollierungsfunktion
function Write-Log {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [ValidateSet("INFO", "WARNING", "ERROR", "SUCCESS")]
        [string]$Level = "INFO"
    )
    
    try {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logMessage = "[$timestamp] [$Level] $Message"
        
        # Ausgabe in verschiedenen Farben je nach Level
        switch ($Level) {
            "INFO" { Write-Host $logMessage -ForegroundColor Cyan }
            "WARNING" { Write-Host $logMessage -ForegroundColor Yellow }
            "ERROR" { Write-Host $logMessage -ForegroundColor Red }
            "SUCCESS" { Write-Host $logMessage -ForegroundColor Green }
        }
        
        # In Logdatei schreiben
        $logDir = Join-Path $PSScriptRoot "Logs"
        if (-not (Test-Path $logDir)) {
            New-Item -ItemType Directory -Path $logDir -Force | Out-Null
        }
        
        $logFile = Join-Path $logDir "easySPO_$(Get-Date -Format 'yyyy-MM-dd').log"
        Add-Content -Path $logFile -Value $logMessage
    }
    catch {
        $errorMessage = ${_}
        Write-Warning "Fehler beim Schreiben des Logs: $errorMessage"
    }
}

#region Modul-Management
# Prüfen der erforderlichen Module
function Check-RequiredModules {
    try {
        Write-Log "Prüfe erforderliche Module..." -Level "INFO"
        
        # Nur das neueste Modul für SharePoint Online prüfen
        $requiredModules = @(
            "PnP.PowerShell"
        )
        
        $moduleStatus = @()
        
        foreach ($moduleName in $requiredModules) {
            $status = "Nicht installiert"
            $version = "N/A"
            $latestVersion = "N/A"
            $notes = ""
            
            # Prüfen, ob das Modul installiert ist
            $module = Get-Module -Name $moduleName -ListAvailable
            
            if ($module) {
                # Nehme die höchste verfügbare Version
                $highestVersion = $module | Sort-Object Version -Descending | Select-Object -First 1
                $version = $highestVersion.Version.ToString()
                $status = "Installiert"
                
                # Prüfen, ob ein Update verfügbar ist
                try {
                    $onlineModule = Find-Module -Name $moduleName -ErrorAction SilentlyContinue
                    if ($onlineModule) {
                        $latestVersion = $onlineModule.Version.ToString()
                        
                        if ([System.Version]$latestVersion -gt [System.Version]$version) {
                            $notes = "Update verfügbar"
                            $status = "Update verfügbar"
                        }
                        else {
                            $notes = "Auf dem neuesten Stand"
                        }
                    }
                    else {
                        $notes = "Online-Informationen nicht verfügbar"
                    }
                }
                catch {
                    $notes = "Fehler bei der Online-Prüfung"
                }
            }
            
            $moduleStatus += [PSCustomObject]@{
                Name = $moduleName
                Status = $status
                Version = $version
                LatestVersion = $latestVersion
                Notes = $notes
            }
        }
        
        # GUI-Liste aktualisieren
        if ($script:GUIElements -and $script:GUIElements.ContainsKey("lvModules")) {
            $script:GUIElements["lvModules"].Dispatcher.Invoke(
                [Action]{
                    $script:GUIElements["lvModules"].ItemsSource = $moduleStatus
                },
                "Normal"
            )
        }
        
        # Überprüfen, ob alle erforderlichen Module installiert sind
        $allModulesInstalled = -not ($moduleStatus | Where-Object { $_.Status -eq "Nicht installiert" })
        return $allModulesInstalled
    }
    catch {
        $errorMessage = ${_}
        Write-Log "Fehler bei der Überprüfung der Module: $errorMessage" -Level "ERROR"
        return $false
    }
}

# Prüfen und installieren der erforderlichen Module
function Test-AndInstallModules {
    param (
        [Parameter(Mandatory = $true)]
        [string[]]$RequiredModules
    )
    
    try {
        foreach ($module in $RequiredModules) {
            Write-Log "Prüfe Modul ${module}..." -Level "INFO"
            
            if (-not (Get-Module -Name $module -ListAvailable)) {
                Write-Log "Modul ${module} nicht gefunden. Installation wird gestartet..." -Level "WARNING"
                
                try {
                    Install-Module -Name $module -Force -AllowClobber -Scope CurrentUser
                    Write-Log "Modul ${module} erfolgreich installiert." -Level "SUCCESS"
                }
                catch {
                    $errorMessage = ${_}
                    Write-Log "Fehler bei der Installation von Modul ${module}: $errorMessage" -Level "ERROR"
                    return $false
                }
            }
            else {
                Write-Log "Modul ${module} bereits installiert." -Level "INFO"
                
                # Prüfe auf Updates
                $installedVersion = (Get-Module -Name $module -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1).Version
                $onlineVersion = Find-Module -Name $module | Select-Object -ExpandProperty Version
                
                if ($onlineVersion -gt $installedVersion) {
                    Write-Log "Update für ${module} verfügbar (Installiert: $installedVersion, Verfügbar: $onlineVersion)" -Level "WARNING"
                }
            }
        }
        
        # Nach der Installation Status aktualisieren
        Check-RequiredModules
        
        return $true
    }
    catch {
        $errorMessage = ${_}
        Write-Log "Fehler bei der Modulprüfung: $errorMessage" -Level "ERROR"
        return $false
    }
}
#endregion
#region SharePoint Online Funktionen
# Verbindung zu SharePoint Online herstellen
function Connect-SharePointOnline {
    param (
        [Parameter(Mandatory = $true)]
        [string]$AdminUrl,
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.PSCredential]$Credentials
    )
    
    try {
        Write-Log "Verbinde mit SharePoint Online Admin-Center: ${AdminUrl}" -Level "INFO"
        
        # Verbindung nur mit PnP.PowerShell herstellen (neuestes Modul)
        Connect-PnPOnline -Url $AdminUrl -Credentials $Credentials
        
        Write-Log "Verbindung zu SharePoint Online hergestellt." -Level "SUCCESS"
        return $true
    }
    catch {
        $errorMessage = ${_}
        Write-Log "Fehler bei der Verbindung zu SharePoint Online: $errorMessage" -Level "ERROR"
        return $false
    }
}

# Mit SharePoint Online verbinden (GUI-Aktion)
function Connect-ToSharePointOnline {
    try {
        Write-Log "Starte Verbindungsaufbau zu SharePoint Online..." -Level "INFO"
        Update-GuiText -ControlName "txtStatus" -Text "Verbindung wird hergestellt..."
        
        # Eingaben aus GUI holen
        $adminUrl = $script:GUIElements["txtAdminUrl"].Text
        $username = $script:GUIElements["txtUsername"].Text
        $password = $script:GUIElements["txtPassword"].SecurePassword
        
        # Eingaben validieren
        if ([string]::IsNullOrWhiteSpace($adminUrl) -or [string]::IsNullOrWhiteSpace($username)) {
            $message = "Bitte geben Sie Admin-URL und Benutzername ein."
            Write-Log $message -Level "WARNING"
            Update-GuiText -ControlName "txtConnectionStatus" -Text "$message`r`n"
            Update-GuiText -ControlName "txtStatus" -Text "Fehlerhafte Eingabe"
            return
        }
        
        # URL-Format korrigieren, wenn nötig
        if (-not $adminUrl.StartsWith("https://")) {
            $adminUrl = "https://" + $adminUrl
            $script:GUIElements["txtAdminUrl"].Text = $adminUrl
        }
        
        # Anmeldedaten erstellen
        $credentials = New-Object System.Management.Automation.PSCredential($username, $password)
        
        # Modulprüfung durchführen
        $modulesOK = Check-RequiredModules
        if (-not $modulesOK) {
            $message = "Es fehlen benötigte Module. Bitte installieren Sie alle erforderlichen Module."
            Write-Log $message -Level "WARNING"
            Update-GuiText -ControlName "txtConnectionStatus" -Text "$message`r`n"
            
            $result = [System.Windows.MessageBox]::Show(
                "Einige erforderliche Module sind nicht installiert. Möchten Sie diese jetzt installieren?", 
                "Module fehlen", 
                [System.Windows.MessageBoxButton]::YesNo, 
                [System.Windows.MessageBoxImage]::Warning
            )
            
            if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
                Install-RequiredModules
            }
            
            Update-GuiText -ControlName "txtStatus" -Text "Bereit"
            return
        }
        
        # Self-Diagnostics durchführen
        $diagnosticsMessage = Perform-SelfDiagnostics
        if ($diagnosticsMessage -ne "OK") {
            Write-Log "Self-Diagnostics fehlgeschlagen: $diagnosticsMessage" -Level "WARNING"
            Update-GuiText -ControlName "txtConnectionStatus" -Text "Self-Diagnostics fehlgeschlagen: $diagnosticsMessage`r`n"
            Update-GuiText -ControlName "txtStatus" -Text "Bereit"
            return
        }
        
        # Verbindung herstellen und Status aktualisieren
        $connectionSuccess = Connect-SharePointOnline -AdminUrl $adminUrl -Credentials $credentials
        
        if ($connectionSuccess) {
            $script:isConnected = $true
            
            # GUI-Elemente aktualisieren
            $script:GUIElements["btnConnect"].IsEnabled = $false
            $script:GUIElements["btnDisconnect"].IsEnabled = $true
            $script:GUIElements["txtAdminUrl"].IsEnabled = $false
            $script:GUIElements["txtUsername"].IsEnabled = $false
            $script:GUIElements["txtPassword"].IsEnabled = $false
            
            # Tabs aktivieren
            $script:GUIElements["tabSiteCollections"].IsEnabled = $true
            $script:GUIElements["tabLists"].IsEnabled = $true
            $script:GUIElements["tabUsers"].IsEnabled = $true
            $script:GUIElements["tabSettings"].IsEnabled = $true
            
            # Status anzeigen
            $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $connectionInfo = "Verbunden mit $adminUrl als $username`r`nVerbindung hergestellt: $timestamp"
            Update-GuiText -ControlName "txtConnectionStatus" -Text $connectionInfo
            Update-GuiText -ControlName "txtConnectedUser" -Text "Verbunden als: $username"
            Update-GuiText -ControlName "txtStatus" -Text "Verbunden"
            
            # Direkt Websitesammlungen laden
            Refresh-SiteCollections
            
            Write-Log "Erfolgreich mit SharePoint Online verbunden." -Level "SUCCESS"
        }
        else {
            Update-GuiText -ControlName "txtConnectionStatus" -Text "Fehler beim Herstellen der Verbindung. Bitte überprüfen Sie die Anmeldedaten."
            Update-GuiText -ControlName "txtStatus" -Text "Verbindungsfehler"
        }
    }
    catch {
        $errorMessage = ${_}
        Write-Log "Fehler beim Verbinden mit SharePoint Online: $errorMessage" -Level "ERROR"
        Update-GuiText -ControlName "txtConnectionStatus" -Text "Fehler: $errorMessage"
        Update-GuiText -ControlName "txtStatus" -Text "Fehler"
    }
}

# Verbindung zu SharePoint Online trennen
function Disconnect-FromSharePointOnline {
    try {
        Write-Log "Trenne Verbindung zu SharePoint Online..." -Level "INFO"
        Update-GuiText -ControlName "txtStatus" -Text "Verbindung wird getrennt..."
        
        # Verbindung nur mit PnP.PowerShell trennen
        Disconnect-PnPOnline -ErrorAction SilentlyContinue
        
        $script:isConnected = $false
        
        # GUI-Status aktualisieren
        $script:GUIElements["btnConnect"].IsEnabled = $true
        $script:GUIElements["btnDisconnect"].IsEnabled = $false
        $script:GUIElements["txtAdminUrl"].IsEnabled = $true
        $script:GUIElements["txtUsername"].IsEnabled = $true
        $script:GUIElements["txtPassword"].IsEnabled = $true
        
        # Tabs deaktivieren
        $script:GUIElements["tabSiteCollections"].IsEnabled = $false
        $script:GUIElements["tabLists"].IsEnabled = $false
        $script:GUIElements["tabUsers"].IsEnabled = $false
        $script:GUIElements["tabSettings"].IsEnabled = $false
        
        # Status aktualisieren
        Update-GuiText -ControlName "txtConnectionStatus" -Text "Verbindung getrennt."
        Update-GuiText -ControlName "txtConnectedUser" -Text "Nicht verbunden"
        Update-GuiText -ControlName "txtStatus" -Text "Bereit"
        
        # DataGrids leeren
        $script:GUIElements["dgSites"].Dispatcher.Invoke(
            [Action]{
                $script:GUIElements["dgSites"].ItemsSource = $null
            },
            "Normal"
        )
        
        $script:GUIElements["dgLists"].Dispatcher.Invoke(
            [Action]{
                $script:GUIElements["dgLists"].ItemsSource = $null
            },
            "Normal"
        )
        
        $script:GUIElements["dgUsers"].Dispatcher.Invoke(
            [Action]{
                $script:GUIElements["dgUsers"].ItemsSource = $null
            },
            "Normal"
        )
        
        # Dropdowns leeren
        $script:GUIElements["cmbSiteUrl"].Dispatcher.Invoke(
            [Action]{
                $script:GUIElements["cmbSiteUrl"].ItemsSource = $null
            },
            "Normal"
        )
        
        $script:GUIElements["cmbUserSiteUrl"].Dispatcher.Invoke(
            [Action]{
                $script:GUIElements["cmbUserSiteUrl"].ItemsSource = $null
            },
            "Normal"
        )
        
        Write-Log "Verbindung zu SharePoint Online erfolgreich getrennt." -Level "SUCCESS"
    }
    catch {
        $errorMessage = ${_}
        Write-Log "Fehler beim Trennen der Verbindung: $errorMessage" -Level "ERROR"
        Update-GuiText -ControlName "txtStatus" -Text "Fehler"
    }
}

# Self-Diagnostics durchführen
function Perform-SelfDiagnostics {
    try {
        Write-Log "Führe Self-Diagnostics durch..." -Level "INFO"
        
        # 1. Prüfen, ob alle erforderlichen Icons existieren
        $iconPaths = @(
            "assets/info.png",
            "assets/settings.png",
            "assets/close.png"
        )
        
        foreach ($iconPath in $iconPaths) {
            $fullPath = Join-Path -Path $PSScriptRoot -ChildPath $iconPath
            if (-not (Test-Path $fullPath)) {
                return "Icon nicht gefunden: $iconPath"
            }
        }
        
        # 2. Prüfen der Benutzerberechtigungen
        if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
            Write-Log "Warnung: Skript läuft nicht mit Administratorrechten" -Level "WARNING"
            # Kein Abbruch, nur Warnung
        }
        
        # 3. Prüfen der PowerShell-Version
        $psVersion = $PSVersionTable.PSVersion
        if ($psVersion.Major -lt 5 -or ($psVersion.Major -eq 5 -and $psVersion.Minor -lt 1)) {
            return "PowerShell-Version zu alt. Mindestens 5.1 erforderlich, gefunden: $($psVersion.ToString())"
        }
        
        Write-Log "Self-Diagnostics erfolgreich abgeschlossen." -Level "INFO"
        return "OK"
    }
    catch {
        $errorMessage = ${_}
        Write-Log "Fehler bei Self-Diagnostics: $errorMessage" -Level "ERROR"
        return "Fehler: $errorMessage"
    }
}

# Websitesammlungen auflisten
function Get-SPOSiteCollections {
    try {
        Write-Log "Hole Websitesammlungen..." -Level "INFO"
        
        $sites = Get-SPOSite -Limit All
        Write-Log "Erfolgreich $($sites.Count) Websitesammlungen gefunden." -Level "SUCCESS"
        
        return $sites
    }
    catch {
        $errorMessage = ${_}
        Write-Log "Fehler beim Abrufen der Websitesammlungen: $errorMessage" -Level "ERROR"
        return $null
    }
}

# Websitesammlungsberechtigungen verwalten
function Get-SPOSitePermissions {
    param (
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl
    )
    
    try {
        Write-Log "Hole Berechtigungen für Site: $SiteUrl" -Level "INFO"
        
        $admins = Get-SPOSiteGroup -Site $SiteUrl
        Write-Log "Berechtigungen für $SiteUrl erfolgreich abgerufen." -Level "SUCCESS"
        
        return $admins
    }
    catch {
        $errorMessage = ${_}
        Write-Log "Fehler beim Abrufen der Berechtigungen: $errorMessage" -Level "ERROR"
        return $null
    }
}

# Listen abrufen
function Get-SPOLists {
    param (
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl
    )
    
    try {
        Write-Log "Hole Listen für Site: $SiteUrl" -Level "INFO"
        
        $ctx = Connect-PnPOnline -Url $SiteUrl -ReturnConnection
        $lists = Get-PnPList -Connection $ctx
        
        Write-Log "Erfolgreich $($lists.Count) Listen gefunden." -Level "SUCCESS"
        return $lists
    }
    catch {
        $errorMessage = ${_}
        Write-Log "Fehler beim Abrufen der Listen: $errorMessage" -Level "ERROR"
        return $null
    }
}

# Neue Websitesammlung erstellen
function New-SPOSiteCollection {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Url,
        [Parameter(Mandatory = $true)]
        [string]$Title,
        [Parameter(Mandatory = $true)]
        [string]$Owner,
        [Parameter(Mandatory = $false)]
        [int]$StorageQuota = 1024,
        [Parameter(Mandatory = $false)]
        [string]$Template = "STS#3"
    )
    
    try {
        Write-Log "Erstelle neue Websitesammlung: $Title ($Url)" -Level "INFO"
        
        New-SPOSite -Url $Url -Title $Title -Owner $Owner -StorageQuota $StorageQuota -Template $Template
        
        Write-Log "Websitesammlung erfolgreich erstellt: $Title" -Level "SUCCESS"
        return $true
    }
    catch {
        $errorMessage = ${_}
        Write-Log "Fehler beim Erstellen der Websitesammlung: $errorMessage" -Level "ERROR"
        return $false
    }
}

# Tenant-weite Einstellungen abrufen
function Get-SPOTenantSettings {
    try {
        Write-Log "Rufe Tenant-Einstellungen ab..." -Level "INFO"
        
        $settings = Get-SPOTenant
        
        Write-Log "Tenant-Einstellungen erfolgreich abgerufen." -Level "SUCCESS"
        return $settings
    }
    catch {
        $errorMessage = ${_}
        Write-Log "Fehler beim Abrufen der Tenant-Einstellungen: $errorMessage" -Level "ERROR"
        return $null
    }
}

# Externe Freigabe konfigurieren
function Set-SPOSharingCapability {
    param (
        [Parameter(Mandatory = $true)]
        [ValidateSet("Disabled", "ExistingExternalUserSharingOnly", "ExternalUserSharingOnly", "ExternalUserAndGuestSharing")]
        [string]$SharingCapability
    )
    
    try {
        Write-Log "Konfiguriere externe Freigabeeinstellungen: $SharingCapability" -Level "INFO"
        
        Set-SPOTenant -SharingCapability $SharingCapability
        
        Write-Log "Freigabeeinstellungen erfolgreich auf '$SharingCapability' gesetzt." -Level "SUCCESS"
        return $true
    }
    catch {
        $errorMessage = ${_}
        Write-Log "Fehler beim Konfigurieren der Freigabeeinstellungen: $errorMessage" -Level "ERROR"
        return $false
    }
}

# SharePoint-Benutzer verwalten
function Get-SPOUsers {
    param (
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl
    )
    
    try {
        Write-Log "Hole Benutzer für Site: $SiteUrl" -Level "INFO"
        
        $ctx = Connect-PnPOnline -Url $SiteUrl -ReturnConnection
        $users = Get-PnPUser -Connection $ctx
        
        Write-Log "Erfolgreich $($users.Count) Benutzer gefunden." -Level "SUCCESS"
        return $users
    }
    catch {
        $errorMessage = ${_}
        Write-Log "Fehler beim Abrufen der Benutzer: $errorMessage" -Level "ERROR"
        return $null
    }
}

# OneDrive-Einstellungen verwalten
function Get-OneDriveSettings {
    try {
        Write-Log "Rufe OneDrive-Einstellungen ab..." -Level "INFO"
        
        $settings = Get-SPOTenantSyncClientRestriction
        
        Write-Log "OneDrive-Einstellungen erfolgreich abgerufen." -Level "SUCCESS"
        return $settings
    }
    catch {
        $errorMessage = ${_}
        Write-Log "Fehler beim Abrufen der OneDrive-Einstellungen: $errorMessage" -Level "ERROR"
        return $null
    }
}
#endregion

#region GUI Definition (Base64-Bilder und XAML-Definition)
# Base64-codierte Icons
$infoIconBase64 = "iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAAABmJLR0QA/wD/AP+gvaeTAAAA/0lEQVRIie3UMUpDQRCA4S8+sNAygufQzgtYJK2FjWfQVxgQjyBYpPcEXsFGewutPILgBWy0ErEQJokbu8m+bPK2EDIwsOz+M/vPLsv+VxWdJc61cIpDrGEa7/iI9xgDDBtyVnERGns4wiZW04qXMYcIpfiHuMYtRnhNc7bzgFksYICDgoncx0RDXOEmtbuVB9jFZ7w/YL0g5EuD0IbD+G6MjSLA9ohGe1jGYx6gn1buJwZcYqkE0MFbbPdXAeAHP0WAdsG740aSK8CoCoB+HsALjVVwF4Bn9OoA3OOiav5ZHmCIvcT5p+MCk0hH0au+xkka/K/qziL+wG71BanPLOeRlsyFAAAAAElFTkSuQmCC"
$settingsIconBase64 = "iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAAABmJLR0QA/wD/AP+gvaeTAAABKklEQVRIie3UvUodURDG8Z8X/IooiI1YJIJFUmlh4QUsLAQbLyBWES+glZ3eQQJ5A0FsFC0s9R6SQoKFhSIhhYKoGItnL4hhc87ZXRH2Dwv7zTDzzOzMzrL3p3b28FTSdh4f8IQt/ArfwAN2mgloxWesoRsrOMMI5vGAmzLAUVL2HSb+IL9Qj7hglQSA32H7qDXgF+xE8D5qGQCm8Brn3iYAV3Ae57oPeAv9EfNndGUB8Ae3cb4zAXiG/YhZzQKgoZXQaQJwFe0mlnMAFmM8nwJwHG0hlx7geciWgRUspgA0eoKVXIDL0JctBUwitLEMaDeZ9AdqS3Z3cJgLcILLOB/nAuxFTH8uwHrEdOcAdGAJg+hDTw7ArzjXWchPMZwDUJMG6yDkE/iIL9jGbfi28ct6A0/+dyqXVR5uAAAAAElFTkSuQmCC"
$closeIconBase64 = "iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAAABmJLR0QA/wD/AP+gvaeTAAAAsklEQVRIie2UQQrCMBBFX6UHcO8B3HsecemNvJiIroXeoEcQKhSXmcG0JGmCUHThXwSSyfx5MCGw4d+RJ3JVYu4PeAJ7oByYXwEP4AZ0ClHlnEqN+ad1GFG8eEMAbr64RtSDMj4CWSQuKnQQ3nhqEoCTS25Q/uoetcEDuAIndVHqIIPgSo2m/Ldt1KAJKduoQRqSmEELVIkaNKAJV9sF2b30oZKpN8+BO3CYO+MNgx8PizXYq+hBOQAAAABJRU5ErkJggg=="

# WPF XAML als String
$xaml = @"
<Window 
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="SharePoint Online Management Tool" Height="1000" Width="1250"
    WindowStartupLocation="CenterScreen" ResizeMode="CanResize"
    Background="#f9f9f9">
    
    <Window.Resources>
        <!-- Allgemeine Farben im Windows 11 Stil -->
        <SolidColorBrush x:Key="PrimaryAccentBrush" Color="#0078d4"/>
        <SolidColorBrush x:Key="PrimaryAccentHoverBrush" Color="#106ebe"/>
        <SolidColorBrush x:Key="PrimaryAccentPressedBrush" Color="#005a9e"/>
        <SolidColorBrush x:Key="BackgroundBrush" Color="#f9f9f9"/>
        <SolidColorBrush x:Key="BorderBrush" Color="#e0e0e0"/>
        <SolidColorBrush x:Key="DisabledBrush" Color="#cccccc"/>
        <SolidColorBrush x:Key="TextBrush" Color="#202020"/>
        <SolidColorBrush x:Key="SecondaryTextBrush" Color="#505050"/>
        
        <!-- Windows 11 Button Style -->
        <Style TargetType="Button">
            <Setter Property="Padding" Value="12,7" />
            <Setter Property="Margin" Value="6" />
            <Setter Property="Background" Value="{StaticResource PrimaryAccentBrush}" />
            <Setter Property="Foreground" Value="White" />
            <Setter Property="BorderThickness" Value="0" />
            <Setter Property="FontSize" Value="13" />
            <Setter Property="FontWeight" Value="Medium" />
            <Setter Property="Height" Value="36" />
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" 
                                BorderThickness="{TemplateBinding BorderThickness}" 
                                BorderBrush="{TemplateBinding BorderBrush}"
                                CornerRadius="4">
                            <Border.Effect>
                                <DropShadowEffect ShadowDepth="1" BlurRadius="3" Opacity="0.2" Direction="270" />
                            </Border.Effect>
                            <ContentPresenter HorizontalAlignment="Center" 
                                              VerticalAlignment="Center"
                                              Margin="{TemplateBinding Padding}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="{StaticResource PrimaryAccentHoverBrush}"/>
                                <Setter Property="Cursor" Value="Hand"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter Property="Background" Value="{StaticResource PrimaryAccentPressedBrush}"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Background" Value="{StaticResource DisabledBrush}"/>
                                <Setter Property="Foreground" Value="#666666"/>
                                <Setter Property="Effect" Value="{x:Null}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        
        <!-- Windows 11 Link Button Style -->
        <Style x:Key="LinkButtonStyle" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
            <Setter Property="Background" Value="Transparent" />
            <Setter Property="Foreground" Value="{StaticResource PrimaryAccentBrush}" />
            <Setter Property="BorderThickness" Value="0" />
            <Setter Property="Padding" Value="5,2" />
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" Padding="{TemplateBinding Padding}">
                            <ContentPresenter VerticalAlignment="Center" />
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Foreground" Value="{StaticResource PrimaryAccentHoverBrush}"/>
                                <Setter Property="TextBlock.TextDecorations" Value="Underline"/>
                                <Setter Property="Cursor" Value="Hand"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Foreground" Value="{StaticResource DisabledBrush}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        
        <!-- Windows 11 TextBox Style -->
        <Style TargetType="TextBox">
            <Setter Property="Padding" Value="10,8" />
            <Setter Property="Margin" Value="6" />
            <Setter Property="BorderThickness" Value="1" />
            <Setter Property="BorderBrush" Value="{StaticResource BorderBrush}" />
            <Setter Property="Background" Value="White" />
            <Setter Property="Height" Value="36" />
            <Setter Property="FontSize" Value="13" />
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TextBox">
                        <Border Background="{TemplateBinding Background}" 
                                BorderThickness="{TemplateBinding BorderThickness}" 
                                BorderBrush="{TemplateBinding BorderBrush}"
                                CornerRadius="4">
                            <ScrollViewer x:Name="PART_ContentHost" Padding="{TemplateBinding Padding}"
                                          VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsFocused" Value="True">
                                <Setter Property="BorderBrush" Value="{StaticResource PrimaryAccentBrush}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        
        <!-- PasswordBox Style -->
        <Style TargetType="PasswordBox">
            <Setter Property="Padding" Value="10,8" />
            <Setter Property="Margin" Value="6" />
            <Setter Property="BorderThickness" Value="1" />
            <Setter Property="BorderBrush" Value="{StaticResource BorderBrush}" />
            <Setter Property="Background" Value="White" />
            <Setter Property="Height" Value="36" />
            <Setter Property="FontSize" Value="13" />
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="PasswordBox">
                        <Border Background="{TemplateBinding Background}" 
                                BorderThickness="{TemplateBinding BorderThickness}" 
                                BorderBrush="{TemplateBinding BorderBrush}"
                                CornerRadius="4">
                            <ScrollViewer x:Name="PART_ContentHost" Padding="{TemplateBinding Padding}"
                                          VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsFocused" Value="True">
                                <Setter Property="BorderBrush" Value="{StaticResource PrimaryAccentBrush}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        
        <!-- ComboBox Style -->
        <Style TargetType="ComboBox">
            <Setter Property="Padding" Value="10,6" />
            <Setter Property="Margin" Value="6" />
            <Setter Property="BorderThickness" Value="1" />
            <Setter Property="BorderBrush" Value="{StaticResource BorderBrush}" />
            <Setter Property="Background" Value="White" />
            <Setter Property="Height" Value="36" />
            <Setter Property="FontSize" Value="13" />
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBox">
                        <Grid>
                            <ToggleButton x:Name="ToggleButton" 
                                          BorderBrush="{TemplateBinding BorderBrush}"
                                          Background="{TemplateBinding Background}"
                                          IsChecked="{Binding Path=IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}"
                                          Focusable="false">
                                <ToggleButton.Template>
                                    <ControlTemplate TargetType="ToggleButton">
                                        <Border x:Name="Border" 
                                                BorderThickness="{TemplateBinding BorderThickness}"
                                                BorderBrush="{TemplateBinding BorderBrush}"
                                                Background="{TemplateBinding Background}"
                                                CornerRadius="4">
                                            <Grid>
                                                <Grid.ColumnDefinitions>
                                                    <ColumnDefinition Width="*" />
                                                    <ColumnDefinition Width="Auto" />
                                                </Grid.ColumnDefinitions>
                                                <ContentPresenter Margin="{TemplateBinding Padding}"
                                                                  HorizontalAlignment="Left"
                                                                  VerticalAlignment="Center" />
                                                <Path x:Name="Arrow" Grid.Column="1"
                                                      Fill="{StaticResource SecondaryTextBrush}"
                                                      HorizontalAlignment="Center"
                                                      VerticalAlignment="Center"
                                                      Data="M0,0 L5,5 L10,0 Z"
                                                      Margin="0,0,10,0"/>
                                            </Grid>
                                        </Border>
                                        <ControlTemplate.Triggers>
                                            <Trigger Property="IsMouseOver" Value="True">
                                                <Setter Property="BorderBrush" Value="{StaticResource PrimaryAccentBrush}"/>
                                            </Trigger>
                                            <Trigger Property="IsChecked" Value="True">
                                                <Setter Property="BorderBrush" Value="{StaticResource PrimaryAccentBrush}"/>
                                            </Trigger>
                                        </ControlTemplate.Triggers>
                                    </ControlTemplate>
                                </ToggleButton.Template>
                            </ToggleButton>
                            <Popup IsOpen="{TemplateBinding IsDropDownOpen}" 
                                   Placement="Bottom"
                                   PopupAnimation="Slide"
                                   AllowsTransparency="True">
                                <Border BorderThickness="1" 
                                        BorderBrush="{StaticResource BorderBrush}"
                                        Background="White"
                                        CornerRadius="4"
                                        Margin="0,2,0,0">
                                    <Border.Effect>
                                        <DropShadowEffect ShadowDepth="2" BlurRadius="8" Opacity="0.2" Direction="270" />
                                    </Border.Effect>
                                    <ScrollViewer MaxHeight="250">
                                        <ItemsPresenter KeyboardNavigation.DirectionalNavigation="Contained" />
                                    </ScrollViewer>
                                </Border>
                            </Popup>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        
        <!-- TabControl Style -->
        <Style TargetType="TabControl">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="0"/>
        </Style>
        
        <!-- TabItem Style -->
        <Style TargetType="TabItem">
            <Setter Property="Padding" Value="16,10" />
            <Setter Property="Margin" Value="2,0,2,0" />
            <Setter Property="FontSize" Value="13" />
            <Setter Property="FontWeight" Value="Medium" />
            <Setter Property="Background" Value="Transparent" />
            <Setter Property="BorderThickness" Value="0" />
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TabItem">
                        <Grid>
                            <Border x:Name="Border"
                                    Background="{TemplateBinding Background}"
                                    BorderThickness="0,0,0,2"
                                    BorderBrush="Transparent"
                                    Margin="{TemplateBinding Margin}">
                                <ContentPresenter ContentSource="Header"
                                                  VerticalAlignment="Center"
                                                  HorizontalAlignment="Center"
                                                  Margin="{TemplateBinding Padding}"
                                                  RecognizesAccessKey="True"/>
                            </Border>
                        </Grid>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Border" Property="BorderBrush" Value="#c7c7c7"/>
                            </Trigger>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="Border" Property="BorderBrush" Value="{StaticResource PrimaryAccentBrush}"/>
                                <Setter Property="Foreground" Value="{StaticResource PrimaryAccentBrush}"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Foreground" Value="{StaticResource DisabledBrush}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        
        <!-- GroupBox Style -->
        <Style TargetType="GroupBox">
            <Setter Property="Padding" Value="12" />
            <Setter Property="Margin" Value="6" />
            <Setter Property="BorderThickness" Value="1" />
            <Setter Property="BorderBrush" Value="{StaticResource BorderBrush}" />
            <Setter Property="Background" Value="White" />
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="GroupBox">
                        <Grid>
                            <Border CornerRadius="6"
                                    Background="{TemplateBinding Background}"
                                    BorderBrush="{TemplateBinding BorderBrush}"
                                    BorderThickness="{TemplateBinding BorderThickness}">
                                <Border.Effect>
                                    <DropShadowEffect ShadowDepth="1" BlurRadius="4" Opacity="0.1" Direction="270" />
                                </Border.Effect>
                            </Border>
                            <DockPanel>
                                <Border DockPanel.Dock="Top" 
                                        Background="{TemplateBinding Background}" 
                                        Padding="10,0,0,0">
                                    <ContentPresenter ContentSource="Header" 
                                                      TextBlock.FontWeight="Medium" 
                                                      TextBlock.FontSize="13"
                                                      TextBlock.Foreground="{StaticResource TextBrush}"
                                                      RecognizesAccessKey="True" />
                                </Border>
                                <ContentPresenter Margin="{TemplateBinding Padding}" />
                            </DockPanel>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        
        <!-- DataGrid Styles -->
        <Style TargetType="DataGrid">
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="BorderBrush" Value="{StaticResource BorderBrush}"/>
            <Setter Property="Background" Value="White"/>
            <Setter Property="RowBackground" Value="White"/>
            <Setter Property="AlternatingRowBackground" Value="#f5f5f5"/>
            <Setter Property="HorizontalGridLinesBrush" Value="#e5e5e5"/>
            <Setter Property="VerticalGridLinesBrush" Value="#e5e5e5"/>
            <Setter Property="RowHeight" Value="36"/>
            <Setter Property="HeadersVisibility" Value="Column"/>
        </Style>
        
        <Style TargetType="DataGridColumnHeader">
            <Setter Property="Background" Value="#f0f0f0"/>
            <Setter Property="Foreground" Value="{StaticResource TextBrush}"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Padding" Value="10,8"/>
            <Setter Property="BorderThickness" Value="0,0,1,1"/>
            <Setter Property="BorderBrush" Value="#e0e0e0"/>
        </Style>
        
        <Style TargetType="DataGridRow">
            <Setter Property="Margin" Value="0"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#e6f2fa"/>
                </Trigger>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="#cce4f7"/>
                    <Setter Property="Foreground" Value="{StaticResource TextBrush}"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        
        <Style TargetType="DataGridCell">
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="10,8"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="DataGridCell">
                        <Border Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}">
                            <ContentPresenter VerticalAlignment="Center" />
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        
        <!-- CheckBox Style -->
        <Style TargetType="CheckBox">
            <Setter Property="Margin" Value="6"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
        </Style>
        
        <!-- RadioButton Style -->
        <Style TargetType="RadioButton">
            <Setter Property="Margin" Value="6"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
        </Style>
        
        <!-- ScrollViewer Style -->
        <Style TargetType="ScrollViewer">
            <Setter Property="Padding" Value="0"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ScrollViewer">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="*"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>
                            
                            <ScrollContentPresenter Grid.Column="0" Grid.Row="0"/>
                            
                            <ScrollBar x:Name="PART_VerticalScrollBar"
                                       Grid.Column="1" Grid.Row="0"
                                       Value="{TemplateBinding VerticalOffset}"
                                       Maximum="{TemplateBinding ScrollableHeight}"
                                       ViewportSize="{TemplateBinding ViewportHeight}"
                                       Visibility="{TemplateBinding ComputedVerticalScrollBarVisibility}"/>
                            
                            <ScrollBar x:Name="PART_HorizontalScrollBar"
                                       Orientation="Horizontal"
                                       Grid.Column="0" Grid.Row="1"
                                       Value="{TemplateBinding HorizontalOffset}"
                                       Maximum="{TemplateBinding ScrollableWidth}"
                                       ViewportSize="{TemplateBinding ViewportWidth}"
                                       Visibility="{TemplateBinding ComputedHorizontalScrollBarVisibility}"/>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        
        <!-- ScrollBar Style für Windows 11 -->
        <Style TargetType="ScrollBar">
            <Setter Property="Width" Value="12"/>
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ScrollBar">
                        <Grid x:Name="Bg" Background="{TemplateBinding Background}" SnapsToDevicePixels="true">
                            <Border Background="Transparent" Margin="-1"/>
                            <Track x:Name="PART_Track"
                                   IsDirectionReversed="true"
                                   IsEnabled="{TemplateBinding IsMouseOver}">
                                <Track.DecreaseRepeatButton>
                                    <RepeatButton Command="ScrollBar.PageUpCommand" Opacity="0" />
                                </Track.DecreaseRepeatButton>
                                <Track.Thumb>
                                    <Thumb x:Name="Thumb" 
                                           Background="#757575"
                                           Margin="2,0">
                                        <Thumb.Template>
                                            <ControlTemplate TargetType="Thumb">
                                                <Rectangle Fill="{TemplateBinding Background}" 
                                                           RadiusX="6" RadiusY="6" 
                                                           Opacity="0.6"/>
                                            </ControlTemplate>
                                        </Thumb.Template>
                                    </Thumb>
                                </Track.Thumb>
                                <Track.IncreaseRepeatButton>
                                    <RepeatButton Command="ScrollBar.PageDownCommand" Opacity="0" />
                                </Track.IncreaseRepeatButton>
                            </Track>
                        </Grid>
                        <ControlTemplate.Triggers>
                            <Trigger Property="Orientation" Value="Horizontal">
                                <Setter Property="Width" Value="Auto" />
                                <Setter Property="Height" Value="12" />
                                <Setter TargetName="Thumb" Property="Margin" Value="0,2" />
                            </Trigger>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Thumb" Property="Background" Value="#606060"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        
        <!-- Header TextBlock Style -->
        <Style x:Key="HeaderTextBlockStyle" TargetType="TextBlock">
            <Setter Property="FontSize" Value="14" />
            <Setter Property="FontWeight" Value="SemiBold" />
            <Setter Property="Margin" Value="6" />
            <Setter Property="Foreground" Value="{StaticResource TextBrush}" />
        </Style>
    </Window.Resources>

    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <!-- Hauptmenü mit Glaseffekt -->
        <DockPanel Grid.Row="0" LastChildFill="False" Background="{StaticResource PrimaryAccentBrush}" x:Name="dockMainHeader">
            <Border CornerRadius="0,0,8,8" Background="Transparent" Padding="2" Margin="0,0,0,-4">
                <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                    <TextBlock Text="SharePoint Online Management Tool" 
                               FontSize="18" FontWeight="SemiBold" Foreground="White"
                               VerticalAlignment="Center" Margin="15,12"/>
                </StackPanel>
            </Border>
            <Button x:Name="btnClose" DockPanel.Dock="Right" 
                    Background="Transparent" Foreground="White" 
                    BorderThickness="0" Width="42" Height="42" 
                    Margin="5" Padding="0" ToolTip="Schließen">
                <Image x:Name="imgClose" Width="18" Height="18"/>
            </Button>
        </DockPanel>
        
        <!-- Hauptinhalt mit Glaseffekt -->
        <TabControl Grid.Row="1" Margin="14,8,14,14" x:Name="tabMain">
            <!-- Verbindungstab -->
            <TabItem Header="Verbindung">
                <Grid Margin="10">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    
                    <GroupBox Grid.Row="0" Header="SharePoint Online Verbindung">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="Auto"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>
                            
                            <TextBlock Grid.Row="0" Grid.Column="0" Text="Admin-URL:" 
                                       VerticalAlignment="Center" Margin="5" Width="120"/>
                            <TextBox Grid.Row="0" Grid.Column="1" x:Name="txtAdminUrl" 
                                     ToolTip="SharePoint Admin URL (https://tenant-admin.sharepoint.com)" 
                                     TabIndex="0"/>
                            
                            <TextBlock Grid.Row="1" Grid.Column="0" Text="Benutzername:" 
                                       VerticalAlignment="Center" Margin="5"/>
                            <TextBox Grid.Row="1" Grid.Column="1" x:Name="txtUsername" 
                                     ToolTip="Benutzername (admin@tenant.onmicrosoft.com)" 
                                     TabIndex="1"/>
                            
                            <TextBlock Grid.Row="2" Grid.Column="0" Text="Passwort:" 
                                       VerticalAlignment="Center" Margin="5"/>
                            <PasswordBox Grid.Row="2" Grid.Column="1" x:Name="txtPassword" 
                                         TabIndex="2"/>
                            
                            <StackPanel Grid.Row="3" Grid.Column="0" Grid.ColumnSpan="2" 
                                        Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0">
                                <Button x:Name="btnConnect" Content="Verbinden" Width="150" TabIndex="3"/>
                                <Button x:Name="btnDisconnect" Content="Trennen" Width="150" IsEnabled="False" TabIndex="4"/>
                            </StackPanel>
                        </Grid>
                    </GroupBox>
                    
                    <GroupBox Grid.Row="1" Header="Modulstatus" Margin="0,10,0,0">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>
                            
                            <TextBlock Grid.Row="0" Text="Erforderliche Module:" Margin="5"/>
                            <ListView Grid.Row="1" x:Name="lvModules" Height="150" Margin="5" BorderThickness="1" 
                                      BorderBrush="{StaticResource BorderBrush}" Background="White">
                                <ListView.View>
                                    <GridView>
                                        <GridViewColumn Header="Modul" Width="220" DisplayMemberBinding="{Binding Name}"/>
                                        <GridViewColumn Header="Status" Width="150" DisplayMemberBinding="{Binding Status}"/>
                                        <GridViewColumn Header="Version" Width="100" DisplayMemberBinding="{Binding Version}"/>
                                        <GridViewColumn Header="Neueste Version" Width="120" DisplayMemberBinding="{Binding LatestVersion}"/>
                                        <GridViewColumn Header="Notizen" Width="250" DisplayMemberBinding="{Binding Notes}"/>
                                    </GridView>
                                </ListView.View>
                            </ListView>
                            
                            <Button Grid.Row="2" Content="Module installieren/aktualisieren" 
                                    HorizontalAlignment="Right" Width="250" Margin="0,10,5,0" 
                                    x:Name="btnInstallModules"/>
                        </Grid>
                    </GroupBox>
                    
                    <GroupBox Grid.Row="2" Header="Verbindungsstatus" Margin="0,10,0,0">
                        <TextBox x:Name="txtConnectionStatus" IsReadOnly="True" 
                                 TextWrapping="Wrap" Height="150" 
                                 VerticalScrollBarVisibility="Auto"/>
                    </GroupBox>
                </Grid>
            </TabItem>
            
            <!-- Websitesammlungstab -->
            <TabItem Header="Websitesammlungen" x:Name="tabSiteCollections">
                <Grid Margin="10">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <StackPanel Grid.Row="0" Orientation="Horizontal">
                        <Button x:Name="btnRefreshSites" Content="Aktualisieren" Width="120"/>
                        <Button x:Name="btnAddSite" Content="Neue Websitesammlung" Width="180"/>
                        <TextBox x:Name="txtSearchSites" Width="300" Margin="20,6,6,6" 
                                 ToolTip="Nach Websitesammlungen suchen"/>
                        <Button x:Name="btnSearchSites" Content="Suchen" Width="80"/>
                    </StackPanel>
                    
                    <DataGrid Grid.Row="1" x:Name="dgSites" Margin="0,10,0,0" 
                              AutoGenerateColumns="False" IsReadOnly="True" 
                              SelectionMode="Single">
                        <DataGrid.Columns>
                            <DataGridTextColumn Header="Titel" Binding="{Binding Title}" Width="200"/>
                            <DataGridTextColumn Header="URL" Binding="{Binding Url}" Width="300"/>
                            <DataGridTextColumn Header="Vorlage" Binding="{Binding Template}" Width="120"/>
                            <DataGridTextColumn Header="Besitzer" Binding="{Binding Owner}" Width="180"/>
                            <DataGridTextColumn Header="Speichernutzung (MB)" Binding="{Binding StorageUsageCurrent}" Width="150"/>
                        </DataGrid.Columns>
                    </DataGrid>
                    
                    <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,10,0,0">
                        <Button x:Name="btnEditSite" Content="Bearbeiten" Width="120"/>
                        <Button x:Name="btnDeleteSite" Content="Löschen" Width="120"/>
                        <Button x:Name="btnSitePermissions" Content="Berechtigungen" Width="120"/>
                    </StackPanel>
                </Grid>
            </TabItem>
            
            <!-- Listentab -->
            <TabItem Header="Listen und Bibliotheken" x:Name="tabLists">
                <Grid Margin="10">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <GroupBox Grid.Row="0" Header="Websiteauswahl">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="Auto"/>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            
                            <TextBlock Grid.Column="0" Text="Website-URL:" VerticalAlignment="Center" Margin="5" Width="100"/>
                            <ComboBox Grid.Column="1" x:Name="cmbSiteUrl" Padding="5" Margin="5"/>
                            <Button Grid.Column="2" Content="Laden" x:Name="btnLoadLists" Width="100"/>
                        </Grid>
                    </GroupBox>
                    
                    <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,10,0,0">
                        <Button x:Name="btnRefreshLists" Content="Aktualisieren" Width="120"/>
                        <Button x:Name="btnAddList" Content="Neue Liste" Width="120"/>
                        <Button x:Name="btnAddLibrary" Content="Neue Bibliothek" Width="150"/>
                    </StackPanel>
                    
                    <DataGrid Grid.Row="2" x:Name="dgLists" Margin="0,10,0,0" 
                              AutoGenerateColumns="False" IsReadOnly="True" 
                              SelectionMode="Single">
                        <DataGrid.Columns>
                            <DataGridTextColumn Header="Titel" Binding="{Binding Title}" Width="200"/>
                            <DataGridTextColumn Header="Typ" Binding="{Binding BaseTemplate}" Width="100"/>
                            <DataGridTextColumn Header="Elemente" Binding="{Binding ItemCount}" Width="80"/>
                            <DataGridTextColumn Header="Erstellt am" Binding="{Binding Created}" Width="150"/>
                            <DataGridTextColumn Header="Beschreibung" Binding="{Binding Description}" Width="250"/>
                            <DataGridCheckBoxColumn Header="Versionsverwaltung" Binding="{Binding EnableVersioning}" Width="130"/>
                        </DataGrid.Columns>
                    </DataGrid>
                    
                    <StackPanel Grid.Row="3" Orientation="Horizontal" Margin="0,10,0,0">
                        <Button x:Name="btnViewList" Content="Anzeigen" Width="120"/>
                        <Button x:Name="btnEditList" Content="Bearbeiten" Width="120"/>
                        <Button x:Name="btnDeleteList" Content="Löschen" Width="120"/>
                        <Button x:Name="btnListPermissions" Content="Berechtigungen" Width="120"/>
                    </StackPanel>
                </Grid>
            </TabItem>
            
            <!-- Benutzertab -->
            <TabItem Header="Benutzer und Berechtigungen" x:Name="tabUsers">
                <Grid Margin="10">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <GroupBox Grid.Row="0" Header="Websiteauswahl">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="Auto"/>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            
                            <TextBlock Grid.Column="0" Text="Website-URL:" VerticalAlignment="Center" Margin="5" Width="100"/>
                            <ComboBox Grid.Column="1" x:Name="cmbUserSiteUrl" Padding="5" Margin="5"/>
                            <Button Grid.Column="2" Content="Laden" x:Name="btnLoadUsers" Width="100"/>
                        </Grid>
                    </GroupBox>
                    
                    <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,10,0,0">
                        <Button x:Name="btnRefreshUsers" Content="Aktualisieren" Width="120"/>
                        <Button x:Name="btnAddUser" Content="Benutzer hinzufügen" Width="150"/>
                        <Button x:Name="btnInviteGuest" Content="Gast einladen" Width="150"/>
                    </StackPanel>
                    
                    <DataGrid Grid.Row="2" x:Name="dgUsers" Margin="0,10,0,0" 
                              AutoGenerateColumns="False" IsReadOnly="True" 
                              SelectionMode="Single">
                        <DataGrid.Columns>
                            <DataGridTextColumn Header="Anzeigename" Binding="{Binding Title}" Width="200"/>
                            <DataGridTextColumn Header="E-Mail" Binding="{Binding Email}" Width="220"/>
                            <DataGridTextColumn Header="Benutzertyp" Binding="{Binding UserType}" Width="120"/>
                            <DataGridTextColumn Header="Gruppen" Binding="{Binding Groups}" Width="*"/>
                        </DataGrid.Columns>
                    </DataGrid>
                    
                    <StackPanel Grid.Row="3" Orientation="Horizontal" Margin="0,10,0,0">
                        <Button x:Name="btnModifyPermissions" Content="Berechtigungen ändern" Width="180"/>
                        <Button x:Name="btnRemoveUser" Content="Benutzer entfernen" Width="150"/>
                    </StackPanel>
                </Grid>
            </TabItem>
            
            <!-- Tenant-Einstellungstab -->
            <TabItem Header="Tenant-Einstellungen" x:Name="tabSettings">
                <ScrollViewer VerticalScrollBarVisibility="Auto">
                    <StackPanel Margin="10">
                        <!-- Freigabeeinstellungen -->
                        <GroupBox Header="Freigabeeinstellungen">
                            <Grid>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>
                                
                                <TextBlock Grid.Row="0" Text="Externe Freigabe" 
                                           Style="{StaticResource HeaderTextBlockStyle}"/>
                                
                                <StackPanel Grid.Row="1" Margin="5">
                                    <RadioButton x:Name="rbSharingDisabled" GroupName="SharingLevel" 
                                                Content="Deaktiviert (keine externe Freigabe)" Margin="5"/>
                                    <RadioButton x:Name="rbSharingExistingOnly" GroupName="SharingLevel" 
                                                Content="Nur vorhandene externe Benutzer" Margin="5"/>
                                    <RadioButton x:Name="rbSharingNewExternal" GroupName="SharingLevel" 
                                                Content="Neue und vorhandene externe Benutzer" Margin="5"/>
                                    <RadioButton x:Name="rbSharingAnonymous" GroupName="SharingLevel" 
                                                Content="Jeder (anonyme Links eingeschlossen)" Margin="5"/>
                                </StackPanel>
                                
                                <CheckBox Grid.Row="2" x:Name="chkRequireAcceptingAccountsToMatchInvitedAccount" 
                                          Content="Anmeldende E-Mail-Adresse muss mit eingeladener E-Mail-Adresse übereinstimmen" 
                                          Margin="5,10,5,5"/>
                                
                                <Button Grid.Row="3" Content="Freigabeeinstellungen speichern" 
                                        HorizontalAlignment="Right" Width="250" Margin="0,10,5,0" 
                                        x:Name="btnSaveSharing"/>
                            </Grid>
                        </GroupBox>
                        
                        <!-- OneDrive-Einstellungen -->
                        <GroupBox Header="OneDrive-Einstellungen" Margin="0,10,0,0">
                            <Grid>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>
                                
                                <TextBlock Grid.Row="0" Text="OneDrive-Synchronisierung" 
                                           Style="{StaticResource HeaderTextBlockStyle}"/>
                                
                                <StackPanel Grid.Row="1" Margin="5">
                                    <CheckBox x:Name="chkBlockMacSync" Content="Mac-Synchronisierung blockieren" Margin="5"/>
                                    <CheckBox x:Name="chkDisableReportProblemDialog" Content="Problemdialog deaktivieren" Margin="5"/>
                                    <StackPanel Orientation="Horizontal">
                                        <TextBlock Text="Speicherplatzlimit (GB):" VerticalAlignment="Center" Margin="5"/>
                                        <TextBox x:Name="txtOneDriveStorageQuota" Width="100" Margin="5"/>
                                    </StackPanel>
                                </StackPanel>
                                
                                <Button Grid.Row="2" Content="OneDrive-Einstellungen speichern" 
                                        HorizontalAlignment="Right" Width="250" Margin="0,10,5,0" 
                                        x:Name="btnSaveOneDrive"/>
                            </Grid>
                        </GroupBox>
                        
                        <!-- Benutzerumfangreiche Einstellungen -->
                        <GroupBox Header="Benutzerumfangreiche Einstellungen" Margin="0,10,0,0">
                            <Grid>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>
                                
                                <TextBlock Grid.Row="0" Text="Allgemeine Einstellungen" 
                                           Style="{StaticResource HeaderTextBlockStyle}"/>
                                
                                <StackPanel Grid.Row="1" Margin="5">
                                    <CheckBox x:Name="chkDisableCompanyWideSharingLinks" Content="Unternehmensweite Freigabelinks deaktivieren" Margin="5"/>
                                    <CheckBox x:Name="chkEnableGuestSignInAcceleration" Content="Beschleunigte Gastanmeldung aktivieren" Margin="5"/>
                                    <CheckBox x:Name="chkUseFindPeopleInPeoplePicker" Content="Personenauswahl für Personensuche nutzen" Margin="5"/>
                                    <CheckBox x:Name="chkNotificationsInSharePoint" Content="Benachrichtigungen in SharePoint aktivieren" Margin="5"/>
                                    <CheckBox x:Name="chkEnableRestrictedAccessControl" Content="Eingeschränkte Zugriffssteuerung aktivieren" Margin="5"/>
                                </StackPanel>
                                
                                <Button Grid.Row="2" Content="Allgemeine Einstellungen speichern" 
                                        HorizontalAlignment="Right" Width="250" Margin="0,10,5,0" 
                                        x:Name="btnSaveGeneralSettings"/>
                            </Grid>
                        </GroupBox>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>
            
            <!-- Protokolltab -->
            <TabItem Header="Protokoll" x:Name="tabLog">
                <Grid Margin="10">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    
                    <StackPanel Grid.Row="0" Orientation="Horizontal">
                        <Button x:Name="btnClearLog" Content="Protokoll löschen" Width="150"/>
                        <Button x:Name="btnSaveLog" Content="Protokoll speichern" Width="150"/>
                    </StackPanel>
                    
                    <TextBox Grid.Row="1" x:Name="txtLog" IsReadOnly="True" 
                             TextWrapping="Wrap" Margin="0,10,0,0" 
                             VerticalScrollBarVisibility="Auto" 
                             FontFamily="Consolas"/>
                </Grid>
            </TabItem>
        </TabControl>
        
        <!-- Statusleiste mit Glaseffekt -->
        <Border Grid.Row="2" Background="#f0f0f0" CornerRadius="0,8,0,0" Padding="1" Margin="-1">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                
                <TextBlock Grid.Column="0" x:Name="txtStatus" Text="Bereit" Margin="14,8" VerticalAlignment="Center"/>
                <Rectangle Grid.Column="1" Width="1" Height="16" Fill="#d0d0d0" Margin="4,0"/>
                <TextBlock Grid.Column="2" x:Name="txtConnectedUser" Text="Nicht verbunden" Margin="14,8" VerticalAlignment="Center"/>
                <TextBlock Grid.Column="3" Text="v0.0.1" Margin="14,8" VerticalAlignment="Center" Opacity="0.6"/>
            </Grid>
        </Border>
        
        <!-- Schatten für das Hauptfenster -->
        <Border Grid.RowSpan="3" BorderThickness="0" Margin="-5">
            <Border.Effect>
                <DropShadowEffect ShadowDepth="3" BlurRadius="12" Opacity="0.3" Direction="270"/>
            </Border.Effect>
        </Border>
    </Grid>
</Window>
"@
#endregion

#region Hilfsfunktionen
# Base64 in ImageSource umwandeln
function Convert-Base64ToImageSource {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Base64String
    )
    
    try {
        $bytes = [System.Convert]::FromBase64String($Base64String)
        $stream = New-Object System.IO.MemoryStream
        $stream.Write($bytes, 0, $bytes.Length)
        $stream.Seek(0, [System.IO.SeekOrigin]::Begin) | Out-Null
        
        $bitmap = New-Object System.Windows.Media.Imaging.BitmapImage
        $bitmap.BeginInit()
        $bitmap.StreamSource = $stream
        $bitmap.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
        $bitmap.EndInit()
        $bitmap.Freeze() # Wichtig für Thread-Sicherheit
        
        return $bitmap
    }
    catch {
        $errorMessage = ${_}
        Write-Log "Fehler beim Konvertieren des Base64-Strings zu ImageSource: $errorMessage" -Level "ERROR"
        return $null
    }
}

# GUI-Textausgabe sichern (Längenbegrenzung, Zeichenfilterung)
function Update-GuiText {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ControlName,
        [Parameter(Mandatory = $true)]
        [string]$Text,
        [int]$MaxLength = 10000
    )
    
    try {
        # Prüfen, ob Steuerelement existiert
        if (-not $script:GUIElements.ContainsKey($ControlName)) {
            Write-Log "Das Steuerelement '$ControlName' existiert nicht." -Level "WARNING"
            return $false
        }
        
        # Zeichenfilterung (nur druckbare ASCII-Zeichen)
        $filteredText = $Text -replace '[^\x20-\x7E\r\n]', '?'
        
        # Längenbegrenzung
        if ($filteredText.Length -gt $MaxLength) {
            $filteredText = $filteredText.Substring(0, $MaxLength - 50) + 
                "`r`n[...gekürzt...]`r`n" +
                $filteredText.Substring($filteredText.Length - 50, 50)
        }
        
        # Thread-sicher in GUI-Element schreiben
        $script:GUIElements[$ControlName].Dispatcher.Invoke(
            [Action]{
                $script:GUIElements[$ControlName].Text = $filteredText
            },
            "Normal"
        )
        
        return $true
    }
    catch {
        $errorMessage = ${_}
        Write-Log "Fehler beim Aktualisieren des GUI-Textes: $errorMessage" -Level "ERROR"
        
        # Fallback-Logging in Datei
        $fallbackFile = Join-Path $PSScriptRoot "Logs\GUI_Fallback_$(Get-Date -Format 'yyyy-MM-dd').log"
        try {
            Add-Content -Path $fallbackFile -Value "[$ControlName] $Text" -ErrorAction SilentlyContinue
        }
        catch {}
        
        return $false
    }
}
#endregion

#region GUI Initialisierung und Event-Handler
# GUI initialisieren
function Initialize-GUI {
    try {
        # XAML laden
        Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
        
        [xml]$xamlDoc = $xaml
        $reader = New-Object System.Xml.XmlNodeReader $xamlDoc
        
        # Versuche das Fenster zu laden
        try {
            $window = [Windows.Markup.XamlReader]::Load($reader)
            if ($null -eq $window) {
                Write-Log "Fehler beim Laden des XAML: Das zurückgegebene Fenster ist null." -Level "ERROR"
                return $null
            }
        }
        catch {
            $errorMessage = ${_}
            Write-Log "Fehler beim Laden des XAML: $errorMessage" -Level "ERROR"
            return $null
        }
        
        # Base64-Bilder laden
        $imgClose = $window.FindName("imgClose")
        if ($null -ne $imgClose) {
            $imgClose.Source = Convert-Base64ToImageSource $closeIconBase64
        }
        else {
            Write-Log "Warnung: Element 'imgClose' nicht gefunden" -Level "WARNING"
        }
        
        # Elementreferenzen speichern - nur relevante UI-Elemente finden, keine Template-Teile
        $script:GUIElements = @{} 
        
        # Liste der zu findenden Steuerelemente (vermeidet Warnungen für Template-Elemente)
        $controlsToFind = @(
            # Hauptfenster-Elemente
            "dockMainHeader", "tabMain", "txtStatus", "txtConnectedUser",
            
            # Verbindungs-Tab
            "txtAdminUrl", "txtUsername", "txtPassword", "btnConnect", "btnDisconnect",
            "lvModules", "btnInstallModules", "txtConnectionStatus",
            
            # Site-Collections Tab
            "tabSiteCollections", "dgSites", "btnRefreshSites", "btnAddSite",
            "txtSearchSites", "btnSearchSites", "btnEditSite", "btnDeleteSite", "btnSitePermissions",
            
            # Listen Tab
            "tabLists", "cmbSiteUrl", "btnLoadLists", "dgLists", "btnRefreshLists", 
            "btnAddList", "btnAddLibrary", "btnViewList", "btnEditList", "btnDeleteList", 
            "btnListPermissions",
            
            # Benutzer Tab
            "tabUsers", "cmbUserSiteUrl", "btnLoadUsers", "dgUsers", "btnRefreshUsers",
            "btnAddUser", "btnInviteGuest", "btnModifyPermissions", "btnRemoveUser",
            
            # Settings Tab
            "tabSettings", "rbSharingDisabled", "rbSharingExistingOnly", "rbSharingNewExternal",
            "rbSharingAnonymous", "chkRequireAcceptingAccountsToMatchInvitedAccount", "btnSaveSharing",
            "chkBlockMacSync", "chkDisableReportProblemDialog", "txtOneDriveStorageQuota", "btnSaveOneDrive",
            "chkDisableCompanyWideSharingLinks", "chkEnableGuestSignInAcceleration", 
            "chkUseFindPeopleInPeoplePicker", "chkNotificationsInSharePoint", 
            "chkEnableRestrictedAccessControl", "btnSaveGeneralSettings",
            
            # Log Tab
            "tabLog", "btnClearLog", "btnSaveLog", "txtLog",
            
            # Close Button
            "btnClose"
        )
        
        # Elemente suchen und speichern
        foreach ($controlName in $controlsToFind) {
            $element = $window.FindName($controlName)
            if ($null -ne $element) {
                $script:GUIElements[$controlName] = $element
            }
            else {
                # Nur wirklich benötigte Elemente verursachen eine Warnung
                if ($controlName -notmatch '^(PART_|Border$|Thumb$|Arrow$)') {
                    Write-Log "Element nicht gefunden: $controlName" -Level "WARNING"
                }
            }
        }
        
        # Prüfen ob wichtige Elemente gefunden wurden
        $requiredElements = @("tabSiteCollections", "tabLists", "tabUsers", "tabSettings", 
                              "btnConnect", "btnDisconnect", "btnInstallModules")
        $missingRequired = $false
        
        foreach ($elementName in $requiredElements) {
            if (-not $script:GUIElements.ContainsKey($elementName)) {
                Write-Log "Kritischer Fehler: Erforderliches Element '$elementName' wurde nicht gefunden." -Level "ERROR"
                $missingRequired = $true
            }
        }
        
        if ($missingRequired) {
            return $null
        }
        
        # Hilfe-Button im Header hinzufügen
        $dockHeader = $window.FindName("dockMainHeader")
        if ($dockHeader) {
            $btnHelp = New-Object System.Windows.Controls.Button
            $btnHelp.Content = "SPO-Cmdlets"
            $btnHelp.ToolTip = "SharePoint Online Cmdlet-Referenz anzeigen"
            $btnHelp.Margin = New-Object System.Windows.Thickness(5)
            $btnHelp.Padding = New-Object System.Windows.Thickness(10, 5, 10, 5)
            $btnHelp.Background = "Transparent"
            $btnHelp.Foreground = "White"
            $btnHelp.BorderThickness = 0
            $btnHelp.Add_Click```powershell
            $btnHelp.Add_Click({ Show-SPOCmdletReference })
            [System.Windows.Controls.DockPanel]::SetDock($btnHelp, "Right")
            $dockHeader.Children.Insert(1, $btnHelp)
        }
        
        # Tab-Steuerungselemente deaktivieren, bis eine Verbindung hergestellt ist
        $script:GUIElements["tabSiteCollections"].IsEnabled = $false
        $script:GUIElements["tabLists"].IsEnabled = $false
        $script:GUIElements["tabUsers"].IsEnabled = $false
        $script:GUIElements["tabSettings"].IsEnabled = $false
        
        # Ereignishandler für Schaltflächen hinzufügen
        try {
            # Close-Button muss immer funktionieren
            $script:GUIElements["btnClose"].Add_Click({ $window.Close() })
            
            # Weitere Ereignishandler registrieren
            $script:GUIElements["btnConnect"].Add_Click({ Connect-ToSharePointOnline })
            $script:GUIElements["btnDisconnect"].Add_Click({ Disconnect-FromSharePointOnline })
            $script:GUIElements["btnInstallModules"].Add_Click({ Install-RequiredModules })
            $script:GUIElements["btnRefreshSites"].Add_Click({ Refresh-SiteCollections })
            $script:GUIElements["btnAddSite"].Add_Click({ Show-NewSiteDialog })
            $script:GUIElements["btnLoadLists"].Add_Click({ Load-ListsFromSite })
            $script:GUIElements["btnLoadUsers"].Add_Click({ Load-UsersFromSite })
            $script:GUIElements["btnClearLog"].Add_Click({ Clear-Log })
            $script:GUIElements["btnSaveLog"].Add_Click({ Save-Log })
            $script:GUIElements["btnSaveSharing"].Add_Click({ Save-SharingSettings })
        }
        catch {
            $errorMessage = ${_}
            Write-Log "Fehler beim Hinzufügen von Event-Handlern: $errorMessage" -Level "ERROR"
            # Wir brechen hier nicht ab, da einige Funktionen dennoch nutzbar sein könnten
        }
        
        # Bei Initialisierung prüfen und Module anzeigen
        Check-RequiredModules
        
        # Protokoll im Hintergrund aktualisieren
        try {
            $timer = New-Object System.Windows.Threading.DispatcherTimer
            $timer.Interval = [TimeSpan]::FromSeconds(5)
            $timer.Add_Tick({
                Update-LogDisplay
            })
            $timer.Start()
        }
        catch {
            $errorMessage = ${_}
            Write-Log "Fehler beim Einrichten des Log-Update-Timers: $errorMessage" -Level "WARNING"
            # Kein kritischer Fehler, Protokoll kann immer noch manuell aktualisiert werden
        }
        
        Write-Log "GUI erfolgreich initialisiert." -Level "SUCCESS"
        return $window
    }
    catch {
        $errorMessage = ${_}
        Write-Log "Kritischer Fehler bei der GUI-Initialisierung: $errorMessage" -Level "ERROR"
        return $null
    }
}
#endregion

#region GUI Event-Handler-Funktionen
# Protokollanzeige aktualisieren
function Update-LogDisplay {
    try {
        # Pfad zur heutigen Logdatei
        $logDir = Join-Path $PSScriptRoot "Logs"
        $logFile = Join-Path $logDir "easySPO_$(Get-Date -Format 'yyyy-MM-dd').log"
        
        # Prüfen, ob die Datei existiert
        if (Test-Path $logFile) {
            # Die letzten Einträge der Logdatei holen (begrenzt auf eine bestimmte Anzahl von Zeilen)
            $logContent = Get-Content -Path $logFile -Tail 100 | Out-String
            
            # GUI-Steuerelement aktualisieren
            Update-GuiText -ControlName "txtLog" -Text $logContent
        }
    }
    catch {
        # Fehler beim Aktualisieren des Protokolls werden leise geschluckt,
        # um keine Endlosschleife bei Fehlern zu verursachen
        Write-Debug "Fehler beim Aktualisieren des Protokolls: $_"
    }
}

# Log-Anzeige leeren
function Clear-Log {
    try {
        # Protokollanzeige leeren
        Update-GuiText -ControlName "txtLog" -Text ""
        Write-Log "Protokollanzeige wurde geleert." -Level "INFO"
    }
    catch {
        $errorMessage = ${_}
        Write-Log "Fehler beim Löschen des Protokolls: $errorMessage" -Level "ERROR"
    }
}

# Protokoll speichern
function Save-Log {
    try {
        $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
        $saveDialog.Title = "Protokoll speichern"
        $saveDialog.Filter = "Protokolldateien (*.log)|*.log|Textdateien (*.txt)|*.txt|Alle Dateien (*.*)|*.*"
        $saveDialog.DefaultExt = "log"
        $saveDialog.FileName = "easySPO_Protokoll_$(Get-Date -Format 'yyyy-MM-dd')"
        
        $result = $saveDialog.ShowDialog()
        
        if ($result -eq $true) {
            # Prüfen, ob txtLog existiert
            if ($script:GUIElements.ContainsKey("txtLog")) {
                $logContent = $script:GUIElements["txtLog"].Text
                Set-Content -Path $saveDialog.FileName -Value $logContent -Encoding UTF8
                Write-Log "Protokoll wurde gespeichert unter: $($saveDialog.FileName)" -Level "SUCCESS"
            }
            else {
                # Alternativ Logdatei kopieren, wenn Textfeld nicht verfügbar
                $logDir = Join-Path $PSScriptRoot "Logs"
                $logFile = Join-Path $logDir "easySPO_$(Get-Date -Format 'yyyy-MM-dd').log"
                
                if (Test-Path $logFile) {
                    Copy-Item -Path $logFile -Destination $saveDialog.FileName -Force
                    Write-Log "Protokoll wurde gespeichert unter: $($saveDialog.FileName)" -Level "SUCCESS"
                }
                else {
                    Write-Log "Keine Protokolldatei gefunden zum Speichern." -Level "WARNING"
                }
            }
        }
    }
    catch {
        $errorMessage = ${_}
        Write-Log "Fehler beim Speichern des Protokolls: $errorMessage" -Level "ERROR"
    }
}

# SharePoint Online Cmdlet-Referenz anzeigen
function Show-SPOCmdletReference {
    try {
        $documentationUrl = "https://learn.microsoft.com/powershell/sharepoint/sharepoint-pnp/sharepoint-pnp-cmdlets"
        Start-Process $documentationUrl
        Write-Log "SharePoint Online Cmdlet-Referenz wurde geöffnet." -Level "INFO"
    }
    catch {
        $errorMessage = ${_}
        Write-Log "Fehler beim Öffnen der SPO Cmdlet-Referenz: $errorMessage" -Level "ERROR"
    }
}

# Installation der erforderlichen Module
function Install-RequiredModules {
    try {
        Write-Log "Starte Installation/Aktualisierung der erforderlichen Module..." -Level "INFO"
        Update-GuiText -ControlName "txtStatus" -Text "Module werden installiert..."
        
        # PnP.PowerShell Module installieren
        $moduleName = "PnP.PowerShell"
        
        # Prüfe, ob Modul bereits installiert ist
        if (Get-Module -Name $moduleName -ListAvailable) {
            # Prüfe, ob Update verfügbar ist
            $installedVersion = (Get-Module -Name $moduleName -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1).Version
            $onlineVersion = Find-Module -Name $moduleName -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Version
            
            if ($onlineVersion -gt $installedVersion) {
                Write-Log "Aktualisiere $moduleName von Version $installedVersion auf $onlineVersion" -Level "INFO"
                Update-Module -Name $moduleName -Force -Confirm:$false
                Write-Log "Modul $moduleName wurde aktualisiert." -Level "SUCCESS"
            }
            else {
                Write-Log "Modul $moduleName ist bereits auf dem neuesten Stand (Version $installedVersion)." -Level "INFO"
            }
        }
        else {
            # Modul installieren
            Write-Log "Installiere $moduleName..." -Level "INFO"
            Install-Module -Name $moduleName -Force -AllowClobber -Scope CurrentUser
            Write-Log "Modul $moduleName wurde installiert." -Level "SUCCESS"
        }
        
        # Modulstatus aktualisieren
        Check-RequiredModules
        
        Write-Log "Installation/Aktualisierung der Module abgeschlossen." -Level "SUCCESS"
        Update-GuiText -ControlName "txtStatus" -Text "Module installiert"
        
        # Kurzer Timer, um Status zurückzusetzen
        $resetTimer = New-Object System.Windows.Threading.DispatcherTimer
        $resetTimer.Interval = [TimeSpan]::FromSeconds(3)
        $resetTimer.Add_Tick({
            Update-GuiText -ControlName "txtStatus" -Text "Bereit"
            $this.Stop()
        })
        $resetTimer.Start()
    }
    catch {
        $errorMessage = ${_}
        Write-Log "Fehler bei der Modulinstallation: $errorMessage" -Level "ERROR"
        Update-GuiText -ControlName "txtStatus" -Text "Fehler bei der Modulinstallation"
    }
}

# Freigabeeinstellungen speichern
function Save-SharingSettings {
    try {
        Write-Log "Speichere Freigabeeinstellungen..." -Level "INFO"
        
        # Ermitteln des ausgewählten Freigabelevels
        $sharingLevel = "Disabled" # Standardwert
        
        if ($script:GUIElements["rbSharingDisabled"].IsChecked) {
            $sharingLevel = "Disabled"
        }
        elseif ($script:GUIElements["rbSharingExistingOnly"].IsChecked) {
            $sharingLevel = "ExistingExternalUserSharingOnly"
        }
        elseif ($script:GUIElements["rbSharingNewExternal"].IsChecked) {
            $sharingLevel = "ExternalUserSharingOnly"
        }
        elseif ($script:GUIElements["rbSharingAnonymous"].IsChecked) {
            $sharingLevel = "ExternalUserAndGuestSharing"
        }
        
        # Überprüfen, ob eine SharePoint-Verbindung besteht
        if (-not $script:isConnected) {
            Write-Log "Keine Verbindung zu SharePoint Online. Einstellungen können nicht gespeichert werden." -Level "WARNING"
            [System.Windows.MessageBox]::Show(
                "Es besteht keine Verbindung zu SharePoint Online. Bitte stellen Sie zuerst eine Verbindung her.", 
                "Freigabeeinstellungen speichern", 
                [System.Windows.MessageBoxButton]::OK, 
                [System.Windows.MessageBoxImage]::Warning
            )
            return
        }
        
        # Freigabeeinstellungen anwenden
        $success = Set-SPOSharingCapability -SharingCapability $sharingLevel
        
        if ($success) {
            Write-Log "Freigabeeinstellungen erfolgreich gespeichert (Sharing: $sharingLevel)." -Level "SUCCESS"
            [System.Windows.MessageBox]::Show(
                "Die Freigabeeinstellungen wurden erfolgreich gespeichert.", 
                "Freigabeeinstellungen speichern", 
                [System.Windows.MessageBoxButton]::OK, 
                [System.Windows.MessageBoxImage]::Information
            )
        }
        else {
            Write-Log "Fehler beim Speichern der Freigabeeinstellungen." -Level "ERROR"
            [System.Windows.MessageBox]::Show(
                "Beim Speichern der Freigabeeinstellungen ist ein Fehler aufgetreten. Bitte überprüfen Sie das Protokoll.", 
                "Freigabeeinstellungen speichern", 
                [System.Windows.MessageBoxButton]::OK, 
                [System.Windows.MessageBoxImage]::Error
            )
        }
    }
    catch {
        $errorMessage = ${_}
        Write-Log "Fehler beim Speichern der Freigabeeinstellungen: $errorMessage" -Level "ERROR"
        [System.Windows.MessageBox]::Show(
            "Beim Speichern der Freigabeeinstellungen ist ein Fehler aufgetreten: $errorMessage", 
            "Freigabeeinstellungen speichern", 
            [System.Windows.MessageBoxButton]::OK, 
            [System.Windows.MessageBoxImage]::Error
        )
    }
}

# Webseiten-Sammlungen aktualisieren
function Refresh-SiteCollections {
    try {
        if (-not $script:isConnected) {
            Write-Log "Keine Verbindung zu SharePoint Online. Websitesammlungen können nicht aktualisiert werden." -Level "WARNING"
            return
        }
        
        Write-Log "Aktualisiere Websitesammlungen..." -Level "INFO"
        Update-GuiText -ControlName "txtStatus" -Text "Lade Websitesammlungen..."
        
        $sites = Get-SPOSiteCollections
        
        if ($null -ne $sites) {
            $script:GUIElements["dgSites"].Dispatcher.Invoke(
                [Action]{
                    $script:GUIElements["dgSites"].ItemsSource = $sites
                },
                "Normal"
            )
            
            # Site-URLs in Dropdowns aktualisieren
            $siteUrls = $sites | Select-Object -ExpandProperty Url
            
            $script:GUIElements["cmbSiteUrl"].Dispatcher.Invoke(
                [Action]{
                    $script:GUIElements["cmbSiteUrl"].ItemsSource = $siteUrls
                },
                "Normal"
            )
            
            $script:GUIElements["cmbUserSiteUrl"].Dispatcher.Invoke(
                [Action]{
                    $script:GUIElements["cmbUserSiteUrl"].ItemsSource = $siteUrls
                },
                "Normal"
            )
            
            Write-Log "$($sites.Count) Websitesammlungen gefunden." -Level "SUCCESS"
        }
        else {
            Write-Log "Keine Websitesammlungen gefunden oder Fehler beim Abrufen." -Level "WARNING"
        }
        
        Update-GuiText -ControlName "txtStatus" -Text "Bereit"
    }
    catch {
        $errorMessage = ${_}
        Write-Log "Fehler beim Aktualisieren der Websitesammlungen: $errorMessage" -Level "ERROR"
        Update-GuiText -ControlName "txtStatus" -Text "Fehler"
    }
}

# Neue Websitesammlung Dialog anzeigen
function Show-NewSiteDialog {
    try {
        # Hier müssten wir eigentlich ein neues WPF-Fenster erstellen und anzeigen
        # Da dies eine komplexere Funktion ist, erstellen wir vorerst nur eine Meldung
        [System.Windows.MessageBox]::Show(
            "Diese Funktion ist noch nicht implementiert.",
            "Neue Websitesammlung",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Information
        )
    }
    catch {
        $errorMessage = ${_}
        Write-Log "Fehler beim Anzeigen des Dialogs für eine neue Websitesammlung: $errorMessage" -Level "ERROR"
    }
}

# Listen einer ausgewählten Site laden
function Load-ListsFromSite {
    try {
        if (-not $script:isConnected) {
            Write-Log "Keine Verbindung zu SharePoint Online. Listen können nicht geladen werden." -Level "WARNING"
            return
        }
        
        $selectedUrl = $script:GUIElements["cmbSiteUrl"].SelectedItem
        
        if ([string]::IsNullOrEmpty($selectedUrl)) {
            Write-Log "Keine Site-URL ausgewählt. Listen können nicht geladen werden." -Level "WARNING"
            [System.Windows.MessageBox]::Show(
                "Bitte wählen Sie eine Website-URL aus.",
                "Listen laden",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Warning
            )
            return
        }
        
        Write-Log "Lade Listen von $selectedUrl..." -Level "INFO"
        Update-GuiText -ControlName "txtStatus" -Text "Lade Listen..."
        
        $lists = Get-SPOLists -SiteUrl $selectedUrl
        
        if ($null -ne $lists) {
            $script:GUIElements["dgLists"].Dispatcher.Invoke(
                [Action]{
                    $script:GUIElements["dgLists"].ItemsSource = $lists
                },
                "Normal"
            )
            
            Write-Log "$($lists.Count) Listen geladen." -Level "SUCCESS"
        }
        else {
            Write-Log "Keine Listen gefunden oder Fehler beim Abrufen." -Level "WARNING"
        }
        
        Update-GuiText -ControlName "txtStatus" -Text "Bereit"
    }
    catch {
        $errorMessage = ${_}
        Write-Log "Fehler beim Laden der Listen: $errorMessage" -Level "ERROR"
        Update-GuiText -ControlName "txtStatus" -Text "Fehler"
    }
}

# Benutzer einer ausgewählten Site laden
function Load-UsersFromSite {
    try {
        if (-not $script:isConnected) {
                       Write-Log "Keine Verbindung zu SharePoint Online. Benutzer können nicht geladen werden." -Level "WARNING"
            return
        }
        
        $selectedUrl = $script:GUIElements["cmbUserSiteUrl"].SelectedItem
        
        if ([string]::IsNullOrEmpty($selectedUrl)) {
            Write-Log "Keine Site-URL ausgewählt. Benutzer können nicht geladen werden." -Level "WARNING"
            [System.Windows.MessageBox]::Show(
                "Bitte wählen Sie eine Website-URL aus.",
                "Benutzer laden",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Warning
            )
            return
        }
        
        Write-Log "Lade Benutzer von $selectedUrl..." -Level "INFO"
        Update-GuiText -ControlName "txtStatus" -Text "Lade Benutzer..."
        
        $users = Get-SPOUsers -SiteUrl $selectedUrl
        
        if ($null -ne $users) {
            $script:GUIElements["dgUsers"].Dispatcher.Invoke(
                [Action]{
                    $script:GUIElements["dgUsers"].ItemsSource = $users
                },
                "Normal"
            )
            
            Write-Log "$($users.Count) Benutzer geladen." -Level "SUCCESS"
        }
        else {
            Write-Log "Keine Benutzer gefunden oder Fehler beim Abrufen." -Level "WARNING"
        }
        
        Update-GuiText -ControlName "txtStatus" -Text "Bereit"
    }
    catch {
        $errorMessage = ${_}
        Write-Log "Fehler beim Laden der Benutzer: $errorMessage" -Level "ERROR"
        Update-GuiText -ControlName "txtStatus" -Text "Fehler"
    }
}
#endregion

#region Hauptteil des Skripts
try {
    Write-Log "Starte SharePoint Online Management Tool..." -Level "INFO"
    
    # GUI initialisieren
    $window = Initialize-GUI
    
    # Verbesserte Fehlerprüfung: Sicherstellen, dass $window tatsächlich ein WPF-Fenster ist
    if ($null -eq $window -or $window -isnot [System.Windows.Window]) {
        $errorMsg = "GUI konnte nicht initialisiert werden. Erhaltenes Objekt: $($window.GetType().FullName)"
        Write-Log $errorMsg -Level "ERROR"
        
        # Einfache Fallback-Fehlermeldung anzeigen
        [System.Windows.MessageBox]::Show(
            "Die Benutzeroberfläche konnte nicht initialisiert werden. Bitte prüfen Sie die Protokolldatei für Details.",
            "SharePoint Online Management Tool - Fehler",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
        
        exit 1
    }
    
    # Globale Variablen 
    $script:isConnected = $false
    
    # GUI anzeigen
    Write-Log "GUI wird angezeigt." -Level "INFO"
    [void]$window.ShowDialog()
    Write-Log "GUI wurde geschlossen." -Level "INFO"
    
    # Aufräumen, wenn das Fenster geschlossen wird
    if ($script:isConnected) {
        Write-Log "Trenne bestehende Verbindungen..." -Level "INFO"
        try {
            Disconnect-PnPOnline -ErrorAction SilentlyContinue
            Disconnect-SPOService -ErrorAction SilentlyContinue
            Write-Log "Verbindungen getrennt." -Level "INFO"
        } 
        catch {
            Write-Log "Fehler beim Trennen der Verbindungen: $_" -Level "WARNING"
            # Kein kritischer Fehler beim Beenden
        }
    }
}
catch {
    $errorMessage = ${_}
    Write-Log "Kritischer Fehler im Hauptteil des Skripts: $errorMessage" -Level "ERROR"
    
    # Fallback-Fehlermeldung anzeigen, wenn möglich
    try {
        [System.Windows.MessageBox]::Show(
            "Ein kritischer Fehler ist aufgetreten:`n`n$errorMessage", 
            "SharePoint Online Management Tool - Fehler",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
    }
    catch {
        # Selbst der Fallback ist fehlgeschlagen
        Write-Host "Kritischer Fehler: $errorMessage" -ForegroundColor Red
    }
    
    exit 1
}
finally {
    # Zusätzliche Aufräumarbeiten
    Write-Log "Programm beendet." -Level "INFO"
}
#endregion