<#
.SYNOPSIS
    Starter für Exchange Online Management Tool mit Microsoft Graph Integration
.DESCRIPTION
    Dieses Skript dient als Einrichtungs- und Startkomponente für die easyIT- Tool Sammlung.
    Es überprüft erforderliche Module, ermöglicht deren Installation und stellt eine, falls nötig, 
    Verbindung zu Exchange Online und Microsoft Graph her, bevor das Hauptskript geladen wird.
#>

# Konfiguration
$config = @{
    OptionalModules = @{
        Azure = $true      # Azure PowerShell Module installieren
        SharePoint = $true # SharePoint Online Management Shell installieren
        Teams = $true      # Microsoft Teams Modul installieren
    }
}

# Benötigte Assemblies laden
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

# Logging initialisieren
$script:logFolder = Join-Path -Path $PSScriptRoot -ChildPath "Logs"
$script:logFile = Join-Path -Path $script:logFolder -ChildPath "easyEXO_Setup.log"

if (-not (Test-Path -Path $script:logFolder)) {
    New-Item -Path $script:logFolder -ItemType Directory -Force | Out-Null
}

function Write-Log {
    param (
        [string]$Message,
        [ValidateSet("Info", "Warning", "Error", "Success")]
        [string]$Type = "Info"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Type] $Message"
    Add-Content -Path $script:logFile -Value $logEntry -Encoding UTF8
    
    # Auch an die Konsole ausgeben
    switch ($Type) {
        "Info"    { Write-Host $logEntry -ForegroundColor Gray }
        "Warning" { Write-Host $logEntry -ForegroundColor Yellow }
        "Error"   { Write-Host $logEntry -ForegroundColor Red }
        "Success" { Write-Host $logEntry -ForegroundColor Green }
    }
}

# Benötigte Module definieren mit kürzeren Beschreibungen und korrigierten Connect-Befehlen
$requiredModules = @(
    @{
        Name = "ExchangeOnlineManagement"
        MinVersion = "3.0.0"
        Description = "Exchange Online V3"
        Optional = $false
        ConnectCmd = "Connect-ExchangeOnline -ShowBanner:$false -ShowProgress $true -ErrorAction Stop"
        ConnectCheck = { Get-PSSession | Where-Object { $_.ConfigurationName -eq "Microsoft.Exchange" -and $_.State -eq "Opened" } }
    },
    @{
        Name = "Microsoft.Graph"
        MinVersion = "1.27.0" 
        Description = "Microsoft Graph SDK"
        Optional = $false
        ConnectCmd = "Connect-MgGraph -Scopes `"User.Read.All`",`"Group.ReadWrite.All`"" 
        ConnectCheck = { Get-MgContext -ErrorAction SilentlyContinue }
    },
    @{
        Name = "PowerShellGet"
        MinVersion = "2.2.5"
        Description = "PS Module Management"
        Optional = $false
        ConnectCmd = $null # Keine Verbindung nötig
        ConnectCheck = $null
    },
    @{
        Name = "Microsoft.Graph.Beta"
        MinVersion = "1.27.0"
        Description = "Microsoft Graph Beta"
        Optional = $true
        ConnectCmd = "Connect-MgGraph -Scopes `"User.Read.All`" -Beta"
        ConnectCheck = { Get-MgContext -ErrorAction SilentlyContinue }
    },
    @{
        Name = "Microsoft.Entra"
        MinVersion = "1.0.0"
        Description = "Entra ID PowerShell"
        Optional = $true
        ConnectCmd = "Connect-Entra -Scopes `"User.Read.All`"" 
        ConnectCheck = { Get-MgContext -ErrorAction SilentlyContinue }
    },
    @{
        Name = "Microsoft.Entra.Beta"
        MinVersion = "1.0.0"
        Description = "Entra ID PowerShell (Beta)"
        Optional = $true
        ConnectCmd = "Connect-EntraBeta -Scopes `"User.Read.All`"" 
        ConnectCheck = { Get-MgContext -ErrorAction SilentlyContinue }
    },
    @{
        Name = "PnP.PowerShell"
        MinVersion = "1.12.0"
        Description = "SharePoint PnP"
        Optional = $true
        ConnectCmd = "Connect-PnPOnline -Interactive"
        ConnectCheck = { Get-PnPConnection -ErrorAction SilentlyContinue }
    },
    @{
        Name = "AzureAD"
        MinVersion = "2.0.2.182"
        Description = "Azure AD (veraltet)"
        Optional = $true
        ConnectCmd = "Connect-AzureAD"
        ConnectCheck = { Get-AzureADCurrentSessionInfo -ErrorAction SilentlyContinue }
    },
    @{
        Name = "MSOnline"
        MinVersion = "1.1.183.57"
        Description = "Azure AD V1 (veraltet)"
        Optional = $true
        ConnectCmd = "Connect-MsolService"
        ConnectCheck = { Get-MsolCompanyInformation -ErrorAction SilentlyContinue }
    },
    @{
        Name = "Microsoft.Online.SharePoint.PowerShell"
        MinVersion = "16.0.0"
        Description = "SharePoint Online"
        Optional = $true
        ConnectCmd = "# SharePoint-Verbindung erfordert Admin-URL"  # Spezieller Platzhalter für SharePoint
        ConnectCheck = { Get-SPOTenant -ErrorAction SilentlyContinue }
    }
)

# Optionale Module je nach Konfiguration hinzufügen
if ($config.OptionalModules.Azure) {
    $requiredModules += @{
        Name = "Az"
        MinVersion = "10.0.0"
        Description = "Azure PowerShell"
        Optional = $true
        ConnectCmd = "Connect-AzAccount"
        ConnectCheck = { Get-AzContext -ErrorAction SilentlyContinue }
    }
}

if ($config.OptionalModules.SharePoint) {
    $requiredModules += @{
        Name = "Microsoft.Online.SharePoint.PowerShell"
        MinVersion = "16.0.0"
        Description = "SharePoint Online"
        Optional = $true
        ConnectCmd = "Connect-SPOService -Url https://TENANT-admin.sharepoint.com"
        ConnectCheck = { Get-SPOTenant -ErrorAction SilentlyContinue }
    }
}

if ($config.OptionalModules.Teams) {
    $requiredModules += @{
        Name = "MicrosoftTeams"
        MinVersion = "5.3.0"
        Description = "Microsoft Teams"
        Optional = $true
        ConnectCmd = "Connect-MicrosoftTeams"
        ConnectCheck = { Get-Team -ErrorAction SilentlyContinue }
    }
}

# Prüfen, ob ein Modul installiert ist
function Test-ModuleInstalled {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ModuleName,
        
        [Parameter(Mandatory = $false)]
        [version]$MinVersion
    )
    
    $module = Get-Module -ListAvailable -Name $ModuleName
    
    if ($null -eq $module) {
        return $false
    }
    
    if ($MinVersion) {
        $latestVersion = $module | Sort-Object Version -Descending | Select-Object -First 1
        if ($latestVersion.Version -lt $MinVersion) {
            return $false
        }
    }
    
    return $true
}

# Modul installieren
function Install-RequiredModule {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ModuleName,
        
        [Parameter(Mandatory = $false)]
        [version]$MinVersion
    )
    
    try {
        # Prüfe, ob NuGet Provider installiert ist
        if (-not (Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue)) {
            Write-Log "NuGet Provider wird installiert..." -Type "Info"
            Install-PackageProvider -Name NuGet -Force -Scope CurrentUser | Out-Null
        }
        
        # Setze PSGallery als vertrauenswürdig
        if ((Get-PSRepository -Name PSGallery).InstallationPolicy -ne "Trusted") {
            Write-Log "PSGallery wird als vertrauenswürdige Quelle gesetzt..." -Type "Info"
            Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
        }
        
        $installParams = @{
            Name = $ModuleName
            Force = $true
            AllowClobber = $true
            Scope = "CurrentUser"
            Repository = "PSGallery"
        }
        
        if ($MinVersion) {
            $installParams.MinimumVersion = $MinVersion
            Write-Log "Installiere $ModuleName (Mindestversion $MinVersion)..." -Type "Info"
        } else {
            Write-Log "Installiere $ModuleName..." -Type "Info"
        }
        
        Install-Module @installParams
        
        Write-Log "Modul $ModuleName erfolgreich installiert." -Type "Success"
        return $true
    }
    catch {
        Write-Log "Fehler bei der Installation von $ModuleName - $($_.Exception.Message)" -Type "Error"
        return $false
    }
}

# Verbindung zu Exchange Online herstellen
function Connect-ToExchangeOnline {
    try {
        # Prüfen, ob das Exchange Online Modul vorhanden ist
        if (-not (Test-ModuleInstalled -ModuleName "ExchangeOnlineManagement")) {
            Write-Log "Exchange Online Management Modul ist nicht installiert." -Type "Error"
            return $false
        }
        
        # Modul importieren
        Import-Module ExchangeOnlineManagement -ErrorAction Stop
        Write-Log "Modul ExchangeOnlineManagement wurde importiert" -Type "Info"
        
        # Prüfen, ob bereits eine Verbindung besteht
        if (Test-ExchangeOnlineConnection) {
            Write-Log "Bereits mit Exchange Online verbunden" -Type "Success"
            return $true
        }
        
        Write-Log "Verbinde mit Exchange Online über moderne Authentifizierung..." -Type "Info"
        
        # Verbindung herstellen mit besserer Fehlerbehandlung
        Connect-ExchangeOnline -ShowBanner:$false -ShowProgress $true -ErrorAction Stop
        
        # Kurz warten, damit sich die Session etablieren kann
        Start-Sleep -Seconds 2
        
        # Verbindung prüfen mit mehreren Versuchen
        $isConnected = $false
        $maxRetries = 3
        $retryCount = 0
        
        while (-not $isConnected -and $retryCount -lt $maxRetries) {
            $retryCount++
            
            if ($retryCount -gt 1) {
                Write-Log "Verbindungscheck Versuch $retryCount..." -Type "Info"
                Start-Sleep -Seconds 2
            }
            
            $isConnected = Test-ExchangeOnlineConnection
        }
        
        if ($isConnected) {
            Write-Log "Exchange Online Verbindung erfolgreich hergestellt" -Type "Success"
            return $true
        } else {
            # Letzte Chance: Prüfe direkt, ob irgendeine EXO-Session existiert
            $anyExoSessions = Get-PSSession | Where-Object { $_.ConfigurationName -eq "Microsoft.Exchange" }
            if ($anyExoSessions.Count -gt 0) {
                Write-Log "Exchange Online Sessions gefunden, Status möglicherweise nicht optimal" -Type "Warning"
                return $true  # Wir akzeptieren auch nicht vollständig offene Sessions
            }
            
            Write-Log "Konnte keine aktive Exchange Online Session verifizieren." -Type "Error"
            return $false
        }
    }
    catch {
        Write-Log "Fehler bei der Verbindung zu Exchange Online: $($_.Exception.Message)" -Type "Error"
        
        # Mehr Kontext zur Fehlerbehebung
        try {
            [System.Windows.MessageBox]::Show(
                "Fehler bei der Verbindung zu Exchange Online: $($_.Exception.Message)", 
                "Verbindungsfehler", 
                [System.Windows.MessageBoxButton]::OK, 
                [System.Windows.MessageBoxImage]::Error)
        }
        catch {}
        
        return $false
    }
}

# Verbindung zu Microsoft Graph herstellen
function Connect-ToMicrosoftGraph {
    param (
        [Parameter(Mandatory = $false)]
        [ValidateSet("Interactive", "DeviceCode", "AccessToken", "ClientApp", "CertificateBased", "ClientSecret", "ManagedIdentity")]
        [string]$AuthMethod = "Interactive",
        
        [Parameter(Mandatory = $false)]
        [string]$ClientId,
        
        [Parameter(Mandatory = $false)]
        [string]$TenantId,
        
        [Parameter(Mandatory = $false)]
        [string]$CertificateThumbprint,
        
        [Parameter(Mandatory = $false)]
        [System.Management.Automation.PSCredential]$ClientSecretCredential,
        
        [Parameter(Mandatory = $false)]
        [string]$AccessToken
    )
    
    try {
        # Prüfen, ob das Microsoft Graph Modul vorhanden ist
        if (-not (Test-ModuleInstalled -ModuleName "Microsoft.Graph")) {
            Write-Log "Microsoft Graph Modul ist nicht installiert." -Type "Error"
            return $false
        }
        
        # Modul importieren
        Import-Module Microsoft.Graph -ErrorAction Stop
        
        # Definiere die erforderlichen Bereiche
        $scopes = @(
            "User.Read.All", 
            "Organization.Read.All",
            "Directory.Read.All",
            "Group.Read.All"
        )
        
        # Prüfen, ob bereits eine Verbindung besteht
        $existingContext = Get-MgContext -ErrorAction SilentlyContinue
        if ($null -ne $existingContext) {
            Write-Log "Bereits mit Microsoft Graph verbunden als: $($existingContext.Account)" -Type "Success"
            return $true
        }
        
        # Je nach Authentifizierungsmethode verbinden
        switch ($AuthMethod) {
            "Interactive" {
                Write-Log "Verbinde mit Microsoft Graph über interaktive Anmeldung..." -Type "Info"
                Connect-MgGraph -Scopes $scopes -ErrorAction Stop
            }
            "DeviceCode" {
                Write-Log "Verbinde mit Microsoft Graph über Device Code Authentifizierung..." -Type "Info"
                Write-Host "`n*** BITTE FOLGEN SIE DEN ANWEISUNGEN FÜR DIE GERÄTE-CODE AUTHENTIFIZIERUNG ***" -ForegroundColor Cyan -BackgroundColor DarkBlue
                Connect-MgGraph -Scopes $scopes -UseDeviceAuthentication -ErrorAction Stop
            }
            "AccessToken" {
                if ([string]::IsNullOrEmpty($AccessToken)) {
                    Write-Log "Für die AccessToken-Methode wird ein Token benötigt." -Type "Error"
                    return $false
                }
                Write-Log "Verbinde mit Microsoft Graph über Access Token..." -Type "Info"
                Connect-MgGraph -AccessToken $AccessToken -ErrorAction Stop
            }
            "ClientApp" {
                if ([string]::IsNullOrEmpty($ClientId) -or [string]::IsNullOrEmpty($TenantId)) {
                    Write-Log "Für die ClientApp-Methode werden ClientId und TenantId benötigt." -Type "Error"
                    return $false
                }
                Write-Log "Verbinde mit Microsoft Graph über Client App (Public Client)..." -Type "Info"
                Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -ErrorAction Stop
            }
            "CertificateBased" {
                if ([string]::IsNullOrEmpty($ClientId) -or [string]::IsNullOrEmpty($TenantId) -or [string]::IsNullOrEmpty($CertificateThumbprint)) {
                    Write-Log "Für die CertificateBased-Methode werden ClientId, TenantId und CertificateThumbprint benötigt." -Type "Error"
                    return $false
                }
                Write-Log "Verbinde mit Microsoft Graph über Zertifikatsbasierte Authentifizierung..." -Type "Info"
                Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -CertificateThumbprint $CertificateThumbprint -ErrorAction Stop
            }
            "ClientSecret" {
                if ([string]::IsNullOrEmpty($TenantId) -or $null -eq $ClientSecretCredential) {
                    Write-Log "Für die ClientSecret-Methode werden TenantId und ClientSecretCredential benötigt." -Type "Error"
                    return $false
                }
                Write-Log "Verbinde mit Microsoft Graph über Client Secret..." -Type "Info"
                Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $ClientSecretCredential -ErrorAction Stop
            }
            "ManagedIdentity" {
                Write-Log "Verbinde mit Microsoft Graph über Azure Managed Identity..." -Type "Info"
                Connect-MgGraph -Identity -ErrorAction Stop
            }
        }
        
        # Verbindung prüfen
        $context = Get-MgContext -ErrorAction Stop
        if ($null -ne $context) {
            Write-Log "Microsoft Graph Verbindung erfolgreich hergestellt als: $($context.Account)" -Type "Success"
            return $true
        }
        
        Write-Log "Unerwarteter Fehler: Verbindung hergestellt, aber kein Kontext gefunden." -Type "Error"
        return $false
    }
    catch {
        Write-Log "Fehler bei der Verbindung zu Microsoft Graph: $($_.Exception.Message)" -Type "Error"
        return $false
    }
}

# Hauptskript laden (überarbeitet, um das Fenster geöffnet zu lassen)
function Load-MainScript {
    param (
        [string]$ScriptPath,
        [switch]$KeepSetupWindowOpen = $true
    )
    
    try {
        if (-not (Test-Path -Path $ScriptPath)) {
            Write-Log "Hauptskript nicht gefunden: $ScriptPath" -Type "Error"
            return $false
        }
        
        Write-Log "Lade Hauptskript: $ScriptPath" -Type "Info"
        
        # Hauptfenster nicht schließen, nur minimieren oder ausblenden, wenn gewünscht
        if ($KeepSetupWindowOpen -and $Global:window) {
            Update-LogDisplay -Message "Hauptfenster bleibt geöffnet, um Sessions beizubehalten" -Type "Info"
            # Fenster ausblenden/minimieren statt schließen
            $Global:window.WindowState = [System.Windows.WindowState]::Minimized
            $Global:window.Title = "easyEXO Setup (Hauptskript läuft...)"
        }
        
        # Prüfen, ob wichtige Verbindungen bestehen
        $exoSession = Get-PSSession | Where-Object { 
            $_.ConfigurationName -eq "Microsoft.Exchange" -and 
            ($_.State -eq "Opened" -or $_.Availability -eq "Available")
        }
        
        if ($null -eq $exoSession) {
            Update-LogDisplay -Message "Warnung: Keine aktive Exchange Online Session gefunden!" -Type "Warning"
            if ([System.Windows.MessageBox]::Show(
                "Es wurde keine aktive Exchange Online Session gefunden. Das Hauptskript funktioniert möglicherweise nicht richtig. Möchten Sie trotzdem fortfahren?", 
                "Keine Exchange-Verbindung", 
                [System.Windows.MessageBoxButton]::YesNo, 
                [System.Windows.MessageBoxImage]::Warning) -eq "No") {
                
                Update-LogDisplay -Message "Hauptskript-Start vom Benutzer abgebrochen." -Type "Info"
                if ($KeepSetupWindowOpen -and $Global:window) {
                    $Global:window.WindowState = [System.Windows.WindowState]::Normal
                    $Global:window.Title = "easyEXO Setup"
                }
                return $false
            }
        }
        
        # Skript in der aktuellen Session ausführen
        Update-LogDisplay -Message "Führe Hauptskript in der aktuellen Session aus..." -Type "Info"
        . $ScriptPath
        
        Write-Log "Hauptskript erfolgreich ausgeführt" -Type "Success"
        
        # Fenster wieder anzeigen
        if ($KeepSetupWindowOpen -and $Global:window) {
            $Global:window.WindowState = [System.Windows.WindowState]::Normal
            $Global:window.Title = "easyEXO Setup"
        }
        
        return $true
    }
    catch {
        Write-Log "Fehler beim Laden des Hauptskripts: $($_.Exception.Message)" -Type "Error"
        
        # Detaillierte Fehlerinformationen anzeigen
        $detailedError = "Fehlerdetails:`n"
        $detailedError += "Fehlertyp: $($_.Exception.GetType().FullName)`n"
        $detailedError += "Fehlermeldung: $($_.Exception.Message)`n"
        $detailedError += "Fehlerort: $($_.InvocationInfo.PositionMessage)`n"
        
        if ($_.Exception.InnerException) {
            $detailedError += "Innere Ausnahme: $($_.Exception.InnerException.Message)`n"
        }
        
        Write-Log $detailedError -Type "Error"
        
        # Fenster wieder anzeigen, falls minimiert
        if ($KeepSetupWindowOpen -and $Global:window) {
            $Global:window.WindowState = [System.Windows.WindowState]::Normal
            $Global:window.Title = "easyEXO Setup (Fehler beim Ausführen des Hauptskripts)"
        }
        
        [System.Windows.MessageBox]::Show(
            "Das Hauptskript konnte nicht ausgeführt werden. Fehlermeldung: $($_.Exception.Message)", 
            "Fehler beim Ausführen des Skripts", 
            [System.Windows.MessageBoxButton]::OK, 
            [System.Windows.MessageBoxImage]::Error
        )
        
        return $false
    }
}

# XAML für die GUI definieren - Event-Bindungen entfernt
$xaml = @"
<Window 
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="easyEXO Setup" 
    Height="900" 
    Width="1200"
    WindowStartupLocation="CenterScreen">
    <Window.Resources>
        <Style TargetType="Button" x:Key="InstallButtonStyle">
            <Setter Property="Padding" Value="5,2"/>
            <Setter Property="Background" Value="#EFEFEF"/>
        </Style>
        <Style TargetType="Button" x:Key="ConnectButtonStyle">
            <Setter Property="Padding" Value="5,2"/>
            <Setter Property="Background" Value="#10E5E5"/>
            <Setter Property="Foreground" Value="Black"/>
            <Setter Property="FontWeight" Value="Bold"/>
        </Style>
        <Style TargetType="Button" x:Key="ScriptButtonStyle">
            <Setter Property="Padding" Value="10,5"/>
            <Setter Property="Margin" Value="0,2"/>
            <Setter Property="Background" Value="#E3F2FD"/>
            <Setter Property="HorizontalContentAlignment" Value="Left"/>
        </Style>
    </Window.Resources>
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Header -->
        <Border Grid.Row="0" Background="#0078D4" Padding="20,15">
            <Grid>
                <StackPanel>
                    <TextBlock Text="easyIT - START EXO Management" FontSize="24" FontWeight="Bold" Foreground="White"/>
                    <TextBlock Text="Connect first with Exchange Online | (optional) Microsoft Graph Tool" FontSize="14" Foreground="White" Margin="0,5,0,0"/>
                </StackPanel>
            </Grid>
        </Border>

        <!-- Content with Split Layout -->
        <Grid Grid.Row="1" Margin="20">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="2*"/> <!-- Status-Bereich: 2/3 -->
                <ColumnDefinition Width="1*"/> <!-- Skript-Auswahl: 1/3 -->
            </Grid.ColumnDefinitions>

            <!-- Linke Seite - Status-Bereich -->
            <Grid Grid.Column="0" Margin="0,0,10,0">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>

                <!-- Module Check Section -->
                <GroupBox Grid.Row="0" Header="Module und Verbindungen" Margin="0,0,0,10" Padding="10">
                    <ListView x:Name="lvModules" Margin="0,10,0,0" BorderThickness="1">
                        <ListView.View>
                            <GridView>
                                <GridViewColumn Header="Modul" DisplayMemberBinding="{Binding Name}" Width="180"/>
                                <GridViewColumn Header="Beschreibung" DisplayMemberBinding="{Binding Description}" Width="180"/>
                                <GridViewColumn Header="Version" DisplayMemberBinding="{Binding InstalledVersion}" Width="80"/>
                                <GridViewColumn Header="Status" DisplayMemberBinding="{Binding Status}" Width="80"/>
                                <GridViewColumn Header="Installation" Width="100">
                                    <GridViewColumn.CellTemplate>
                                        <DataTemplate>
                                            <Button Content="Installieren" 
                                                    Style="{StaticResource InstallButtonStyle}"
                                                    IsEnabled="{Binding InstallEnabled}"
                                                    Tag="{Binding Name}"/>
                                        </DataTemplate>
                                    </GridViewColumn.CellTemplate>
                                </GridViewColumn>
                                <GridViewColumn Header="Verbindung" Width="100">
                                    <GridViewColumn.CellTemplate>
                                        <DataTemplate>
                                            <Button Content="Verbinden" 
                                                    Style="{StaticResource ConnectButtonStyle}"
                                                    IsEnabled="{Binding ConnectEnabled}"
                                                    Tag="{Binding Name}"/>
                                        </DataTemplate>
                                    </GridViewColumn.CellTemplate>
                                </GridViewColumn>
                            </GridView>
                        </ListView.View>
                    </ListView>
                </GroupBox>

                <!-- Status Log -->
                <GroupBox Grid.Row="1" Header="Status und Log" Margin="0,10,0,0" Padding="10">
                    <ScrollViewer VerticalScrollBarVisibility="Auto">
                        <TextBox x:Name="txtLog" IsReadOnly="True" TextWrapping="Wrap" 
                                 FontFamily="Consolas" Background="#f5f5f5" BorderThickness="0"/>
                    </ScrollViewer>
                </GroupBox>
            </Grid>

            <!-- Rechte Seite - Skript-Auswahl -->
            <GroupBox Grid.Column="1" Header="Verfügbare PowerShell-Skripte" Margin="10,0,0,0" Padding="10">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <!-- Suchfeld -->
                    <TextBox Grid.Row="0" x:Name="txtSearchScripts" Margin="0,0,0,10" 
                             Height="25" VerticalContentAlignment="Center"
                             BorderBrush="Gray" BorderThickness="1" Padding="5,0"
                             Text="Skriptsuche..." />
                    
                    <!-- Skriptliste -->
                    <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto">
                        <StackPanel x:Name="spScriptList" Orientation="Vertical" Margin="0,5,0,5">
                            <!-- Skriptbuttons werden hier dynamisch hinzugefügt -->
                        </StackPanel>
                    </ScrollViewer>
                    
                    <!-- Aktualisieren-Button -->
                    <Button Grid.Row="2" x:Name="btnRefreshScripts" 
                            Content="Skriptliste aktualisieren" Margin="0,5,0,0" 
                            Height="30" Background="#F0F0F0"/>
                </Grid>
            </GroupBox>
        </Grid>

        <!-- Footer with Buttons -->
        <Border Grid.Row="2" Background="#f0f0f0" Padding="20,15">
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                <Button x:Name="btnRefresh" Content="Module aktualisieren" Width="150" Height="30" Margin="0,0,10,0" Background="#FFE4FFE1"/>
                <Button x:Name="btnExit" Content="Beenden" Width="150" Height="30" Background="#FFFF7272" Margin="10,0,0,0"/>
            </StackPanel>
        </Border>
    </Grid>
</Window>
"@

# XAML für den Graph Auth Dialog
$graphAuthXaml = @"
<Window 
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Microsoft Graph Verbindungsoptionen" 
    Height="500" 
    Width="700"
    WindowStartupLocation="CenterScreen" Background="#FFF5FCFD">
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Header -->
        <TextBlock Grid.Row="0" Text="LIZENZ ANALYSE  |  Microsoft Graph Authentifizierungsmethode" 
                   FontSize="16" FontWeight="Bold" Margin="0,0,0,15"/>

        <!-- Auth Method Selection -->
        <ComboBox Grid.Row="1" x:Name="cmbAuthMethod" Margin="0,0,0,15" 
                  SelectedIndex="0">
            <ComboBoxItem Content="Interaktive Anmeldung (Browser)"/>
            <ComboBoxItem Content="Gerätecode-Authentifizierung"/>
            <ComboBoxItem Content="Anmeldung mit vorhandenem Access Token"/>
            <ComboBoxItem Content="Anmeldung über eigene Azure AD App (Public Client)"/>
            <ComboBoxItem Content="App-Only: Zertifikatsbasierte Authentifizierung"/>
            <ComboBoxItem Content="App-Only: Client-Secret basierte Authentifizierung"/>
            <ComboBoxItem Content="Azure Managed Identity"/>
        </ComboBox>

        <!-- Parameter Fields -->
        <Grid Grid.Row="2" x:Name="grdParameters">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="150"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

            <!-- AccessToken -->
            <StackPanel Grid.Row="0" Grid.ColumnSpan="2" x:Name="pnlAccessToken" Visibility="Collapsed">
                <TextBlock Text="Access Token:" Margin="0,5"/>
                <TextBox x:Name="txtAccessToken" Height="80" TextWrapping="Wrap" AcceptsReturn="True"/>
                <TextBlock Text="Hinweis: Token muss alle benötigten Berechtigungen enthalten" 
                           Margin="0,5" FontStyle="Italic" Foreground="Gray"/>
            </StackPanel>

            <!-- ClientId und TenantId für ClientApp -->
            <StackPanel Grid.Row="1" Grid.ColumnSpan="2" x:Name="pnlClientApp" Visibility="Collapsed">
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="150"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>

                    <TextBlock Grid.Row="0" Grid.Column="0" Text="Client ID:" Margin="0,5"/>
                    <TextBox Grid.Row="0" Grid.Column="1" x:Name="txtClientId" Margin="0,5"/>

                    <TextBlock Grid.Row="1" Grid.Column="0" Text="Tenant ID:" Margin="0,5"/>
                    <TextBox Grid.Row="1" Grid.Column="1" x:Name="txtTenantId" Margin="0,5"/>
                </Grid>
            </StackPanel>

            <!-- Certificate -->
            <StackPanel Grid.Row="2" Grid.ColumnSpan="2" x:Name="pnlCertificate" Visibility="Collapsed">
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="150"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <TextBlock Grid.Column="0" Text="Zertifikat Thumbprint:" Margin="0,5"/>
                    <TextBox Grid.Column="1" x:Name="txtCertThumbprint" Margin="0,5"/>
                </Grid>
            </StackPanel>

            <!-- Client Secret -->
            <StackPanel Grid.Row="3" Grid.ColumnSpan="2" x:Name="pnlClientSecret" Visibility="Collapsed">
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="150"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <TextBlock Grid.Column="0" Text="Client Secret:" Margin="0,5"/>
                    <PasswordBox Grid.Column="1" x:Name="txtClientSecret" Margin="0,5"/>
                </Grid>
            </StackPanel>

            <!-- Info Text -->
            <TextBlock Grid.Row="4" Grid.ColumnSpan="2" x:Name="txtInfo" 
                       TextWrapping="Wrap" Margin="0,20,0,0"/>
        </Grid>

        <!-- Buttons -->
        <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,20,0,0">
            <Button x:Name="btnCancel" Content="Abbrechen" Width="100" Height="30" Margin="0,0,10,0" Background="#FFFF7272"/>
            <Button x:Name="btnConnect" Content="Verbinden" Width="100" Height="30" Background="#FF87FF77"/>
        </StackPanel>
    </Grid>
</Window>
"@

# XAML für den SharePoint-Admin-URL-Dialog
$spoConnectXaml = @"
<Window 
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="SharePoint Online Verbindung" 
    Height="220" 
    Width="600"
    WindowStartupLocation="CenterScreen" Background="#FFF5FCFD">
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>

        <!-- Header -->
        <TextBlock Grid.Row="0" Text="Verbindung zu SharePoint Online" 
                   FontSize="16" FontWeight="Bold" Margin="0,0,0,15"/>
        
        <!-- Beschreibung -->
        <TextBlock Grid.Row="1" Text="Bitte geben Sie die SharePoint Admin Center URL ein. Format: https://TENANT-admin.sharepoint.com" 
                   TextWrapping="Wrap" Margin="0,0,0,10"/>
        
        <!-- URL Eingabefeld -->
        <TextBox Grid.Row="2" x:Name="txtSpoUrl" Margin="0,5,0,15" Height="25"
                 Text="https://" FontSize="14"/>
        
        <!-- Buttons -->
        <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Right">
            <Button x:Name="btnSpoCancel" Content="Abbrechen" Width="100" Height="30" Margin="0,0,10,0" Background="#FFFF7272"/>
            <Button x:Name="btnSpoConnect" Content="Verbinden" Width="100" Height="30" Background="#FF87FF77"/>
        </StackPanel>
    </Grid>
</Window>
"@

# Event-Handler Scriptblocks definieren (vor der Verwendung)
$installModuleClickHandlerScript = {
    param($sender, $e)
    
    $button = [System.Windows.Controls.Button]$sender
    $moduleName = $button.Tag.ToString()
    
    # Finde das entsprechende Modul in $requiredModules
    $moduleInfo = $requiredModules | Where-Object { $_.Name -eq $moduleName } | Select-Object -First 1
    
    if ($moduleInfo) {
        Update-LogDisplay -Message "Installation von $moduleName wird gestartet..." -Type "Info"
        
        # Button deaktivieren während der Installation
        $button.IsEnabled = $false
        $button.Content = "Installiere..."
        
        # Modul installieren
        $success = Install-RequiredModule -ModuleName $moduleName -MinVersion $moduleInfo.MinVersion
        
        if ($success) {
            Update-LogDisplay -Message "Modul $moduleName erfolgreich installiert." -Type "Success"
        } else {
            Update-LogDisplay -Message "Fehler bei der Installation von $moduleName." -Type "Error"
        }
        
        # Modulliste aktualisieren
        Check-AllModules
    }
}

# Verbesserte Funktion für den Connect-Modul-Handler - mit SharePoint-Spezialfall
$connectModuleClickHandlerScript = {
    param($sender, $e)
    
    $button = [System.Windows.Controls.Button]$sender
    $moduleName = $button.Tag.ToString()
    
    # Finde das entsprechende Modul in $requiredModules
    $moduleInfo = $requiredModules | Where-Object { $_.Name -eq $moduleName } | Select-Object -First 1
    
    if ($moduleInfo) {
        Update-LogDisplay -Message "Verbindung zu $moduleName wird hergestellt..." -Type "Info"
        
        # Button deaktivieren während der Verbindung
        $button.IsEnabled = $false
        $button.Content = "Verbinde..."
        
        try {
            # Modul vorab importieren, um sicherzustellen, dass alle Cmdlets verfügbar sind
            if (-not (Get-Module -Name $moduleName -ErrorAction SilentlyContinue)) {
                Update-LogDisplay -Message "Importiere Modul $moduleName..." -Type "Info"
                Import-Module -Name $moduleName -ErrorAction Stop -Verbose:$false
            }
            
            # Spezialfall: SharePoint Online - URL-Dialog anzeigen
            if ($moduleName -eq "Microsoft.Online.SharePoint.PowerShell") {
                $spoUrl = Show-SPOConnectDialog
                
                if ([string]::IsNullOrEmpty($spoUrl)) {
                    Update-LogDisplay -Message "SharePoint-Verbindung abgebrochen." -Type "Warning"
                    $button.IsEnabled = $true
                    $button.Content = "Verbinden"
                    return
                }
                
                Update-LogDisplay -Message "Verbinde mit SharePoint Online: $spoUrl" -Type "Info"
                Connect-SPOService -Url $spoUrl -ErrorAction Stop
            }
            # Spezieller Fall: Exchange Online - verbesserte Verbindungsmethode
            elseif ($moduleName -eq "ExchangeOnlineManagement") {
                Update-LogDisplay -Message "Starte Exchange Online Verbindung mit verbesserter Methode..." -Type "Info"
                $connected = Connect-ToExchangeOnline
                if (-not $connected) {
                    throw "Exchange Online Verbindung konnte nicht hergestellt werden"
                }
            }
            # Spezieller Fall: Microsoft Graph - verbesserte Methode mit Dialog
            elseif ($moduleName -like "Microsoft.Graph*") {
                Update-LogDisplay -Message "Starte Microsoft Graph Verbindung..." -Type "Info"
                
                # Graph Auth Dialog anzeigen
                $authParams = Show-GraphAuthDialog
                if ($null -eq $authParams) {
                    Update-LogDisplay -Message "Microsoft Graph Verbindung abgebrochen." -Type "Warning"
                    $button.IsEnabled = $true
                    $button.Content = "Verbinden"
                    return
                }
                
                # Connect-Methode mit den vom Dialog zurückgegebenen Parametern aufrufen
                Connect-ToMicrosoftGraph @authParams
            }
            # Für alle anderen Module
            else {
                # Connect-Befehl ausführen
                $scriptBlock = [ScriptBlock]::Create($moduleInfo.ConnectCmd)
                & $scriptBlock
            }
            
            # Nach der Verbindung einige Sekunden warten
            Start-Sleep -Seconds 3
            
            # Verbindungsstatus mit verbesserter Fehlerbehandlung prüfen
            $isConnected = $false
            
            # Exchange Online Spezialfall
            if ($moduleName -eq "ExchangeOnlineManagement") {
                $isConnected = Test-ExchangeOnlineConnection
            }
            # Für andere Module
            elseif ($moduleInfo.ConnectCheck -ne $null) {
                try {
                    $checkResult = & $moduleInfo.ConnectCheck
                    $isConnected = ($null -ne $checkResult)
                } catch {
                    $isConnected = $false
                }
            }
            
            if ($isConnected) {
                Update-LogDisplay -Message "Verbindung zu $moduleName erfolgreich hergestellt." -Type "Success"
                $button.Content = "Verbunden"
                $button.IsEnabled = $false
            } else {
                Update-LogDisplay -Message "Verbindung zu $moduleName konnte nicht verifiziert werden." -Type "Warning"
                $button.Content = "Verbinden"
                $button.IsEnabled = $true
            }
        } catch {
            Update-LogDisplay -Message "Fehler bei der Verbindung zu $moduleName - $($_.Exception.Message)" -Type "Error"
            $button.Content = "Verbinden"
            $button.IsEnabled = $true
        }
    }
}

# XAML laden und GUI erstellen - überarbeitete Version
function Import-XamlGui {
    $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($xaml))
    
    try {
        $window = [System.Windows.Markup.XamlReader]::Load($reader)
        
        # Alle benannten Elemente verfügbar machen
        $namedNodes = $xaml.Split("`n") | 
            Where-Object { $_ -match 'x:Name="([^"]*)"' } | 
            ForEach-Object { $matches[1] }
        
        $namedNodes | ForEach-Object {
            $name = $_
            $element = $window.FindName($name)
            
            # Globale Variable für jedes benannte Element erstellen
            Set-Variable -Name $name -Value $element -Scope Global
        }
        
        # Event-Handler für die Buttons registrieren
        try {
            # Event-Handler für Module aktualisieren
            $btnRefresh = $window.FindName("btnRefresh")
            if ($btnRefresh) {
                $btnRefresh.Add_Click({
                    param($sender, $e)
                    Check-AllModules
                })
            }
            
            # Event-Handler für Beenden-Button
            $btnExit = $window.FindName("btnExit")
            if ($btnExit) {
                $btnExit.Add_Click({
                    $Global:window.Close()
                })
            }
            
            # Event-Handler für Skriptliste aktualisieren
            $btnRefreshScripts = $window.FindName("btnRefreshScripts")
            if ($btnRefreshScripts) {
                $btnRefreshScripts.Add_Click({
                    Load-ScriptList
                })
            }
            
            # Event-Handler für Skriptsuche
            $txtSearchScripts = $window.FindName("txtSearchScripts")
            if ($txtSearchScripts) {
                # Platzhaltertext entfernen, wenn Fokus erhalten
                $txtSearchScripts.Add_GotFocus({
                    if ($this.Text -eq "Skriptsuche...") {
                        $this.Text = ""
                    }
                })
                
                # Platzhaltertext hinzufügen, wenn Fokus verloren und leer
                $txtSearchScripts.Add_LostFocus({
                    if ([string]::IsNullOrWhiteSpace($this.Text)) {
                        $this.Text = "Skriptsuche..."
                    }
                })
                
                # Filterung der Skripte bei Texteingabe
                $txtSearchScripts.Add_TextChanged({
                    $searchText = $this.Text
                    
                    # Nichts tun, wenn Platzhaltertext
                    if ($searchText -eq "Skriptsuche...") {
                        return
                    }
                    
                    $spScriptList = $Global:window.FindName("spScriptList")
                    if ($null -eq $spScriptList) {
                        return
                    }
                    
                    # Alle Skriptbuttons durchgehen
                    foreach ($child in $spScriptList.Children) {
                        if ($child -is [System.Windows.Controls.Button]) {
                            $scriptButton = $child
                            $buttonContent = $scriptButton.Content
                            
                            # Textinhalte des Buttons extrahieren
                            $buttonText = ""
                            if ($buttonContent -is [System.Windows.Controls.StackPanel]) {
                                foreach ($element in $buttonContent.Children) {
                                    if ($element -is [System.Windows.Controls.TextBlock]) {
                                        $buttonText += $element.Text + " "
                                    }
                                }
                            }
                            
                            # Sichtbarkeit basierend auf Suchtext setzen
                            if ([string]::IsNullOrWhiteSpace($searchText) -or 
                                $buttonText -like "*$searchText*") {
                                $scriptButton.Visibility = "Visible"
                            } else {
                                $scriptButton.Visibility = "Collapsed"
                            }
                        }
                    }
                })
            }
            
            # Event-Handler für Button-Klicks in der ListView
            $lvModules = $window.FindName("lvModules")
            if ($lvModules) {
                $lvModules.AddHandler(
                    [System.Windows.Controls.Button]::ClickEvent,
                    [System.Windows.RoutedEventHandler]{
                        param($sender, $e)
                        
                        # Den genauen Button finden, der geklickt wurde
                        $button = $e.OriginalSource -as [System.Windows.Controls.Button]
                        if ($button -and $button.Tag) {
                            if ($button.Content -eq "Installieren") {
                                # Installieren-Button wurde geklickt
                                & $installModuleClickHandlerScript $button $e
                            } elseif ($button.Content -eq "Verbinden") {
                                # Verbinden-Button wurde geklickt
                                & $connectModuleClickHandlerScript $button $e
                            }
                        }
                        
                        # Event als behandelt markieren
                        $e.Handled = $true
                    }
                )
            }
            
        } catch {
            Write-Log "Fehler beim Registrieren von Event-Handlern: $($_.Exception.Message)" -Type "Warning"
        }
        
        return $window
    }
    catch {
        Write-Log "Fehler beim Laden des XAML: $($_.Exception.Message)" -Type "Error"
        throw
    }
}

# Graph Auth Dialog laden und anzeigen
function Show-GraphAuthDialog {
    $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($graphAuthXaml))
    $window = [System.Windows.Markup.XamlReader]::Load($reader)
    
    # Alle benannten Elemente verfügbar machen
    $namedNodes = $graphAuthXaml.Split("`n") | 
        Where-Object { $_ -match 'x:Name="([^"]*)"' } | 
        ForEach-Object { $matches[1] }
    
    $controls = @{}
    $namedNodes | ForEach-Object {
        $name = $_
        $control = $window.FindName($name)
        $controls[$name] = $control
    }
    
    # Info Text für verschiedene Authentifizierungsmethoden
    $authMethodInfos = @{
        0 = "Interaktive Anmeldung im Browser. Unterstützt MFA und alle delegierten Berechtigungen."
        1 = "Gerätecode-Authentifizierung erlaubt die Anmeldung ohne direkten Browserzugang."
        2 = "Nutzt ein vorhandenes Access Token (z.B. aus eigenem OAuth Flow)."
        3 = "Verwendet eine registrierte Azure AD App (öffentlicher Client)."
        4 = "App-Only Zugriff mit Zertifikatsauthentifizierung (höchste Sicherheit)."
        5 = "App-Only Zugriff mit Client Secret (einfacher, aber weniger sicher)."
        6 = "Nutzt die verwaltete Identität des Azure-Dienstes (nur in Azure-Umgebungen)."
    }
    
    # Event-Handler für Änderung der Auth-Methode - jetzt programmatisch hinzufügen
    $authMethodChangedScript = {
        param($sender, $e)
        
        $selectedIndex = $sender.SelectedIndex
        
        # Alle Panels ausblenden
        $controls["pnlAccessToken"].Visibility = "Collapsed"
        $controls["pnlClientApp"].Visibility = "Collapsed"
        $controls["pnlCertificate"].Visibility = "Collapsed"
        $controls["pnlClientSecret"].Visibility = "Collapsed"
        
        # Info-Text setzen
        $controls["txtInfo"].Text = $authMethodInfos[$selectedIndex]
        
        # Je nach ausgewählter Methode entsprechende Panels einblenden
        switch ($selectedIndex) {
            2 { # Access Token
                $controls["pnlAccessToken"].Visibility = "Visible"
            }
            3 { # Client App
                $controls["pnlClientApp"].Visibility = "Visible"
            }
            4 { # Certificate
                $controls["pnlClientApp"].Visibility = "Visible"
                $controls["pnlCertificate"].Visibility = "Visible"
            }
            5 { # Client Secret
                $controls["pnlClientApp"].Visibility = "Visible"
                $controls["pnlClientSecret"].Visibility = "Visible"
            }
        }
    }
    
    # Event-Handler programmatisch registrieren
    $controls["cmbAuthMethod"].Add_SelectionChanged($authMethodChangedScript)
    
    # Initialen Zustand setzen (manuell statt über Event auslösen)
    $initialIndex = $controls["cmbAuthMethod"].SelectedIndex
    $controls["txtInfo"].Text = $authMethodInfos[$initialIndex]
    
    # Event-Handler für den Connect-Button
    $controls["btnConnect"].Add_Click({
        $window.DialogResult = $true
        $window.Close()
    })
    
    # Event-Handler für den Cancel-Button
    $controls["btnCancel"].Add_Click({
        $window.DialogResult = $false
        $window.Close()
    })
    
    # Dialog anzeigen und Ergebnis zurückgeben
    $result = $window.ShowDialog()
    
    if ($result) {
        $selectedIndex = $controls["cmbAuthMethod"].SelectedIndex
        
        # Authentifizierungsmethode und Parameter basierend auf Auswahl
        $authParams = @{}
        
        switch ($selectedIndex) {
            0 { # Interaktive Anmeldung
                $authParams.AuthMethod = "Interactive"
            }
            1 { # Device Code
                $authParams.AuthMethod = "DeviceCode"
            }
            2 { # Access Token
                $authParams.AuthMethod = "AccessToken"
                $authParams.AccessToken = $controls["txtAccessToken"].Text
            }
            3 { # Client App
                $authParams.AuthMethod = "ClientApp"
                $authParams.ClientId = $controls["txtClientId"].Text
                $authParams.TenantId = $controls["txtTenantId"].Text
            }
            4 { # Certificate
                $authParams.AuthMethod = "CertificateBased"
                $authParams.ClientId = $controls["txtClientId"].Text
                $authParams.TenantId = $controls["txtTenantId"].Text
                $authParams.CertificateThumbprint = $controls["txtCertThumbprint"].Text
            }
            5 { # Client Secret
                $authParams.AuthMethod = "ClientSecret"
                $authParams.TenantId = $controls["txtTenantId"].Text
                
                # Client Secret Credential erstellen
                $clientId = $controls["txtClientId"].Text
                $secureSecret = ConvertTo-SecureString $controls["txtClientSecret"].Password -AsPlainText -Force
                $credential = New-Object System.Management.Automation.PSCredential ($clientId, $secureSecret)
                $authParams.ClientSecretCredential = $credential
            }
            6 { # Managed Identity
                $authParams.AuthMethod = "ManagedIdentity"
            }
        }
        
        return $authParams
    }
    
    return $null
}

# SharePoint Admin URL Dialog anzeigen
function Show-SPOConnectDialog {
    $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($spoConnectXaml))
    $window = [System.Windows.Markup.XamlReader]::Load($reader)
    
    $txtSpoUrl = $window.FindName("txtSpoUrl")
    $btnSpoConnect = $window.FindName("btnSpoConnect")
    $btnSpoCancel = $window.FindName("btnSpoCancel")
    
    # Standardwert für die URL setzen
    if ([string]::IsNullOrEmpty($txtSpoUrl.Text)) {
        $txtSpoUrl.Text = "https://"
    }
    
    # Event-Handler für den Connect-Button
    $btnSpoConnect.Add_Click({
        $window.DialogResult = $true
        $window.Close()
    })
    
    # Event-Handler für den Cancel-Button
    $btnSpoCancel.Add_Click({
        $window.DialogResult = $false
        $window.Close()
    })
    
    # Dialog anzeigen und URL zurückgeben
    $result = $window.ShowDialog()
    
    if ($result) {
        return $txtSpoUrl.Text
    }
    
    return $null
}

# Moduldaten vorbereiten - angepasst für ListView
function Get-ModuleData {
    $moduleData = @()
    
    foreach ($module in $requiredModules) {
        $moduleName = $module.Name
        $minVersion = $module.MinVersion
        $description = $module.Description
        $optional = $module.Optional
        $connectCmd = $module.ConnectCmd
        
        $installed = Test-ModuleInstalled -ModuleName $moduleName -MinVersion $minVersion
        $installedModule = Get-Module -ListAvailable -Name $moduleName | Sort-Object Version -Descending | Select-Object -First 1
        $installedVersion = if ($installedModule) { $installedModule.Version.ToString() } else { "Nicht installiert" }
        
        $status = if ($installed) { 
            "Bereit" 
        } elseif ($optional) { 
            "Optional" 
        } else { 
            "Erforderlich" 
        }
        
        $installEnabled = -not $installed
        
        # Connect-Button aktivieren, wenn:
        # 1. Modul installiert ist UND
        # 2. Ein Connect-Befehl vorhanden ist UND
        # 3. (Modul nicht optional ist ODER wir alle anzeigen)
        $connectEnabled = $installed -and (-not [string]::IsNullOrEmpty($connectCmd))
        
        # Prüfen, ob bereits verbunden
        $isConnected = $false
        if ($installed -and $module.ConnectCheck -ne $null) {
            try {
                $checkResult = & $module.ConnectCheck
                $isConnected = ($checkResult -ne $null)
            } catch {
                # Fehler beim Prüfen der Verbindung - nicht verbunden
                $isConnected = $false
            }
        }
        
        $moduleData += [PSCustomObject]@{
            Name = $moduleName
            Description = $description
            MinVersion = $minVersion
            InstalledVersion = $installedVersion
            Status = $status
            InstallEnabled = $installEnabled
            ConnectEnabled = $connectEnabled -and (-not $isConnected)
            Optional = $optional
            ConnectCmd = $connectCmd
            IsConnected = $isConnected
        }
    }
    
    return $moduleData
}

# Log-Anzeige aktualisieren
function Update-LogDisplay {
    param (
        [string]$Message,
        [ValidateSet("Info", "Warning", "Error", "Success")]
        [string]$Type = "Info"
    )
    
    # Log in die Datei schreiben
    Write-Log -Message $Message -Type $Type
    
    # TextBox aktualisieren
    $timestamp = Get-Date -Format "HH:mm:ss"
    $logEntry = "[$timestamp] "
    
    switch ($Type) {
        "Info"    { $logEntry += "[INFO] " }
        "Warning" { $logEntry += "[WARNUNG] " }
        "Error"   { $logEntry += "[FEHLER] " }
        "Success" { $logEntry += "[ERFOLG] " }
    }
    
    $logEntry += "$Message"
    
    if ($Global:txtLog) {
        $Global:txtLog.AppendText("$logEntry`n")
        $Global:txtLog.ScrollToEnd()
    }
}

# Prüfung aller Module durchführen - angepasst für ListView
function Check-AllModules {
    Update-LogDisplay -Message "Überprüfe Module..." -Type "Info"
    
    $moduleData = Get-ModuleData
    if ($Global:lvModules) {
        $Global:lvModules.ItemsSource = $moduleData
        Update-LogDisplay -Message "Module in ListView geladen." -Type "Info"
    } else {
        Update-LogDisplay -Message "ListView-Element nicht gefunden!" -Type "Error"
    }
    
    # Nur erforderliche (nicht optionale) Module prüfen
}

# Ereignishandler für Install-Buttons in DataGrid registrieren - verbesserte Version
function Register-ButtonEventHandlers {
    try {
        Update-LogDisplay -Message "Registriere Button-Ereignishandler..." -Type "Info"
        
        # Warten auf Rendering der UI-Elemente mit mehreren Versuchen
        $maxRetries = 5
        $retryCount = 0
        $success = $false
        
        while (-not $success -and $retryCount -lt $maxRetries) {
            try {
                $retryCount++
                $Global:window.Dispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Render, [System.Action]{})
                Start-Sleep -Milliseconds 100
                
                # Suche alle Zeilen in der DataGrid
                $rows = $Global:dgModules.Items
                if ($null -eq $rows -or $rows.Count -eq 0) {
                    Update-LogDisplay -Message "Keine Zeilen in der DataGrid gefunden (Versuch $retryCount)." -Type "Warning"
                    Start-Sleep -Seconds 1
                    continue
                }
                
                # DataGrid vollständig rendern
                $Global:dgModules.UpdateLayout()
                
                # Wir verwenden die Click-Events der Buttons direkt im Code-Behind
                $success = $true
                Update-LogDisplay -Message "Button-Handler werden direkt über Events registriert." -Type "Info"
            }
            catch {
                Update-LogDisplay -Message "Fehler beim Registrieren der Button-Handler (Versuch $retryCount): $($_.Exception.Message)" -Type "Warning"
                Start-Sleep -Seconds 1
            }
        }
        
        if (-not $success) {
            Update-LogDisplay -Message "Konnte Button-Handler nach $maxRetries Versuchen nicht registrieren." -Type "Error"
        }
    }
    catch {
        Update-LogDisplay -Message "Kritischer Fehler beim Registrieren der Button-Handler: $($_.Exception.Message)" -Type "Error"
    }
}

# Event-Handler für den Installationsbutton
function ModuleInstall_Click {
    param($sender, $e)
    
    $moduleName = $sender.Tag
    
    # Finde das entsprechende Modul in $requiredModules
    $moduleInfo = $requiredModules | Where-Object { $_.Name -eq $moduleName } | Select-Object -First 1
    
    if ($moduleInfo) {
        Update-LogDisplay -Message "Installation von $moduleName wird gestartet..." -Type "Info"
        
        # Button deaktivieren während der Installation
        $sender.IsEnabled = $false
        $sender.Content = "Installiere..."
        
        # Modul installieren
        $success = Install-RequiredModule -ModuleName $moduleName -MinVersion $moduleInfo.MinVersion
        
        if ($success) {
            Update-LogDisplay -Message "Modul $moduleName erfolgreich installiert." -Type "Success"
        } else {
            Update-LogDisplay -Message "Fehler bei der Installation von $moduleName." -Type "Error"
        }
        
        # Modulliste aktualisieren
        Check-AllModules
    }
}

# Event-Handler für den Connect-Button
function ModuleConnect_Click {
    param($sender, $e)
    
    $moduleName = $sender.Tag
    
    # Finde das entsprechende Modul in $requiredModules
    $moduleInfo = $requiredModules | Where-Object { $_.Name -eq $moduleName } | Select-Object -First 1
    
    if ($moduleInfo -and $moduleInfo.ConnectCmd) {
        Update-LogDisplay -Message "Verbindung zu $moduleName wird hergestellt..." -Type "Info"
        
        # Button deaktivieren während der Verbindung
        $sender.IsEnabled = $false
        $sender.Content = "Verbinde..."
        
        try {
            # Connect-Befehl ausführen
            $scriptBlock = [ScriptBlock]::Create($moduleInfo.ConnectCmd)
            Invoke-Command -ScriptBlock $scriptBlock
            
            # Verbindungsstatus prüfen
            $isConnected = $false
            if ($moduleInfo.ConnectCheck -ne $null) {
                try {
                    $checkResult = & $moduleInfo.ConnectCheck
                    $isConnected = ($checkResult -ne $null)
                } catch {
                    $isConnected = $false
                }
            }
            
            if ($isConnected) {
                Update-LogDisplay -Message "Verbindung zu $moduleName erfolgreich hergestellt." -Type "Success"
                $sender.Content = "Verbunden"
                $sender.IsEnabled = $false
                
                # Hauptskript-Button aktivieren, wenn mindestens ein Dienst verbunden ist
                $Global:btnStart.IsEnabled = $true
            } else {
                Update-LogDisplay -Message "Verbindung zu $moduleName konnte nicht verifiziert werden." -Type "Warning"
                $sender.Content = "Verbinden"
                $sender.IsEnabled = $true
            }
        } catch {
            Update-LogDisplay -Message "Fehler bei der Verbindung zu $moduleName - $($_.Exception.Message)" -Type "Error"
            $sender.Content = "Verbinden"
            $sender.IsEnabled = $true
        }
    }
}

# Event-Handler für den Refresh-Button
function Refresh_Click {
    param($sender, $e)
    
    Check-AllModules
}

# Verbesserte Verbindungsprüfung für Exchange Online
function Test-ExchangeOnlineConnection {
    try {
        Write-Log "Prüfe Exchange Online Verbindung..." -Type "Info"
        
        # 1. Prüfen auf aktive Exchange-Sessions
        $exoSessions = Get-PSSession | Where-Object { 
            $_.ConfigurationName -eq "Microsoft.Exchange" -and 
            ($_.State -eq "Opened" -or $_.State -eq "Available" -or $_.Availability -eq "Available")
        }
        
        if ($exoSessions.Count -gt 0) {
            Write-Log "Exchange Online Verbindung gefunden: $($exoSessions.Count) aktive Session(s)" -Type "Success"
            return $true
        }
        
        # 2. Alternative Prüfung: Versuche einen einfachen Exchange-Befehl
        try {
            # Get-AcceptedDomain ist ein schneller Befehl, der eine EXO-Verbindung bestätigt
            $domains = Get-AcceptedDomain -ErrorAction Stop | Select-Object -First 1
            if ($null -ne $domains) {
                Write-Log "Exchange Online Verbindung bestätigt durch erfolgreichen Befehlsaufruf" -Type "Success"
                return $true
            }
        } 
        catch {
            Write-Log "Exchange Online Befehlsausführung fehlgeschlagen: $($_.Exception.Message)" -Type "Warning"
        }
        
        # 3. Zusätzliche Prüfung mit Get-OrganizationConfig als Alternative
        try {
            $orgConfig = Get-OrganizationConfig -ErrorAction Stop
            if ($null -ne $orgConfig) {
                Write-Log "Exchange Online Verbindung bestätigt durch OrganizationConfig-Abruf" -Type "Success"
                return $true
            }
        }
        catch {
            Write-Log "Konnte OrganizationConfig nicht abrufen: $($_.Exception.Message)" -Type "Warning"
        }
        
        Write-Log "Keine aktive Exchange Online Verbindung gefunden." -Type "Warning"
        return $false
    }
    catch {
        Write-Log "Fehler bei der Prüfung der Exchange Online Verbindung: $($_.Exception.Message)" -Type "Error"
        return $false
    }
}

# Event-Handler für den Start-Button - mit verbesserter Sessionprüfung
function Start_Click {
    param($sender, $e)
    
    # Pfad zum Hauptskript bestimmen
    $mainScriptPath = Join-Path -Path $PSScriptRoot -ChildPath "easyEXO\easyEXO_V0.0.6.ps1"
    
    # Prüfen, ob das Hauptskript existiert
    if (-not (Test-Path -Path $mainScriptPath)) {
        Update-LogDisplay -Message "Hauptskript nicht gefunden: $mainScriptPath" -Type "Error"
        [System.Windows.MessageBox]::Show("Das Hauptskript konnte nicht gefunden werden.", "Fehler", 
            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        return
    }
    
    # Syntaxprüfung des Skripts vor der Ausführung
    Update-LogDisplay -Message "Validiere Hauptskript..." -Type "Info"
    try {
        $syntaxErrors = $null
        [System.Management.Automation.PSParser]::Tokenize((Get-Content -Path $mainScriptPath -Raw), [ref]$syntaxErrors)
        if ($syntaxErrors -and $syntaxErrors.Count -gt 0) {
            $errorMessage = "Das Hauptskript enthält Syntaxfehler:`n"
            foreach ($error in $syntaxErrors) {
                $errorMessage += "Zeile $($error.Token.StartLine), Spalte $($error.Token.StartColumn): $($error.Message)`n"
            }
            
            Update-LogDisplay -Message "Syntaxfehler im Hauptskript gefunden." -Type "Error"
            [System.Windows.MessageBox]::Show($errorMessage, "Syntaxfehler", 
                [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            return
        }
        
        Update-LogDisplay -Message "Hauptskript-Validierung erfolgreich." -Type "Success"
    }
    catch {
        Update-LogDisplay -Message "Fehler bei der Skriptvalidierung: $($_.Exception.Message)" -Type "Error"
        [System.Windows.MessageBox]::Show("Fehler bei der Syntaxprüfung: $($_.Exception.Message)", 
            "Validierungsfehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        return
    }
    
    # Prüfe die Verbindungen vor dem Start
    $exoConnected = Test-ExchangeOnlineConnection
    if (-not $exoConnected) {
        if ([System.Windows.MessageBox]::Show(
            "Es wurde keine aktive Exchange Online Verbindung gefunden. Das Hauptskript benötigt möglicherweise diese Verbindung. Möchten Sie trotzdem fortfahren?", 
            "Fehlende Verbindung", 
            [System.Windows.MessageBoxButton]::YesNo, 
            [System.Windows.MessageBoxImage]::Warning) -ne 'Yes') {
            return
        }
    }
    
    Update-LogDisplay -Message "Starte Hauptskript in der aktuellen Session: $mainScriptPath" -Type "Info"
    
    # Hauptskript ausführen, aber Fenster offen lassen
    try {
        Load-MainScript -ScriptPath $mainScriptPath -KeepSetupWindowOpen
    } catch {
        $errorMessage = $_.Exception.Message
        Write-Log "Fehler beim Ausführen des Hauptskripts: $errorMessage" -Type "Error"
    }
}

# Implementierung der Update-ScriptList Funktion (mit genehmigtem Verb)
function Update-ScriptList {
    try {
        Update-LogDisplay -Message "Lade verfügbare PowerShell Skripte..." -Type "Info"
        
        # Stack Panel für die Skriptliste holen
        $spScriptList = $Global:window.FindName("spScriptList")
        if ($null -eq $spScriptList) {
            Update-LogDisplay -Message "Skriptliste-Element nicht gefunden!" -Type "Error"
            return
        }
        
        # Stack Panel leeren
        $spScriptList.Children.Clear()
        
        # Basisordner sicherstellen (auch wenn $PSScriptRoot null ist)
        $baseFolder = if ([string]::IsNullOrEmpty($PSScriptRoot)) {
            # Fallback-Pfad verwenden - aktuelles Verzeichnis
            $currentPath = (Get-Location).Path
            Update-LogDisplay -Message "PSScriptRoot ist null, verwende aktuelles Verzeichnis: $currentPath" -Type "Warning"
            $currentPath
        } else {
            $PSScriptRoot
        }
        
        # Pfade definieren, in denen nach Skripten gesucht werden soll
        $scriptFolders = @(
            $baseFolder
        )
        
        # Unterordner hinzufügen, wenn sie existieren
        $potentialSubFolders = @("easyEXO", "easyEXOMigrate", "easySPO")
        foreach ($subFolder in $potentialSubFolders) {
            $fullPath = Join-Path -Path $baseFolder -ChildPath $subFolder
            if (Test-Path -Path $fullPath -ErrorAction SilentlyContinue) {
                $scriptFolders += $fullPath
            }
        }
        
        Update-LogDisplay -Message "Durchsuche folgende Ordner nach Skripten: $($scriptFolders -join ', ')" -Type "Info"
        
        # Skripte sammeln (PowerShell-Dateien)
        $scripts = @()
        foreach ($folder in $scriptFolders) {
            if (-not [string]::IsNullOrEmpty($folder) -and (Test-Path -Path $folder -ErrorAction SilentlyContinue)) {
                try {
                    $folderScripts = Get-ChildItem -Path $folder -Filter "*.ps1" -File -ErrorAction SilentlyContinue | 
                                    Where-Object { $_.Name -ne (Split-Path -Leaf $MyInvocation.MyCommand.Path) }
                    if ($folderScripts) {
                        $scripts += $folderScripts
                        Update-LogDisplay -Message "Gefundene Skripte in $folder - $($folderScripts.Count)" -Type "Info"
                    }
                } catch {
                    Update-LogDisplay -Message "Fehler beim Durchsuchen von $folder - $($_.Exception.Message)" -Type "Error"
                }
            }
        }
        
        if ($scripts.Count -eq 0) {
            Update-LogDisplay -Message "Keine PowerShell-Skripte gefunden." -Type "Warning"
            
            # Placeholder anzeigen
            $textBlock = New-Object System.Windows.Controls.TextBlock
            $textBlock.Text = "Keine Skripte gefunden."
            $textBlock.Margin = "10,5"
            $textBlock.Foreground = "Gray"
            $textBlock.FontStyle = "Italic"
            
            $spScriptList.Children.Add($textBlock)
            return
        }
        
        # Skripte nach Name sortieren
        $scripts = $scripts | Sort-Object -Property Name
        
        # Scripts in Kategorien gruppieren
        $scriptsByCategory = @{}
        
        foreach ($script in $scripts) {
            # Kategorie aus Dateinamen oder Pfad ableiten
            $category = "Allgemein"
            
            # Aus Verzeichnis ableiten
            $parentDir = Split-Path -Leaf (Split-Path -Parent $script.FullName)
            if ($parentDir -ne (Split-Path -Leaf $PSScriptRoot)) {
                $category = $parentDir
            }
            
            # Aus Dateinamen-Präfix ableiten
            if ($script.Name -match "^(easy\w+)_") {
                $category = $matches[1]
            }
            
            # In Kategorieliste hinzufügen
            if (-not $scriptsByCategory.ContainsKey($category)) {
                $scriptsByCategory[$category] = @()
            }
            $scriptsByCategory[$category] += $script
        }
        
        # Kategorien in Alphabetischer Reihenfolge durchgehen
        foreach ($category in ($scriptsByCategory.Keys | Sort-Object)) {
            # Kategorie-Header erstellen
            $headerBorder = New-Object System.Windows.Controls.Border
            $headerBorder.Background = "#E3F2FD"
            $headerBorder.Margin = "0,10,0,5"
            $headerBorder.Padding = "5"
            $headerBorder.CornerRadius = "3"
            
            $headerText = New-Object System.Windows.Controls.TextBlock
            $headerText.Text = $category
            $headerText.FontWeight = "Bold"
            
            $headerBorder.Child = $headerText
            $spScriptList.Children.Add($headerBorder)
            
            # Skripte in dieser Kategorie hinzufügen
            foreach ($script in $scriptsByCategory[$category]) {
                # Skriptname und Beschreibung extrahieren
                $scriptName = $script.Name
                $scriptDescription = ""
                
                # Versuchen, Synopsis aus Skript-Header zu extrahieren
                try {
                    if (Test-Path -Path $script.FullName -ErrorAction SilentlyContinue) {
                        $content = Get-Content -Path $script.FullName -TotalCount 30 -ErrorAction SilentlyContinue
                        $synopsisMatch = $content | Select-String -Pattern "\.SYNOPSIS\s+(.*)" -ErrorAction SilentlyContinue
                        if ($synopsisMatch) {
                            $scriptDescription = $synopsisMatch.Matches.Groups[1].Value.Trim()
                        }
                    }
                } catch {
                    # Fehler ignorieren
                }
                
                # Button für das Skript erstellen
                $button = New-Object System.Windows.Controls.Button
                $button.Style = $Global:window.Resources["ScriptButtonStyle"]
                $button.Tag = $script.FullName
                $button.Margin = "5,2"
                $button.HorizontalContentAlignment = "Left"
                
                # Stack Panel für Button-Inhalt
                $buttonContent = New-Object System.Windows.Controls.StackPanel
                $buttonContent.Orientation = "Vertical"
                
                # Skript-Name
                $nameText = New-Object System.Windows.Controls.TextBlock
                $nameText.Text = $scriptName
                $nameText.FontWeight = "Bold"
                $buttonContent.Children.Add($nameText)
                
                # Skript-Beschreibung (nur wenn vorhanden)
                if (-not [string]::IsNullOrWhiteSpace($scriptDescription)) {
                    $descText = New-Object System.Windows.Controls.TextBlock
                    $descText.Text = $scriptDescription
                    $descText.TextTrimming = "CharacterEllipsis"
                    $descText.Opacity = 0.7
                    $buttonContent.Children.Add($descText)
                }
                
                $button.Content = $buttonContent
                $button.Add_Click({
                    param($clickSource, $eventArgs)
                    
                    $scriptPath = $clickSource.Tag
                    
                    # Bestätigungsdialog anzeigen
                    $scriptName = Split-Path -Leaf $scriptPath
                    $result = [System.Windows.MessageBox]::Show(
                        "Möchten Sie das Skript '$scriptName' in der aktuellen Session ausführen?",
                        "Skript ausführen",
                        [System.Windows.MessageBoxButton]::YesNo,
                        [System.Windows.MessageBoxImage]::Question
                    )
                    
                    if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
                        # Hauptskript in aktueller Session ausführen
                        try {
                            Update-LogDisplay -Message "Führe Skript aus: $scriptPath" -Type "Info"
                            Load-MainScript -ScriptPath $scriptPath -KeepSetupWindowOpen
                        }
                        catch {
                            Update-LogDisplay -Message "Fehler beim Ausführen des Skripts - $($_.Exception.Message)" -Type "Error"
                            [System.Windows.MessageBox]::Show(
                                "Fehler beim Ausführen des Skripts: $($_.Exception.Message)",
                                "Skriptfehler",
                                [System.Windows.MessageBoxButton]::OK,
                                [System.Windows.MessageBoxImage]::Error
                            )
                        }
                    }
                })
                
                # Button zur Liste hinzufügen
                $spScriptList.Children.Add($button)
            }
        }
        
        Update-LogDisplay -Message "Skriptliste wurde geladen - $($scripts.Count) Skripte gefunden." -Type "Success"
    }
    catch {
        Update-LogDisplay -Message "Fehler beim Laden der Skriptliste - $($_.Exception.Message)" -Type "Error"
    }
}

# Alias für Update-ScriptList erstellen
function Load-ScriptList {
    Update-ScriptList
}

# Initialisierung und Anzeige der GUI mit verbesserter Fehlerbehandlung
try {
    Write-Log "Starte easyEXO Setup Tool..." -Type "Info"
    
    # GUI laden
    $Global:window = Import-XamlGui
    if (-not $Global:window) {
        throw "GUI konnte nicht geladen werden."
    }
    
    # Module überprüfen
    Check-AllModules
    # Skriptliste laden
    Update-ScriptList
    
    # GUI anzeigen
    [void]$Global:window.ShowDialog()
}
catch {
    $errorMessage = $_.Exception.Message
    $lineNumber = $_.InvocationInfo.ScriptLineNumber
    Write-Log "Fehler bei der Ausführung des Setup-Tools (Zeile $lineNumber) - $errorMessage" -Type "Error"
    
    [System.Windows.MessageBox]::Show("Ein Fehler ist aufgetreten - $errorMessage", "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
}

# Verbesserte PnP-Verbindungsmethode
function Connect-ToPnPOnline {
    param (
        [Parameter(Mandatory = $false)]
        [string]$Url = ""
    )
    
    try {
        # Prüfen, ob das Modul installiert ist
        if (-not (Test-ModuleInstalled -ModuleName "PnP.PowerShell")) {
            Update-LogDisplay -Message "PnP.PowerShell Modul nicht installiert" -Type "Error"
            return $false
        }
        
        # Modul importieren
        Import-Module -Name PnP.PowerShell -ErrorAction Stop
        Update-LogDisplay -Message "PnP.PowerShell Modul importiert" -Type "Info"
        
        # Wenn keine URL angegeben wurde, Dialog anzeigen
        if ([string]::IsNullOrEmpty($Url)) {
            # Dialog für SharePoint URL anzeigen (wiederverwendbar)
            $spoDialog = [Windows.Forms.Form]@{
                Width = 450
                Height = 200
                Text = "SharePoint Site Verbindung"
                StartPosition = "CenterScreen"
                FormBorderStyle = "FixedDialog"
                MaximizeBox = $false
                MinimizeBox = $false
            }
            
            $lblInfo = [Windows.Forms.Label]@{
                Text = "Bitte geben Sie die vollständige SharePoint Site URL ein:"
                Location = New-Object Drawing.Point(20, 20)
                Size = New-Object Drawing.Size(400, 20)
            }
            
            $txtUrl = [Windows.Forms.TextBox]@{
                Text = "https://"
                Location = New-Object Drawing.Point(20, 50)
                Size = New-Object Drawing.Size(400, 20)
            }
            
            $btnConnect = [Windows.Forms.Button]@{
                Text = "Verbinden"
                Location = New-Object Drawing.Point(245, 120)
                Size = New-Object Drawing.Size(100, 30)
                DialogResult = [Windows.Forms.DialogResult]::OK
            }
            
            $btnCancel = [Windows.Forms.Button]@{
                Text = "Abbrechen"
                Location = New-Object Drawing.Point(345, 120)
                Size = New-Object Drawing.Size(80, 30)
                DialogResult = [Windows.Forms.DialogResult]::Cancel
            }
            
            $spoDialog.Controls.AddRange(@($lblInfo, $txtUrl, $btnConnect, $btnCancel))
            $spoDialog.AcceptButton = $btnConnect
            $spoDialog.CancelButton = $btnCancel
            
            $result = $spoDialog.ShowDialog()
            
            if ($result -eq [Windows.Forms.DialogResult]::OK) {
                $Url = $txtUrl.Text.Trim()
            } else {
                Update-LogDisplay -Message "PnP Verbindung vom Benutzer abgebrochen" -Type "Warning"
                return $false
            }
        }
        
        if ([string]::IsNullOrEmpty($Url)) {
            Update-LogDisplay -Message "Keine SharePoint URL angegeben" -Type "Error"
            return $false
        }
        
        Update-LogDisplay -Message "Verbinde mit SharePoint Site: $Url" -Type "Info"
        Connect-PnPOnline -Url $Url -Interactive -ErrorAction Stop
        
        # Verbindung prüfen
        $context = Get-PnPContext -ErrorAction Stop
        if ($null -ne $context) {
            Update-LogDisplay -Message "PnP-Verbindung erfolgreich hergestellt" -Type "Success"
            return $true
        }
        
        Update-LogDisplay -Message "PnP-Verbindung konnte nicht verifiziert werden" -Type "Warning"
        return $false
    }
    catch {
        Update-LogDisplay -Message "Fehler bei der PnP-Verbindung: $($_.Exception.Message)" -Type "Error"
        return $false
    }
}

# Verbesserte SharePoint Online Verbindungsfunktion
function Connect-ToSharePointOnline {
    param(
        [string]$AdminUrl = ""
    )
    
    try {
        # Prüfen, ob das Modul installiert ist
        if (-not (Test-ModuleInstalled -ModuleName "Microsoft.Online.SharePoint.PowerShell")) {
            Update-LogDisplay -Message "SharePoint Online PowerShell Modul nicht installiert" -Type "Error"
            return $false
        }
        
        # Modul importieren mit Kompatibilitätsprüfung
        try {
            Import-Module -Name Microsoft.Online.SharePoint.PowerShell -DisableNameChecking -ErrorAction Stop
            Update-LogDisplay -Message "SharePoint Online PowerShell Modul importiert" -Type "Info"
        }
        catch {
            Update-LogDisplay -Message "Problem beim Laden des SharePoint Moduls: $($_.Exception.Message)" -Type "Error"
            
            # Prüfen auf bekanntes Kompatibilitätsproblem
            if ($_.Exception.Message -like "*Could not load type*SharePointTenantSettingCategory*") {
                Update-LogDisplay -Message "Bekanntes Kompatibilitätsproblem mit SharePoint-Modul erkannt" -Type "Warning"
                Update-LogDisplay -Message "Versuche alternativen Ansatz mit PnP PowerShell..." -Type "Info"
                
                # PnP PowerShell als Alternative vorschlagen
                if (Test-ModuleInstalled -ModuleName "PnP.PowerShell") {
                    if ([System.Windows.MessageBox]::Show(
                        "Das SharePoint Online PowerShell Modul hat ein Kompatibilitätsproblem. Möchten Sie stattdessen PnP PowerShell verwenden?",
                        "Modul-Kompatibilitätsproblem",
                        [System.Windows.MessageBoxButton]::YesNo,
                        [System.Windows.MessageBoxImage]::Question) -eq 'Yes') {
                        
                        return Connect-ToPnPOnline
                    }
                } else {
                    Update-LogDisplay -Message "Bitte installieren Sie das PnP.PowerShell Modul als Alternative" -Type "Warning"
                }
            }
            
            return $false
        }
        
        # Wenn keine URL angegeben wurde, SharePoint Admin URL Dialog anzeigen
        if ([string]::IsNullOrEmpty($AdminUrl)) {
            $AdminUrl = Show-SPOConnectDialog
            
            if ([string]::IsNullOrEmpty($AdminUrl)) {
                Update-LogDisplay -Message "SharePoint Admin URL nicht angegeben" -Type "Warning"
                return $false
            }
        }
        
        Update-LogDisplay -Message "Verbinde mit SharePoint Online Admin: $AdminUrl" -Type "Info"
        
        try {
            # Verbindung mit besserer Fehlerbehandlung herstellen
            Connect-SPOService -Url $AdminUrl -ErrorAction Stop
            
            # Verbindung testen
            $tenantInfo = Get-SPOTenant -ErrorAction Stop
            
            if ($null -ne $tenantInfo) {
                Update-LogDisplay -Message "SharePoint Online Verbindung erfolgreich hergestellt" -Type "Success"
                return $true
            }
        }
        catch [System.Management.Automation.CommandNotFoundException] {
            Update-LogDisplay -Message "SharePoint Online Befehl nicht gefunden. Dies kann auf ein Problem mit dem Modul hinweisen." -Type "Error"
            return $false
        }
        catch {
            Update-LogDisplay -Message "Fehler bei der SharePoint Online Verbindung - $($_.Exception.Message)" -Type "Error"
            return $false
        }
        
        Update-LogDisplay -Message "SharePoint Online Verbindung konnte nicht verifiziert werden" -Type "Warning"
        return $false
    }
    catch {
        Update-LogDisplay -Message "Fehler beim SharePoint Online Verbindungsversuch - $($_.Exception.Message)" -Type "Error"
        return $false
    }
}

# Verbesserte MSOnline Verbindungsfunktion
function Connect-ToMSOnline {
    try {
        # Prüfen, ob das Modul installiert ist
        if (-not (Test-ModuleInstalled -ModuleName "MSOnline")) {
            Update-LogDisplay -Message "MSOnline Modul nicht installiert" -Type "Error"
            return $false
        }
        
        # Modul importieren
        Import-Module -Name MSOnline -ErrorAction Stop
        Update-LogDisplay -Message "MSOnline Modul importiert" -Type "Info"
        
        # Prüfen auf bestehende Verbindung
        try {
            $companyInfo = Get-MsolCompanyInformation -ErrorAction SilentlyContinue
            if ($companyInfo) {
                Update-LogDisplay -Message "Bereits mit MSOnline verbunden" -Type "Success"
                return $true
            }
        }
        catch {
            # Ignorieren - wir versuchen neu zu verbinden
        }
        
        # Verbindung mit geänderter Authentication-Methode herstellen
        Update-LogDisplay -Message "Verbinde mit MSOnline Service..." -Type "Info"
        
        try {
            # Moderne Authentifizierung mit verbesserter Methode
            $auth = Get-Credential -Message "MSOnline Anmeldedaten eingeben" -ErrorAction Stop
            if ($null -eq $auth) {
                Update-LogDisplay -Message "Anmeldung abgebrochen" -Type "Warning"
                return $false
            }
            
            # Anmelden mit ADAL-Authentifizierung vermeiden
            Connect-MsolService -Credential $auth -ErrorAction Stop
            
            # Verbindung prüfen
            $companyInfo = Get-MsolCompanyInformation -ErrorAction Stop
            if ($companyInfo) {
                Update-LogDisplay -Message "MSOnline Verbindung erfolgreich hergestellt" -Type "Success"
                return $true
            }
        }
        catch {
            Update-LogDisplay -Message "Fehler bei der MSOnline Verbindung: $($_.Exception.Message)" -Type "Error"
            
            # Spezifischer Fehler für Redirect-URI Problem
            if ($_.Exception.Message -like "*Only loopback redirect uri is supported*") {
                Update-LogDisplay -Message "MSOnline hat ein Problem mit dem Authentifizierungs-Redirect. Versuchen Sie die Microsoft.Graph-Module als Alternative." -Type "Warning"
            }
            
            return $false
        }
        
        Update-LogDisplay -Message "MSOnline Verbindung konnte nicht verifiziert werden" -Type "Warning"
        return $false
    }
    catch {
        Update-LogDisplay -Message "Fehler beim MSOnline Verbindungsversuch: $($_.Exception.Message)" -Type "Error"
        return $false
    }
}

# Event-Handler für den Connect-Modul-Handler - mit verbesserten Methoden
$connectModuleClickHandlerScript = {
    param($sender, $e)
    
    $button = [System.Windows.Controls.Button]$sender
    $moduleName = $button.Tag.ToString()
    
    # Finde das entsprechende Modul in $requiredModules
    $moduleInfo = $requiredModules | Where-Object { $_.Name -eq $moduleName } | Select-Object -First 1
    
    if ($moduleInfo) {
        Update-LogDisplay -Message "Verbindung zu $moduleName wird hergestellt..." -Type "Info"
        
        # Button deaktivieren während der Verbindung
        $button.IsEnabled = $false
        $button.Content = "Verbinde..."
        
        try {
            # Modul vorab importieren, um sicherzustellen, dass alle Cmdlets verfügbar sind
            if (-not (Get-Module -Name $moduleName -ErrorAction SilentlyContinue)) {
                Update-LogDisplay -Message "Importiere Modul $moduleName..." -Type "Info"
                Import-Module -Name $moduleName -ErrorAction Stop -Verbose:$false
            }
            
            $isConnected = $false
            
            # Spezialfälle mit verbesserten Verbindungsmethoden
            switch ($moduleName) {
                "Microsoft.Online.SharePoint.PowerShell" {
                    $isConnected = Connect-ToSharePointOnline
                }
                "ExchangeOnlineManagement" {
                    $isConnected = Connect-ToExchangeOnline
                }
                "Microsoft.Graph" {
                    $authParams = Show-GraphAuthDialog
                    if ($null -ne $authParams) {
                        $isConnected = Connect-ToMicrosoftGraph @authParams
                    } else {
                        Update-LogDisplay -Message "Graph-Verbindung abgebrochen." -Type "Warning"
                    }
                }
                "PnP.PowerShell" {
                    $isConnected = Connect-ToPnPOnline
                }
                "MSOnline" {
                    $isConnected = Connect-ToMSOnline
                }
                default {
                    # Standard-Verbindungsmethode für andere Module
                    if ($moduleInfo.ConnectCmd) {
                        $scriptBlock = [ScriptBlock]::Create($moduleInfo.ConnectCmd)
                        & $scriptBlock
                        
                        # Kurz warten und Verbindung prüfen
                        Start-Sleep -Seconds 2
                        
                        if ($moduleInfo.ConnectCheck -ne $null) {
                            try {
                                $checkResult = & $moduleInfo.ConnectCheck
                                $isConnected = ($null -ne $checkResult)
                            } catch {
                                $isConnected = $false
                            }
                        } else {
                            # Ohne Prüfmethode nehmen wir an, dass die Verbindung erfolgreich war
                            $isConnected = $true
                        }
                    }
                }
            }
            
            if ($isConnected) {
                Update-LogDisplay -Message "Verbindung zu $moduleName erfolgreich hergestellt." -Type "Success"
                $button.Content = "Verbunden"
                $button.IsEnabled = $false
            } else {
                Update-LogDisplay -Message "Verbindung zu $moduleName konnte nicht hergestellt werden." -Type "Warning"
                $button.Content = "Verbinden"
                $button.IsEnabled = $true
            }
        } catch {
            Update-LogDisplay -Message "Fehler bei der Verbindung zu $moduleName - $($_.Exception.Message)" -Type "Error"
            $button.Content = "Verbinden"
            $button.IsEnabled = $true
        }
    }
}

# SIG # Begin signature block
# MIIbywYJKoZIhvcNAQcCoIIbvDCCG7gCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCLAk3b+dlfQA+t
# sokdQeevPbxj9yguLXsHxqKuzp9BnqCCFhcwggMQMIIB+KADAgECAhB3jzsyX9Cg
# jEi+sBC2rBMTMA0GCSqGSIb3DQEBCwUAMCAxHjAcBgNVBAMMFVBoaW5JVC1QU3Nj
# cmlwdHNfU2lnbjAeFw0yNTA3MDUwODI4MTZaFw0yNzA3MDUwODM4MTZaMCAxHjAc
# BgNVBAMMFVBoaW5JVC1QU3NjcmlwdHNfU2lnbjCCASIwDQYJKoZIhvcNAQEBBQAD
# ggEPADCCAQoCggEBALmz3o//iDA5MvAndTjGX7/AvzTSACClfuUR9WYK0f6Ut2dI
# mPxn+Y9pZlLjXIpZT0H2Lvxq5aSI+aYeFtuJ8/0lULYNCVT31Bf+HxervRBKsUyi
# W9+4PH6STxo3Pl4l56UNQMcWLPNjDORWRPWHn0f99iNtjI+L4tUC/LoWSs3obzxN
# 3uTypzlaPBxis2qFSTR5SWqFdZdRkcuI5LNsJjyc/QWdTYRrfmVqp0QrvcxzCv8u
# EiVuni6jkXfiE6wz+oeI3L2iR+ywmU6CUX4tPWoS9VTtmm7AhEpasRTmrrnSg20Q
# jiBa1eH5TyLAH3TcYMxhfMbN9a2xDX5pzM65EJUCAwEAAaNGMEQwDgYDVR0PAQH/
# BAQDAgeAMBMGA1UdJQQMMAoGCCsGAQUFBwMDMB0GA1UdDgQWBBQO7XOqiE/EYi+n
# IaR6YO5M2MUuVTANBgkqhkiG9w0BAQsFAAOCAQEAjYOKIwBu1pfbdvEFFaR/uY88
# peKPk0NnvNEc3dpGdOv+Fsgbz27JPvItITFd6AKMoN1W48YjQLaU22M2jdhjGN5i
# FSobznP5KgQCDkRsuoDKiIOTiKAAknjhoBaCCEZGw8SZgKJtWzbST36Thsdd/won
# ihLsuoLxfcFnmBfrXh3rTIvTwvfujob68s0Sf5derHP/F+nphTymlg+y4VTEAijk
# g2dhy8RAsbS2JYZT7K5aEJpPXMiOLBqd7oTGfM7y5sLk2LIM4cT8hzgz3v5yPMkF
# H2MdR//K403e1EKH9MsGuGAJZddVN8ppaiESoPLoXrgnw2SY5KCmhYw1xRFdjTCC
# BY0wggR1oAMCAQICEA6bGI750C3n79tQ4ghAGFowDQYJKoZIhvcNAQEMBQAwZTEL
# MAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3
# LmRpZ2ljZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBSb290
# IENBMB4XDTIyMDgwMTAwMDAwMFoXDTMxMTEwOTIzNTk1OVowYjELMAkGA1UEBhMC
# VVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0
# LmNvbTEhMB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MIICIjANBgkq
# hkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAv+aQc2jeu+RdSjwwIjBpM+zCpyUuySE9
# 8orYWcLhKac9WKt2ms2uexuEDcQwH/MbpDgW61bGl20dq7J58soR0uRf1gU8Ug9S
# H8aeFaV+vp+pVxZZVXKvaJNwwrK6dZlqczKU0RBEEC7fgvMHhOZ0O21x4i0MG+4g
# 1ckgHWMpLc7sXk7Ik/ghYZs06wXGXuxbGrzryc/NrDRAX7F6Zu53yEioZldXn1RY
# jgwrt0+nMNlW7sp7XeOtyU9e5TXnMcvak17cjo+A2raRmECQecN4x7axxLVqGDgD
# EI3Y1DekLgV9iPWCPhCRcKtVgkEy19sEcypukQF8IUzUvK4bA3VdeGbZOjFEmjNA
# vwjXWkmkwuapoGfdpCe8oU85tRFYF/ckXEaPZPfBaYh2mHY9WV1CdoeJl2l6SPDg
# ohIbZpp0yt5LHucOY67m1O+SkjqePdwA5EUlibaaRBkrfsCUtNJhbesz2cXfSwQA
# zH0clcOP9yGyshG3u3/y1YxwLEFgqrFjGESVGnZifvaAsPvoZKYz0YkH4b235kOk
# GLimdwHhD5QMIR2yVCkliWzlDlJRR3S+Jqy2QXXeeqxfjT/JvNNBERJb5RBQ6zHF
# ynIWIgnffEx1P2PsIV/EIFFrb7GrhotPwtZFX50g/KEexcCPorF+CiaZ9eRpL5gd
# LfXZqbId5RsCAwEAAaOCATowggE2MA8GA1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYE
# FOzX44LScV1kTN8uZz/nupiuHA9PMB8GA1UdIwQYMBaAFEXroq/0ksuCMS1Ri6en
# IZ3zbcgPMA4GA1UdDwEB/wQEAwIBhjB5BggrBgEFBQcBAQRtMGswJAYIKwYBBQUH
# MAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3aHR0cDov
# L2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNy
# dDBFBgNVHR8EPjA8MDqgOKA2hjRodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGln
# aUNlcnRBc3N1cmVkSURSb290Q0EuY3JsMBEGA1UdIAQKMAgwBgYEVR0gADANBgkq
# hkiG9w0BAQwFAAOCAQEAcKC/Q1xV5zhfoKN0Gz22Ftf3v1cHvZqsoYcs7IVeqRq7
# IviHGmlUIu2kiHdtvRoU9BNKei8ttzjv9P+Aufih9/Jy3iS8UgPITtAq3votVs/5
# 9PesMHqai7Je1M/RQ0SbQyHrlnKhSLSZy51PpwYDE3cnRNTnf+hZqPC/Lwum6fI0
# POz3A8eHqNJMQBk1RmppVLC4oVaO7KTVPeix3P0c2PR3WlxUjG/voVA9/HYJaISf
# b8rbII01YBwCA8sgsKxYoA5AY8WYIsGyWfVVa88nq2x2zm8jLfR+cWojayL/ErhU
# LSd+2DrZ8LaHlv1b0VysGMNNn3O3AamfV6peKOK5lDCCBq4wggSWoAMCAQICEAc2
# N7ckVHzYR6z9KGYqXlswDQYJKoZIhvcNAQELBQAwYjELMAkGA1UEBhMCVVMxFTAT
# BgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEh
# MB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MB4XDTIyMDMyMzAwMDAw
# MFoXDTM3MDMyMjIzNTk1OVowYzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lD
# ZXJ0LCBJbmMuMTswOQYDVQQDEzJEaWdpQ2VydCBUcnVzdGVkIEc0IFJTQTQwOTYg
# U0hBMjU2IFRpbWVTdGFtcGluZyBDQTCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCC
# AgoCggIBAMaGNQZJs8E9cklRVcclA8TykTepl1Gh1tKD0Z5Mom2gsMyD+Vr2EaFE
# FUJfpIjzaPp985yJC3+dH54PMx9QEwsmc5Zt+FeoAn39Q7SE2hHxc7Gz7iuAhIoi
# GN/r2j3EF3+rGSs+QtxnjupRPfDWVtTnKC3r07G1decfBmWNlCnT2exp39mQh0YA
# e9tEQYncfGpXevA3eZ9drMvohGS0UvJ2R/dhgxndX7RUCyFobjchu0CsX7LeSn3O
# 9TkSZ+8OpWNs5KbFHc02DVzV5huowWR0QKfAcsW6Th+xtVhNef7Xj3OTrCw54qVI
# 1vCwMROpVymWJy71h6aPTnYVVSZwmCZ/oBpHIEPjQ2OAe3VuJyWQmDo4EbP29p7m
# O1vsgd4iFNmCKseSv6De4z6ic/rnH1pslPJSlRErWHRAKKtzQ87fSqEcazjFKfPK
# qpZzQmiftkaznTqj1QPgv/CiPMpC3BhIfxQ0z9JMq++bPf4OuGQq+nUoJEHtQr8F
# nGZJUlD0UfM2SU2LINIsVzV5K6jzRWC8I41Y99xh3pP+OcD5sjClTNfpmEpYPtMD
# iP6zj9NeS3YSUZPJjAw7W4oiqMEmCPkUEBIDfV8ju2TjY+Cm4T72wnSyPx4Jduyr
# XUZ14mCjWAkBKAAOhFTuzuldyF4wEr1GnrXTdrnSDmuZDNIztM2xAgMBAAGjggFd
# MIIBWTASBgNVHRMBAf8ECDAGAQH/AgEAMB0GA1UdDgQWBBS6FtltTYUvcyl2mi91
# jGogj57IbzAfBgNVHSMEGDAWgBTs1+OC0nFdZEzfLmc/57qYrhwPTzAOBgNVHQ8B
# Af8EBAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwgwdwYIKwYBBQUHAQEEazBpMCQG
# CCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQQYIKwYBBQUHMAKG
# NWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRSb290
# RzQuY3J0MEMGA1UdHwQ8MDowOKA2oDSGMmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNv
# bS9EaWdpQ2VydFRydXN0ZWRSb290RzQuY3JsMCAGA1UdIAQZMBcwCAYGZ4EMAQQC
# MAsGCWCGSAGG/WwHATANBgkqhkiG9w0BAQsFAAOCAgEAfVmOwJO2b5ipRCIBfmbW
# 2CFC4bAYLhBNE88wU86/GPvHUF3iSyn7cIoNqilp/GnBzx0H6T5gyNgL5Vxb122H
# +oQgJTQxZ822EpZvxFBMYh0MCIKoFr2pVs8Vc40BIiXOlWk/R3f7cnQU1/+rT4os
# equFzUNf7WC2qk+RZp4snuCKrOX9jLxkJodskr2dfNBwCnzvqLx1T7pa96kQsl3p
# /yhUifDVinF2ZdrM8HKjI/rAJ4JErpknG6skHibBt94q6/aesXmZgaNWhqsKRcnf
# xI2g55j7+6adcq/Ex8HBanHZxhOACcS2n82HhyS7T6NJuXdmkfFynOlLAlKnN36T
# U6w7HQhJD5TNOXrd/yVjmScsPT9rp/Fmw0HNT7ZAmyEhQNC3EyTN3B14OuSereU0
# cZLXJmvkOHOrpgFPvT87eK1MrfvElXvtCl8zOYdBeHo46Zzh3SP9HSjTx/no8Zhf
# +yvYfvJGnXUsHicsJttvFXseGYs2uJPU5vIXmVnKcPA3v5gA3yAWTyf7YGcWoWa6
# 3VXAOimGsJigK+2VQbc61RWYMbRiCQ8KvYHZE/6/pNHzV9m8BPqC3jLfBInwAM1d
# wvnQI38AC+R2AibZ8GV2QqYphwlHK+Z/GqSFD/yYlvZVVCsfgPrA8g4r5db7qS9E
# FUrnEw4d2zc4GqEr9u3WfPwwgga8MIIEpKADAgECAhALrma8Wrp/lYfG+ekE4zME
# MA0GCSqGSIb3DQEBCwUAMGMxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2Vy
# dCwgSW5jLjE7MDkGA1UEAxMyRGlnaUNlcnQgVHJ1c3RlZCBHNCBSU0E0MDk2IFNI
# QTI1NiBUaW1lU3RhbXBpbmcgQ0EwHhcNMjQwOTI2MDAwMDAwWhcNMzUxMTI1MjM1
# OTU5WjBCMQswCQYDVQQGEwJVUzERMA8GA1UEChMIRGlnaUNlcnQxIDAeBgNVBAMT
# F0RpZ2lDZXJ0IFRpbWVzdGFtcCAyMDI0MIICIjANBgkqhkiG9w0BAQEFAAOCAg8A
# MIICCgKCAgEAvmpzn/aVIauWMLpbbeZZo7Xo/ZEfGMSIO2qZ46XB/QowIEMSvgjE
# dEZ3v4vrrTHleW1JWGErrjOL0J4L0HqVR1czSzvUQ5xF7z4IQmn7dHY7yijvoQ7u
# jm0u6yXF2v1CrzZopykD07/9fpAT4BxpT9vJoJqAsP8YuhRvflJ9YeHjes4fduks
# THulntq9WelRWY++TFPxzZrbILRYynyEy7rS1lHQKFpXvo2GePfsMRhNf1F41nyE
# g5h7iOXv+vjX0K8RhUisfqw3TTLHj1uhS66YX2LZPxS4oaf33rp9HlfqSBePejlY
# eEdU740GKQM7SaVSH3TbBL8R6HwX9QVpGnXPlKdE4fBIn5BBFnV+KwPxRNUNK6lY
# k2y1WSKour4hJN0SMkoaNV8hyyADiX1xuTxKaXN12HgR+8WulU2d6zhzXomJ2Ple
# I9V2yfmfXSPGYanGgxzqI+ShoOGLomMd3mJt92nm7Mheng/TBeSA2z4I78JpwGpT
# RHiT7yHqBiV2ngUIyCtd0pZ8zg3S7bk4QC4RrcnKJ3FbjyPAGogmoiZ33c1HG93V
# p6lJ415ERcC7bFQMRbxqrMVANiav1k425zYyFMyLNyE1QulQSgDpW9rtvVcIH7Wv
# G9sqYup9j8z9J1XqbBZPJ5XLln8mS8wWmdDLnBHXgYly/p1DhoQo5fkCAwEAAaOC
# AYswggGHMA4GA1UdDwEB/wQEAwIHgDAMBgNVHRMBAf8EAjAAMBYGA1UdJQEB/wQM
# MAoGCCsGAQUFBwMIMCAGA1UdIAQZMBcwCAYGZ4EMAQQCMAsGCWCGSAGG/WwHATAf
# BgNVHSMEGDAWgBS6FtltTYUvcyl2mi91jGogj57IbzAdBgNVHQ4EFgQUn1csA3cO
# KBWQZqVjXu5Pkh92oFswWgYDVR0fBFMwUTBPoE2gS4ZJaHR0cDovL2NybDMuZGln
# aWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0UlNBNDA5NlNIQTI1NlRpbWVTdGFt
# cGluZ0NBLmNybDCBkAYIKwYBBQUHAQEEgYMwgYAwJAYIKwYBBQUHMAGGGGh0dHA6
# Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBYBggrBgEFBQcwAoZMaHR0cDovL2NhY2VydHMu
# ZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0UlNBNDA5NlNIQTI1NlRpbWVT
# dGFtcGluZ0NBLmNydDANBgkqhkiG9w0BAQsFAAOCAgEAPa0eH3aZW+M4hBJH2UOR
# 9hHbm04IHdEoT8/T3HuBSyZeq3jSi5GXeWP7xCKhVireKCnCs+8GZl2uVYFvQe+p
# PTScVJeCZSsMo1JCoZN2mMew/L4tpqVNbSpWO9QGFwfMEy60HofN6V51sMLMXNTL
# fhVqs+e8haupWiArSozyAmGH/6oMQAh078qRh6wvJNU6gnh5OruCP1QUAvVSu4kq
# VOcJVozZR5RRb/zPd++PGE3qF1P3xWvYViUJLsxtvge/mzA75oBfFZSbdakHJe2B
# VDGIGVNVjOp8sNt70+kEoMF+T6tptMUNlehSR7vM+C13v9+9ZOUKzfRUAYSyyEmY
# tsnpltD/GWX8eM70ls1V6QG/ZOB6b6Yum1HvIiulqJ1Elesj5TMHq8CWT/xrW7tw
# ipXTJ5/i5pkU5E16RSBAdOp12aw8IQhhA/vEbFkEiF2abhuFixUDobZaA0VhqAsM
# HOmaT3XThZDNi5U2zHKhUs5uHHdG6BoQau75KiNbh0c+hatSF+02kULkftARjsyE
# pHKsF7u5zKRbt5oK5YGwFvgc4pEVUNytmB3BpIiowOIIuDgP5M9WArHYSAR16gc0
# dP2XdkMEP5eBsX7bf/MGN4K3HP50v/01ZHo/Z5lGLvNwQ7XHBx1yomzLP8lx4Q1z
# ZKDyHcp4VQJLu2kWTsKsOqQxggUKMIIFBgIBATA0MCAxHjAcBgNVBAMMFVBoaW5J
# VC1QU3NjcmlwdHNfU2lnbgIQd487Ml/QoIxIvrAQtqwTEzANBglghkgBZQMEAgEF
# AKCBhDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgor
# BgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3
# DQEJBDEiBCAiCuKpRvrFkSJkmWobXCUwAC+v2fHwVaee7uLF2QNSLDANBgkqhkiG
# 9w0BAQEFAASCAQCPC53N5NgWOdhlegX2lxIKoHvxE87lPVTXX77NMzuc2EPVRH3w
# qwGwucB9RTb/yA4Koqj2mQ1mP72JUXoMYm5obypQmSX/8EPf5nQTPuEAcmzD/AZm
# bCLCWeTOSlmk3icboEPsHsmoAN+bqsUo9ct73rOhKBd61aowcCRcY0uOFCd9K9mR
# ZWvLI0bac3L+gPnVGCiBsvP2Sdx8yH4/e8CyEif0G/5bxj21coYYwh+ITQHevc9S
# msLEOip1+wp8ViQD84vj0d6fpeGo/EDchcqDB9m1Fz+pg1SieD1u+/JiB3VqHq+m
# aaKxx8SThNoSAILv0aPL16jVcUeVAaUCK45/oYIDIDCCAxwGCSqGSIb3DQEJBjGC
# Aw0wggMJAgEBMHcwYzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJ
# bmMuMTswOQYDVQQDEzJEaWdpQ2VydCBUcnVzdGVkIEc0IFJTQTQwOTYgU0hBMjU2
# IFRpbWVTdGFtcGluZyBDQQIQC65mvFq6f5WHxvnpBOMzBDANBglghkgBZQMEAgEF
# AKBpMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTI1
# MDcwNTEwMTIwOVowLwYJKoZIhvcNAQkEMSIEIHn8wAk2gdTPHFiJiKw9TtqEG6Gq
# cbNUJbRK6BTsFm0LMA0GCSqGSIb3DQEBAQUABIICAJslpjlZXmmZy/84gVf8nNng
# /4hfbOWXb0MOwSDa4RSMqfRePhke4cHQUF89eKjF6U5XdFxJ27nBoPvlPPLi6yVk
# +OHlMVvrdJJhiHxLcnzdGg3KJIG7+UUZnT4S/iw7sGitrafXr5naPgj36DUIVQR0
# O2kgBLGrYMZaGrs+SLWIMGcmHa+pfmfSWioK4FsDPyK+cweNg9CQMUI8Li/VOY3n
# uIM6emDhpRPJSmkQVAH1ECzf9qDVilNqH78hfv2JiZ/RGDsKjo81r735C7uNyfHQ
# K/iYcB1AiGwTwRRy2eHMyX+CBH0pBVKTbthAM4ilErO5sBXRTRaSR+8z+TQ3PK9x
# U2RoLOz7YaNDiitQafssfQfZthRLXDROdMz/ntOqG1u94XJ2Nn+ysti+BRz+G5O/
# 1RUT3lgU5h/hU9D6LWgyyp4iqlDVDkW5/NybiRpUa38gXXpoltpu/FtFNtCSBsLS
# aLjKi+JBDWJM7lh94vrNGKTpAaYqa7bcEk2JIhmaW/SgiNvTewMQ1rKEVc8iShsS
# CiRzRtyyx+bxOIHSg+YdByu1X+Uv0HTL+vtKiLvcYe8aOpGmOErWAoy6bERlj4xO
# 1n3MfWlKFGpRMcDbmBxw5arm3Y6hAtjH2Lzke+2juqnGRR5WTeB6wZ4ib4odI03W
# hS1jUj2fFcAsKVdK0AJX
# SIG # End signature block
