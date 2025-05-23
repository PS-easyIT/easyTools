# easyWSAudit - Windows Server Audit Tool
# Version: 0.0.1

# Debug-Modus - Auf $true setzen für ausführliche Logging-Informationen
$DEBUG = $true
$DebugLogPath = "$env:TEMP\easyWSAudit_Debug.log"

# Debug-Funktion
function Write-DebugLog {
    param(
        [string]$Message,
        [string]$Source = "General"
    )
    
    if ($DEBUG) {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
        $logMessage = "[$timestamp] [$Source] $Message"
        
        # Ausgabe in die Konsole
        Write-Host $logMessage -ForegroundColor Yellow
        
        # Ausgabe in die Log-Datei
        Add-Content -Path $DebugLogPath -Value $logMessage
        
        # Wenn das Debug-Display vorhanden ist, dort auch anzeigen
        if ($script:txtDebugOutput -and $window) {
            $dispatcher = [System.Windows.Threading.Dispatcher]::CurrentDispatcher
            $dispatcher.Invoke([Action]{
                $script:txtDebugOutput.Text += "$logMessage`r`n"
                $script:txtDebugOutput.ScrollToEnd()
            }, "Normal")
        }
    }
}

# Starte das Debug-Log
if ($DEBUG) {
    $startMessage = "=== easyWSAudit Debug Log gestartet $(Get-Date -Format "yyyy-MM-dd HH:mm:ss") ==="
    Set-Content -Path $DebugLogPath -Value $startMessage -Force
    Write-Host $startMessage -ForegroundColor Cyan
}

# Importiere notwendige Module
Write-DebugLog "Importiere notwendige Module..." "Init"
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms # Für den SaveFileDialog
Write-DebugLog "Module importiert" "Init"

# Definiere die Commands für verschiedene Server-Rollen und Systeminformationen
$commands = @(
    @{Name="Installierte Features und Rollen"; Command="Get-WindowsFeature | Where-Object { `$_.Installed -eq `$true }"; Type="PowerShell"},
    @{Name="Systeminformationen"; Command="systeminfo"; Type="CMD"},
    @{Name="Event-Log Übersicht"; Command="Get-EventLog -LogName System -Newest 100"; Type="PowerShell"},
    @{Name="Netzwerkkonfiguration"; Command="Get-NetIPConfiguration"; Type="PowerShell"},
    @{Name="Aktive Netzwerkverbindungen"; Command="netstat -an"; Type="CMD"},
    @{Name="CPU-Informationen"; Command="Get-CimInstance -ClassName Win32_Processor"; Type="PowerShell"},
    @{Name="Volume-Informationen"; Command="Get-Volume"; Type="PowerShell"},
    @{Name="Active Directory Domain Services"; Command="Get-WindowsFeature AD-Domain-Services"; Type="PowerShell"; FeatureName="AD-Domain-Services"},
    @{Name="Active Directory Certificate Services"; Command="Get-WindowsFeature ADCS-Cert-Authority"; Type="PowerShell"; FeatureName="ADCS-Cert-Authority"},
    @{Name="Active Directory Federation Services"; Command="Get-WindowsFeature ADFS-Federation"; Type="PowerShell"; FeatureName="ADFS-Federation"},
    @{Name="Active Directory Lightweight Directory Services"; Command="Get-WindowsFeature ADLDS"; Type="PowerShell"; FeatureName="ADLDS"},
    @{Name="DHCP Server"; Command="Get-WindowsFeature DHCP"; Type="PowerShell"; FeatureName="DHCP"},
    @{Name="DNS Server"; Command="Get-WindowsFeature DNS"; Type="PowerShell"; FeatureName="DNS"},
    @{Name="File and Storage Services"; Command="Get-WindowsFeature FS-FileServer"; Type="PowerShell"; FeatureName="FS-FileServer"},
    @{Name="Hyper-V"; Command="Get-WindowsFeature Hyper-V"; Type="PowerShell"; FeatureName="Hyper-V"},
    @{Name="Print Services"; Command="Get-WindowsFeature Print-Services"; Type="PowerShell"; FeatureName="Print-Services"},
    @{Name="Remote Desktop Services"; Command="Get-WindowsFeature RDS-RD-Server"; Type="PowerShell"; FeatureName="RDS-RD-Server"},
    @{Name="Terminal Services - aktive Sitzungen"; Command="qwinsta"; Type="CMD"},
    @{Name="Windows Deployment Services"; Command="Get-WindowsFeature WDS"; Type="PowerShell"; FeatureName="WDS"},
    @{Name="Windows Server Update Services"; Command="Get-WindowsFeature UpdateServices"; Type="PowerShell"; FeatureName="UpdateServices"},
    @{Name="Web Server (IIS)"; Command="Get-WindowsFeature Web-Server"; Type="PowerShell"; FeatureName="Web-Server"},
    @{Name="Network Policy and Access Services"; Command="Get-WindowsFeature NPAS"; Type="PowerShell"; FeatureName="NPAS"},
    @{Name="Failover Cluster Services"; Command="Get-Cluster"; Type="PowerShell"; FeatureName="FailoverClusters"}
)

# Funktion zum Ausführen von PowerShell-Befehlen
function Invoke-PSCommand {
    param(
        [string]$Command
    )
    try {
        Write-DebugLog "Ausführen von PowerShell-Befehl: $Command" "CommandExec"
        $result = Invoke-Expression -Command $Command | Out-String
        Write-DebugLog "PowerShell-Befehl erfolgreich ausgeführt. Ergebnis-Länge: $($result.Length)" "CommandExec"
        return $result
    }
    catch {
        $errorMsg = "Fehler bei der Ausführung des Befehls: $Command`r`n$($_.Exception.Message)"
        Write-DebugLog "FEHLER: $errorMsg" "CommandExec"
        return $errorMsg
    }
}

# Funktion zum Ausführen von CMD-Befehlen
function Invoke-CMDCommand {
    param(
        [string]$Command
    )
    try {
        Write-DebugLog "Ausführen von CMD-Befehl: $Command" "CommandExec"
        $result = cmd /c $Command 2>&1 | Out-String
        Write-DebugLog "CMD-Befehl erfolgreich ausgeführt. Ergebnis-Länge: $($result.Length)" "CommandExec"
        return $result
    }
    catch {
        $errorMsg = "Fehler bei der Ausführung des Befehls: $Command`r`n$($_.Exception.Message)"
        Write-DebugLog "FEHLER: $errorMsg" "CommandExec"
        return $errorMsg
    }
}

# Funktion zum Prüfen, ob eine bestimmte Serverrolle installiert ist
function Test-ServerRole {
    param(
        [string]$FeatureName
    )
    
    try {
        Write-DebugLog "Prüfe Serverrolle: $FeatureName" "RoleCheck"
        $feature = Get-WindowsFeature -Name $FeatureName -ErrorAction SilentlyContinue
        if ($feature -and $feature.Installed) {
            Write-DebugLog "Serverrolle $FeatureName ist installiert" "RoleCheck"
            return $true
        }
        Write-DebugLog "Serverrolle $FeatureName ist NICHT installiert" "RoleCheck"
        return $false
    }
    catch {
        Write-DebugLog "FEHLER beim Prüfen der Serverrolle $FeatureName - $($_.Exception.Message)" "RoleCheck"
        return $false
    }
}

# Funktion zum Generieren des HTML-Exports
function Export-AuditToHTML {
    param(
        [hashtable]$Results,
        [string]$FilePath
    )
    
    Write-DebugLog "Starte HTML-Export nach: $FilePath" "Export"
    $htmlHeader = @"
<!DOCTYPE html>
<html>
<head>
    <title>Windows Server Audit Report</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        h1 { color: #0066cc; }
        h2 { color: #003366; margin-top: 20px; border-bottom: 1px solid #ccc; padding-bottom: 5px; }
        pre { background-color: #f5f5f5; padding: 10px; border: 1px solid #ddd; white-space: pre-wrap; }
        .section { margin-bottom: 30px; }
        .timestamp { color: #666; font-style: italic; margin-bottom: 20px; }
    </style>
</head>
<body>
    <h1>Windows Server Audit Report</h1>
    <div class="timestamp">Report generiert am $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</div>
    <div class="timestamp">Server: $env:COMPUTERNAME</div>
"@

    $htmlFooter = @"
</body>
</html>
"@

    $htmlBody = ""
    
    foreach ($key in $Results.Keys) {
        $htmlBody += @"
    <div class="section">
        <h2>$key</h2>
        <pre>$($Results[$key])</pre>
    </div>
"@
    }
    
    $htmlContent = $htmlHeader + $htmlBody + $htmlFooter
    $htmlContent | Out-File -FilePath $FilePath -Encoding utf8
}

# Variable für die Audit-Ergebnisse
$global:auditResults = @{}

# XAML UI Definition - Füge Debug-Anzeige hinzu
[xml]$xaml = @"
<Window 
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="easyWSAudit - Windows Server Audit Tool"
    Height="800" Width="1200" WindowStartupLocation="CenterScreen"
    Background="#F9F9F9">
    <Window.Resources>
        <Style x:Key="ModernButton" TargetType="Button">
            <Setter Property="Background" Value="#0078D4"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="Padding" Value="15,8"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Margin" Value="5"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" 
                                CornerRadius="4" 
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#106EBE"/>
                </Trigger>
                <Trigger Property="IsEnabled" Value="False">
                    <Setter Property="Background" Value="#CCCCCC"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style x:Key="NavButton" TargetType="Button">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="#333333"/>
            <Setter Property="Padding" Value="15,10"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="HorizontalContentAlignment" Value="Left"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" Background="{TemplateBinding Background}" 
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Left" VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#E8E8E8"/>
                </Trigger>
                <Trigger Property="IsEnabled" Value="False">
                    <Setter Property="Foreground" Value="#999999"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style x:Key="CheckboxStyle" TargetType="CheckBox">
            <Setter Property="Margin" Value="5,3"/>
            <Setter Property="Padding" Value="5,0,0,0"/>
        </Style>
        <Style TargetType="ScrollViewer">
            <Setter Property="Padding" Value="5"/>
        </Style>
    </Window.Resources>
    <Grid>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="250"/>
            <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>

        <!-- Linke Navigation -->
        <Border Grid.Column="0" Background="White" BorderThickness="0,0,1,0" BorderBrush="#E0E0E0">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="80"/>
                    <RowDefinition Height="*"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                
                <!-- Header -->
                <Border Grid.Row="0" Background="#0078D4" Padding="15">
                    <TextBlock Text="easyWSAudit" FontSize="22" Foreground="White" VerticalAlignment="Center"/>
                </Border>
                
                <!-- Navigation Buttons -->
                <StackPanel Grid.Row="1" Margin="0,10,0,0">
                    <Button x:Name="btnOverview" Content="Übersicht" Style="{StaticResource NavButton}"/>
                    <Button x:Name="btnAudit" Content="Audit durchführen" Style="{StaticResource NavButton}"/>
                    <Button x:Name="btnResults" Content="Audit Ergebnisse" Style="{StaticResource NavButton}"/>
                    <Button x:Name="btnExport" Content="Export" Style="{StaticResource NavButton}"/>
                    <Button x:Name="btnDebug" Content="Debug" Style="{StaticResource NavButton}" Visibility="Collapsed"/>
                </StackPanel>
                
                <!-- Status -->
                <Border Grid.Row="2" Background="#F0F0F0" Padding="15,10">
                    <TextBlock x:Name="txtStatus" Text="Status: Bereit" FontSize="12"/>
                </Border>
            </Grid>
        </Border>
        
        <!-- Rechter Bereich - Hauptinhalt -->
        <Grid Grid.Column="1" Grid.Row="0" Margin="0">
            <!-- Seiten als Grid-Elemente -->
            
            <!-- Übersichts-Seite -->
            <Grid x:Name="pageOverview" Visibility="Visible" Margin="20">
                <StackPanel>
                    <TextBlock Text="Windows Server Audit Tool" FontSize="24" FontWeight="SemiBold" Margin="0,0,0,15"/>
                    <TextBlock TextWrapping="Wrap" Margin="0,0,0,10">
                        Willkommen beim easyWSAudit Tool. Mit diesem Tool können Sie umfassende Systeminformationen sammeln 
                        und als HTML-Bericht exportieren.
                    </TextBlock>
                    <TextBlock TextWrapping="Wrap" Margin="0,0,0,15">
                        Das Tool erkennt automatisch installierte Serverrollen und passt die verfügbaren Optionen entsprechend an.
                    </TextBlock>
                    <Button Content="Audit starten" Style="{StaticResource ModernButton}" Width="150" HorizontalAlignment="Left" x:Name="btnStartAudit"/>
                </StackPanel>
            </Grid>
            
            <!-- Audit-Seite -->
            <Grid x:Name="pageAudit" Visibility="Collapsed" Margin="20">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                
                <TextBlock Grid.Row="0" Text="Audit-Optionen" FontSize="20" FontWeight="SemiBold" Margin="0,0,0,15"/>
                
                <Border Grid.Row="1" Background="White" CornerRadius="4" BorderThickness="1" BorderBrush="#E0E0E0">
                    <ScrollViewer VerticalScrollBarVisibility="Auto">
                        <StackPanel x:Name="spOptions" Margin="10"/>
                    </ScrollViewer>
                </Border>
                
                <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,15,0,0" HorizontalAlignment="Left">
                    <Button Content="Alle auswählen" x:Name="btnSelectAll" Style="{StaticResource ModernButton}" Margin="0,0,10,0"/>
                    <Button Content="Keinen auswählen" x:Name="btnSelectNone" Style="{StaticResource ModernButton}" Margin="0,0,10,0"/>
                    <Button Content="Audit ausführen" x:Name="btnRunAudit" Style="{StaticResource ModernButton}"/>
                </StackPanel>
            </Grid>
            
            <!-- Ergebnis-Seite -->
            <Grid x:Name="pageResults" Visibility="Collapsed" Margin="20">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                
                <TextBlock Grid.Row="0" Text="Audit-Ergebnisse" FontSize="20" FontWeight="SemiBold" Margin="0,0,0,15"/>
                
                <Border Grid.Row="1" Background="White" CornerRadius="4" BorderThickness="1" BorderBrush="#E0E0E0">
                    <ScrollViewer VerticalScrollBarVisibility="Auto">
                        <TextBox x:Name="txtResults" Background="Transparent" BorderThickness="0" 
                                 FontFamily="Consolas" IsReadOnly="True" TextWrapping="Wrap"/>
                    </ScrollViewer>
                </Border>
                
                <Button Grid.Row="2" Content="Als HTML exportieren" x:Name="btnExportHTML" 
                        Style="{StaticResource ModernButton}" Margin="0,15,0,0" HorizontalAlignment="Left" 
                        IsEnabled="False"/>
            </Grid>
            
            <!-- Export-Seite -->
            <Grid x:Name="pageExport" Visibility="Collapsed" Margin="20">
                <StackPanel>
                    <TextBlock Text="Export-Optionen" FontSize="20" FontWeight="SemiBold" Margin="0,0,0,15"/>
                    <TextBlock TextWrapping="Wrap" Margin="0,0,0,15">
                        Hier können Sie den erstellten Audit-Bericht als HTML-Datei exportieren.
                    </TextBlock>
                    <Button Content="Als HTML exportieren" Style="{StaticResource ModernButton}" Width="200" 
                            HorizontalAlignment="Left" x:Name="btnExportPage" IsEnabled="False"/>
                    <TextBlock x:Name="txtExportStatus" Margin="0,15,0,0"/>
                </StackPanel>
            </Grid>
            
            <!-- Debug-Seite -->
            <Grid x:Name="pageDebug" Visibility="Collapsed" Margin="20">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                
                <TextBlock Grid.Row="0" Text="Debug-Informationen" FontSize="20" FontWeight="SemiBold" Margin="0,0,0,15"/>
                
                <Border Grid.Row="1" Background="#1E1E1E" CornerRadius="4" BorderThickness="1" BorderBrush="#333333">
                    <ScrollViewer VerticalScrollBarVisibility="Auto">
                        <TextBox x:Name="txtDebugOutput" Background="Transparent" BorderThickness="0" 
                                 FontFamily="Consolas" IsReadOnly="True" TextWrapping="Wrap"
                                 Foreground="#CCCCCC" Padding="10"/>
                    </ScrollViewer>
                </Border>
                
                <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,15,0,0" HorizontalAlignment="Left">
                    <Button Content="Log-Datei öffnen" x:Name="btnOpenLog" Style="{StaticResource ModernButton}" Margin="0,0,10,0"/>
                    <Button Content="Log leeren" x:Name="btnClearLog" Style="{StaticResource ModernButton}"/>
                </StackPanel>
            </Grid>
        </Grid>
    </Grid>
</Window>
"@

# Lade das XAML
Write-DebugLog "Lade XAML für UI..." "UI"
$reader = [System.Xml.XmlNodeReader]::new($xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)
Write-DebugLog "XAML geladen, Fenster erstellt" "UI"

# Hole die UI-Elemente
$btnOverview = $window.FindName("btnOverview")
$btnAudit = $window.FindName("btnAudit")
$btnResults = $window.FindName("btnResults")
$btnExport = $window.FindName("btnExport")
$btnDebug = $window.FindName("btnDebug")
$txtStatus = $window.FindName("txtStatus")

$pageOverview = $window.FindName("pageOverview")
$pageAudit = $window.FindName("pageAudit")
$pageResults = $window.FindName("pageResults")
$pageExport = $window.FindName("pageExport")
$pageDebug = $window.FindName("pageDebug")

$btnStartAudit = $window.FindName("btnStartAudit")
$spOptions = $window.FindName("spOptions")
$btnSelectAll = $window.FindName("btnSelectAll")
$btnSelectNone = $window.FindName("btnSelectNone")
$btnRunAudit = $window.FindName("btnRunAudit")
$txtResults = $window.FindName("txtResults")
$btnExportHTML = $window.FindName("btnExportHTML")
$btnExportPage = $window.FindName("btnExportPage")
$txtExportStatus = $window.FindName("txtExportStatus")

# Debug-Elemente
$script:txtDebugOutput = $window.FindName("txtDebugOutput")
$btnOpenLog = $window.FindName("btnOpenLog")
$btnClearLog = $window.FindName("btnClearLog")

# Wenn Debug aktiv, Debug-Button anzeigen
if ($DEBUG) {
    $btnDebug.Visibility = "Visible"
}

Write-DebugLog "UI-Elemente initialisiert" "UI"

# Funktion zum Anzeigen der gewählten Seite
function Show-Page {
    param (
        [string]$PageName
    )
    
    Write-DebugLog "Wechsle zur Seite: $PageName" "Navigation"
    
    $pageOverview.Visibility = "Collapsed"
    $pageAudit.Visibility = "Collapsed"
    $pageResults.Visibility = "Collapsed"
    $pageExport.Visibility = "Collapsed"
    $pageDebug.Visibility = "Collapsed"
    
    switch ($PageName) {
        "Overview" { $pageOverview.Visibility = "Visible" }
        "Audit" { $pageAudit.Visibility = "Visible" }
        "Results" { $pageResults.Visibility = "Visible" }
        "Export" { $pageExport.Visibility = "Visible" }
        "Debug" { $pageDebug.Visibility = "Visible" }
    }
}

# Navigation-Button-Events
$btnOverview.Add_Click({ 
    Write-DebugLog "btnOverview geklickt" "UI"
    Show-Page -PageName "Overview" 
})
$btnAudit.Add_Click({ 
    Write-DebugLog "btnAudit geklickt" "UI"
    Show-Page -PageName "Audit" 
})
$btnResults.Add_Click({ 
    Write-DebugLog "btnResults geklickt" "UI"
    Show-Page -PageName "Results" 
})
$btnExport.Add_Click({ 
    Write-DebugLog "btnExport geklickt" "UI"
    Show-Page -PageName "Export" 
})
$btnDebug.Add_Click({ 
    Write-DebugLog "btnDebug geklickt" "UI"
    Show-Page -PageName "Debug" 
})

# Von der Startseite zum Audit wechseln
$btnStartAudit.Add_Click({ 
    Write-DebugLog "btnStartAudit geklickt" "UI"
    Show-Page -PageName "Audit" 
})

# Dictionary für Checkboxen
$checkboxes = @{}

# Erstelle die Checkboxen für die Audit-Optionen
Write-DebugLog "Erstelle Checkboxen für Audit-Optionen..." "UI"
foreach ($cmd in $commands) {
    Write-DebugLog "  Erstelle Checkbox für: $($cmd.Name)" "UI"
    $checkbox = New-Object System.Windows.Controls.CheckBox
    $checkbox.Content = $cmd.Name
    $checkbox.IsChecked = $true
    $checkbox.Style = $window.FindResource("CheckboxStyle")
    
    # Überprüfe, ob diese Option mit einer Serverrolle verbunden ist
    if ($cmd.ContainsKey("FeatureName")) {
        Write-DebugLog "  Option ist mit Serverrolle verbunden: $($cmd.FeatureName)" "UI"
        $isRoleInstalled = Test-ServerRole -FeatureName $cmd.FeatureName
        if (-not $isRoleInstalled) {
            Write-DebugLog "  Rolle nicht installiert, deaktiviere Checkbox" "UI"
            $checkbox.IsEnabled = $false
            $checkbox.Content = "$($cmd.Name) (Nicht installiert)"
            $checkbox.IsChecked = $false
        }
    }
    
    $spOptions.Children.Add($checkbox)
    $checkboxes[$cmd.Name] = $checkbox
}
Write-DebugLog "Checkboxen erstellt" "UI"

# "Alle auswählen" Button-Funktion
$btnSelectAll.Add_Click({
    Write-DebugLog "btnSelectAll geklickt - Wähle alle Checkboxen" "UI"
    foreach ($key in $checkboxes.Keys) {
        if ($checkboxes[$key].IsEnabled) {
            $checkboxes[$key].IsChecked = $true
        }
    }
})

# "Keinen auswählen" Button-Funktion
$btnSelectNone.Add_Click({
    Write-DebugLog "btnSelectNone geklickt - Deselektiere alle Checkboxen" "UI"
    foreach ($key in $checkboxes.Keys) {
        $checkboxes[$key].IsChecked = $false
    }
})

# "Audit ausführen" Button-Funktion
$btnRunAudit.Add_Click({
    Write-DebugLog "btnRunAudit geklickt - Starte Audit" "Audit"
    $txtResults.Text = "Audit wird ausgeführt, bitte warten..."
    $txtStatus.Text = "Status: Audit läuft..."
    
    # Sammle ausgewählte Befehle
    $selectedCommands = @()
    foreach ($cmd in $commands) {
        if ($checkboxes[$cmd.Name].IsChecked) {
            Write-DebugLog "Ausgewählter Befehl: $($cmd.Name)" "Audit"
            $selectedCommands += $cmd
        }
    }
    
    Write-DebugLog "Ausgewählte Befehle: $($selectedCommands.Count)" "Audit"
    
    # Starte einen Hintergrund-Job für die Audit-Ausführung
    $job = [PowerShell]::Create().AddScript({
        param($selectedCommands, $debugMode, $debugLogPath)
        
        # Debug-Funktion im Job-Kontext
        function Write-JobDebugLog {
            param([string]$Message, [string]$Source = "JobExec")
            
            if ($debugMode) {
                $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
                $logMessage = "[$timestamp] [JOB-$Source] $Message"
                
                # Ausgabe in die Log-Datei
                Add-Content -Path $debugLogPath -Value $logMessage
            }
        }
        
        Write-JobDebugLog "Job gestartet mit $($selectedCommands.Count) Befehlen" "Init"
        
        # Definiere die benötigten Funktionen im Job-Kontext
        function Invoke-PSCommand {
            param([string]$Command)
            try {
                Write-JobDebugLog "Führe PowerShell-Befehl aus: $Command" "CmdExec"
                $result = Invoke-Expression -Command $Command | Out-String
                Write-JobDebugLog "PowerShell-Befehl erfolgreich ausgeführt. Ergebnis-Länge: $($result.Length)" "CmdExec"
                return $result
            } catch {
                $errorMsg = "Fehler bei der Ausführung des Befehls: $Command`r`n$($_.Exception.Message)"
                Write-JobDebugLog "FEHLER: $errorMsg" "CmdExec"
                return $errorMsg
            }
        }

        function Invoke-CMDCommand {
            param([string]$Command)
            try {
                Write-JobDebugLog "Führe CMD-Befehl aus: $Command" "CmdExec"
                $result = cmd /c $Command 2>&1 | Out-String
                Write-JobDebugLog "CMD-Befehl erfolgreich ausgeführt. Ergebnis-Länge: $($result.Length)" "CmdExec"
                return $result
            } catch {
                $errorMsg = "Fehler bei der Ausführung des Befehls: $Command`r`n$($_.Exception.Message)"
                Write-JobDebugLog "FEHLER: $errorMsg" "CmdExec"
                return $errorMsg
            }
        }
        
        $results = @{}
        $resultText = "Audit wird ausgeführt, bitte warten...`r`n`r`n"
        
        for ($i=0; $i -lt $selectedCommands.Count; $i++) {
            $cmd = $selectedCommands[$i]
            Write-JobDebugLog "Verarbeite Befehl ($($i+1)/$($selectedCommands.Count)): $($cmd.Name)" "Progress"
            $resultText += "Führe aus: $($cmd.Name)...`r`n"
            
            if ($cmd.Type -eq "PowerShell") {
                $result = Invoke-PSCommand -Command $cmd.Command
            } else {
                $result = Invoke-CMDCommand -Command $cmd.Command
            }
            
            $results[$cmd.Name] = $result
            $resultText += "`r`n==== $($cmd.Name) ====`r`n$result`r`n`r`n"
            Write-JobDebugLog "Befehl abgeschlossen: $($cmd.Name)" "Progress"
        }
        
        $resultText += "Audit abgeschlossen."
        Write-JobDebugLog "Job abgeschlossen, $($results.Count) Ergebnisse gesammelt" "Complete"
        
        return @{
            Results = $results
            ResultText = $resultText
        }
    }).AddArgument($selectedCommands).AddArgument($DEBUG).AddArgument($DebugLogPath)
    
    Write-DebugLog "PowerShell-Job erstellt, beginne Ausführung..." "Audit"
    $asyncResult = $job.BeginInvoke()
    Write-DebugLog "Job asynchron gestartet mit AsyncResult: $($asyncResult.GetType().Name)" "Audit"
    
    $timer = New-Object System.Windows.Threading.DispatcherTimer
    $timer.Interval = [TimeSpan]::FromMilliseconds(500)
    Write-DebugLog "Timer erstellt, intervall: 500ms" "Audit"
    
    $timer.Add_Tick({
        try {
            # Prüfen, ob der Job noch läuft oder abgeschlossen ist
            if ($asyncResult.IsCompleted) {
                Write-DebugLog "Job ist abgeschlossen, verarbeite Ergebnisse..." "AuditTimer"
                $timer.Stop()
                Write-DebugLog "Timer gestoppt" "AuditTimer"
                $txtStatus.Text = "Status: Verarbeite Ergebnisse..."
                
                # Ergebnisse abrufen
                try {
                    Write-DebugLog "Rufe Ergebnisse vom Job ab..." "AuditTimer"
                    $result = $job.EndInvoke($asyncResult)
                    Write-DebugLog "Ergebnisse abgerufen" "AuditTimer"
                    
                    # Prüfen, ob Ergebnisse zurückgegeben wurden
                    if ($null -eq $result) {
                        Write-DebugLog "FEHLER: Keine Ergebnisse vom Audit erhalten" "AuditTimer"
                        $txtResults.Text = "Fehler: Keine Ergebnisse vom Audit erhalten."
                        $txtStatus.Text = "Status: Fehler - keine Ergebnisse"
                    } else {
                        # Ergebnisse speichern und anzeigen
                        Write-DebugLog "Ergebnisse erhalten: $($result.Results.Count) Einträge" "AuditTimer"
                        $global:auditResults = $result.Results
                        $txtResults.Text = $result.ResultText
                        $txtStatus.Text = "Status: Audit erfolgreich abgeschlossen"
                        
                        # Aktiviere die Export-Buttons
                        Write-DebugLog "Aktiviere Export-Buttons" "AuditTimer"
                        $btnExportHTML.IsEnabled = $true
                        $btnExportPage.IsEnabled = $true
                        
                        # Wechsle zur Ergebnisseite
                        Write-DebugLog "Wechsle zur Ergebnisseite" "AuditTimer"
                        Show-Page -PageName "Results"
                    }
                }
                catch {
                    $errorMessage = "Fehler beim Abrufen der Audit-Ergebnisse: $($_.Exception.Message)"
                    Write-DebugLog "FEHLER: $errorMessage" "AuditTimer"
                    Write-Host $errorMessage
                    $txtResults.Text = $errorMessage
                    $txtStatus.Text = "Status: Fehler beim Audit"
                }
                finally {
                    # Bereinige den Job
                    try {
                        Write-DebugLog "Versuche Job zu bereinigen" "AuditTimer"
                        $job.Dispose()
                        Write-DebugLog "Job bereinigt" "AuditTimer"
                    } catch {
                        Write-DebugLog "FEHLER beim Bereinigen des Jobs: $($_.Exception.Message)" "AuditTimer"
                        # Ignoriere Fehler beim Bereinigen
                    }
                }
            } else {
                # Aktualisiere den Status während der Job läuft
                $dots = "." * ((Get-Date).Second % 4 + 1)
                $txtStatus.Text = "Status: Audit läuft$dots"
            }
        } catch {
            $timer.Stop()
            $errorMessage = "Unerwarteter Fehler im Timer-Event: $($_.Exception.Message)"
            Write-DebugLog "KRITISCHER FEHLER: $errorMessage" "AuditTimer"
            $txtResults.Text = $errorMessage
            $txtStatus.Text = "Status: Fehler im Timer"
            try { $job.Dispose() } catch {}
        }
    })
    
    Write-DebugLog "Starte Timer" "Audit"
    $timer.Start()
    
    # Sofort umschalten zur Ergebnisseite, um Fortschritt zu sehen
    Write-DebugLog "Wechsle zur Ergebnisseite, um Fortschritt zu zeigen" "Audit"
    Show-Page -PageName "Results"
})

# Export-Button-Funktionalität
$btnExportHTML.Add_Click({
    Write-DebugLog "btnExportHTML geklickt" "Export"
    Export-AuditResults
})

$btnExportPage.Add_Click({
    Write-DebugLog "btnExportPage geklickt" "Export"
    Export-AuditResults
    $txtExportStatus.Text = "Bericht wurde erfolgreich exportiert."
})

function Export-AuditResults {
    Write-DebugLog "Starte Export-AuditResults Funktion" "Export"
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "HTML Files (*.html)|*.html"
    $saveFileDialog.Title = "Speichern Sie den Audit-Bericht"
    $saveFileDialog.FileName = "ServerAudit_$($env:COMPUTERNAME)_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
    
    Write-DebugLog "Zeige SaveFileDialog" "Export"
    $dialogResult = $saveFileDialog.ShowDialog()
    
    if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
        Write-DebugLog "Datei wird gespeichert unter: $($saveFileDialog.FileName)" "Export"
        $txtStatus.Text = "Status: Exportiere HTML..."
        Export-AuditToHTML -Results $global:auditResults -FilePath $saveFileDialog.FileName
        $txtStatus.Text = "Status: Export abgeschlossen"
        Write-DebugLog "Export abgeschlossen" "Export"
        [System.Windows.MessageBox]::Show("Bericht wurde gespeichert unter: $($saveFileDialog.FileName)", "Export abgeschlossen", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
    } else {
        Write-DebugLog "Export abgebrochen durch Benutzer" "Export"
    }
}

# Debug-Funktionen
$btnOpenLog.Add_Click({
    Write-DebugLog "btnOpenLog geklickt" "Debug"
    if (Test-Path $DebugLogPath) {
        Start-Process notepad.exe -ArgumentList $DebugLogPath
    } else {
        [System.Windows.MessageBox]::Show("Log-Datei nicht gefunden.", "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
})

$btnClearLog.Add_Click({
    Write-DebugLog "btnClearLog geklickt - Lösche Debug-Ausgabe" "Debug"
    $script:txtDebugOutput.Text = ""
    if ($DEBUG) {
        $clearMessage = "=== Debug-Log gelöscht: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss") ==="
        Set-Content -Path $DebugLogPath -Value $clearMessage -Force
    }
})

Write-DebugLog "Skript initialisierung abgeschlossen, zeige UI" "Init"

# Zeige das Fenster an
$null = $window.ShowDialog()
Write-DebugLog "UI geschlossen, Skript beendet" "Shutdown"
