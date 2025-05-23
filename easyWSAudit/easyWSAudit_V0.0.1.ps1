# easyWSAudit - Windows Server Audit Tool
# Version: 0.0.2 - Vereinfacht ohne asynchrone Ausführung

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
            try {
                $script:txtDebugOutput.Dispatcher.Invoke([Action]{
                    $script:txtDebugOutput.Text += "$logMessage`r`n"
                    $script:txtDebugOutput.ScrollToEnd()
                }, "Normal")
            } catch {
                # Ignoriere Dispatcher-Fehler
            }
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
    @{Name="Systeminformationen"; Command="Get-ComputerInfo"; Type="PowerShell"; Category="System"},
    @{Name="Betriebssystem Details"; Command="Get-CimInstance Win32_OperatingSystem"; Type="PowerShell"; Category="System"},
    @{Name="Hardware Informationen"; Command="Get-CimInstance Win32_ComputerSystem"; Type="PowerShell"; Category="Hardware"},
    @{Name="CPU Informationen"; Command="Get-CimInstance Win32_Processor"; Type="PowerShell"; Category="Hardware"},
    @{Name="Arbeitsspeicher Details"; Command="Get-CimInstance Win32_PhysicalMemory"; Type="PowerShell"; Category="Hardware"},
    @{Name="Festplatten Informationen"; Command="Get-CimInstance Win32_LogicalDisk"; Type="PowerShell"; Category="Storage"},
    @{Name="Volume Informationen"; Command="Get-Volume"; Type="PowerShell"; Category="Storage"},
    @{Name="Installierte Features und Rollen"; Command="Get-WindowsFeature | Where-Object { `$_.Installed -eq `$true }"; Type="PowerShell"; Category="Features"},
    @{Name="Installierte Programme"; Command="Get-CimInstance Win32_Product | Select-Object Name, Version, Vendor | Sort-Object Name"; Type="PowerShell"; Category="Software"},
    @{Name="Windows Updates"; Command="Get-HotFix | Sort-Object InstalledOn -Descending | Select-Object -First 20"; Type="PowerShell"; Category="Updates"},
    @{Name="Netzwerkkonfiguration"; Command="Get-NetIPConfiguration"; Type="PowerShell"; Category="Network"},
    @{Name="Netzwerkadapter"; Command="Get-NetAdapter"; Type="PowerShell"; Category="Network"},
    @{Name="Aktive Netzwerkverbindungen"; Command="Get-NetTCPConnection | Where-Object State -eq 'Listen' | Select-Object LocalAddress, LocalPort, OwningProcess"; Type="PowerShell"; Category="Network"},
    @{Name="Firewall Regeln"; Command="Get-NetFirewallRule | Where-Object Enabled -eq 'True' | Select-Object DisplayName, Direction, Action | Sort-Object DisplayName"; Type="PowerShell"; Category="Security"},
    @{Name="Services (Automatisch)"; Command="Get-Service | Where-Object StartType -eq 'Automatic' | Sort-Object Status, Name"; Type="PowerShell"; Category="Services"},
    @{Name="Services (Laufend)"; Command="Get-Service | Where-Object Status -eq 'Running' | Sort-Object Name"; Type="PowerShell"; Category="Services"},
    @{Name="Geplante Aufgaben"; Command="Get-ScheduledTask | Where-Object State -eq 'Ready' | Select-Object TaskName, TaskPath, State"; Type="PowerShell"; Category="Tasks"},
    @{Name="Event-Log System (Letzte 24h)"; Command="Get-WinEvent -FilterHashtable @{LogName='System'; StartTime=(Get-Date).AddDays(-1)} -MaxEvents 50 | Select-Object TimeCreated, Id, LevelDisplayName, Message"; Type="PowerShell"; Category="Events"},
    @{Name="Event-Log Application (Letzte 24h)"; Command="Get-WinEvent -FilterHashtable @{LogName='Application'; StartTime=(Get-Date).AddDays(-1)} -MaxEvents 50 | Select-Object TimeCreated, Id, LevelDisplayName, Message"; Type="PowerShell"; Category="Events"},
    @{Name="Lokale Benutzer"; Command="Get-LocalUser | Select-Object Name, Enabled, LastLogon, PasswordRequired"; Type="PowerShell"; Category="Security"},
    @{Name="Lokale Gruppen"; Command="Get-LocalGroup | Select-Object Name, Description"; Type="PowerShell"; Category="Security"},
    
    # Server-Rollen spezifische Commands
    @{Name="Active Directory Domain Services"; Command="Get-ADDomain; Get-ADForest"; Type="PowerShell"; FeatureName="AD-Domain-Services"; Category="AD"},
    @{Name="DNS Server Zonen"; Command="Get-DnsServerZone"; Type="PowerShell"; FeatureName="DNS"; Category="DNS"},
    @{Name="DHCP Bereiche"; Command="Get-DhcpServerv4Scope"; Type="PowerShell"; FeatureName="DHCP"; Category="DHCP"},
    @{Name="Hyper-V VMs"; Command="Get-VM"; Type="PowerShell"; FeatureName="Hyper-V"; Category="Hyper-V"},
    @{Name="IIS Websites"; Command="Get-Website"; Type="PowerShell"; FeatureName="Web-Server"; Category="IIS"},
    @{Name="Cluster Informationen"; Command="Get-Cluster; Get-ClusterNode"; Type="PowerShell"; FeatureName="Failover-Clustering"; Category="Cluster"}
)

# Funktion zum Ausführen von PowerShell-Befehlen
function Invoke-PSCommand {
    param(
        [string]$Command
    )
    try {
        Write-DebugLog "Ausführen von PowerShell-Befehl: $Command" "CommandExec"
        
        # Spezielle Behandlung für bestimmte Befehle
        if ($Command -like "*Get-ComputerInfo*") {
            $result = Get-ComputerInfo | Format-List | Out-String
        } else {
            $result = Invoke-Expression -Command $Command | Format-Table -AutoSize | Out-String
        }
        
        Write-DebugLog "PowerShell-Befehl erfolgreich ausgeführt. Ergebnis-Länge: $($result.Length)" "CommandExec"
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
    
    $serverInfo = @{
        ServerName = $env:COMPUTERNAME
        ReportDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Domain = $env:USERDOMAIN
        User = $env:USERNAME
    }
    
    $htmlHeader = @"
<!DOCTYPE html>
<html>
<head>
    <title>Windows Server Audit Report - $($serverInfo.ServerName)</title>
    <meta charset="utf-8">
    <style>
        body { 
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; 
            margin: 0; 
            padding: 20px; 
            background-color: #f5f5f5;
        }
        .container {
            max-width: 1200px;
            margin: 0 auto;
            background-color: white;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            overflow: hidden;
        }
        .header {
            background: linear-gradient(135deg, #0078d4, #106ebe);
            color: white;
            padding: 30px;
            text-align: center;
        }
        .header h1 { 
            margin: 0; 
            font-size: 2.5em; 
            font-weight: 300;
        }
        .server-info {
            background-color: #f8f9fa;
            padding: 20px;
            border-bottom: 1px solid #dee2e6;
        }
        .info-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 15px;
        }
        .info-item {
            background: white;
            padding: 15px;
            border-radius: 4px;
            border-left: 4px solid #0078d4;
        }
        .info-label {
            font-weight: 600;
            color: #495057;
            font-size: 0.9em;
        }
        .info-value {
            font-size: 1.1em;
            color: #212529;
            margin-top: 5px;
        }
        .nav-tabs {
            display: flex;
            background-color: #f8f9fa;
            border-bottom: 1px solid #dee2e6;
            overflow-x: auto;
        }
        .nav-tab {
            padding: 15px 25px;
            background: none;
            border: none;
            cursor: pointer;
            border-bottom: 3px solid transparent;
            transition: all 0.3s;
            white-space: nowrap;
        }
        .nav-tab:hover {
            background-color: #e9ecef;
        }
        .nav-tab.active {
            border-bottom-color: #0078d4;
            background-color: white;
            font-weight: 600;
        }
        .tab-content {
            display: none;
            padding: 30px;
        }
        .tab-content.active {
            display: block;
        }
        .section {
            margin-bottom: 40px;
            background: white;
            border-radius: 6px;
            overflow: hidden;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        }
        .section-header {
            background: linear-gradient(90deg, #f8f9fa, #e9ecef);
            padding: 20px;
            border-bottom: 1px solid #dee2e6;
        }
        .section-title {
            font-size: 1.4em;
            font-weight: 600;
            color: #495057;
            margin: 0;
        }
        .section-content {
            padding: 20px;
        }
        pre { 
            background-color: #f8f9fa; 
            padding: 20px; 
            border: 1px solid #dee2e6; 
            border-radius: 4px;
            white-space: pre-wrap; 
            font-family: 'Consolas', 'Monaco', monospace;
            font-size: 0.9em;
            line-height: 1.4;
            overflow-x: auto;
        }
        .timestamp { 
            color: #6c757d; 
            font-style: italic; 
            text-align: center;
            padding: 20px;
            border-top: 1px solid #dee2e6;
            background-color: #f8f9fa;
        }
        .category-summary {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            margin-bottom: 30px;
        }
        .category-card {
            background: linear-gradient(135deg, #ffffff, #f8f9fa);
            border: 1px solid #dee2e6;
            border-radius: 6px;
            padding: 20px;
            text-align: center;
            transition: transform 0.2s;
        }
        .category-card:hover {
            transform: translateY(-2px);
            box-shadow: 0 4px 12px rgba(0,0,0,0.1);
        }
        .category-count {
            font-size: 2em;
            font-weight: bold;
            color: #0078d4;
        }
        .category-label {
            font-size: 0.9em;
            color: #6c757d;
            margin-top: 5px;
        }
    </style>
    <script>
        function showTab(tabName) {
            // Hide all tab contents
            var contents = document.querySelectorAll('.tab-content');
            contents.forEach(function(content) {
                content.classList.remove('active');
            });
            
            // Remove active class from all tabs
            var tabs = document.querySelectorAll('.nav-tab');
            tabs.forEach(function(tab) {
                tab.classList.remove('active');
            });
            
            // Show selected tab content
            document.getElementById(tabName).classList.add('active');
            
            // Add active class to clicked tab
            event.target.classList.add('active');
        }
        
        window.onload = function() {
            // Show first tab by default
            document.querySelector('.nav-tab').click();
        }
    </script>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🖥️ Windows Server Audit Report</h1>
        </div>
        
        <div class="server-info">
            <div class="info-grid">
                <div class="info-item">
                    <div class="info-label">Server Name</div>
                    <div class="info-value">$($serverInfo.ServerName)</div>
                </div>
                <div class="info-item">
                    <div class="info-label">Report Date</div>
                    <div class="info-value">$($serverInfo.ReportDate)</div>
                </div>
                <div class="info-item">
                    <div class="info-label">Domain</div>
                    <div class="info-value">$($serverInfo.Domain)</div>
                </div>
                <div class="info-item">
                    <div class="info-label">Generated by</div>
                    <div class="info-value">$($serverInfo.User)</div>
                </div>
            </div>
        </div>
"@

    # Gruppiere Ergebnisse nach Kategorien
    $categories = @{}
    foreach ($cmd in $commands) {
        $category = if ($cmd.Category) { $cmd.Category } else { "Allgemein" }
        if (-not $categories.ContainsKey($category)) {
            $categories[$category] = @()
        }
        if ($Results.ContainsKey($cmd.Name)) {
            $categories[$category] += @{
                Name = $cmd.Name
                Result = $Results[$cmd.Name]
            }
        }
    }

    # Navigation Tabs
    $navTabs = ""
    $tabContents = ""
    
    foreach ($category in $categories.Keys | Sort-Object) {
        $tabId = $category -replace '[^a-zA-Z0-9]', ''
        $navTabs += "<button class='nav-tab' onclick='showTab(`"$tabId`")'>$category ($($categories[$category].Count))</button>"
        
        $tabContent = "<div id='$tabId' class='tab-content'>"
        
        foreach ($item in $categories[$category]) {
            $tabContent += @"
            <div class="section">
                <div class="section-header">
                    <h3 class="section-title">$($item.Name)</h3>
                </div>
                <div class="section-content">
                    <pre>$($item.Result)</pre>
                </div>
            </div>
"@
        }
        
        $tabContent += "</div>"
        $tabContents += $tabContent
    }

    $htmlNav = @"
        <div class="nav-tabs">
            $navTabs
        </div>
"@

    $htmlFooter = @"
        $htmlNav
        $tabContents
        <div class="timestamp">
            Report generiert mit easyWSAudit am $($serverInfo.ReportDate)
        </div>
    </div>
</body>
</html>
"@

    $htmlContent = $htmlHeader + $htmlFooter
    $htmlContent | Out-File -FilePath $FilePath -Encoding utf8
    Write-DebugLog "HTML-Export abgeschlossen" "Export"
}

# Variable für die Audit-Ergebnisse
$global:auditResults = @{}

# XAML UI Definition - Vereinfachte Version
[xml]$xaml = @"
<Window 
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="easyWSAudit - Windows Server Audit Tool"
    Height="900" Width="1400" WindowStartupLocation="CenterScreen"
    Background="#F5F5F5">
    <Window.Resources>
        <Style x:Key="ModernButton" TargetType="Button">
            <Setter Property="Background" Value="#0078D4"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="Padding" Value="15,10"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Margin" Value="5"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
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
        
        <Style x:Key="CategoryHeader" TargetType="TextBlock">
            <Setter Property="FontSize" Value="16"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="Foreground" Value="#0078D4"/>
            <Setter Property="Margin" Value="0,15,0,5"/>
        </Style>
        
        <Style x:Key="CheckboxStyle" TargetType="CheckBox">
            <Setter Property="Margin" Value="20,3,5,3"/>
            <Setter Property="Padding" Value="5,0,0,0"/>
        </Style>
    </Window.Resources>
    
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <!-- Header -->
        <Border Grid.Row="0" Background="#0078D4" Padding="20">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel Grid.Column="0">
                    <TextBlock Text="easyWSAudit" FontSize="28" Foreground="White" FontWeight="Light"/>
                    <TextBlock Text="Windows Server Audit Tool" FontSize="14" Foreground="#CCE7FF" Margin="0,5,0,0"/>
                </StackPanel>
                <StackPanel Grid.Column="1" Orientation="Horizontal">
                    <TextBlock x:Name="txtServerName" Text="" FontSize="14" Foreground="White" VerticalAlignment="Center" Margin="0,0,20,0"/>
                    <Button Content="🔄 Vollständiges Audit" x:Name="btnFullAudit" Style="{StaticResource ModernButton}" Background="#28A745"/>
                </StackPanel>
            </Grid>
        </Border>
        
        <!-- Hauptinhalt -->
        <TabControl Grid.Row="1" Margin="0" Background="Transparent" BorderThickness="0">
            <TabItem Header="📋 Audit Konfiguration" FontSize="14">
                <Grid Margin="20">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="400"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    
                    <!-- Linke Seite - Optionen -->
                    <Border Grid.Column="0" Background="White" CornerRadius="8" Padding="20" Margin="0,0,10,0">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>
                            
                            <TextBlock Grid.Row="0" Text="Audit-Kategorien auswählen" FontSize="18" FontWeight="SemiBold" Margin="0,0,0,15"/>
                            
                            <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" Margin="0,0,0,15">
                                <StackPanel x:Name="spOptions"/>
                            </ScrollViewer>
                            
                            <StackPanel Grid.Row="2">
                                <Button Content="✅ Alle auswählen" x:Name="btnSelectAll" Style="{StaticResource ModernButton}" Background="#28A745" Margin="0,0,0,5"/>
                                <Button Content="❌ Alle abwählen" x:Name="btnSelectNone" Style="{StaticResource ModernButton}" Background="#DC3545" Margin="0,0,0,5"/>
                                <Button Content="🚀 Audit starten" x:Name="btnRunAudit" Style="{StaticResource ModernButton}" Background="#FFC107" Foreground="Black" FontWeight="Bold"/>
                            </StackPanel>
                        </Grid>
                    </Border>
                    
                    <!-- Rechte Seite - Fortschritt -->
                    <Border Grid.Column="1" Background="White" CornerRadius="8" Padding="20" Margin="10,0,0,0">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>
                            
                            <TextBlock Grid.Row="0" Text="Audit-Fortschritt" FontSize="18" FontWeight="SemiBold" Margin="0,0,0,15"/>
                            
                            <StackPanel Grid.Row="1" Margin="0,0,0,15">
                                <ProgressBar x:Name="progressBar" Height="20" Margin="0,0,0,10"/>
                                <TextBlock x:Name="txtProgress" Text="Bereit für Audit" HorizontalAlignment="Center" FontSize="12" Foreground="#666"/>
                            </StackPanel>
                            
                            <Border Grid.Row="2" Background="#F8F9FA" CornerRadius="4" Padding="15">
                                <ScrollViewer VerticalScrollBarVisibility="Auto">
                                    <TextBlock x:Name="txtStatusLog" Text="Bereit..." FontFamily="Consolas" FontSize="11" Foreground="#495057"/>
                                </ScrollViewer>
                            </Border>
                        </Grid>
                    </Border>
                </Grid>
            </TabItem>
            
            <TabItem Header="📊 Audit Ergebnisse" FontSize="14">
                <Grid Margin="20">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    
                    <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,15">
                        <Button Content="💾 Als HTML exportieren" x:Name="btnExportHTML" Style="{StaticResource ModernButton}" Background="#17A2B8" IsEnabled="False"/>
                        <Button Content="📋 In Zwischenablage" x:Name="btnCopyToClipboard" Style="{StaticResource ModernButton}" Background="#6C757D" IsEnabled="False" Margin="10,0,0,0"/>
                    </StackPanel>
                    
                    <Border Grid.Row="1" Background="White" CornerRadius="8" BorderThickness="1" BorderBrush="#DEE2E6">
                        <ScrollViewer VerticalScrollBarVisibility="Auto">
                            <TextBox x:Name="txtResults" Background="Transparent" BorderThickness="0" 
                                     FontFamily="Consolas" IsReadOnly="True" TextWrapping="Wrap" Padding="20" FontSize="12"/>
                        </ScrollViewer>
                    </Border>
                </Grid>
            </TabItem>
            
            <TabItem Header="🔧 Debug" FontSize="14">
                <Grid Margin="20">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <TextBlock Grid.Row="0" Text="Debug-Informationen" FontSize="18" FontWeight="SemiBold" Margin="0,0,0,15"/>
                    
                    <Border Grid.Row="1" Background="#1E1E1E" CornerRadius="8" BorderThickness="1" BorderBrush="#333333">
                        <ScrollViewer VerticalScrollBarVisibility="Auto">
                            <TextBox x:Name="txtDebugOutput" Background="Transparent" BorderThickness="0" 
                                     FontFamily="Consolas" IsReadOnly="True" TextWrapping="Wrap"
                                     Foreground="#00FF00" Padding="15" FontSize="11"/>
                        </ScrollViewer>
                    </Border>
                    
                    <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,15,0,0" HorizontalAlignment="Left">
                        <Button Content="📝 Log-Datei öffnen" x:Name="btnOpenLog" Style="{StaticResource ModernButton}"/>
                        <Button Content="🗑️ Log leeren" x:Name="btnClearLog" Style="{StaticResource ModernButton}" Background="#DC3545" Margin="10,0,0,0"/>
                    </StackPanel>
                </Grid>
            </TabItem>
        </TabControl>
        
        <!-- Footer -->
        <Border Grid.Row="2" Background="#F8F9FA" BorderThickness="0,1,0,0" BorderBrush="#DEE2E6" Padding="20,10">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock x:Name="txtStatus" Text="Status: Bereit" VerticalAlignment="Center" FontSize="12" Foreground="#6C757D"/>
                <TextBlock Grid.Column="1" Text="easyWSAudit v0.0.2" VerticalAlignment="Center" FontSize="12" Foreground="#6C757D"/>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

# Lade das XAML
Write-DebugLog "Lade XAML für UI..." "UI"
$reader = [System.Xml.XmlNodeReader]::new($xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)
Write-DebugLog "XAML geladen, Fenster erstellt" "UI"

# Hole die UI-Elemente
$txtServerName = $window.FindName("txtServerName")
$btnFullAudit = $window.FindName("btnFullAudit")
$spOptions = $window.FindName("spOptions")
$btnSelectAll = $window.FindName("btnSelectAll")
$btnSelectNone = $window.FindName("btnSelectNone")
$btnRunAudit = $window.FindName("btnRunAudit")
$progressBar = $window.FindName("progressBar")
$txtProgress = $window.FindName("txtProgress")
$txtStatusLog = $window.FindName("txtStatusLog")
$txtResults = $window.FindName("txtResults")
$btnExportHTML = $window.FindName("btnExportHTML")
$btnCopyToClipboard = $window.FindName("btnCopyToClipboard")
$txtStatus = $window.FindName("txtStatus")

# Debug-Elemente
$script:txtDebugOutput = $window.FindName("txtDebugOutput")
$btnOpenLog = $window.FindName("btnOpenLog")
$btnClearLog = $window.FindName("btnClearLog")

Write-DebugLog "UI-Elemente initialisiert" "UI"

# Servername anzeigen
$txtServerName.Text = "Server: $env:COMPUTERNAME"

# Dictionary für Checkboxen
$checkboxes = @{}

# Erstelle die Checkboxen für die Audit-Optionen gruppiert nach Kategorien
Write-DebugLog "Erstelle Checkboxen für Audit-Optionen..." "UI"

$categories = @{}
foreach ($cmd in $commands) {
    $category = if ($cmd.Category) { $cmd.Category } else { "Allgemein" }
    if (-not $categories.ContainsKey($category)) {
        $categories[$category] = @()
    }
    $categories[$category] += $cmd
}

foreach ($category in $categories.Keys | Sort-Object) {
    # Kategorie-Header
    $categoryHeader = New-Object System.Windows.Controls.TextBlock
    $categoryHeader.Text = "📁 $category"
    $categoryHeader.Style = $window.FindResource("CategoryHeader")
    $spOptions.Children.Add($categoryHeader)
    
    # Checkboxen für diese Kategorie
    foreach ($cmd in $categories[$category]) {
        $checkbox = New-Object System.Windows.Controls.CheckBox
        $checkbox.Content = $cmd.Name
        $checkbox.IsChecked = $true
        $checkbox.Style = $window.FindResource("CheckboxStyle")
        
        # Überprüfe, ob diese Option mit einer Serverrolle verbunden ist
        if ($cmd.ContainsKey("FeatureName")) {
            $isRoleInstalled = Test-ServerRole -FeatureName $cmd.FeatureName
            if (-not $isRoleInstalled) {
                $checkbox.IsEnabled = $false
                $checkbox.Content = "$($cmd.Name) (Nicht installiert)"
                $checkbox.IsChecked = $false
            }
        }
        
        $spOptions.Children.Add($checkbox)
        $checkboxes[$cmd.Name] = $checkbox
    }
}
Write-DebugLog "Checkboxen erstellt für $($checkboxes.Count) Optionen" "UI"

# Button-Event-Handler

# "Alle auswählen" Button
$btnSelectAll.Add_Click({
    Write-DebugLog "Alle Optionen auswählen" "UI"
    foreach ($key in $checkboxes.Keys) {
        if ($checkboxes[$key].IsEnabled) {
            $checkboxes[$key].IsChecked = $true
        }
    }
})

# "Alle abwählen" Button
$btnSelectNone.Add_Click({
    Write-DebugLog "Alle Optionen abwählen" "UI"
    foreach ($key in $checkboxes.Keys) {
        $checkboxes[$key].IsChecked = $false
    }
})

# Vollständiges Audit Button
$btnFullAudit.Add_Click({
    Write-DebugLog "Vollständiges Audit gestartet" "Audit"
    # Alle verfügbaren Optionen auswählen
    foreach ($key in $checkboxes.Keys) {
        if ($checkboxes[$key].IsEnabled) {
            $checkboxes[$key].IsChecked = $true
        }
    }
    # Audit starten
    Start-AuditProcess
})

# "Audit starten" Button
$btnRunAudit.Add_Click({
    Write-DebugLog "Benutzerdefiniertes Audit gestartet" "Audit"
    Start-AuditProcess
})

# Hauptfunktion für die Audit-Durchführung (Synchron)
function Start-AuditProcess {
    # UI vorbereiten
    $btnRunAudit.IsEnabled = $false
    $btnFullAudit.IsEnabled = $false
    $btnExportHTML.IsEnabled = $false
    $btnCopyToClipboard.IsEnabled = $false
    
    $txtResults.Text = ""
    $txtStatusLog.Text = ""
    $progressBar.Value = 0
    $txtStatus.Text = "Status: Audit läuft..."
    
    # Sammle ausgewählte Befehle
    $selectedCommands = @()
    foreach ($cmd in $commands) {
        if ($checkboxes[$cmd.Name].IsChecked) {
            $selectedCommands += $cmd
        }
    }
    
    Write-DebugLog "Starte Audit mit $($selectedCommands.Count) ausgewählten Befehlen" "Audit"
    
    $global:auditResults = @{}
    $allResults = ""
    $progressStep = 100.0 / $selectedCommands.Count
    $currentProgress = 0
    
    for ($i = 0; $i -lt $selectedCommands.Count; $i++) {
        $cmd = $selectedCommands[$i]
        
        # UI aktualisieren
        $window.Dispatcher.Invoke([Action]{
            $txtProgress.Text = "Verarbeite: $($cmd.Name) ($($i+1)/$($selectedCommands.Count))"
            $txtStatusLog.Text += "➤ $($cmd.Name)...`r`n"
            $progressBar.Value = $currentProgress
        }, "Normal")
        
        Write-DebugLog "Führe aus ($($i+1)/$($selectedCommands.Count)): $($cmd.Name)" "Audit"
        
        try {
            if ($cmd.Type -eq "PowerShell") {
                $result = Invoke-PSCommand -Command $cmd.Command
            } else {
                $result = "CMD-Befehle werden in dieser Version nicht unterstützt"
            }
            
            $global:auditResults[$cmd.Name] = $result
            $allResults += "`r`n=== $($cmd.Name) ===`r`n$result`r`n"
            
            # Erfolg in Status-Log
            $window.Dispatcher.Invoke([Action]{
                $txtStatusLog.Text += "  ✅ Erfolgreich`r`n"
            }, "Normal")
            
        } catch {
            $errorMsg = "Fehler: $($_.Exception.Message)"
            $global:auditResults[$cmd.Name] = $errorMsg
            $allResults += "`r`n=== $($cmd.Name) ===`r`n$errorMsg`r`n"
            
            # Fehler in Status-Log
            $window.Dispatcher.Invoke([Action]{
                $txtStatusLog.Text += "  ❌ Fehler: $($_.Exception.Message)`r`n"
            }, "Normal")
            
            Write-DebugLog "FEHLER bei $($cmd.Name): $($_.Exception.Message)" "Audit"
        }
        
        $currentProgress += $progressStep
        
        # Kurze Pause für UI-Updates
        Start-Sleep -Milliseconds 50
    }
    
    # Audit abgeschlossen
    $window.Dispatcher.Invoke([Action]{
        $progressBar.Value = 100
        $txtProgress.Text = "Audit abgeschlossen! $($selectedCommands.Count) Befehle ausgeführt."
        $txtStatusLog.Text += "`r`n🎉 Audit erfolgreich abgeschlossen!"
        $txtResults.Text = $allResults
        $txtStatus.Text = "Status: Audit abgeschlossen - $($global:auditResults.Count) Ergebnisse"
        
        # Buttons wieder aktivieren
        $btnRunAudit.IsEnabled = $true
        $btnFullAudit.IsEnabled = $true
        $btnExportHTML.IsEnabled = $true
        $btnCopyToClipboard.IsEnabled = $true
    }, "Normal")
    
    Write-DebugLog "Audit abgeschlossen mit $($global:auditResults.Count) Ergebnissen" "Audit"
}

# Export-Button-Funktionalität
$btnExportHTML.Add_Click({
    Write-DebugLog "HTML-Export gestartet" "Export"
    
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "HTML Files (*.html)|*.html"
    $saveFileDialog.Title = "Speichern Sie den Audit-Bericht"
    $saveFileDialog.FileName = "ServerAudit_$($env:COMPUTERNAME)_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
    
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtStatus.Text = "Status: Exportiere HTML..."
        
        try {
            Export-AuditToHTML -Results $global:auditResults -FilePath $saveFileDialog.FileName
            $txtStatus.Text = "Status: Export erfolgreich abgeschlossen"
            [System.Windows.MessageBox]::Show("Bericht wurde erfolgreich exportiert:`r`n$($saveFileDialog.FileName)", "Export erfolgreich", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        } catch {
            $txtStatus.Text = "Status: Fehler beim Export"
            [System.Windows.MessageBox]::Show("Fehler beim Export:`r`n$($_.Exception.Message)", "Export Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        }
    }
})

# Zwischenablage-Button
$btnCopyToClipboard.Add_Click({
    Write-DebugLog "Kopiere Ergebnisse in Zwischenablage" "UI"
    try {
        $txtResults.Text | Set-Clipboard
        $txtStatus.Text = "Status: Ergebnisse in Zwischenablage kopiert"
        [System.Windows.MessageBox]::Show("Audit-Ergebnisse wurden in die Zwischenablage kopiert.", "Kopiert", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
    } catch {
        [System.Windows.MessageBox]::Show("Fehler beim Kopieren in die Zwischenablage.", "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
})

# Debug-Funktionen
$btnOpenLog.Add_Click({
    Write-DebugLog "Öffne Log-Datei" "Debug"
    if (Test-Path $DebugLogPath) {
        Start-Process notepad.exe -ArgumentList $DebugLogPath
    } else {
        [System.Windows.MessageBox]::Show("Log-Datei nicht gefunden.", "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
})

$btnClearLog.Add_Click({
    Write-DebugLog "Debug-Log wird geleert" "Debug"
    $script:txtDebugOutput.Text = ""
    if ($DEBUG) {
        $clearMessage = "=== Debug-Log gelöscht: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss") ==="
        Set-Content -Path $DebugLogPath -Value $clearMessage -Force
        Write-Host $clearMessage -ForegroundColor Cyan
    }
})

Write-DebugLog "UI-Initialisierung abgeschlossen" "Init"

# Zeige das Fenster an
$txtStatus.Text = "Status: Bereit für Audit"
Write-DebugLog "Zeige Hauptfenster" "UI"
$null = $window.ShowDialog()
Write-DebugLog "Anwendung geschlossen" "Shutdown"
