# Move the Write-Host function definition to the top of the file and make it globally accessible
function Write-Log {
    param (
        [string]$Message,
        [string]$Type = "INFO"
    )
    
    try {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logLine = "[$timestamp] [$Type] $Message"
        Write-Host $logLine
        
        # Add logging to file
        $logFile = Join-Path -Path $PSScriptRoot -ChildPath "easyWSUS.log"
        Add-Content -Path $logFile -Value $logLine -ErrorAction SilentlyContinue
    }
    catch {
        # Fallback if logging fails
        Write-Host "Error in logging: $($_.Exception.Message)"
    }
}

# Load required assemblies for WPF
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Xaml

# Load required assemblies for Windows Forms
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Windows.Forms.DataVisualization

# Global variable to store GUI elements
$script:GUI = @{}
$script:timers = @{}

# Funktion zum Laden der XAML-Datei
function Load-XamlGUI {
    param (
        [string]$XamlPath
    )
    
    try {
        Write-Host "Versuche XAML zu laden von: $XamlPath"
        
        # Prüfen, ob die Datei existiert
        if (-not (Test-Path -Path $XamlPath)) {
            throw "XAML-Datei nicht gefunden: $XamlPath"
        }
        
        # XAML-Datei als Text laden
        $xamlContent = Get-Content -Path $XamlPath -Raw -ErrorAction Stop
        
        if ([string]::IsNullOrWhiteSpace($xamlContent)) {
            throw "XAML-Datei ist leer oder enthält keine Daten."
        }
        
        # Namespace-Definitionen beibehalten, aber Klassen-Attribut und Click-Events entfernen
        $xamlContent = $xamlContent -replace 'x:Class="[^"]*"', ''
        $xamlContent = $xamlContent -replace 'Click="[^"]*"', '' # Event-Handler aus XAML entfernen
        
        # XAML als XML-Dokument laden
        [xml]$xaml = $xamlContent
        
        if ($null -eq $xaml) {
            throw "XAML konnte nicht als XML geladen werden."
        }
        
        # Namespace für XAML-Elemente hinzufügen
        $nsManager = New-Object System.Xml.XmlNamespaceManager($xaml.NameTable)
        $nsManager.AddNamespace("x", "http://schemas.microsoft.com/winfx/2006/xaml")
        
        # XAML-Reader erstellen und Fenster laden
        $reader = New-Object System.Xml.XmlNodeReader $xaml
        $window = [Windows.Markup.XamlReader]::Load($reader)
        
        if ($null -eq $window) {
            throw "Window konnte nicht aus XAML erstellt werden."
        }
        
        Write-Host "XAML erfolgreich geladen."
        
        # GUI-Elemente in eine Hashtable für den Zugriff speichern
        $GUI = @{}
        $nodes = $xaml.SelectNodes("//*[@x:Name]", $nsManager)
        
        if ($null -eq $nodes -or $nodes.Count -eq 0) {
            Write-Host "Warnung: Keine benannten Elemente in der XAML gefunden." "WARNING"
        } else {
            Write-Host "$($nodes.Count) benannte Elemente gefunden."
        }
        
        foreach ($node in $nodes) {
            try {
                $name = $node.GetAttribute("Name", "http://schemas.microsoft.com/winfx/2006/xaml")
                
                if ([string]::IsNullOrEmpty($name)) {
                    Write-Host "Element ohne Namen übersprungen." "WARNING"
                    continue
                }
                
                $element = $window.FindName($name)
                
                if ($null -ne $element) {
                    $GUI[$name] = $element
                    Write-Host "Element geladen: $name" "INFO"
                } else {
                    Write-Host "Element '$name' konnte nicht gefunden werden." "WARNING"
                }
            }
            catch {
                Write-Host "Fehler beim Laden des Elements '$($node.Name)': $($_.Exception.Message)" "ERROR"
            }
        }
        
        return @{
            Window = $window
            GUI = $GUI
        }
    }
    catch {
        Write-Host "Fehler beim Laden der XAML: $($_.Exception.Message)" "ERROR"
        Write-Host "Stack Trace: $($_.Exception.StackTrace)" "ERROR"
        return $null
    }
}

# Funktion für direkten Zugriff auf Microsoft Dokumentation
function Open-MicrosoftDocs {
    param (
        [Parameter(Mandatory = $true)]
        [string]$SearchTerm,
        [ValidateSet("All", "WindowsUpdate", "WindowsServer", "TroubleShooting")]
        [string]$Category = "WindowsUpdate"
    )
    
    try {
        # System.Web.HttpUtility nutzen
        Add-Type -AssemblyName System.Web
        
        $encodedTerm = [System.Web.HttpUtility]::UrlEncode($SearchTerm)
        
        # URL basierend auf Kategorie anpassen
        $baseUrl = "https://learn.microsoft.com/de-de/search/?terms=$encodedTerm"
        
        switch ($Category) {
            "WindowsUpdate" { 
                $url = "$baseUrl&category=Documentation&filter=products-Windows%20Update" 
            }
            "WindowsServer" { 
                $url = "$baseUrl&category=Documentation&filter=products-Windows%20Server" 
            }
            "TroubleShooting" { 
                $url = "$baseUrl&category=Documentation&view=troubleshoot" 
            }
            default { 
                $url = $baseUrl 
            }
        }
        
        Write-Host "Öffne Microsoft Docs mit Suche nach: $SearchTerm (Kategorie: $Category)"
        Start-Process $url
        
        return $true
    }
    catch {
        Write-Host "Fehler beim Öffnen der Microsoft Docs: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

# Funktion zum Starten/Stoppen/Neustarten von Diensten
function global:Set-ServiceOperation {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ServiceName,
     
        [Parameter(Mandatory=$true)]
        [ValidateSet("Start", "Stop", "Restart")]
        [string]$Operation
    )
    
    try {
        # Prüfen, ob Administrator-Rechte vorhanden sind
        $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        if (-not $isAdmin) {
            Write-Host "Administratorrechte erforderlich für Dienst-Operationen" "ERROR"
            return $false
        }
        
        # Dienst abrufen
        $service = Get-Service -Name $ServiceName -ErrorAction Stop
        
        switch ($Operation) {
            "Start" {
                if ($service.Status -ne "Running") {
                    $service | Start-Service -ErrorAction Stop
                    Write-Host "Dienst '$ServiceName' wurde gestartet" "SUCCESS"
                } else {
                    Write-Host "Dienst '$ServiceName' läuft bereits"
                }
            }
            "Stop" {
                if ($service.Status -ne "Stopped") {
                    $service | Stop-Service -Force -ErrorAction Stop
                    Write-Host "Dienst '$ServiceName' wurde gestoppt" "SUCCESS"
                } else {
                    Write-Host "Dienst '$ServiceName' ist bereits gestoppt"
                }
            }
            "Restart" {
                $service | Restart-Service -Force -ErrorAction Stop
                Write-Host "Dienst '$ServiceName' wurde neu gestartet" "SUCCESS"
            }
        }
        
        return $true
    }
    catch {
        Write-Host "Fehler bei der Dienst-Operation ($Operation) für '$ServiceName': $($_.Exception.Message)" "ERROR"
        return $false
    }
}

# Funktion zum Starten der Windows Update-Problembehandlung
function Run-UpdateTroubleshooter {
    try {
        $diagPath = "$env:SystemRoot\System32\msdt.exe"
        
        if (-not (Test-Path -Path $diagPath)) {
            throw "Problembehandlungstool (msdt.exe) wurde nicht gefunden"
        }
        
        $arguments = "/id WindowsUpdateDiagnostic"
        Write-Host "Starte Windows Update-Problembehandlung..."
        Start-Process $diagPath -ArgumentList $arguments
        
        return $true
    }
    catch {
        Write-Host "Fehler beim Starten der Windows Update-Problembehandlung: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

# Service-Button Click-Handler für die DataGrid-Zellen
function btnStartService_Click {
    param($sender, $e)
    
    try {
        $serviceName = $sender.Tag
        Write-Host "Starte Dienst: $serviceName"
        
        Set-ServiceOperation -ServiceName $serviceName -Operation "Start"
        
        # Dienste-View aktualisieren
        Update-ServiceStatus -GUI $script:GUI
    }
    catch {
        Write-Host "Fehler beim Starten des Dienstes $serviceName`: $($_.Exception.Message)" "ERROR"
        [System.Windows.MessageBox]::Show("Fehler beim Starten des Dienstes: $($_.Exception.Message)", "Fehler", 
            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
}

function btnStopService_Click {
    param($sender, $e)
    
    try {
        $serviceName = $sender.Tag
        Write-Host "Stoppe Dienst: $serviceName"
        
        Set-ServiceOperation -ServiceName $serviceName -Operation "Stop"
        
        # Dienste-View aktualisieren
        Update-ServiceStatus -GUI $script:GUI
    }
    catch {
        Write-Host "Fehler beim Stoppen des Dienstes $serviceName`: $($_.Exception.Message)" "ERROR"
        [System.Windows.MessageBox]::Show("Fehler beim Stoppen des Dienstes: $($_.Exception.Message)", "Fehler", 
            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
}

function btnRestartService_Click {
    param($sender, $e)
    
    try {
        $serviceName = $sender.Tag
        Write-Host "Starte Dienst neu: $serviceName"
        
        Set-ServiceOperation -ServiceName $serviceName -Operation "Restart"
        
        # Dienste-View aktualisieren
        Update-ServiceStatus -GUI $script:GUI
    }
    catch {
        Write-Host "Fehler beim Neustarten des Dienstes $serviceName`: $($_.Exception.Message)" "ERROR"
        [System.Windows.MessageBox]::Show("Fehler beim Neustarten des Dienstes: $($_.Exception.Message)", "Fehler", 
            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
}

# Funktion zum Abrufen des Status relevanter Windows Update Dienste
function Get-ServiceStatus {
    try {
        $services = Get-Service -Name wuauserv, BITS, cryptsvc, TrustedInstaller, DoSvc, UsoSvc, WaaSMedicSvc -ErrorAction SilentlyContinue | 
            Select-Object @{Name="Name"; Expression={$_.Name}}, 
                          @{Name="DisplayName"; Expression={$_.DisplayName}},
                          @{Name="Status"; Expression={$_.Status}},
                          @{Name="StartType"; Expression={(Get-Service $_.Name).StartType}}
        
        Write-Host "Dienst-Status abgerufen: $($services.Count) Dienste gefunden"
        return $services
    }
    catch {
        Write-Host "Fehler beim Abrufen der Dienst-Status: $($_.Exception.Message)" "ERROR"
        return $null
    }
}

# Funktion zum Neustarten aller Windows Update Dienste
function Restart-UpdateServices {
    try {
        # Prüfen, ob Administrator-Rechte vorhanden sind
        $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        if (-not $isAdmin) {
            Write-Host "Administratorrechte erforderlich zum Neustart der Dienste" "ERROR"
            return $false
        }
        
        # Dienste stoppen in spezifischer Reihenfolge
        Write-Host "Stoppe Update-Dienste..."
        Get-Service -Name wuauserv, BITS, cryptsvc, UsoSvc, DoSvc, WaaSMedicSvc -ErrorAction SilentlyContinue | Stop-Service -Force -ErrorAction SilentlyContinue
        
        # Kurze Wartezeit
        Start-Sleep -Seconds 2
        
        # Dienste starten in umgekehrter Reihenfolge
        Write-Host "Starte Update-Dienste..."
        Get-Service -Name cryptsvc, BITS, wuauserv, WaaSMedicSvc, DoSvc, UsoSvc -ErrorAction SilentlyContinue | Start-Service -ErrorAction SilentlyContinue
        
        Write-Host "Update-Dienste wurden neu gestartet" "SUCCESS"
        return $true
    }
    catch {
        Write-Host "Fehler beim Neustart der Update-Dienste: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

# CPU-, RAM- und Festplattenauslastung abrufen
function Get-SystemUsage {
    try {
        # RAM-Auslastung 
        $osInfo = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
        $totalRAM = [math]::Round($osInfo.TotalVisibleMemorySize / 1MB, 2)
        $freeRAM = [math]::Round($osInfo.FreePhysicalMemory / 1MB, 2)
        $usedRAM = $totalRAM - $freeRAM
        $ramUsagePercent = [math]::Round(($usedRAM / $totalRAM) * 100, 2)
        
        # Festplattenauslastung 
        $diskC = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'" -ErrorAction Stop
        $totalDisk = [math]::Round($diskC.Size / 1GB, 2)
        $freeDisk = [math]::Round($diskC.FreeSpace / 1GB, 2)
        $usedDisk = $totalDisk - $freeDisk
        $diskUsagePercent = [math]::Round(($usedDisk / $totalDisk) * 100, 2)
        
        $result = [PSCustomObject]@{
            TotalRAM = $totalRAM
            FreeRAM = $freeRAM
            UsedRAM = $usedRAM
            RAMUsagePercent = $ramUsagePercent
            
            TotalDiskSpace = $totalDisk
            FreeDiskSpace = $freeDisk
            UsedDiskSpace = $usedDisk
            DiskUsagePercent = $diskUsagePercent
            
            Timestamp = Get-Date
        }
        
        Write-Host "Systemauslastung erfolgreich abgerufen: RAM $($ramUsagePercent)%, Disk $($diskUsagePercent)%"
        return $result
    }
    catch {
        Write-Host "Fehler beim Abrufen der Systemauslastung: $($_.Exception.Message)" "ERROR"
        # Fallback-Ergebnis zurückgeben mit Default-Werten
        return [PSCustomObject]@{
            CPUUsage = 0
            CPUUsagePercent = 0
            TotalRAM = 0
            FreeRAM = 0
            UsedRAM = 0
            RAMUsagePercent = 0
            TotalDiskSpace = 0
            FreeDiskSpace = 0
            UsedDiskSpace = 0
            DiskUsagePercent = 0
            Timestamp = Get-Date
        }
    }
}

# Event-Handler registrieren
function Register-EventHandlers {
    param (
        $GUI,
        $Window
    )
    
    try {
        # Dashboard Button
        if ($GUI.ContainsKey("btnDashboard") -and $null -ne $GUI["btnDashboard"]) {
            $GUI["btnDashboard"].Add_Click({
                $GUI["Dashboard"].Visibility = "Visible"
                $GUI["WUSASettings"].Visibility = "Collapsed"
                $GUI["Settings"].Visibility = "Collapsed"
                $GUI["Troubleshooting"].Visibility = "Collapsed"
                $GUI["ServiceMonitor"].Visibility = "Collapsed"
                $GUI["SystemStats"].Visibility = "Collapsed"
                $GUI["LogViewer"].Visibility = "Collapsed"
                $GUI["UpdateManager"].Visibility = "Collapsed"
                $GUI["UpdateHistory"].Visibility = "Collapsed"
                $GUI["Scheduler"].Visibility = "Collapsed"
                $GUI["Toolbox"].Visibility = "Collapsed"
                $GUI["SystemReport"].Visibility = "Collapsed"
                
                Write-Host "Dashboard Tab ausgewählt"
            })
        }
        
        # Update Manager Button
        if ($GUI.ContainsKey("btnUpdateManager") -and $null -ne $GUI["btnUpdateManager"]) {
            $GUI["btnUpdateManager"].Add_Click({
                $GUI["Dashboard"].Visibility = "Collapsed"
                $GUI["WUSASettings"].Visibility = "Collapsed"
                $GUI["Settings"].Visibility = "Collapsed"
                $GUI["Troubleshooting"].Visibility = "Collapsed"
                $GUI["ServiceMonitor"].Visibility = "Collapsed"
                $GUI["SystemStats"].Visibility = "Collapsed"
                $GUI["LogViewer"].Visibility = "Collapsed"
                $GUI["UpdateWartung"].Visibility = "Visible"
                $GUI["UpdateHistory"].Visibility = "Collapsed"
                $GUI["Scheduler"].Visibility = "Collapsed"
                $GUI["Toolbox"].Visibility = "Collapsed"
                $GUI["SystemReport"].Visibility = "Collapsed"
                
                Write-Host "Update Manager Tab ausgewählt"
            })
        }
        
        # Update History Button
        if ($GUI.ContainsKey("btnUpdateHistory") -and $null -ne $GUI["btnUpdateHistory"]) {
            $GUI["btnUpdateHistory"].Add_Click({
                $GUI["Dashboard"].Visibility = "Collapsed"
                $GUI["WUSASettings"].Visibility = "Collapsed"
                $GUI["Settings"].Visibility = "Collapsed"
                $GUI["Troubleshooting"].Visibility = "Collapsed"
                $GUI["ServiceMonitor"].Visibility = "Collapsed"
                $GUI["SystemStats"].Visibility = "Collapsed"
                $GUI["LogViewer"].Visibility = "Collapsed"
                $GUI["UpdateManager"].Visibility = "Collapsed"
                $GUI["UpdateHistory"].Visibility = "Visible"
                $GUI["Scheduler"].Visibility = "Collapsed"
                $GUI["Toolbox"].Visibility = "Collapsed"
                $GUI["SystemReport"].Visibility = "Collapsed"
                
                Write-Host "Update History Tab ausgewählt"
            })
        }

        # Service Monitor Button
        if ($GUI.ContainsKey("btnServiceMonitor") -and $null -ne $GUI["btnServiceMonitor"]) {
            $GUI["btnServiceMonitor"].Add_Click({
                $GUI["Dashboard"].Visibility = "Collapsed"
                $GUI["WUSASettings"].Visibility = "Collapsed"
                $GUI["Settings"].Visibility = "Collapsed"
                $GUI["Troubleshooting"].Visibility = "Collapsed"
                $GUI["ServiceMonitor"].Visibility = "Visible"
                $GUI["SystemStats"].Visibility = "Collapsed"
                $GUI["LogViewer"].Visibility = "Collapsed"
                $GUI["UpdateManager"].Visibility = "Collapsed"
                $GUI["UpdateHistory"].Visibility = "Collapsed"
                $GUI["Scheduler"].Visibility = "Collapsed"
                $GUI["Toolbox"].Visibility = "Collapsed"
                $GUI["SystemReport"].Visibility = "Collapsed"
                
                Write-Host "Service Monitor Tab ausgewählt"
                
                # Dienste-Status aktualisieren
                Update-ServiceMonitorUI -GUI $GUI
            })
        }
        
        # System Info Button
        if ($GUI.ContainsKey("btnSystemStats") -and $null -ne $GUI["btnSystemStats"]) {
            $GUI["btnSystemStats"].Add_Click({
                $GUI["Dashboard"].Visibility = "Collapsed"
                $GUI["WUSASettings"].Visibility = "Collapsed"
                $GUI["Settings"].Visibility = "Collapsed"
                $GUI["Troubleshooting"].Visibility = "Collapsed"
                $GUI["ServiceMonitor"].Visibility = "Collapsed"
                $GUI["SystemStats"].Visibility = "Visible"
                $GUI["LogViewer"].Visibility = "Collapsed"
                $GUI["UpdateManager"].Visibility = "Collapsed"
                $GUI["UpdateHistory"].Visibility = "Collapsed"
                $GUI["Scheduler"].Visibility = "Collapsed"
                $GUI["Toolbox"].Visibility = "Collapsed"
                $GUI["SystemReport"].Visibility = "Collapsed"
                
                Write-Host "System Stats Tab ausgewählt"
                
                # Systemressourcen laden
                Update-SystemResources -GUI $GUI

            })
        }
        
        # Log Viewer Button
        if ($GUI.ContainsKey("btnLogViewer") -and $null -ne $GUI["btnLogViewer"]) {
            $GUI["btnLogViewer"].Add_Click({
                $GUI["Dashboard"].Visibility = "Collapsed"
                $GUI["WUSASettings"].Visibility = "Collapsed"
                $GUI["Settings"].Visibility = "Collapsed"
                $GUI["Troubleshooting"].Visibility = "Collapsed"
                $GUI["ServiceMonitor"].Visibility = "Collapsed"
                $GUI["SystemStats"].Visibility = "Collapsed"
                $GUI["LogViewer"].Visibility = "Visible"
                $GUI["UpdateManager"].Visibility = "Collapsed"
                $GUI["UpdateHistory"].Visibility = "Collapsed"
                $GUI["Scheduler"].Visibility = "Collapsed"
                $GUI["Toolbox"].Visibility = "Collapsed"
                $GUI["SystemReport"].Visibility = "Collapsed"
                
                Write-Host "Log Viewer Tab ausgewählt"
                
                # Log-Daten initial laden, wenn noch nicht geschehen
                if ($GUI.ContainsKey("txtLogContent") -and [string]::IsNullOrEmpty($GUI["txtLogContent"].Text)) {
                    Update-LogView -GUI $GUI
                }
            })
        }
        
        # Toolbox Button
        if ($GUI.ContainsKey("btnToolbox") -and $null -ne $GUI["btnToolbox"]) {
            $GUI["btnToolbox"].Add_Click({
                $GUI["Dashboard"].Visibility = "Collapsed"
                $GUI["WUSASettings"].Visibility = "Collapsed"
                $GUI["Settings"].Visibility = "Collapsed"
                $GUI["Troubleshooting"].Visibility = "Collapsed"
                $GUI["ServiceMonitor"].Visibility = "Collapsed"
                $GUI["SystemStats"].Visibility = "Collapsed"
                $GUI["LogViewer"].Visibility = "Collapsed"
                $GUI["UpdateManager"].Visibility = "Collapsed"
                $GUI["UpdateHistory"].Visibility = "Collapsed"
                $GUI["Scheduler"].Visibility = "Collapsed"
                $GUI["Toolbox"].Visibility = "Visible"
                $GUI["SystemReport"].Visibility = "Collapsed"
                
                Write-Host "Toolbox Tab ausgewählt"
            })
        }
        
        # System Report Button
        if ($GUI.ContainsKey("btnSystemReport") -and $null -ne $GUI["btnSystemReport"]) {
            $GUI["btnSystemReport"].Add_Click({
                $GUI["Dashboard"].Visibility = "Collapsed"
                $GUI["WUSASettings"].Visibility = "Collapsed"
                $GUI["Settings"].Visibility = "Collapsed"
                $GUI["Troubleshooting"].Visibility = "Collapsed"
                $GUI["ServiceMonitor"].Visibility = "Collapsed"
                $GUI["SystemStats"].Visibility = "Collapsed"
                $GUI["LogViewer"].Visibility = "Collapsed"
                $GUI["UpdateManager"].Visibility = "Collapsed"
                $GUI["UpdateHistory"].Visibility = "Collapsed"
                $GUI["Scheduler"].Visibility = "Collapsed"
                $GUI["Toolbox"].Visibility = "Collapsed"
                $GUI["SystemReport"].Visibility = "Visible"
                
                Write-Host "System Report Tab ausgewählt"
            })
        }
        
        # WUSA Settings Button
        if ($GUI.ContainsKey("btnWUSASettings") -and $null -ne $GUI["btnWUSASettings"]) {
            $GUI["btnWUSASettings"].Add_Click({
                $GUI["Dashboard"].Visibility = "Collapsed"
                $GUI["WUSASettings"].Visibility = "Visible"
                $GUI["Settings"].Visibility = "Collapsed"
                $GUI["Troubleshooting"].Visibility = "Collapsed"
                $GUI["ServiceMonitor"].Visibility = "Collapsed"
                $GUI["SystemStats"].Visibility = "Collapsed"
                $GUI["LogViewer"].Visibility = "Collapsed"
                $GUI["UpdateManager"].Visibility = "Collapsed"
                $GUI["UpdateHistory"].Visibility = "Collapsed"
                $GUI["Scheduler"].Visibility = "Collapsed"
                $GUI["Toolbox"].Visibility = "Collapsed"
                $GUI["SystemReport"].Visibility = "Collapsed"
                
                Write-Host "WUSA Settings Tab ausgewählt"
            })
        }
        
        # Troubleshooting Button
        if ($GUI.ContainsKey("btnTroubleshooting") -and $null -ne $GUI["btnTroubleshooting"]) {
            $GUI["btnTroubleshooting"].Add_Click({
                $GUI["Dashboard"].Visibility = "Collapsed"
                $GUI["WUSASettings"].Visibility = "Collapsed"
                $GUI["Settings"].Visibility = "Collapsed"
                $GUI["Troubleshooting"].Visibility = "Visible"
                $GUI["ServiceMonitor"].Visibility = "Collapsed"
                $GUI["SystemStats"].Visibility = "Collapsed"
                $GUI["LogViewer"].Visibility = "Collapsed"
                $GUI["UpdateManager"].Visibility = "Collapsed"
                $GUI["UpdateHistory"].Visibility = "Collapsed"
                $GUI["Scheduler"].Visibility = "Collapsed"
                $GUI["Toolbox"].Visibility = "Collapsed"
                $GUI["SystemReport"].Visibility = "Collapsed"
                
                Write-Host "Troubleshooting Tab ausgewählt"
            })
        }
        
        # Refresh Services Button
        if ($GUI.ContainsKey("btnRefreshServices") -and $null -ne $GUI["btnRefreshServices"]) {
            $GUI["btnRefreshServices"].Add_Click({
                Update-ServiceMonitorUI -GUI $GUI
            })
        }
        
        # Restart All Services Button
        if ($GUI.ContainsKey("btnRestartAllServices") -and $null -ne $GUI["btnRestartAllServices"]) {
            $GUI["btnRestartAllServices"].Add_Click({
                $result = Restart-UpdateServices
                
                if ($result) {
                    [System.Windows.MessageBox]::Show("Windows Update-Dienste wurden erfolgreich neu gestartet.", "Erfolg", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
                    Update-ServiceMonitorUI -GUI $GUI
                } else {
                    [System.Windows.MessageBox]::Show("Fehler beim Neustarten der Windows Update-Dienste.", "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                }
            })
        }
        
        # Registriere die Service-spezifischen Button-Events
        Register-ServiceButtonEvents -GUI $GUI
        
        # Refresh Resources Button
        if ($GUI.ContainsKey("btnRefreshResources") -and $null -ne $GUI["btnRefreshResources"]) {
            $GUI["btnRefreshResources"].Add_Click({
                Write-Host "Aktualisiere Systemressourcen..."
                Update-SystemResources -GUI $GUI
            })
        }
        
        # Log Load Button
        if ($GUI.ContainsKey("btnLoadLog") -and $null -ne $GUI["btnLoadLog"]) {
            $GUI["btnLoadLog"].Add_Click({
                Write-Host "Lade Log-Daten..."
                Update-LogView -GUI $GUI
            })
        }
        
        # Log Type ComboBox
        if ($GUI.ContainsKey("cmbLogType") -and $null -ne $GUI["cmbLogType"]) {
            $GUI["cmbLogType"].Add_SelectionChanged({
                if ($GUI["LogViewer"].Visibility -eq "Visible") {
                    Update-LogView -GUI $GUI
                }
            })
        }
        
        # Log Lines ComboBox
        if ($GUI.ContainsKey("cmbLogLines") -and $null -ne $GUI["cmbLogLines"]) {
            $GUI["cmbLogLines"].Add_SelectionChanged({
                if ($GUI["LogViewer"].Visibility -eq "Visible") {
                    Update-LogView -GUI $GUI
                }
            })
        }
        
        # Export Log Button
        if ($GUI.ContainsKey("btnExportLog") -and $null -ne $GUI["btnExportLog"]) {
            $GUI["btnExportLog"].Add_Click({
                # Speicherdialog öffnen
                $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
                $saveDialog.Filter = "Log-Dateien (*.log)|*.log|Textdateien (*.txt)|*.txt|Alle Dateien (*.*)|*.*"
                $saveDialog.DefaultExt = ".log"
                $saveDialog.Title = "Log-Datei speichern"
                $saveDialog.FileName = "WindowsUpdate_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
                
                if ($saveDialog.ShowDialog()) {
                    try {
                        $GUI["txtLogContent"].Text | Out-File -FilePath $saveDialog.FileName -Encoding utf8
                        $GUI["txtLogStatus"].Text = "Log gespeichert unter: $($saveDialog.FileName)"
                    }
                    catch {
                        $GUI["txtLogStatus"].Text = "Fehler beim Speichern: $($_.Exception.Message)"
                    }
                }
            })
        }
        
        # Clear Log Button
        if ($GUI.ContainsKey("btnClearLogSearch") -and $null -ne $GUI["btnClearLogSearch"]) {
            $GUI["btnClearLogSearch"].Add_Click({
                $GUI["txtLogSearch"].Text = ""
                Update-LogView -GUI $GUI
                $GUI["txtLogStatus"].Text = "Suche zurückgesetzt"
            })
        }
        
        # Close Button
        if ($GUI.ContainsKey("btnClose") -and $null -ne $GUI["btnClose"]) {
            $GUI["btnClose"].Add_Click({
                # Stop all timers if they exist
                if ($script:timers) {
                    foreach ($timer in $script:timers.Values) {
                        if ($timer -and $timer.Enabled) {
                            $timer.Stop()
                        }
                    }
                }
                
                $Window.Close()
                Write-Host "Anwendung wird geschlossen"
            })
        }
        
        Write-Host "Event-Handler erfolgreich registriert"
        return $true
    }
    catch {
        Write-Host "Fehler beim Registrieren der Event-Handler: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

function global:Update-ServiceMonitorUI {
    param (
        $GUI
    )
    
    try {
        # Dienste abrufen
        $services = Get-ServiceStatus
        
        if (-not $services) {
            Write-Host "Keine Dienste gefunden" "WARNING"
            return $false
        }
        
        # Service-Map für die UI-Elemente
        $serviceMap = @{
            "wuauserv" = @{
                Status = "txtStatusWUAU"
                StartType = "txtStartTypeWUAU"
                StartButton = "btnStartWUAU"
                StopButton = "btnStopWUAU"
                RestartButton = "btnRestartWUAU"
            }
            "BITS" = @{
                Status = "txtStatusBITS"
                StartType = "txtStartTypeBITS"
                StartButton = "btnStartBITS"
                StopButton = "btnStopBITS"
                RestartButton = "btnRestartBITS"
            }
            "cryptsvc" = @{
                Status = "txtStatusCRYPT"
                StartType = "txtStartTypeCRYPT"
                StartButton = "btnStartCRYPT"
                StopButton = "btnStopCRYPT"
                RestartButton = "btnRestartCRYPT"
            }
            "TrustedInstaller" = @{
                Status = "txtStatusTRUST"
                StartType = "txtStartTypeTRUST"
                StartButton = "btnStartTRUST"
                StopButton = "btnStopTRUST"
                RestartButton = "btnRestartTRUST"
            }
            "DoSvc" = @{
                Status = "txtStatusDOSVC"
                StartType = "txtStartTypeDOSVC"
                StartButton = "btnStartDOSVC"
                StopButton = "btnStopDOSVC"
                RestartButton = "btnRestartDOSVC"
            }
            "UsoSvc" = @{
                Status = "txtStatusUSO"
                StartType = "txtStartTypeUSO"
                StartButton = "btnStartUSO"
                StopButton = "btnStopUSO"
                RestartButton = "btnRestartUSO"
            }
            "WaaSMedicSvc" = @{
                Status = "txtStatusMEDIC"
                StartType = "txtStartTypeMEDIC"
                StartButton = "btnStartMEDIC"
                StopButton = "btnStopMEDIC"
                RestartButton = "btnRestartMEDIC"
            }
        }
        
        # Für jeden Dienst die UI aktualisieren
        foreach ($service in $services) {
            $serviceName = $service.Name
            
            if ($serviceMap.ContainsKey($serviceName)) {
                $uiElements = $serviceMap[$serviceName]
                $isRunning = $service.Status -eq "Running"
                
                # Status-Text setzen
                if ($GUI.ContainsKey($uiElements.Status)) {
                    $GUI[$uiElements.Status].Text = $service.Status.ToString()
                    
                    # Färbung je nach Status
                    if ($isRunning) {
                        $GUI[$uiElements.Status].Foreground = [System.Windows.Media.Brushes]::Green
                    } else {
                        $GUI[$uiElements.Status].Foreground = [System.Windows.Media.Brushes]::Red
                    }
                }
                
                # StartType-Text setzen
                if ($GUI.ContainsKey($uiElements.StartType)) {
                    $GUI[$uiElements.StartType].Text = $service.StartType.ToString()
                }
                
                # Buttons aktivieren/deaktivieren je nach Status
                if ($GUI.ContainsKey($uiElements.StartButton)) {
                    $GUI[$uiElements.StartButton].IsEnabled = (-not $isRunning)
                }
                
                if ($GUI.ContainsKey($uiElements.StopButton)) {
                    $GUI[$uiElements.StopButton].IsEnabled = $isRunning
                }
                
                if ($GUI.ContainsKey($uiElements.RestartButton)) {
                    $GUI[$uiElements.RestartButton].IsEnabled = $isRunning
                }
            }
        }
        
        # Statusinfo aktualisieren
        if ($GUI.ContainsKey("txtServiceInfo")) {
            $runningCount = ($services | Where-Object { $_.Status -eq "Running" }).Count
            $stoppedCount = ($services | Where-Object { $_.Status -eq "Stopped" }).Count
            
            $GUI["txtServiceInfo"].Text = "Status: $runningCount Dienste laufen, $stoppedCount Dienste gestoppt. Letzte Aktualisierung: $(Get-Date -Format 'HH:mm:ss')"
        }
        
        Write-Host "Dienste-Status erfolgreich aktualisiert: $($services.Count) Dienste" "SUCCESS"
        return $true
    }
    catch {
        Write-Host "Fehler bei der Aktualisierung des Dienste-Status: $($_.Exception.Message)" "ERROR"
        return $false
    }
}


function Register-ServiceButtonEvents {
    param (
        $GUI
    )
    
    try {
        # Service-Map für Button-Events
        $serviceButtons = @(
            @{ ServiceName = "wuauserv"; StartButton = "btnStartWUAU"; StopButton = "btnStopWUAU"; RestartButton = "btnRestartWUAU" },
            @{ ServiceName = "BITS"; StartButton = "btnStartBITS"; StopButton = "btnStopBITS"; RestartButton = "btnRestartBITS" },
            @{ ServiceName = "cryptsvc"; StartButton = "btnStartCRYPT"; StopButton = "btnStopCRYPT"; RestartButton = "btnRestartCRYPT" },
            @{ ServiceName = "TrustedInstaller"; StartButton = "btnStartTRUST"; StopButton = "btnStopTRUST"; RestartButton = "btnRestartTRUST" },
            @{ ServiceName = "DoSvc"; StartButton = "btnStartDOSVC"; StopButton = "btnStopDOSVC"; RestartButton = "btnRestartDOSVC" },
            @{ ServiceName = "UsoSvc"; StartButton = "btnStartUSO"; StopButton = "btnStopUSO"; RestartButton = "btnRestartUSO" },
            @{ ServiceName = "WaaSMedicSvc"; StartButton = "btnStartMEDIC"; StopButton = "btnStopMEDIC"; RestartButton = "btnRestartMEDIC" }
        )
        
        foreach ($serviceButton in $serviceButtons) {
            $serviceName = $serviceButton.ServiceName
            
            # Start Button
            if ($GUI.ContainsKey($serviceButton.StartButton)) {
                $GUI[$serviceButton.StartButton].Add_Click({
                    $svcName = $serviceName
                    Set-ServiceOperation -ServiceName $svcName -Operation "Start"
                    Update-ServiceMonitorUI -GUI $script:GUI
                }.GetNewClosure())
            }
            
            # Stop Button
            if ($GUI.ContainsKey($serviceButton.StopButton)) {
                $GUI[$serviceButton.StopButton].Add_Click({
                    $svcName = $serviceName
                    Set-ServiceOperation -ServiceName $svcName -Operation "Stop"
                    Update-ServiceMonitorUI -GUI $script:GUI
                }.GetNewClosure())
            }
            
            # Restart Button
            if ($GUI.ContainsKey($serviceButton.RestartButton)) {
                $GUI[$serviceButton.RestartButton].Add_Click({
                    $svcName = $serviceName
                    Set-ServiceOperation -ServiceName $svcName -Operation "Restart"
                    Update-ServiceMonitorUI -GUI $script:GUI
                }.GetNewClosure())
            }
        }
        
        Write-Host "Service-Button-Events erfolgreich registriert" "SUCCESS"
        return $true
    }
    catch {
        Write-Host "Fehler beim Registrieren der Service-Button-Events: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

# Funktion zum Aktualisieren der Systemressourcen in der GUI
function Update-SystemResources {
    param (
        $GUI
    )
    
    try {
        # Update-Statistiken anzeigen
        if ($GUI.ContainsKey("chartUpdateStats")) {
            try {
                # Canvas leeren
                $GUI["chartUpdateStats"].Children.Clear()
                
                # Update-Statistiken abrufen
                $updateSession = New-Object -ComObject Microsoft.Update.Session
                $updateSearcher = $updateSession.CreateUpdateSearcher()
                
                # Gesamtanzahl der Updates im Verlauf abrufen
                $totalHistoryCount = $updateSearcher.GetTotalHistoryCount()
                
                # Update-Verlauf für statistiche Auswertung abrufen (max. 1000 Einträge)
                $historyCount = [Math]::Min($totalHistoryCount, 1000)
                $history = $updateSearcher.QueryHistory(0, $historyCount)
                
                # Statistiken berechnen
                $totalCount = $historyCount # Gesamtzahl der geprüften Updates
                $successCount = 0           # Erfolgreiche Updates
                $failedCount = 0            # Fehlgeschlagene Updates
                $pendingCount = 0           # Ausstehende/unvollständige Updates
                $otherCount = 0             # Andere Statusmeldungen
                
                foreach ($entry in $history) {
                    switch ($entry.ResultCode) {
                    2 { $successCount++ } # Erfolgreich
                    4 { $failedCount++ }  # Fehlgeschlagen
                    3 { $pendingCount++ } # Unvollständig
                    default { $otherCount++ } # Andere (nicht gestartet, abgebrochen, etc.)
                    }
                }
                
                # Konfiguration für die Grafik
                $canvasWidth = $GUI["chartUpdateStats"].ActualWidth
                $canvasHeight = $GUI["chartUpdateStats"].ActualHeight
                
                # Fallback für Größen, wenn ActualWidth/Height noch 0 sind
                if ($canvasWidth -lt 10) { $canvasWidth = 300 }
                if ($canvasHeight -lt 10) { $canvasHeight = 150 }
                
                # Abstand zwischen Elementen
                $margin = 10
                $barHeight = 20
                $labelWidth = 100
                $valueWidth = 40
                $top = $margin
                $maxBarWidth = $canvasWidth - (2 * $margin) - $labelWidth - $valueWidth
                
                # Hintergrund für die Statistik
                $background = New-Object System.Windows.Shapes.Rectangle
                $background.Width = $canvasWidth - ($margin * 2)
                $background.Height = $canvasHeight - ($margin * 2)
                $background.Fill = [System.Windows.Media.Brushes]::WhiteSmoke
                $background.Stroke = [System.Windows.Media.Brushes]::LightGray
                $background.StrokeThickness = 1
                [System.Windows.Controls.Canvas]::SetLeft($background, $margin)
                [System.Windows.Controls.Canvas]::SetTop($background, $margin)
                $GUI["chartUpdateStats"].Children.Add($background)
                
                # Titel
                $titleText = New-Object System.Windows.Controls.TextBlock
                $titleText.Text = "Windows Update Statistik"
                $titleText.FontWeight = "Bold"
                $titleText.Foreground = [System.Windows.Media.Brushes]::Black
                [System.Windows.Controls.Canvas]::SetLeft($titleText, $margin + 5)
                [System.Windows.Controls.Canvas]::SetTop($titleText, $top + 5)
                $GUI["chartUpdateStats"].Children.Add($titleText)
                $top += 30
                
                # Gesamtzahl
                $totalText = New-Object System.Windows.Controls.TextBlock
                $totalText.Text = "Gesamtzahl der Updates:"
                $totalText.Foreground = [System.Windows.Media.Brushes]::Black
                [System.Windows.Controls.Canvas]::SetLeft($totalText, $margin + 5)
                [System.Windows.Controls.Canvas]::SetTop($totalText, $top)
                $GUI["chartUpdateStats"].Children.Add($totalText)
                
                $totalValueText = New-Object System.Windows.Controls.TextBlock
                $totalValueText.Text = $totalHistoryCount.ToString()
                $totalValueText.FontWeight = "Bold"
                $totalValueText.Foreground = [System.Windows.Media.Brushes]::Black
                [System.Windows.Controls.Canvas]::SetLeft($totalValueText, $canvasWidth - $margin - $valueWidth - 30)
                [System.Windows.Controls.Canvas]::SetTop($totalValueText, $top)
                $GUI["chartUpdateStats"].Children.Add($totalValueText)
                $top += 25
                
                # Erfolgreiche Updates
                $successBar = New-Object System.Windows.Shapes.Rectangle
                $successWidth = if ($totalCount -gt 0) { ($successCount / $totalCount) * $maxBarWidth } else { 0 }
                $successBar.Width = [Math]::Max(0, $successWidth)
                $successBar.Height = $barHeight
                $successBar.Fill = [System.Windows.Media.Brushes]::Green
                [System.Windows.Controls.Canvas]::SetLeft($successBar, $margin + $labelWidth)
                [System.Windows.Controls.Canvas]::SetTop($successBar, $top)
                $GUI["chartUpdateStats"].Children.Add($successBar)
                
                $successText = New-Object System.Windows.Controls.TextBlock
                $successText.Text = "Erfolgreich:"
                $successText.Foreground = [System.Windows.Media.Brushes]::Black
                [System.Windows.Controls.Canvas]::SetLeft($successText, $margin + 5)
                [System.Windows.Controls.Canvas]::SetTop($successText, $top + 2)
                $GUI["chartUpdateStats"].Children.Add($successText)
                
                $successValueText = New-Object System.Windows.Controls.TextBlock
                $successValueText.Text = "$successCount"
                $successValueText.Foreground = if ($successBar.Width -gt 40) { [System.Windows.Media.Brushes]::White } else { [System.Windows.Media.Brushes]::Black }
                $successValueText.FontWeight = "Bold"
                $successValuePosition = if ($successBar.Width -gt 40) {
                    $margin + $labelWidth + 5
                } else {
                    $margin + $labelWidth + $successBar.Width + 5
                }
                [System.Windows.Controls.Canvas]::SetLeft($successValueText, $successValuePosition)
                [System.Windows.Controls.Canvas]::SetTop($successValueText, $top + 2)
                $GUI["chartUpdateStats"].Children.Add($successValueText)
                $top += $barHeight + 5
                
                # Fehlgeschlagene Updates
                $failedBar = New-Object System.Windows.Shapes.Rectangle
                $failedWidth = if ($totalCount -gt 0) { ($failedCount / $totalCount) * $maxBarWidth } else { 0 }
                $failedBar.Width = [Math]::Max(0, $failedWidth)
                $failedBar.Height = $barHeight
                $failedBar.Fill = [System.Windows.Media.Brushes]::Red
                [System.Windows.Controls.Canvas]::SetLeft($failedBar, $margin + $labelWidth)
                [System.Windows.Controls.Canvas]::SetTop($failedBar, $top)
                $GUI["chartUpdateStats"].Children.Add($failedBar)
                
                $failedText = New-Object System.Windows.Controls.TextBlock
                $failedText.Text = "Fehlgeschlagen:"
                $failedText.Foreground = [System.Windows.Media.Brushes]::Black
                [System.Windows.Controls.Canvas]::SetLeft($failedText, $margin + 5)
                [System.Windows.Controls.Canvas]::SetTop($failedText, $top + 2)
                $GUI["chartUpdateStats"].Children.Add($failedText)
                
                $failedValueText = New-Object System.Windows.Controls.TextBlock
                $failedValueText.Text = "$failedCount"
                $failedValueText.Foreground = if ($failedBar.Width -gt 40) { [System.Windows.Media.Brushes]::White } else { [System.Windows.Media.Brushes]::Black }
                $failedValueText.FontWeight = "Bold"
                $failedValuePosition = if ($failedBar.Width -gt 40) {
                    $margin + $labelWidth + 5
                } else {
                    $margin + $labelWidth + $failedBar.Width + 5
                }
                [System.Windows.Controls.Canvas]::SetLeft($failedValueText, $failedValuePosition)
                [System.Windows.Controls.Canvas]::SetTop($failedValueText, $top + 2)
                $GUI["chartUpdateStats"].Children.Add($failedValueText)
                $top += $barHeight + 5
                
                # Ausstehende Updates
                $pendingBar = New-Object System.Windows.Shapes.Rectangle
                $pendingWidth = if ($totalCount -gt 0) { ($pendingCount / $totalCount) * $maxBarWidth } else { 0 }
                $pendingBar.Width = [Math]::Max(0, $pendingWidth)
                $pendingBar.Height = $barHeight
                $pendingBar.Fill = [System.Windows.Media.Brushes]::Orange
                [System.Windows.Controls.Canvas]::SetLeft($pendingBar, $margin + $labelWidth)
                [System.Windows.Controls.Canvas]::SetTop($pendingBar, $top)
                $GUI["chartUpdateStats"].Children.Add($pendingBar)
                
                $pendingText = New-Object System.Windows.Controls.TextBlock
                $pendingText.Text = "Unvollständig:"
                $pendingText.Foreground = [System.Windows.Media.Brushes]::Black
                [System.Windows.Controls.Canvas]::SetLeft($pendingText, $margin + 5)
                [System.Windows.Controls.Canvas]::SetTop($pendingText, $top + 2)
                $GUI["chartUpdateStats"].Children.Add($pendingText)
                
                $pendingValueText = New-Object System.Windows.Controls.TextBlock
                $pendingValueText.Text = "$pendingCount"
                $pendingValueText.Foreground = if ($pendingBar.Width -gt 40) { [System.Windows.Media.Brushes]::White } else { [System.Windows.Media.Brushes]::Black }
                $pendingValueText.FontWeight = "Bold"
                $pendingValuePosition = if ($pendingBar.Width -gt 40) {
                    $margin + $labelWidth + 5
                } else {
                    $margin + $labelWidth + $pendingBar.Width + 5
                }
                [System.Windows.Controls.Canvas]::SetLeft($pendingValueText, $pendingValuePosition)
                [System.Windows.Controls.Canvas]::SetTop($pendingValueText, $top + 2)
                $GUI["chartUpdateStats"].Children.Add($pendingValueText)
                
                # Andere Updates (optional hinzufügen, wenn Platz ist)
                if ($otherCount -gt 0 -and ($top + $barHeight + 5) -lt ($canvasHeight - 2 * $margin)) {
                    $top += $barHeight + 5
                    
                    $otherBar = New-Object System.Windows.Shapes.Rectangle
                    $otherWidth = if ($totalCount -gt 0) { ($otherCount / $totalCount) * $maxBarWidth } else { 0 }
                    $otherBar.Width = [Math]::Max(0, $otherWidth)
                    $otherBar.Height = $barHeight
                    $otherBar.Fill = [System.Windows.Media.Brushes]::Gray
                    [System.Windows.Controls.Canvas]::SetLeft($otherBar, $margin + $labelWidth)
                    [System.Windows.Controls.Canvas]::SetTop($otherBar, $top)
                    $GUI["chartUpdateStats"].Children.Add($otherBar)
                    
                    $otherText = New-Object System.Windows.Controls.TextBlock
                    $otherText.Text = "Sonstige:"
                    $otherText.Foreground = [System.Windows.Media.Brushes]::Black
                    [System.Windows.Controls.Canvas]::SetLeft($otherText, $margin + 5)
                    [System.Windows.Controls.Canvas]::SetTop($otherText, $top + 2)
                    $GUI["chartUpdateStats"].Children.Add($otherText)
                    
                    $otherValueText = New-Object System.Windows.Controls.TextBlock
                    $otherValueText.Text = "$otherCount"
                    $otherValueText.Foreground = if ($otherBar.Width -gt 40) { [System.Windows.Media.Brushes]::White } else { [System.Windows.Media.Brushes]::Black }
                    $otherValueText.FontWeight = "Bold"
                    $otherValuePosition = if ($otherBar.Width -gt 40) {
                        $margin + $labelWidth + 5
                    } else {
                        $margin + $labelWidth + $otherBar.Width + 5
                    }
                    [System.Windows.Controls.Canvas]::SetLeft($otherValueText, $otherValuePosition)
                    [System.Windows.Controls.Canvas]::SetTop($otherValueText, $top + 2)
                    $GUI["chartUpdateStats"].Children.Add($otherValueText)
                }
                
                Write-Log "Update-Statistiken erfolgreich aktualisiert" "INFO"
            }
            catch {
                Write-Log "Fehler beim Anzeigen der Update-Statistiken: $($_.Exception.Message)" "ERROR"
            }
        }
        
        # Statische Systeminformationen anzeigen
        if ($GUI.ContainsKey("txtProcessorInfo")) {
            $cpuInfo = Get-CimInstance -ClassName Win32_Processor | Select-Object -First 1
            $GUI["txtProcessorInfo"].Text = "$($cpuInfo.Name) ($($cpuInfo.NumberOfCores) Kerne)"
        }
        
        if ($GUI.ContainsKey("txtMemoryInfo")) {
            $osInfo = Get-CimInstance -ClassName Win32_OperatingSystem
            $totalRAM = [math]::Round($osInfo.TotalVisibleMemorySize / 1MB, 2)
            $GUI["txtMemoryInfo"].Text = "Gesamt: $totalRAM GB" 
        }
        
        if ($GUI.ContainsKey("txtOSInfo")) {
            $osInfo = Get-CimInstance -ClassName Win32_OperatingSystem
            $GUI["txtOSInfo"].Text = "$($osInfo.Caption) ($($osInfo.Version), Build $($osInfo.BuildNumber))"
        }
        
        if ($GUI.ContainsKey("txtDiskInfo")) {
            $diskInfo = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType=3"
            $diskText = ""
            foreach ($disk in $diskInfo) {
                $totalGB = [math]::Round($disk.Size / 1GB, 2)
                $freeGB = [math]::Round($disk.FreeSpace / 1GB, 2)
                $diskText += "Laufwerk $($disk.DeviceID): $totalGB GB gesamt, $freeGB GB frei`r`n"
            }
            $GUI["txtDiskInfo"].Text = $diskText.TrimEnd("`r`n")
        }
        
        if ($GUI.ContainsKey("txtUpdateInfo")) {
            $updateSession = New-Object -ComObject Microsoft.Update.Session
            $updateSearcher = $updateSession.CreateUpdateSearcher()
            try {
                $pendingCount = $updateSearcher.GetTotalHistoryCount()
                $lastUpdate = $updateSearcher.QueryHistory(0, 1) | Select-Object -First 1
                
                if ($lastUpdate) {
                    $GUI["txtUpdateInfo"].Text = "Letzte Installation: $($lastUpdate.Date), KB$($lastUpdate.Title -replace '^.*\((KB\d+)\).*$', '$1')"
                } else {
                    $GUI["txtUpdateInfo"].Text = "Keine Update-Informationen verfügbar."
                }
            } catch {
                $GUI["txtUpdateInfo"].Text = "Fehler beim Abfragen der Update-Informationen."
            }
        }
        
        if ($GUI.ContainsKey("txtNetworkInfo")) {
            $network = Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -Filter "IPEnabled=True" | Select-Object -First 1
            if ($network) {
                $GUI["txtNetworkInfo"].Text = "Adapter: $($network.Description)`r`nIP: $($network.IPAddress[0])`r`nGateway: $($network.DefaultIPGateway[0])"
            } else {
                $GUI["txtNetworkInfo"].Text = "Keine Netzwerkadapter gefunden."
            }
        }
        
        # Letzte Aktualisierung anzeigen
        if ($GUI.ContainsKey("txtLastRefresh")) {
            $GUI["txtLastRefresh"].Text = "Letzte Aktualisierung: " + (Get-Date -Format 'HH:mm:ss')
        }
        
        # Deaktiviere alle Timer zur Systemressourcenüberwachung
        if ($script:timers.ContainsKey("ResourceTimer") -and $script:timers["ResourceTimer"]) {
            $script:timers["ResourceTimer"].Stop()
            Write-Log "Ressourcen-Aktualisierungstimer gestoppt" "INFO"
        }
        
        Write-Log "Systeminformationen erfolgreich aktualisiert" "INFO"
        return $true
    }
    catch {
        Write-Log "Fehler bei der Aktualisierung der Systeminformationen: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

# Update-LogView Funktion
function Update-LogView {
    param (
        $GUI
    )
    
    try {
        if (-not ($GUI.ContainsKey("cmbLogType") -and $GUI.ContainsKey("cmbLogLines") -and $GUI.ContainsKey("txtLogContent"))) {
            Write-Host "Log-Viewer UI-Elemente nicht gefunden" "WARNING"
            return $false
        }
        
        # Bestimmen, welches Log angezeigt werden soll
        $logName = "WindowsUpdate"
        switch ($GUI["cmbLogType"].SelectedIndex) {
            0 { $logName = "WindowsUpdate" }
            1 { $logName = "CBS" }
            2 { $logName = "DISM" }
            3 { $logName = "SetupAPI" }
            default { $logName = "WindowsUpdate" }
        }
        
        # Bestimmen der anzuzeigenden Zeilen
        $tailLines = 100
        switch ($GUI["cmbLogLines"].SelectedIndex) {
            0 { $tailLines = 100 }
            1 { $tailLines = 500 }
            2 { $tailLines = 1000 }
            3 { $tailLines = 0 } # Alle Zeilen
            default { $tailLines = 100 }
        }
        
        # Suchbegriff abrufen
        $searchString = ""
        if ($GUI.ContainsKey("txtLogSearch")) {
            $searchString = $GUI["txtLogSearch"].Text.Trim()
        }
        
        # Log-Dateien laden (Platzhalter-Funktion, die eigentliche Implementierung sollte hier eingesetzt werden)
        # Hier sollte die Funktion Get-UpdateLogContent implementiert sein
        $logContent = Get-UpdateLogContent -LogType $logName -TailLines $tailLines -SearchString $searchString
        
        # In TextBox anzeigen
        $GUI["txtLogContent"].Text = $logContent -join "`r`n"
        
        # Status aktualisieren
        if ($GUI.ContainsKey("txtLogStatus")) {
            $GUI["txtLogStatus"].Text = "Log-Typ: $logName, Zeilen: " + $(if ($tailLines -eq 0) { "Alle" } else { $tailLines }) + ". Geladen: $(Get-Date -Format 'HH:mm:ss')"
        }
        
        Write-Host "Log-Ansicht für '$logName' aktualisiert"
        return $true
    }
    catch {
        Write-Host "Fehler beim Aktualisieren der Log-Ansicht: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

# Funktion zum Laden und Anzeigen verschiedener Windows Update Logs
function Get-UpdateLogContent {
    param (
        [Parameter(Mandatory=$true)]
        [ValidateSet("WindowsUpdate", "CBS", "DISM", "SetupAPI")]
        [string]$LogType,
        
        [Parameter(Mandatory=$false)]
        [int]$TailLines = 100,
        
        [Parameter(Mandatory=$false)]
        [string]$SearchString = ""
    )
    
    try {
        # Pfad zum Log basierend auf Typ ermitteln
        $logPath = switch ($LogType) {
            "WindowsUpdate" { "$env:SystemRoot\WindowsUpdate.log" }
            "CBS" { "$env:SystemRoot\Logs\CBS\CBS.log" }
            "DISM" { "$env:SystemRoot\Logs\DISM\DISM.log" }
            "SetupAPI" { "$env:SystemRoot\inf\setupapi.dev.log" }
            default { "$env:SystemRoot\WindowsUpdate.log" }
        }
        
        # Prüfen, ob das Log existiert
        if (-not (Test-Path -Path $logPath)) {
            # Bei Windows 10/11 ist das WindowsUpdate.log nicht direkt lesbar
            if ($LogType -eq "WindowsUpdate" -and (Get-CimInstance -ClassName Win32_OperatingSystem).Version -ge "10.0") {
                Write-Host "Windows 10+ erkannt, nutze PowerShell-Befehl zum Abrufen des Windows Update-Logs" "INFO"
                
                # Temporären Pfad für das konvertierte Log erstellen
                $tempFile = [System.IO.Path]::GetTempFileName()
                
                # Get-WindowsUpdateLog ausführen, um das ETL-Log zu konvertieren
                try {
                    # Versuchen, Get-WindowsUpdateLog zu verwenden
                    Get-WindowsUpdateLog -LogPath $tempFile -ErrorAction Stop | Out-Null
                    
                    if (Test-Path -Path $tempFile) {
                        $logPath = $tempFile
                    } else {
                        throw "Get-WindowsUpdateLog konnte keine lesbare Logdatei erzeugen."
                    }
                }
                catch {
                    # Fallback-Meldung, wenn Get-WindowsUpdateLog nicht funktioniert
                    Write-Host "Get-WindowsUpdateLog fehlgeschlagen: $($_.Exception.Message)" "ERROR"
                    return @("Windows Update Log nicht verfügbar. Bitte führen Sie PowerShell als Administrator aus und versuchen Sie es erneut.", 
                             "Alternatives Kommando: Get-WindowsUpdateLog -LogPath $env:TEMP\WindowsUpdate.log")
                }
            } else {
                Write-Host "Log-Datei nicht gefunden: $logPath" "WARNING"
                return @("Log-Datei nicht gefunden: $logPath")
            }
        }
        
        # Log-Inhalt abrufen
        $logContent = if ($TailLines -eq 0) {
            # Alle Zeilen lesen
            Get-Content -Path $logPath -ErrorAction Stop
        } else {
            # Letzte X Zeilen lesen
            Get-Content -Path $logPath -Tail $TailLines -ErrorAction Stop
        }
        
        # Nach Suchbegriff filtern, falls vorhanden
        if (-not [string]::IsNullOrEmpty($SearchString)) {
            $logContent = $logContent | Where-Object { $_ -like "*$SearchString*" }
            
            if ($logContent.Count -eq 0) {
                return @("Keine Treffer für Suchbegriff '$SearchString' gefunden.")
            }
        }
        
        Write-Host "Log-Inhalt für '$LogType' erfolgreich abgerufen: $($logContent.Count) Zeilen" "SUCCESS"
        return $logContent
    }
    catch {
        Write-Host "Fehler beim Abrufen des Log-Inhalts für '$LogType': $($_.Exception.Message)" "ERROR"
        return @("Fehler beim Lesen der Log-Datei: $($_.Exception.Message)")
    }
    finally {
        # Temporäre Datei löschen, falls vorhanden
        if (($LogType -eq "WindowsUpdate") -and ($tempFile) -and (Test-Path -Path $tempFile)) {
            Remove-Item -Path $tempFile -Force -ErrorAction SilentlyContinue
        }
    }
}

# Windows Update Error Codes Database
function Get-WindowsUpdateErrorCodes {
    return @(
        [PSCustomObject]@{
            ErrorCode = "0x80072EE7"
            Description = "Servername oder Adresse konnte nicht aufgelöst werden (DNS-Problem). Dies tritt typischerweise bei falschen Netzwerkeinstellungen auf."
            Command = "ipconfig /flushdns"
            Documentation = "Microsoft Q&A: Fehler 0x80072EE7 Erklärung"
        },
        [PSCustomObject]@{
            ErrorCode = "0x8024402C"
            Description = "Der Proxyserver oder Zielservername kann nicht aufgelöst werden (ERROR_WINHTTP_NAME_NOT_RESOLVED). Häufig verursacht durch falsche Proxy- oder DNS-Einstellungen."
            Command = "netsh winhttp reset proxy"
            Documentation = "Proxy/Name resolution issues – ConfigMgr Troubleshooting"
        },
        [PSCustomObject]@{
            ErrorCode = "0x8024401B"
            Description = "HTTP-Status 407 – Proxy-Authentifizierung erforderlich. Der Windows Update-Client kann aufgrund einer Proxy-Einstellung nicht auf den Update-Server zugreifen."
            Command = "Kein spezifischer Befehl (Proxy-Einstellungen korrekt konfigurieren)"
            Documentation = "Microsoft Learn: WU_E_PT_HTTP_STATUS_PROXY_AUTH_REQ"
        },
        [PSCustomObject]@{
            ErrorCode = "0x80244018"
            Description = "HTTP-Status 403 – Zugriff verweigert (Forbidden). Der Update-Server lehnt die Anfrage ab (möglicherweise durch Sicherheitsfilter)."
            Command = "Kein spezifischer Befehl (Zugriffsberechtigungen/Firewall prüfen)"
            Documentation = "Windows Update Error 0x80244018 Beschreibung"
        },
        [PSCustomObject]@{
            ErrorCode = "0x80244019"
            Description = "HTTP-Status 404 – Objekt nicht gefunden. Die angeforderte Update-Datei ist auf dem Server (WSUS/Windows Update) nicht vorhanden."
            Command = "netsh winhttp import proxy source=ie"
            Documentation = "Microsoft Q&A: Lösungen für Fehler 0x80244019"
        },
        [PSCustomObject]@{
            ErrorCode = "0x80244022"
            Description = "HTTP-Status 503 – Dienst nicht verfügbar. Der Update-Dienst ist überlastet oder temporär nicht erreichbar."
            Command = "Kein spezifischer Befehl (Netzwerkverbindung und Serverstatus prüfen)"
            Documentation = "Microsoft Learn: HTTP 503 bei Windows Update"
        },
        [PSCustomObject]@{
            ErrorCode = "0x80072EFD / 0x80072EFE / 0x80D02002"
            Description = "Timeout beim Herstellen der Verbindung. Die Verbindung zum Update-Dienst brach ab oder dauerte zu lange, oft durch Firewall/Proxy verursacht."
            Command = "Kein spezifischer Befehl (Firewall/Proxy prüfen, Netzwerk-Monitoring)"
            Documentation = "Microsoft Learn: TIME_OUT_ERRORS (0x80072EFD etc.)"
        },
        [PSCustomObject]@{
            ErrorCode = "0x80072EE2"
            Description = "WININET_E_TIMEOUT; Zeitüberschreitung der Verbindung. Der Windows Update-Agent konnte keine Verbindung mit dem Update-Server (WSUS, ConfigMgr oder Microsoft) herstellen."
            Command = "Kein spezifischer Befehl (Netzwerkverbindung sicherstellen, Endpunkte erreichbar machen)"
            Documentation = "Microsoft Learn: WININET_E_TIMEOUT Fehler"
        },
        [PSCustomObject]@{
            ErrorCode = "0x80072F8F"
            Description = "WININET_E_DECODING_FAILED; TLS/SSL-Fehler. Die empfangenen Update-Daten konnten nicht decodiert werden, oft weil TLS 1.2 nicht aktiviert oder Uhrzeit/Datum falsch sind."
            Command = "Kein spezifischer Befehl (TLS 1.2 aktivieren bzw. Update KB3140245 installieren)"
            Documentation = "Microsoft Learn: TLS-Fehler 0x80072F8F"
        },
        [PSCustomObject]@{
            ErrorCode = "0x80200053"
            Description = "BG_E_VALIDATION_FAILED – Dateivalidierung fehlgeschlagen. Möglicherweise filtert eine Firewall Downloads, was zu fehlerhaften Antworten führt."
            Command = "Kein spezifischer Befehl (Download-Filter in Firewalls/Proxys entfernen)"
            Documentation = "Microsoft Learn: BG_E_VALIDATION_FAILED"
        },
        [PSCustomObject]@{
            ErrorCode = "0x80244007"
            Description = "WU_E_PT_SOAPCLIENT_SOAPFAULT – SOAP-Fehler beim Update-Scan. Windows kann das Authentifizierungs-Cookie für WSUS nicht erneuern."
            Command = "Kein spezifischer Befehl (Server-Einstellungen prüfen, ggf. WSUS-Client resetten)"
            Documentation = "Microsoft Support: Fehler 0x80244007 bei WSUS-Scan"
        },
        [PSCustomObject]@{
            ErrorCode = "0x8024D009"
            Description = "WU_E_SETUP_SKIP_UPDATE – Self-Update des Windows Update-Agent wurde übersprungen. Tritt auf, wenn der WSUS-Server den WUA nicht aktualisiert (SelfUpdate nicht funktional)."
            Command = "Kein spezifischer Befehl (WSUS SelfUpdate reparieren, Berechtigungen und virtuelles Verzeichnis prüfen)"
            Documentation = "Microsoft KB920659 – WSUS SelfUpdate Problem"
        },
        [PSCustomObject]@{
            ErrorCode = "0x8024402F"
            Description = "WU_E_PT_ECP_SUCCEEDED_WITH_ERRORS – Externe CAB-Verarbeitung mit Fehlern abgeschlossen. Häufig durch Webfilter-Software (z.B. Lightspeed Rocket) verursacht, die Updates blockiert."
            Command = "Kein spezifischer Befehl (Webfilter-Ausnahmen konfigurieren)"
            Documentation = "Microsoft Learn: 0x8024402F Beschreibung"
        },
        [PSCustomObject]@{
            ErrorCode = "0x8024A10A"
            Description = "USO_E_SERVICE_SHUTTING_DOWN – Windows Update-Dienst wurde aufgrund von Inaktivität beendet. Tritt auf, wenn das System während des Updatevorgangs längere Zeit untätig ist."
            Command = "Kein spezifischer Befehl (System während Updates aktiv halten)"
            Documentation = "Microsoft Learn: Fehler 0x8024A10A"
        },
        [PSCustomObject]@{
            ErrorCode = "0x80070422"
            Description = "ERROR_SERVICE_DISABLED – Der Windows Update-Dienst ist deaktiviert oder nicht gestartet."
            Command = "sc config wuauserv start= auto && sc start wuauserv"
            Documentation = "Microsoft Learn: Fehler 0x80070422"
        },
        [PSCustomObject]@{
            ErrorCode = "0x80240020"
            Description = "WU_E_NO_INTERACTIVE_USER – Kein Benutzer angemeldet. Die Installation erfordert einen interaktiven Benutzer (ggf. Neustart ausstehend)."
            Command = "Kein spezifischer Befehl (Benutzer anmelden und Update erneut starten)"
            Documentation = "Microsoft Learn: 0x80240020 Erklärung"
        },
        [PSCustomObject]@{
            ErrorCode = "0x80242014"
            Description = "WU_E_UH_POSTREBOOTSTILLPENDING – Update wartet auf Neustart. Bei einigen Updates ist ein Neustart erforderlich, um die Installation abzuschließen."
            Command = "Kein spezifischer Befehl (System neu starten, um ausstehende Updates zu installieren)"
            Documentation = "Microsoft Learn: 0x80242014 Erklärung"
        },
        [PSCustomObject]@{
            ErrorCode = "0x80070BC9"
            Description = "ERROR_FAIL_REBOOT_REQUIRED – Ausstehender Neustart erforderlich. Häufig blockiert eine GPO (TrustedInstaller auf Manuell) die Fortsetzung des Updates bis zum Neustart."
            Command = "dism /image:C:\ /cleanup-image /revertpendingactions"
            Documentation = "Microsoft Learn: Fehler 0x80070BC9 GPO-Issue"
        },
        [PSCustomObject]@{
            ErrorCode = "0x800706BE"
            Description = "RPC-Kommunikationsfehler – Remote Procedure Call schlug fehl. Tritt z.B. auf, wenn ein vorheriges kumulatives Update unvollständig installiert wurde."
            Command = "Kein spezifischer Befehl (defekte Update-Registrierungseinträge reparieren, siehe Doku)"
            Documentation = "Microsoft Learn: 0x800706BE Update-Fehler"
        },
        [PSCustomObject]@{
            ErrorCode = "0x80240022"
            Description = "WU_E_ALL_UPDATES_FAILED – Alle Updates konnten nicht installiert werden. Häufig blockiert Antivirensoftware den Zugriff auf bestimmte Ordner (z.B. SoftwareDistribution)."
            Command = "Kein spezifischer Befehl (Virenschutz prüfen oder temporär deaktivieren)"
            Documentation = "Microsoft Learn: WU_E_ALL_UPDATES_FAILED"
        },
        [PSCustomObject]@{
            ErrorCode = "0x80246017"
            Description = "WU_E_DM_UNAUTHORIZED_LOCAL_USER – Download wurde verweigert. Der lokale Benutzer hat nicht genügend Rechte, Updates herunterzuladen/zu installieren."
            Command = "Kein spezifischer Befehl (als Administrator anmelden und erneut versuchen)"
            Documentation = "Microsoft Learn: 0x80246017 Erklärung"
        },
        [PSCustomObject]@{
            ErrorCode = "0x8007000D"
            Description = "ERROR_INVALID_DATA – Ungültige oder beschädigte Update-Daten. Möglicherweise wurde ein Updatepaket fehlerhaft heruntergeladen."
            Command = "Kein spezifischer Befehl (Update erneut herunterladen und Installation neu starten)"
            Documentation = "Microsoft Learn: 0x8007000D Erklärung"
        },
        [PSCustomObject]@{
            ErrorCode = "0x80070490"
            Description = "ERROR_NOT_FOUND – Ein zum Update gehöriges Element (z.B. Treiber-Architekturwert in der Registry) wurde nicht gefunden."
            Command = "Kein spezifischer Befehl (fehlenden Wert in der Registry ergänzen – siehe Dokumentation)"
            Documentation = "Microsoft Learn: Fehler 0x80070490 bei Treiber-Update"
        },
        [PSCustomObject]@{
            ErrorCode = "0x80242006"
            Description = "WU_E_UH_INVALIDMETADATA – Ungültige Metadaten im Update. Die Update-Informationen sind fehlerhaft oder inkonsistent."
            Command = "Ren %systemroot%\SoftwareDistribution\DataStore DataStore.bak"
            Documentation = "Microsoft Learn: 0x80242006 Problembehebung"
        },
        [PSCustomObject]@{
            ErrorCode = "0x8024200D"
            Description = "SUS_E_UH_NEEDANOTHERDOWNLOAD – Es sind weitere Daten erforderlich. Teile des Updates fehlen (erneuter Download erforderlich)."
            Command = "Kein spezifischer Befehl (Update erneut herunterladen lassen, ggf. neuesten Servicing Stack installieren)"
            Documentation = "BornCity: Fehler 0x8024200D – - Need another download"
        },
        [PSCustomObject]@{
            ErrorCode = "0x80246007"
            Description = "SUS_E_DM_NOTDOWNLOADED – Das Update wurde nicht heruntergeladen. Der Download ist unvollständig oder wurde verworfen."
            Command = "bitsadmin /reset"
            Documentation = "WoSHub: Fehlercode 0x80246007 Bedeutung"
        },
        [PSCustomObject]@{
            ErrorCode = "0x800705B4"
            Description = "ERROR_TIMEOUT – Zeitlimit überschritten. Der Vorgang (Update-Suche oder -Installation) dauerte zu lange und wurde abgebrochen."
            Command = "Kein spezifischer Befehl (später erneut versuchen, ggf. Windows Update-Problembehandlung ausführen)"
            Documentation = "WoSHub: ERROR_TIMEOUT 0x800705B4"
        },
        [PSCustomObject]@{
            ErrorCode = "0x8024000B"
            Description = "WU_E_CALL_CANCELLED – Vorgang abgebrochen. Der Updatevorgang wurde vom Benutzer oder Dienst (z.B. dem Scheduler) abgebrochen."
            Command = "Kein spezifischer Befehl (Update erneut anstoßen und nicht manuell abbrechen)"
            Documentation = "Microsoft Learn: 0x8024000B Beschreibung"
        },
        [PSCustomObject]@{
            ErrorCode = "0x80070643"
            Description = "Installationsfehler 0x80070643 bedeutet -Fatal error during installation-. Häufig sind beschädigte Systemdateien oder .NET-Framework-Probleme die Ursache."
            Command = "sfc /scannow"
            Documentation = "MS Answers: 0x80070643 = fatal error"
        },
        [PSCustomObject]@{
            ErrorCode = "0x800B0109"
            Description = "TRUST_E_CERT_SIGNATURE – Eine Zertifikatkette endete in einem nicht vertrauenswürdigen Stammzertifikat. Das Update-Zertifikat ist nicht vertrauenswürdig (oder Update-Datei beschädigt)."
            Command = "sfc /scannow"
            Documentation = "PatchMyPC: Fehler 0x800B0109 (Signaturproblem)"
        },
        [PSCustomObject]@{
            ErrorCode = "0x80070005"
            Description = "E_ACCESSDENIED – Zugriff verweigert. Der Update-Prozess hat nicht die erforderlichen Berechtigungen auf eine Datei oder Registry-Schlüssel."
            Command = "Kein spezifischer Befehl (Berechtigungen der betroffenen Dateien/Registrypfade anpassen; als Administrator ausführen)"
            Documentation = "Microsoft Learn: Fehler 0x80070005 (Zugriff verweigert)"
        },
        [PSCustomObject]@{
            ErrorCode = "0x80070570"
            Description = "ERROR_FILE_CORRUPT – Datei oder Verzeichnis ist beschädigt. Der Komponentenspeicher (Component Store) enthält korrupte Dateien."
            Command = "DISM /Online /Cleanup-Image /RestoreHealth + sfc /scannow"
            Documentation = "Microsoft Learn: 0x80070570 (Komponentenspeicher)"
        },
        [PSCustomObject]@{
            ErrorCode = "0x80070003"
            Description = "ERROR_PATH_NOT_FOUND – Ein benötigter Pfad wurde nicht gefunden. Möglicherweise verweist ein Komponenten-Eintrag auf einen ungültigen Pfad."
            Command = "Kein spezifischer Befehl (Logs in %windir%\Logs\CBS\CBS.log nach dem Fehler durchsuchen und Pfadangaben prüfen)"
            Documentation = "Microsoft Learn: 0x80070003 Beschreibung"
        },
        [PSCustomObject]@{
            ErrorCode = "0x80070020"
            Description = "ERROR_SHARING_VIOLATION – Freigabekonflikt beim Zugriff auf eine Datei. Häufig durch Virenscanner oder Backup-Software verursacht, die Dateien sperren."
            Command = "Kein spezifischer Befehl (Clean Boot durchführen und mit Process Monitor den Prozess ermitteln, der die Datei blockiert)"
            Documentation = "Microsoft Learn: Fehler 0x80070020 (Sharing Violation)"
        },
        [PSCustomObject]@{
            ErrorCode = "0x80073701"
            Description = "ERROR_SXS_ASSEMBLY_MISSING – Referenzierte Assembly konnte nicht gefunden werden. Der Komponentenspeicher ist inkonsistent (teilweise installierte Komponente)."
            Command = "DISM /Online /Cleanup-Image /RestoreHealth + sfc /scannow"
            Documentation = "Microsoft Learn: 0x80073701 (Assembly missing)"
        },
        [PSCustomObject]@{
            ErrorCode = "0x8007371B"
            Description = "ERROR_SXS_TRANSACTION_CLOSURE_INCOMPLETE – Transaktion nicht vollständig. Mindestens ein erforderliches Element der Installation fehlt (Komponentenspeicher beschädigt)."
            Command = "DISM /Online /Cleanup-Image /RestoreHealth + sfc /scannow"
            Documentation = "Microsoft Learn: 0x8007371B (Transaktion unvollständig)"
        },
        [PSCustomObject]@{
            ErrorCode = "0x80073712"
            Description = "ERROR_SXS_COMPONENT_STORE_CORRUPT – Der Komponentenspeicher ist beschädigt. Windows kann benötigte Komponenten nicht mehr registrieren oder installieren."
            Command = "DISM /Online /Cleanup-Image /RestoreHealth"
            Documentation = "MS Community: Fehler 0x80073712 = Komponentenspeicher beschädigt"
        },
        [PSCustomObject]@{
            ErrorCode = "0x800F081F"
            Description = "CBS_E_SOURCE_MISSING – Quellpaket/Datei nicht gefunden. Für die Installation benötigte Komponentenquelle ist nicht verfügbar (häufig .NET-Framework-Problem)."
            Command = "DISM /Online /Cleanup-Image /RestoreHealth /Source:<Pfad>"
            Documentation = "Microsoft Learn: 0x800F081F (Quelle fehlt)"
        },
        [PSCustomObject]@{
            ErrorCode = "0x800F0831"
            Description = "CBS_E_STORE_CORRUPTION – Komponentenspeicher ist beschädigt. Ähnlich zu 0x80073712 – der Windows Component Store enthält Inkonsistenzen."
            Command = "DISM /Online /Cleanup-Image /RestoreHealth + sfc /scannow"
            Documentation = "Microsoft Learn: 0x800F0831 (Store corruption)"
        },
        [PSCustomObject]@{
            ErrorCode = "0x800F0821"
            Description = "CBS_E_ABORT – Vorgang vom Client abgebrochen (Timeout). Die Komponenteninstallation dauerte zu lange und wurde vom Watchdog abgebrochen."
            Command = "Kein spezifischer Befehl (mehr Ressourcen bereitstellen – CPU/RAM erhöhen; Patch KB4493473 oder höher installieren)"
            Documentation = "Microsoft Learn: 0x800F0821 (CBS Timeout)"
        },
        [PSCustomObject]@{
            ErrorCode = "0x800F0825"
            Description = "CBS_E_CANNOT_UNINSTALL – Paket kann nicht deinstalliert werden. Tritt meist bei Komponentenspeicher-Beschädigung auf (teilinstallierte Komponente)."
            Command = "DISM /Online /Cleanup-Image /RestoreHealth + sfc /scannow, danach Neustart"
            Documentation = "Microsoft Learn: 0x800F0825 (Komponentenspeicher)"
        },
        [PSCustomObject]@{
            ErrorCode = "0x800F0920"
            Description = "CBS_E_HANG_DETECTED – Hänger bei der Update-Verarbeitung erkannt. (Nachfolger von 0x800F0821) Die Installation scheint eingefroren."
            Command = "Kein spezifischer Befehl (VM-Ressourcen erhöhen, Timeout verlängern, Patch KB4493473+ installieren)"
            Documentation = "Microsoft Learn: 0x800F0920 (Hang detected)"
        },
        [PSCustomObject]@{
            ErrorCode = "0x800F0922"
            Description = "CBS_E_INSTALLERS_FAILED – Installationsroutine schlug fehl. Z.B. könnte ein kumulatives Update (Server 2016) nicht installiert werden, weil Lizenz-/Produktschlüssel nicht aktualisiert wurden."
            Command = "Kein spezifischer Befehl (Schreibberechtigungen für *C:\Windows\System32\spp* für 'Benutzer' und 'Netzwerkdienst' erteilen)"
            Documentation = "Microsoft Learn: 0x800F0922 (Installers failed)"
        },
        [PSCustomObject]@{
            ErrorCode = "0xC1900101"
            Description = "Allgemeiner Installationsfehler beim Upgrade (Rollback). Fehlercodes beginnend mit 0xC1900101 deuten meist auf Treiberprobleme hin, die zu einem Abbruch des Upgrades führen."
            Command = "Kein einzelner Befehl (Treiber aktualisieren, überflüssige Hardware entfernen, genügend Speicherplatz bereitstellen)"
            Documentation = "Microsoft Support: 0xC1900101 – Treiberfehler"
        },
        [PSCustomObject]@{
            ErrorCode = "0xC1900107"
            Description = "Ein vorheriger Update-/Upgrade-Vorgang ist noch nicht abgeschlossen (cleanup operation still pending). Ein Neustart des Systems ist erforderlich, bevor das Upgrade erneut versucht werden kann."
            Command = "Neustart des Systems durchführen; falls der Fehler bleibt, temporäre Windows Update-Dateien bereinigen (z.B. Datenträgerbereinigung)"
            Documentation = "Microsoft Support: Fehler 0xC1900107 (Cleanup pending)"
        },
        [PSCustomObject]@{
            ErrorCode = "0xC1900201"
            Description = "Systemreservierte Partition konnte nicht aktualisiert werden (unzureichender Platz). Das Upgrade schlägt fehl, weil die System Reserved-Partition zu klein ist oder nicht aktualisiert werden kann."
            Command = "Kein spezifischer Befehl (Speicherplatz auf der reservierten Partition vergrößern, z.B. nicht benötigte Sprachpakete entfernen oder Partition erweitern)"
            Documentation = "TechCommunity: Fehler 0xC1900201 – System Reserved Partition"
        }
    )
}

# Funktion zum Nachschlagen von Fehlercode-Informationen
function Get-UpdateErrorInfo {
    param (
        [string]$ErrorCode
    )
    
    try {
        # Fehlercodes laden
        $errorCodes = Get-WindowsUpdateErrorCodes
        
        # Fehlercode normalisieren (alles in Kleinbuchstaben, ohne Leerzeichen)
        $normalizedErrorCode = $ErrorCode.ToLower().Trim()
        
        # Nach Fehlercode suchen
        $errorInfo = $errorCodes | Where-Object { $_.ErrorCode.ToLower().Contains($normalizedErrorCode) }
        
        if ($errorInfo) {
            return $errorInfo
        }
        else {
            return $null
        }
    }
    catch {
        Write-Host "Fehler beim Nachschlagen des Fehlercodes '$ErrorCode': $($_.Exception.Message)" "ERROR"
        return $null
    }
}

# Funktion zum Anzeigen der Fehlercode-Informationen in der UI
function Show-ErrorCodeInfo {
    param (
        [string]$ErrorCode,
        $GUI
    )
    
    try {
        # Fehlermeldung löschen
        if ($GUI.ContainsKey("txtErrorMessage")) {
            $GUI["txtErrorMessage"].Text = ""
        }
        
        # Fehlercode-Informationen abrufen
        $errorInfo = Get-UpdateErrorInfo -ErrorCode $ErrorCode
        
        if ($null -eq $errorInfo) {
            if ($GUI.ContainsKey("txtErrorMessage")) {
                $GUI["txtErrorMessage"].Text = "Keine Informationen für Fehlercode '$ErrorCode' gefunden."
            }
            return $false
        }
        
        # Wenn mehrere Ergebnisse gefunden wurden, das erste nehmen
        if ($errorInfo -is [array]) {
            $errorInfo = $errorInfo[0]
        }
        
        # Fehlercode-Informationen anzeigen
        if ($GUI.ContainsKey("txtErrorCode")) {
            $GUI["txtErrorCode"].Text = $errorInfo.ErrorCode
        }
        
        if ($GUI.ContainsKey("txtErrorDescription")) {
            $GUI["txtErrorDescription"].Text = $errorInfo.Description
        }
        
        if ($GUI.ContainsKey("txtErrorCommand")) {
            $GUI["txtErrorCommand"].Text = $errorInfo.Command
        }
        
        if ($GUI.ContainsKey("txtErrorDocumentation")) {
            $GUI["txtErrorDocumentation"].Text = $errorInfo.Documentation
        }
        
        Write-Host "Fehlercode-Informationen für '$ErrorCode' angezeigt"
        return $true
    }
    catch {
        Write-Host "Fehler beim Anzeigen der Fehlercode-Informationen: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

# Funktion zum Abrufen einer Windows Update Diagnose
function Start-WindowsUpdateDiagnose {
    param (
        [Parameter(Mandatory=$false)]
        [string]$ErrorCode = ""
    )
    
    try {
        # Diagnose-Tool starten
        $result = Run-UpdateTroubleshooter
        
        if ($result) {
            Write-Host "Windows Update-Problembehandlung gestartet" "SUCCESS"
            return $true
        }
        else {
            Write-Host "Windows Update-Problembehandlung konnte nicht gestartet werden" "ERROR"
            return $false
        }
    }
    catch {
        Write-Host "Fehler beim Starten der Windows Update-Problembehandlung: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

# Main entry point for the GUI initialization
function Initialize-GUI {
    try {
        # Pfad zur XAML-Datei
        $xamlPath = Join-Path -Path $PSScriptRoot -ChildPath "GUI\SimpleGUI.xaml"
        
        # Prüfen, ob die XAML-Datei existiert
        if (-not (Test-Path -Path $xamlPath)) {
            Write-Host "XAML-Datei nicht gefunden unter: $xamlPath" "WARNING"
            
            # Alternative Pfade prüfen
            $alternativePaths = @(
                (Join-Path -Path $PSScriptRoot -ChildPath "SimpleGUI.xaml"),
                (Join-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -ChildPath "GUI\SimpleGUI.xaml"),
                (Get-ChildItem -Path $PSScriptRoot -Filter "*.xaml" -Recurse | Select-Object -First 1 -ExpandProperty FullName)
            )
            
            foreach ($altPath in $alternativePaths) {
                if ($altPath -and (Test-Path -Path $altPath)) {
                    Write-Host "Alternative XAML-Datei gefunden: $altPath" "INFO"
                    $xamlPath = $altPath
                    break
                }
            }
            
            # Wenn keine XAML gefunden wurde, nach Bestätigung erstellen
            if (-not (Test-Path -Path $xamlPath)) {
                $createNewFolder = $true
                $xamlDir = Join-Path -Path $PSScriptRoot -ChildPath "GUI"
                
                if (-not (Test-Path -Path $xamlDir)) {
                    try {
                        New-Item -Path $xamlDir -ItemType Directory -Force | Out-Null
                        Write-Host "Verzeichnis erstellt: $xamlDir" "INFO"
                    } catch {
                        $createNewFolder = $false
                    }
                }
                
                if ($createNewFolder) {
                    try {
                        # XAML-Inhalt aus dem bekannten Skript extrahieren und speichern
                        $scriptContent = Get-Content -Path $MyInvocation.MyCommand.Path -Raw
                        if ($scriptContent -match '<file>\r?\n```xaml.*?// filepath:.*?\r?\n(.*?)\r?\n```\r?\n</file>') {
                            $xamlContent = $Matches[1]
                            $xamlContent | Out-File -FilePath $xamlPath -Encoding utf8
                        } else {
                            throw "Konnte XAML-Inhalt nicht aus dem Skript extrahieren."
                        }
                    } catch {
                        
                        # Als letzten Ausweg interaktiven Dialog anzeigen
                        $userInput = [System.Windows.Forms.MessageBox]::Show(
                            "Die XAML-Datei (SimpleGUI.xaml) wurde nicht gefunden.`n`nMöchten Sie die Datei manuell auswählen?",
                            "XAML-Datei fehlt",
                            [System.Windows.Forms.MessageBoxButtons]::YesNo,
                            [System.Windows.Forms.MessageBoxIcon]::Question
                        )
                        
                        if ($userInput -eq 'Yes') {
                            $openDialog = New-Object Microsoft.Win32.OpenFileDialog
                            $openDialog.Filter = "XAML-Dateien (*.xaml)|*.xaml|Alle Dateien (*.*)|*.*"
                            $openDialog.Title = "XAML-Datei auswählen"
                            
                            if ($openDialog.ShowDialog()) {
                                $xamlPath = $openDialog.FileName
                                } else {
                                throw "Keine XAML-Datei ausgewählt. Beende Anwendung."
                            }
                        } else {
                            throw "Keine XAML-Datei verfügbar. Beende Anwendung."
                        }
                    }
                }
            }
        }
        
        # XAML laden und GUI-Objekte abrufen
        $uiData = Load-XamlGUI -XamlPath $xamlPath
        
        if ($null -eq $uiData) {
            throw "Fehler beim Laden der XAML-Datei"
        }
        
        $window = $uiData.Window
        $script:GUI = $uiData.GUI
        
        # Event-Handler registrieren
        $success = Register-AllEventHandlers -GUI $script:GUI -Window $window
        
        if (-not $success) {
            Write-Host "Warnung: Einige Event-Handler konnten nicht registriert werden" "WARNING"
        }
        
        # WUSA Settings Event-Handler registrieren
        $wusaSuccess = Register-WusaSettingsEventHandlers -GUI $script:GUI -Window $window
        
        if (-not $wusaSuccess) {
            Write-Host "Warnung: WUSA Settings Event-Handler konnten nicht registriert werden" "WARNING"
        }
        
        # Event-Handler für den Settings-Button
        if ($script:GUI.ContainsKey("btnWUSASettings") -and $null -ne $script:GUI["btnWUSASettings"]) {
            # Direkt Event-Handler hinzufügen (keine RemoveHandler erforderlich)
            $script:GUI["btnWUSASettings"].Add_Click({
                Show-WUSASettings -GUI $script:GUI
            })
             
            Write-Host "Event-Handler für btnWUSASettings (WUSA Settings Tab) registriert" "INFO"
        }
        
        # Initial den Dashboard-Tab zeigen
        $script:GUI["Dashboard"].Visibility = "Visible"
        
        # Überprüfen, ob es sich um ein gültiges Window-Objekt handelt
        if ($window -is [System.Windows.Window]) {
            Write-Host "Anzeigen des Fensters..."
            
            # Systemressourcenanzeige initial füllen
            $resources = Get-SystemUsage
            if ($resources) {
                try {
                    # Safely set initial values
                    if ($script:GUI.ContainsKey("progressCPU")) {
                        $script:GUI["progressCPU"].Dispatcher.Invoke([Action]{
                            $script:GUI["progressCPU"].Value = [double]$resources.CPUUsagePercent
                        })
                    }
                    
                    if ($script:GUI.ContainsKey("progressRAM")) {
                        $script:GUI["progressRAM"].Dispatcher.Invoke([Action]{
                            $script:GUI["progressRAM"].Value = [double]$resources.RAMUsagePercent
                        })
                    }
                    
                    if ($script:GUI.ContainsKey("progressDisk")) {
                        $script:GUI["progressDisk"].Dispatcher.Invoke([Action]{
                            $script:GUI["progressDisk"].Value = [double]$resources.DiskUsagePercent
                        })
                    }
                    
                    if ($script:GUI.ContainsKey("txtCPUValue")) {
                        $script:GUI["txtCPUValue"].Text = "$([Math]::Round($resources.CPUUsagePercent))%"
                    }
                    
                    if ($script:GUI.ContainsKey("txtRAMValue")) {
                        $script:GUI["txtRAMValue"].Text = "$([Math]::Round($resources.RAMUsagePercent))%"
                    }
                    
                    if ($script:GUI.ContainsKey("txtDiskValue")) {
                        $script:GUI["txtDiskValue"].Text = "$([Math]::Round($resources.DiskUsagePercent))%"
                    }
                } catch {
                    Write-Host "Fehler beim Initialisieren der Systemressourcenanzeige: $($_.Exception.Message)" "WARNING"
                }
            }
            
            # Window anzeigen
            $result = $window.ShowDialog()
        }
        else {
            throw "Das Fenster-Objekt ist kein gültiges System.Windows.Window-Objekt"
        }
    }
    catch {
        Write-Host "Kritischer Fehler: $($_.Exception.Message)" "ERROR"
        Write-Host "Stack Trace: $($_.Exception.StackTrace)" "ERROR"
        
        # Benutzerfreundliche Fehlermeldung anzeigen
        try {
            [System.Windows.MessageBox]::Show(
                "Ein Fehler ist aufgetreten: $($_.Exception.Message)`n`nBitte stellen Sie sicher, dass die XAML-Datei (SimpleGUI.xaml) vorhanden und korrekt formatiert ist.",
                "Fehler", 
                "OK", 
                "Error"
            )
        }
        catch {
            Write-Host "Fehler beim Anzeigen der Fehlermeldung: $($_.Exception.Message)"
            Write-Host "Grundlegender Fehler: Die XAML-Datei (GUI\SimpleGUI.xaml) konnte nicht geladen werden."
            Write-Host "Bitte stellen Sie sicher, dass diese Datei existiert und gültig ist."
        }
    }
}

# Funktion zum Generieren eines detaillierten Systemberichts
function Generate-SystemReport {
    param (
        [switch]$IncludeSystemInfo = $true,
        [switch]$IncludeUpdateConfig = $true,
        [switch]$IncludeUpdateHistory = $true,
        [switch]$IncludeServices = $true,
        [switch]$IncludeDiskSpace = $true,
        [switch]$IncludeEventLogs = $false,
        [switch]$IncludeNetworkConfig = $false,
        [ValidateSet("TXT", "HTML", "CSV")]
        [string]$Format = "TXT"
    )
    
    try {
        $reportContent = @()
        $reportContent += "=== WINDOWS UPDATE SYSTEMBERICHT ==="
        $reportContent += "Erstellt am: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        $reportContent += "Computername: $env:COMPUTERNAME"
        $reportContent += ""
        
        # Systeminformationen
        if ($IncludeSystemInfo) {
            $reportContent += "=== SYSTEMINFORMATIONEN ==="
            
            # Betriebssystem-Infos
            $osInfo = Get-CimInstance -ClassName Win32_OperatingSystem
            $reportContent += "Betriebssystem: $($osInfo.Caption) $($osInfo.Version)"
            $reportContent += "Build: $($osInfo.BuildNumber)"
            $reportContent += "Architektur: $($osInfo.OSArchitecture)"
            $reportContent += "Installation: $($osInfo.InstallDate)"
            $reportContent += "Letzter Boot: $($osInfo.LastBootUpTime)"
            
            # Prozessor-Infos
            $cpuInfo = Get-CimInstance -ClassName Win32_Processor -Property Name, NumberOfCores, NumberOfLogicalProcessors
            $reportContent += ""
            $reportContent += "Prozessor: $($cpuInfo.Name)"
            $reportContent += "Kerne: $($cpuInfo.NumberOfCores)"
            $reportContent += "Logische Prozessoren: $($cpuInfo.NumberOfLogicalProcessors)"
            
            # RAM-Infos
            $totalRAM = [math]::Round($osInfo.TotalVisibleMemorySize / 1MB, 2)
            $freeRAM = [math]::Round($osInfo.FreePhysicalMemory / 1MB, 2)
            $reportContent += ""
            $reportContent += "RAM Gesamt: $totalRAM GB"
            $reportContent += "RAM Frei: $freeRAM GB"
            $reportContent += ""
        }
        
        # Windows Update Konfiguration
        if ($IncludeUpdateConfig) {
            $reportContent += "=== WINDOWS UPDATE KONFIGURATION ==="
            
            # WSUS-Server und Update-Einstellungen abfragen
            $auSettings = Get-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" -ErrorAction SilentlyContinue
            $wuSettings = Get-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate" -ErrorAction SilentlyContinue
            
            if ($wuSettings) {
                $reportContent += "WSUS-Server: " + ($wuSettings.WUServer ?? "Nicht konfiguriert")
                $reportContent += "WSUS Status-Server: " + ($wuSettings.WUStatusServer ?? "Nicht konfiguriert")
            } else {
                $reportContent += "WSUS-Server: Nicht konfiguriert (Microsoft Update)"
            }
            
            if ($auSettings) {
                $auOptions = @{
                    0 = "Nicht konfiguriert"
                    1 = "Automatische Updates deaktiviert"
                    2 = "Benachrichtigung vor Download und Installation"
                    3 = "Automatischer Download und Benachrichtigung zur Installation"
                    4 = "Automatischer Download und geplante Installation"
                    5 = "Von lokaler Verwaltung festgelegt"
                }
                
                $reportContent += "AUOptions: " + ($auOptions[[int]$auSettings.AUOptions] ?? "Nicht konfiguriert")
                $reportContent += "ScheduledInstallDay: " + ($auSettings.ScheduledInstallDay ?? "Nicht konfiguriert")
                $reportContent += "ScheduledInstallTime: " + ($auSettings.ScheduledInstallTime ?? "Nicht konfiguriert")
                $reportContent += "UseWUServer: " + ($auSettings.UseWUServer ?? "Nicht konfiguriert")
            } else {
                $reportContent += "Automatische Updates: Standardeinstellung (Windows-Standard)"
            }
            $reportContent += ""
        }
        
        # Update-Verlauf
        if ($IncludeUpdateHistory) {
            $reportContent += "=== WINDOWS UPDATE VERLAUF ==="
            
            # Update-Verlauf abrufen (letzte 20 Einträge)
            $updateHistoryCount = 20
            $updateHistory = Get-HotFix | Sort-Object -Property InstalledOn -Descending | Select-Object -First $updateHistoryCount
            
            if ($updateHistory) {
                foreach ($update in $updateHistory) {
                    $reportContent += "KB$($update.HotFixID) - Installiert am: $($update.InstalledOn) - Beschreibung: $($update.Description)"
                }
            } else {
                $reportContent += "Keine Update-Verlaufsdaten verfügbar."
            }
            $reportContent += ""
        }
        
        # Dienste-Status
        if ($IncludeServices) {
            $reportContent += "=== WINDOWS UPDATE RELEVANTE DIENSTE ==="
            
            # Update-relevante Dienste abfragen
            $services = Get-Service -Name wuauserv, BITS, cryptsvc, TrustedInstaller, DoSvc, UsoSvc, WaaSMedicSvc -ErrorAction SilentlyContinue
            
            if ($services) {
                foreach ($service in $services) {
                    $startType = (Get-CimInstance -ClassName Win32_Service -Filter "Name='$($service.Name)'").StartMode
                    $reportContent += "$($service.DisplayName) - Status: $($service.Status) - Starttyp: $startType"
                }
            } else {
                $reportContent += "Keine Dienste-Informationen verfügbar."
            }
            $reportContent += ""
        }
        
        # Speicherplatz
        if ($IncludeDiskSpace) {
            $reportContent += "=== SPEICHERPLATZ ==="
            
            # Festplatten-Informationen abrufen
            $drives = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType=3"
            
            if ($drives) {
                foreach ($drive in $drives) {
                    $totalGB = [math]::Round($drive.Size / 1GB, 2)
                    $freeGB = [math]::Round($drive.FreeSpace / 1GB, 2)
                    $usedGB = $totalGB - $freeGB
                    $percentFree = [math]::Round(($drive.FreeSpace / $drive.Size) * 100, 2)
                    
                    $reportContent += "Laufwerk $($drive.DeviceID) - $($drive.VolumeName)"
                    $reportContent += "  Gesamt: $totalGB GB"
                    $reportContent += "  Frei: $freeGB GB ($percentFree% frei)"
                    $reportContent += "  Belegt: $usedGB GB"
                    $reportContent += ""
                }
            } else {
                $reportContent += "Keine Festplatten-Informationen verfügbar."
                $reportContent += ""
            }
        }
        
        # Ereignis-Logs
        if ($IncludeEventLogs) {
            $reportContent += "=== WINDOWS UPDATE EREIGNISSE (LETZTE 24 STUNDEN) ==="
            
            # Update-relevante Ereignis-Logs abrufen
            $yesterday = (Get-Date).AddDays(-1)
            $events = Get-WinEvent -LogName "System" -FilterXPath "*[System[(EventID=19 or EventID=20 or EventID=43) and TimeCreated[@SystemTime>='$($yesterday.ToUniversalTime().ToString("o"))']]]" -ErrorAction SilentlyContinue | Select-Object -First 20
            
            if ($events) {
                foreach ($event in $events) {
                    $reportContent += "Datum: $($event.TimeCreated) - ID: $($event.Id) - Quelle: $($event.ProviderName)"
                    $reportContent += "Nachricht: $($event.Message)"
                    $reportContent += ""
                }
            } else {
                $reportContent += "Keine relevanten Ereignisse in den letzten 24 Stunden gefunden."
                $reportContent += ""
            }
        }
        
        # Netzwerkkonfiguration
        if ($IncludeNetworkConfig) {
            $reportContent += "=== NETZWERKKONFIGURATION ==="
            
            # Netzwerk-Adapter-Konfiguration abrufen
            $adapters = Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -Filter "IPEnabled=True"
            
            if ($adapters) {
                foreach ($adapter in $adapters) {
                    $reportContent += "Adapter: $($adapter.Description)"
                    $reportContent += "  MAC-Adresse: $($adapter.MACAddress)"
                    $reportContent += "  IP-Adresse: $($adapter.IPAddress -join ', ')"
                    $reportContent += "  Subnetzmaske: $($adapter.IPSubnet -join ', ')"
                    $reportContent += "  Gateway: $($adapter.DefaultIPGateway -join ', ')"
                    $reportContent += "  DNS-Server: $($adapter.DNSServerSearchOrder -join ', ')"
                    $reportContent += "  DHCP aktiviert: $($adapter.DHCPEnabled)"
                    $reportContent += ""
                }
            } else {
                $reportContent += "Keine Netzwerk-Adapter-Informationen verfügbar."
                $reportContent += ""
            }
        }
        
        # Formatieren des Berichts je nach Ausgabeformat
        switch ($Format) {
            "HTML" {
                $htmlContent = "<html><head><title>Windows Update Systembericht</title>"
                $htmlContent += "<style type='text/css'>"
                $htmlContent += "body { font-family: Arial, sans-serif; margin: 20px; }"
                $htmlContent += "h1 { color: #0066cc; }"
                $htmlContent += "h2 { color: #0099cc; margin-top: 20px; }"
                $htmlContent += ".section { margin-bottom: 20px; padding: 10px; border: 1px solid #ddd; }"
                $htmlContent += "table { border-collapse: collapse; width: 100%; }"
                $htmlContent += "th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }"
                $htmlContent += "th { background-color: #f2f2f2; }"
                $htmlContent += "</style></head><body>"
                $htmlContent += "<h1>Windows Update Systembericht</h1>"
                $htmlContent += "<p>Erstellt am: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>"
                $htmlContent += "<p>Computername: $env:COMPUTERNAME</p>"
                
                $currentSection = ""
                
                foreach ($line in $reportContent) {
                    if ($line -match "^===(.+)===$") {
                        if ($currentSection) {
                            $htmlContent += "</div>"
                        }
                        
                        $currentSection = $matches[1].Trim()
                        $htmlContent += "<h2>$currentSection</h2>"
                        $htmlContent += "<div class='section'>"
                    } elseif ($line -ne "") {
                        $htmlContent += "$line<br>"
                    }
                }
                
                if ($currentSection) {
                    $htmlContent += "</div>"
                }
                
                $htmlContent += "</body></html>"
                return $htmlContent
            }
            "CSV" {
                $csvData = @()
                $currentSection = ""
                $propertyBag = @{ }
                
                foreach ($line in $reportContent) {
                    if ($line -match "^===(.+)===") {
                        if ($currentSection -and $propertyBag.Count -gt 0) {
                            $csvData += [PSCustomObject]$propertyBag
                        }
                        $currentSection = $matches[1].Trim()
                        $propertyBag = @{ "Section" = $currentSection }
                    } elseif ($line -match "^([^:]+):\s*(.*)$") {
                        $propertyName = $matches[1].Trim() -replace '\s+', '_'
                        $propertyValue = $matches[2].Trim()
                        $propertyBag[$propertyName] = $propertyValue
                    }
                }
                
                if ($currentSection -and $propertyBag.Count -gt 0) {
                    $csvData += [PSCustomObject]$propertyBag
                }
                
                return $csvData | ConvertTo-Csv -NoTypeInformation
            }
            default { # TXT
                return $reportContent -join "`r`n"
            }
        }
    }
    catch {
        Write-Host "Fehler bei der Generierung des Systemberichts: $($_.Exception.Message)" "ERROR"
        return "FEHLER: $($_.Exception.Message)"
    }
}

# Funktion zum Öffnen von Windows-Tools
function Open-WindowsTool {
    param (
        [Parameter(Mandatory=$true)]
        [ValidateSet(
            "EventViewer", "ResourceMonitor", "ReliabilityMonitor", 
            "DeviceManager", "Services", "DiskManagement", "RegistryEditor"
        )]
        [string]$ToolName
    )
    
    try {
        Write-Host "Öffne Windows-Tool: $ToolName"
        
        switch ($ToolName) {
            "EventViewer" { 
                Start-Process "eventvwr.msc"
            }
            "ResourceMonitor" { 
                Start-Process "resmon.exe"
            }
            "ReliabilityMonitor" { 
                Start-Process "perfmon.exe" -ArgumentList "/rel"
            }
            "DeviceManager" { 
                Start-Process "devmgmt.msc"
            }
            "Services" { 
                Start-Process "services.msc"
            }
            "DiskManagement" { 
                Start-Process "diskmgmt.msc"
            }
            "RegistryEditor" { 
                Start-Process "regedit.exe"
            }
            default {
                Write-Host "Unbekanntes Windows-Tool: $ToolName" "ERROR"
                return $false
            }
        }
        
        return $true
    }
    catch {
        Write-Host "Fehler beim Öffnen des Windows-Tools '$ToolName': $($_.Exception.Message)" "ERROR"
        return $false
    }
}

# Funktion zur Registrierung der Event-Handler für Toolbox-Buttons
function Register-ToolboxEventHandlers {
    param (
        $GUI
    )
    
    try {
        # Event Viewer Button
        if ($GUI.ContainsKey("btnEventViewer") -and $null -ne $GUI["btnEventViewer"]) {
            $GUI["btnEventViewer"].Add_Click({
                $result = Open-WindowsTool -ToolName "EventViewer"
                if (-not $result) {
                    $GUI["txtToolOutput"].Text = "Fehler beim Öffnen der Ereignisanzeige."
                } else {
                    $GUI["txtToolOutput"].Text = "Ereignisanzeige wurde gestartet."
                }
            })
        }
        
        # Resource Monitor Button
        if ($GUI.ContainsKey("btnResourceMonitor") -and $null -ne $GUI["btnResourceMonitor"]) {
            $GUI["btnResourceMonitor"].Add_Click({
                $result = Open-WindowsTool -ToolName "ResourceMonitor"
                if (-not $result) {
                    $GUI["txtToolOutput"].Text = "Fehler beim Öffnen des Ressourcenmonitors."
                } else {
                    $GUI["txtToolOutput"].Text = "Ressourcenmonitor wurde gestartet."
                }
            })
        }
        
        # Reliability Monitor Button
        if ($GUI.ContainsKey("btnReliabilityMonitor") -and $null -ne $GUI["btnReliabilityMonitor"]) {
            $GUI["btnReliabilityMonitor"].Add_Click({
                $result = Open-WindowsTool -ToolName "ReliabilityMonitor"
                if (-not $result) {
                    $GUI["txtToolOutput"].Text = "Fehler beim Öffnen der Zuverlässigkeitsüberwachung."
                } else {
                    $GUI["txtToolOutput"].Text = "Zuverlässigkeitsüberwachung wurde gestartet."
                }
            })
        }
        
        # Device Manager Button
        if ($GUI.ContainsKey("btnDeviceManager") -and $null -ne $GUI["btnDeviceManager"]) {
            $GUI["btnDeviceManager"].Add_Click({
                $result = Open-WindowsTool -ToolName "DeviceManager"
                if (-not $result) {
                    $GUI["txtToolOutput"].Text = "Fehler beim Öffnen des Geräte-Managers."
                } else {
                    $GUI["txtToolOutput"].Text = "Geräte-Manager wurde gestartet."
                }
            })
        }
        
        # Services Button
        if ($GUI.ContainsKey("btnServices") -and $null -ne $GUI["btnServices"]) {
            $GUI["btnServices"].Add_Click({
                $result = Open-WindowsTool -ToolName "Services"
                if (-not $result) {
                    $GUI["txtToolOutput"].Text = "Fehler beim Öffnen der Dienste-Konsole."
                } else {
                    $GUI["txtToolOutput"].Text = "Dienste-Konsole wurde gestartet."
                }
            })
        }
        
        # Disk Management Button
        if ($GUI.ContainsKey("btnDiskManagement") -and $null -ne $GUI["btnDiskManagement"]) {
            $GUI["btnDiskManagement"].Add_Click({
                $result = Open-WindowsTool -ToolName "DiskManagement"
                if (-not $result) {
                    $GUI["txtToolOutput"].Text = "Fehler beim Öffnen der Datenträgerverwaltung."
                } else {
                    $GUI["txtToolOutput"].Text = "Datenträgerverwaltung wurde gestartet."
                }
            })
        }
        
        # Registry Editor Button
        if ($GUI.ContainsKey("btnRegistryEditor") -and $null -ne $GUI["btnRegistryEditor"]) {
            $GUI["btnRegistryEditor"].Add_Click({
                $result = Open-WindowsTool -ToolName "RegistryEditor"
                if (-not $result) {
                    $GUI["txtToolOutput"].Text = "Fehler beim Öffnen des Registrierungs-Editors."
                } else {
                    $GUI["txtToolOutput"].Text = "Registrierungs-Editor wurde gestartet."
                }
            })
        }
        
        # Update Troubleshooter Button
        if ($GUI.ContainsKey("btnUpdateTroubleshooter") -and $null -ne $GUI["btnUpdateTroubleshooter"]) {
            $GUI["btnUpdateTroubleshooter"].Add_Click({
                $GUI["txtToolOutput"].Text = "Starte Windows Update-Problembehandlung..."
                $result = Run-UpdateTroubleshooter
                
                if ($result) {
                    $GUI["txtToolOutput"].Text = "Windows Update-Problembehandlung wurde gestartet."
                } else {
                    $GUI["txtToolOutput"].Text = "Fehler beim Starten der Windows Update-Problembehandlung."
                }
            })
        }
        
        # Reset Components Button
        if ($GUI.ContainsKey("btnResetComponents") -and $null -ne $GUI["btnResetComponents"]) {
            $GUI["btnResetComponents"].Add_Click({
                $GUI["txtToolOutput"].Text = "Windows Update-Komponenten werden zurückgesetzt..."
                
                try {
                    # Administrator-Rechte prüfen
                    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
                    
                    if (-not $isAdmin) {
                        $GUI["txtToolOutput"].Text = "Für das Zurücksetzen der Update-Komponenten sind Administratorrechte erforderlich."
                        return
                    }
                    
                    # Update-Dienste stoppen
                    $GUI["txtToolOutput"].AppendText("`r`nStoppe Windows Update-Dienste...")
                    Stop-Service -Name wuauserv, BITS, cryptsvc, TrustedInstaller -Force -ErrorAction SilentlyContinue
                    
                    # Software Distribution und CatRoot-Ordner umbenennen
                    $GUI["txtToolOutput"].AppendText("`r`nBenenne Cache-Ordner um...")
                    Rename-Item -Path "$env:SystemRoot\SoftwareDistribution" -NewName "SoftwareDistribution.old" -ErrorAction SilentlyContinue
                    Rename-Item -Path "$env:SystemRoot\System32\catroot2" -NewName "catroot2.old" -ErrorAction SilentlyContinue
                    
                    # Update-Dienste neustarten
                    $GUI["txtToolOutput"].AppendText("`r`nStarte Windows Update-Dienste...")
                    Start-Service -Name cryptsvc, BITS, wuauserv, TrustedInstaller -ErrorAction SilentlyContinue
                    
                    # WUAUSERV zurücksetzen
                    $GUI["txtToolOutput"].AppendText("`r`nSetze Windows Update-Dienst zurück...")
                    & sc.exe sdset wuauserv "D:(A;;CCLCSWLOCRRC;;;AU)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;SY)"
                    
                    # REG.DAT zurücksetzen
                    $GUI["txtToolOutput"].AppendText("`r`nBereinige REG.DAT Eintrag...")
                    $regPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate"
                    Remove-ItemProperty -Path $regPath -Name "AccountDomainSid" -ErrorAction SilentlyContinue
                    Remove-ItemProperty -Path $regPath -Name "PingID" -ErrorAction SilentlyContinue
                    Remove-ItemProperty -Path $regPath -Name "SusClientId" -ErrorAction SilentlyContinue
                    
                    # Windows Update-Agent zurücksetzen
                    $GUI["txtToolOutput"].AppendText("`r`nRegistriere Windows Update-Komponenten neu...")
                    $components = @(
                        "$env:SystemRoot\System32\wuaueng.dll",
                        "$env:SystemRoot\System32\wuapi.dll",
                        "$env:SystemRoot\System32\wups.dll",
                        "$env:SystemRoot\System32\wups2.dll"
                    )
                    
                    foreach ($component in $components) {
                        if (Test-Path -Path $component) {
                            & regsvr32.exe /s $component
                        }
                    }
                    
                    # WSUS-Client-Registrierung zurücksetzen
                    $GUI["txtToolOutput"].AppendText("`r`nSetze WSUS-Client zurück...")
                    & wuauclt.exe /resetauthorization /detectnow
                    
                    $GUI["txtToolOutput"].AppendText("`r`n`r`nWindows Update-Komponenten wurden erfolgreich zurückgesetzt. Bitte starten Sie den Computer neu.")
                }
                catch {
                    $GUI["txtToolOutput"].AppendText("`r`nFehler beim Zurücksetzen der Update-Komponenten: $($_.Exception.Message)")
                }
            })
        }
        
        # Clear Update Cache Button
        if ($GUI.ContainsKey("btnClearUpdateCache") -and $null -ne $GUI["btnClearUpdateCache"]) {
            $GUI["btnClearUpdateCache"].Add_Click({
                $GUI["txtToolOutput"].Text = "Update-Cache wird geleert..."
                
                try {
                    # Administrator-Rechte prüfen
                    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
                    
                    if (-not $isAdmin) {
                        $GUI["txtToolOutput"].Text = "Für das Leeren des Update-Caches sind Administratorrechte erforderlich."
                        return
                    }
                    
                    # Windows Update-Dienste stoppen
                    $GUI["txtToolOutput"].AppendText("`r`nStoppe Windows Update-Dienste...")
                    Stop-Service -Name wuauserv, BITS -Force -ErrorAction SilentlyContinue
                    
                    # Software Distribution-Ordner leeren
                    $GUI["txtToolOutput"].AppendText("`r`nBereinige SoftwareDistribution-Ordner...")
                    Get-ChildItem -Path "$env:SystemRoot\SoftwareDistribution" -Recurse | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
                    
                    # Windows Update-Dienste starten
                    $GUI["txtToolOutput"].AppendText("`r`nStarte Windows Update-Dienste...")
                    Start-Service -Name BITS, wuauserv -ErrorAction SilentlyContinue
                    
                    $GUI["txtToolOutput"].AppendText("`r`n`r`nUpdate-Cache wurde erfolgreich geleert.")
                }
                catch {
                    $GUI["txtToolOutput"].AppendText("`r`nFehler beim Leeren des Update-Caches: $($_.Exception.Message)")
                }
            })
        }
        
        # DISM Health Button
        if ($GUI.ContainsKey("btnDISMHealth") -and $null -ne $GUI["btnDISMHealth"]) {
            $GUI["btnDISMHealth"].Add_Click({
                $GUI["txtToolOutput"].Text = "Starte DISM-Systemreparatur..."
                
                try {
                    # Administrator-Rechte prüfen
                    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
                    
                    if (-not $isAdmin) {
                        $GUI["txtToolOutput"].Text = "Für die DISM-Systemreparatur sind Administratorrechte erforderlich."
                        return
                    }
                    
                    # DISM-Reparatur starten
                    $GUI["txtToolOutput"].AppendText("`r`nFühre DISM /Online /Cleanup-Image /RestoreHealth aus...`r`nDies kann einige Minuten dauern...")
                    
                    $process = Start-Process -FilePath "dism.exe" -ArgumentList "/Online /Cleanup-Image /RestoreHealth" -Wait -PassThru -NoNewWindow -ErrorAction Stop
                    
                    if ($process.ExitCode -eq 0) {
                        $GUI["txtToolOutput"].AppendText("`r`n`r`nDISM-Systemreparatur erfolgreich abgeschlossen.")
                    } else {
                        $GUI["txtToolOutput"].AppendText("`r`n`r`nDISM-Systemreparatur mit Fehlercode $($process.ExitCode) beendet.")
                    }
                }
                catch {
                    $GUI["txtToolOutput"].AppendText("`r`nFehler bei der DISM-Systemreparatur: $($_.Exception.Message)")
                }
            })
        }
        
        # SFC Scan Button
        if ($GUI.ContainsKey("btnSFCScan") -and $null -ne $GUI["btnSFCScan"]) {
            $GUI["btnSFCScan"].Add_Click({
                $GUI["txtToolOutput"].Text = "Starte SFC-Systemdateiscan..."
                
                try {
                    # Administrator-Rechte prüfen
                    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
                    
                    if (-not $isAdmin) {
                        $GUI["txtToolOutput"].Text = "Für den SFC-Systemdateiscan sind Administratorrechte erforderlich."
                        return
                    }
                    
                    # SFC-Scan starten
                    $GUI["txtToolOutput"].AppendText("`r`nFühre sfc /scannow aus...`r`nDies kann einige Minuten dauern...")
                    
                    $process = Start-Process -FilePath "sfc.exe" -ArgumentList "/scannow" -Wait -PassThru -NoNewWindow -ErrorAction Stop
                    
                    if ($process.ExitCode -eq 0) {
                        $GUI["txtToolOutput"].AppendText("`r`n`r`nSFC-Systemdateiscan erfolgreich abgeschlossen.")
                    } else {
                        $GUI["txtToolOutput"].AppendText("`r`n`r`nSFC-Systemdateiscan mit Fehlercode $($process.ExitCode) beendet.")
                    }
                }
                catch {
                    $GUI["txtToolOutput"].AppendText("`r`nFehler beim SFC-Systemdateiscan: $($_.Exception.Message)")
                }
            })
        }
        
        # Search Microsoft Documentation Button
        if ($GUI.ContainsKey("btnSearchDocs") -and $null -ne $GUI["btnSearchDocs"] -and $GUI.ContainsKey("txtSearchDocs")) {
            $GUI["btnSearchDocs"].Add_Click({
                $searchTerm = $GUI["txtSearchDocs"].Text.Trim()
                
                if ([string]::IsNullOrEmpty($searchTerm)) {
                    $GUI["txtToolOutput"].Text = "Bitte einen Suchbegriff eingeben."
                    return
                }
                
                $GUI["txtToolOutput"].Text = "Öffne Microsoft Dokumentation für Suche nach: $searchTerm"
                
                $result = Open-MicrosoftDocs -SearchTerm $searchTerm -Category "WindowsUpdate"
                
                if (-not $result) {
                    $GUI["txtToolOutput"].Text += "`r`nFehler beim Öffnen der Microsoft Dokumentation."
                }
            })
        }
        
        Write-Host "Toolbox-Event-Handler erfolgreich registriert"
        return $true
    }
    catch {
        Write-Host "Fehler beim Registrieren der Toolbox-Event-Handler: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

# Funktion zum Generieren und Anzeigen des Systemberichts in der GUI
function Update-SystemReportView {
    param (
        $GUI
    )
    
    try {
        Write-Host "Generiere Systembericht..."
        
        if (-not $GUI.ContainsKey("chkReportSystemInfo") -or -not $GUI.ContainsKey("txtReportPreview") -or -not $GUI.ContainsKey("txtReportStatus")) {
            Write-Host "Erforderliche UI-Elemente für Systembericht nicht gefunden" "WARNING"
            return $false
        }
        
        # Report-Optionen aus GUI-Elementen auslesen
        $includeSystemInfo = $GUI["chkReportSystemInfo"].IsChecked -eq $true
        $includeUpdateConfig = $GUI["chkReportUpdateConfig"].IsChecked -eq $true
        $includeUpdateHistory = $GUI["chkReportUpdateHistory"].IsChecked -eq $true
        $includeServices = $GUI["chkReportServices"].IsChecked -eq $true
        $includeDiskSpace = $GUI["chkReportDiskSpace"].IsChecked -eq $true
        $includeEventLogs = $GUI["chkReportEventLogs"].IsChecked -eq $true
        $includeNetworkConfig = $GUI["chkReportNetworkConfig"].IsChecked -eq $true
        
        # Format bestimmen
        $reportFormat = "TXT"
        if ($GUI.ContainsKey("radFormatHtml") -and $GUI["radFormatHtml"].IsChecked) {
            $reportFormat = "HTML"
        }
        elseif ($GUI.ContainsKey("radFormatCsv") -and $GUI["radFormatCsv"].IsChecked) {
            $reportFormat = "CSV"
        }
        
        # Statusmeldung aktualisieren
        $GUI["txtReportStatus"].Text = "Generiere Systembericht..."
        
        # Report generieren
        $report = Generate-SystemReport -IncludeSystemInfo:$includeSystemInfo `
                                       -IncludeUpdateConfig:$includeUpdateConfig `
                                       -IncludeUpdateHistory:$includeUpdateHistory `
                                       -IncludeServices:$includeServices `
                                       -IncludeDiskSpace:$includeDiskSpace `
                                       -IncludeEventLogs:$includeEventLogs `
                                       -IncludeNetworkConfig:$includeNetworkConfig `
                                       -Format $reportFormat
        
        # Report in Preview-Textbox anzeigen
        $GUI["txtReportPreview"].Text = $report
        
        # Status aktualisieren
        $GUI["txtReportStatus"].Text = "Systembericht wurde erfolgreich generiert: $(Get-Date -Format 'HH:mm:ss')"
        
        Write-Host "Systembericht erfolgreich generiert"
        return $true
    }
    catch {
        Write-Host "Fehler beim Generieren des Systemberichts: $($_.Exception.Message)" "ERROR"
        
        if ($GUI.ContainsKey("txtReportStatus")) {
            $GUI["txtReportStatus"].Text = "Fehler beim Generieren des Berichts: $($_.Exception.Message)"
        }
        
        return $false
    }
}

# Funktion zum Exportieren des Systemberichts
function Export-SystemReport {
    param (
        $GUI
    )
    
    try {
        # Prüfen, ob der Bericht generiert wurde
        if (-not $GUI.ContainsKey("txtReportPreview") -or [string]::IsNullOrEmpty($GUI["txtReportPreview"].Text)) {
            Write-Host "Kein Bericht zum Exportieren vorhanden" "WARNING"
            return $false
        }
        
        # Format bestimmen
        $reportFormat = "txt"
        $filter = "Textdateien (*.txt)|*.txt|Alle Dateien (*.*)|*.*"
        
        if ($GUI.ContainsKey("radFormatHtml") -and $GUI["radFormatHtml"].IsChecked) {
            $reportFormat = "html"
            $filter = "HTML-Dateien (*.html)|*.html|Alle Dateien (*.*)|*.*"
        }
        elseif ($GUI.ContainsKey("radFormatCsv") -and $GUI["radFormatCsv"].IsChecked) {
            $reportFormat = "csv"
            $filter = "CSV-Dateien (*.csv)|*.csv|Alle Dateien (*.*)|*.*"
        }
        
        # Speicherdialog öffnen
        $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
        $saveDialog.Filter = $filter
        $saveDialog.DefaultExt = ".$reportFormat"
        $saveDialog.Title = "Systembericht speichern"
        $saveDialog.FileName = "WindowsUpdate_Systembericht_$(Get-Date -Format 'yyyyMMdd').$reportFormat"
        
        if ($saveDialog.ShowDialog()) {
            # Bericht in Datei speichern
            $GUI["txtReportPreview"].Text | Out-File -FilePath $saveDialog.FileName -Encoding utf8
            
            # Status aktualisieren
            if ($GUI.ContainsKey("txtReportStatus")) {
                $GUI["txtReportStatus"].Text = "Bericht wurde unter '$($saveDialog.FileName)' gespeichert."
            }
            
            return $true
        }
        else {
            return $false
        }
    }
    catch {
        Write-Host "Fehler beim Exportieren des Systemberichts: $($_.Exception.Message)" "ERROR"
        
        if ($GUI.ContainsKey("txtReportStatus")) {
            $GUI["txtReportStatus"].Text = "Fehler beim Exportieren des Berichts: $($_.Exception.Message)"
        }
        
        return $false
    }
}

# Funktion zum Registrieren der Event-Handler für den Systembericht
function Register-SystemReportEventHandlers {
    param (
        $GUI
    )
    
    try {
        # Generate Report Button
        if ($GUI.ContainsKey("btnGenerateReport") -and $null -ne $GUI["btnGenerateReport"]) {
            $GUI["btnGenerateReport"].Add_Click({
                Update-SystemReportView -GUI $GUI
            })
        }
        
        # Export Report Button
        if ($GUI.ContainsKey("btnExportReport") -and $null -ne $GUI["btnExportReport"]) {
            $GUI["btnExportReport"].Add_Click({
                Export-SystemReport -GUI $GUI
            })
        }
        
        # Report-Format RadioButtons für Live-Preview
        if ($GUI.ContainsKey("radFormatTxt") -and $null -ne $GUI["radFormatTxt"]) {
            $GUI["radFormatTxt"].Add_Checked({
                if ($GUI["SystemReport"].Visibility -eq "Visible") {
                    Update-SystemReportView -GUI $GUI
                }
            })
        }
        
        if ($GUI.ContainsKey("radFormatHtml") -and $null -ne $GUI["radFormatHtml"]) {
            $GUI["radFormatHtml"].Add_Checked({
                if ($GUI["SystemReport"].Visibility -eq "Visible") {
                    Update-SystemReportView -GUI $GUI
                }
            })
        }
        
        if ($GUI.ContainsKey("radFormatCsv") -and $null -ne $GUI["radFormatCsv"]) {
            $GUI["radFormatCsv"].Add_Checked({
                if ($GUI["SystemReport"].Visibility -eq "Visible") {
                    Update-SystemReportView -GUI $GUI
                }
            })
        }
        
        # Checkboxen für Report-Optionen mit Live-Preview
        $reportCheckboxes = @(
            "chkReportSystemInfo", "chkReportUpdateConfig", "chkReportUpdateHistory",
            "chkReportServices", "chkReportDiskSpace", "chkReportEventLogs", "chkReportNetworkConfig"
        )
        
        foreach ($checkboxName in $reportCheckboxes) {
            if ($GUI.ContainsKey($checkboxName) -and $null -ne $GUI[$checkboxName]) {
                $GUI[$checkboxName].Add_Checked({
                    if ($GUI["SystemReport"].Visibility -eq "Visible") {
                        Update-SystemReportView -GUI $GUI
                    }
                })
                
                $GUI[$checkboxName].Add_Unchecked({
                    if ($GUI["SystemReport"].Visibility -eq "Visible") {
                        Update-SystemReportView -GUI $GUI
                    }
                })
            }
        }
        
        Write-Host "System Report Event Handler erfolgreich registriert"
        return $true
    }
    catch {
        Write-Host "Fehler beim Registrieren der System Report Event Handler: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

# Funktion zum Registrieren der Error-Handler für Troubleshooting
function Register-TroubleshootingEventHandlers {
    param (
        $GUI
    )
    
    try {
        # Initialisierung der Fehlercode-Liste in der DataGrid
        if ($GUI.ContainsKey("errorCodesDataGrid")) {
            # Windows Update Fehlercodes laden
            $errorCodes = Get-WindowsUpdateErrorCodes
            
            if ($errorCodes) {
                # ItemsSource mit Fehlercode-Mapping setzen
                $formattedCodes = $errorCodes | ForEach-Object {
                    [PSCustomObject]@{
                        Fehlercode = $_.ErrorCode
                        Beschreibung = $_.Description
                        Befehl = $_.Command
                        Dokumentation = $_.Documentation
                    }
                }
                
                $GUI["errorCodesDataGrid"].ItemsSource = $formattedCodes
            }
        }
        
        # Such-Placeholder für Fehlercode-Feld
        if ($GUI.ContainsKey("txtErrorSearch") -and $GUI.ContainsKey("txtErrorSearchPlaceholder")) {
            $GUI["txtErrorSearch"].Add_GotFocus({
                $GUI["txtErrorSearchPlaceholder"].Visibility = "Hidden"
            })
            
            $GUI["txtErrorSearch"].Add_LostFocus({
                if ([string]::IsNullOrEmpty($GUI["txtErrorSearch"].Text)) {
                    $GUI["txtErrorSearchPlaceholder"].Visibility = "Visible"
                }
            })
        }
        
        # Search Error Button
        if ($GUI.ContainsKey("btnSearchError") -and $null -ne $GUI["btnSearchError"]) {
            $GUI["btnSearchError"].Add_Click({
                if ($GUI.ContainsKey("txtErrorSearch") -and $null -ne $GUI["txtErrorSearch"]) {
                    $searchTerm = $GUI["txtErrorSearch"].Text.Trim()
                    
                    # Wenn kein Fehlercode eingegeben wurde, Meldung anzeigen
                    if ([string]::IsNullOrEmpty($searchTerm)) {
                        [System.Windows.MessageBox]::Show("Bitte geben Sie einen Fehlercode ein.", "Fehlercode fehlt", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
                        return
                    }
                    
                    # Fehlercode-Informationen abrufen
                    $errorInfo = Get-UpdateErrorInfo -ErrorCode $searchTerm
                    
                    if ($null -eq $errorInfo) {
                        [System.Windows.MessageBox]::Show("Keine Informationen für Fehlercode '$searchTerm' gefunden.", "Keine Informationen", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
                        return
                    }
                    
                    # Wenn mehrere Ergebnisse gefunden wurden, das erste nehmen
                    if ($errorInfo -is [array]) {
                        $errorInfo = $errorInfo[0]
                    }
                    
                    # Fehlercode-Informationen in die UI schreiben
                    if ($GUI.ContainsKey("txtErrorCode")) {
                        $GUI["txtErrorCode"].Text = $errorInfo.ErrorCode
                    }
                    
                    if ($GUI.ContainsKey("txtErrorDescription")) {
                        $GUI["txtErrorDescription"].Text = $errorInfo.Description
                    }
                    
                    if ($GUI.ContainsKey("txtErrorCommand")) {
                        $GUI["txtErrorCommand"].Text = $errorInfo.Command
                    }
                    
                    if ($GUI.ContainsKey("txtErrorDocumentation")) {
                        $GUI["txtErrorDocumentation"].Text = $errorInfo.Documentation
                    }
                    
                    # DataGrid auf entsprechenden Fehlercode setzen
                    if ($GUI.ContainsKey("errorCodesDataGrid")) {
                        $items = $GUI["errorCodesDataGrid"].ItemsSource
                        $item = $items | Where-Object { $_.Fehlercode -like "*$searchTerm*" } | Select-Object -First 1
                        
                        if ($item) {
                            $GUI["errorCodesDataGrid"].SelectedItem = $item
                            $GUI["errorCodesDataGrid"].ScrollIntoView($item)
                        }
                    }
                }
            })
        }
        
        # Execute Command Button
        if ($GUI.ContainsKey("btnExecuteCommand") -and $null -ne $GUI["btnExecuteCommand"]) {
            $GUI["btnExecuteCommand"].Add_Click({
                try {
                    # Befehl aus Textfeld auslesen
                    if (-not $GUI.ContainsKey("txtErrorCommand") -or [string]::IsNullOrWhiteSpace($GUI["txtErrorCommand"].Text)) {
                        [System.Windows.MessageBox]::Show("Kein Befehl zum Ausführen verfügbar.", "Kein Befehl", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
                        return
                    }
                    
                    $command = $GUI["txtErrorCommand"].Text.Trim()
                    
                    # Wenn Befehl leer oder "Kein spezifischer Befehl" ist, abbrechen
                    if ([string]::IsNullOrWhiteSpace($command) -or $command -like "Kein spezifischer Befehl*") {
                        [System.Windows.MessageBox]::Show("Für diesen Fehler ist kein automatisierter Befehl verfügbar.", "Kein Befehl", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
                        return
                    }
                    
                    # Bestätigung einholen
                    $confirmation = [System.Windows.MessageBox]::Show(
                        "Möchten Sie den folgenden Befehl ausführen?`n`n$command`n`nDies könnte Auswirkungen auf Ihr System haben.",
                        "Befehl ausführen",
                        [System.Windows.MessageBoxButton]::YesNo,
                        [System.Windows.MessageBoxImage]::Warning
                    )
                    
                    if ($confirmation -eq "No") {
                        return
                    }
                    
                    # Befehl ausführen
                    $process = Start-Process -FilePath "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -Command `"$command`"" -Wait -PassThru -NoNewWindow
                    
                    if ($process.ExitCode -eq 0) {
                        [System.Windows.MessageBox]::Show("Der Befehl wurde erfolgreich ausgeführt.", "Erfolg", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
                    } else {
                        [System.Windows.MessageBox]::Show("Der Befehl wurde mit Fehlercode $($process.ExitCode) beendet.", "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                    }
                }
                catch {
                    [System.Windows.MessageBox]::Show("Fehler beim Ausführen des Befehls: $($_.Exception.Message)", "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                }
            })
        }
        
        # Show Documentation Button
        if ($GUI.ContainsKey("btnShowDocs") -and $null -ne $GUI["btnShowDocs"]) {
            $GUI["btnShowDocs"].Add_Click({
                if ($GUI.ContainsKey("errorCodesDataGrid") -and $null -ne $GUI["errorCodesDataGrid"]) {
                    $selectedItem = $GUI["errorCodesDataGrid"].SelectedItem
                    
                    if ($null -eq $selectedItem) {
                        [System.Windows.MessageBox]::Show("Bitte wählen Sie einen Fehlercode aus der Liste aus.", "Kein Fehlercode ausgewählt", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
                        return
                    }
                    
                    $errorCode = $selectedItem.Fehlercode
                    $docLink = $selectedItem.Dokumentation
                    
                    if ([string]::IsNullOrEmpty($docLink) -or $docLink -eq "-") {
                        # Wenn kein direkter Link verfügbar ist, nach dem Fehlercode suchen
                        Open-MicrosoftDocs -SearchTerm $errorCode -Category "WindowsUpdate"
                    } else {
                        # Versuchen, den Link in einem Browser zu öffnen
                        try {
                            Start-Process $docLink
                        }
                        catch {
                            # Bei Fehler fallback auf Suche
                            Open-MicrosoftDocs -SearchTerm $errorCode -Category "WindowsUpdate"
                        }
                    }
                }
            })
        }
        
        # Run Diagnostics Button
        if ($GUI.ContainsKey("btnRunDiagnostics") -and $null -ne $GUI["btnRunDiagnostics"]) {
            $GUI["btnRunDiagnostics"].Add_Click({
                # Ausgewählten Fehlercode abrufen, falls vorhanden
                $errorCode = ""
                
                if ($GUI.ContainsKey("errorCodesDataGrid") -and $null -ne $GUI["errorCodesDataGrid"]) {
                    $selectedItem = $GUI["errorCodesDataGrid"].SelectedItem
                    
                    if ($null -ne $selectedItem) {
                        $errorCode = $selectedItem.Fehlercode
                        # Use the error code for the troubleshooter
                        Write-Host "Starte Windows Update-Problembehandlung für Fehlercode: $errorCode"
                    }
                }
                
                # Windows Update Diagnose starten
                $result = Run-UpdateTroubleshooter
                
                if (-not $result) {
                    [System.Windows.MessageBox]::Show(
                        "Die Windows Update-Problembehandlung konnte nicht gestartet werden.", 
                        "Fehler", 
                        [System.Windows.MessageBoxButton]::OK, 
                        [System.Windows.MessageBoxImage]::Error
                    )
                }
            })
        }
        
        Write-Host "Troubleshooting-Event-Handler erfolgreich registriert"
        return $true
    }
    catch {
        Write-Host "Fehler beim Registrieren der Troubleshooting-Event-Handler: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

# Register all remaining event handlers for Dashboard buttons, Search buttons, etc.
function Register-AllEventHandlers {
    param (
        $GUI,
        $Window
    )
    
    try {
        # First register the main navigation event handlers
        Register-EventHandlers -GUI $GUI -Window $Window
        
        # Then register the toolbox event handlers
        Register-ToolboxEventHandlers -GUI $GUI
        
        # Register system report event handlers
        Register-SystemReportEventHandlers -GUI $GUI
        
        # Register troubleshooting event handlers
        Register-TroubleshootingEventHandlers -GUI $GUI
        
        # Initialize Dashboard System Info on load
        if ($GUI.ContainsKey("Dashboard") -and $GUI.ContainsKey("txtSystemInfo") -and $GUI.ContainsKey("txtWsusStatus")) {
            # Initial load of system information
            try {
                $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem
                $operatingSystem = Get-CimInstance -ClassName Win32_OperatingSystem
                
                $GUI["txtSystemInfo"].Text = "$($computerSystem.Manufacturer) $($computerSystem.Model)`r`n"
                $GUI["txtSystemInfo"].Text += "$($operatingSystem.Caption) ($($operatingSystem.Version))`r`n"
                $GUI["txtSystemInfo"].Text += "$($operatingSystem.OSArchitecture), Build $($operatingSystem.BuildNumber)`r`n"
                $GUI["txtSystemInfo"].Text += "Computername: $($computerSystem.Name)"
                
                # Windows Update-Status abrufen
                $updateServices = Get-Service -Name wuauserv, BITS, UsoSvc, WaaSMedicSvc -ErrorAction SilentlyContinue
                $wuStatus = "Betriebsbereit"
                
                if ($updateServices) {
                    $stoppedServices = $updateServices | Where-Object { $_.Status -ne "Running" }
                    
                    if ($stoppedServices -and $stoppedServices.Count -gt 0) {
                        $wuStatus = "Eingeschränkt (Dienste inaktiv: $($stoppedServices.Name -join ', '))"
                    }
                }
                
                $GUI["txtWsusStatus"].Text = $wuStatus
                
                # Letzte Update-Prüfung abrufen
                $lastCheck = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\Results\Detect" -ErrorAction SilentlyContinue
                if ($lastCheck -and $lastCheck.LastSuccessTime) {
                    $GUI["txtLastCheck"].Text = "Letzte Prüfung: $($lastCheck.LastSuccessTime)"
                } else {
                    $GUI["txtLastCheck"].Text = "Letzte Prüfung: Unbekannt"
                }
                
                # Letzte Update-Installation abrufen
                $lastInstall = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\Results\Install" -ErrorAction SilentlyContinue
                if ($lastInstall -and $lastInstall.LastSuccessTime) {
                    $GUI["txtLastInstall"].Text = "Letzte Installation: $($lastInstall.LastSuccessTime)"
                } else {
                    $GUI["txtLastInstall"].Text = "Letzte Installation: Unbekannt"
                }
                
                # Nächste geplante Installation abrufen
                $autoUpdate = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update" -ErrorAction SilentlyContinue
                if ($autoUpdate -and $autoUpdate.AUOptions -eq 4) {
                    $wuSchedule = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update" -Name "ScheduledInstallDay", "ScheduledInstallTime" -ErrorAction SilentlyContinue
                    if ($wuSchedule) {
                        $days = @("Jeden Tag", "Sonntag", "Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag")
                        $day = $days[$wuSchedule.ScheduledInstallDay]
                        $time = "$($wuSchedule.ScheduledInstallTime):00"
                        
                        $GUI["txtNextScheduled"].Text = "Nächste geplante Installation: $day um $time Uhr"
                    }
                } else {
                    $GUI["txtNextScheduled"].Text = "Nächste geplante Installation: Nicht konfiguriert"
                }
            }
            catch {
                Write-Host "Fehler beim Initialisieren der Dashboard-Informationen: $($_.Exception.Message)" "ERROR"
            }
        }
        
        # Update History Buttons
        if ($GUI.ContainsKey("btnRefreshHistory") -and $null -ne $GUI["btnRefreshHistory"]) {
            $GUI["btnRefreshHistory"].Add_Click({
                try {
                    # Anzahl der anzuzeigenden Updates bestimmen
                    $historyCount = 50 # Standard
                    
                    if ($GUI.ContainsKey("cmbHistoryCount") -and $null -ne $GUI["cmbHistoryCount"]) {
                        switch ($GUI["cmbHistoryCount"].SelectedIndex) {
                            0 { $historyCount = 10 }
                            1 { $historyCount = 25 }
                            2 { $historyCount = 50 }
                            3 { $historyCount = 100 }
                            4 { $historyCount = 0 } # Alle
                        }
                    }
                    
                    # Windows Update COM-Objekt erstellen
                    $updateSession = New-Object -ComObject Microsoft.Update.Session
                    $updateSearcher = $updateSession.CreateUpdateSearcher()
                    
                    # Gesamtanzahl der Updates im Verlauf abrufen
                    $totalHistoryCount = $updateSearcher.GetTotalHistoryCount()
                    
                    # Update-Verlauf abrufen
                    if ($historyCount -eq 0 -or $historyCount -gt $totalHistoryCount) {
                        $historyCount = $totalHistoryCount
                    }
                    
                    $history = $updateSearcher.QueryHistory(0, $historyCount)
                    
                    # Update-Verlauf formatieren
                    $updateHistory = @()
                    
                    foreach ($entry in $history) {
                        # Result-Code interpretieren
                        $resultString = switch ($entry.ResultCode) {
                            0 { "Nicht gestartet" }
                            1 { "In Bearbeitung" }
                            2 { "Erfolgreich" }
                            3 { "Unvollständig" }
                            4 { "Fehlgeschlagen" }
                            5 { "Abgebrochen" }
                            default { "Unbekannt ($($entry.ResultCode))" }
                        }
                        
                        # Operation-Code interpretieren
                        $operationString = switch ($entry.Operation) {
                            0 { "Unbekannt" }
                            1 { "Installation" }
                            2 { "Deinstallation" }
                            3 { "Andere" }
                            default { "Unbekannt ($($entry.Operation))" }
                        }
                        
                        # KB number is removed since it's already in the title
                        $updateHistory += [PSCustomObject]@{
                            Datum = $entry.Date
                            Titel = $entry.Title
                            Operation = $operationString
                            Status = $resultString
                            Fehlercode = if ($entry.ResultCode -eq 2) { "N/A" } else { "0x" + [System.Convert]::ToString($entry.HResult, 16) }
                        }
                    }
                    
                    # DataGrid aktualisieren
                    $GUI["gridUpdateHistory"].ItemsSource = $updateHistory
                    
                    # Status aktualisieren
                    if ($GUI.ContainsKey("txtHistoryStatus")) {
                        $GUI["txtHistoryStatus"].Text = "$($updateHistory.Count) Updates im Verlauf gefunden. Letzter Scan: $(Get-Date -Format 'HH:mm:ss')"
                    }
                }
                catch {
                    Write-Host "Fehler beim Abrufen des Update-Verlaufs: $($_.Exception.Message)" "ERROR"
                    if ($GUI.ContainsKey("txtHistoryStatus")) {
                        $GUI["txtHistoryStatus"].Text = "Fehler beim Abrufen des Verlaufs: $($_.Exception.Message)"
                    }
                }
            })
        }
        
        if ($GUI.ContainsKey("btnExportHistory") -and $null -ne $GUI["btnExportHistory"]) {
            $GUI["btnExportHistory"].Add_Click({
                try {
                    # Speicherdialog öffnen
                    $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
                    $saveDialog.Filter = "CSV-Dateien (*.csv)|*.csv|Alle Dateien (*.*)|*.*"
                    $saveDialog.DefaultExt = ".csv"
                    $saveDialog.Title = "Update-Verlauf exportieren"
                    $saveDialog.FileName = "WindowsUpdate_History_$(Get-Date -Format 'yyyyMMdd').csv"
                    
                    if ($saveDialog.ShowDialog()) {
                        # Aktuellen ItemsSource als CSV exportieren
                        $GUI["gridUpdateHistory"].ItemsSource | Export-Csv -Path $saveDialog.FileName -NoTypeInformation -Encoding UTF8
                        
                        [System.Windows.MessageBox]::Show(
                            "Update-Verlauf wurde erfolgreich nach '$($saveDialog.FileName)' exportiert.", 
                            "Export erfolgreich", 
                            [System.Windows.MessageBoxButton]::OK, 
                            [System.Windows.MessageBoxImage]::Information
                        )
                    }
                }
                catch {
                    Write-Host "Fehler beim Exportieren des Update-Verlaufs - $($_.Exception.Message)" "ERROR"
                    
                    [System.Windows.MessageBox]::Show(
                        "Fehler beim Exportieren des Update-Verlaufs: $($_.Exception.Message)", 
                        "Export fehlgeschlagen", 
                        [System.Windows.MessageBoxButton]::OK, 
                        [System.Windows.MessageBoxImage]::Error
                    )
                }
            })
        }
        
        return $true
    }
    catch {
        Write-Host "Fehler beim Registrieren der Event-Handler: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

# Funktion zum Abrufen der Windows Update Registry-Einstellungen
function Get-WindowsUpdateRegistrySettings {
    param (
        [Parameter(Mandatory=$true)]
        [ValidateSet("WU", "WUAU", "Current")]
        [string]$SettingsType
    )
    
    try {
        # Registry-Pfade basierend auf dem Typ definieren
        $regPath = switch ($SettingsType) {
            "WU" { "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate" }
            "WUAU" { "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" }
            "Current" { "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate" }
        }
        
        # Einstellungen-Array erstellen
        $settings = @()
        
        # Prüfen, ob der Pfad existiert
        if (Test-Path -Path $regPath) {
            # Registry-Werte abrufen
            $regValues = Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue
            
            # Properties durchlaufen
            foreach ($prop in $regValues.PSObject.Properties) {
                # System-Properties ausschließen
                if ($prop.Name -notmatch "^(PSPath|PSParentPath|PSChildName|PSDrive|PSProvider)$") {
                    # Beschreibung basierend auf dem Namen zuordnen
                    $description = Get-RegistryValueDescription -Path $regPath -Name $prop.Name
                    
                    # Wert formatieren
                    $value = Format-RegistryValue -Value $prop.Value -Type $prop.TypeNameOfValue
                    
                    # Zum Array hinzufügen
                    $settings += [PSCustomObject]@{
                        Key = $prop.Name
                        Value = $value
                        Description = $description
                    }
                }
            }
        }
        
        Write-Host "Gefunden: $($settings.Count) Einstellungen für $SettingsType" "INFO"
        return $settings
    }
    catch {
        Write-Host "Fehler beim Abrufen der Registry-Einstellungen: $($_.Exception.Message)" "ERROR"
        return @()
    }
}

# Hilfsfunktion zur Formatierung von Registry-Werten
function Format-RegistryValue {
    param (
        $Value,
        [string]$Type
    )
    
    try {
        # Formatierung basierend auf Typ
        switch -Regex ($Type) {
            "Int32|UInt32|Int64|UInt64|Boolean" {
                return $Value
            }
            "Binary" {
                # Binärwert in Hex konvertieren
                return ($Value | ForEach-Object { $_.ToString("X2") }) -join " "
            }
            default {
                # String oder andere Typen
                return "$Value"
            }
        }
    }
    catch {
        return "$Value"
    }
}

# Hilfsfunktion zur Beschreibung von Registry-Werten
function Get-RegistryValueDescription {
    param (
        [string]$Path,
        [string]$Name
    )
    
    # Bekannte Registry-Werte und ihre Beschreibungen
    $knownValues = @{
        # Windows Update\AU Einstellungen
        "AUOptions" = @{
            Description = "Konfiguration für automatische Updates"
            Values = @{
                "1" = "Automatische Updates deaktivieren"
                "2" = "Vor Download benachrichtigen"
                "3" = "Automatisch herunterladen und vor Installation benachrichtigen"
                "4" = "Automatisch herunterladen und für Installation planen"
                "5" = "Lokale Administratoren können Einstellungen festlegen"
            }
        }
        "NoAutoUpdate" = @{
            Description = "Automatische Updates deaktivieren"
            Values = @{
                "0" = "Automatische Updates sind aktiviert"
                "1" = "Automatische Updates sind deaktiviert"
            }
        }
        "ScheduledInstallDay" = @{
            Description = "Tag für geplante Installationen"
            Values = @{
                "0" = "Jeden Tag"
                "1" = "Sonntag"
                "2" = "Montag"
                "3" = "Dienstag"
                "4" = "Mittwoch"
                "5" = "Donnerstag"
                "6" = "Freitag"
                "7" = "Samstag"
            }
        }
        "ScheduledInstallTime" = @{
            Description = "Uhrzeit für geplante Installationen (Stunde, 0-23)"
        }
        "UseWUServer" = @{
            Description = "Windows Update Server verwenden"
            Values = @{
                "0" = "WSUS-Server nicht verwenden"
                "1" = "WSUS-Server verwenden"
            }
        }
        "NoAutoRebootWithLoggedOnUsers" = @{
            Description = "Automatischer Neustart mit angemeldeten Benutzern"
            Values = @{
                "0" = "Neustart kann mit angemeldeten Benutzern erfolgen"
                "1" = "Kein automatischer Neustart mit angemeldeten Benutzern"
            }
        }
        
        # Windows Update Einstellungen
        "WUServer" = @{
            Description = "URL des Windows Update-Servers (WSUS)"
        }
        "WUStatusServer" = @{
            Description = "URL des Windows Update Status-Servers"
        }
        "TargetGroup" = @{
            Description = "WSUS-Zielgruppe für diesen Computer"
        }
        "TargetGroupEnabled" = @{
            Description = "WSUS-Zielgruppe aktiviert"
            Values = @{
                "0" = "Deaktiviert"
                "1" = "Aktiviert"
            }
        }
        "DisableDualScan" = @{
            Description = "Dual Scan deaktivieren"
            Values = @{
                "0" = "Dual Scan aktiviert"
                "1" = "Dual Scan deaktiviert"
            }
        }
        "ExcludeWUDriversInQualityUpdate" = @{
            Description = "Treiber aus Windows Update ausschließen"
            Values = @{
                "0" = "Treiber einbeziehen"
                "1" = "Treiber ausschließen"
            }
        }
        "DisableMicrosoftUpdate" = @{
            Description = "Microsoft Update deaktivieren"
            Values = @{
                "0" = "Microsoft Update aktiviert"
                "1" = "Microsoft Update deaktiviert"
            }
        }
        
        # Aktuelle Windows Update Einstellungen (Current)
        "LastWUAutoupdate" = @{
            Description = "Letztes automatisches Update"
        }
        "LastOnlineScanTimeForAppCategory" = @{
            Description = "Letzte Online-Scan-Zeit für App-Kategorie"
        }
        "EnableFeaturedSoftware" = @{
            Description = "Empfohlene Software anzeigen"
            Values = @{
                "0" = "Nicht anzeigen"
                "1" = "Anzeigen"
            }
        }
    }
    
    # Standard-Beschreibung, falls nichts passendes gefunden wird
    $description = "Registry-Wert für Windows Update"
    
    # Bekannte Werte prüfen
    if ($knownValues.ContainsKey($Name)) {
        $valueInfo = $knownValues[$Name]
        $description = $valueInfo.Description
    }
    
    return $description
}

# Funktion zum Füllen der DataGrids mit Windows Update-Einstellungen
function Update-WindowsUpdateSettingsUI {
    param (
        $GUI
    )
    
    try {
        Write-Host "Aktualisiere Windows Update-Einstellungen-Ansicht..." "INFO"
        
        # Windows Update Einstellungen abrufen
        $wuSettings = Get-WindowsUpdateRegistrySettings -SettingsType "WU"
        
        # Windows Update\AU Einstellungen abrufen
        $wuauSettings = Get-WindowsUpdateRegistrySettings -SettingsType "WUAU"
        
        # Aktuelle Einstellungen abrufen
        $currentSettings = Get-WindowsUpdateRegistrySettings -SettingsType "Current"
        
        # Windows Update DataGrid aktualisieren
        if ($GUI.ContainsKey("gridWUSettings")) {
            $GUI["gridWUSettings"].ItemsSource = $null # Wichtig: ItemsSource zuerst löschen
            
            # Wenn keine Einstellungen vorhanden sind, Platzhalter einfügen
            if ($wuSettings.Count -eq 0) {
                $placeholderWU = @([PSCustomObject]@{
                    Key = "KEIN WSUS ERKANNT"
                    Value = "-"
                    Description = "Keine Windows Update Server-Einstellungen konfiguriert"
                })
                $GUI["gridWUSettings"].ItemsSource = $placeholderWU
            } else {
                $GUI["gridWUSettings"].ItemsSource = $wuSettings
            }
            
            # Zeige "Keine Einstellungen" Text, wenn keine Daten verfügbar
            if ($GUI.ContainsKey("txtNoWUSettings")) {
                $GUI["txtNoWUSettings"].Visibility = "Collapsed" # Immer ausblenden, da nun der Platzhalter im Grid steht
            }
        }
        
        # Windows Update\AU DataGrid aktualisieren
        if ($GUI.ContainsKey("gridWUAUSettings")) {
            $GUI["gridWUAUSettings"].ItemsSource = $null # Wichtig: ItemsSource zuerst löschen
            
            # Wenn keine Einstellungen vorhanden sind, Platzhalter einfügen
            if ($wuauSettings.Count -eq 0) {
                $placeholderWUAU = @([PSCustomObject]@{
                    Key = "KEIN WSUS ERKANNT"
                    Value = "-"
                    Description = "Keine automatischen Update-Einstellungen konfiguriert"
                })
                $GUI["gridWUAUSettings"].ItemsSource = $placeholderWUAU
            } else {
                $GUI["gridWUAUSettings"].ItemsSource = $wuauSettings
            }
            
            # Zeige "Keine Einstellungen" Text, wenn keine Daten verfügbar
            if ($GUI.ContainsKey("txtNoWUAUSettings")) {
                $GUI["txtNoWUAUSettings"].Visibility = "Collapsed" # Immer ausblenden, da nun der Platzhalter im Grid steht
            }
        }
        
        # Aktuelle Einstellungen DataGrid aktualisieren
        if ($GUI.ContainsKey("gridCurrentWUSettings")) {
            $GUI["gridCurrentWUSettings"].ItemsSource = $null # Wichtig: ItemsSource zuerst löschen
            
            # Wenn keine Einstellungen vorhanden sind, Platzhalter einfügen
            if ($currentSettings.Count -eq 0) {
                $placeholderCurrent = @([PSCustomObject]@{
                    Key = "KEINE EINSTELLUNGEN"
                    Value = "-"
                    Description = "Keine aktuellen Windows Update-Einstellungen gefunden"
                })
                $GUI["gridCurrentWUSettings"].ItemsSource = $placeholderCurrent
            } else {
                $GUI["gridCurrentWUSettings"].ItemsSource = $currentSettings
            }
            
            # Zeige "Keine Einstellungen" Text, wenn keine Daten verfügbar
            if ($GUI.ContainsKey("txtNoCurrentSettings")) {
                $GUI["txtNoCurrentSettings"].Visibility = "Collapsed" # Immer ausblenden, da nun der Platzhalter im Grid steht
            }
        }
        
        # Status aktualisieren (falls vorhanden)
        if ($GUI.ContainsKey("txtSettingsStatus")) {
            $GUI["txtSettingsStatus"].Text = "Einstellungen aktualisiert: $(Get-Date -Format 'HH:mm:ss')"
            $GUI["txtSettingsStatus"].Visibility = "Visible"
        }
        
        Write-Host "Windows Update-Einstellungen erfolgreich aktualisiert" "SUCCESS"
        return $true
    }
    catch {
        Write-Host "Fehler beim Aktualisieren der Windows Update-Einstellungen-Ansicht: $($_.Exception.Message)" "ERROR"
        
        # Fehlerstatus anzeigen
        if ($GUI.ContainsKey("txtSettingsStatus")) {
            $GUI["txtSettingsStatus"].Text = "Fehler: $($_.Exception.Message)"
            $GUI["txtSettingsStatus"].Visibility = "Visible"
        }
        
        return $false
    }
}

# Funktion zum Zurücksetzen der Windows Update-Einstellungen
function Reset-WindowsUpdateSettings {
    try {
        # Prüfen, ob Administrator-Rechte vorhanden sind
        $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        if (-not $isAdmin) {
            throw "Administrator-Rechte erforderlich, um Windows Update-Einstellungen zurückzusetzen"
        }
        
        # Bestätigung vom Benutzer einholen
        $confirmation = [System.Windows.MessageBox]::Show(
            "Möchten Sie die Windows Update-Einstellungen zurücksetzen? Dies entfernt alle WSUS-spezifischen Konfigurationen und stellt die Standardwerte wieder her.",
            "Einstellungen zurücksetzen",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Warning
        )
        
        if ($confirmation -ne "Yes") {
            return $false
        }
        
        # Windows Update Dienste stoppen
        Write-Host "Stoppe Windows Update-Dienste..." "INFO"
        Stop-Service -Name wuauserv, BITS -Force -ErrorAction SilentlyContinue
        
        # Windows Update-Einstellungen zurücksetzen
        Write-Host "Entferne Windows Update-Richtlinien aus der Registry..." "INFO"
        
        # Registry-Schlüssel für WindowsUpdate-Richtlinien entfernen
        Remove-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate" -Recurse -Force -ErrorAction SilentlyContinue
        
        # Software Distribution-Ordner umbenennen (optional)
        if ((Test-Path -Path "$env:SystemRoot\SoftwareDistribution") -and -not (Test-Path -Path "$env:SystemRoot\SoftwareDistribution.old")) {
            Write-Host "Benenne SoftwareDistribution-Ordner um..." "INFO"
            Rename-Item -Path "$env:SystemRoot\SoftwareDistribution" -NewName "SoftwareDistribution.old" -ErrorAction SilentlyContinue
        }
        
        # Windows Update Dienste neu starten
        Write-Host "Starte Windows Update-Dienste neu..." "INFO"
        Start-Service -Name BITS, wuauserv -ErrorAction SilentlyContinue
        
        # Windows Update-Client zurücksetzen
        Write-Host "Setze Windows Update-Client zurück..." "INFO"
        $process = Start-Process -FilePath "wuauclt.exe" -ArgumentList "/resetauthorization /detectnow" -Wait -PassThru -WindowStyle Hidden
        
        Write-Host "Windows Update-Einstellungen wurden erfolgreich zurückgesetzt" "SUCCESS"
        return $true
    }
    catch {
        Write-Host "Fehler beim Zurücksetzen der Windows Update-Einstellungen: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

# WICHTIG: Event-Handler für die Windows Update-Einstellungen-Buttons registrieren
# Diese Funktion muss nach dem Laden der GUI aufgerufen werden
function Register-WusaSettingsEventHandlers {
    param (
        $GUI,
        $Window
    )
    
    try {
        # XAML-definierte Clickhandler mit PowerShell-Funktionen verknüpfen
        # Statt Event-Handler zu überschreiben, registrieren wir die PowerShell-Funktionen für diese Events
        
        # btnRefreshWuSettings_Click Event definieren
        $Window.Add_SourceInitialized({
            # Event-Handler für Refresh-Button registrieren
            try {
                if ($null -ne $Window.FindName("btnRefreshWuSettings")) {
                    $refreshButton = $Window.FindName("btnRefreshWuSettings")
                    $refreshButton.Add_Click({
                        try {
                            Write-Host "Aktualisiere Windows Update-Einstellungen nach Button-Klick..." "INFO"
                            $result = Update-WindowsUpdateSettingsUI -GUI $script:GUI
                            
                            if ($result -and $script:GUI.ContainsKey("txtSettingsStatus")) {
                                $script:GUI["txtSettingsStatus"].Text = "Einstellungen erfolgreich aktualisiert: $(Get-Date -Format 'HH:mm:ss')"
                                $script:GUI["txtSettingsStatus"].Visibility = "Visible"
                            }
                        }
                        catch {
                            Write-Host "Fehler beim Aktualisieren der Windows Update-Einstellungen: $($_.Exception.Message)" "ERROR"
                            if ($script:GUI.ContainsKey("txtSettingsStatus")) {
                                $script:GUI["txtSettingsStatus"].Text = "Fehler: $($_.Exception.Message)"
                                $script:GUI["txtSettingsStatus"].Visibility = "Visible"
                            }
                            
                            [System.Windows.MessageBox]::Show(
                                "Fehler beim Aktualisieren der Windows Update-Einstellungen: $($_.Exception.Message)",
                                "Fehler",
                                [System.Windows.MessageBoxButton]::OK,
                                [System.Windows.MessageBoxImage]::Error
                            )
                        }
                    })
                    Write-Host "Event-Handler für btnRefreshWuSettings erfolgreich registriert" "INFO"
                } else {
                    Write-Host "Button 'btnRefreshWuSettings' nicht gefunden in GUI" "WARNING"
                }
            } catch {
                Write-Host "Fehler beim Registrieren des Refresh-Button-Handlers: $($_.Exception.Message)" "ERROR"
            }
            
            # Event-Handler für Reset-Button registrieren
            try {
                if ($null -ne $Window.FindName("btnResetWuSettings")) {
                    $resetButton = $Window.FindName("btnResetWuSettings")
                    $resetButton.Add_Click({
                        try {
                            Write-Host "Setze Windows Update-Einstellungen zurück nach Button-Klick..." "INFO"
                            $result = Reset-WindowsUpdateSettings
                            
                            if ($result) {
                                Write-Host "Windows Update-Einstellungen wurden zurückgesetzt" "SUCCESS"
                                
                                # UI aktualisieren
                                Update-WindowsUpdateSettingsUI -GUI $script:GUI
                                
                                if ($script:GUI.ContainsKey("txtSettingsStatus")) {
                                    $script:GUI["txtSettingsStatus"].Text = "Einstellungen erfolgreich zurückgesetzt: $(Get-Date -Format 'HH:mm:ss')"
                                    $script:GUI["txtSettingsStatus"].Visibility = "Visible"
                                }
                                
                                [System.Windows.MessageBox]::Show(
                                    "Windows Update-Einstellungen wurden erfolgreich zurückgesetzt. Ein Neustart wird empfohlen, um die Änderungen zu übernehmen.",
                                    "Einstellungen zurückgesetzt",
                                    [System.Windows.MessageBoxButton]::OK,
                                    [System.Windows.MessageBoxImage]::Information
                                )
                            }
                        }
                        catch {
                            Write-Host "Fehler beim Zurücksetzen der Windows Update-Einstellungen: $($_.Exception.Message)" "ERROR"
                            if ($script:GUI.ContainsKey("txtSettingsStatus")) {
                                $script:GUI["txtSettingsStatus"].Text = "Fehler beim Zurücksetzen: $($_.Exception.Message)"
                                $script:GUI["txtSettingsStatus"].Visibility = "Visible"
                            }
                            
                            [System.Windows.MessageBox]::Show(
                                "Fehler beim Zurücksetzen der Windows Update-Einstellungen: $($_.Exception.Message)",
                                "Fehler",
                                [System.Windows.MessageBoxButton]::OK,
                                [System.Windows.MessageBoxImage]::Error
                            )
                        }
                    })
                    Write-Host "Event-Handler für btnResetWuSettings erfolgreich registriert" "INFO"
                } else {
                    Write-Host "Button 'btnResetWuSettings' nicht gefunden in GUI" "WARNING"
                }
            } catch {
                Write-Host "Fehler beim Registrieren des Reset-Button-Handlers: $($_.Exception.Message)" "ERROR"
            }
        })
        
        Write-Host "Alle WUSA-Settings Event-Handler erfolgreich registriert" "SUCCESS"
        return $true
    }
    catch {
        Write-Host "Fehler beim Registrieren der WUSA-Settings Event-Handler: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

# Funktion zum Anzeigen und Initialisieren des WUSASettings-Tabs
function Show-WUSASettings {
    param(
        $GUI
    )
    
    try {
        Write-Host "Initialisiere WUSA Settings Tab..." "INFO"
        
        # Status zurücksetzen
        if ($GUI.ContainsKey("txtSettingsStatus")) {
            $GUI["txtSettingsStatus"].Text = "Lade Windows Update-Einstellungen..."
            $GUI["txtSettingsStatus"].Visibility = "Visible"
        }
        
        # Tab anzeigen und alle anderen ausblenden
        $tabsToHide = @("Dashboard", "Settings", "Troubleshooting", "ServiceMonitor", 
                       "SystemStats", "LogViewer", "UpdateManager", "UpdateHistory", 
                       "Scheduler", "Toolbox", "SystemReport")
        
        foreach ($tab in $tabsToHide) {
            if ($GUI.ContainsKey($tab)) {
                $GUI[$tab].Visibility = "Collapsed"
            }
        }
        
        # WUSA-Settings-Tab anzeigen
        if ($GUI.ContainsKey("WUSASettings")) {
            $GUI["WUSASettings"].Visibility = "Visible"
        }
        
        # Windows Update-Einstellungen initialisieren/aktualisieren
# Dieser Event-Handler wird nun innerhalb von Initialize-GUI registriert
# Dadurch wird sichergestellt, dass die GUI bereits geladen ist
        return $true
    }
    catch {
        Write-Host "Fehler beim Initialisieren des WUSA Settings Tab: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

# Event-Handler für den Settings-Button zum Umschalten auf den WUSA Settings-Tab
if ($GUI.ContainsKey("btnWUSASettings") -and $null -ne $GUI["btnWUSASettings"]) {
    # Bestehenden Event-Handler entfernen
# Wenn der Tab beim Start angezeigt werden soll, aktivieren wir ihn hier
# Show-WUSASettings -GUI $script:GUI

# Start GUI
Initialize-GUI "Event-Handler für btnWUSASettings (WUSA Settings Tab) neu registriert" "INFO"
}

# Start GUI
Initialize-GUI