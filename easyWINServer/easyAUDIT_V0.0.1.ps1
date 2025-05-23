<#
.SYNOPSIS
    System-Reporting und Audit-Tool
.DESCRIPTION
    Dieses Tool sammelt umfangreiche Systeminformationen,
    bereitet diese in einer GUI auf und ermöglicht den Export als HTML-Bericht.
.NOTES
    Version: 0.0.1
    Autor: System-Administrator
    Erstellungsdatum: $(Get-Date -Format "dd.MM.yyyy")
    Voraussetzungen: PowerShell 5.1 oder höher, Administratorrechte
#>

#region Logging-Funktionen - muss zuerst definiert werden
# Globale Logdatei-Variable
$script:logFile = "$PSScriptRoot\audit_log.txt"
$script:fallbackLogFile = "$env:TEMP\audit_fallback_log.txt"

function Write-Log {
    param (
        [string]$Message,
        [string]$Level = "INFO"
    )
    
    try {
        if ([string]::IsNullOrEmpty($Message)) {
            $Message = "Leere Lognachricht"
        }

        # Sonderzeichen filtern - nur druckbare ASCII-Zeichen zulassen
        $filteredMessage = [System.Text.RegularExpressions.Regex]::Replace(
            $Message, 
            '[^\x20-\x7E]', 
            ' '
        )
        
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logEntry = "[$timestamp] [$Level] $filteredMessage"
        
        # In Datei schreiben mit Fehlerbehandlung
        try {
            if ($null -ne $script:logFile -and $script:logFile -ne "") {
                Add-Content -Path $script:logFile -Value $logEntry -ErrorAction Stop
            }
        }
        catch {
            # Fallback in temporäre Datei
            try {
                Add-Content -Path $script:fallbackLogFile -Value "[$timestamp] [FALLBACK] $filteredMessage" -ErrorAction SilentlyContinue
            }
            catch {
                # Silent catch - keine weitere Aktion, um endlose Loops zu vermeiden
            }
        }
        
        # Ausgabe im PowerShell-Fenster mit Farbe je nach Level
        switch ($Level) {
            "ERROR" { Write-Host $logEntry -ForegroundColor Red }
            "WARNING" { Write-Host $logEntry -ForegroundColor Yellow }
            "SUCCESS" { Write-Host $logEntry -ForegroundColor Green }
            default { Write-Host $logEntry -ForegroundColor White }
        }
    }
    catch {
        # Absoluter Fallback als einfache Konsolen-Ausgabe
        $failMsg = "[CRITICAL-FALLBACK] Logging fehlgeschlagen: $($_.Exception.Message)"
        Write-Host $failMsg -ForegroundColor Red
    }
}
#endregion

#region Initialisierung und Module
# Administrator-Rechte prüfen
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    try {
        $errorMessage = "Das Skript benötigt Administratorrechte. Bitte starten Sie PowerShell als Administrator."
        [System.Windows.Forms.MessageBox]::Show($errorMessage, "Administratorrechte erforderlich", "OK", "Error")
    }
    catch {
        Write-Host "Das Skript benötigt Administratorrechte. Bitte starten Sie PowerShell als Administrator." -ForegroundColor Red
    }
    exit
}

# Benötigte Module importieren
try {
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms
    Write-Log "UI-Module erfolgreich geladen" -Level "INFO"
}
catch {
    Write-Log "Fehler beim Laden der UI-Module: $($_.Exception.Message)" -Level "ERROR"
    Write-Host "Ein Fehler ist beim Laden der UI-Module aufgetreten. Das Tool kann nicht gestartet werden." -ForegroundColor Red
    exit
}

# Globale Variablen und Konfiguration
$script:appName = "Easy System Audit Tool"
$script:themeColor = "#007ACC"
$script:reportData = @{ }
$script:isServer = $false
#endregion

#region Self-Diagnostics
function Test-RequiredAssets {
    try {
        $assetsPath = "$PSScriptRoot\assets"
        
        if (-not (Test-Path $assetsPath)) {
            # Erstelle Assets-Verzeichnis, falls nicht vorhanden
            try {
                New-Item -Path $assetsPath -ItemType Directory -Force | Out-Null
                Write-Log "Assets-Verzeichnis wurde erstellt: $assetsPath" -Level "INFO"
            }
            catch {
                Write-Log "Konnte Assets-Verzeichnis nicht erstellen: $($_.Exception.Message)" -Level "WARNING"
            }
        }
        
        $requiredIcons = @("info.png", "settings.png", "close.png")
        $missingIcons = @()

        foreach ($icon in $requiredIcons) {
            if (-not (Test-Path "$assetsPath\$icon")) {
                $missingIcons += $icon
            }
        }

        if ($missingIcons.Count -gt 0) {
            Write-Log "WARNUNG: Fehlende Icons: $($missingIcons -join ', ')" -Level "WARNING"
            return $false
        }
        
        return $true
    }
    catch {
        Write-Log "Fehler bei der Überprüfung der benötigten Assets: $_" -Level "ERROR"
        return $false
    }
}

# Systemtyp bestimmen (Server oder Client)
function Get-SystemType {
    try {
        $os = Get-WmiObject -Class Win32_OperatingSystem
        if ($os.ProductType -eq 1) {
            # Workstation
            $script:isServer = $false
            Write-Log "System identifiziert als Client: $($os.Caption)" -Level "INFO"
            return "Client"
        }
        else {
            # Server
            $script:isServer = $true
            Write-Log "System identifiziert als Server: $($os.Caption)" -Level "INFO"
            return "Server"
        }
    }
    catch {
        Write-Log "Fehler bei der Bestimmung des Systemtyps: $($_.Exception.Message)" -Level "ERROR"
        # Im Zweifelsfall als Client behandeln
        $script:isServer = $false
        return "Client"
    }
}
#endregion

#region XAML GUI Definition
[xml]$xaml = @'
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Easy System Audit Tool" 
    Height="700" 
    Width="1000" 
    WindowStartupLocation="CenterScreen"
    ResizeMode="CanResize"
    MinWidth="800"
    MinHeight="600">
    
    <Window.Resources>
        <Style x:Key="NavigationButton" TargetType="Button">
            <Setter Property="Height" Value="40"/>
            <Setter Property="Margin" Value="5"/>
            <Setter Property="Background" Value="#007ACC"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#005A9E"/>
                </Trigger>
                <Trigger Property="IsPressed" Value="True">
                    <Setter Property="Background" Value="#003E6B"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        
        <Style x:Key="StandardButton" TargetType="Button">
            <Setter Property="Height" Value="30"/>
            <Setter Property="Padding" Value="15,0"/>
            <Setter Property="Background" Value="#007ACC"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#005A9E"/>
                </Trigger>
                <Trigger Property="IsPressed" Value="True">
                    <Setter Property="Background" Value="#003E6B"/>
                </Trigger>
                <Trigger Property="IsEnabled" Value="False">
                    <Setter Property="Background" Value="#CCCCCC"/>
                    <Setter Property="Foreground" Value="#666666"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        
        <Style x:Key="TabButtonStyle" TargetType="Button">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="15,10"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#E0E0E0"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        
        <Style x:Key="TabActiveButtonStyle" TargetType="Button">
            <Setter Property="Background" Value="#007ACC"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="15,10"/>
            <Setter Property="Cursor" Value="Hand"/>
        </Style>
    </Window.Resources>
    
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="60"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="40"/>
        </Grid.RowDefinitions>
        
        <!-- Header -->
        <Border Grid.Row="0" Background="#007ACC">
            <Grid>
                <StackPanel Orientation="Horizontal" VerticalAlignment="Center" Margin="20,0,0,0">
                    <TextBlock Text="Easy System Audit Tool" FontSize="20" Foreground="White" VerticalAlignment="Center"/>
                </StackPanel>
                <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,0,10,0">
                    <Button x:Name="btnSettings" Width="30" Height="30" Margin="5" Style="{StaticResource TabButtonStyle}" ToolTip="Einstellungen">
                        <Image Source="assets/settings.png" Width="16" Height="16"/>
                    </Button>
                    <Button x:Name="btnInfo" Width="30" Height="30" Margin="5" Style="{StaticResource TabButtonStyle}" ToolTip="Info">
                        <Image Source="assets/info.png" Width="16" Height="16"/>
                    </Button>
                </StackPanel>
            </Grid>
        </Border>
        
        <!-- Content Area -->
        <Grid Grid.Row="1">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="200"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            
            <!-- Navigation Menu -->
            <Border Background="#F0F0F0" Grid.Column="0">
                <StackPanel Margin="0,10,0,0">
                    <Button x:Name="btnSystemInfo" Content="Systeminformationen" Style="{StaticResource TabButtonStyle}" HorizontalContentAlignment="Left"/>
                    <Button x:Name="btnHardware" Content="Hardware" Style="{StaticResource TabButtonStyle}" HorizontalContentAlignment="Left"/>
                    <Button x:Name="btnSoftware" Content="Software" Style="{StaticResource TabButtonStyle}" HorizontalContentAlignment="Left"/>
                    <Button x:Name="btnNetwork" Content="Netzwerk" Style="{StaticResource TabButtonStyle}" HorizontalContentAlignment="Left"/>
                    <Button x:Name="btnUsers" Content="Benutzer und Gruppen" Style="{StaticResource TabButtonStyle}" HorizontalContentAlignment="Left"/>
                    <Button x:Name="btnServices" Content="Dienste" Style="{StaticResource TabButtonStyle}" HorizontalContentAlignment="Left"/>
                    <Button x:Name="btnEventLogs" Content="Ereignisprotokolle" Style="{StaticResource TabButtonStyle}" HorizontalContentAlignment="Left"/>
                    <Button x:Name="btnServerRoles" Content="Serverrollen" Style="{StaticResource TabButtonStyle}" HorizontalContentAlignment="Left"/>
                </StackPanel>
            </Border>
            
            <!-- Main Content -->
            <Grid Grid.Column="1" Margin="10">
                <!-- System Info Panel -->
                <Grid x:Name="pnlSystemInfo" Visibility="Collapsed">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <TextBlock Text="Systeminformationen" FontSize="18" FontWeight="Bold" Grid.Row="0" Margin="0,0,0,10"/>
                    
                    <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto">
                        <StackPanel>
                            <GroupBox Header="Betriebssystem" Margin="0,0,0,10" Padding="5">
                                <Grid>
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="180"/>
                                        <ColumnDefinition Width="*"/>
                                    </Grid.ColumnDefinitions>
                                    <Grid.RowDefinitions>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                    </Grid.RowDefinitions>
                                    
                                    <TextBlock Text="Name:" Grid.Row="0" Grid.Column="0" Margin="0,5,0,5"/>
                                    <TextBlock x:Name="txtOSName" Text="" Grid.Row="0" Grid.Column="1" Margin="0,5,0,5"/>
                                    
                                    <TextBlock Text="Version:" Grid.Row="1" Grid.Column="0" Margin="0,5,0,5"/>
                                    <TextBlock x:Name="txtOSVersion" Text="" Grid.Row="1" Grid.Column="1" Margin="0,5,0,5"/>
                                    
                                    <TextBlock Text="Build:" Grid.Row="2" Grid.Column="0" Margin="0,5,0,5"/>
                                    <TextBlock x:Name="txtOSBuild" Text="" Grid.Row="2" Grid.Column="1" Margin="0,5,0,5"/>
                                    
                                    <TextBlock Text="Architektur:" Grid.Row="3" Grid.Column="0" Margin="0,5,0,5"/>
                                    <TextBlock x:Name="txtOSArchitecture" Text="" Grid.Row="3" Grid.Column="1" Margin="0,5,0,5"/>
                                    
                                    <TextBlock Text="Installationsdatum:" Grid.Row="4" Grid.Column="0" Margin="0,5,0,5"/>
                                    <TextBlock x:Name="txtOSInstallDate" Text="" Grid.Row="4" Grid.Column="1" Margin="0,5,0,5"/>
                                    
                                    <TextBlock Text="Letzte Aktualisierung:" Grid.Row="5" Grid.Column="0" Margin="0,5,0,5"/>
                                    <TextBlock x:Name="txtOSLastUpdate" Text="" Grid.Row="5" Grid.Column="1" Margin="0,5,0,5"/>
                                </Grid>
                            </GroupBox>
                            
                            <GroupBox Header="System" Margin="0,0,0,10" Padding="5">
                                <Grid>
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="180"/>
                                        <ColumnDefinition Width="*"/>
                                    </Grid.ColumnDefinitions>
                                    <Grid.RowDefinitions>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                    </Grid.RowDefinitions>
                                    
                                    <TextBlock Text="Computername:" Grid.Row="0" Grid.Column="0" Margin="0,5,0,5"/>
                                    <TextBlock x:Name="txtComputerName" Text="" Grid.Row="0" Grid.Column="1" Margin="0,5,0,5"/>
                                    
                                    <TextBlock Text="Domain/Arbeitsgruppe:" Grid.Row="1" Grid.Column="0" Margin="0,5,0,5"/>
                                    <TextBlock x:Name="txtDomain" Text="" Grid.Row="1" Grid.Column="1" Margin="0,5,0,5"/>
                                    
                                    <TextBlock Text="Systemtyp:" Grid.Row="2" Grid.Column="0" Margin="0,5,0,5"/>
                                    <TextBlock x:Name="txtSystemType" Text="" Grid.Row="2" Grid.Column="1" Margin="0,5,0,5"/>
                                    
                                    <TextBlock Text="Uptime:" Grid.Row="3" Grid.Column="0" Margin="0,5,0,5"/>
                                    <TextBlock x:Name="txtUptime" Text="" Grid.Row="3" Grid.Column="1" Margin="0,5,0,5"/>
                                    
                                    <TextBlock Text="Zeitzone:" Grid.Row="4" Grid.Column="0" Margin="0,5,0,5"/>
                                    <TextBlock x:Name="txtTimeZone" Text="" Grid.Row="4" Grid.Column="1" Margin="0,5,0,5"/>
                                </Grid>
                            </GroupBox>
                        </StackPanel>
                    </ScrollViewer>
                    
                    <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0">
                        <Button x:Name="btnRefreshSystemInfo" Content="Aktualisieren" Style="{StaticResource StandardButton}" Margin="0,0,10,0"/>
                        <Button x:Name="btnExportSystemInfo" Content="Exportieren" Style="{StaticResource StandardButton}"/>
                    </StackPanel>
                </Grid>
                
                <!-- Hardware Panel -->
                <Grid x:Name="pnlHardware" Visibility="Collapsed">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <TextBlock Text="Hardware" FontSize="18" FontWeight="Bold" Grid.Row="0" Margin="0,0,0,10"/>
                    
                    <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto">
                        <StackPanel>
                            <GroupBox Header="Prozessor" Margin="0,0,0,10" Padding="5">
                                <DataGrid x:Name="dgProcessors" AutoGenerateColumns="False" IsReadOnly="True" HeadersVisibility="Column" GridLinesVisibility="Horizontal" 
                                          CanUserSortColumns="True" CanUserResizeColumns="True" CanUserReorderColumns="True" MaxHeight="150">
                                    <DataGrid.Columns>
                                        <DataGridTextColumn Header="Name" Binding="{Binding Name}" Width="*"/>
                                        <DataGridTextColumn Header="Kerne" Binding="{Binding Cores}" Width="80"/>
                                        <DataGridTextColumn Header="Threads" Binding="{Binding Threads}" Width="80"/>
                                        <DataGridTextColumn Header="Geschwindigkeit" Binding="{Binding Speed}" Width="120"/>
                                    </DataGrid.Columns>
                                </DataGrid>
                            </GroupBox>
                            
                            <GroupBox Header="Arbeitsspeicher" Margin="0,0,0,10" Padding="5">
                                <StackPanel>
                                    <Grid>
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="180"/>
                                            <ColumnDefinition Width="*"/>
                                        </Grid.ColumnDefinitions>
                                        <Grid.RowDefinitions>
                                            <RowDefinition Height="Auto"/>
                                            <RowDefinition Height="Auto"/>
                                            <RowDefinition Height="Auto"/>
                                        </Grid.RowDefinitions>
                                        
                                        <TextBlock Text="Gesamt:" Grid.Row="0" Grid.Column="0" Margin="0,5,0,5"/>
                                        <TextBlock x:Name="txtTotalRAM" Text="" Grid.Row="0" Grid.Column="1" Margin="0,5,0,5"/>
                                        
                                        <TextBlock Text="Verfügbar:" Grid.Row="1" Grid.Column="0" Margin="0,5,0,5"/>
                                        <TextBlock x:Name="txtAvailableRAM" Text="" Grid.Row="1" Grid.Column="1" Margin="0,5,0,5"/>
                                        
                                        <TextBlock Text="Verwendung:" Grid.Row="2" Grid.Column="0" Margin="0,5,0,5"/>
                                        <StackPanel Grid.Row="2" Grid.Column="1" Orientation="Horizontal" Margin="0,5,0,5">
                                            <TextBlock x:Name="txtRAMUsage" Text=""/>
                                            <ProgressBar x:Name="pbRAMUsage" Width="200" Height="15" Maximum="100" Margin="10,0,0,0"/>
                                        </StackPanel>
                                    </Grid>
                                    
                                    <TextBlock Text="RAM Module:" Margin="0,10,0,5"/>
                                    <DataGrid x:Name="dgRAMModules" AutoGenerateColumns="False" IsReadOnly="True" HeadersVisibility="Column" GridLinesVisibility="Horizontal" 
                                              CanUserSortColumns="True" CanUserResizeColumns="True" CanUserReorderColumns="True" MaxHeight="150">
                                        <DataGrid.Columns>
                                            <DataGridTextColumn Header="Slot" Binding="{Binding Slot}" Width="80"/>
                                            <DataGridTextColumn Header="Kapazität" Binding="{Binding Capacity}" Width="100"/>
                                            <DataGridTextColumn Header="Typ" Binding="{Binding Type}" Width="100"/>
                                            <DataGridTextColumn Header="Takt" Binding="{Binding Speed}" Width="100"/>
                                            <DataGridTextColumn Header="Hersteller" Binding="{Binding Manufacturer}" Width="*"/>
                                        </DataGrid.Columns>
                                    </DataGrid>
                                </StackPanel>
                            </GroupBox>
                            
                            <GroupBox Header="Speichermedien" Margin="0,0,0,10" Padding="5">
                                <DataGrid x:Name="dgDiskDrives" AutoGenerateColumns="False" IsReadOnly="True" HeadersVisibility="Column" GridLinesVisibility="Horizontal" 
                                          CanUserSortColumns="True" CanUserResizeColumns="True" CanUserReorderColumns="True" MaxHeight="200">
                                    <DataGrid.Columns>
                                        <DataGridTextColumn Header="Laufwerk" Binding="{Binding Drive}" Width="80"/>
                                        <DataGridTextColumn Header="Modell" Binding="{Binding Model}" Width="*"/>
                                        <DataGridTextColumn Header="Typ" Binding="{Binding Type}" Width="100"/>
                                        <DataGridTextColumn Header="Größe" Binding="{Binding Size}" Width="100"/>
                                        <DataGridTextColumn Header="Freier Speicher" Binding="{Binding FreeSpace}" Width="120"/>
                                        <DataGridTextColumn Header="% Frei" Binding="{Binding PercentFree}" Width="80"/>
                                    </DataGrid.Columns>
                                </DataGrid>
                            </GroupBox>
                        </StackPanel>
                    </ScrollViewer>
                    
                    <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0">
                        <Button x:Name="btnRefreshHardware" Content="Aktualisieren" Style="{StaticResource StandardButton}" Margin="0,0,10,0"/>
                        <Button x:Name="btnExportHardware" Content="Exportieren" Style="{StaticResource StandardButton}"/>
                    </StackPanel>
                </Grid>
                
                <!-- Software Panel -->
                <Grid x:Name="pnlSoftware" Visibility="Collapsed">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <TextBlock Text="Software" FontSize="18" FontWeight="Bold" Grid.Row="0" Margin="0,0,0,10"/>
                    
                    <Grid Grid.Row="1">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>
                        
                        <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,10">
                            <TextBlock Text="Suche:" VerticalAlignment="Center" Margin="0,0,10,0"/>
                            <TextBox x:Name="txtSoftwareSearch" Width="250" Height="25" VerticalContentAlignment="Center"/>
                        </StackPanel>
                        
                        <DataGrid x:Name="dgInstalledSoftware" Grid.Row="1" AutoGenerateColumns="False" IsReadOnly="True" HeadersVisibility="Column" 
                                  GridLinesVisibility="Horizontal" CanUserSortColumns="True" CanUserResizeColumns="True" CanUserReorderColumns="True">
                            <DataGrid.Columns>
                                <DataGridTextColumn Header="Name" Binding="{Binding Name}" Width="3*"/>
                                <DataGridTextColumn Header="Version" Binding="{Binding Version}" Width="1*"/>
                                <DataGridTextColumn Header="Hersteller" Binding="{Binding Publisher}" Width="2*"/>
                                <DataGridTextColumn Header="Installationsdatum" Binding="{Binding InstallDate}" Width="1.5*"/>
                                <DataGridTextColumn Header="Größe" Binding="{Binding Size}" Width="1*"/>
                            </DataGrid.Columns>
                        </DataGrid>
                    </Grid>
                    
                    <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0">
                        <Button x:Name="btnRefreshSoftware" Content="Aktualisieren" Style="{StaticResource StandardButton}" Margin="0,0,10,0"/>
                        <Button x:Name="btnExportSoftware" Content="Exportieren" Style="{StaticResource StandardButton}"/>
                    </StackPanel>
                </Grid>
                
                <!-- Network Panel -->
                <Grid x:Name="pnlNetwork" Visibility="Collapsed">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <TextBlock Text="Netzwerk" FontSize="18" FontWeight="Bold" Grid.Row="0" Margin="0,0,0,10"/>
                    
                    <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto">
                        <StackPanel>
                            <GroupBox Header="Netzwerkadapter" Margin="0,0,0,10" Padding="5">
                                <DataGrid x:Name="dgNetworkAdapters" AutoGenerateColumns="False" IsReadOnly="True" HeadersVisibility="Column" GridLinesVisibility="Horizontal" 
                                          CanUserSortColumns="True" CanUserResizeColumns="True" CanUserReorderColumns="True" MaxHeight="200">
                                    <DataGrid.Columns>
                                        <DataGridTextColumn Header="Name" Binding="{Binding Name}" Width="2*"/>
                                        <DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="1*"/>
                                        <DataGridTextColumn Header="IP-Adresse" Binding="{Binding IPAddress}" Width="1.2*"/>
                                        <DataGridTextColumn Header="Subnetzmaske" Binding="{Binding SubnetMask}" Width="1.2*"/>
                                        <DataGridTextColumn Header="Gateway" Binding="{Binding Gateway}" Width="1.2*"/>
                                        <DataGridTextColumn Header="DNS-Server" Binding="{Binding DNSServers}" Width="1.5*"/>
                                        <DataGridTextColumn Header="MAC-Adresse" Binding="{Binding MACAddress}" Width="1.5*"/>
                                    </DataGrid.Columns>
                                </DataGrid>
                            </GroupBox>
                            
                            <GroupBox Header="Netzwerkverbindungen" Margin="0,0,0,10" Padding="5">
                                <DataGrid x:Name="dgNetworkConnections" AutoGenerateColumns="False" IsReadOnly="True" HeadersVisibility="Column" GridLinesVisibility="Horizontal" 
                                          CanUserSortColumns="True" CanUserResizeColumns="True" CanUserReorderColumns="True" MaxHeight="200">
                                    <DataGrid.Columns>
                                        <DataGridTextColumn Header="Protokoll" Binding="{Binding Protocol}" Width="1*"/>
                                        <DataGridTextColumn Header="Lokale Adresse" Binding="{Binding LocalAddress}" Width="2*"/>
                                        <DataGridTextColumn Header="Remote-Adresse" Binding="{Binding RemoteAddress}" Width="2*"/>
                                        <DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="1*"/>
                                        <DataGridTextColumn Header="Prozess" Binding="{Binding Process}" Width="1.5*"/>
                                    </DataGrid.Columns>
                                </DataGrid>
                            </GroupBox>
                            
                            <GroupBox Header="Firewallregeln" Margin="0,0,0,10" Padding="5">
                                <DataGrid x:Name="dgFirewallRules" AutoGenerateColumns="False" IsReadOnly="True" HeadersVisibility="Column" GridLinesVisibility="Horizontal" 
                                          CanUserSortColumns="True" CanUserResizeColumns="True" CanUserReorderColumns="True" MaxHeight="200">
                                    <DataGrid.Columns>
                                        <DataGridTextColumn Header="Name" Binding="{Binding Name}" Width="2*"/>
                                        <DataGridTextColumn Header="Richtung" Binding="{Binding Direction}" Width="1*"/>
                                        <DataGridTextColumn Header="Aktion" Binding="{Binding Action}" Width="1*"/>
                                        <DataGridTextColumn Header="Protokoll" Binding="{Binding Protocol}" Width="1*"/>
                                        <DataGridTextColumn Header="Lokaler Port" Binding="{Binding LocalPort}" Width="1*"/>
                                        <DataGridTextColumn Header="Remote-Port" Binding="{Binding RemotePort}" Width="1*"/>
                                        <DataGridTextColumn Header="Aktiviert" Binding="{Binding Enabled}" Width="1*"/>
                                    </DataGrid.Columns>
                                </DataGrid>
                            </GroupBox>
                        </StackPanel>
                    </ScrollViewer>
                    
                    <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0">
                        <Button x:Name="btnRefreshNetwork" Content="Aktualisieren" Style="{StaticResource StandardButton}" Margin="0,0,10,0"/>
                        <Button x:Name="btnExportNetwork" Content="Exportieren" Style="{StaticResource StandardButton}"/>
                    </StackPanel>
                </Grid>
                
                <!-- Users and Groups Panel -->
                <Grid x:Name="pnlUsers" Visibility="Collapsed">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <TextBlock Text="Benutzer und Gruppen" FontSize="18" FontWeight="Bold" Grid.Row="0" Margin="0,0,0,10"/>
                    
                    <TabControl Grid.Row="1" Margin="0,10,0,0">
                        <TabItem Header="Lokale Benutzer">
                            <DataGrid x:Name="dgLocalUsers" AutoGenerateColumns="False" IsReadOnly="True" HeadersVisibility="Column" GridLinesVisibility="Horizontal"
                                      CanUserSortColumns="True" CanUserResizeColumns="True" CanUserReorderColumns="True">
                                <DataGrid.Columns>
                                    <DataGridTextColumn Header="Benutzername" Binding="{Binding Name}" Width="1.5*"/>
                                    <DataGridTextColumn Header="Vollständiger Name" Binding="{Binding FullName}" Width="2*"/>
                                    <DataGridTextColumn Header="Beschreibung" Binding="{Binding Description}" Width="2*"/>
                                    <DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="1*"/>
                                    <DataGridTextColumn Header="Letzte Anmeldung" Binding="{Binding LastLogon}" Width="1.5*"/>
                                </DataGrid.Columns>
                            </DataGrid>
                        </TabItem>
                        <TabItem Header="Lokale Gruppen">
                            <DataGrid x:Name="dgLocalGroups" AutoGenerateColumns="False" IsReadOnly="True" HeadersVisibility="Column" GridLinesVisibility="Horizontal"
                                      CanUserSortColumns="True" CanUserResizeColumns="True" CanUserReorderColumns="True">
                                <DataGrid.Columns>
                                    <DataGridTextColumn Header="Gruppenname" Binding="{Binding Name}" Width="1.5*"/>
                                    <DataGridTextColumn Header="Beschreibung" Binding="{Binding Description}" Width="2*"/>
                                    <DataGridTextColumn Header="Anzahl Mitglieder" Binding="{Binding MemberCount}" Width="1*"/>
                                </DataGrid.Columns>
                            </DataGrid>
                        </TabItem>
                        <TabItem Header="Gruppenmitgliedschaften">
                            <Grid>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="*"/>
                                </Grid.RowDefinitions>
                                
                                <StackPanel Orientation="Horizontal" Margin="0,5,0,10">
                                    <TextBlock Text="Benutzer/Gruppe:" VerticalAlignment="Center" Margin="0,0,10,0"/>
                                    <ComboBox x:Name="cmbUserGroupSelect" Width="250" Height="25" VerticalContentAlignment="Center"/>
                                    <Button x:Name="btnShowMembership" Content="Anzeigen" Style="{StaticResource StandardButton}" Margin="10,0,0,0"/>
                                </StackPanel>
                                
                                <DataGrid x:Name="dgMemberships" Grid.Row="1" AutoGenerateColumns="False" IsReadOnly="True" HeadersVisibility="Column" GridLinesVisibility="Horizontal"
                                          CanUserSortColumns="True" CanUserResizeColumns="True" CanUserReorderColumns="True">
                                    <DataGrid.Columns>
                                        <DataGridTextColumn Header="Gruppenname" Binding="{Binding GroupName}" Width="1.5*"/>
                                        <DataGridTextColumn Header="Beschreibung" Binding="{Binding Description}" Width="2*"/>
                                        <DataGridTextColumn Header="Typ" Binding="{Binding Type}" Width="1*"/>
                                    </DataGrid.Columns>
                                </DataGrid>
                            </Grid>
                        </TabItem>
                    </TabControl>
                    
                    <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0">
                        <Button x:Name="btnRefreshUsers" Content="Aktualisieren" Style="{StaticResource StandardButton}" Margin="0,0,10,0"/>
                        <Button x:Name="btnExportUsers" Content="Exportieren" Style="{StaticResource StandardButton}"/>
                    </StackPanel>
                </Grid>
                
                <!-- Services Panel -->
                <Grid x:Name="pnlServices" Visibility="Collapsed">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <TextBlock Text="Dienste" FontSize="18" FontWeight="Bold" Grid.Row="0" Margin="0,0,0,10"/>
                    
                    <Grid Grid.Row="1">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>
                        
                        <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,10">
                            <TextBlock Text="Filter:" VerticalAlignment="Center" Margin="0,0,10,0"/>
                            <ComboBox x:Name="cmbServiceStatus" Width="150" Height="25" VerticalContentAlignment="Center">
                                <ComboBoxItem Content="Alle Dienste" IsSelected="True"/>
                                <ComboBoxItem Content="Ausgeführt"/>
                                <ComboBoxItem Content="Angehalten"/>
                                <ComboBoxItem Content="Automatisch"/>
                                <ComboBoxItem Content="Manuell"/>
                                <ComboBoxItem Content="Deaktiviert"/>
                            </ComboBox>
                            <TextBlock Text="Suche:" VerticalAlignment="Center" Margin="20,0,10,0"/>
                            <TextBox x:Name="txtServiceSearch" Width="250" Height="25" VerticalContentAlignment="Center"/>
                        </StackPanel>
                        
                        <DataGrid x:Name="dgServices" Grid.Row="1" AutoGenerateColumns="False" IsReadOnly="True" HeadersVisibility="Column" 
                                  GridLinesVisibility="Horizontal" CanUserSortColumns="True" CanUserResizeColumns="True" CanUserReorderColumns="True">
                            <DataGrid.Columns>
                                <DataGridTextColumn Header="Dienst" Binding="{Binding DisplayName}" Width="2*"/>
                                <DataGridTextColumn Header="Name" Binding="{Binding ServiceName}" Width="1.5*"/>
                                <DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="1*"/>
                                <DataGridTextColumn Header="Starttyp" Binding="{Binding StartupType}" Width="1*"/>
                                <DataGridTextColumn Header="Anmeldung als" Binding="{Binding LogOnAs}" Width="1.5*"/>
                                <DataGridTextColumn Header="Beschreibung" Binding="{Binding Description}" Width="3*"/>
                            </DataGrid.Columns>
                        </DataGrid>
                    </Grid>
                    
                    <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0">
                        <Button x:Name="btnRefreshServices" Content="Aktualisieren" Style="{StaticResource StandardButton}" Margin="0,0,10,0"/>
                        <Button x:Name="btnExportServices" Content="Exportieren" Style="{StaticResource StandardButton}"/>
                    </StackPanel>
                </Grid>
                
                <!-- Event Logs Panel -->
                <Grid x:Name="pnlEventLogs" Visibility="Collapsed">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <TextBlock Text="Ereignisprotokolle" FontSize="18" FontWeight="Bold" Grid.Row="0" Margin="0,0,0,10"/>
                    
                    <Grid Grid.Row="1">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>
                        
                        <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,10">
                            <TextBlock Text="Protokoll:" VerticalAlignment="Center" Margin="0,0,10,0"/>
                            <ComboBox x:Name="cmbEventLog" Width="150" Height="25" VerticalContentAlignment="Center">
                                <ComboBoxItem Content="System" IsSelected="True"/>
                                <ComboBoxItem Content="Anwendung"/>
                                <ComboBoxItem Content="Sicherheit"/>
                            </ComboBox>
                            <TextBlock Text="Typ:" VerticalAlignment="Center" Margin="20,0,10,0"/>
                            <ComboBox x:Name="cmbEventType" Width="150" Height="25" VerticalContentAlignment="Center">
                                <ComboBoxItem Content="Alle" IsSelected="True"/>
                                <ComboBoxItem Content="Fehler"/>
                                <ComboBoxItem Content="Warnung"/>
                                <ComboBoxItem Content="Information"/>
                            </ComboBox>
                            <TextBlock Text="Zeitraum:" VerticalAlignment="Center" Margin="20,0,10,0"/>
                            <ComboBox x:Name="cmbEventTimespan" Width="150" Height="25" VerticalContentAlignment="Center">
                                <ComboBoxItem Content="Letzte 24 Stunden" IsSelected="True"/>
                                <ComboBoxItem Content="Letzte 7 Tage"/>
                                <ComboBoxItem Content="Letzte 30 Tage"/>
                                <ComboBoxItem Content="Alle"/>
                            </ComboBox>
                        </StackPanel>
                        
                        <DataGrid x:Name="dgEventLogs" Grid.Row="1" AutoGenerateColumns="False" IsReadOnly="True" HeadersVisibility="Column" 
                                  GridLinesVisibility="Horizontal" CanUserSortColumns="True" CanUserResizeColumns="True" CanUserReorderColumns="True">
                            <DataGrid.Columns>
                                <DataGridTextColumn Header="Zeitstempel" Binding="{Binding TimeGenerated}" Width="1.5*"/>
                                <DataGridTextColumn Header="Ereignis-ID" Binding="{Binding EventID}" Width="1*"/>
                                <DataGridTextColumn Header="Level" Binding="{Binding EntryType}" Width="1*"/>
                                <DataGridTextColumn Header="Quelle" Binding="{Binding Source}" Width="1.5*"/>
                                <DataGridTextColumn Header="Nachricht" Binding="{Binding Message}" Width="4*"/>
                            </DataGrid.Columns>
                        </DataGrid>
                    </Grid>
                    
                    <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0">
                        <Button x:Name="btnRefreshEventLogs" Content="Aktualisieren" Style="{StaticResource StandardButton}" Margin="0,0,10,0"/>
                        <Button x:Name="btnExportEventLogs" Content="Exportieren" Style="{StaticResource StandardButton}"/>
                    </StackPanel>
                </Grid>
                
                <!-- Server Roles Panel -->
                <Grid x:Name="pnlServerRoles" Visibility="Collapsed">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <TextBlock Text="Serverrollen" FontSize="18" FontWeight="Bold" Grid.Row="0" Margin="0,0,0,10"/>
                    
                    <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>
                            
                            <!-- Linke Spalte: Rollen -->
                            <GroupBox Header="Serverrollen" Grid.Column="0" Margin="0,0,5,0" Padding="5">
                                <StackPanel>
                                    <CheckBox x:Name="chkRoleADDS" Content="Active Directory Domain Services" Margin="0,5,0,5"/>
                                    <CheckBox x:Name="chkRoleADCS" Content="Active Directory Certificate Services" Margin="0,5,0,5"/>
                                    <CheckBox x:Name="chkRoleDNS" Content="DNS-Server" Margin="0,5,0,5"/>
                                    <CheckBox x:Name="chkRoleDHCP" Content="DHCP-Server" Margin="0,5,0,5"/>
                                    <CheckBox x:Name="chkRoleFileServer" Content="Dateiserver" Margin="0,5,0,5"/>
                                    <CheckBox x:Name="chkRolePrintServer" Content="Druckserver" Margin="0,5,0,5"/>
                                    <CheckBox x:Name="chkRoleWebServer" Content="Webserver" Margin="0,5,0,5"/>
                                    <CheckBox x:Name="chkRoleWSUS" Content="Windows Server Update Services" Margin="0,5,0,5"/>
                                    <CheckBox x:Name="chkRoleRemoteDesktop" Content="Remote Desktop Services" Margin="0,5,0,5"/>
                                    <CheckBox x:Name="chkRoleADFS" Content="Active Directory Federation Services" Margin="0,5,0,5"/>
                                </StackPanel>
                            </GroupBox>
                            
                            <!-- Rechte Spalte: Features -->
                            <GroupBox Header="Features" Grid.Column="1" Margin="5,0,0,0" Padding="5">
                                <StackPanel>
                                    <CheckBox x:Name="chkFeatureNET" Content=".NET Framework" Margin="0,5,0,5"/>
                                    <CheckBox x:Name="chkFeatureNetworkPolicy" Content="Network Policy Server" Margin="0,5,0,5"/>
                                    <CheckBox x:Name="chkFeatureRemoteAccess" Content="Remote Access" Margin="0,5,0,5"/>
                                    <CheckBox x:Name="chkFeatureFSRM" Content="File Server Resource Manager" Margin="0,5,0,5"/>
                                    <CheckBox x:Name="chkFeatureDFS" Content="DFS-Namespaces" Margin="0,5,0,5"/>
                                    <CheckBox x:Name="chkFeatureDFSR" Content="DFS-Replikation" Margin="0,5,0,5"/>
                                    <CheckBox x:Name="chkFeatureFax" Content="Faxserver" Margin="0,5,0,5"/>
                                    <CheckBox x:Name="chkFeatureWINS" Content="WINS" Margin="0,5,0,5"/>
                                    <CheckBox x:Name="chkFeatureSMTP" Content="SMTP-Server" Margin="0,5,0,5"/>
                                    <CheckBox x:Name="chkFeatureWAS" Content="Windows-Prozessaktivierungsdienst" Margin="0,5,0,5"/>
                                </StackPanel>
                            </GroupBox>
                        </Grid>
                    </ScrollViewer>
                    
                    <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0">
                        <Button x:Name="btnRefreshServerRoles" Content="Aktualisieren" Style="{StaticResource StandardButton}" Margin="0,0,10,0"/>
                        <Button x:Name="btnExportServerRoles" Content="Exportieren" Style="{StaticResource StandardButton}"/>
                    </StackPanel>
                </Grid>
            </Grid>
        </Grid>
        
        <!-- Footer -->
        <Border Grid.Row="2" Background="#F0F0F0" BorderThickness="0,1,0,0" BorderBrush="#CCCCCC">
            <Grid>
                <StackPanel Orientation="Horizontal" Margin="10,0">
                    <TextBlock x:Name="txtStatus" Text="Bereit" VerticalAlignment="Center"/>
                </StackPanel>
                <Button x:Name="btnGenerateReport" Content="Gesamtbericht erstellen" Style="{StaticResource StandardButton}" HorizontalAlignment="Right" Margin="0,0,10,0" VerticalAlignment="Center"/>
            </Grid>
        </Border>
    </Grid>
</Window>
'@

# GUI in PowerShell laden
try {
    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $window = [Windows.Markup.XamlReader]::Load($reader)
    
    Write-Log "GUI erfolgreich geladen" -Level "INFO"
}
catch {
    Write-Log "Fehler beim Laden der GUI: $_" -Level "ERROR"
    [System.Windows.Forms.MessageBox]::Show("Ein unerwarteter Fehler ist aufgetreten. Bitte wenden Sie sich an den Support." + "`n`nDetails: " + $_.Exception.Message, "Fehler", "OK", "Error")
    exit
}
#endregion

#region GUI-Elemente an PowerShell binden
# Hauptbuttons
$btnSystemInfo = $window.FindName("btnSystemInfo")
$btnHardware = $window.FindName("btnHardware")
$btnSoftware = $window.FindName("btnSoftware")
$btnNetwork = $window.FindName("btnNetwork")
$btnUsers = $window.FindName("btnUsers")
$btnServices = $window.FindName("btnServices")
$btnEventLogs = $window.FindName("btnEventLogs")
$btnServerRoles = $window.FindName("btnServerRoles")

# Steuerungspanels
$pnlSystemInfo = $window.FindName("pnlSystemInfo")
$pnlHardware = $window.FindName("pnlHardware")
$pnlSoftware = $window.FindName("pnlSoftware")
$pnlNetwork = $window.FindName("pnlNetwork")
$pnlUsers = $window.FindName("pnlUsers")
$pnlServices = $window.FindName("pnlServices")
$pnlEventLogs = $window.FindName("pnlEventLogs")
$pnlServerRoles = $window.FindName("pnlServerRoles")

# Systeminfo-Panel Elemente
$txtOSName = $window.FindName("txtOSName")
$txtOSVersion = $window.FindName("txtOSVersion")
$txtOSBuild = $window.FindName("txtOSBuild")
$txtOSArchitecture = $window.FindName("txtOSArchitecture")
$txtOSInstallDate = $window.FindName("txtOSInstallDate")
$txtOSLastUpdate = $window.FindName("txtOSLastUpdate")
$txtComputerName = $window.FindName("txtComputerName")
$txtDomain = $window.FindName("txtDomain")
$txtSystemType = $window.FindName("txtSystemType")
$txtUptime = $window.FindName("txtUptime")
$txtTimeZone = $window.FindName("txtTimeZone")

# Hardware-Panel Elemente
$dgProcessors = $window.FindName("dgProcessors")
$txtTotalRAM = $window.FindName("txtTotalRAM")
$txtAvailableRAM = $window.FindName("txtAvailableRAM")
$txtRAMUsage = $window.FindName("txtRAMUsage")
$pbRAMUsage = $window.FindName("pbRAMUsage")
$dgRAMModules = $window.FindName("dgRAMModules")
$dgDiskDrives = $window.FindName("dgDiskDrives")

# Software-Panel Elemente
$txtSoftwareSearch = $window.FindName("txtSoftwareSearch")
$dgInstalledSoftware = $window.FindName("dgInstalledSoftware")

# Netzwerk-Panel Elemente
$dgNetworkAdapters = $window.FindName("dgNetworkAdapters")
$dgNetworkConnections = $window.FindName("dgNetworkConnections")
$dgFirewallRules = $window.FindName("dgFirewallRules")

# Benutzer-Panel Elemente
$dgLocalUsers = $window.FindName("dgLocalUsers")
$dgLocalGroups = $window.FindName("dgLocalGroups")
$cmbUserGroupSelect = $window.FindName("cmbUserGroupSelect")
$btnShowMembership = $window.FindName("btnShowMembership")
$dgMemberships = $window.FindName("dgMemberships")

# Dienste-Panel Elemente
$cmbServiceStatus = $window.FindName("cmbServiceStatus")
$txtServiceSearch = $window.FindName("txtServiceSearch")
$dgServices = $window.FindName("dgServices")

# Ereignisprotokolle-Panel Elemente
$cmbEventLog = $window.FindName("cmbEventLog")
$cmbEventType = $window.FindName("cmbEventType")
$cmbEventTimespan = $window.FindName("cmbEventTimespan")
$dgEventLogs = $window.FindName("dgEventLogs")

# Serverrollen-Panel Elemente
$chkRoleADDS = $window.FindName("chkRoleADDS")
$chkRoleADCS = $window.FindName("chkRoleADCS")
$chkRoleDNS = $window.FindName("chkRoleDNS")
$chkRoleDHCP = $window.FindName("chkRoleDHCP")
$chkRoleFileServer = $window.FindName("chkRoleFileServer")
$chkRolePrintServer = $window.FindName("chkRolePrintServer")
$chkRoleWebServer = $window.FindName("chkRoleWebServer")
$chkRoleWSUS = $window.FindName("chkRoleWSUS")
$chkRoleRemoteDesktop = $window.FindName("chkRoleRemoteDesktop")
$chkRoleADFS = $window.FindName("chkRoleADFS")

# Feature-Checkboxen
$chkFeatureNET = $window.FindName("chkFeatureNET")
$chkFeatureNetworkPolicy = $window.FindName("chkFeatureNetworkPolicy")
$chkFeatureRemoteAccess = $window.FindName("chkFeatureRemoteAccess")
$chkFeatureFSRM = $window.FindName("chkFeatureFSRM")
$chkFeatureDFS = $window.FindName("chkFeatureDFS")
$chkFeatureDFSR = $window.FindName("chkFeatureDFSR")
$chkFeatureFax = $window.FindName("chkFeatureFax")
$chkFeatureWINS = $window.FindName("chkFeatureWINS")
$chkFeatureSMTP = $window.FindName("chkFeatureSMTP")
$chkFeatureWAS = $window.FindName("chkFeatureWAS")

# Aktionsbuttons
$btnRefreshSystemInfo = $window.FindName("btnRefreshSystemInfo")
$btnExportSystemInfo = $window.FindName("btnExportSystemInfo")
$btnRefreshHardware = $window.FindName("btnRefreshHardware")
$btnExportHardware = $window.FindName("btnExportHardware")
$btnRefreshSoftware = $window.FindName("btnRefreshSoftware")
$btnExportSoftware = $window.FindName("btnExportSoftware")
$btnRefreshNetwork = $window.FindName("btnRefreshNetwork")
$btnExportNetwork = $window.FindName("btnExportNetwork")
$btnRefreshUsers = $window.FindName("btnRefreshUsers")
$btnExportUsers = $window.FindName("btnExportUsers")
$btnRefreshServices = $window.FindName("btnRefreshServices")
$btnExportServices = $window.FindName("btnExportServices")
$btnRefreshEventLogs = $window.FindName("btnRefreshEventLogs")
$btnExportEventLogs = $window.FindName("btnExportEventLogs")
$btnRefreshServerRoles = $window.FindName("btnRefreshServerRoles")
$btnExportServerRoles = $window.FindName("btnExportServerRoles")
$btnGenerateReport = $window.FindName("btnGenerateReport")

# Weitere UI-Elemente
$btnSettings = $window.FindName("btnSettings")
$btnInfo = $window.FindName("btnInfo")
$txtStatus = $window.FindName("txtStatus")
#endregion

#region GUI-Funktionen
# Funktion zum Wechseln der Ansicht
function Switch-View {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ViewName
    )
    
    try {
        # Alle Panels ausblenden
        $pnlSystemInfo.Visibility = "Collapsed"
        $pnlHardware.Visibility = "Collapsed"
        $pnlSoftware.Visibility = "Collapsed"
        $pnlNetwork.Visibility = "Collapsed"
        $pnlUsers.Visibility = "Collapsed"
        $pnlServices.Visibility = "Collapsed"
        $pnlEventLogs.Visibility = "Collapsed"
        $pnlServerRoles.Visibility = "Collapsed"
        
        # Alle Navigationsbuttons zurücksetzen
        $btnSystemInfo.Style = $window.FindResource("TabButtonStyle")
        $btnHardware.Style = $window.FindResource("TabButtonStyle")
        $btnSoftware.Style = $window.FindResource("TabButtonStyle")
        $btnNetwork.Style = $window.FindResource("TabButtonStyle")
        $btnUsers.Style = $window.FindResource("TabButtonStyle")
        $btnServices.Style = $window.FindResource("TabButtonStyle")
        $btnEventLogs.Style = $window.FindResource("TabButtonStyle")
        $btnServerRoles.Style = $window.FindResource("TabButtonStyle")
        
        # Gewünschtes Panel anzeigen und Button markieren
        switch ($ViewName) {
            "SystemInfo" {
                $pnlSystemInfo.Visibility = "Visible"
                $btnSystemInfo.Style = $window.FindResource("TabActiveButtonStyle")
                Update-SystemInfo
            }
            "Hardware" {
                $pnlHardware.Visibility = "Visible"
                $btnHardware.Style = $window.FindResource("TabActiveButtonStyle")
                Update-HardwareInfo
            }
            "Software" {
                $pnlSoftware.Visibility = "Visible"
                $btnSoftware.Style = $window.FindResource("TabActiveButtonStyle")
                Update-SoftwareInfo
            }
            "Network" {
                $pnlNetwork.Visibility = "Visible"
                $btnNetwork.Style = $window.FindResource("TabActiveButtonStyle")
                Update-NetworkInfo
            }
            "Users" {
                $pnlUsers.Visibility = "Visible"
                $btnUsers.Style = $window.FindResource("TabActiveButtonStyle")
                Update-UsersInfo
            }
            "Services" {
                $pnlServices.Visibility = "Visible"
                $btnServices.Style = $window.FindResource("TabActiveButtonStyle")
                Update-ServicesInfo
            }
            "EventLogs" {
                $pnlEventLogs.Visibility = "Visible"
                $btnEventLogs.Style = $window.FindResource("TabActiveButtonStyle")
                Update-EventLogsInfo
            }
            "ServerRoles" {
                $pnlServerRoles.Visibility = "Visible"
                $btnServerRoles.Style = $window.FindResource("TabActiveButtonStyle")
                Update-ServerRolesInfo
            }
        }
        
        Write-Log "Ansicht gewechselt zu: $ViewName" -Level "INFO"
        $txtStatus.Text = "Bereit"
    }
    catch {
        Write-Log "Fehler beim Wechseln der Ansicht: $($_.Exception.Message)" -Level "ERROR"
        $txtStatus.Text = "Fehler beim Laden der Ansicht"
    }
}

# Funktion zum sicheren Aktualisieren von UI-Elementen
function Update-UIElement {
    param (
        [Parameter(Mandatory = $true)]
        [object]$Element,
        
        [Parameter(Mandatory = $true)]
        [string]$PropertyName,
        
        [Parameter(Mandatory = $true)]
        [object]$Value,
        
        [int]$MaxLength = 10000
    )
    
    try {
        if ($null -eq $Element) {
            Write-Log "UI-Element ist null, kann nicht aktualisiert werden" -Level "WARNING"
            return
        }
        
        if ($PropertyName -eq "Text" -and $Value -is [string]) {
            # Sonderzeichen filtern - nur druckbare ASCII-Zeichen zulassen
            $filteredValue = [System.Text.RegularExpressions.Regex]::Replace(
                $Value, 
                '[^\x20-\x7E\r\n]', 
                ' '
            )
            
            # Text auf maximale Länge beschränken
            if ($filteredValue.Length -gt $MaxLength) {
                $filteredValue = $filteredValue.Substring(0, $MaxLength) + "... (gekürzt)"
                Write-Log "Text für UI-Element gekürzt (Länge > $MaxLength Zeichen)" -Level "INFO"
            }
            
            $Element.$PropertyName = $filteredValue
        }
        elseif ($PropertyName -eq "ItemsSource") {
            $Element.$PropertyName = $Value
        }
        elseif ($PropertyName -eq "Value" -and $Element -is [System.Windows.Controls.ProgressBar]) {
            if ($Value -is [int] -or $Value -is [double]) {
                $Element.$PropertyName = [Math]::Min([Math]::Max($Value, 0), 100)
            }
            else {
                $Element.$PropertyName = 0
            }
        }
        elseif ($PropertyName -eq "IsChecked" -and $Element -is [System.Windows.Controls.CheckBox]) {
            $Element.$PropertyName = [bool]$Value
        }
        else {
            $Element.$PropertyName = $Value
        }
    }
    catch {
        Write-Log "Fehler beim Aktualisieren des UI-Elements ($PropertyName): $($_.Exception.Message)" -Level "ERROR"
        try {
            Add-Content -Path $script:fallbackLogFile -Value "UI-Update Fehler: $($_.Exception.Message)" -ErrorAction SilentlyContinue
        }
        catch {}
    }
}

# Sicherer Zugriff auf WMI/CIM
function Get-SafeWMIObject {
    param (
        [string]$Class,
        [string]$NameSpace = "root\cimv2",
        [string]$ComputerName = ".",
        [switch]$CIM
    )
    
    try {
        if ($CIM) {
            return Get-CimInstance -Namespace $NameSpace -ClassName $Class -ComputerName $ComputerName -ErrorAction Stop
        }
        else {
            return Get-WmiObject -Namespace $NameSpace -Class $Class -ComputerName $ComputerName -ErrorAction Stop
        }
    }
    catch {
        Write-Log "Fehler bei WMI/CIM-Abfrage für Klasse '$Class': $($_.Exception.Message)" -Level "ERROR"
        return $null
    }
}
#endregion

#region Datenerfassungsfunktionen
# System-Informationen
function Update-SystemInfo {
    try {
        $txtStatus.Text = "Sammle Systeminformationen..."
        
        # Betriebssystem-Informationen
        $os = Get-SafeWMIObject -Class Win32_OperatingSystem
        if ($os) {
            Update-UIElement -Element $txtOSName -PropertyName "Text" -Value $os.Caption
            Update-UIElement -Element $txtOSVersion -PropertyName "Text" -Value "$($os.Version) (Build $($os.BuildNumber))"
            Update-UIElement -Element $txtOSBuild -PropertyName "Text" -Value $os.BuildNumber
            Update-UIElement -Element $txtOSArchitecture -PropertyName "Text" -Value $os.OSArchitecture
            
            # Installationsdatum umwandeln
            try {
                $installDate = [Management.ManagementDateTimeConverter]::ToDateTime($os.InstallDate)
                Update-UIElement -Element $txtOSInstallDate -PropertyName "Text" -Value $installDate.ToString("dd.MM.yyyy HH:mm:ss")
            }
            catch {
                Update-UIElement -Element $txtOSInstallDate -PropertyName "Text" -Value "Unbekannt"
                Write-Log "Fehler bei Konvertierung des Installationsdatums: $($_.Exception.Message)" -Level "WARNING"
            }
            
            # Letzte Windows-Aktualisierung (über Registry)
            try {
                $lastUpdate = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\Results\Install" -Name "LastSuccessTime" -ErrorAction SilentlyContinue
                if ($lastUpdate) {
                    Update-UIElement -Element $txtOSLastUpdate -PropertyName "Text" -Value $lastUpdate.LastSuccessTime
                }
                else {
                    Update-UIElement -Element $txtOSLastUpdate -PropertyName "Text" -Value "Keine Informationen verfügbar"
                }
            }
            catch {
                Update-UIElement -Element $txtOSLastUpdate -PropertyName "Text" -Value "Keine Informationen verfügbar"
                Write-Log "Fehler bei Abfrage der letzten Windows-Updates: $($_.Exception.Message)" -Level "WARNING"
            }
        }
        
        # Computersystem-Informationen
        $cs = Get-SafeWMIObject -Class Win32_ComputerSystem
        if ($cs) {
            Update-UIElement -Element $txtComputerName -PropertyName "Text" -Value $cs.Name
            
            if ($cs.PartOfDomain) {
                Update-UIElement -Element $txtDomain -PropertyName "Text" -Value $cs.Domain
            }
            else {
                Update-UIElement -Element $txtDomain -PropertyName "Text" -Value "Arbeitsgruppe: $($cs.Workgroup)"
            }
            
            $systemTypeText = Get-SystemType
            Update-UIElement -Element $txtSystemType -PropertyName "Text" -Value $systemTypeText
        }
        
        # Uptime berechnen
        if ($os) {
            try {
                $lastBoot = [Management.ManagementDateTimeConverter]::ToDateTime($os.LastBootUpTime)
                $uptime = (Get-Date) - $lastBoot
                $uptimeText = "$($uptime.Days) Tage, $($uptime.Hours) Stunden, $($uptime.Minutes) Minuten"
                Update-UIElement -Element $txtUptime -PropertyName "Text" -Value $uptimeText
            }
            catch {
                Update-UIElement -Element $txtUptime -PropertyName "Text" -Value "Unbekannt"
                Write-Log "Fehler bei Berechnung der Uptime: $($_.Exception.Message)" -Level "WARNING"
            }
        }
        
        # Zeitzone
        try {
            $timeZone = Get-TimeZone
            Update-UIElement -Element $txtTimeZone -PropertyName "Text" -Value "$($timeZone.DisplayName) (UTC$($timeZone.BaseUtcOffset))"
        }
        catch {
            Update-UIElement -Element $txtTimeZone -PropertyName "Text" -Value "Unbekannt"
            Write-Log "Fehler bei Abfrage der Zeitzone: $($_.Exception.Message)" -Level "WARNING"
        }
        
        # Daten im globalen Report-Dictionary speichern
        $script:reportData.SystemInfo = @{
            OSName = $os.Caption
            OSVersion = "$($os.Version) (Build $($os.BuildNumber))"
            OSArchitecture = $os.OSArchitecture
            InstallDate = if ($installDate) { $installDate.ToString("dd.MM.yyyy HH:mm:ss") } else { "Unbekannt" }
            LastUpdate = $txtOSLastUpdate.Text
            ComputerName = $cs.Name
            Domain = $txtDomain.Text
            SystemType = $systemTypeText
            Uptime = $uptimeText
            TimeZone = $txtTimeZone.Text
        }
        
        $txtStatus.Text = "Systeminformationen geladen"
    }
    catch {
        Write-Log "Fehler beim Aktualisieren der Systeminformationen: $($_.Exception.Message)" -Level "ERROR"
        $txtStatus.Text = "Fehler beim Laden der Systeminformationen"
    }
}

# Hardware-Informationen
function Update-HardwareInfo {
    try {
        $txtStatus.Text = "Sammle Hardware-Informationen..."
        
        # Prozessor-Informationen
        $processors = @()
        $cpus = Get-SafeWMIObject -Class Win32_Processor
        
        if ($cpus) {
            if ($cpus -is [array]) {
                foreach ($cpu in $cpus) {
                    $processors += [PSCustomObject]@{
                        Name = $cpu.Name
                        Cores = $cpu.NumberOfCores
                        Threads = $cpu.NumberOfLogicalProcessors
                        Speed = "$($cpu.MaxClockSpeed) MHz"
                    }
                }
            }
            else {
                $processors += [PSCustomObject]@{
                    Name = $cpus.Name
                    Cores = $cpus.NumberOfCores
                    Threads = $cpus.NumberOfLogicalProcessors
                    Speed = "$($cpus.MaxClockSpeed) MHz"
                }
            }
        }
        
        Update-UIElement -Element $dgProcessors -PropertyName "ItemsSource" -Value $processors
        
        # RAM-Informationen
        $cs = Get-SafeWMIObject -Class Win32_ComputerSystem
        $os = Get-SafeWMIObject -Class Win32_OperatingSystem
        
        if ($cs -and $os) {
            # Gesamt-RAM
            $totalRAMGB = [math]::Round($cs.TotalPhysicalMemory / 1GB, 2)
            Update-UIElement -Element $txtTotalRAM -PropertyName "Text" -Value "$totalRAMGB GB"
            
            # Verfügbarer RAM
            $availableRAMGB = [math]::Round($os.FreePhysicalMemory / 1MB, 2)
            Update-UIElement -Element $txtAvailableRAM -PropertyName "Text" -Value "$availableRAMGB GB"
            
            # RAM-Nutzung in Prozent
            $usedRAMPerc = [math]::Round(100 - (($os.FreePhysicalMemory / 1MB) / ($cs.TotalPhysicalMemory / 1GB) * 100), 0)
            Update-UIElement -Element $txtRAMUsage -PropertyName "Text" -Value "$usedRAMPerc%"
            Update-UIElement -Element $pbRAMUsage -PropertyName "Value" -Value $usedRAMPerc
        }
        
        # RAM-Module Details
        $ramModules = @()
        $modules = Get-SafeWMIObject -Class Win32_PhysicalMemory
        
        if ($modules) {
            if ($modules -is [array]) {
                foreach ($module in $modules) {
                    $ramModules += [PSCustomObject]@{
                        Slot = $module.DeviceLocator
                        Capacity = "$([math]::Round($module.Capacity / 1GB, 0)) GB"
                        Type = switch ($module.MemoryType) {
                            24 { "DDR3" }
                            26 { "DDR4" }
                            27 { "DDR5" }
                            default { "Typ $($module.MemoryType)" }
                        }
                        Speed = "$($module.Speed) MHz"
                        Manufacturer = $module.Manufacturer
                    }
                }
            }
            else {
                $ramModules += [PSCustomObject]@{
                    Slot = $modules.DeviceLocator
                    Capacity = "$([math]::Round($modules.Capacity / 1GB, 0)) GB"
                    Type = switch ($modules.MemoryType) {
                        24 { "DDR3" }
                        26 { "DDR4" }
                        27 { "DDR5" }
                        default { "Typ $($modules.MemoryType)" }
                    }
                    Speed = "$($modules.Speed) MHz"
                    Manufacturer = $modules.Manufacturer
                }
            }
        }
        
        Update-UIElement -Element $dgRAMModules -PropertyName "ItemsSource" -Value $ramModules
        
        # Festplatten-Informationen
        $diskDrives = @()
        $logicalDisks = Get-SafeWMIObject -Class Win32_LogicalDisk -CIM | Where-Object { $_.DriveType -eq 3 }
        
        if ($logicalDisks) {
            if ($logicalDisks -is [array]) {
                foreach ($disk in $logicalDisks) {
                    $sizeGB = [math]::Round($disk.Size / 1GB, 2)
                    $freeSpaceGB = [math]::Round($disk.FreeSpace / 1GB, 2)
                    $percentFree = [math]::Round(($disk.FreeSpace / $disk.Size) * 100, 0)
                    
                    # Versuchen, das physische Laufwerk zu finden
                    try {
                        $physicalDisk = Get-Disk -Number ($disk.DeviceId -replace '[^\d]', '') -ErrorAction SilentlyContinue
                        $diskModel = if ($physicalDisk) { $physicalDisk.FriendlyName } else { "Unbekannt" }
                        $diskType = if ($physicalDisk) { 
                            switch ($physicalDisk.MediaType) {
                                "SSD" { "SSD" }
                                "HDD" { "HDD" }
                                "SCM" { "Storage Class Memory" }
                                "Unspecified" { "Unspezifiziert" }
                                default { $physicalDisk.MediaType }
                            }
                        } else { "Unbekannt" }
                    }
                    catch {
                        $diskModel = "Unbekannt"
                        $diskType = "Unbekannt"
                        Write-Log "Fehler beim Abrufen von physischen Festplatteninformationen: $($_.Exception.Message)" -Level "WARNING"
                    }
                    
                    $diskDrives += [PSCustomObject]@{
                        Drive = $disk.DeviceID
                        Model = $diskModel
                        Type = $diskType
                        Size = "$sizeGB GB"
                        FreeSpace = "$freeSpaceGB GB"
                        PercentFree = "$percentFree%"
                    }
                }
            }
            else {
                $sizeGB = [math]::Round($logicalDisks.Size / 1GB, 2)
                $freeSpaceGB = [math]::Round($logicalDisks.FreeSpace / 1GB, 2)
                $percentFree = [math]::Round(($logicalDisks.FreeSpace / $logicalDisks.Size) * 100, 0)
                
                # Versuchen, das physische Laufwerk zu finden
                try {
                    $physicalDisk = Get-Disk -Number ($logicalDisks.DeviceId -replace '[^\d]', '') -ErrorAction SilentlyContinue
                    $diskModel = if ($physicalDisk) { $physicalDisk.FriendlyName } else { "Unbekannt" }
                    $diskType = if ($physicalDisk) { 
                        switch ($physicalDisk.MediaType) {
                            "SSD" { "SSD" }
                            "HDD" { "HDD" }
                            "SCM" { "Storage Class Memory" }
                            "Unspecified" { "Unspezifiziert" }
                            default { $physicalDisk.MediaType }
                        }
                    } else { "Unbekannt" }
                }
                catch {
                    $diskModel = "Unbekannt"
                    $diskType = "Unbekannt"
                    Write-Log "Fehler beim Abrufen von physischen Festplatteninformationen: $($_.Exception.Message)" -Level "WARNING"
                }
                
                $diskDrives += [PSCustomObject]@{
                    Drive = $logicalDisks.DeviceID
                    Model = $diskModel
                    Type = $diskType
                    Size = "$sizeGB GB"
                    FreeSpace = "$freeSpaceGB GB"
                    PercentFree = "$percentFree%"
                }
            }
        }
        
        Update-UIElement -Element $dgDiskDrives -PropertyName "ItemsSource" -Value $diskDrives
        
        # Hardware-Daten im globalen Report-Dictionary speichern
        $script:reportData.Hardware = @{
            Processors = $processors
            TotalRAM = "$totalRAMGB GB"
            AvailableRAM = "$availableRAMGB GB"
            RAMUsagePercent = $usedRAMPerc
            RAMModules = $ramModules
            DiskDrives = $diskDrives
        }
        
        $txtStatus.Text = "Hardware-Informationen geladen"
    }
    catch {
        Write-Log "Fehler beim Aktualisieren der Hardware-Informationen: $($_.Exception.Message)" -Level "ERROR"
        $txtStatus.Text = "Fehler beim Laden der Hardware-Informationen"
    }
}
#endregion