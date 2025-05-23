<#
.SYNOPSIS
    Windows Server 2012 Migration Tool
.DESCRIPTION
    Dieses Tool automatisiert die Migration von Windows Server 2012 Domänencontrollern,
    DNS- und DHCP-Diensten auf neue Windows Server.
.NOTES
    Version: 0.0.1
    Autor: System-Administrator
    Erstellungsdatum: $(Get-Date -Format "dd.MM.yyyy")
    Voraussetzungen: PowerShell 5.1, Administratorrechte, Active Directory Module
#>

# Hilfsvariable für Doppelpunkt definieren, um Syntaxfehler zu vermeiden
$colon = ":"

#region Logging-Funktionen - muss zuerst definiert werden
# Globale Logdatei-Variable
$script:logFile = "$PSScriptRoot\migration_log.txt"
$script:fallbackLogFile = "$env:TEMP\migration_fallback_log.txt"

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
    
    # AD-Module prüfen und importieren
    if (Get-Module -ListAvailable -Name ActiveDirectory) {
        Import-Module ActiveDirectory -ErrorAction Stop
        Write-Log "Active Directory Module erfolgreich geladen" -Level "INFO"
    } else {
        Write-Log "Active Directory Module nicht verfügbar - einige Funktionen werden nicht verfügbar sein" -Level "WARNING"
    }
}
catch {
    Write-Log "Fehler beim Laden der Module: $($_.Exception.Message)" -Level "ERROR"
    [System.Windows.Forms.MessageBox]::Show("Ein Fehler ist beim Laden der Module aufgetreten. Möglicherweise sind einige Funktionen nicht verfügbar.", "Fehler", "OK", "Error")
}

# Globale Variablen und Konfiguration
$script:appName = "Easy Windows Server 2012 Migration Tool"
$script:themeColor = "#007ACC"
$script:logFile = "$PSScriptRoot\migration_log.txt"
$script:footerText = "© " + (Get-Date).Year + " Windows Server Management" 
$script:footerWebsite = "www.easyit.com"
#endregion

#region XAML GUI Definition
# XAML aus Datei laden
$xamlPath = Join-Path $PSScriptRoot "GUI\MainWindow.xaml"
if (Test-Path $xamlPath) {
    [xml]$xaml = Get-Content -Path $xamlPath -Raw
    # Remove x:Class attribute if it exists
    if ($xaml.Window.Class) {
        $xaml.Window.RemoveAttribute("Class")
    }
    if ($xaml.Window.GetAttribute("x:Class")) {
        $xaml.Window.RemoveAttribute("x:Class")
    }
    Write-Log "XAML-Datei erfolgreich geladen: $xamlPath" -Level "INFO"
} else {
    # Fallback zum eingebetteten XAML, wenn die externe Datei nicht gefunden wird
    Write-Log "XAML-Datei nicht gefunden: $xamlPath, verwende eingebettetes XAML" -Level "WARNING"
    [xml]$xaml = @'
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Windows Server 2012 Migration Tool" 
    Height="800" 
    Width="1200" 
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
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" CornerRadius="3">
                            <ContentPresenter HorizontalAlignment="Left" VerticalAlignment="Center" Margin="10,0,0,0"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
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
            <Setter Property="Height" Value="25"/>
            <Setter Property="Width" Value="150"/>
            <Setter Property="Background" Value="#007ACC"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" CornerRadius="3">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#005A9E"/>
                </Trigger>
                <Trigger Property="IsPressed" Value="True">
                    <Setter Property="Background" Value="#003E6B"/>
                </Trigger>
            </Style.Triggers>
        </Style>
    </Window.Resources>
    
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="80"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="60"/>
        </Grid.RowDefinitions>
        
        <!-- Header -->
        <Border Grid.Row="0" Background="#007ACC">
            <StackPanel Orientation="Horizontal" VerticalAlignment="Center" Margin="20,0,0,0">
                <TextBlock Text="Windows Server 2012 Migration Tool" FontSize="24" Foreground="White" VerticalAlignment="Center"/>
            </StackPanel>
        </Border>
        
        <Grid Grid.Row="1">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="220"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            
            <!-- Navigation Panel -->
            <Border Grid.Column="0" Background="#F0F0F0">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <!-- Navigation Buttons -->
                    <StackPanel Grid.Row="0">
                        <TextBlock Text="Navigation" FontSize="16" Margin="10,15,0,15" FontWeight="Bold"/>
                        <Button x:Name="btnHome" Content="Übersicht" Style="{StaticResource NavigationButton}"/>
                        <Button x:Name="btnDiscovery" Content="1. Ermittlung" Style="{StaticResource NavigationButton}"/>
                        <Button x:Name="btnInstallation" Content="2. Installation" Style="{StaticResource NavigationButton}"/>
                        <Button x:Name="btnMigration" Content="3. Migration" Style="{StaticResource NavigationButton}"/>
                        <Button x:Name="btnValidation" Content="4. Validierung" Style="{StaticResource NavigationButton}"/>
                        <Button x:Name="btnDecommission" Content="5. Abschaltung" Style="{StaticResource NavigationButton}"/>
                    </StackPanel>
                    
                    <!-- Bottom Navigation Icons -->
                    <StackPanel Grid.Row="1" Orientation="Horizontal" HorizontalAlignment="Center" Margin="0,0,0,10">
                        <Button x:Name="btnInfo" Width="32" Height="32" Margin="5" ToolTip="Information">
                            <Image Source="/assets/info.png"/>
                        </Button>
                        <Button x:Name="btnSettings" Width="32" Height="32" Margin="5" ToolTip="Einstellungen">
                            <Image Source="/assets/settings.png"/>
                        </Button>
                        <Button x:Name="btnClose" Width="32" Height="32" Margin="5" ToolTip="Beenden">
                            <Image Source="/assets/close.png"/>
                        </Button>
                    </StackPanel>
                </Grid>
            </Border>
            
            <!-- Content Area -->
            <Grid Grid.Column="1" Margin="10">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>
                
                <!-- Tab Header -->
                <TextBlock x:Name="txtCurrentTab" Grid.Row="0" Text="Übersicht" FontSize="20" FontWeight="Bold" Margin="0,0,0,10"/>
                
                <!-- Content Frame -->
                <Border Grid.Row="1" BorderBrush="#CCCCCC" BorderThickness="1" Padding="10">
                    <Grid>
                        <!-- Content Panels -->
                        <Grid x:Name="panelHome" Visibility="Visible">
                            <StackPanel>
                                <TextBlock Text="Willkommen beim Windows Server 2012 Migration Tool" FontSize="18" Margin="0,0,0,20"/>
                                <TextBlock TextWrapping="Wrap">
                                    Dieses Tool führt Sie durch den Prozess der Migration von Windows Server 2012 Domänencontrollern, 
                                    DNS- und DHCP-Diensten auf neue Windows Server. Folgen Sie den Schritten in der Navigation links, 
                                    um durch den Migrationsprozess geführt zu werden.
                                </TextBlock>
                                <TextBlock Margin="0,20,0,10" FontWeight="Bold">Migrationsphasen:</TextBlock>
                                <TextBlock TextWrapping="Wrap" Margin="10,0,0,5">1. Ermittlung - Analyse der bestehenden Umgebung</TextBlock>
                                <TextBlock TextWrapping="Wrap" Margin="10,0,0,5">2. Installation - Einrichtung der neuen Server</TextBlock>
                                <TextBlock TextWrapping="Wrap" Margin="10,0,0,5">3. Migration - Übertragung der Rollen und Dienste</TextBlock>
                                <TextBlock TextWrapping="Wrap" Margin="10,0,0,5">4. Validierung - Überprüfung der erfolgreichen Migration</TextBlock>
                                <TextBlock TextWrapping="Wrap" Margin="10,0,0,5">5. Abschaltung - Sichere Außerbetriebnahme der alten Server</TextBlock>
                            </StackPanel>
                        </Grid>
                        
                        <Grid x:Name="panelDiscovery" Visibility="Collapsed">
                            <Grid>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="*"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>
                                
                                <StackPanel Grid.Row="0" Margin="0,0,0,10">
                                    <TextBlock Text="Ermittlung der aktuellen Umgebung" FontSize="16" FontWeight="Bold" Margin="0,0,0,10"/>
                                    <TextBlock TextWrapping="Wrap" Margin="0,0,0,10">
                                        In diesem Schritt wird die bestehende Active Directory-Umgebung analysiert, um Domänencontroller, FSMO-Rollen, Serverdienste und installierte Rollen zu ermitteln.
                                    </TextBlock>
                                    <Button x:Name="btnStartDiscovery" Content="Umgebung analysieren" Style="{StaticResource StandardButton}" HorizontalAlignment="Left" Margin="0,10,0,10"/>
                                </StackPanel>
                                
                                <TabControl Grid.Row="1" x:Name="tabDiscovery" Visibility="Collapsed">
                                    <TabItem Header="Domänencontroller">
                                        <Grid Margin="5">
                                            <DataGrid x:Name="dgDomainControllers" AutoGenerateColumns="False" IsReadOnly="True">
                                                <DataGrid.Columns>
                                                    <DataGridTextColumn Header="Name" Binding="{Binding Name}" Width="150"/>
                                                    <DataGridTextColumn Header="Standort" Binding="{Binding Site}" Width="100"/>
                                                    <DataGridTextColumn Header="IP-Adresse" Binding="{Binding IPAddress}" Width="120"/>
                                                    <DataGridTextColumn Header="Betriebssystem" Binding="{Binding OperatingSystem}" Width="180"/>
                                                    <DataGridTextColumn Header="Version" Binding="{Binding OSVersion}" Width="80"/>
                                                    <DataGridCheckBoxColumn Header="Global Catalog" Binding="{Binding IsGlobalCatalog}" Width="100"/>
                                                    <DataGridCheckBoxColumn Header="Erreichbar" Binding="{Binding IsOnline}" Width="80"/>
                                                </DataGrid.Columns>
                                            </DataGrid>
                                        </Grid>
                                    </TabItem>
                                    <TabItem Header="FSMO-Rollen">
                                        <Grid Margin="5">
                                            <ListView x:Name="lvFSMORoles">
                                                <ListView.View>
                                                    <GridView>
                                                        <GridViewColumn Header="Rollenname" Width="250" DisplayMemberBinding="{Binding RoleName}"/>
                                                        <GridViewColumn Header="Inhaber" Width="200" DisplayMemberBinding="{Binding RoleOwner}"/>
                                                    </GridView>
                                                </ListView.View>
                                            </ListView>
                                        </Grid>
                                    </TabItem>
                                    <TabItem Header="DNS/DHCP-Server">
                                        <Grid Margin="5">
                                            <DataGrid x:Name="dgServices" AutoGenerateColumns="False" IsReadOnly="True">
                                                <DataGrid.Columns>
                                                    <DataGridTextColumn Header="Servername" Binding="{Binding ServerName}" Width="150"/>
                                                    <DataGridCheckBoxColumn Header="DNS-Dienst" Binding="{Binding HasDNS}" Width="100"/>
                                                    <DataGridCheckBoxColumn Header="DHCP-Dienst" Binding="{Binding HasDHCP}" Width="100"/>
                                                    <DataGridTextColumn Header="DNS-Status" Binding="{Binding DNSStatus}" Width="100"/>
                                                    <DataGridTextColumn Header="DHCP-Status" Binding="{Binding DHCPStatus}" Width="100"/>
                                                    <DataGridTextColumn Header="Betriebssystem" Binding="{Binding OperatingSystem}" Width="180"/>
                                                </DataGrid.Columns>
                                            </DataGrid>
                                        </Grid>
                                    </TabItem>
                                    <TabItem Header="Installierte Rollen">
                                        <Grid Margin="5">
                                            <DataGrid x:Name="dgServerRoles" AutoGenerateColumns="False" IsReadOnly="True">
                                                <DataGrid.Columns>
                                                    <DataGridTextColumn Header="Servername" Binding="{Binding ServerName}" Width="150"/>
                                                    <DataGridTextColumn Header="Rollenname" Binding="{Binding RoleName}" Width="200"/>
                                                    <DataGridTextColumn Header="Status" Binding="{Binding InstallState}" Width="100"/>
                                                    <DataGridTextColumn Header="Features" Binding="{Binding Features}" Width="280"/>
                                                </DataGrid.Columns>
                                            </DataGrid>
                                        </Grid>
                                    </TabItem>
                                    <TabItem Header="DNS-Konfiguration">
                                        <Grid Margin="5">
                                            <DataGrid x:Name="dgDnsConfig" AutoGenerateColumns="False" IsReadOnly="True">
                                                <DataGrid.Columns>
                                                    <DataGridTextColumn Header="DNS-Server" Binding="{Binding ServerName}" Width="150"/>
                                                    <DataGridTextColumn Header="Zonen" Binding="{Binding Zones}" Width="200"/>
                                                    <DataGridTextColumn Header="Forwarders" Binding="{Binding Forwarders}" Width="150"/>
                                                    <DataGridTextColumn Header="Zonen-Replikation" Binding="{Binding ZoneReplication}" Width="150"/>
                                                    <DataGridCheckBoxColumn Header="Rekursion aktiviert" Binding="{Binding RecursionEnabled}" Width="120"/>
                                                </DataGrid.Columns>
                                            </DataGrid>
                                        </Grid>
                                    </TabItem>
                                    <TabItem Header="DHCP-Konfiguration">
                                        <Grid Margin="5">
                                            <DataGrid x:Name="dgDhcpConfig" AutoGenerateColumns="False" IsReadOnly="True">
                                                <DataGrid.Columns>
                                                    <DataGridTextColumn Header="DHCP-Server" Binding="{Binding ServerName}" Width="150"/>
                                                    <DataGridTextColumn Header="Bereiche" Binding="{Binding Scopes}" Width="280"/>
                                                    <DataGridTextColumn Header="Optionen" Binding="{Binding ServerOptions}" Width="200"/>
                                                    <DataGridCheckBoxColumn Header="Failover-Konfiguration" Binding="{Binding HasFailover}" Width="140"/>
                                                </DataGrid.Columns>
                                            </DataGrid>
                                        </Grid>
                                    </TabItem>
                                    <TabItem Header="AD-Struktur">
                                        <Grid Margin="5">
                                            <ListView x:Name="lvADStructure">
                                                <ListView.View>
                                                    <GridView>
                                                        <GridViewColumn Header="Typ" Width="150" DisplayMemberBinding="{Binding Type}"/>
                                                        <GridViewColumn Header="Name" Width="200" DisplayMemberBinding="{Binding Name}"/>
                                                        <GridViewColumn Header="Beschreibung" Width="300" DisplayMemberBinding="{Binding Description}"/>
                                                    </GridView>
                                                </ListView.View>
                                            </ListView>
                                        </Grid>
                                    </TabItem>
                                    <TabItem Header="Zusammenfassung">
                                        <ScrollViewer VerticalScrollBarVisibility="Auto">
                                            <TextBox x:Name="txtSummary" IsReadOnly="True" TextWrapping="Wrap" AcceptsReturn="True" 
                                                    VerticalAlignment="Stretch" HorizontalAlignment="Stretch" Margin="5"/>
                                        </ScrollViewer>
                                    </TabItem>
                                </TabControl>
                                
                                <!-- Export Button Panel -->
                                <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,10,0,0" HorizontalAlignment="Right">
                                    <Button x:Name="btnExportHTML" Content="Als HTML exportieren" Style="{StaticResource StandardButton}" Margin="0,0,5,0" Visibility="Collapsed"/>
                                    <Button x:Name="btnPrintReport" Content="Bericht drucken" Style="{StaticResource StandardButton}" Margin="5,0,0,0" Visibility="Collapsed"/>
                                </StackPanel>
                            </Grid>
                        </Grid>
                        
                        <Grid x:Name="panelInstallation" Visibility="Collapsed">
                            <Grid>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="*"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>
                                
                                <!-- Überschrift und Beschreibung -->
                                <StackPanel Grid.Row="0" Margin="0,0,0,10">
                                    <TextBlock Text="Serverrollen Installation" FontSize="18" FontWeight="Bold" Margin="0,0,0,10"/>
                                    <TextBlock TextWrapping="Wrap">
                                        In diesem Bereich können Sie Serverrollen auf lokalen oder entfernten Servern installieren.
                                        Wählen Sie den Zielserver und die zu installierenden Rollen und Features aus.
                                    </TextBlock>
                                    <Border Background="#f0f7ff" BorderBrush="#c3daf9" BorderThickness="1" Margin="0,10,0,5" Padding="10">
                                        <TextBlock TextWrapping="Wrap">
                                            Für die Installation von Rollen auf entfernten Servern werden entsprechende Berechtigungen benötigt.
                                            Die Installation erfolgt über PowerShell Remoting und das Windows-Feature-Modul.
                                        </TextBlock>
                                    </Border>
                                </StackPanel>
                                
                                <!-- Server- und Rollenauswahl -->
                                <Grid Grid.Row="1" Margin="0,10,0,0">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="*"/>
                                    </Grid.ColumnDefinitions>
                                    
                                    <!-- Linke Seite: Serverauswahl -->
                                    <GroupBox Grid.Column="0" Header="Zielserver" Margin="0,0,5,0">
                                        <Grid Margin="5">
                                            <Grid.RowDefinitions>
                                                <RowDefinition Height="Auto"/>
                                                <RowDefinition Height="Auto"/>
                                                <RowDefinition Height="Auto"/>
                                                <RowDefinition Height="*"/>
                                                <RowDefinition Height="Auto"/>
                                            </Grid.RowDefinitions>
                                            
                                            <RadioButton x:Name="rbLocalServer" Grid.Row="0" Content="Lokaler Server" IsChecked="True" Margin="0,5,0,5"/>
                                            <RadioButton x:Name="rbRemoteServer" Grid.Row="1" Content="Entfernter Server" Margin="0,5,0,5"/>
                                            
                                            <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,5,0,5">
                                                <TextBlock Text="Servername:" VerticalAlignment="Center" Width="80"/>
                                                <TextBox x:Name="txtServerName" Width="200" IsEnabled="{Binding IsChecked, ElementName=rbRemoteServer}"/>
                                            </StackPanel>
                                            
                                            <GroupBox Grid.Row="3" Header="Domänenserver" Margin="0,5,0,5">
                                                <ListView x:Name="lvDomainServers" Height="150">
                                                    <ListView.View>
                                                        <GridView>
                                                            <GridViewColumn Header="Servername" Width="150" DisplayMemberBinding="{Binding Name}"/>
                                                            <GridViewColumn Header="Betriebssystem" Width="200" DisplayMemberBinding="{Binding OS}"/>
                                                        </GridView>
                                                    </ListView.View>
                                                </ListView>
                                            </GroupBox>
                                            
                                            <Button x:Name="btnRefreshServers" Grid.Row="4" Content="Server aktualisieren" 
                                                    Style="{StaticResource StandardButton}" HorizontalAlignment="Left" Margin="0,5,0,0"/>
                                        </Grid>
                                    </GroupBox>
                                    
                                    <!-- Rechte Seite: Rollenauswahl -->
                                    <GroupBox Grid.Column="1" Header="Serverrollen und Features" Margin="5,0,0,0">
                                        <Grid Margin="5">
                                            <Grid.RowDefinitions>
                                                <RowDefinition Height="Auto"/>
                                                <RowDefinition Height="*"/>
                                                <RowDefinition Height="Auto"/>
                                            </Grid.RowDefinitions>
                                            
                                            <TextBlock Grid.Row="0" Text="Wählen Sie die zu installierenden Rollen und Features:" Margin="0,0,0,5"/>
                                            
                                            <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto">
                                                <StackPanel>
                                                    <!-- Domänendienste -->
                                                    <GroupBox Header="Active Directory" Margin="0,5,0,5">
                                                        <StackPanel>
                                                            <CheckBox x:Name="chkADDS" Content="Active Directory-Domänendienste" Margin="0,5,0,0"/>
                                                            <CheckBox x:Name="chkADCS" Content="Active Directory-Zertifikatsdienste" Margin="0,5,0,0"/>
                                                            <CheckBox x:Name="chkADLDS" Content="Active Directory-Lightweight Directory Services" Margin="0,5,0,0"/>
                                                            <CheckBox x:Name="chkADRMS" Content="Active Directory-Rechteverwaltungsdienste" Margin="0,5,0,0"/>
                                                        </StackPanel>
                                                    </GroupBox>
                                                    
                                                    <!-- Netzwerkdienste -->
                                                    <GroupBox Header="Netzwerkdienste" Margin="0,5,0,5">
                                                        <StackPanel>
                                                            <CheckBox x:Name="chkDNS" Content="DNS-Server" Margin="0,5,0,0"/>
                                                            <CheckBox x:Name="chkDHCP" Content="DHCP-Server" Margin="0,5,0,0"/>
                                                            <CheckBox x:Name="chkWINS" Content="WINS-Server" Margin="0,5,0,0"/>
                                                        </StackPanel>
                                                    </GroupBox>
                                                    
                                                    <!-- Dateidieste -->
                                                    <GroupBox Header="Dateidienste" Margin="0,5,0,5">
                                                        <StackPanel>
                                                            <CheckBox x:Name="chkFileServer" Content="Dateiserver" Margin="0,5,0,0"/>
                                                            <CheckBox x:Name="chkDFS" Content="DFS-Namespaces" Margin="0,5,0,0"/>
                                                            <CheckBox x:Name="chkDFSR" Content="DFS-Replikation" Margin="0,5,0,0"/>
                                                        </StackPanel>
                                                    </GroupBox>
                                                    
                                                    <!-- Anwendungsdienste -->
                                                    <GroupBox Header="Anwendungsdienste" Margin="0,5,0,5">
                                                        <StackPanel>
                                                            <CheckBox x:Name="chkWebServer" Content="Webserver (IIS)" Margin="0,5,0,0"/>
                                                            <CheckBox x:Name="chkWSUS" Content="Windows Server Update Services" Margin="0,5,0,0"/>
                                                        </StackPanel>
                                                    </GroupBox>
                                                </StackPanel>
                                            </ScrollViewer>
                                            
                                            <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0">
                                                <Button x:Name="btnCheckAll" Content="Alle auswählen" Width="120" Margin="0,0,10,0"/>
                                                <Button x:Name="btnUncheckAll" Content="Alle abwählen" Width="120"/>
                                            </StackPanel>
                                        </Grid>
                                    </GroupBox>
                                </Grid>
                                
                                <!-- Statusbereich und Aktionsbuttons -->
                                <Grid Grid.Row="2" Margin="0,15,0,0">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
                                    
                                    <!-- Statusanzeige -->
                                    <Border Grid.Column="0" BorderBrush="#CCCCCC" BorderThickness="1" Background="#F9F9F9" MinHeight="60" Padding="10">
                                        <TextBlock x:Name="txtInstallationStatus" Text="Bereit für die Installation von Serverrollen." 
                                                  TextWrapping="Wrap" VerticalAlignment="Center"/>
                                    </Border>
                                    
                                    <!-- Aktionsbuttons -->
                                    <StackPanel Grid.Column="1" Orientation="Vertical" Margin="10,0,0,0">
                                        <Button x:Name="btnCheckRequirements" Content="Voraussetzungen prüfen" Style="{StaticResource StandardButton}" Margin="0,0,0,10"/>
                                        <Button x:Name="btnInstallRoles" Content="Rollen installieren" Style="{StaticResource StandardButton}"/>
                                    </StackPanel>
                                </Grid>
                            </Grid>
                        </Grid>
                        
                        <Grid x:Name="panelMigration" Visibility="Collapsed">
                            <TextBlock Text="Migration - Übertragung der Rollen und Dienste" FontSize="16"/>
                        </Grid>
                        
                        <Grid x:Name="panelValidation" Visibility="Collapsed">
                            <Grid>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="*"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>
                                
                                <!-- Überschrift und Beschreibung -->
                                <StackPanel Grid.Row="0" Margin="0,0,0,10">
                                    <TextBlock Text="Servervalidierung" FontSize="18" FontWeight="Bold" Margin="0,0,0,10"/>
                                    <TextBlock TextWrapping="Wrap">
                                        Diese Validierung überprüft den aktuellen Server, auf dem dieses Tool ausgeführt wird. Es werden verschiedene
                                        Aspekte wie Active Directory, DNS, DHCP und allgemeine Serverstabilität untersucht.
                                    </TextBlock>
                                    <Border Background="#f0f7ff" BorderBrush="#c3daf9" BorderThickness="1" Margin="0,10,0,5" Padding="10">
                                        <StackPanel>
                                            <TextBlock TextWrapping="Wrap" FontWeight="SemiBold">
                                                <Run>Validierung wird durchgeführt für Server: </Run>
                                                <Run x:Name="txtLocalServerName" FontWeight="Bold"/>
                                            </TextBlock>
                                            <TextBlock TextWrapping="Wrap" Margin="0,5,0,0">
                                                Wählen Sie die zu überprüfenden Bereiche und starten Sie die Validierung.
                                            </TextBlock>
                                        </StackPanel>
                                    </Border>
                                </StackPanel>
                                
                                <!-- Prüfoptionen -->
                                <Grid Grid.Row="1" Margin="0,0,0,15">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="*"/>
                                    </Grid.ColumnDefinitions>
                                    <Grid.RowDefinitions>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                    </Grid.RowDefinitions>
                                    
                                    <TextBlock Grid.Row="0" Grid.ColumnSpan="4" Text="Zu überprüfende Bereiche auswählen:" FontWeight="SemiBold" Margin="0,0,0,5"/>
                                    
                                    <CheckBox x:Name="chkValidateBasics" Grid.Row="1" Grid.Column="0" Content="Server-Grundfunktionen" IsChecked="True" Margin="0,5,0,5"/>
                                    <CheckBox x:Name="chkValidateAD" Grid.Row="1" Grid.Column="1" Content="Active Directory" IsChecked="True" Margin="0,5,0,5"/>
                                    <CheckBox x:Name="chkValidateDNS" Grid.Row="1" Grid.Column="2" Content="DNS-Dienst" IsChecked="True" Margin="0,5,0,5"/>
                                    <CheckBox x:Name="chkValidateDHCP" Grid.Row="1" Grid.Column="3" Content="DHCP-Dienst" IsChecked="True" Margin="0,5,0,5"/>
                                    
                                    <CheckBox x:Name="chkValidateRoles" Grid.Row="2" Grid.Column="0" Content="FSMO-Rollen" IsChecked="True" Margin="0,5,0,5"/>
                                    <CheckBox x:Name="chkValidateGPO" Grid.Row="2" Grid.Column="1" Content="Gruppenrichtlinien" IsChecked="True" Margin="0,5,0,5"/>
                                    <CheckBox x:Name="chkValidatePerf" Grid.Row="2" Grid.Column="2" Content="Performance" IsChecked="True" Margin="0,5,0,5"/>
                                    <CheckBox x:Name="chkValidateServices" Grid.Row="2" Grid.Column="3" Content="Kritische Dienste" IsChecked="True" Margin="0,5,0,5"/>
                                    
                                    <Button x:Name="btnStartValidation" Grid.Row="3" Grid.Column="3" Content="Validierung starten" 
                                            Style="{StaticResource StandardButton}" HorizontalAlignment="Right" Margin="0,10,0,0" Height="30"/>
                                </Grid>
                                
                                <!-- Validierungsstatus -->
                                <TextBlock x:Name="txtValidationStatus" Grid.Row="1" Text="" FontWeight="Bold" Margin="0,5,0,5" Visibility="Collapsed"/>
                                
                                <!-- Validierungsergebnisse -->
                                <TabControl Grid.Row="2" x:Name="tabValidation" Visibility="Collapsed">
                                    <TabItem Header="Zusammenfassung">
                                        <ScrollViewer VerticalScrollBarVisibility="Auto">
                                            <StackPanel>
                                                <DataGrid x:Name="dgValidationSummary" AutoGenerateColumns="False" IsReadOnly="True" 
                                                          HeadersVisibility="Column" GridLinesVisibility="Horizontal" Margin="0,10,0,0">
                                                    <DataGrid.Columns>
                                                        <DataGridTextColumn Header="Kategorie" Binding="{Binding Category}" Width="150"/>
                                                        <DataGridTextColumn Header="Komponente" Binding="{Binding Component}" Width="200"/>
                                                        <DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="100"/>
                                                        <DataGridTextColumn Header="Details" Binding="{Binding Details}" Width="*"/>
                                                    </DataGrid.Columns>
                                                </DataGrid>
                                            </StackPanel>
                                        </ScrollViewer>
                                    </TabItem>
                                    
                                    <!-- Weitere Tab-Inhalte für Detailergebnisse... -->
                                    
                                    <TabItem Header="Bericht">
                                        <TextBox x:Name="txtValidationReport" IsReadOnly="True" TextWrapping="Wrap" 
                                                 VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" 
                                                 FontFamily="Consolas" FontSize="12"/>
                                    </TabItem>
                                </TabControl>
                                
                                <!-- Export Button Panel -->
                                <StackPanel Grid.Row="3" Orientation="Horizontal" Margin="0,10,0,0" HorizontalAlignment="Right">
                                    <Button x:Name="btnSaveValidationHTML" Content="Als HTML speichern" Style="{StaticResource StandardButton}" Margin="0,0,10,0" Visibility="Collapsed"/>
                                    <Button x:Name="btnPrintValidation" Content="Drucken" Style="{StaticResource StandardButton}" Visibility="Collapsed"/>
                                </StackPanel>
                            </Grid>
                        </Grid>
                        
                        <Grid x:Name="panelDecommission" Visibility="Collapsed">
                            <TextBlock Text="Abschaltung - Sichere Außerbetriebnahme der alten Server" FontSize="16"/>
                        </Grid>
                    </Grid>
                </Border>
            </Grid>
        </Grid>
        
        <!-- Footer -->
        <Border Grid.Row="2" Background="#F0F0F0" BorderBrush="#CCCCCC" BorderThickness="0,1,0,0">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                
                <TextBlock Grid.Column="0" Text="© 08.04.2025 Windows Server 2012 Migration Tool" Margin="15,0,0,0" VerticalAlignment="Center"/>
                <TextBlock Grid.Column="1" Margin="0,0,15,0" VerticalAlignment="Center">
                    <Hyperlink x:Name="linkWebsite" NavigateUri="https://support.microsoft.com/">Support-Website</Hyperlink>
                </TextBlock>
            </Grid>
        </Border>
    </Grid>
</Window>
'@
}

# Platzhalter im XAML ersetzen (korrigierte Muster)
$xaml.Window.Title = $xaml.Window.Title -replace "{AppName}", $script:appName

# ThemeBrush und DarkModeBrush ersetzen
$resources = $xaml.SelectNodes("//Window/Window.Resources/*")
foreach ($resource in $resources) {
    if ($resource.Name -eq "ThemeBrush" -or $resource.Name -eq "DarkModeBrush") {
        $resource.Color = $resource.Color -replace "{ThemeColor}", $script:themeColor
    }
}

# Footer-Texte ersetzen
$textBlocks = $xaml.SelectNodes("//TextBlock")
foreach ($textBlock in $textBlocks) {
    if ($textBlock.Text -eq "{FooterText}") {
        $textBlock.Text = $script:footerText
    }
    elseif ($textBlock.Text -eq "{FooterWebsite}") {
        $textBlock.Text = $script:footerWebsite
    }
}

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
# Navigationselemente
$btnHome = $window.FindName("btnHome")
$btnDiscovery = $window.FindName("btnDiscovery")
$btnInstallation = $window.FindName("btnInstallation")
$btnMigration = $window.FindName("btnMigration")
$btnValidation = $window.FindName("btnValidation")
$btnDecommission = $window.FindName("btnDecommission")

# Steuerungselemente
$btnInfo = $window.FindName("btnInfo")
$btnSettings = $window.FindName("btnSettings")
$btnClose = $window.FindName("btnClose")

# Inhaltsbereich
$txtCurrentTab = $window.FindName("txtCurrentTab")
$panelHome = $window.FindName("panelHome")
$panelDiscovery = $window.FindName("panelDiscovery")
$panelInstallation = $window.FindName("panelInstallation")
$panelMigration = $window.FindName("panelMigration")
$panelValidation = $window.FindName("panelValidation")
$panelDecommission = $window.FindName("panelDecommission")

# Discovery-Bereich
$btnStartDiscovery = $window.FindName("btnStartDiscovery")
$tabDiscovery = $window.FindName("tabDiscovery")
$dgDomainControllers = $window.FindName("dgDomainControllers")
$lvFSMORoles = $window.FindName("lvFSMORoles")
$dgServices = $window.FindName("dgServices")
$txtSummary = $window.FindName("txtSummary")

# Neue erweiterte Discovery-Elemente
$dgServerRoles = $window.FindName("dgServerRoles")
$dgDnsConfig = $window.FindName("dgDnsConfig")
$dgDhcpConfig = $window.FindName("dgDhcpConfig")
$lvADStructure = $window.FindName("lvADStructure")

# Export-Button
$btnExportHTML = $window.FindName("btnExportHTML")
$btnPrintReport = $window.FindName("btnPrintReport")

# Validierungs-Elemente (aktualisiert - keine ComboBox mehr)
$txtLocalServerName = $window.FindName("txtLocalServerName")
$btnStartValidation = $window.FindName("btnStartValidation")
$tabValidation = $window.FindName("tabValidation")
$txtValidationStatus = $window.FindName("txtValidationStatus")
$dgValidationSummary = $window.FindName("dgValidationSummary")
$txtValidationReport = $window.FindName("txtValidationReport")

# Validierungs-Prüfoptionen
$chkValidateBasics = $window.FindName("chkValidateBasics")
$chkValidateAD = $window.FindName("chkValidateAD")
$chkValidateDNS = $window.FindName("chkValidateDNS")
$chkValidateDHCP = $window.FindName("chkValidateDHCP")
$chkValidateRoles = $window.FindName("chkValidateRoles")
$chkValidateGPO = $window.FindName("chkValidateGPO")
$chkValidatePerf = $window.FindName("chkValidatePerf")
$chkValidateServices = $window.FindName("chkValidateServices")

# Validierungs-Detailgrids
$dgValidationAD = $window.FindName("dgValidationAD")
$dgValidationReplication = $window.FindName("dgValidationReplication")
$dgValidationFSMO = $window.FindName("dgValidationFSMO")
$dgValidationDNSService = $window.FindName("dgValidationDNSService")
$dgValidationDNSZones = $window.FindName("dgValidationDNSZones")
$dgValidationDNSResolution = $window.FindName("dgValidationDNSResolution")
$dgValidationDHCPService = $window.FindName("dgValidationDHCPService")
$dgValidationDHCPScopes = $window.FindName("dgValidationDHCPScopes")
$dgValidationDHCPFailover = $window.FindName("dgValidationDHCPFailover")
$dgValidationServices = $window.FindName("dgValidationServices")
$dgValidationPerformance = $window.FindName("dgValidationPerformance")
$dgValidationEventLogs = $window.FindName("dgValidationEventLogs")
$dgValidationGPO = $window.FindName("dgValidationGPO")
$dgValidationSYSVOL = $window.FindName("dgValidationSYSVOL")

# Validierungs-Exportbuttons
$btnSaveValidationHTML = $window.FindName("btnSaveValidationHTML")
$btnPrintValidation = $window.FindName("btnPrintValidation")

# Footer
$linkWebsite = $window.FindName("linkWebsite")
#endregion

#region GUI-Funktionen
# Funktion zum Wechseln der Ansicht
function Switch-View {
    param (
        [string]$ViewName,
        [string]$TabTitle
    )
    
    try {
        # Alle Panels ausblenden
        $panelHome.Visibility = "Collapsed"
        $panelDiscovery.Visibility = "Collapsed"
        $panelInstallation.Visibility = "Collapsed"
        $panelMigration.Visibility = "Collapsed"
        $panelValidation.Visibility = "Collapsed"
        $panelDecommission.Visibility = "Collapsed"
        
        # Ausgewähltes Panel anzeigen
        switch ($ViewName) {
            "Home" { $panelHome.Visibility = "Visible" }
            "Discovery" { $panelDiscovery.Visibility = "Visible" }
            "Installation" { $panelInstallation.Visibility = "Visible" }
            "Migration" { $panelMigration.Visibility = "Visible" }
            "Validation" { $panelValidation.Visibility = "Visible" }
            "Decommission" { $panelDecommission.Visibility = "Visible" }
        }
        
        # Tab-Titel aktualisieren
        $txtCurrentTab.Text = $TabTitle
        
        Write-Log "Ansicht gewechselt zu: $ViewName" -Level "INFO"
    }
    catch {
        Write-Log "Fehler beim Wechseln der Ansicht: $_" -Level "ERROR"
        [System.Windows.Forms.MessageBox]::Show("Ein unerwarteter Fehler ist aufgetreten. Bitte wenden Sie sich an den Support.", "Fehler", "OK", "Error")
    }
}

# Funktion zum Öffnen einer URL
function Open-URL {
    param (
        [string]$URL
    )
    
    try {
        Start-Process $URL
        Write-Log "URL geöffnet: $URL" -Level "INFO"
    }
    catch {
        Write-Log "Fehler beim Öffnen der URL: $_" -Level "ERROR"
        [System.Windows.Forms.MessageBox]::Show("Fehler beim Öffnen der URL.", "Fehler", "OK", "Error")
    }
}
#endregion

#region Self-Diagnostics
function Test-RequiredAssets {
    try {
        $assetsPath = "$PSScriptRoot\assets"
        $requiredIcons = @("info.png", "settings.png", "close.png")
        $missingIcons = @()

        # Erstellen des Assets-Verzeichnisses, falls es nicht existiert
        if (-not (Test-Path $assetsPath)) {
            New-Item -Path $assetsPath -ItemType Directory -Force | Out-Null
            Write-Log "Assets-Verzeichnis wurde erstellt: $assetsPath" -Level "INFO"
        }

        foreach ($icon in $requiredIcons) {
            if (-not (Test-Path "$assetsPath\$icon")) {
                $missingIcons += $icon
                # Versuchen, ein leeres Icon zu erstellen, damit die Anwendung nicht abstürzt
                try {
                    $emptyIcon = [System.Drawing.Bitmap]::new(24, 24)
                    $emptyIcon.Save("$assetsPath\$icon")
                    Write-Log "Leeres Platzhalter-Icon erstellt für: $icon" -Level "WARNING"
                } catch {
                    Write-Log "Konnte kein Platzhalter-Icon erstellen für: $icon" -Level "WARNING"
                }
            }
        }

        if ($missingIcons.Count -gt 0) {
            Write-Log "WARNUNG: Fehlende Icons wurden durch Platzhalter ersetzt: $($missingIcons -join ', ')" -Level "WARNING"
            return $false
        }
        
        return $true
    }
    catch {
        Write-Log "Fehler bei der Überprüfung der benötigten Assets: $_" -Level "ERROR"
        return $false
    }
}

# Self-Diagnostics beim Start ausführen
$assetsOK = Test-RequiredAssets
if (-not $assetsOK) {
    [System.Windows.Forms.MessageBox]::Show("Es fehlen einige erforderliche Ressourcen. Die Anwendung kann möglicherweise nicht korrekt funktionieren.", "Warnung", "OK", "Warning")
}
#endregion

#region Event-Handler
# Navigation Event-Handler
$btnHome.Add_Click({ Switch-View -ViewName "Home" -TabTitle "Übersicht" })
$btnDiscovery.Add_Click({ Switch-View -ViewName "Discovery" -TabTitle "Ermittlung" })
$btnInstallation.Add_Click({ Switch-View -ViewName "Installation" -TabTitle "Installation" })
$btnMigration.Add_Click({ Switch-View -ViewName "Migration" -TabTitle "Migration" })
$btnValidation.Add_Click({ 
    Switch-View -ViewName "Validation" -TabTitle "Validierung"
    # Initialisieren der Validierungs-UI
    Initialize-ValidationUI
})
$btnDecommission.Add_Click({ Switch-View -ViewName "Decommission" -TabTitle "Abschaltung" })

# Steuerungselemente Event-Handler
$btnInfo.Add_Click({
    [System.Windows.Forms.MessageBox]::Show("Windows Server 2012 Migration Tool`nVersion 0.0.1`n`nDieses Tool unterstützt die Migration von Windows Server 2012 Domänencontrollern, DNS- und DHCP-Diensten auf neue Windows Server.", "Information", "OK", "Information")
})

$btnSettings.Add_Click({
    [System.Windows.Forms.MessageBox]::Show("Einstellungen sind in dieser Version noch nicht verfügbar.", "Einstellungen", "OK", "Information")
})

$btnClose.Add_Click({
    $window.Close()
})

# Footer Event-Handler
$linkWebsite.Add_Click({
    Open-URL -URL "https://support.microsoft.com/"
})

# Discovery Event-Handler
$btnStartDiscovery.Add_Click({
    try {
        # Cursor auf "Warten" setzen
        $window.Cursor = [System.Windows.Input.Cursors]::Wait
        
        # Analyse durchführen
        $analysisResults = Get-ADEnvironmentInfo
        
        if ($null -ne $analysisResults) {
            # Ergebnisse anzeigen
            Show-EnvironmentAnalysis -AnalysisResults $analysisResults
            Write-Log "Umgebungsanalyse erfolgreich durchgeführt" -Level "SUCCESS"
        } else {
            Write-Log "Umgebungsanalyse konnte nicht abgeschlossen werden" -Level "ERROR"
            [System.Windows.Forms.MessageBox]::Show("Die Analyse konnte nicht abgeschlossen werden. Bitte prüfen Sie die Logs für Details.", "Analyse fehlgeschlagen", "OK", "Error")
        }
    }
    catch {
        Write-Log "Fehler bei der Umgebungsanalyse: $($_.Exception.Message)" -Level "ERROR"
        [System.Windows.Forms.MessageBox]::Show("Ein Fehler ist bei der Umgebungsanalyse aufgetreten: $_", "Fehler", "OK", "Error")
    }
    finally {
        # Cursor zurücksetzen
        $window.Cursor = $null
    }
})

# HTML-Export Button Handler
$btnExportHTML.Add_Click({
    try {
        if ($null -eq $script:lastAnalysisResults) {
            [System.Windows.Forms.MessageBox]::Show("Es sind keine Analyseergebnisse zum Exportieren verfügbar. Bitte führen Sie zuerst eine Analyse durch.", "Keine Daten", "OK", "Warning")
            return
        }

        # Speicherdialog anzeigen
        $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
        $saveDialog.Filter = "HTML-Dateien (*.html)|*.html|Alle Dateien (*.*)|*.*"
        $saveDialog.Title = "HTML-Bericht speichern"
        $saveDialog.DefaultExt = "html"
        $saveDialog.FileName = "WS2012_Migration_Analyse_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
        
        $result = $saveDialog.ShowDialog()
        
        if ($result -eq $true) {
            $window.Cursor = [System.Windows.Input.Cursors]::Wait
            
            $exportSuccess = Export-AnalysisToHTML -AnalysisResults $script:lastAnalysisResults -FilePath $saveDialog.FileName
            
            if ($exportSuccess) {
                [System.Windows.Forms.MessageBox]::Show("Der Bericht wurde erfolgreich gespeichert:`n$($saveDialog.FileName)", "Export erfolgreich", "OK", "Information")
                
                # Optionally open the HTML report
                Start-Process $saveDialog.FileName
            }
            else {
                [System.Windows.Forms.MessageBox]::Show("Beim Exportieren des Berichts ist ein Fehler aufgetreten. Bitte prüfen Sie die Logdateien.", "Exportfehler", "OK", "Error")
            }
            
            $window.Cursor = $null
        }
    }
    catch {
        Write-Log "Fehler beim HTML-Export: $($_.Exception.Message)" -Level "ERROR"
        [System.Windows.Forms.MessageBox]::Show("Beim Exportieren ist ein Fehler aufgetreten: $_", "Fehler", "OK", "Error")
        $window.Cursor = $null
    }
})

# Bericht drucken Handler
$btnPrintReport.Add_Click({
    try {
        [System.Windows.Forms.MessageBox]::Show("Die Druckfunktion ist in dieser Version noch nicht implementiert. Bitte exportieren Sie den Bericht als HTML-Datei und drucken Sie ihn aus dem Browser.", "Information", "OK", "Information")
    }
    catch {
        Write-Log "Fehler beim Verarbeiten des Druckbefehls: $($_.Exception.Message)" -Level "ERROR"
    }
})

# Event-Handler für den Validierungsstart (aktualisiert)
$btnStartValidation.Add_Click({
    try {
        # Cursor auf "Warten" setzen
        $window.Cursor = [System.Windows.Input.Cursors]::Wait
        
        # Lokalen Server verwenden
        $serverToValidate = $env:COMPUTERNAME
        
        # Validation Tab Control zurücksetzen
        if ($null -ne $tabValidation) {
            $tabValidation.SelectedIndex = 0
        }
        
        # Validierungsstatus sichtbar machen und aktualisieren
        $txtValidationStatus.Visibility = "Visible"
        $txtValidationStatus.Text = "Validierung läuft..."
        
        # Überprüfen, welche Kategorien validiert werden sollen
        $validateBasics = $chkValidateBasics.IsChecked
        $validateAD = $chkValidateAD.IsChecked
        $validateDNS = $chkValidateDNS.IsChecked
        $validateDHCP = $chkValidateDHCP.IsChecked
        $validateRoles = $chkValidateRoles.IsChecked
        $validateGPO = $chkValidateGPO.IsChecked
        $validatePerf = $chkValidatePerf.IsChecked
        $validateServices = $chkValidateServices.IsChecked
        
        # Validierung durchführen
        $validationResults = Start-ServerValidation -ServerName $serverToValidate -ValidateBasics:$validateBasics -ValidateAD:$validateAD `
                            -ValidateDNS:$validateDNS -ValidateDHCP:$validateDHCP -ValidateRoles:$validateRoles -ValidateGPO:$validateGPO `
                            -ValidatePerf:$validatePerf -ValidateServices:$validateServices
        
        if ($null -ne $validationResults) {
            # TabControl und Export-Buttons sichtbar machen
            $tabValidation.Visibility = "Visible"
            $btnSaveValidationHTML.Visibility = "Visible"
            $btnPrintValidation.Visibility = "Visible"
            
            # Ergebnisse anzeigen
            Show-ValidationResults -ValidationResults $validationResults
            Write-Log "Servervalidierung für '$serverToValidate' erfolgreich durchgeführt" -Level "SUCCESS"
        } else {
            $txtValidationStatus.Text = "Validierung fehlgeschlagen"
            $txtValidationStatus.Foreground = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Colors]::Red)
            Write-Log "Servervalidierung für '$serverToValidate' konnte nicht abgeschlossen werden" -Level "ERROR"
            [System.Windows.Forms.MessageBox]::Show("Die Servervalidierung konnte nicht abgeschlossen werden. Bitte prüfen Sie die Logs für Details.", "Validierung fehlgeschlagen", "OK", "Error")
        }
    }
    catch {
        $txtValidationStatus.Text = "Fehler bei der Validierung"
        $txtValidationStatus.Foreground = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Colors]::Red)
        $txtValidationStatus.Visibility = "Visible"
        Write-Log "Fehler bei der Servervalidierung: $($_.Exception.Message)" -Level "ERROR"
        [System.Windows.Forms.MessageBox]::Show("Ein Fehler ist bei der Servervalidierung aufgetreten: $_", "Fehler", "OK", "Error")
    }
    finally {
        # Cursor zurücksetzen
        $window.Cursor = $null
    }
})

# HTML-Export Button für Validierungsbericht
$btnSaveValidationHTML.Add_Click({
    try {
        if ($null -eq $script:lastValidationResults) {
            [System.Windows.Forms.MessageBox]::Show("Es sind keine Validierungsergebnisse zum Exportieren verfügbar. Bitte führen Sie zuerst eine Validierung durch.", "Keine Daten", "OK", "Warning")
            return
        }

        # Speicherdialog anzeigen
        $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
        $saveDialog.Filter = "HTML-Dateien (*.html)|*.html|Alle Dateien (*.*)|*.*"
        $saveDialog.Title = "Validierungsbericht speichern"
        $saveDialog.DefaultExt = "html"
        $saveDialog.FileName = "WS2012_Migration_Validierung_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
        
        $result = $saveDialog.ShowDialog()
        
        if ($result -eq $true) {
            $window.Cursor = [System.Windows.Input.Cursors]::Wait
            
            $exportSuccess = Export-ValidationToHTML -ValidationResults $script:lastValidationResults -FilePath $saveDialog.FileName
            
            if ($exportSuccess) {
                [System.Windows.Forms.MessageBox]::Show("Der Validierungsbericht wurde erfolgreich gespeichert:`n$($saveDialog.FileName)", "Export erfolgreich", "OK", "Information")
                
                # Optional den HTML-Bericht öffnen
                Start-Process $saveDialog.FileName
            }
            else {
                [System.Windows.Forms.MessageBox]::Show("Beim Exportieren des Validierungsberichts ist ein Fehler aufgetreten. Bitte prüfen Sie die Logdateien.", "Exportfehler", "OK", "Error")
            }
            
            $window.Cursor = $null
        }
    }
    catch {
        Write-Log "Fehler beim HTML-Export des Validierungsberichts: $($_.Exception.Message)" -Level "ERROR"
        [System.Windows.Forms.MessageBox]::Show("Beim Exportieren ist ein Fehler aufgetreten: $_", "Fehler", "OK", "Error")
        $window.Cursor = $null
    }
})

# Bericht drucken Handler für Validierung
$btnPrintValidation.Add_Click({
    try {
        [System.Windows.Forms.MessageBox]::Show("Die Druckfunktion ist in dieser Version noch nicht implementiert. Bitte exportieren Sie den Bericht als HTML-Datei und drucken Sie ihn aus dem Browser.", "Information", "OK", "Information")
    }
    catch {
        Write-Log "Fehler beim Verarbeiten des Druckbefehls für Validierungsbericht: $($_.Exception.Message)" -Level "ERROR"
    }
})

# Event-Handler für Servernamen-Auswahl
$rbLocalServer.Add_Click({
    $txtServerName.IsEnabled = $false
    $lvDomainServers.IsEnabled = $false
})

$rbRemoteServer.Add_Click({
    $txtServerName.IsEnabled = $true
    $lvDomainServers.IsEnabled = $true
})

# Event-Handler für Domänenserverliste füllen
$btnRefreshServers.Add_Click({
    try {
        # Cursor auf "Warten" setzen
        $window.Cursor = [System.Windows.Input.Cursors]::Wait
        
        # Domänenserver abrufen
        $servers = Get-DomainServers
        
        if ($servers.Count -gt 0) {
            # ListView mit Servern füllen
            $lvDomainServers.ItemsSource = $servers
            Write-Log "Serverliste erfolgreich aktualisiert - $($servers.Count) Server gefunden" -Level "INFO"
        } else {
            [System.Windows.Forms.MessageBox]::Show("Es wurden keine Domänenserver gefunden.", "Information", "OK", "Information")
            Write-Log "Keine Domänenserver gefunden" -Level "WARNING"
        }
    }
    catch {
        Write-Log "Fehler beim Abrufen der Domänenserver: $($_.Exception.Message)" -Level "ERROR"
        [System.Windows.Forms.MessageBox]::Show("Fehler beim Abrufen der Serverliste: $_", "Fehler", "OK", "Error")
    }
    finally {
        # Cursor zurücksetzen
        $window.Cursor = $null
    }
})

# Event-Handler für Serverauswahl aus der Liste
$lvDomainServers.Add_SelectionChanged({
    if ($lvDomainServers.SelectedItem) {
        $selectedServer = $lvDomainServers.SelectedItem
        $txtServerName.Text = $selectedServer.Name
    }
})

# Event-Handler für "Alle auswählen" Button
$btnCheckAll.Add_Click({
    # Alle Checkboxen im Rollen-Bereich aktivieren
    $checkboxes = $window.FindName("panelInstallation").FindName("GroupBox").FindName("StackPanel").Children | 
                 Where-Object { $_ -is [System.Windows.Controls.CheckBox] }
    
    foreach ($checkbox in $checkboxes) {
        $checkbox.IsChecked = $true
    }
})

# Event-Handler für "Alle abwählen" Button
$btnUncheckAll.Add_Click({
    # Alle Checkboxen im Rollen-Bereich deaktivieren
    $checkboxes = $window.FindName("panelInstallation").FindName("GroupBox").FindName("StackPanel").Children | 
                 Where-Object { $_ -is [System.Windows.Controls.CheckBox] }
    
    foreach ($checkbox in $checkboxes) {
        $checkbox.IsChecked = $false
    }
})

# Event-Handler für "Voraussetzungen prüfen" Button
$btnCheckRequirements.Add_Click({
    try {
        # Cursor auf "Warten" setzen
        $window.Cursor = [System.Windows.Input.Cursors]::Wait
        
        # Zielserver ermitteln
        $targetServer = if ($rbLocalServer.IsChecked) {
            $env:COMPUTERNAME
        } else {
            $txtServerName.Text
        }
        
        if ([string]::IsNullOrEmpty($targetServer)) {
            [System.Windows.Forms.MessageBox]::Show("Bitte geben Sie einen Servernamen ein.", "Eingabefehler", "OK", "Warning")
            return
        }
        
        # Ausgewählte Rollen ermitteln
        $selectedRoles = @()
        
        # AD-Rollen
        if ($chkADDS.IsChecked) { $selectedRoles += "AD-Domain-Services" }
        if ($chkADCS.IsChecked) { $selectedRoles += "AD-Certificate" }
        if ($chkADLDS.IsChecked) { $selectedRoles += "ADLDS" }
        if ($chkADRMS.IsChecked) { $selectedRoles += "ADRMS" }
        
        # Netzwerkdienste
        if ($chkDNS.IsChecked) { $selectedRoles += "DNS" }
        if ($chkDHCP.IsChecked) { $selectedRoles += "DHCP" }
        if ($chkWINS.IsChecked) { $selectedRoles += "WINS" }
        
        # Dateidienste
        if ($chkFileServer.IsChecked) { $selectedRoles += "File-Services" }
        if ($chkDFS.IsChecked) { $selectedRoles += "FS-DFS-Namespace" }
        if ($chkDFSR.IsChecked) { $selectedRoles += "FS-DFS-Replication" }
        
        # Anwendungsdienste
        if ($chkWebServer.IsChecked) { $selectedRoles += "Web-Server" }
        if ($chkWSUS.IsChecked) { $selectedRoles += "UpdateServices" }
        
        if ($selectedRoles.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Bitte wählen Sie mindestens eine Rolle aus.", "Eingabefehler", "OK", "Warning")
            return
        }
        
        # Status aktualisieren
        $txtInstallationStatus.Text = "Überprüfe Voraussetzungen für $targetServer..."
        
        # Voraussetzungen prüfen
        $requirements = Test-ServerRoleInstallationRequirements -ServerName $targetServer -Roles $selectedRoles
        
        # Prüfergebnis anzeigen
        $resultText = "Ergebnis der Voraussetzungsprüfung für $targetServer -`n`n"
        
        foreach ($req in $requirements) {
            $statusSymbol = switch ($req.Status) {
                "OK" { "[✓]" }
                "WARNUNG" { "[!]" }
                "FEHLER" { "[✕]" }
                default { "[-]" }
            }
            
            $resultText += "$statusSymbol $($req.Requirement)$colon $($req.Details)`n"
        }
        
        # Prüfung auf kritische Fehler
        $hasCriticalErrors = ($requirements | Where-Object { $_.Status -eq "FEHLER" }).Count -gt 0
        
        # Status aktualisieren
        if ($hasCriticalErrors) {
            $txtInstallationStatus.Text = "Es wurden kritische Probleme festgestellt. Installation kann nicht fortgesetzt werden."
        } else {
            $txtInstallationStatus.Text = "Voraussetzungsprüfung abgeschlossen. Bereit für die Installation."
        }
        
        # Ergebnis anzeigen
        [System.Windows.Forms.MessageBox]::Show($resultText, "Voraussetzungsprüfung", "OK", "Information")
    }
    catch {
        Write-Log "Fehler bei der Überprüfung der Voraussetzungen: $($_.Exception.Message)" -Level "ERROR"
        [System.Windows.Forms.MessageBox]::Show("Fehler bei der Überprüfung der Voraussetzungen: $_", "Fehler", "OK", "Error")
    }
    finally {
        # Cursor zurücksetzen
        $window.Cursor = $null
    }
})

# Event-Handler für "Rollen installieren" Button
$btnInstallRoles.Add_Click({
    try {
        # Zielserver ermitteln
        $targetServer = if ($rbLocalServer.IsChecked) {
            $env:COMPUTERNAME
        } else {
            $txtServerName.Text
        }
        
        if ([string]::IsNullOrEmpty($targetServer)) {
            [System.Windows.Forms.MessageBox]::Show("Bitte geben Sie einen Servernamen ein.", "Eingabefehler", "OK", "Warning")
            return
        }
        
        # Ausgewählte Rollen ermitteln
        $selectedRoles = @()
        
        # AD-Rollen
        if ($chkADDS.IsChecked) { $selectedRoles += "AD-Domain-Services" }
        if ($chkADCS.IsChecked) { $selectedRoles += "AD-Certificate" }
        if ($chkADLDS.IsChecked) { $selectedRoles += "ADLDS" }
        if ($chkADRMS.IsChecked) { $selectedRoles += "ADRMS" }
        
        # Netzwerkdienste
        if ($chkDNS.IsChecked) { $selectedRoles += "DNS" }
        if ($chkDHCP.IsChecked) { $selectedRoles += "DHCP" }
        if ($chkWINS.IsChecked) { $selectedRoles += "WINS" }
        
        # Dateidienste
        if ($chkFileServer.IsChecked) { $selectedRoles += "File-Services" }
        if ($chkDFS.IsChecked) { $selectedRoles += "FS-DFS-Namespace" }
        if ($chkDFSR.IsChecked) { $selectedRoles += "FS-DFS-Replication" }
        
        # Anwendungsdienste
        if ($chkWebServer.IsChecked) { $selectedRoles += "Web-Server" }
        if ($chkWSUS.IsChecked) { $selectedRoles += "UpdateServices" }
        
        if ($selectedRoles.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Bitte wählen Sie mindestens eine Rolle aus.", "Eingabefehler", "OK", "Warning")
            return
        }
        
        # Bestätigung vor der Installation
        $confirmMessage = "Möchten Sie die folgenden Rollen auf $targetServer installieren?$colon`n`n"
        $confirmMessage += ($selectedRoles -join "`n")
        
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            $confirmMessage, 
            "Installation bestätigen", 
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
            return
        }
        
        # Cursor auf "Warten" setzen
        $window.Cursor = [System.Windows.Input.Cursors]::Wait
        
        # Status aktualisieren
        $txtInstallationStatus.Text = "Installiere Rollen auf $targetServer..."
        
        # Fortschritts-Callback-Funktion
        $progressCallback = {
            param($message, $percentComplete)
            
            # Status im UI aktualisieren (Thread-Safe)
            $window.Dispatcher.Invoke([action]{
                $txtInstallationStatus.Text = $message
            })
        }
        
        # Rollen installieren
        $installationResult = Install-ServerRoles -ServerName $targetServer -Roles $selectedRoles -ProgressCallback $progressCallback
        
        # Installationsergebnis verarbeiten
        $resultText = "Installationsergebnis für $targetServer$colon`n`n"
        
        foreach ($result in $installationResult.Results) {
            $statusSymbol = switch ($result.Status) {
                "ERFOLGREICH" { "[✓]" }
                "FEHLER" { "[✕]" }
                default { "[-]" }
            }
            
            $resultText += "$statusSymbol $($result.Role)$colon $($result.Details)`n"
        }
        
        # Status aktualisieren
        if ($installationResult.Success) {
            $txtInstallationStatus.Text = "Installation erfolgreich abgeschlossen."
            
            if ($installationResult.RestartRequired) {
                $txtInstallationStatus.Text += " Ein Neustart ist erforderlich."
                $resultText += "`nEin Neustart des Servers ist erforderlich, um die Installation abzuschließen."
            }
        } else {
            $txtInstallationStatus.Text = "Bei der Installation sind Fehler aufgetreten."
        }
        
        # Ergebnis anzeigen
        [System.Windows.Forms.MessageBox]::Show($resultText, "Installationsergebnis", "OK", "Information")
        
        # Neustart, falls erforderlich
        if ($installationResult.RestartRequired) {
            $restartConfirm = [System.Windows.Forms.MessageBox]::Show(
                "Möchten Sie den Server $targetServer jetzt neu starten?", 
                "Neustart erforderlich", 
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question
            )
            
            if ($restartConfirm -eq [System.Windows.Forms.DialogResult]::Yes) {
                $txtInstallationStatus.Text = "Starte Server $targetServer neu..."
                
                # Neustart durchführen
                $restartResult = Restart-RemoteServer -ServerName $targetServer
                
                if ($restartResult) {
                    $txtInstallationStatus.Text = "Server $targetServer wurde neu gestartet."
                } else {
                    $txtInstallationStatus.Text = "Konnte Server $targetServer nicht neu starten oder Neustart bestätigen."
                }
            }
        }
    }
    catch {
        Write-Log "Fehler bei der Installation von Serverrollen: $($_.Exception.Message)" -Level "ERROR"
        [System.Windows.Forms.MessageBox]::Show("Fehler bei der Installation$colon $_", "Fehler", "OK", "Error")
        $txtInstallationStatus.Text = "Fehler bei der Installation$colon $($_.Exception.Message)"
    }
    finally {
        # Cursor zurücksetzen
        $window.Cursor = $null
    }
})

# Beim Laden des Installations-Panels die Domänenserverliste aktualisieren
$btnInstallation.Add_Click({
    # Nach dem Wechsel zur Installation-Ansicht
    Switch-View -ViewName "Installation" -TabTitle "Installation"
    
    # Standardmäßig lokalen Server auswählen
    $rbLocalServer.IsChecked = $true
    $txtServerName.IsEnabled = $false
    $lvDomainServers.IsEnabled = $false
    
    # Standardstatus setzen
    $txtInstallationStatus.Text = "Bereit für die Installation von Serverrollen."
    
    # Domänenserverliste aktualisieren
    try {
        # Cursor auf "Warten" setzen
        $window.Cursor = [System.Windows.Input.Cursors]::Wait
        
        # Domänenserver abrufen
        $servers = Get-DomainServers
        
        if ($servers.Count -gt 0) {
            # ListView mit Servern füllen
            $lvDomainServers.ItemsSource = $servers
            Write-Log "Serverliste für Installation erfolgreich geladen - $($servers.Count) Server gefunden" -Level "INFO"
        }
    }
    catch {
        Write-Log "Fehler beim Laden der Domänenserverliste: $($_.Exception.Message)" -Level "WARNING"
    }
    finally {
        # Cursor zurücksetzen
        $window.Cursor = $null
    }
})
#endregion

#region Hauptfunktionen für die Migration
function Get-ServerStatus {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ServerName
    )
    
    try {
        $results = @{
            ADStructure = @()
            # Weitere Eigenschaften hier...
        }
        
        # Forest-Informationen
        $forest = Get-ADForest -ErrorAction SilentlyContinue
        if ($forest) {
            $results.ADStructure += [PSCustomObject]@{
                Type = "Forest"
                Name = $forest.Name
                Description = "Gesamtstruktur (Forest-Functional-Level$colon $($forest.ForestMode))"
            }
            
            # Domains im Forest
            foreach ($domainName in $forest.Domains) {
                try {
                    $domain = Get-ADDomain -Identity $domainName -ErrorAction SilentlyContinue
                    if ($domain) {
                        $results.ADStructure += [PSCustomObject]@{
                            Type = "Domäne"
                            Name = $domain.Name
                            Description = "Domänen-Funktionsebene$colon $($domain.DomainMode)"
                        }
                        
                        # Sites für diese Domäne
                        $sites = Get-ADReplicationSite -Filter * -Server $domainName -ErrorAction SilentlyContinue
                        foreach ($site in $sites) {
                            $results.ADStructure += [PSCustomObject]@{
                                Type = "  Site"
                                Name = $site.Name
                                Description = "Active Directory-Standort"
                            }
                            
                            # Subnets für diesen Site
                            $subnets = Get-ADReplicationSubnet -Filter "site -eq '$($site.DistinguishedName)'" -Server $domainName -ErrorAction SilentlyContinue
                            foreach ($subnet in $subnets) {
                                $results.ADStructure += [PSCustomObject]@{
                                    Type = "    Subnetz"
                                    Name = $subnet.Name
                                    Description = "Zugewiesen zu Standort $($site.Name)"
                                }
                            }
                        }
                        
                        # Vertrauensstellungen
                        $trusts = Get-ADTrust -Filter * -Server $domainName -ErrorAction SilentlyContinue
                        foreach ($trust in $trusts) {
                            $results.ADStructure += [PSCustomObject]@{
                                Type = "  Vertrauensstellung"
                                Name = $trust.Name
                                Description = "Vertrauenstyp$colon $($trust.TrustType), Richtung$colon $($trust.TrustDirection)"
                            }
                        }
                    }
                }
                catch {
                    Write-Log "Fehler beim Abrufen von Details für Domäne $domainName - $($_.Exception.Message)" -Level "WARNING"
                }
            }
        }

        Write-Log "Active Directory-Strukturanalyse abgeschlossen" -Level "INFO"
        
        return $results
    }
    catch {
        Write-Log "Fehler bei Get-ServerStatus für $ServerName`: $($_.Exception.Message)" -Level "ERROR"
        return $null
    }
}

# Funktion zum Anzeigen der Analyseergebnisse
function Show-EnvironmentAnalysis {
    param (
        [Parameter(Mandatory = $true)]
        [object]$AnalysisResults
    )
    
    try {
        # DataGrid mit Domänencontrollern befüllen
        $dgDomainControllers.ItemsSource = $AnalysisResults.DomainControllers
        
        # ListView mit FSMO-Rollen befüllen
        $lvFSMORoles.ItemsSource = $AnalysisResults.FSMORoles
        
        # DataGrid mit Diensten befüllen
        $dgServices.ItemsSource = $AnalysisResults.ServiceServers
        
        # Neue Grids mit erweiterten Informationen befüllen
        $dgServerRoles.ItemsSource = $AnalysisResults.ServerRoles
        $dgDnsConfig.ItemsSource = $AnalysisResults.DnsConfig
        $dgDhcpConfig.ItemsSource = $AnalysisResults.DhcpConfig
        $lvADStructure.ItemsSource = $AnalysisResults.ADStructure
        
        # Zusammenfassung aktualisieren
        $txtSummary.Text = $AnalysisResults.Summary
        
        # Tab-Steuerung anzeigen und Export-Button aktivieren
        $tabDiscovery.Visibility = "Visible"
        $btnExportHTML.Visibility = "Visible"
        $btnPrintReport.Visibility = "Visible"
        
        # Globale Variable zum Speichern der Analyseergebnisse für Export
        $script:lastAnalysisResults = $AnalysisResults
        
        Write-Log "Analyseergebnisse erfolgreich angezeigt" -Level "SUCCESS"
    }
    catch {
        Write-Log "Fehler beim Anzeigen der Analyseergebnisse$colon $_" -Level "ERROR"
        [System.Windows.Forms.MessageBox]::Show("Ein Fehler ist beim Anzeigen der Analyseergebnisse aufgetreten.", "Fehler", "OK", "Error")
    }
}

# Funktion zum Exportieren der Analyseergebnisse als HTML-Bericht
function Export-AnalysisToHTML {
    param (
        [Parameter(Mandatory = $true)]
        [object]$AnalysisResults,
        [string]$FilePath
    )
    
    try {
        $htmlHeader = @"
<!DOCTYPE html>
<html>
<head>
    <title>Windows Server 2012 Migration - Umgebungsanalyse</title>
    <style>
        body { 
            font-family: 'Segoe UI', Arial, sans-serif; 
            margin: 20px; 
            color: #333;
            line-height: 1.5;
        }
        h1 { 
            color: #007ACC; 
            border-bottom: 1px solid #007ACC;
            padding-bottom: 10px;
        }
        h2 { 
            color: #007ACC; 
            margin-top: 20px;
            border-bottom: 1px solid #ccc;
            padding-bottom: 5px;
        }
        table { 
            border-collapse: collapse; 
            width: 100%; 
            margin-bottom: 20px;
        }
        th { 
            background-color: #007ACC; 
            color: white; 
            text-align: left; 
            padding: 8px; 
        }
        td { 
            padding: 8px; 
            border-bottom: 1px solid #ddd; 
        }
        tr:nth-child(even) { 
            background-color: #f9f9f9; 
        }
        tr:hover { 
            background-color: #f1f1f1; 
        }
        .summary { 
            background-color: #f5f5f5;
            padding: 15px;
            border-left: 4px solid #007ACC;
            white-space: pre-wrap;
            font-family: Consolas, monospace;
        }
        .footer {
            margin-top: 30px;
            text-align: center;
            font-size: 0.8em;
            color: #777;
        }
        .status-running { color: green; }
        .status-stopped { color: red; }
        .true-value { color: green; }
        .false-value { color: red; }
    </style>
</head>
<body>
    <h1>Windows Server 2012 Migration - Umgebungsanalyse</h1>
    <p><strong>Erstellungsdatum$colon</strong> $(Get-Date -Format "dd.MM.yyyy HH:mm:ss")</p>
"@

        $htmlDomainControllers = "<h2>Domänencontroller</h2>"
        if ($AnalysisResults.DomainControllers.Count -gt 0) {
            $htmlDomainControllers += "<table><tr><th>Name</th><th>Standort</th><th>IP-Adresse</th><th>Betriebssystem</th><th>Version</th><th>Global Catalog</th><th>Erreichbar</th></tr>"
            foreach ($dc in $AnalysisResults.DomainControllers) {
                $gcClass = if ($dc.IsGlobalCatalog) { "true-value" } else { "false-value" }
                $onlineClass = if ($dc.IsOnline) { "true-value" } else { "false-value" }
                
                $htmlDomainControllers += "<tr><td>$($dc.Name)</td><td>$($dc.Site)</td><td>$($dc.IPAddress)</td><td>$($dc.OperatingSystem)</td><td>$($dc.OSVersion)</td><td class='$gcClass'>$($dc.IsGlobalCatalog)</td><td class='$onlineClass'>$($dc.IsOnline)</td></tr>"
            }
            $htmlDomainControllers += "</table>"
        } else {
            $htmlDomainControllers += "<p>Keine Domänencontroller gefunden.</p>"
        }

        $htmlFSMORoles = "<h2>FSMO-Rollen</h2>"
        if ($AnalysisResults.FSMORoles.Count -gt 0) {
            $htmlFSMORoles += "<table><tr><th>Rollenname</th><th>Inhaber</th></tr>"
            foreach ($role in $AnalysisResults.FSMORoles) {
                $htmlFSMORoles += "<tr><td>$($role.RoleName)</td><td>$($role.RoleOwner)</td></tr>"
            }
            $htmlFSMORoles += "</table>"
        } else {
            $htmlFSMORoles += "<p>Keine FSMO-Rollen gefunden.</p>"
        }

        $htmlServices = "<h2>DNS/DHCP-Server</h2>"
        if ($AnalysisResults.ServiceServers.Count -gt 0) {
            $htmlServices += "<table><tr><th>Servername</th><th>DNS-Dienst</th><th>DHCP-Dienst</th><th>DNS-Status</th><th>DHCP-Status</th><th>Betriebssystem</th></tr>"
            foreach ($server in $AnalysisResults.ServiceServers) {
                $dnsClass = if ($server.HasDNS) { "true-value" } else { "false-value" }
                $dhcpClass = if ($server.HasDHCP) { "true-value" } else { "false-value" }
                $dnsStatusClass = if ($server.DNSStatus -eq "Running") { "status-running" } else { "status-stopped" }
                $dhcpStatusClass = if ($server.DHCPStatus -eq "Running") { "status-running" } else { "status-stopped" }
                
                $htmlServices += "<tr><td>$($server.ServerName)</td><td class='$dnsClass'>$($server.HasDNS)</td><td class='$dhcpClass'>$($server.HasDHCP)</td><td class='$dnsStatusClass'>$($server.DNSStatus)</td><td class='$dhcpStatusClass'>$($server.DHCPStatus)</td><td>$($server.OperatingSystem)</td></tr>"
            }
            $htmlServices += "</table>"
        } else {
            $htmlServices += "<p>Keine DNS/DHCP-Server gefunden.</p>"
        }

        $htmlServerRoles = "<h2>Installierte Rollen</h2>"
        if ($AnalysisResults.ServerRoles.Count -gt 0) {
            $htmlServerRoles += "<table><tr><th>Servername</th><th>Rollenname</th><th>Status</th><th>Features</th></tr>"
            foreach ($role in $AnalysisResults.ServerRoles) {
                $htmlServerRoles += "<tr><td>$($role.ServerName)</td><td>$($role.RoleName)</td><td>$($role.InstallState)</td><td>$($role.Features)</td></tr>"
            }
            $htmlServerRoles += "</table>"
        } else {
            $htmlServerRoles += "<p>Keine installierten Rollen gefunden.</p>"
        }

        $htmlDnsConfig = "<h2>DNS-Konfiguration</h2>"
        if ($AnalysisResults.DnsConfig.Count -gt 0) {
            $htmlDnsConfig += "<table><tr><th>DNS-Server</th><th>Zonen</th><th>Forwarders</th><th>Zonen-Replikation</th><th>Rekursion aktiviert</th></tr>"
            foreach ($dnsServer in $AnalysisResults.DnsConfig) {
                $recursionClass = if ($dnsServer.RecursionEnabled) { "true-value" } else { "false-value" }
        $htmlDhcpConfig = "<h2>DHCP-Konfiguration</h2>"
        if ($AnalysisResults.DhcpConfig.Count -gt 0) {
            $htmlDhcpConfig += "<table><tr><th>DHCP-Server</th><th>Bereiche</th><th>Optionen</th><th>Failover-Konfiguration</th></tr>"
            foreach ($dhcpServer in $AnalysisResults.DhcpConfig) {
                $failoverClass = if ($dhcpServer.HasFailover) { "true-value" } else { "false-value" }
                $htmlDhcpConfig += "<tr><td>$($dhcpServer.ServerName)</td><td>$($dhcpServer.Scopes)</td><td>$($dhcpServer.ServerOptions)</td><td class='$failoverClass'>$($dhcpServer.HasFailover)</td></tr>"
            }
            $htmlDhcpConfig += "</table>"
        } else {
            $htmlDhcpConfig += "<p>Keine DHCP-Konfiguration gefunden.</p>"
        }

        $htmlADStructure = "<h2>Active Directory-Struktur</h2>"
        if ($AnalysisResults.ADStructure.Count -gt 0) {
            $htmlADStructure += "<table><tr><th>Typ</th><th>Name</th><th>Beschreibung</th></tr>"
            foreach ($item in $AnalysisResults.ADStructure) {
                $htmlADStructure += "<tr><td>$($item.Type)</td><td>$($item.Name)</td><td>$($item.Description)</td></tr>"
            }
            $htmlADStructure += "</table>"
        } else {
            $htmlADStructure += "<p>Keine Active Directory-Strukturdaten gefunden.</p>"
        }

        $htmlSummary = "<h2>Zusammenfassung</h2><pre class='summary'>$($AnalysisResults.Summary)</pre>"

        $htmlFooter = @"
    <div class="footer">
        <p>Erstellt mit Windows Server 2012 Migration Tool | $(Get-Date -Format "yyyy")</p>
    </div>
</body>
</html>
"@

        $fullHTML = $htmlHeader + $htmlDomainControllers + $htmlFSMORoles + $htmlServices + $htmlServerRoles + $htmlDnsConfig + $htmlDhcpConfig + $htmlADStructure + $htmlSummary + $htmlFooter

        # Datei speichern
        $fullHTML | Out-File -FilePath $FilePath -Encoding UTF8

        Write-Log "Analyse-Bericht erfolgreich als HTML exportiert nach $FilePath" -Level "SUCCESS"
        return $true
    }
    catch {
        Write-Log "Fehler beim Exportieren der Analyse als HTML$colon $($_.Exception.Message)" -Level "ERROR"
        return $false
    }
}
#endregion

#region Validierungsfunktionen
# Funktion zum Initialisieren der Validierungs-UI - vereinfacht für lokalen Server
function Initialize-ValidationUI {
    try {
        # Lokalen Server automatisch auswählen und anzeigen
        $localServer = $env:COMPUTERNAME
        
        # Servernamen anzeigen
        if ($null -ne $txtLocalServerName) {
            $txtLocalServerName.Text = $localServer
        }
        
        Write-Log "Validierungs-UI erfolgreich für lokalen Server $localServer initialisiert" -Level "INFO"
    }
    catch {
        Write-Log "Fehler beim Initialisieren der Validierungs-UI$colon $($_.Exception.Message)" -Level "ERROR"
    }
}

# Hauptfunktion für die Servervalidierung
function Start-ServerValidation {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ServerName,
        [bool]$ValidateBasics = $true,
        [bool]$ValidateAD = $true,
        [bool]$ValidateDNS = $true,
        [bool]$ValidateDHCP = $true,
        [bool]$ValidateRoles = $true,
        [bool]$ValidateGPO = $true,
        [bool]$ValidatePerf = $true,
        [bool]$ValidateServices = $true
    )
    
    try {
        Write-Log "Starte Validierung für Server$colon $ServerName" -Level "INFO"
        
        $results = @{
            ServerName = $ServerName
            Timestamp = Get-Date
            Summary = @()
            BasicTests = @()
            ADTests = @()
            Replication = @()
            FSMORoles = @()
            DNSService = @()
            DNSZones = @()
            DNSResolution = @()
            DHCPService = @()
            DHCPScopes = @()
            DHCPFailover = @()
            Services = @()
            Performance = @()
            EventLogs = @()
            GPO = @()
            SYSVOL = @()
            Report = ""
        }
        
        # Server-Erreichbarkeit prüfen
        $pingSuccessful = Test-Connection -ComputerName $ServerName -Count 1 -Quiet
        
        if (-not $pingSuccessful) {
            Write-Log "Server $ServerName ist nicht erreichbar (Ping fehlgeschlagen)" -Level "ERROR"
            
            # Minimalen Ergebnissatz zurückgeben mit Fehlermeldung
            $results.Summary += [PSCustomObject]@{
                Category = "Grundlegend"
                Component = "Konnektivität"
                Status = "FEHLER"
                Details = "Server ist nicht per ICMP (Ping) erreichbar"
            }
            
            $results.BasicTests += [PSCustomObject]@{
                TestName = "Ping-Erreichbarkeit"
                Result = "FEHLER"
                Details = "Server antwortet nicht auf Ping-Anfragen"
            }
            
            $results.Report = "SERVER NICHT ERREICHBAR$colon $ServerName konnte nicht erreicht werden (Ping fehlgeschlagen)."
            
            return $results
        }
        
        # Grundlegende Tests durchführen (immer)
        $basicTests = Test-ServerBasics -ServerName $ServerName
        $results.BasicTests = $basicTests
        
        # Zusammenfassung der grundlegenden Tests hinzufügen
        foreach ($basicTest in $basicTests) {
            $results.Summary += [PSCustomObject]@{
                Category = "Grundlegend"
                Component = $basicTest.TestName
                Status = $basicTest.Result
                Details = $basicTest.Details
            }
        }
        
        # Active Directory Tests
        if ($ValidateAD) {
            # AD-Dienst validieren
            $adTests = Test-ADServices -ServerName $ServerName
            $results.ADTests = $adTests
            
            # Replikationsstatus
            $replTests = Test-ADReplication -ServerName $ServerName
            $results.Replication = $replTests
            
            # Zusammenfassung der AD-Tests hinzufügen
            foreach ($adTest in $adTests) {
                $results.Summary += [PSCustomObject]@{
                    Category = "Active Directory"
                    Component = $adTest.TestName
                    Status = $adTest.Result
                    Details = $adTest.Details
                }
            }
            
            # Zusammenfassung der Replikationstests hinzufügen
            $replStatus = if ($replTests.Count -gt 0) {
                $failedRepl = $replTests | Where-Object { $_.Status -ne "OK" }
                if ($failedRepl) {
                    "WARNUNG"
                } else {
                    "OK"
                }
            } else {
                "Nicht geprüft"
            }
            
            $results.Summary += [PSCustomObject]@{
                Category = "Active Directory"
                Component = "Replikation"
                Status = $replStatus
                Details = "Replikation mit $($replTests.Count) Partnern geprüft"
            }
        }
        
        # FSMO-Rollen prüfen
        if ($ValidateRoles) {
            $fsmoTests = Test-FSMORoles -ServerName $ServerName
            $results.FSMORoles = $fsmoTests
            
            # Zusammenfassung der FSMO-Tests hinzufügen, wenn Rollen auf dem Server vorhanden
            $fsmoOnServer = $fsmoTests | Where-Object { $_.Owner -like "*$ServerName*" }
            if ($fsmoOnServer) {
                $fsmoStatus = if (($fsmoOnServer | Where-Object { $_.Availability -ne "OK" }).Count -gt 0) {
                    "WARNUNG"
                } else {
                    "OK"
                }
                
                $results.Summary += [PSCustomObject]@{
                    Category = "Active Directory"
                    Component = "FSMO-Rollen"
                    Status = $fsmoStatus
                    Details = "$($fsmoOnServer.Count) FSMO-Rollen auf diesem Server"
                }
            }
        }
        
        # DNS-Tests durchführen
        if ($ValidateDNS) {
            # DNS-Dienst validieren
            $dnsServiceTests = Test-DNSService -ServerName $ServerName
            $results.DNSService = $dnsServiceTests
            
            # DNS-Zusammenfassungsstatus hinzufügen
            $dnsServiceStatus = if ($dnsServiceTests.Count -gt 0) {
                $failedDns = $dnsServiceTests | Where-Object { $_.Result -ne "OK" }
                if ($failedDns) {
                    "FEHLER"
                } else {
                    "OK"
                }
            } else {
                "Nicht installiert"
            }
            
            $results.Summary += [PSCustomObject]@{
                Category = "DNS"
                Component = "DNS-Dienst"
                Status = $dnsServiceStatus
                Details = if ($dnsServiceStatus -eq "Nicht installiert") { "DNS-Rolle nicht installiert" } else { "DNS-Dienst wurde geprüft" }
            }
            
            # Wenn DNS installiert ist, weitere Tests durchführen
            if ($dnsServiceStatus -ne "Nicht installiert") {
                # DNS-Zonen validieren
                $dnsZones = Test-DNSZones -ServerName $ServerName
                $results.DNSZones = $dnsZones
                
                # DNS-Auflösungstests durchführen
                $dnsResolution = Test-DNSResolution -ServerName $ServerName
                $results.DNSResolution = $dnsResolution
                
                # DNS-Zonen-Zusammenfassung hinzufügen
                $zonesStatus = if ($dnsZones.Count -gt 0) {
                    $failedZones = $dnsZones | Where-Object { $_.Status -ne "OK" }
                    if ($failedZones.Count -gt 0) {
                        "WARNUNG"
                    } else {
                        "OK"
                    }
                } else {
                    "Keine Zonen"
                }
                
                $results.Summary += [PSCustomObject]@{
                    Category = "DNS"
                    Component = "DNS-Zonen"
                    Status = $zonesStatus
                    Details = "$($dnsZones.Count) DNS-Zonen geprüft"
                }
                
                # DNS-Auflösungszusammenfassung hinzufügen
                $resolutionStatus = if ($dnsResolution.Count -gt 0) {
                    $failedResolution = $dnsResolution | Where-Object { $_.Status -ne "OK" }
                    if ($failedResolution.Count -gt 0) {
                        "WARNUNG"
                    } else {
                        "OK"
                    }
                } else {
                    "Nicht getestet"
                }
                
                $results.Summary += [PSCustomObject]@{
                    Category = "DNS"
                    Component = "DNS-Auflösung"
                    Status = $resolutionStatus
                    Details = "$($dnsResolution.Count) DNS-Auflösungstests durchgeführt"
                }
            }
        }
        
        # DHCP-Tests durchführen
        if ($ValidateDHCP) {
            # DHCP-Dienst validieren
            $dhcpServiceTests = Test-DHCPService -ServerName $ServerName
            $results.DHCPService = $dhcpServiceTests
            
            # DHCP-Zusammenfassungsstatus hinzufügen
            $dhcpServiceStatus = if ($dhcpServiceTests.Count -gt 0) {
                $failedDhcp = $dhcpServiceTests | Where-Object { $_.Result -ne "OK" }
                if ($failedDhcp) {
                    "FEHLER"
                } else {
                    "OK"
                }
            } else {
                "Nicht installiert"
            }
            
            $results.Summary += [PSCustomObject]@{
                Category = "DHCP"
                Component = "DHCP-Dienst"
                Status = $dhcpServiceStatus
                Details = if ($dhcpServiceStatus -eq "Nicht installiert") { "DHCP-Rolle nicht installiert" } else { "DHCP-Dienst wurde geprüft" }
            }
            
            # Wenn DHCP installiert ist, weitere Tests durchführen
            if ($dhcpServiceStatus -ne "Nicht installiert") {
                # DHCP-Bereiche validieren
                $dhcpScopes = Test-DHCPScopes -ServerName $ServerName
                $results.DHCPScopes = $dhcpScopes
                
                # DHCP-Failover prüfen
                $dhcpFailover = Test-DHCPFailover -ServerName $ServerName
                $results.DHCPFailover = $dhcpFailover
                
                # DHCP-Bereichs-Zusammenfassung hinzufügen
                $scopesStatus = if ($dhcpScopes.Count -gt 0) {
                    $failedScopes = $dhcpScopes | Where-Object { $_.Status -ne "OK" }
                    if ($failedScopes.Count -gt 0) {
                        "WARNUNG"
                    } else {
                        "OK"
                    }
                } else {
                    "Keine Bereiche"
                }
                
                $results.Summary += [PSCustomObject]@{
                    Category = "DHCP"
                    Component = "DHCP-Bereiche"
                    Status = $scopesStatus
                    Details = "$($dhcpScopes.Count) DHCP-Bereiche geprüft"
                }
                
                # DHCP-Failover-Zusammenfassung hinzufügen, falls vorhanden
                if ($dhcpFailover.Count -gt 0) {
                    $failoverStatus = if (($dhcpFailover | Where-Object { $_.Status -ne "OK" }).Count -gt 0) {
                        "WARNUNG"
                    } else {
                        "OK"
                    }
                    
                    $results.Summary += [PSCustomObject]@{
                        Category = "DHCP"
                        Component = "DHCP-Failover"
                        Status = $failoverStatus
                        Details = "$($dhcpFailover.Count) DHCP-Failover-Beziehungen geprüft"
                    }
                }
            }
        }
        
        # Dienste-Tests durchführen
        if ($ValidateServices) {
            $serviceTests = Test-CriticalServices -ServerName $ServerName
            $results.Services = $serviceTests
            
            $servicesStatus = if ($serviceTests.Count -gt 0) {
                $failedServices = $serviceTests | Where-Object { $_.Status -ne "Running" }
                if ($failedServices.Count -gt 0) {
                    "WARNUNG"
                } else {
                    "OK"
                }
            } else {
                "Nicht geprüft"
            }
            
            $results.Summary += [PSCustomObject]@{
                Category = "Dienste"
                Component = "Kritische Dienste"
                Status = $servicesStatus
                Details = "$($serviceTests.Count) Dienste geprüft, $($serviceTests | Where-Object { $_.Status -ne 'Running' } | Measure-Object).Count mit Problemen"
            }
        }
        
        # Performance-Tests durchführen
        if ($ValidatePerf) {
            $perfTests = Test-ServerPerformance -ServerName $ServerName
            $results.Performance = $perfTests
            
            # Ereignisprotokolle prüfen
            $eventTests = Test-EventLogs -ServerName $ServerName
            $results.EventLogs = $eventTests
            
            $perfStatus = if ($perfTests.Count -gt 0) {
                $failedPerf = $perfTests | Where-Object { $_.Status -ne "OK" }
                if ($failedPerf.Count -gt 0) {
                    "WARNUNG"
                } else {
                    "OK"
                }
            } else {
                "Nicht geprüft"
            }
            
            $results.Summary += [PSCustomObject]@{
                Category = "Performance"
                Component = "Leistungsmetriken"
                Status = $perfStatus
                Details = "$($perfTests.Count) Performance-Metriken geprüft"
            }
            
            $eventStatus = if ($eventTests.Count -gt 0) {
                $criticalEvents = $eventTests | Where-Object { $_.Severity -eq "Error" -and $_.Count -gt 0 }
                if ($criticalEvents.Count -gt 0) {
                    "WARNUNG"
                } else {
                    "OK"
                }
            } else {
                "Nicht geprüft"
            }
            
            $results.Summary += [PSCustomObject]@{
                Category = "Ereignisse"
                Component = "Ereignisprotokolle"
                Status = $eventStatus
                Details = "$($eventTests.Count) Protokolle geprüft, $($eventTests | Where-Object { $_.Severity -eq 'Error' -and $_.Count -gt 0 } | Measure-Object).Count mit kritischen Fehlern"
            }
        }
        
        # Gruppenrichtlinien-Tests durchführen
        if ($ValidateGPO) {
            $gpoTests = Test-GroupPolicy -ServerName $ServerName
            $results.GPO = $gpoTests
            
            $sysvolTests = Test-SYSVOL -ServerName $ServerName
            $results.SYSVOL = $sysvolTests
            
            $gpoStatus = if ($gpoTests.Count -gt 0) {
                $failedGpo = $gpoTests | Where-Object { $_.ReplicationStatus -ne "OK" }
                if ($failedGpo.Count -gt 0) {
                    "WARNUNG"
                } else {
                    "OK"
                }
            } else {
                "Nicht geprüft"
            }
            
            $results.Summary += [PSCustomObject]@{
                Category = "Gruppenrichtlinien"
                Component = "GPO-Replikation"
                Status = $gpoStatus
                Details = "$($gpoTests.Count) Gruppenrichtlinienobjekte geprüft"
            }
            
            $sysvolStatus = if ($sysvolTests.Count -gt 0) {
                $failedSysvol = $sysvolTests | Where-Object { $_.Status -ne "OK" }
                if ($failedSysvol.Count -gt 0) {
                    "WARNUNG"
                } else {
                    "OK"
                }
            } else {
                "Nicht geprüft"
            }
            
            $results.Summary += [PSCustomObject]@{
                Category = "Gruppenrichtlinien"
                Component = "SYSVOL"
                Status = $sysvolStatus
                Details = "SYSVOL-Status auf $($sysvolTests.Count) Servern geprüft"
            }
        }
        
        # Textbericht erstellen
        $report = @"
=============================================
Validierungsbericht für Server$colon $ServerName
Zeitpunkt$colon $((Get-Date).ToString('dd.MM.yyyy HH:mm:ss'))
=============================================

ZUSAMMENFASSUNG:
"@
        
        foreach ($summaryItem in $results.Summary) {
            $statusText = switch ($summaryItem.Status) {
                "OK" { "[ERFOLG]  " }
                "WARNUNG" { "[WARNUNG] " }
                "FEHLER" { "[FEHLER]   " }
                default { "[$($summaryItem.Status)]" }
            }
            $report += "`n$statusText $($summaryItem.Category) - $($summaryItem.Component)$colon $($summaryItem.Details)"
        }
        
        $report += @"


=============================================
DETAILIERTE ERGEBNISSE:
=============================================

GRUNDLEGENDE TESTS:
"@
        
        foreach ($test in $results.BasicTests) {
            $report += "`n[$($test.Result)] $($test.TestName)$colon $($test.Details)"
        }
        
        if ($ValidateAD) {
            $report += @"


ACTIVE DIRECTORY TESTS:
"@
            foreach ($test in $results.ADTests) {
                $report += "`n[$($test.Result)] $($test.TestName)$colon $($test.Details)"
            }
            
            if ($results.Replication.Count -gt 0) {
                $report += @"


AD REPLIKATION:
"@
                foreach ($replItem in $results.Replication) {
                    $report += "`n[$($replItem.Status)] $($replItem.SourceServer) -> $($replItem.DestinationServer)$colon $($replItem.Details)"
                }
            }
        }
        
        if ($ValidateRoles -and $results.FSMORoles.Count -gt 0) {
            $report += @"


FSMO ROLLEN:
"@
            foreach ($role in $results.FSMORoles) {
                $report += "`n[$($role.Availability)] $($role.Role)$colon $($role.Owner)"
            }
        }
        
        # DNS-Berichtdetails
        if ($ValidateDNS -and $results.DNSService.Count -gt 0) {
            $report += @"


DNS-DIENST:
"@
            foreach ($dns in $results.DNSService) {
                $report += "`n[$($dns.Result)] $($dns.TestName)$colon $($dns.Details)"
            }
            
            if ($results.DNSZones.Count -gt 0) {
                $report += @"


DNS-ZONEN:
"@
                foreach ($zone in $results.DNSZones) {
                    $report += "`n[$($zone.Status)] $($zone.ZoneName) ($($zone.ZoneType))$colon $($zone.Details)"
                }
            }
            
            if ($results.DNSResolution.Count -gt 0) {
                $report += @"


DNS-AUFLÖSUNG:
"@
                foreach ($res in $results.DNSResolution) {
                    $report += "`n[$($res.Status)] $($res.Hostname) -> $($res.ResolvedIP)$colon $($res.ResponseTime) ms"
                }
            }
        }
        
        # DHCP-Berichtdetails
        if ($ValidateDHCP -and $results.DHCPService.Count -gt 0) {
            $report += @"


DHCP-DIENST:
"@
            foreach ($dhcp in $results.DHCPService) {
                $report += "`n[$($dhcp.Result)] $($dhcp.TestName)$colon $($dhcp.Details)"
            }
            
            if ($results.DHCPScopes.Count -gt 0) {
                $report += @"


DHCP-BEREICHE:
"@
                foreach ($scope in $results.DHCPScopes) {
                    $report += "`n[$($scope.Status)] $($scope.ScopeId) ($($scope.ScopeName))$colon $($scope.Details)"
                }
            }
            
            if ($results.DHCPFailover.Count -gt 0) {
                $report += @"


DHCP-FAILOVER:
"@
                foreach ($failover in $results.DHCPFailover) {
                    $report += "`n[$($failover.Status)] $($failover.ScopeId) mit $($failover.PartnerServer) ($($failover.Mode))$colon $($failover.Details)"
                }
            }
        }
        
        # Dienste und Performance Berichtdetails
        if ($ValidateServices -and $results.Services.Count -gt 0) {
            $report += @"


KRITISCHE DIENSTE:
"@
            foreach ($service in $results.Services) {
                $report += "`n[$($service.Status)] $($service.DisplayName) ($($service.ServiceName))$colon $($service.StartType)"
            }
        }
        
        if ($ValidatePerf) {
            if ($results.Performance.Count -gt 0) {
                $report += @"


PERFORMANCE-METRIKEN:
"@
                foreach ($perf in $results.Performance) {
                    $report += "`n[$($perf.Status)] $($perf.MetricName)$colon $($perf.Value) $($perf.Unit) - $($perf.Details)"
                }
            }
            
            if ($results.EventLogs.Count -gt 0) {
                $report += @"


EREIGNISPROTOKOLLE:
"@
                foreach ($event in $results.EventLogs) {
                    $report += "`n[$($event.Severity)] $($event.LogName)$colon $($event.Count) Ereignisse, letztes$colon $($event.LatestEvent)"
                }
            }
        }
        
        # Gruppenrichtlinien Berichtdetails
        if ($ValidateGPO) {
            if ($results.GPO.Count -gt 0) {
                $report += @"


GRUPPENRICHTLINIEN-REPLIKATION:
"@
                foreach ($gpo in $results.GPO) {
                    $report += "`n[$($gpo.ReplicationStatus)] $($gpo.GPOName) (Ver $($gpo.Version))$colon $($gpo.Details)"
                }
            }
            
            if ($results.SYSVOL.Count -gt 0) {
                $report += @"


SYSVOL-REPLIKATION:
"@
                foreach ($sysvol in $results.SYSVOL) {
                    $report += "`n[$($sysvol.Status)] $($sysvol.Server) ($($sysvol.ReplicationType))$colon $($sysvol.Details)"
                }
            }
        }
        
        $report += @"


=============================================
VALIDIERUNGSBERICHT ENDE
=============================================
"@
        
        $results.Report = $report
        
        Write-Log "Servervalidierung für $ServerName erfolgreich abgeschlossen" -Level "SUCCESS"
        return $results
    }
    catch {
        Write-Log "Kritischer Fehler bei der Servervalidierung für $ServerName`$colon $($_.Exception.Message)" -Level "ERROR"
        return $null
    }
}

# Funktion zum Anzeigen der Validierungsergebnisse
function Show-ValidationResults {
    param (
        [Parameter(Mandatory = $true)]
        [object]$ValidationResults
    )
    
    try {
        # Globale Variable zum Speichern der Validierungsergebnisse für Export setzen
        $script:lastValidationResults = $ValidationResults
        
        # Status aktualisieren
        $overallStatus = "OK"
        
        # Prüfen, ob Fehler oder Warnungen vorhanden sind
        if (($ValidationResults.Summary | Where-Object { $_.Status -eq "FEHLER" }).Count -gt 0) {
            $overallStatus = "FEHLER"
            $txtValidationStatus.Text = "Validierung abgeschlossen mit Fehlern"
            $txtValidationStatus.Foreground = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Colors]::Red)
        }
        elseif (($ValidationResults.Summary | Where-Object { $_.Status -eq "WARNUNG" }).Count -gt 0) {
            $overallStatus = "WARNUNG"
            $txtValidationStatus.Text = "Validierung abgeschlossen mit Warnungen"
            $txtValidationStatus.Foreground = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Colors]::Orange)
        }
        else {
            $txtValidationStatus.Text = "Validierung erfolgreich abgeschlossen"
            $txtValidationStatus.Foreground = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Colors]::Green)
        }
        
        # Zusammenfassung anzeigen
        $dgValidationSummary.ItemsSource = $ValidationResults.Summary
        
        # Active Directory Ergebnisse anzeigen
        $dgValidationAD.ItemsSource = $ValidationResults.ADTests
        $dgValidationReplication.ItemsSource = $ValidationResults.Replication
        $dgValidationFSMO.ItemsSource = $ValidationResults.FSMORoles
        
        # DNS Ergebnisse anzeigen
        $dgValidationDNSService.ItemsSource = $ValidationResults.DNSService
        $dgValidationDNSZones.ItemsSource = $ValidationResults.DNSZones
        $dgValidationDNSResolution.ItemsSource = $ValidationResults.DNSResolution
        
        # DHCP Ergebnisse anzeigen
        $dgValidationDHCPService.ItemsSource = $ValidationResults.DHCPService
        $dgValidationDHCPScopes.ItemsSource = $ValidationResults.DHCPScopes
        $dgValidationDHCPFailover.ItemsSource = $ValidationResults.DHCPFailover
        
        # Dienste und Performance Ergebnisse anzeigen
        $dgValidationServices.ItemsSource = $ValidationResults.Services
        $dgValidationPerformance.ItemsSource = $ValidationResults.Performance
        $dgValidationEventLogs.ItemsSource = $ValidationResults.EventLogs
        
        # Gruppenrichtlinien Ergebnisse anzeigen
        $dgValidationGPO.ItemsSource = $ValidationResults.GPO
        $dgValidationSYSVOL.ItemsSource = $ValidationResults.SYSVOL
        
        # Textbericht anzeigen
        $txtValidationReport.Text = $ValidationResults.Report
    }
    catch {
        Write-Log "Fehler beim Anzeigen der Validierungsergebnisse$colon $($_.Exception.Message)" -Level "ERROR"
        [System.Windows.Forms.MessageBox]::Show("Ein Fehler ist beim Anzeigen der Validierungsergebnisse aufgetreten. Details finden Sie im Logfile.", "Fehler", "OK", "Error")
    }
}

# Funktion zum Testen der grundlegenden Serverstatus
function Test-ServerBasics {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ServerName
    )
    
    $results = @()
    
    try {
        # Ping-Test
        $ping = Test-Connection -ComputerName $ServerName -Count 2 -ErrorAction SilentlyContinue
        $pingResult = if ($ping) {
            $avgTime = [math]::Round(($ping | Measure-Object -Property ResponseTime -Average).Average, 2)
            [PSCustomObject]@{
                TestName = "Ping-Erreichbarkeit"
                Result = "OK"
                Details = "Antwortzeit$colon $avgTime ms"
            }
        } else {
            [PSCustomObject]@{
                TestName = "Ping-Erreichbarkeit"
                Result = "FEHLER"
                Details = "Server antwortet nicht auf Ping-Anfragen"
            }
        }
        $results += $pingResult
        
        # Remote Registry Service für WMI-Zugriffe prüfen
        try {
            $regService = Get-Service -ComputerName $ServerName -Name "RemoteRegistry" -ErrorAction SilentlyContinue
            $regStatus = if ($regService.Status -eq "Running") {
                [PSCustomObject]@{
                    TestName = "Remote Registry"
                    Result = "OK"
                    Details = "Der Remote Registry Dienst läuft"
                }
            } else {
                [PSCustomObject]@{
                    TestName = "Remote Registry"
                    Result = "WARNUNG"
                    Details = "Der Remote Registry Dienst ist nicht gestartet (Status$colon $($regService.Status))"
                }
            }
            $results += $regStatus
        }
        catch {
            $results += [PSCustomObject]@{
                TestName = "Remote Registry"
                Result = "FEHLER"
                Details = "Konnte Remote Registry Dienst nicht prüfen$colon $($_.Exception.Message)"
            }
        }
        
        # Systeminfo abrufen
        try {
            $computerSystem = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $ServerName -ErrorAction Stop
            $os = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $ServerName -ErrorAction Stop
            
            # Betriebssystem-Informationen
            $results += [PSCustomObject]@{
                TestName = "Betriebssystem"
                Result = "OK"
                Details = "$($os.Caption) $($os.CSDVersion) (Build $($os.BuildNumber))"
            }
            
            # Uptime berechnen
            $bootTime = [Management.ManagementDateTimeConverter]::ToDateTime($os.LastBootUpTime)
            $uptime = (Get-Date) - $bootTime
            $uptimeString = "{0} Tage, {1} Stunden, {2} Minuten" -f $uptime.Days, $uptime.Hours, $uptime.Minutes
            
            $uptimeResult = if ($uptime.Days -gt 30) {
                [PSCustomObject]@{
                    TestName = "System-Uptime"
                    Result = "WARNUNG"
                    Details = "System läuft seit $uptimeString (Neustart empfohlen)"
                }
            } else {
                [PSCustomObject]@{
                    TestName = "System-Uptime"
                    Result = "OK"
                    Details = "System läuft seit $uptimeString"
                }
            }
            $results += $uptimeResult
            
            # Speicherauslastung
            $totalRAM = [math]::Round($computerSystem.TotalPhysicalMemory / 1GB, 2)
            $freeRAM = [math]::Round($os.FreePhysicalMemory / 1MB, 2)
            $usedRAM = $totalRAM - ($freeRAM / 1024)
            $ramPercentUsed = [math]::Round(($usedRAM / $totalRAM) * 100, 2)
            
            $ramResult = if ($ramPercentUsed -gt 90) {
                [PSCustomObject]@{
                    TestName = "Arbeitsspeicher"
                    Result = "WARNUNG"
                    Details = "Speicherauslastung bei $ramPercentUsed% ($usedRAM GB von $totalRAM GB verwendet)"
                }
            } else {
                [PSCustomObject]@{
                    TestName = "Arbeitsspeicher"
                    Result = "OK"
                    Details = "Speicherauslastung bei $ramPercentUsed% ($usedRAM GB von $totalRAM GB verwendet)"
                }
            }
            $results += $ramResult
        }
        catch {
            $results += [PSCustomObject]@{
                TestName = "Systeminformationen"
                Result = "FEHLER"
                Details = "Konnte Systeminformationen nicht abrufen$colon $($_.Exception.Message)"
            }
        }
        
        # Festplattenplatz prüfen
        try {
            $disks = Get-WmiObject -Class Win32_LogicalDisk -Filter "DriveType = 3" -ComputerName $ServerName -ErrorAction Stop
            foreach ($disk in $disks) {
                $size = [math]::Round($disk.Size / 1GB, 2)
                $free = [math]::Round($disk.FreeSpace / 1GB, 2)
                $percentFree = [math]::Round(($free / $size) * 100, 2)
                
                $diskResult = if ($percentFree -lt 15) {
                    [PSCustomObject]@{
                        TestName = "Laufwerk $($disk.DeviceID)"
                        Result = "WARNUNG"
                        Details = "Nur $percentFree% freier Speicherplatz ($free GB von $size GB)"
                    }
                } elseif ($percentFree -lt 5) {
                    [PSCustomObject]@{
                        TestName = "Laufwerk $($disk.DeviceID)"
                        Result = "FEHLER"
                        Details = "Kritisch wenig Speicherplatz$colon $percentFree% frei ($free GB von $size GB)"
                    }
                } else {
                    [PSCustomObject]@{
                        TestName = "Laufwerk $($disk.DeviceID)"
                        Result = "OK"
                        Details = "$percentFree% freier Speicherplatz ($free GB von $size GB)"
                    }
                }
                $results += $diskResult
            }
        }
        catch {
            $results += [PSCustomObject]@{
                TestName = "Festplattenplatz"
                Result = "FEHLER"
                Details = "Konnte Festplatteninformationen nicht abrufen$colon $($_.Exception.Message)"
            }
        }
        
        # Netzwerkschnittstellen prüfen
        try {
            $adapters = Get-WmiObject -Class Win32_NetworkAdapter -Filter "NetEnabled='True'" -ComputerName $ServerName -ErrorAction Stop
            $adapterConfigs = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -Filter "IPEnabled='True'" -ComputerName $ServerName -ErrorAction Stop
            
            $activeAdapters = $adapters | Where-Object { $_.NetEnabled }
            
            if ($activeAdapters.Count -gt 0) {
                $results += [PSCustomObject]@{
                    TestName = "Netzwerkadapter"
                    Result = "OK"
                    Details = "$($activeAdapters.Count) aktive Netzwerkadapter gefunden"
                }
                
                # IP-Konfigurationen anzeigen
                foreach ($config in $adapterConfigs) {
                    $adapterName = ($adapters | Where-Object { $_.DeviceID -eq $config.Index }).Name
                    if ($adapterName) {
                        $results += [PSCustomObject]@{
                            TestName = "IP$colon $adapterName"
                            Result = "OK"
                            Details = "$($config.IPAddress -join ', ') | Gateway$colon $($config.DefaultIPGateway -join ', ')"
                        }
                    }
                }
            } else {
                $results += [PSCustomObject]@{
                    TestName = "Netzwerkadapter"
                    Result = "WARNUNG"
                    Details = "Keine aktiven Netzwerkadapter gefunden"
                }
            }
        }
        catch {
            $results += [PSCustomObject]@{
                TestName = "Netzwerkadapter"
                Result = "FEHLER"
                Details = "Konnte Netzwerkadapter nicht prüfen$colon $($_.Exception.Message)"
            }
        }
        
        return $results
    }
    catch {
        Write-Log "Fehler bei grundlegenden Serverchecks für $ServerName`$colon $($_.Exception.Message)" -Level "ERROR"
        return @(
            [PSCustomObject]@{
                TestName = "Server-Basischecks"
                Result = "FEHLER"
                Details = "Ein allgemeiner Fehler ist aufgetreten$colon $($_.Exception.Message)"
            }
        )
    }
}

# Funktion zum Testen der Active Directory-Dienste
function Test-ADServices {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ServerName
    )
    
    $results = @()
    
    try {
        # Prüfen, ob AD-Module verfügbar
        if (-not (Get-Module -Name ActiveDirectory -ErrorAction SilentlyContinue)) {
            try {
                Import-Module ActiveDirectory -ErrorAction Stop
            }
            catch {
                return @(
                    [PSCustomObject]@{
                        TestName = "AD-Module"
                        Result = "FEHLER"
                        Details = "Active Directory-PowerShell-Module nicht verfügbar"
                    }
                )
            }
        }
        
        # Prüfen, ob der Server ein Domänencontroller ist
        try {
            $isDC = $false
            $dc = Get-ADDomainController -Identity $ServerName -ErrorAction SilentlyContinue
            if ($dc) {
                $isDC = $true
                $results += [PSCustomObject]@{
                    TestName = "Domänencontroller"
                    Result = "OK"
                    Details = "Server ist ein Domänencontroller in Domäne $($dc.Domain)"
                }
                
                # Weitere DC-spezifische Tests
                # NTDS-Dienst prüfen
                $ntdsService = Get-Service -ComputerName $ServerName -Name "NTDS" -ErrorAction SilentlyContinue
                if ($ntdsService -and $ntdsService.Status -eq "Running") {
                    $results += [PSCustomObject]@{
                        TestName = "NTDS-Dienst"
                        Result = "OK"
                        Details = "Active Directory-Domain Services läuft"
                    }
                } else {
                    $results += [PSCustomObject]@{
                        TestName = "NTDS-Dienst"
                        Result = "FEHLER"
                        Details = "NTDS-Dienst ist nicht aktiv"
                    }
                }
                
                # Kerberos Key Distribution Center prüfen
                $kdcService = Get-Service -ComputerName $ServerName -Name "Kdc" -ErrorAction SilentlyContinue
                if ($kdcService -and $kdcService.Status -eq "Running") {
                    $results += [PSCustomObject]@{
                        TestName = "Kerberos KDC"
                        Result = "OK"
                        Details = "Kerberos Key Distribution Center läuft"
                    }
                } else {
                    $results += [PSCustomObject]@{
                        TestName = "Kerberos KDC"
                        Result = "FEHLER"
                        Details = "KDC-Dienst ist nicht aktiv"
                    }
                }
                
                # Globaler Katalog prüfen
                if ($dc.IsGlobalCatalog) {
                    $results += [PSCustomObject]@{
                        TestName = "Globaler Katalog"
                        Result = "OK"
                        Details = "Server ist ein Globaler Katalog"
                    }
                } else {
                    $results += [PSCustomObject]@{
                        TestName = "Globaler Katalog"
                        Result = "INFO"
                        Details = "Server ist kein Globaler Katalog"
                    }
                }
                
                # SYSVOL-Pfad prüfen
                try {
                    $sysvol = "\\$ServerName\SYSVOL"
                    $sysvolTest = Test-Path -Path $sysvol -ErrorAction Stop
                    if ($sysvolTest) {
                        $results += [PSCustomObject]@{
                            TestName = "SYSVOL-Freigabe"
                            Result = "OK"
                            Details = "SYSVOL-Freigabe ist erreichbar"
                        }
                    } else {
                        $results += [PSCustomObject]@{
                            TestName = "SYSVOL-Freigabe"
                            Result = "FEHLER"
                            Details = "SYSVOL-Freigabe konnte nicht erreicht werden"
                        }
                    }
                }
                catch {
                    $results += [PSCustomObject]@{
                        TestName = "SYSVOL-Freigabe"
                        Result = "FEHLER"
                        Details = "Fehler beim Prüfen der SYSVOL-Freigabe$colon $($_.Exception.Message)"
                    }
                }
            } else {
                $results += [PSCustomObject]@{
                    TestName = "Domänencontroller"
                    Result = "INFO"
                    Details = "Server ist kein Domänencontroller"
                }
            }
        }
        catch {
            $results += [PSCustomObject]@{
                TestName = "Domänencontroller-Status"
                Result = "FEHLER"
                Details = "Fehler beim Prüfen des DC-Status$colon $($_.Exception.Message)"
            }
        }
        
        return $results
    }
    catch {
        Write-Log "Fehler bei AD-Dienste-Checks für $ServerName`$colon $($_.Exception.Message)" -Level "ERROR"
        return @(
            [PSCustomObject]@{
                TestName = "AD-Dienste"
                Result = "FEHLER"
                Details = "Ein allgemeiner Fehler ist aufgetreten$colon $($_.Exception.Message)"
            }
        )
    }
}

# Funktion zum Testen der AD-Replikation
function Test-ADReplication {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ServerName
    )
    
    $results = @()
    
    try {
        try {
            $dc = Get-ADDomainController -Identity $ServerName -ErrorAction SilentlyContinue
            $isDC = ($null -ne $dc)
        }
        catch {
            return @()  # Kein DC, leere Liste zurückgeben
        }
        
        if (-not $isDC) {
            return @()  # Server ist kein DC, leere Liste zurückgeben
        }
        
        # Replikationspartner ermitteln
        $replPartners = Get-ADReplicationPartnerMetadata -Target $ServerName -ErrorAction SilentlyContinue
        
        if (-not $replPartners) {
            return @(
                [PSCustomObject]@{
                    SourceServer = $ServerName
                    DestinationServer = "N/A"
                    LastReplication = "Unbekannt"
                    Status = "WARNUNG"
                    Details = "Keine Replikationspartner gefunden"
                }
            )
        }
        
        # Replikationsstatus für jeden Partner prüfen
        foreach ($partner in $replPartners) {
            $lastReplTime = $partner.LastReplicationSuccess
            $timeSinceRepl = (Get-Date) - $lastReplTime
            
            $status = "OK"
            $details = "Letzte erfolgreiche Replikation vor $([math]::Round($timeSinceRepl.TotalHours, 2)) Stunden"
            
            # Warnung, wenn Replikation älter als 24 Stunden
            if ($timeSinceRepl.TotalHours -gt 24) {
                $status = "WARNUNG"
                $details = "Replikation ist älter als 24 Stunden ($([math]::Round($timeSinceRepl.TotalHours, 2)) Stunden)"
            }
            
            # Fehler, wenn Replikation älter als 48 Stunden
            if ($timeSinceRepl.TotalHours -gt 48) {
                $status = "FEHLER"
                $details = "Replikation ist älter als 48 Stunden ($([math]::Round($timeSinceRepl.TotalHours, 2)) Stunden)"
            }
            
            $results += [PSCustomObject]@{
                SourceServer = $ServerName
                DestinationServer = $partner.PartnerName
                LastReplication = $lastReplTime.ToString("dd.MM.yyyy HH:mm:ss")
                Status = $status
                Details = $details
            }
        }
        
        # Expliziten Replikationsstatus mit repadmin abfragen, wenn verfügbar
        try {
            $repadminOutput = Invoke-Command -ComputerName $ServerName -ScriptBlock {
                repadmin /replsummary
            } -ErrorAction SilentlyContinue
            
            if ($repadminOutput -match "Fehler") {
                $results += [PSCustomObject]@{
                    SourceServer = $ServerName
                    DestinationServer = "Diverse"
                    LastReplication = "N/A"
                    Status = "WARNUNG"
                    Details = "repadmin meldet Replikationsprobleme"
                }
            }
        }
        catch {
            # Silent catch - repadmin ist optional
        }
        
        return $results
    }
    catch {
        Write-Log "Fehler bei AD-Replikations-Checks für $ServerName`: $($_.Exception.Message)" -Level "ERROR"
        return @(
            [PSCustomObject]@{
                SourceServer = $ServerName
                DestinationServer = "FEHLER"
                LastReplication = "N/A"
                Status = "FEHLER"
                Details = "Ein allgemeiner Fehler ist aufgetreten: $($_.Exception.Message)"
            }
        )
    }
}

# Funktion zum Testen der FSMO-Rollen
function Test-FSMORoles {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ServerName
    )
    
    $results = @()
    
    try {
        # Forest FSMO-Rollen
        try {
            $forest = Get-ADForest -ErrorAction Stop
            
            # Schema Master
            $schemaMaster = $forest.SchemaMaster
            $schemaAvailability = "OK"
            
            # Ping-Test am Schema Master
            if (-not (Test-Connection -ComputerName $schemaMaster -Count 1 -Quiet -ErrorAction SilentlyContinue)) {
                $schemaAvailability = "FEHLER"
            }
            
            $results += [PSCustomObject]@{
                Role = "Schema Master"
                Owner = $schemaMaster
                Availability = $schemaAvailability
                Details = "Forest-weiter Schema Master"
            }
            
            # Domain Naming Master
            $domainNamingMaster = $forest.DomainNamingMaster
            $domainNamingAvailability = "OK"
            
            # Ping-Test am Domain Naming Master
            if (-not (Test-Connection -ComputerName $domainNamingMaster -Count 1 -Quiet -ErrorAction SilentlyContinue)) {
                $domainNamingAvailability = "FEHLER"
            }
            
            $results += [PSCustomObject]@{
                Role = "Domain Naming Master"
                Owner = $domainNamingMaster
                Availability = $domainNamingAvailability
                Details = "Forest-weiter Domain Naming Master"
            }
        }
        catch {
            Write-Log "Fehler beim Abrufen der Forest FSMO-Rollen: $($_.Exception.Message)" -Level "WARNING"
        }
        
        # Domain FSMO-Rollen
        try {
            $domain = Get-ADDomain -ErrorAction Stop
            
            # PDC Emulator
            $pdcEmulator = $domain.PDCEmulator
            $pdcAvailability = "OK"
            
            # Ping-Test am PDC Emulator
            if (-not (Test-Connection -ComputerName $pdcEmulator -Count 1 -Quiet -ErrorAction SilentlyContinue)) {
                $pdcAvailability = "FEHLER"
            }
            
            $results += [PSCustomObject]@{
                Role = "PDC Emulator"
                Owner = $pdcEmulator
                Availability = $pdcAvailability
                Details = "Primary Domain Controller Emulator (Zeitserver)"
            }
            
            # RID Master
            $ridMaster = $domain.RIDMaster
            $ridAvailability = "OK"
            
            # Ping-Test am RID Master
            if (-not (Test-Connection -ComputerName $ridMaster -Count 1 -Quiet -ErrorAction SilentlyContinue)) {
                $ridAvailability = "FEHLER"
            }
            
            $results += [PSCustomObject]@{
                Role = "RID Master"
                Owner = $ridMaster
                Availability = $ridAvailability
                Details = "Relative ID Master (verwaltet SID-Pools)"
            }
            
            # Infrastructure Master
            $infraMaster = $domain.InfrastructureMaster
            $infraAvailability = "OK"
            
            # Ping-Test am Infrastructure Master
            if (-not (Test-Connection -ComputerName $infraMaster -Count 1 -Quiet -ErrorAction SilentlyContinue)) {
                $infraAvailability = "FEHLER"
            }
            
            $results += [PSCustomObject]@{
                Role = "Infrastructure Master"
                Owner = $infraMaster
                Availability = $infraAvailability
                Details = "Infrastructure Master (verwaltet Referenzen zu Objekten)"
            }
        }
        catch {
            Write-Log "Fehler beim Abrufen der Domain FSMO-Rollen: $($_.Exception.Message)" -Level "WARNING"
        }
        
        return $results
    }
    catch {
        Write-Log "Fehler bei FSMO-Rollen-Checks für $ServerName`: $($_.Exception.Message)" -Level "ERROR"
        return @(
            [PSCustomObject]@{
                Role = "FSMO-Fehler"
                Owner = "N/A"
                Availability = "FEHLER"
                Details = "Ein allgemeiner Fehler ist aufgetreten: $($_.Exception.Message)"
            }
        )
    }
}

# Funktion zum Testen des DNS-Dienstes
function Test-DNSService {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ServerName
    )
    
    $results = @()
    
    try {
        # DNS-Dienst prüfen
        $dnsService = Get-Service -ComputerName $ServerName -Name "DNS" -ErrorAction SilentlyContinue
        
        if ($null -eq $dnsService) {
            # DNS ist nicht installiert
            return @()
        }
        
        # DNS-Dienst Status
        $dnsServiceStatus = if ($dnsService.Status -eq "Running") {
            [PSCustomObject]@{
                TestName = "DNS-Dienst"
                Result = "OK"
                Details = "DNS-Dienst läuft"
            }
        } else {
            [PSCustomObject]@{
                TestName = "DNS-Dienst"
                Result = "FEHLER"
                Details = "DNS-Dienst hat Status: $($dnsService.Status)"
            }
        }
        $results += $dnsServiceStatus
        
        # DNS-Server-Konfiguration prüfen
        try {
            if (Get-Command Get-DnsServer -ErrorAction SilentlyContinue) {
                $dnsServer = Get-DnsServer -ComputerName $ServerName -ErrorAction Stop
                
                # Rekursion prüfen - Ternary-Operator durch if-else ersetzt
                $recursionText = if ($dnsServer.ServerSetting.EnableRecursion) { 'aktiviert' } else { 'deaktiviert' }
                $results += [PSCustomObject]@{
                    TestName = "DNS-Rekursion"
                    Result = "OK"
                    Details = "Rekursion ist $recursionText"
                }
                
                # Root-Hints prüfen
                $rootHints = Get-DnsServerRootHint -ComputerName $ServerName -ErrorAction SilentlyContinue
                if ($rootHints -and $rootHints.Count -gt 0) {
                    $results += [PSCustomObject]@{
                        TestName = "DNS-Root-Hints"
                        Result = "OK"
                        Details = "$($rootHints.Count) Root-Hints konfiguriert"
                    }
                } else {
                    $results += [PSCustomObject]@{
                        TestName = "DNS-Root-Hints"
                        Result = "INFO"
                        Details = "Keine Root-Hints konfiguriert (möglicherweise Forwarder-Konfiguration)"
                    }
                }
                
                # Forwarder prüfen
                $forwarders = Get-DnsServerForwarder -ComputerName $ServerName -ErrorAction SilentlyContinue
                if ($forwarders.IPAddress -and $forwarders.IPAddress.Count -gt 0) {
                    $results += [PSCustomObject]@{
                        TestName = "DNS-Forwarder"
                        Result = "OK"
                        Details = "$($forwarders.IPAddress.Count) Forwarder konfiguriert: $($forwarders.IPAddress -join ', ')"
                    }
                } else {
                    $results += [PSCustomObject]@{
                        TestName = "DNS-Forwarder"
                        Result = "INFO"
                        Details = "Keine Forwarder konfiguriert"
                    }
                }
            } else {
                $results += [PSCustomObject]@{
                    TestName = "DNS-PowerShell-Module"
                    Result = "INFO"
                    Details = "DNS PowerShell-Module nicht verfügbar für erweiterte Prüfungen"
                }
            }
        }
        catch {
            $results += [PSCustomObject]@{
                TestName = "DNS-Konfiguration"
                Result = "WARNUNG"
                Details = "Konnte DNS-Konfiguration nicht prüfen: $($_.Exception.Message)"
            }
        }
        
        # DNS-Auflösungscheck
        try {
            $dnsResult = Resolve-DnsName -Name "www.microsoft.com" -Server $ServerName -ErrorAction Stop
            if ($dnsResult) {
                $results += [PSCustomObject]@{
                    TestName = "DNS-Auflösung"
                    Result = "OK"
                    Details = "DNS-Server kann externe Domains auflösen"
                }
            }
        }
        catch {
            $results += [PSCustomObject]@{
                TestName = "DNS-Auflösung"
                Result = "FEHLER"
                Details = "DNS-Server kann externe Domains nicht auflösen: $($_.Exception.Message)"
            }
        }
        
        return $results
    }
    catch {
        Write-Log "Fehler bei DNS-Dienst-Checks für $ServerName`: $($_.Exception.Message)" -Level "ERROR"
        return @(
            [PSCustomObject]@{
                TestName = "DNS-Dienst"
                Result = "FEHLER"
                Details = "Ein allgemeiner Fehler ist aufgetreten: $($_.Exception.Message)"
            }
        )
    }
}

# Funktion zum Testen der DNS-Zonen
function Test-DNSZones {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ServerName
    )
    
    $results = @()
    
    try {
        # DNS-Dienst prüfen
        $dnsService = Get-Service -ComputerName $ServerName -Name "DNS" -ErrorAction SilentlyContinue
        
        if ($null -eq $dnsService -or $dnsService.Status -ne "Running") {
            # DNS ist nicht installiert oder nicht aktiv
            return @()
        }
        
        # DNS-Zonen abrufen
        try {
            if (Get-Command Get-DnsServerZone -ErrorAction SilentlyContinue) {
                $zones = Get-DnsServerZone -ComputerName $ServerName -ErrorAction Stop
                
                foreach ($zone in $zones) {
                    $zoneStatus = if ($zone.ZoneType -eq "Primary" -and $zone.DynamicUpdate -eq "None") {
                        "WARNUNG"
                    } else {
                        "OK"
                    }
                    
                    $replicationInfo = switch ($zone.ReplicationScope) {
                        "Domain" { "AD-integriert (Domäne)" }
                        "Forest" { "AD-integriert (Forest)" }
                        "Legacy" { "Primäre Datei-basierte Zone" }
                        "Custom" { "AD-integriert (Benutzerdefiniert)" }
                        default { "Unbekannt: $($zone.ReplicationScope)" }
                    }
                    
                    $results += [PSCustomObject]@{
                        ZoneName = $zone.ZoneName
                        ZoneType = $zone.ZoneType
                        Status = $zoneStatus
                        Replication = $replicationInfo
                        Details = "Dynamische Updates: $($zone.DynamicUpdate)"
                    }
                }
            }
        }
        catch {
            $results += [PSCustomObject]@{
                ZoneName = "Fehler"
                ZoneType = "N/A"
                Status = "FEHLER"
                Replication = "N/A"
                Details = "Konnte DNS-Zonen nicht abrufen: $($_.Exception.Message)"
            }
        }
        
        return $results
    }
    catch {
        Write-Log "Fehler bei DNS-Zonen-Checks für $ServerName`: $($_.Exception.Message)" -Level "ERROR"
        return @(
            [PSCustomObject]@{
                ZoneName = "DNS-Zonen-Fehler"
                ZoneType = "N/A"
                Status = "FEHLER"
                Replication = "N/A"
                Details = "Ein allgemeiner Fehler ist aufgetreten: $($_.Exception.Message)"
            }
        )
    }
}

# Funktion zum Testen der DNS-Auflösung
function Test-DNSResolution {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ServerName
    )
    
    $results = @()
    
    try {
        # DNS-Dienst prüfen
        $dnsService = Get-Service -ComputerName $ServerName -Name "DNS" -ErrorAction SilentlyContinue
        
        if ($null -eq $dnsService -or $dnsService.Status -ne "Running") {
            # DNS ist nicht installiert oder nicht aktiv
            return @()
        }
        
        # Liste der zu prüfenden Hostnamen
        $hostnames = @(
            "www.microsoft.com",
            "www.google.com",
            $env:COMPUTERNAME,
            $ServerName
        )
        
        # Domain ermitteln und prüfen
        try {
            $domain = (Get-WmiObject Win32_ComputerSystem).Domain
            if ($domain -ne "WORKGROUP") {
                $hostnames += $domain
            }
        }
        catch {
            Write-Log "Warnung: Konnte Domäne für DNS-Test nicht ermitteln: $($_.Exception.Message)" -Level "WARNING"
        }
        
        # DNS-Auflösung für jeden Hostname testen
        foreach ($hostname in $hostnames) {
            try {
                $startTime = Measure-Command {
                    $resolution = Resolve-DnsName -Name $hostname -Server $ServerName -ErrorAction Stop
                }
                
                $ip = ($resolution | Where-Object { $_.Type -in @("A", "AAAA") } | Select-Object -First 1).IPAddress
                
                $results += [PSCustomObject]@{
                    Hostname = $hostname
                    ResolvedIP = $ip
                    Status = "OK"
                    ResponseTime = [math]::Round($startTime.TotalMilliseconds, 2)
                    Details = "Auflösung erfolgreich"
                }
            }
            catch {
                $results += [PSCustomObject]@{
                    Hostname = $hostname
                    ResolvedIP = "Keine"
                    Status = "FEHLER"
                    ResponseTime = 0
                    Details = "Konnte nicht aufgelöst werden: $($_.Exception.Message)"
                }
            }
        }
        
        return $results
    }
    catch {
        Write-Log "Fehler bei DNS-Auflösungs-Checks für $ServerName`: $($_.Exception.Message)" -Level "ERROR"
        return @(
            [PSCustomObject]@{
                Hostname = "DNS-Auflösungs-Fehler"
                ResolvedIP = "N/A"
                Status = "FEHLER"
                ResponseTime = 0
                Details = "Ein allgemeiner Fehler ist aufgetreten: $($_.Exception.Message)"
            }
        )
    }
}

# Funktion zum Testen des DHCP-Dienstes
function Test-DHCPService {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ServerName
    )
    
    $results = @()
    
    try {
        # DHCP-Dienst prüfen
        $dhcpService = Get-Service -ComputerName $ServerName -Name "DHCPServer" -ErrorAction SilentlyContinue
        
        if ($null -eq $dhcpService) {
            # DHCP ist nicht installiert
            return @()
        }
        
        # DHCP-Dienst Status
        $dhcpServiceStatus = if ($dhcpService.Status -eq "Running") {
            [PSCustomObject]@{
                TestName = "DHCP-Dienst"
                Result = "OK"
                Details = "DHCP-Dienst läuft"
            }
        } else {
            [PSCustomObject]@{
                TestName = "DHCP-Dienst"
                Result = "FEHLER"
                Details = "DHCP-Dienst hat Status: $($dhcpService.Status)"
            }
        }
        $results += $dhcpServiceStatus
        
        # DHCP-Server-Konfiguration prüfen
        try {
            if (Get-Command Get-DhcpServerSetting -ErrorAction SilentlyContinue) {
                $dhcpSettings = Get-DhcpServerSetting -ComputerName $ServerName -ErrorAction Stop
                
                # Autorisierung prüfen
                $dhcpAuth = if ($dhcpSettings.IsAuthorized) {
                    [PSCustomObject]@{
                        TestName = "DHCP-Autorisierung"
                        Result = "OK"
                        Details = "DHCP-Server ist in der Domäne autorisiert"
                    }
                } else {
                    [PSCustomObject]@{
                        TestName = "DHCP-Autorisierung"
                        Result = "FEHLER"
                        Details = "DHCP-Server ist in der Domäne nicht autorisiert"
                    }
                }
                $results += $dhcpAuth
                
                # Konflikterkennungs-Prüfung
                $results += [PSCustomObject]@{
                    TestName = "DHCP-Konflikterkennung"
                    Result = "INFO"
                    Details = "Konflikterkennung ist auf $($dhcpSettings.ConflictDetectionAttempts) Versuche eingestellt"
                }
            }
        }
        catch {
            $results += [PSCustomObject]@{
                TestName = "DHCP-Konfiguration"
                Result = "WARNUNG"
                Details = "Konnte DHCP-Konfiguration nicht prüfen: $($_.Exception.Message)"
            }
        }
        
        # DHCP-Leistungsindikatoren
        try {
            $dhcpLeasesTotal = 0
            $dhcpScopesTotal = 0
            
            if (Get-Command Get-DhcpServerv4Scope -ErrorAction SilentlyContinue) {
                $scopes = Get-DhcpServerv4Scope -ComputerName $ServerName -ErrorAction SilentlyContinue
                $dhcpScopesTotal = ($scopes | Measure-Object).Count
                
                foreach ($scope in $scopes) {
                    $stats = Get-DhcpServerv4ScopeStatistics -ComputerName $ServerName -ScopeId $scope.ScopeId -ErrorAction SilentlyContinue
                    if ($stats) {
                        $dhcpLeasesTotal += $stats.AddressesInUse
                    }
                }
                
                $results += [PSCustomObject]@{
                    TestName = "DHCP-Nutzung"
                    Result = "OK"
                    Details = "$dhcpLeasesTotal aktive Leases in $dhcpScopesTotal Bereichen"
                }
            }
        }
        catch {
            $results += [PSCustomObject]@{
                TestName = "DHCP-Statistiken"
                Result = "WARNUNG"
                Details = "Konnte DHCP-Statistiken nicht abrufen: $($_.Exception.Message)"
            }
        }
        
        return $results
    }
    catch {
        Write-Log "Fehler bei DHCP-Dienst-Checks für $ServerName`: $($_.Exception.Message)" -Level "ERROR"
        return @(
            [PSCustomObject]@{
                TestName = "DHCP-Dienst"
                Result = "FEHLER"
                Details = "Ein allgemeiner Fehler ist aufgetreten: $($_.Exception.Message)"
            }
        )
    }
}

# Funktion zum Testen der DHCP-Bereiche
function Test-DHCPScopes {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ServerName
    )
    
    $results = @()
    
    try {
        # DHCP-Dienst prüfen
        $dhcpService = Get-Service -ComputerName $ServerName -Name "DHCPServer" -ErrorAction SilentlyContinue
        
        if ($null -eq $dhcpService -or $dhcpService.Status -ne "Running") {
            # DHCP ist nicht installiert oder nicht aktiv
            return @()
        }
        
        # DHCP-Bereiche abrufen
        try {
            if (Get-Command Get-DhcpServerv4Scope -ErrorAction SilentlyContinue) {
                $scopes = Get-DhcpServerv4Scope -ComputerName $ServerName -ErrorAction Stop
                
                foreach ($scope in $scopes) {
                    $stats = Get-DhcpServerv4ScopeStatistics -ComputerName $ServerName -ScopeId $scope.ScopeId -ErrorAction SilentlyContinue
                    
                    $usage = if ($stats) {
                        [math]::Round(($stats.AddressesInUse / $stats.AddressesTotal) * 100, 2)
                    } else {
                        "Unbekannt"
                    }
                    
                    $scopeStatus = "OK"
                    $details = "Bereich ist aktiv"
                    
                    # Warnung, wenn Bereich zu voll ist
                    if ($usage -is [double] -and $usage -gt 85) {
                        $scopeStatus = "WARNUNG"
                        $details = "Belegung über 85% - Erweiterung empfohlen"
                    }
                    
                    # Warnung, wenn Bereich deaktiviert ist
                    if (-not $scope.State -or $scope.State -ne "Active") {
                        $scopeStatus = "WARNUNG"
                        $details = "Bereich ist nicht aktiv (Status: $($scope.State))"
                    }
                    
                    $results += [PSCustomObject]@{
                        ScopeId = $scope.ScopeId
                        ScopeName = $scope.Name
                        Status = $scopeStatus
                        Usage = "$usage%"
                        Leases = if ($stats) { $stats.AddressesInUse } else { "Unbekannt" }
                        Details = $details
                    }
                }
            }
        }
        catch {
            $results += [PSCustomObject]@{
                ScopeId = "Fehler"
                ScopeName = "N/A"
                Status = "FEHLER"
                Usage = "N/A"
                Leases = "N/A"
                Details = "Konnte DHCP-Bereiche nicht abrufen: $($_.Exception.Message)"
            }
        }
        
        return $results
    }
    catch {
        Write-Log "Fehler bei DHCP-Bereiche-Checks für $ServerName`: $($_.Exception.Message)" -Level "ERROR"
        return @(
            [PSCustomObject]@{
                ScopeId = "DHCP-Bereiche-Fehler"
                ScopeName = "N/A"
                Status = "FEHLER"
                Usage = "N/A"
                Leases = "N/A"
                Details = "Ein allgemeiner Fehler ist aufgetreten: $($_.Exception.Message)"
            }
        )
    }
}

# Funktion zum Testen von DHCP-Failover
function Test-DHCPFailover {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ServerName
    )
    
    $results = @()
    
    try {
        # DHCP-Dienst prüfen
        $dhcpService = Get-Service -ComputerName $ServerName -Name "DHCPServer" -ErrorAction SilentlyContinue
        
        if ($null -eq $dhcpService -or $dhcpService.Status -ne "Running") {
            # DHCP ist nicht installiert oder nicht aktiv
            return @()
        }
        
        # DHCP-Failover-Beziehungen abrufen
        try {
            if (Get-Command Get-DhcpServerv4Failover -ErrorAction SilentlyContinue) {
                $failovers = Get-DhcpServerv4Failover -ComputerName $ServerName -ErrorAction SilentlyContinue
                
                if (-not $failovers -or $failovers.Count -eq 0) {
                    return @()  # Keine Failover-Beziehungen
                }
                
                foreach ($failover in $failovers) {
                    $failoverStatus = "OK"
                    $details = "Failover-Beziehung ist aktiv"
                    
                    # Partner-Server prüfen
                    $partnerReachable = Test-Connection -ComputerName $failover.PartnerServer -Count 1 -Quiet -ErrorAction SilentlyContinue
                    if (-not $partnerReachable) {
                        $failoverStatus = "WARNUNG"
                        $details = "Failover-Partner ist nicht erreichbar"
                    }
                    
                    $mode = switch ($failover.Mode) {
                        "LoadBalance" { "Lastverteilung ($($failover.LoadBalancePercent)%)" }
                        "HotStandby" { "Heiße Bereitschaft" }
                        default { $failover.Mode }
                    }
                    
                    foreach ($scopeId in $failover.ScopeId) {
                        $results += [PSCustomObject]@{
                            ScopeId = $scopeId
                            PartnerServer = $failover.PartnerServer
                            Mode = $mode
                            Status = $failoverStatus
                            Details = $details
                        }
                    }
                }
            }
        }
        catch {
            $results += [PSCustomObject]@{
                ScopeId = "Fehler"
                PartnerServer = "N/A"
                Mode = "N/A"
                Status = "FEHLER"
                Details = "Konnte DHCP-Failover nicht prüfen: $($_.Exception.Message)"
            }
        }
        
        return $results
    }
    catch {
        Write-Log "Fehler bei DHCP-Failover-Checks für $ServerName`: $($_.Exception.Message)" -Level "ERROR"
        return @(
            [PSCustomObject]@{
                ScopeId = "DHCP-Failover-Fehler"
                PartnerServer = "N/A"
                Mode = "N/A"
                Status = "FEHLER"
                Details = "Ein allgemeiner Fehler ist aufgetreten: $($_.Exception.Message)"
            }
        )
    }
}

# Funktion zum Testen kritischer Dienste
function Test-CriticalServices {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ServerName
    )
    
    $results = @()
    
    try {
        # Liste kritischer Dienste je nach Serverrolle
        $criticalServices = @(
            # Basis-Serverdienste
            @{Name = "LanmanServer"; DisplayName = "Server"; Critical = $true},
            @{Name = "LanmanWorkstation"; DisplayName = "Workstation"; Critical = $true},
            @{Name = "W32Time"; DisplayName = "Windows Time"; Critical = $true},
            @{Name = "Dnscache"; DisplayName = "DNS Client"; Critical = $true},
            @{Name = "RpcSs"; DisplayName = "Remote Procedure Call (RPC)"; Critical = $true},
            @{Name = "EventLog"; DisplayName = "Windows Event Log"; Critical = $true}
        )
        
        # Domänencontroller-spezifische Dienste
        try {
            $isDC = Get-ADDomainController -Identity $ServerName -ErrorAction SilentlyContinue
            if ($isDC) {
                $criticalServices += @(
                    @{Name = "NTDS"; DisplayName = "Active Directory Domain Services"; Critical = $true},
                    @{Name = "Kdc"; DisplayName = "Kerberos Key Distribution Center"; Critical = $true},
                    @{Name = "DFSR"; DisplayName = "DFS Replication"; Critical = $false},
                    @{Name = "Netlogon"; DisplayName = "Netlogon"; Critical = $true}
                )
            }
        }
        catch {
            # Fehlerbehandlung - Server ist möglicherweise kein DC oder AD-Module fehlen
        }
        
        # DNS-/DHCP-Server-Dienste
        $dnsService = Get-Service -ComputerName $ServerName -Name "DNS" -ErrorAction SilentlyContinue
        if ($dnsService) {
            $criticalServices += @{Name = "DNS"; DisplayName = "DNS Server"; Critical = $true}
        }
        
        $dhcpService = Get-Service -ComputerName $ServerName -Name "DHCPServer" -ErrorAction SilentlyContinue
        if ($dhcpService) {
            $criticalServices += @{Name = "DHCPServer"; DisplayName = "DHCP Server"; Critical = $true}
        }
        
        # Dienststatus abrufen
        foreach ($service in $criticalServices) {
            try {
                $svc = Get-Service -ComputerName $ServerName -Name $service.Name -ErrorAction SilentlyContinue
                if ($svc) {
                    $status = $svc.Status
                    $startType = (Get-WmiObject -Class Win32_Service -Filter "Name='$($service.Name)'" -ComputerName $ServerName -ErrorAction SilentlyContinue).StartMode
                    
                    $details = if ($status -ne "Running" -and $service.Critical) {
                        "Kritischer Dienst ist nicht aktiv!"
                    } elseif ($status -ne "Running") {
                        "Dienst ist nicht aktiv"
                    } else {
                        "Dienst läuft normal"
                    }
                    
                    $results += [PSCustomObject]@{
                        ServiceName = $service.Name
                        DisplayName = $service.DisplayName
                        Status = $status
                        StartType = $startType
                        Details = $details
                    }
                }
            }
            catch {
                $results += [PSCustomObject]@{
                    ServiceName = $service.Name
                    DisplayName = $service.DisplayName
                    Status = "FEHLER"
                    StartType = "Unbekannt"
                    Details = "Konnte Dienststatus nicht abrufen: $($_.Exception.Message)"
                }
            }
        }
        
        return $results
    }
    catch {
        Write-Log "Fehler bei kritischen Dienste-Checks für $ServerName`: $($_.Exception.Message)" -Level "ERROR"
        return @(
            [PSCustomObject]@{
                ServiceName = "Dienste-Fehler"
                DisplayName = "N/A"
                Status = "FEHLER"
                StartType = "N/A"
                Details = "Ein allgemeiner Fehler ist aufgetreten: $($_.Exception.Message)"
            }
        )
    }
}

# Funktion zum Testen der Serverperformance
function Test-ServerPerformance {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ServerName
    )
    
    $results = @()
    
    try {
        # Performance-Metriken abfragen
        
        # CPU-Auslastung
        try {
            $cpuLoad = Get-WmiObject -ComputerName $ServerName -Class Win32_Processor -ErrorAction Stop | 
                       Measure-Object -Property LoadPercentage -Average | 
                       Select-Object -ExpandProperty Average
            
            $cpuStatus = if ($cpuLoad -gt 85) {
                "WARNUNG"
            } else {
                "OK"
            }
            
            $results += [PSCustomObject]@{
                MetricName = "CPU-Auslastung"
                Value = $cpuLoad
                Unit = "%"
                Status = $cpuStatus
                Details = if ($cpuLoad -gt 85) { "CPU-Auslastung ist hoch" } else { "CPU-Auslastung ist normal" }
            }
        }
        catch {
            $results += [PSCustomObject]@{
                MetricName = "CPU-Auslastung"
                Value = "N/A"
                Unit = "%"
                Status = "FEHLER"
                Details = "Konnte CPU-Auslastung nicht abrufen: $($_.Exception.Message)"
            }
        }
        
        # Speicherauslastung
        try {
            $osMemory = Get-WmiObject -ComputerName $ServerName -Class Win32_OperatingSystem -ErrorAction Stop
            $totalMemory = [math]::Round($osMemory.TotalVisibleMemorySize / 1MB, 2)
            $freeMemory = [math]::Round($osMemory.FreePhysicalMemory / 1MB, 2)
            $usedMemory = $totalMemory - $freeMemory
            $memoryUsagePercent = [math]::Round(($usedMemory / $totalMemory) * 100, 2)
            
            $memoryStatus = if ($memoryUsagePercent -gt 85) {
                "WARNUNG"
            } else {
                "OK"
            }
            
            $results += [PSCustomObject]@{
                MetricName = "Arbeitsspeichernutzung"
                Value = $memoryUsagePercent
                Unit = "%"
                Status = $memoryStatus
                Details = "Verwendet $usedMemory GB von $totalMemory GB"
            }
        }
        catch {
            $results += [PSCustomObject]@{
                MetricName = "Arbeitsspeichernutzung"
                Value = "N/A"
                Unit = "%"
                Status = "FEHLER"
                Details = "Konnte Arbeitsspeichernutzung nicht abrufen: $($_.Exception.Message)"
            }
        }
        
        # Festplattenbelegung
        try {
            $disks = Get-WmiObject -ComputerName $ServerName -Class Win32_LogicalDisk -Filter "DriveType = 3" -ErrorAction Stop
            
            foreach ($disk in $disks) {
                $freeSpacePercent = [math]::Round(($disk.FreeSpace / $disk.Size) * 100, 2)
                $totalSizeGB = [math]::Round($disk.Size / 1GB, 2)
                $freeSpaceGB = [math]::Round($disk.FreeSpace / 1GB, 2)
                
                $diskStatus = if ($freeSpacePercent -lt 15) {
                    "WARNUNG"
                } elseif ($freeSpacePercent -lt 5) {
                    "FEHLER"
                } else {
                    "OK"
                }
                
                $results += [PSCustomObject]@{
                    MetricName = "Festplatte $($disk.DeviceID)"
                    Value = $freeSpacePercent
                    Unit = "% frei"
                    Status = $diskStatus
                    Details = "$freeSpaceGB GB von $totalSizeGB GB verfügbar"
                }
            }
        }
        catch {
            $results += [PSCustomObject]@{
                MetricName = "Festplattenbelegung"
                Value = "N/A"
                Unit = "%"
                Status = "FEHLER"
                Details = "Konnte Festplattenbelegung nicht abrufen: $($_.Exception.Message)"
            }
        }
        
        # Netzwerknutzung (vereinfacht, da genaue Netzwerknutzung schwieriger in einer Momentaufnahme zu bekommen ist)
        try {
            $networkAdapters = Get-WmiObject -ComputerName $ServerName -Class Win32_NetworkAdapter -Filter "NetEnabled='True'" -ErrorAction Stop
            
            $results += [PSCustomObject]@{
                MetricName = "Netzwerkadapter"
                Value = $networkAdapters.Count
                Unit = "aktiv"
                Status = "OK"
                Details = "$($networkAdapters.Count) aktive Netzwerkverbindungen"
            }
        }
        catch {
            $results += [PSCustomObject]@{
                MetricName = "Netzwerkadapter"
                Value = "N/A"
                Unit = ""
                Status = "FEHLER"
                Details = "Konnte Netzwerkadapter-Status nicht abrufen: $($_.Exception.Message)"
            }
        }
        
        # Systemuptime
        try {
            $os = Get-WmiObject -ComputerName $ServerName -Class Win32_OperatingSystem -ErrorAction Stop
            $bootTime = [Management.ManagementDateTimeConverter]::ToDateTime($os.LastBootUpTime)
            $uptime = (Get-Date) - $bootTime
            $uptimeDisplay = "{0}d {1}h {2}m" -f $uptime.Days, $uptime.Hours, $uptime.Minutes
            
            $uptimeStatus = if ($uptime.Days -gt 60) {
                "WARNUNG"
            } else {
                "OK"
            }
            
            $results += [PSCustomObject]@{
                MetricName = "System-Uptime"
                Value = $uptime.Days
                Unit = "Tage"
                Status = $uptimeStatus
                Details = "System läuft seit $uptimeDisplay"
            }
        }
        catch {
            $results += [PSCustomObject]@{
                MetricName = "System-Uptime"
                Value = "N/A"
                Unit = "Tage"
                Status = "FEHLER"
                Details = "Konnte System-Uptime nicht abrufen: $($_.Exception.Message)"
            }
        }
        
        return $results
    }
    catch {
        Write-Log "Fehler bei Performance-Checks für $ServerName`: $($_.Exception.Message)" -Level "ERROR"
        return @(
            [PSCustomObject]@{
                MetricName = "Performance-Fehler"
                Value = "N/A"
                Unit = "N/A"
                Status = "FEHLER"
                Details = "Ein allgemeiner Fehler ist aufgetreten: $($_.Exception.Message)"
            }
        )
    }
}

# Funktion zum Testen der Ereignisprotokolle
function Test-EventLogs {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ServerName
    )
    
    $results = @()
    
    try {
        # Zu prüfende Ereignisprotokolle
        $logNames = @("System", "Application", "Security", "Directory Service", "DNS Server", "DFS Replication")
        
        # Zeitrahmen für die Prüfung (letzte 24 Stunden)
        $startTime = (Get-Date).AddHours(-24)
        
        foreach ($logName in $logNames) {
            try {
                # Protokoll prüfen, ob es existiert
                $log = Get-EventLog -List | Where-Object { $_.LogDisplayName -eq $logName -or $_.Log -eq $logName } -ErrorAction SilentlyContinue
                
                if (-not $log) {
                    continue  # Protokoll nicht gefunden, überspringen
                }
                
                # Fehler und Warnungen im Protokoll zählen
                $errorEvents = Get-EventLog -ComputerName $ServerName -LogName $log.Log -EntryType Error -After $startTime -ErrorAction SilentlyContinue
                $warningEvents = Get-EventLog -ComputerName $ServerName -LogName $log.Log -EntryType Warning -After $startTime -ErrorAction SilentlyContinue
                
                $errorCount = ($errorEvents | Measure-Object).Count
                $warningCount = ($warningEvents | Measure-Object).Count
                
                # Schweregrad bestimmen
                $severity = if ($errorCount -gt 0) {
                    "Error"
                } elseif ($warningCount -gt 0) {
                    "Warning"
                } else {
                    "Information"
                }
                
                # Neuestes Ereignis finden
                $latestEvent = if ($errorCount -gt 0) {
                    $errorEvents | Sort-Object TimeGenerated -Descending | Select-Object -First 1
                } elseif ($warningCount -gt 0) {
                    $warningEvents | Sort-Object TimeGenerated -Descending | Select-Object -First 1
                } else {
                    $null
                }
                
                $latestEventText = if ($latestEvent) {
                    "$($latestEvent.TimeGenerated.ToString("dd.MM.yyyy HH:mm:ss")) - $($latestEvent.Message.Split([Environment]::NewLine)[0])"
                } else {
                    "Keine relevanten Ereignisse"
                }
                
                # Gekürzte Version für Details
                if ($latestEventText.Length -gt 100) {
                    $latestEventText = $latestEventText.Substring(0, 100) + "..."
                }
                
                $results += [PSCustomObject]@{
                    LogName = $logName
                    Severity = $severity
                    Count = $errorCount + $warningCount
                    LatestEvent = $latestEventText
                    Details = "Fehler: $errorCount, Warnungen: $warningCount in den letzten 24h"
                }
            }
            catch {
                $results += [PSCustomObject]@{
                    LogName = $logName
                    Severity = "Error"
                    Count = 0
                    LatestEvent = "Fehler beim Zugriff"
                    Details = "Konnte Ereignisprotokoll nicht abfragen: $($_.Exception.Message)"
                }
            }
        }
        
        return $results
    }
    catch {
        Write-Log "Fehler bei Ereignisprotokoll-Checks für $ServerName`: $($_.Exception.Message)" -Level "ERROR"
        return @(
            [PSCustomObject]@{
                LogName = "Ereignisprotokollfehler"
                Severity = "Error"
                Count = 0
                LatestEvent = "N/A"
                Details = "Ein allgemeiner Fehler ist aufgetreten: $($_.Exception.Message)"
            }
        )
    }
}

# Funktion zum Testen der Gruppenrichtlinien-Replikation
function Test-GroupPolicy {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ServerName
    )
    
    $results = @()
    
    try {
        # Prüfen, ob der Server ein Domänencontroller ist
        $isDC = $false
        try {
            $dc = Get-ADDomainController -Identity $ServerName -ErrorAction SilentlyContinue
            $isDC = ($null -ne $dc)
        }
        catch {
            return @()  # Kein DC, leere Liste zurückgeben
        }
        
        if (-not $isDC) {
            return @()  # Server ist kein DC, leere Liste zurückgeben
        }
        
        # GPOs in der Domäne abrufen und prüfen
        try {
            if (Get-Command Get-GPO -ErrorAction SilentlyContinue) {
                $domain = Get-ADDomain -ErrorAction Stop
                $gpos = Get-GPO -All -Domain $domain.DNSRoot -ErrorAction Stop
                
                foreach ($gpo in $gpos) {
                    $gpoStatus = "OK"
                    $details = "GPO-Version: User $($gpo.User.DSVersion), Computer $($gpo.Computer.DSVersion)"
                    
                    # Bei Bedarf hier weitere GPO-Prüfungen einbauen
                    
                    $results += [PSCustomObject]@{
                        GPOName = $gpo.DisplayName
                        Version = "$($gpo.User.DSVersion)/$($gpo.Computer.DSVersion)"
                        ReplicationStatus = $gpoStatus
                        LastChange = $gpo.ModificationTime.ToString("dd.MM.yyyy HH:mm:ss")
                        Details = $details
                    }
                }
            } else {
                $results += [PSCustomObject]@{
                    GPOName = "GroupPolicy-Module nicht verfügbar"
                    Version = "N/A"
                    ReplicationStatus = "INFO"
                    LastChange = "N/A"
                    Details = "GroupPolicy PowerShell-Module nicht geladen"
                }
            }
        }
        catch {
            $results += [PSCustomObject]@{
                GPOName = "GPO-Abfragefehler"
                Version = "N/A"
                ReplicationStatus = "FEHLER"
                LastChange = "N/A"
                Details = "Fehler beim Abrufen von GPOs: $($_.Exception.Message)"
            }
        }
        
        return $results
    }
    catch {
        Write-Log "Fehler bei Gruppenrichtlinien-Checks für $ServerName`: $($_.Exception.Message)" -Level "ERROR"
        return @(
            [PSCustomObject]@{
                GPOName = "Gruppenrichtlinien-Fehler"
                Version = "N/A"
                ReplicationStatus = "FEHLER"
                LastChange = "N/A"
                Details = "Ein allgemeiner Fehler ist aufgetreten: $($_.Exception.Message)"
            }
        )
    }
}

# Funktion zum Testen der SYSVOL-Replikation
function Test-SYSVOL {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ServerName
    )
    
    $results = @()
    
    try {
        # Prüfen, ob der Server ein Domänencontroller ist
        $isDC = $false
        try {
            $dc = Get-ADDomainController -Identity $ServerName -ErrorAction SilentlyContinue
            $isDC = ($null -ne $dc)
        }
        catch {
            return @()  # Kein DC, leere Liste zurückgeben
        }
        
        if (-not $isDC) {
            return @()  # Server ist kein DC, leere Liste zurückgeben
        }
        
        # SYSVOL-Freigabe prüfen
        try {
            $sysvolPath = "\\$ServerName\SYSVOL"
            $sysvolAccessible = Test-Path -Path $sysvolPath -ErrorAction Stop
            
            if ($sysvolAccessible) {
                $results += [PSCustomObject]@{
                    Server = $ServerName
                    ReplicationType = "SYSVOL-Freigabe"
                    Status = "OK"
                    Details = "SYSVOL-Freigabe ist zugänglich"
                }
            } else {
                $results += [PSCustomObject]@{
                    Server = $ServerName
                    ReplicationType = "SYSVOL-Freigabe"
                    Status = "FEHLER"
                    Details = "SYSVOL-Freigabe ist nicht zugänglich"
                }
            }
        }
        catch {
            $results += [PSCustomObject]@{
                Server = $ServerName
                ReplicationType = "SYSVOL-Freigabe"
                Status = "FEHLER"
                Details = "Fehler beim Zugriff auf SYSVOL: $($_.Exception.Message)"
            }
        }
        
        # DFSR-Status prüfen (ab Windows Server 2008)
        try {
            $dfsrService = Get-Service -ComputerName $ServerName -Name DFSR -ErrorAction SilentlyContinue
            
            if ($dfsrService) {
                $dfsrStatus = if ($dfsrService.Status -eq "Running") {
                    "OK"
                } else {
                    "FEHLER"
                }
                
                $results += [PSCustomObject]@{
                    Server = $ServerName
                    ReplicationType = "DFS-Replication"
                    Status = $dfsrStatus
                    Details = "DFSR-Dienststatus: $($dfsrService.Status)"
                }
                
                # DFSR-Event-Logs nach Fehlern durchsuchen
                try {
                    $dfsrEvents = Get-EventLog -ComputerName $ServerName -LogName "DFS Replication" -EntryType Error -Newest 5 -ErrorAction SilentlyContinue
                    
                    if ($dfsrEvents -and $dfsrEvents.Count -gt 0) {
                        $results += [PSCustomObject]@{
                            Server = $ServerName
                            ReplicationType = "DFS-Replication Ereignisse"
                            Status = "WARNUNG"
                            Details = "$($dfsrEvents.Count) DFSR-Fehlerereignisse gefunden"
                        }
                    }
                }
                catch {
                    # DFSR-Ereignislog möglicherweise nicht verfügbar
                }
            } else {
                # FRS prüfen (Windows Server 2003/2008)
                $frsService = Get-Service -ComputerName $ServerName -Name NtFrs -ErrorAction SilentlyContinue
                
                if ($frsService) {
                    $frsStatus = if ($frsService.Status -eq "Running") {
                        "OK"
                    } else {
                        "FEHLER"
                    }
                    
                    $results += [PSCustomObject]@{
                        Server = $ServerName
                        ReplicationType = "File Replication Service (FRS)"
                        Status = $frsStatus
                        Details = "FRS-Dienststatus: $($frsService.Status) - Veraltete Technologie"
                    }
                }
            }
        }
        catch {
            $results += [PSCustomObject]@{
                Server = $ServerName
                ReplicationType = "Replikationsdienst"
                Status = "FEHLER"
                Details = "Fehler beim Prüfen der Replikationsdienste: $($_.Exception.Message)"
            }
        }
        
        return $results
    }
    catch {
        Write-Log "Fehler bei SYSVOL-Checks für $ServerName`: $($_.Exception.Message)" -Level "ERROR"
        return @(
            [PSCustomObject]@{
                Server = $ServerName
                ReplicationType = "SYSVOL-Fehler"
                Status = "FEHLER"
                Details = "Ein allgemeiner Fehler ist aufgetreten: $($_.Exception.Message)"
            }
        )
    }
}

# Funktion zum Exportieren der Validierungsergebnisse als HTML-Bericht
function Export-ValidationToHTML {
    param (
        [Parameter(Mandatory = $true)]
        [object]$ValidationResults,
        [string]$FilePath
    )
    
    try {
        $currentYear = (Get-Date -Format "yyyy")
        
        $htmlHeader = @"
<!DOCTYPE html>
<html>
<head>
    <title>Windows Server 2012 Migration - Validierungsbericht</title>
    <style>
        body { 
            font-family: 'Segoe UI', Arial, sans-serif; 
            margin: 20px; 
            color: #333;
            line-height: 1.5;
        }
        h1 { 
            color: #007ACC; 
            border-bottom: 1px solid #007ACC;
            padding-bottom: 10px;
        }
        h2 { 
            color: #007ACC; 
            margin-top: 20px;
            border-bottom: 1px solid #ccc;
            padding-bottom: 5px;
        }
        table { 
            border-collapse: collapse; 
            width: 100%; 
            margin-bottom: 20px;
        }
        th { 
            background-color: #007ACC; 
            color: white; 
            text-align: left; 
            padding: 8px; 
        }
        td { 
            padding: 8px; 
            border-bottom: 1px solid #ddd; 
        }
        tr:nth-child(even) { 
            background-color: #f9f9f9; 
        }
        tr:hover { 
            background-color: #f1f1f1; 
        }
        .summary { 
            background-color: #f5f5f5;
            padding: 15px;
            border-left: 4px solid #007ACC;
            white-space: pre-wrap;
            font-family: Consolas, monospace;
        }
        .footer {
            margin-top: 30px;
            text-align: center;
            font-size: 0.8em;
            color: #777;
        }
        .status-ok { color: green; }
        .status-warning { color: orange; }
        .status-error { color: red; }
        .status-info { color: blue; }
        
        /* Neue Styling für Management-Zusammenfassung */
        .exec-summary {
            background-color: #f9f9f9;
            padding: 20px;
            border-radius: 5px;
            margin-bottom: 30px;
            border: 1px solid #ddd;
        }
        .exec-summary h2 {
            margin-top: 0;
            color: #333;
            border-bottom: 2px solid #007ACC;
            padding-bottom: 10px;
        }
        .status-box {
            padding: 15px;
            border-radius: 5px;
            margin: 10px 0;
            font-weight: bold;
        }
        .status-good {
            background-color: #dff0d8;
            border-left: 5px solid #5cb85c;
        }
        .status-warning {
            background-color: #fcf8e3;
            border-left: 5px solid #f0ad4e;
        }
        .status-critical {
            background-color: #f2dede;
            border-left: 5px solid #d9534f;
        }
        .action-needed {
            background-color: #e8f4f8;
            border-left: 5px solid #5bc0de;
            padding: 15px;
            margin-top: 20px;
            border-radius: 5px;
        }
    </style>
</head>
<body>
    <h1>Windows Server 2012 Migration - Validierungsbericht</h1>
    <p><strong>Server:</strong> $($ValidationResults.ServerName)</p>
    <p><strong>Erstellungsdatum:</strong> $($ValidationResults.Timestamp.ToString("dd.MM.yyyy HH:mm:ss"))</p>
"@

        # Generiere Management-Zusammenfassung
        # Status des Servers bestimmen
        $errorCount = ($ValidationResults.Summary | Where-Object { $_.Status -eq "FEHLER" }).Count
        $warningCount = ($ValidationResults.Summary | Where-Object { $_.Status -eq "WARNUNG" }).Count
        $okCount = ($ValidationResults.Summary | Where-Object { $_.Status -eq "OK" }).Count
        
        # Leicht verständliche Gesamtbewertung
        $overallStatus = if ($errorCount -gt 0) {
            "kritisch"
        } elseif ($warningCount -gt 0) {
            "mit Einschränkungen funktionsfähig"
        } else {
            "voll funktionsfähig"
        }
        
        $overallStatusClass = if ($errorCount -gt 0) {
            "status-critical"
        } elseif ($warningCount -gt 0) {
            "status-warning"
        } else {
            "status-good"
        }

        # Sammlung von Problem-Zusammenfassungen für Handlungsempfehlungen
        $problemAreas = @()
        
        if (($ValidationResults.BasicTests | Where-Object { $_.Result -ne "OK" }).Count -gt 0) {
            $problemAreas += "Server-Basisfunktionen (möglicherweise Performance- oder Festplattenprobleme)"
        }
        
        if (($ValidationResults.ADTests | Where-Object { $_.Result -ne "OK" }).Count -gt 0) {
            $problemAreas += "Active Directory-Dienste"
        }
        
        if (($ValidationResults.Replication | Where-Object { $_.Status -ne "OK" }).Count -gt 0) {
            $problemAreas += "Active Directory-Replikation"
        }
        
        if (($ValidationResults.DNSService | Where-Object { $_.Result -ne "OK" }).Count -gt 0) {
            $problemAreas += "DNS-Dienst"
        }
        
        if (($ValidationResults.DHCPService | Where-Object { $_.Result -ne "OK" }).Count -gt 0) {
            $problemAreas += "DHCP-Dienst"
        }
        
        if (($ValidationResults.Services | Where-Object { $_.Status -ne "Running" }).Count -gt 0) {
            $problemAreas += "Kritische Windows-Dienste"
        }
        
        if (($ValidationResults.Performance | Where-Object { $_.Status -ne "OK" }).Count -gt 0) {
            $problemAreas += "Serverleistung"
        }
        
        # Handlungsempfehlungen basierend auf Status
        $actionNeeded = if ($errorCount -gt 0) {
            "Der Server weist kritische Probleme auf, die umgehend behoben werden müssen, um die Systemstabilität zu gewährleisten."
        } elseif ($warningCount -gt 0) {
            "Der Server funktioniert, zeigt aber Warnungen, die bei Gelegenheit überprüft werden sollten."
        } else {
            "Der Server läuft optimal. Es sind keine unmittelbaren Maßnahmen erforderlich."
        }
        
        # Detailliertere Handlungsempfehlungen bei Problemen
        $detailedAction = if ($problemAreas.Count -gt 0) {
            "Besondere Aufmerksamkeit erfordern: " + ($problemAreas -join ", ")
        } else {
            "Alle geprüften Bereiche sind in Ordnung."
        }
        
        # Management-Zusammenfassung erstellen
        $managementSummary = @"
<div class="exec-summary">
    <h2>Management-Zusammenfassung</h2>
    <p>
        Dieser Bericht fasst die Validierungsergebnisse für den Server <strong>$($ValidationResults.ServerName)</strong> zusammen. 
        Der Server wurde auf Funktionalität, Performance und Konfiguration geprüft.
    </p>
    
    <div class="status-box $overallStatusClass">
        Gesamtstatus: Server ist $overallStatus
    </div>
    
    <p><strong>Prüfungsergebnisse:</strong></p>
    <ul>
        <li>$okCount Bereiche zeigen optimale Funktion</li>
        <li>$warningCount Bereiche zeigen Warnungen</li>
        <li>$errorCount Bereiche zeigen kritische Fehler</li>
    </ul>
    
    <div class="action-needed">
        <p><strong>Handlungsempfehlung:</strong></p>
        <p>$actionNeeded</p>
        <p>$detailedAction</p>
    </div>
</div>
"@

        # Zusammenfassung
        $htmlSummary = "<h2>Zusammenfassung</h2>"
        $htmlSummary += "<table><tr><th>Kategorie</th><th>Komponente</th><th>Status</th><th>Details</th></tr>"
        
        foreach ($item in $ValidationResults.Summary) {
            $statusClass = switch ($item.Status) {
                "OK" { "status-ok" }
                "WARNUNG" { "status-warning" }
                "FEHLER" { "status-error" }
                default { "status-info" }
            }
            
            $htmlSummary += "<tr><td>$($item.Category)</td><td>$($item.Component)</td><td class='$statusClass'>$($item.Status)</td><td>$($item.Details)</td></tr>"
        }
        
        $htmlSummary += "</table>"

        # Grundlegende Tests
        $htmlBasicTests = "<h2>Grundlegende Tests</h2>"
        $htmlBasicTests += "<table><tr><th>Test</th><th>Ergebnis</th><th>Details</th></tr>"
        
        foreach ($test in $ValidationResults.BasicTests) {
            $statusClass = switch ($test.Result) {
                "OK" { "status-ok" }
                "WARNUNG" { "status-warning" }
                "FEHLER" { "status-error" }
                default { "status-info" }
            }
            
            $htmlBasicTests += "<tr><td>$($test.TestName)</td><td class='$statusClass'>$($test.Result)</td><td>$($test.Details)</td></tr>"
        }
        
        $htmlBasicTests += "</table>"

        # Active Directory Tests
        $htmlADTests = ""
        if ($ValidationResults.ADTests.Count -gt 0) {
            $htmlADTests = "<h2>Active Directory Tests</h2>"
            $htmlADTests += "<table><tr><th>Test</th><th>Ergebnis</th><th>Details</th></tr>"
            
            foreach ($test in $ValidationResults.ADTests) {
                $statusClass = switch ($test.Result) {
                    "OK" { "status-ok" }
                    "WARNUNG" { "status-warning" }
                    "FEHLER" { "status-error" }
                    default { "status-info" }
                }
                
                $htmlADTests += "<tr><td>$($test.TestName)</td><td class='$statusClass'>$($test.Result)</td><td>$($test.Details)</td></tr>"
            }
            
            $htmlADTests += "</table>"
        }

        # AD-Replikation
        $htmlReplication = ""
        if ($ValidationResults.Replication.Count -gt 0) {
            $htmlReplication = "<h2>Active Directory-Replikation</h2>"
            $htmlReplication += "<table><tr><th>Quellserver</th><th>Zielserver</th><th>Letzte Replikation</th><th>Status</th><th>Details</th></tr>"
            
            foreach ($repl in $ValidationResults.Replication) {
                $statusClass = switch ($repl.Status) {
                    "OK" { "status-ok" }
                    "WARNUNG" { "status-warning" }
                    "FEHLER" { "status-error" }
                    default { "status-info" }
                }
                
                $htmlReplication += "<tr><td>$($repl.SourceServer)</td><td>$($repl.DestinationServer)</td><td>$($repl.LastReplication)</td><td class='$statusClass'>$($repl.Status)</td><td>$($repl.Details)</td></tr>"
            }
            
            $htmlReplication += "</table>"
        }

        # FSMO-Rollen
        $htmlFSMO = ""
        if ($ValidationResults.FSMORoles.Count -gt 0) {
            $htmlFSMO = "<h2>FSMO-Rollen</h2>"
            $htmlFSMO += "<table><tr><th>Rolle</th><th>Inhaber</th><th>Erreichbarkeit</th><th>Details</th></tr>"
            
            foreach ($role in $ValidationResults.FSMORoles) {
                $statusClass = switch ($role.Availability) {
                    "OK" { "status-ok" }
                    "WARNUNG" { "status-warning" }
                    "FEHLER" { "status-error" }
                    default { "status-info" }
                }
                
                $htmlFSMO += "<tr><td>$($role.Role)</td><td>$($role.Owner)</td><td class='$statusClass'>$($role.Availability)</td><td>$($role.Details)</td></tr>"
            }
            
            $htmlFSMO += "</table>"
        }

        # DNS-Dienst
        $htmlDNSService = ""
        if ($ValidationResults.DNSService.Count -gt 0) {
            $htmlDNSService = "<h2>DNS-Dienst</h2>"
            $htmlDNSService += "<table><tr><th>Test</th><th>Ergebnis</th><th>Details</th></tr>"
            
            foreach ($dns in $ValidationResults.DNSService) {
                $statusClass = switch ($dns.Result) {
                    "OK" { "status-ok" }
                    "WARNUNG" { "status-warning" }
                    "FEHLER" { "status-error" }
                    default { "status-info" }
                }
                
                $htmlDNSService += "<tr><td>$($dns.TestName)</td><td class='$statusClass'>$($dns.Result)</td><td>$($dns.Details)</td></tr>"
            }
            
            $htmlDNSService += "</table>"
        }

        # DNS-Zonen
        $htmlDNSZones = ""
        if ($ValidationResults.DNSZones.Count -gt 0) {
            $htmlDNSZones = "<h2>DNS-Zonen</h2>"
            $htmlDNSZones += "<table><tr><th>Zone</th><th>Typ</th><th>Status</th><th>Replikation</th><th>Details</th></tr>"
            
            foreach ($zone in $ValidationResults.DNSZones) {
                $statusClass = switch ($zone.Status) {
                    "OK" { "status-ok" }
                    "WARNUNG" { "status-warning" }
                    "FEHLER" { "status-error" }
                    default { "status-info" }
                }
                
                $htmlDNSZones += "<tr><td>$($zone.ZoneName)</td><td>$($zone.ZoneType)</td><td class='$statusClass'>$($zone.Status)</td><td>$($zone.Replication)</td><td>$($zone.Details)</td></tr>"
            }
            
            $htmlDNSZones += "</table>"
        }

        # DNS-Auflösung
        $htmlDNSResolution = ""
        if ($ValidationResults.DNSResolution.Count -gt 0) {
            $htmlDNSResolution = "<h2>DNS-Auflösungstests</h2>"
            $htmlDNSResolution += "<table><tr><th>Hostname</th><th>Aufgelöste IP</th><th>Status</th><th>Antwortzeit (ms)</th><th>Details</th></tr>"
            
            foreach ($res in $ValidationResults.DNSResolution) {
                $statusClass = switch ($res.Status) {
                    "OK" { "status-ok" }
                    "WARNUNG" { "status-warning" }
                    "FEHLER" { "status-error" }
                    default { "status-info" }
                }
                
                $htmlDNSResolution += "<tr><td>$($res.Hostname)</td><td>$($res.ResolvedIP)</td><td class='$statusClass'>$($res.Status)</td><td>$($res.ResponseTime)</td><td>$($res.Details)</td></tr>"
            }
            
            $htmlDNSResolution += "</table>"
        }

        # DHCP-Dienst
        $htmlDHCPService = ""
        if ($ValidationResults.DHCPService.Count -gt 0) {
            $htmlDHCPService = "<h2>DHCP-Dienst</h2>"
            $htmlDHCPService += "<table><tr><th>Test</th><th>Ergebnis</th><th>Details</th></tr>"
            
            foreach ($dhcp in $ValidationResults.DHCPService) {
                $statusClass = switch ($dhcp.Result) {
                    "OK" { "status-ok" }
                    "WARNUNG" { "status-warning" }
                    "FEHLER" { "status-error" }
                    default { "status-info" }
                }
                
                $htmlDHCPService += "<tr><td>$($dhcp.TestName)</td><td class='$statusClass'>$($dhcp.Result)</td><td>$($dhcp.Details)</td></tr>"
            }
            
            $htmlDHCPService += "</table>"
        }

        # DHCP-Bereiche
        $htmlDHCPScopes = ""
        if ($ValidationResults.DHCPScopes.Count -gt 0) {
            $htmlDHCPScopes = "<h2>DHCP-Bereiche</h2>"
            $htmlDHCPScopes += "<table><tr><th>Bereich</th><th>Name</th><th>Status</th><th>Belegung</th><th>Leases</th><th>Details</th></tr>"
            
            foreach ($scope in $ValidationResults.DHCPScopes) {
                $statusClass = switch ($scope.Status) {
                    "OK" { "status-ok" }
                    "WARNUNG" { "status-warning" }
                    "FEHLER" { "status-error" }
                    default { "status-info" }
                }
                
                $htmlDHCPScopes += "<tr><td>$($scope.ScopeId)</td><td>$($scope.ScopeName)</td><td class='$statusClass'>$($scope.Status)</td><td>$($scope.Usage)</td><td>$($scope.Leases)</td><td>$($scope.Details)</td></tr>"
            }
            
            $htmlDHCPScopes += "</table>"
        }

        # DHCP-Failover
        $htmlDHCPFailover = ""
        if ($ValidationResults.DHCPFailover.Count -gt 0) {
            $htmlDHCPFailover = "<h2>DHCP-Failover</h2>"
            $htmlDHCPFailover += "<table><tr><th>Bereich</th><th>Partner</th><th>Modus</th><th>Status</th><th>Details</th></tr>"
            
            foreach ($failover in $ValidationResults.DHCPFailover) {
                $statusClass = switch ($failover.Status) {
                    "OK" { "status-ok" }
                    "WARNUNG" { "status-warning" }
                    "FEHLER" { "status-error" }
                    default { "status-info" }
                }
                
                $htmlDHCPFailover += "<tr><td>$($failover.ScopeId)</td><td>$($failover.PartnerServer)</td><td>$($failover.Mode)</td><td class='$statusClass'>$($failover.Status)</td><td>$($failover.Details)</td></tr>"
            }
            
            $htmlDHCPFailover += "</table>"
        }

        # Kritische Dienste
        $htmlServices = ""
        if ($ValidationResults.Services.Count -gt 0) {
            $htmlServices = "<h2>Kritische Dienste</h2>"
            $htmlServices += "<table><tr><th>Dienst</th><th>Anzeigename</th><th>Status</th><th>Starttyp</th><th>Details</th></tr>"
            
            foreach ($service in $ValidationResults.Services) {
                $statusClass = if ($service.Status -eq "Running") {
                    "status-ok"
                } else {
                    "status-warning"
                }
                
                $htmlServices += "<tr><td>$($service.ServiceName)</td><td>$($service.DisplayName)</td><td class='$statusClass'>$($service.Status)</td><td>$($service.StartType)</td><td>$($service.Details)</td></tr>"
            }
            
            $htmlServices += "</table>"
        }

        # Performance-Metriken
        $htmlPerformance = ""
        if ($ValidationResults.Performance.Count -gt 0) {
            $htmlPerformance = "<h2>Performance-Metriken</h2>"
            $htmlPerformance += "<table><tr><th>Metrik</th><th>Wert</th><th>Einheit</th><th>Status</th><th>Details</th></tr>"
            
            foreach ($perf in $ValidationResults.Performance) {
                $statusClass = switch ($perf.Status) {
                    "OK" { "status-ok" }
                    "WARNUNG" { "status-warning" }
                    "FEHLER" { "status-error" }
                    default { "status-info" }
                }
                
                $htmlPerformance += "<tr><td>$($perf.MetricName)</td><td>$($perf.Value)</td><td>$($perf.Unit)</td><td class='$statusClass'>$($perf.Status)</td><td>$($perf.Details)</td></tr>"
            }
            
            $htmlPerformance += "</table>"
        }

        # Ereignisprotokolle
        $htmlEventLogs = ""
        if ($ValidationResults.EventLogs.Count -gt 0) {
            $htmlEventLogs = "<h2>Ereignisprotokolle</h2>"
            $htmlEventLogs += "<table><tr><th>Protokoll</th><th>Schweregrad</th><th>Anzahl</th><th>Neuestes Ereignis</th><th>Details</th></tr>"
            
            foreach ($log in $ValidationResults.EventLogs) {
                $statusClass = switch ($log.Severity) {
                    "Error" { "status-error" }
                    "Warning" { "status-warning" }
                    "Information" { "status-info" }
                    default { "status-info" }
                }
                
                $htmlEventLogs += "<tr><td>$($log.LogName)</td><td class='$statusClass'>$($log.Severity)</td><td>$($log.Count)</td><td>$($log.LatestEvent)</td><td>$($log.Details)</td></tr>"
            }
            
            $htmlEventLogs += "</table>"
        }

        # GPO-Replikation
        $htmlGPO = ""
        if ($ValidationResults.GPO.Count -gt 0) {
            $htmlGPO = "<h2>GPO-Replikationsstatus</h2>"
            $htmlGPO += "<table><tr><th>GPO</th><th>Version</th><th>Replikationsstatus</th><th>Letzte Aenderung</th><th>Details</th></tr>"
            
            foreach ($gpo in $ValidationResults.GPO) {
                $statusClass = switch ($gpo.ReplicationStatus) {
                    "OK" { "status-ok" }
                    "WARNUNG" { "status-warning" }
                    "FEHLER" { "status-error" }
                    default { "status-info" }
                }
                
                $htmlGPO += "<tr><td>$($gpo.GPOName)</td><td>$($gpo.Version)</td><td class='$statusClass'>$($gpo.ReplicationStatus)</td><td>$($gpo.LastChange)</td><td>$($gpo.Details)</td></tr>"
            }
            
            $htmlGPO += "</table>"
        }

        # SYSVOL-Replikation
        $htmlSYSVOL = ""
        if ($ValidationResults.SYSVOL.Count -gt 0) {
            $htmlSYSVOL = "<h2>SYSVOL-Replikation</h2>"
            $htmlSYSVOL += "<table><tr><th>Server</th><th>Replikationsmechanismus</th><th>Status</th><th>Details</th></tr>"
            
            foreach ($sysvol in $ValidationResults.SYSVOL) {
                $statusClass = switch ($sysvol.Status) {
                    "OK" { "status-ok" }
                    "WARNUNG" { "status-warning" }
                    "FEHLER" { "status-error" }
                    default { "status-info" }
                }
                
                $htmlSYSVOL += "<tr><td>$($sysvol.Server)</td><td>$($sysvol.ReplicationType)</td><td class='$statusClass'>$($sysvol.Status)</td><td>$($sysvol.Details)</td></tr>"
            }
            
            $htmlSYSVOL += "</table>"
        }

        # Textbericht
        $htmlReport = ""
        if (-not [string]::IsNullOrEmpty($ValidationResults.Report)) {
            $htmlReport = "<h2>Vollständiger Textbericht</h2>"
            $htmlReport += "<pre class='summary'>$($ValidationResults.Report)</pre>"
        }

        # Footer - Korrektur des in-string Expressions
        $footerContent = "Erstellt mit Windows Server 2012 Migration Tool | $currentYear"
        $htmlFooter = @"
    <div class="footer">
        <p>$footerContent</p>
    </div>
</body>
</html>
"@

        # HTML-Datei zusammensetzen
        $fullHTML = $htmlHeader + 
                    $managementSummary +
                    $htmlSummary + 
                    $htmlBasicTests + 
                    $htmlADTests + 
                    $htmlReplication + 
                    $htmlFSMO + 
                    $htmlDNSService + 
                    $htmlDNSZones + 
                    $htmlDNSResolution + 
                    $htmlDHCPService + 
                    $htmlDHCPScopes + 
                    $htmlDHCPFailover + 
                    $htmlServices + 
                    $htmlPerformance + 
                    $htmlEventLogs + 
                    $htmlGPO + 
                    $htmlSYSVOL + 
                    $htmlReport + 
                    $htmlFooter

        # Datei speichern
        $fullHTML | Out-File -FilePath $FilePath -Encoding UTF8

        $logMessage = "Validierungsbericht erfolgreich als HTML exportiert nach $FilePath"
        Write-Log $logMessage -Level "SUCCESS"
        return $true
    }
    catch {
        $errorMsg = "Fehler beim Exportieren des Validierungsberichts als HTML: $($_.Exception.Message)"
        Write-Log $errorMsg -Level "ERROR"
        return $false
    }
}
#endregion

#region Installation von Serverrollen und Features

# Funktion zum Abrufen aller Serverrollen und Features
function Get-AvailableServerRoles {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ServerName
    )
    
    try {
        Write-Log "Rufe verfügbare Serverrollen für $ServerName ab" -Level "INFO"
        
        # Prüfen, ob Server erreichbar ist
        $serverReachable = Test-Connection -ComputerName $ServerName -Count 1 -Quiet
        if (-not $serverReachable) {
            Write-Log "Server $ServerName ist nicht erreichbar" -Level "ERROR"
            return $null
        }
        
        # Verfügbare Serverrollen und Features abfragen
        $script = {
            # Prüfen, ob das Windows Feature-Modul verfügbar ist (Server 2008 R2 und neuer)
            if (Get-Command Get-WindowsFeature -ErrorAction SilentlyContinue) {
                Get-WindowsFeature | Select-Object Name, DisplayName, InstallState, Description
            } else {
                # Für ältere Server-Versionen alternative Methode verwenden (weniger detailliert)
                $features = @()
                $serverManagerCmd = "$env:SystemRoot\system32\servermanagercmd.exe"
                
                if (Test-Path $serverManagerCmd) {
                    $output = & $serverManagerCmd -query
                    $features = $output | Where-Object { $_ -match "^\[X\]" -or $_ -match "^\[ \]" } | ForEach-Object {
                        $state = if ($_ -match "^\[X\]") { "Installed" } else { "Available" }
                        $name = $_.Substring(3).Trim()
                        
                        [PSCustomObject]@{
                            Name = $name
                            DisplayName = $name
                            InstallState = $state
                            Description = ""
                        }
                    }
                }
                
                $features
            }
        }
        
        # Führe Skript remote oder lokal aus
        if ($ServerName -eq $env:COMPUTERNAME) {
            $roles = & $script
        } else {
            $session = New-PSSession -ComputerName $ServerName -ErrorAction Stop
            $roles = Invoke-Command -Session $session -ScriptBlock $script
            Remove-PSSession $session
        }
        
        Write-Log "Erfolgreich $($roles.Count) Serverrollen für $ServerName abgerufen" -Level "SUCCESS"
        return $roles
    }
    catch {
        Write-Log "Fehler beim Abrufen der Serverrollen für $ServerName`: $($_.Exception.Message)" -Level "ERROR"
        return $null
    }
}

# Funktion zum Abrufen aller domänenverbundenen Server
function Get-DomainServers {
    param (
        [switch]$OnlyWindows
    )
    
    try {
        Write-Log "Rufe Server aus der Domäne ab" -Level "INFO"
        
        # Prüfen, ob AD-Modul verfügbar ist
        if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
            Write-Log "Active Directory-Modul ist nicht verfügbar" -Level "WARNING"
            return @()
        }
        
        # Active Directory-Modul laden
        Import-Module ActiveDirectory -ErrorAction Stop
        
        # Serverfilter definieren
        $filter = "operatingSystem -like '*Windows*' -and operatingSystem -like '*Server*'"
        
        # Server aus der Domäne abrufen
        $servers = Get-ADComputer -Filter $filter -Properties Name, OperatingSystem, OperatingSystemVersion, Enabled |
                   Where-Object { $_.Enabled -eq $true } |
                   Select-Object Name, @{Name="OS"; Expression={$_.OperatingSystem}}, 
                                @{Name="Version"; Expression={$_.OperatingSystemVersion}},
                                @{Name="Online"; Expression={
                                    Test-Connection -ComputerName $_.Name -Count 1 -Quiet -ErrorAction SilentlyContinue
                                }}
        
        Write-Log "$($servers.Count) Server in der Domäne gefunden" -Level "SUCCESS"
        return $servers
    }
    catch {
        Write-Log "Fehler beim Abrufen der Domänenserver: $($_.Exception.Message)" -Level "ERROR"
        return @()
    }
}

# Funktion zum Überprüfen der Voraussetzungen für die Rolleninstallation
function Test-ServerRoleInstallationRequirements {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ServerName,
        
        [Parameter(Mandatory = $true)]
        [string[]]$Roles
    )
    
    try {
        Write-Log "Überprüfe Voraussetzungen für Rolleninstallation auf $ServerName" -Level "INFO"
        
        $requirements = @()
        
        # Prüfen, ob Server erreichbar ist
        $serverReachable = Test-Connection -ComputerName $ServerName -Count 1 -Quiet
        if (-not $serverReachable) {
            $requirements += [PSCustomObject]@{
                Requirement = "Servererreichbarkeit"
                Status = "FEHLER"
                Details = "Server $ServerName ist nicht erreichbar"
            }
            return $requirements
        } else {
            $requirements += [PSCustomObject]@{
                Requirement = "Servererreichbarkeit"
                Status = "OK"
                Details = "Server $ServerName ist erreichbar"
            }
        }
        
        # Skript zum Überprüfen der Voraussetzungen
        $script = {
            param($RolesToCheck)
            
            $results = @()
            
            # Betriebssystem-Version prüfen
            $os = Get-WmiObject -Class Win32_OperatingSystem
            $osVersion = $os.Version
            $osCaption = $os.Caption
            
            $results += [PSCustomObject]@{
                Requirement = "Betriebssystem"
                Status = "INFO"
                Details = "$osCaption (Version $osVersion)"
            }
            
            # Freier Festplattenplatz auf C:
            $disk = Get-WmiObject -Class Win32_LogicalDisk -Filter "DeviceID = 'C:'"
            $freeSpace = [math]::Round($disk.FreeSpace / 1GB, 2)
            $totalSpace = [math]::Round($disk.Size / 1GB, 2)
            $percentFree = [math]::Round(($disk.FreeSpace / $disk.Size) * 100, 2)
            
            if ($freeSpace -lt 10) {
                $status = "WARNUNG"
                $details = "Nur $freeSpace GB freier Speicherplatz ($percentFree%) - mindestens 10 GB empfohlen"
            } else {
                $status = "OK"
                $details = "$freeSpace GB freier Speicherplatz ($percentFree%)"
            }
            
            $results += [PSCustomObject]@{
                Requirement = "Festplattenplatz"
                Status = $status
                Details = $details
            }
            
            # Arbeitsspeicher prüfen
            $totalRAM = [math]::Round((Get-WmiObject -Class Win32_ComputerSystem).TotalPhysicalMemory / 1GB, 2)
            
            if ($totalRAM -lt 2) {
                $status = "WARNUNG"
                $details = "Nur $totalRAM GB RAM vorhanden - mindestens 2 GB empfohlen"
            } else {
                $status = "OK"
                $details = "$totalRAM GB RAM vorhanden"
            }
            
            $results += [PSCustomObject]@{
                Requirement = "Arbeitsspeicher"
                Status = $status
                Details = $details
            }
            
            # Serverrollen-spezifische Prüfungen
            if (Get-Command Get-WindowsFeature -ErrorAction SilentlyContinue) {
                # AD DS Überprüfung
                if ($RolesToCheck -contains "AD-Domain-Services") {
                    # Prüfe, ob der Server bereits ein DC ist
                    $isDC = (Get-WmiObject -Class Win32_ComputerSystem).DomainRole -ge 4
                    
                    if ($isDC) {
                        $results += [PSCustomObject]@{
                            Requirement = "Active Directory Domain Services"
                            Status = "WARNUNG"
                            Details = "Server ist bereits ein Domänencontroller"
                        }
                    } else {
                        $results += [PSCustomObject]@{
                            Requirement = "Active Directory Domain Services"
                            Status = "OK"
                            Details = "Server kann zu einem Domänencontroller hochgestuft werden"
                        }
                    }
                }
                
                # DNS-Server Überprüfung
                if ($RolesToCheck -contains "DNS") {
                    $dnsService = Get-Service -Name DNS -ErrorAction SilentlyContinue
                    
                    if ($dnsService) {
                        $results += [PSCustomObject]@{
                            Requirement = "DNS-Server"
                            Status = "INFO"
                            Details = "DNS-Server-Rolle ist bereits installiert"
                        }
                    } else {
                        $results += [PSCustomObject]@{
                            Requirement = "DNS-Server"
                            Status = "OK"
                            Details = "DNS-Server-Rolle kann installiert werden"
                        }
                    }
                }
                
                # DHCP-Server Überprüfung
                if ($RolesToCheck -contains "DHCP") {
                    $dhcpService = Get-Service -Name DHCPServer -ErrorAction SilentlyContinue
                    
                    if ($dhcpService) {
                        $results += [PSCustomObject]@{
                            Requirement = "DHCP-Server"
                            Status = "INFO"
                            Details = "DHCP-Server-Rolle ist bereits installiert"
                        }
                    } else {
                        # Prüfen, ob die IP-Adresse statisch ist
                        $networkConfig = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -Filter "IPEnabled = 'True'"
                        $staticIp = $networkConfig | Where-Object { $_.DHCPEnabled -eq $false -and $_.IPAddress -ne $null }
                        
                        if ($staticIp) {
                            $results += [PSCustomObject]@{
                                Requirement = "DHCP-Server"
                                Status = "OK"
                                Details = "DHCP-Server-Rolle kann installiert werden, statische IP-Adresse erkannt"
                            }
                        } else {
                            $results += [PSCustomObject]@{
                                Requirement = "DHCP-Server"
                                Status = "WARNUNG"
                                Details = "DHCP-Server benötigt eine statische IP-Adresse"
                            }
                        }
                    }
                }
            }
            
            return $results
        }
        
        # Skript remote oder lokal ausführen
        if ($ServerName -eq $env:COMPUTERNAME) {
            $remoteRequirements = & $script -RolesToCheck $Roles
        } else {
            $session = New-PSSession -ComputerName $ServerName -ErrorAction Stop
            $remoteRequirements = Invoke-Command -Session $session -ScriptBlock $script -ArgumentList (,$Roles)
            Remove-PSSession $session
        }
        
        # Remote-Ergebnisse in die Gesamtliste einfügen
        foreach ($req in $remoteRequirements) {
            $requirements += $req
        }
        
        # PowerShell Remoting-Voraussetzungen prüfen (nur relevant für Remote-Installation)
        if ($ServerName -ne $env:COMPUTERNAME) {
            # WinRM-Dienst auf dem Zielserver prüfen
            try {
                $remoteWinRM = Invoke-Command -ComputerName $ServerName -ScriptBlock {
                    Get-Service -Name WinRM
                } -ErrorAction Stop
                
                $requirements += [PSCustomObject]@{
                    Requirement = "PowerShell Remoting"
                    Status = "OK"
                    Details = "WinRM-Dienst läuft auf Zielserver, Remote-Installation möglich"
                }
            }
            catch {
                $requirements += [PSCustomObject]@{
                    Requirement = "PowerShell Remoting"
                    Status = "FEHLER"
                    Details = "WinRM-Dienst auf Zielserver nicht erreichbar: $($_.Exception.Message)"
                }
            }
        }
        
        Write-Log "Voraussetzungsprüfung für $ServerName abgeschlossen" -Level "SUCCESS"
        return $requirements
    }
    catch {
        Write-Log "Fehler bei der Überprüfung der Voraussetzungen für $ServerName`: $($_.Exception.Message)" -Level "ERROR"
        return @(
            [PSCustomObject]@{
                Requirement = "Allgemein"
                Status = "FEHLER"
                Details = "Konnte Voraussetzungen nicht prüfen: $($_.Exception.Message)"
            }
        )
    }
}

# Funktion zum Installieren von Serverrollen
function Install-ServerRoles {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ServerName,
        
        [Parameter(Mandatory = $true)]
        [string[]]$Roles,
        
        [scriptblock]$ProgressCallback = $null
    )
    
    try {
        Write-Log "Starte Installation der Serverrollen auf $ServerName" -Level "INFO"
        
        $result = @{
            Success = $false
            RestartRequired = $false
            Results = @()
        }
        
        # Prüfen, ob Server erreichbar ist
        $serverReachable = Test-Connection -ComputerName $ServerName -Count 1 -Quiet
        if (-not $serverReachable) {
            Write-Log "Server $ServerName ist nicht erreichbar" -Level "ERROR"
            $result.Results += @{
                Role = "Verbindung"
                Status = "FEHLER"
                Details = "Server ist nicht erreichbar"
            }
            return $result
        }
        
        # Fortschrittsupdate
        if ($null -ne $ProgressCallback) {
            $ProgressCallback.Invoke("Verbindung zu $ServerName hergestellt...", 5)
        }
        
        # Vorhandene Rollen prüfen
        $script = {
            param($rolesToInstall)
            
            $result = @{
                InstalledRoles = @()
                FailedRoles = @()
                RestartRequired = $false
            }
            
            try {
                # ServerManager-Modul importieren
                Import-Module ServerManager -ErrorAction Stop
                
                foreach ($role in $rolesToInstall) {
                    try {
                        # Prüfen, ob Rolle bereits installiert
                        $roleInstalled = Get-WindowsFeature -Name $role -ErrorAction SilentlyContinue
                        
                        if ($roleInstalled -and $roleInstalled.Installed) {
                            $result.InstalledRoles += @{
                                Role = $role
                                Status = "BEREITS INSTALLIERT"
                                Details = "Rolle bereits installiert"
                            }
                            continue
                        }
                        
                        # Rolle installieren
                        $installation = Install-WindowsFeature -Name $role -ErrorAction Stop
                        
                        if ($installation.Success) {
                            $result.InstalledRoles += @{
                                Role = $role
                                Status = "ERFOLGREICH"
                                Details = "Installation erfolgreich"
                            }
                            
                            # Prüfen, ob Neustart erforderlich
                            if ($installation.RestartNeeded -eq "Yes") {
                                $result.RestartRequired = $true
                            }
                        } else {
                            $result.FailedRoles += @{
                                Role = $role
                                Status = "FEHLER"
                                Details = "Installation fehlgeschlagen"
                            }
                        }
                    }
                    catch {
                        $result.FailedRoles += @{
                            Role = $role
                            Status = "FEHLER"
                            Details = "Fehler bei der Installation: $($_.Exception.Message)"
                        }
                    }
                }
            }
            catch {
                $result.FailedRoles += @{
                    Role = "ServerManager"
                    Status = "FEHLER"
                    Details = "Konnte ServerManager-Modul nicht laden: $($_.Exception.Message)"
                }
            }
            
            return $result
        }
        
        # Fortschrittsupdate
        if ($null -ne $ProgressCallback) {
            $ProgressCallback.Invoke("Installiere ausgewählte Rollen...", 25)
        }
        
        # Skript auf Zielserver ausführen
        if ($ServerName -eq $env:COMPUTERNAME) {
            $installationResult = & $script -rolesToInstall $Roles
        } else {
            $installationResult = Invoke-Command -ComputerName $ServerName -ScriptBlock $script -ArgumentList (,$Roles) -ErrorAction Stop
        }
        
        # Fortschrittsupdate
        if ($null -ne $ProgressCallback) {
            $ProgressCallback.Invoke("Verarbeite Ergebnisse...", 75)
        }
        
        # Ergebnisse verarbeiten
        if ($installationResult.InstalledRoles) {
            $result.Results += $installationResult.InstalledRoles
        }
        
        if ($installationResult.FailedRoles) {
            $result.Results += $installationResult.FailedRoles
        }
        
        # Neustart erforderlich?
        $result.RestartRequired = $installationResult.RestartRequired
        
        # Erfolg bestimmen
        $result.Success = ($installationResult.FailedRoles.Count -eq 0)
        
        # Fortschrittsupdate
        if ($null -ne $ProgressCallback) {
            $ProgressCallback.Invoke("Installation abgeschlossen", 100)
        }
        
        Write-Log "Rolleninstallation auf $ServerName abgeschlossen. Erfolg: $($result.Success), Neustart erforderlich: $($result.RestartRequired)" -Level "INFO"
        return $result
    }
    catch {
        Write-Log "Fehler bei der Installation von Serverrollen auf $ServerName`: $($_.Exception.Message)" -Level "ERROR"
        return @{
            Success = $false
            RestartRequired = $false
            Results = @(
                @{
                    Role = "Allgemein"
                    Status = "FEHLER"
                    Details = "Ein unerwarteter Fehler ist aufgetreten: $($_.Exception.Message)"
                }
            )
        }
    }
}

# Funktion zum Testen der FSMO-Rollen
function Test-FSMORoles {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ServerName
    )
    
    $results = @()
    
    try {
        # Forest FSMO-Rollen
        try {
            $forest = Get-ADForest -ErrorAction Stop
            
            # Schema Master
            $schemaMaster = $forest.SchemaMaster
            $schemaAvailability = "OK"
            
            # Ping-Test am Schema Master
            if (-not (Test-Connection -ComputerName $schemaMaster -Count 1 -Quiet -ErrorAction SilentlyContinue)) {
                $schemaAvailability = "FEHLER"
            }
            
            $results += [PSCustomObject]@{
                Role = "Schema Master"
                Owner = $schemaMaster
                Availability = $schemaAvailability
                Details = "Forest-weiter Schema Master"
            }
            
            # Domain Naming Master
            $domainNamingMaster = $forest.DomainNamingMaster
            $domainNamingAvailability = "OK"
            
            # Ping-Test am Domain Naming Master
            if (-not (Test-Connection -ComputerName $domainNamingMaster -Count 1 -Quiet -ErrorAction SilentlyContinue)) {
                $domainNamingAvailability = "FEHLER"
            }
            
            $results += [PSCustomObject]@{
                Role = "Domain Naming Master"
                Owner = $domainNamingMaster
                Availability = $domainNamingAvailability
                Details = "Forest-weiter Domain Naming Master"
            }
        }
        catch {
            Write-Log "Fehler beim Abrufen der Forest FSMO-Rollen: $($_.Exception.Message)" -Level "WARNING"
        }
        
        # Domain FSMO-Rollen
        try {
            $domain = Get-ADDomain -ErrorAction Stop
            
            # PDC Emulator
            $pdcEmulator = $domain.PDCEmulator
            $pdcAvailability = "OK"
            
            # Ping-Test am PDC Emulator
            if (-not (Test-Connection -ComputerName $pdcEmulator -Count 1 -Quiet -ErrorAction SilentlyContinue)) {
                $pdcAvailability = "FEHLER"
            }
            
            $results += [PSCustomObject]@{
                Role = "PDC Emulator"
                Owner = $pdcEmulator
                Availability = $pdcAvailability
                Details = "Primary Domain Controller Emulator (Zeitserver)"
            }
            
            # RID Master
            $ridMaster = $domain.RIDMaster
            $ridAvailability = "OK"
            
            # Ping-Test am RID Master
            if (-not (Test-Connection -ComputerName $ridMaster -Count 1 -Quiet -ErrorAction SilentlyContinue)) {
                $ridAvailability = "FEHLER"
            }
            
            $results += [PSCustomObject]@{
                Role = "RID Master"
                Owner = $ridMaster
                Availability = $ridAvailability
                Details = "Relative ID Master (verwaltet SID-Pools)"
            }
            
            # Infrastructure Master
            $infraMaster = $domain.InfrastructureMaster
            $infraAvailability = "OK"
            
            # Ping-Test am Infrastructure Master
            if (-not (Test-Connection -ComputerName $infraMaster -Count 1 -Quiet -ErrorAction SilentlyContinue)) {
                $infraAvailability = "FEHLER"
            }
            
            $results += [PSCustomObject]@{
                Role = "Infrastructure Master"
                Owner = $infraMaster
                Availability = $infraAvailability
                Details = "Infrastructure Master (verwaltet Referenzen zu Objekten)"
            }
        }
        catch {
            Write-Log "Fehler beim Abrufen der Domain FSMO-Rollen: $($_.Exception.Message)" -Level "WARNING"
        }
        
        return $results
    }
    catch {
        Write-Log "Fehler bei FSMO-Rollen-Checks für $ServerName`: $($_.Exception.Message)" -Level "ERROR"
        return @(
            [PSCustomObject]@{
                Role = "FSMO-Fehler"
                Owner = "N/A"
                Availability = "FEHLER"
                Details = "Ein allgemeiner Fehler ist aufgetreten: $($_.Exception.Message)"
            }
        )
    }
}

# Funktion zum Neustart eines Remote-Servers
function Restart-RemoteServer {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ServerName,
        
        [int]$TimeoutSeconds = 300
    )
    
    try {
        Write-Log "Starte Neustart des Servers $ServerName" -Level "INFO"
        
        # Prüfen, ob es sich um den lokalen Computer handelt
        if ($ServerName -eq $env:COMPUTERNAME) {
            # MessageBox mit Neustart-Information anzeigen
            $result = [System.Windows.Forms.MessageBox]::Show(
                "Der lokale Server muss neu gestartet werden, um die Installation abzuschließen. Möchten Sie jetzt neu starten?",
                "Neustart erforderlich", 
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question
            )
            
            if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
                # Lokalen Neustart veranlassen
                Write-Log "Lokaler Server wird neu gestartet" -Level "WARNING"
                Restart-Computer -Force
                return $true
            } else {
                Write-Log "Neustart wurde abgebrochen" -Level "WARNING"
                return $false
            }
        } else {
            # Remote-Server neu starten
            $result = Restart-Computer -ComputerName $ServerName -Force -Wait -For PowerShell -Timeout $TimeoutSeconds -ErrorAction Stop
            
            if ($result) {
                Write-Log "Server $ServerName wurde erfolgreich neu gestartet" -Level "SUCCESS"
                return $true
            } else {
                Write-Log "Konnte nicht bestätigen, dass Server $ServerName nach dem Neustart wieder verfügbar ist" -Level "WARNING"
                return $false
            }
        }
    }
    catch {
        Write-Log "Fehler beim Neustart des Servers $ServerName`: $($_.Exception.Message)" -Level "ERROR"
        return $false
    }
}

# Funktion zur Übersetzung von Rollenbezeichnern in Windows Feature-Namen
function Convert-RoleIdentifierToFeatureName {
    param (
        [string]$RoleIdentifier
    )
    
    $roleMapping = @{
        "ADDS" = "AD-Domain-Services"
        "ADCS" = "AD-Certificate"
        "ADLDS" = "ADLDS"
        "ADRMS" = "ADRMS"
        "DNS" = "DNS"
        "DHCP" = "DHCP"
        "WINS" = "WINS"
        "FileServer" = "File-Services"
        "DFS" = "FS-DFS-Namespace"
        "DFSR" = "FS-DFS-Replication"
        "WebServer" = "Web-Server"
        "WSUS" = "UpdateServices"
    }
    
    if ($roleMapping.ContainsKey($RoleIdentifier)) {
        return $roleMapping[$RoleIdentifier]
    } else {
        return $RoleIdentifier
    }
}
#endregion

# Fenster anzeigen und Anwendung starten
try {
    # Self-Diagnostics noch einmal ausführen
    $assetsOK = Test-RequiredAssets
    if (-not $assetsOK) {
        Write-Log "Warnung: Fehlende Assets könnten die Funktionalität beeinträchtigen" -Level "WARNING"
    }
    
    # Fenster modal anzeigen (blockiert bis Fenster geschlossen wird)
    Write-Log "Starte GUI-Anwendung" -Level "INFO"
    $result = $window.ShowDialog()
    Write-Log "GUI-Anwendung wurde beendet mit Ergebnis: $result" -Level "INFO"
}
catch {
    Write-Log "Kritischer Fehler beim Starten der GUI: $($_.Exception.Message)" -Level "ERROR"
    [System.Windows.Forms.MessageBox]::Show("Die Anwendung konnte nicht gestartet werden.`n`nFehlerdetails:`n$($_.Exception.Message)", "Kritischer Fehler", "OK", "Error")
}
