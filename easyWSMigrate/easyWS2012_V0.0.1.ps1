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
# Eingebettetes XAML für moderne Windows 11 UI
Write-Log "Verwende eingebettetes XAML für moderne Windows 11 UI" -Level "INFO"
$XAML = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="{AppName}" 
    Height="800" 
    Width="1200" 
    WindowStartupLocation="CenterScreen"
    ResizeMode="CanResize"
    MinWidth="900"
    MinHeight="600"
    Background="#FFFFFF">
    
    <Window.Resources>
        <!-- Theme Colors -->
        <SolidColorBrush x:Key="ThemeBrush" Color="{ThemeColor}"/>
        <SolidColorBrush x:Key="DarkModeBrush" Color="{ThemeColor}"/>
        
        <!-- Navigation Button Style -->
        <Style x:Key="NavigationButton" TargetType="Button">
            <Setter Property="Height" Value="46"/>
            <Setter Property="Margin" Value="8,4"/>
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="#333333"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="FontWeight" Value="Normal"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="HorizontalContentAlignment" Value="Left"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" Background="{TemplateBinding Background}" CornerRadius="4" Padding="12,0">
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="28"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                <ContentPresenter x:Name="icon" Grid.Column="0" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                <TextBlock x:Name="text" Grid.Column="1" Text="{TemplateBinding Content}" VerticalAlignment="Center" Margin="12,0,0,0"/>
                            </Grid>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#F5F5F5"/>
                                <Setter Property="Foreground" Value="#000000"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#EEEEEE"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Foreground" Value="#AAAAAA"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        
        <!-- Active Navigation Button Style -->
        <Style x:Key="ActiveNavigationButton" TargetType="Button" BasedOn="{StaticResource NavigationButton}">
            <Setter Property="Background" Value="#F0F0F0"/>
            <Setter Property="Foreground" Value="{StaticResource ThemeBrush}"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="BorderThickness" Value="0,0,4,0"/>
            <Setter Property="BorderBrush" Value="{StaticResource ThemeBrush}"/>
        </Style>
        
        <!-- Standard Button Style -->
        <Style x:Key="StandardButton" TargetType="Button">
            <Setter Property="Height" Value="36"/>
            <Setter Property="MinWidth" Value="120"/>
            <Setter Property="Padding" Value="16,0"/>
            <Setter Property="Background" Value="{StaticResource ThemeBrush}"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" Background="{TemplateBinding Background}" CornerRadius="4">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Opacity" Value="0.9"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Opacity" Value="0.8"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="border" Property="Opacity" Value="0.5"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        
        <!-- Secondary Button Style -->
        <Style x:Key="SecondaryButton" TargetType="Button">
            <Setter Property="Height" Value="36"/>
            <Setter Property="MinWidth" Value="120"/>
            <Setter Property="Padding" Value="16,0"/>
            <Setter Property="Background" Value="#F5F5F5"/>
            <Setter Property="Foreground" Value="#333333"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="BorderBrush" Value="#DDDDDD"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" 
                                BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="4">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#EEEEEE"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#E0E0E0"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="border" Property="Opacity" Value="0.5"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        
        <!-- Card Style -->
        <Style x:Key="CardStyle" TargetType="Border">
            <Setter Property="Background" Value="White"/>
            <Setter Property="BorderBrush" Value="#E0E0E0"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="CornerRadius" Value="6"/>
            <Setter Property="Padding" Value="16"/>
            <Setter Property="Margin" Value="0,0,0,16"/>
            <Setter Property="Effect">
                <Setter.Value>
                    <DropShadowEffect ShadowDepth="1" Direction="315" BlurRadius="4" Opacity="0.1" Color="#000000"/>
                </Setter.Value>
            </Setter>
        </Style>
        
        <!-- Tab Control Style -->
        <Style TargetType="TabControl">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="0"/>
        </Style>
        
        <Style TargetType="TabItem">
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TabItem">
                        <Border x:Name="Border" BorderThickness="0,0,0,2" BorderBrush="Transparent" Margin="0,0,16,0" 
                                Background="Transparent" Padding="6,8">
                            <ContentPresenter x:Name="ContentSite" ContentSource="Header" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="Border" Property="BorderBrush" Value="{StaticResource ThemeBrush}"/>
                                <Setter Property="Foreground" Value="{StaticResource ThemeBrush}"/>
                                <Setter Property="FontWeight" Value="SemiBold"/>
                            </Trigger>
                            <Trigger Property="IsSelected" Value="False">
                                <Setter Property="Foreground" Value="#666666"/>
                            </Trigger>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Border" Property="Background" Value="#F5F5F5"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        
        <!-- DataGrid Styles -->
        <Style TargetType="DataGrid">
            <Setter Property="Background" Value="White"/>
            <Setter Property="BorderBrush" Value="#E0E0E0"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="RowBackground" Value="White"/>
            <Setter Property="AlternatingRowBackground" Value="#F8F8F8"/>
            <Setter Property="HorizontalGridLinesBrush" Value="#E0E0E0"/>
            <Setter Property="VerticalGridLinesBrush" Value="#E0E0E0"/>
            <Setter Property="RowHeight" Value="32"/>
            <Setter Property="HeadersVisibility" Value="Column"/>
        </Style>
        
        <Style TargetType="DataGridColumnHeader">
            <Setter Property="Background" Value="#F5F5F5"/>
            <Setter Property="Foreground" Value="#333333"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Padding" Value="8"/>
            <Setter Property="BorderBrush" Value="#E0E0E0"/>
            <Setter Property="BorderThickness" Value="0,0,1,1"/>
        </Style>
    </Window.Resources>
    
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="64"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="48"/>
        </Grid.RowDefinitions>
        
        <!-- Header -->
        <Border Grid.Row="0" Background="{StaticResource ThemeBrush}" BorderBrush="#E0E0E0" BorderThickness="0,0,0,1">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="250"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                
                <TextBlock Grid.Column="0" Text="{AppName}" FontSize="18" FontWeight="SemiBold" 
                           Foreground="White" VerticalAlignment="Center" Margin="24,0,0,0"/>
                
                <TextBlock Grid.Column="1" x:Name="txtCurrentPageTitle" Text="Dashboard" FontSize="16" 
                           Foreground="White" VerticalAlignment="Center" Margin="24,0,0,0"/>
                
                <StackPanel Grid.Column="2" Orientation="Horizontal" Margin="0,0,16,0">
                    <Button x:Name="btnHelp" Width="36" Height="36" Background="Transparent" BorderThickness="0" 
                            ToolTip="Help" Margin="8,0">
                        <Path Data="M12,2C6.48,2 2,6.48 2,12s4.48,10 10,10 10-4.48 10-10S17.52,2 12,2zm1,17h-2v-2h2v2zm2.07-7.75l-.9.92c-.5.51-.86.97-1.04,1.69-.08.32-.13.68-.13,1.14h-2v-.5c0-.46.08-.9.22-1.31.2-.58.53-1.1.95-1.52l1.24-1.26c.46-.44.68-1.1.55-1.8-.13-.72-.69-1.33-1.39-1.53-1.11-.31-2.14.32-2.47,1.27-.12.35-.47.59-.85.59h-.55c-.56,0-.96-.53-.81-1.07.57-1.91 2.38-3.21 4.43-3.15 1.55.05 2.92.96 3.6 2.27.61 1.17.49 2.57-.3 3.6l-.79.82c-.28.27-.51.54-.68.81-.17.27-.3.52-.37.78-.05.2-.08.4-.08.58h-2c0-.16.05-.3.1-.45.14-.36.35-.68.64-.95z" 
                               Fill="White" Stretch="Uniform" Width="20" Height="20"/>
                    </Button>
                    <Button x:Name="btnSettingsNav" Width="36" Height="36" Background="Transparent" BorderThickness="0" 
                            ToolTip="Settings" Margin="8,0">
                        <Path Data="M12,15.5A3.5,3.5 0 0,1 8.5,12A3.5,3.5 0 0,1 12,8.5A3.5,3.5 0 0,1 15.5,12A3.5,3.5 0 0,1 12,15.5M19.43,12.97C19.47,12.65 19.5,12.33 19.5,12C19.5,11.67 19.47,11.34 19.43,11L21.54,9.37C21.73,9.22 21.78,8.95 21.66,8.73L19.66,5.27C19.54,5.05 19.27,4.96 19.05,5.05L16.56,6.05C16.04,5.66 15.5,5.32 14.87,5.07L14.5,2.42C14.46,2.18 14.25,2 14,2H10C9.75,2 9.54,2.18 9.5,2.42L9.13,5.07C8.5,5.32 7.96,5.66 7.44,6.05L4.95,5.05C4.73,4.96 4.46,5.05 4.34,5.27L2.34,8.73C2.22,8.95 2.27,9.22 2.46,9.37L4.57,11C4.53,11.34 4.5,11.67 4.5,12C4.5,12.33 4.53,12.65 4.57,12.97L2.46,14.63C2.27,14.78 2.22,15.05 2.34,15.27L4.34,18.73C4.46,18.95 4.73,19.03 4.95,18.95L7.44,17.94C7.96,18.34 8.5,18.68 9.13,18.93L9.5,21.58C9.54,21.82 9.75,22 10,22H14C14.25,22 14.46,21.82 14.5,21.58L14.87,18.93C15.5,18.67 16.04,18.34 16.56,17.94L19.05,18.95C19.27,19.03 19.54,18.95 19.66,18.73L21.66,15.27C21.78,15.05 21.73,14.78 21.54,14.63L19.43,12.97Z" 
                              Fill="White" Stretch="Uniform" Width="20" Height="20"/>
                    </Button>
                </StackPanel>
            </Grid>
        </Border>
        
        <!-- Main Content Area -->
        <Grid Grid.Row="1" Background="#F5F5F5">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="250"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            
            <!-- Navigation Panel -->
            <Border Grid.Column="0" Background="White" BorderBrush="#E0E0E0" BorderThickness="0,0,1,0">
                <ScrollViewer VerticalScrollBarVisibility="Auto">
                    <StackPanel Margin="16">
                        <TextBlock Text="Main Menu" FontSize="14" FontWeight="SemiBold" Margin="0,0,0,16" Foreground="#666666"/>
                        
                        <ListBox x:Name="listNavigation" BorderThickness="0" Background="Transparent">
                            <ListBoxItem x:Name="navDashboard" Content="Dashboard" IsSelected="True"/>
                            <ListBoxItem x:Name="navScan" Content="System Scan"/>
                            <ListBoxItem x:Name="navValidation" Content="Server Validation"/>
                            <ListBoxItem x:Name="navDecommission" Content="Server Decommission"/>
                            <ListBoxItem x:Name="navReports" Content="Reports"/>
                            <ListBoxItem x:Name="navSettings" Content="Settings"/>
                        </ListBox>
                        
                        <TextBlock Text="Migration Tools" FontSize="14" FontWeight="SemiBold" Margin="0,24,0,16" Foreground="#666666"/>
                        <ListBox x:Name="listMigrationTools" BorderThickness="0" Background="Transparent">
                            <ListBoxItem x:Name="toolServerUpgrade" Content="Server Upgrade"/>
                            <ListBoxItem x:Name="toolRoleMigration" Content="Role Migrations"/>
                            <ListBoxItem x:Name="toolRoleInstall" Content="Install Roles"/>
                            <ListBoxItem x:Name="toolDataMigration" Content="Data Migration"/>
                            <ListBoxItem x:Name="toolClusterUpgrade" Content="Cluster Upgrade"/>
                            <ListBoxItem x:Name="toolStorageMigration" Content="Storage Migration"/>
                        </ListBox>
                        
                        <TextBlock Text="Audit Tools" FontSize="14" FontWeight="SemiBold" Margin="0,24,0,16" Foreground="#666666"/>
                        <ListBox x:Name="listAuditTools" BorderThickness="0" Background="Transparent">
                            <ListBoxItem x:Name="auditServer" Content="Server Audit"/>
                            <ListBoxItem x:Name="auditRoles" Content="Role Audit"/>
                            <ListBoxItem x:Name="auditNetwork" Content="Network Audit"/>
                            <ListBoxItem x:Name="auditSecurity" Content="Security Audit"/>
                            <ListBoxItem x:Name="auditPerformance" Content="Performance Audit"/>
                            <ListBoxItem x:Name="auditCompliance" Content="Compliance Check"/>
                        </ListBox>
                    </StackPanel>
                </ScrollViewer>
            </Border>
            
            <!-- Content Area -->
            <ScrollViewer Grid.Column="1" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
                <Grid x:Name="contentArea" Margin="24">
                <!-- Home Panel -->
                <StackPanel x:Name="panelHome" Visibility="Visible">
                    <TextBlock Text="Welcome to the Windows Server Migration Tool" FontSize="24" FontWeight="Bold" Margin="0,0,0,20"/>
                    <TextBlock Text="This tool assists you in migrating Windows Server 2012/2016/2019 roles and services to new servers." FontSize="14" Margin="0,0,0,20" TextWrapping="Wrap"/>
                    
                    <Border Style="{StaticResource CardStyle}">
                        <StackPanel>
                            <TextBlock Text="Getting Started" FontSize="18" FontWeight="SemiBold" Margin="0,0,0,10"/>
                            <TextBlock Text="1. Perform a system analysis (System Scan)." Margin="0,5" TextWrapping="Wrap"/>
                            <TextBlock Text="2. Use the Audit Tools for a detailed inventory." Margin="0,5" TextWrapping="Wrap"/>
                            <TextBlock Text="3. Plan your migration using the Migration Tools." Margin="0,5" TextWrapping="Wrap"/>
                            <TextBlock Text="4. Validate the configuration after migration." Margin="0,5" TextWrapping="Wrap"/>
                            <TextBlock Text="5. Perform server decommissioning for legacy systems." Margin="0,5" TextWrapping="Wrap"/>
                        </StackPanel>
                    </Border>
                     <Border Style="{StaticResource CardStyle}">
                        <StackPanel>
                            <TextBlock Text="About This Tool" FontSize="18" FontWeight="SemiBold" Margin="0,0,0,10"/>
                            <TextBlock TextWrapping="Wrap">
                                The Easy Windows Server Migration Tool is designed to provide administrators with a central hub for planning, executing, and verifying Windows Server migrations. It combines analysis, audit, and migration tools to simplify the process and minimize errors.
                                <LineBreak/><LineBreak/>
                                Please note that this tool is an assistant. Careful planning and knowledge of your specific environment are essential for a successful migration. Always test all steps in a non-production environment first.
                            </TextBlock>
                        </StackPanel>
                    </Border>
                </StackPanel>
                
                <!-- Discovery Panel (System Scan) -->
                <StackPanel x:Name="panelDiscovery" Visibility="Collapsed">
                    <TextBlock Text="System Analysis (Scan)" FontSize="24" FontWeight="Bold" Margin="0,0,0,20"/>
                     <Border Style="{StaticResource CardStyle}">
                        <StackPanel>
                            <TextBlock Text="Gain System Overview" FontWeight="SemiBold" Margin="0,0,0,10"/>
                            <TextBlock Text="Start an analysis of your current server environment to identify domain controllers, FSMO roles, and important services. This information forms the basis for your migration planning." TextWrapping="Wrap" Margin="0,0,0,15"/>
                            <Button x:Name="btnStartDiscovery" Content="Start Analysis" Style="{StaticResource StandardButton}" HorizontalAlignment="Left"/>
                        </StackPanel>
                    </Border>
                    
                    <TabControl x:Name="tabDiscovery" Visibility="Collapsed" Margin="0,10,0,0">
                        <TabItem Header="Domain Controllers">
                            <DataGrid x:Name="dgDomainControllers" Margin="10"/>
                        </TabItem>
                        <TabItem Header="FSMO Roles">
                            <ListView x:Name="lvFSMORoles" Margin="10"/>
                        </TabItem>
                        <TabItem Header="Services">
                            <DataGrid x:Name="dgServices" Margin="10"/>
                        </TabItem>
                        <TabItem Header="Summary">
                            <TextBox x:Name="txtSummary" Margin="10" IsReadOnly="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" Height="200"/>
                        </TabItem>
                    </TabControl>
                    
                    <StackPanel Orientation="Horizontal" Margin="0,20,0,0" Visibility="Collapsed" x:Name="exportButtons">
                        <Button x:Name="btnExportHTML" Content="Export HTML" Style="{StaticResource StandardButton}" Margin="0,0,10,0"/>
                        <Button x:Name="btnPrintReport" Content="Print Report" Style="{StaticResource SecondaryButton}"/>
                    </StackPanel>
                </StackPanel>
                
                <!-- Installation Panel (Install Roles) -->
                <StackPanel x:Name="panelInstallation" Visibility="Collapsed">
                    <TextBlock Text="Role Installation" FontSize="24" FontWeight="Bold" Margin="0,0,0,20"/>
                    <Border Style="{StaticResource CardStyle}">
                        <StackPanel>
                            <TextBlock Text="Install Server Roles and Features" FontWeight="SemiBold" Margin="0,0,0,10"/>
                            <TextBlock Text="Install new server roles and features on local or remote servers. This is often a preparatory step for a side-by-side migration or setting up a new server." TextWrapping="Wrap" Margin="0,0,0,15"/>
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>
                                
                                <TextBlock Grid.Row="0" Grid.Column="0" Text="Target Server:" Margin="0,5,10,5" VerticalAlignment="Center"/>
                                <TextBox Grid.Row="0" Grid.Column="1" x:Name="txtInstallationServer" Text="localhost" Margin="0,5"/>
                                
                                <TextBlock Grid.Row="1" Grid.Column="0" Text="Available Roles:" Margin="0,5,10,5" VerticalAlignment="Top"/>
                                <ScrollViewer Grid.Row="1" Grid.Column="1" Height="150" Margin="0,5" VerticalScrollBarVisibility="Auto">
                                    <ListBox x:Name="listAvailableRoles" SelectionMode="Multiple" Background="Transparent" BorderThickness="0">
                                        <ListBoxItem Content="Active Directory Domain Services" Tag="ADDomainServices"/>
                                        <ListBoxItem Content="DNS Server" Tag="DNSServer"/>
                                        <ListBoxItem Content="DHCP Server" Tag="DHCPServer"/>
                                        <ListBoxItem Content="File and Storage Services" Tag="FileServices"/>
                                        <ListBoxItem Content="Web Server (IIS)" Tag="WebServerIIS"/>
                                        <ListBoxItem Content="Windows Server Update Services" Tag="UpdateServices"/>
                                        <ListBoxItem Content="Hyper V" Tag="HyperV"/>
                                        <ListBoxItem Content="Remote Desktop Services" Tag="RDSServices"/>
                                        <ListBoxItem Content="Print and Document Services" Tag="PrintServices"/>
                                        <ListBoxItem Content="Network Policy and Access Services" Tag="NPAS"/>
                                        <ListBoxItem Content="Remote Access" Tag="RemoteAccess"/>
                                    </ListBox>
                                </ScrollViewer>
                                
                                <Button Grid.Row="2" Grid.Column="1" x:Name="btnStartInstallation" Content="Start Installation" Style="{StaticResource StandardButton}" Margin="0,15,0,0"/>
                            </Grid>
                        </StackPanel>
                    </Border>
                    
                    <TextBlock x:Name="txtInstallationStatus" Text="Ready for Installation" Margin="0,10,0,0"/>
                    <ProgressBar x:Name="progressInstallation" Visibility="Collapsed" Height="20" Margin="0,10,0,0"/>
                    
                    <Border x:Name="installationResults" Style="{StaticResource CardStyle}" Visibility="Collapsed" Margin="0,10,0,0">
                        <StackPanel>
                            <TextBlock Text="Installation Results" FontWeight="SemiBold" Margin="0,0,0,10"/>
                            <DataGrid x:Name="dgInstallationResults" MaxHeight="200" VerticalScrollBarVisibility="Auto"/>
                        </StackPanel>
                    </Border>
                </StackPanel>
                
                <!-- Data Migration Panel -->
                <StackPanel x:Name="panelDataMigration" Visibility="Collapsed">
                    <TextBlock Text="Data Migration" FontSize="24" FontWeight="Bold" Margin="0,0,0,20"/>
                    <Border Style="{StaticResource CardStyle}">
                        <StackPanel>
                            <TextBlock Text="Data and Configuration Migration" FontWeight="SemiBold" Margin="0,0,0,10"/>
                            <TextBlock Text="Migrate specific data and configurations between servers. This can be part of a side-by-side migration or used for transferring individual datasets." TextWrapping="Wrap" Margin="0,0,0,15"/>
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                
                                <StackPanel Grid.Column="0" Margin="0,0,10,0">
                                    <TextBlock Text="Available Migration Types:" FontWeight="SemiBold" Margin="0,0,0,10"/>
                                    <CheckBox x:Name="chkMigrateUserData" Content="User Data" Margin="0,5"/>
                                    <CheckBox x:Name="chkMigrateShares" Content="File Shares" Margin="0,5"/>
                                    <CheckBox x:Name="chkMigratePrinters" Content="Printers" Margin="0,5"/>
                                    <CheckBox x:Name="chkMigrateRegistry" Content="Registry Settings (Selection)" Margin="0,5"/>
                                    <CheckBox x:Name="chkMigrateScheduledTasks" Content="Scheduled Tasks" Margin="0,5"/>
                                </StackPanel>
                                
                                <StackPanel Grid.Column="1" Margin="10,0,0,0">
                                    <TextBlock Text="Server Configuration:" FontWeight="SemiBold" Margin="0,0,0,10"/>
                                    <TextBlock Text="Source Server:" Margin="0,5,0,2"/>
                                    <TextBox x:Name="txtMigrationSource" Margin="0,0,0,10"/>
                                    <TextBlock Text="Target Server:" Margin="0,5,0,2"/>
                                    <TextBox x:Name="txtMigrationTarget" Margin="0,0,0,10"/>
                                    <Button x:Name="btnStartDataMigration" Content="Start Migration" Style="{StaticResource StandardButton}" Margin="0,10,0,0"/>
                                </StackPanel>
                            </Grid>
                        </StackPanel>
                    </Border>
                    <TextBlock x:Name="txtDataMigrationStatus" Visibility="Collapsed" Margin="0,10,0,0"/>
                    <ProgressBar x:Name="progressDataMigration" Visibility="Collapsed" Height="20" Margin="0,10,0,0"/>
                    <Border x:Name="dataMigrationResults" Style="{StaticResource CardStyle}" Visibility="Collapsed" Margin="0,10,0,0">
                        <StackPanel>
                            <TextBlock Text="Migration Results" FontWeight="SemiBold" Margin="0,0,0,10"/>
                            <DataGrid x:Name="dgDataMigrationResults" MaxHeight="200" VerticalScrollBarVisibility="Auto"/>
                        </StackPanel>
                    </Border>
                </StackPanel>
                
                <!-- Validation Panel -->
                <StackPanel x:Name="panelValidation" Visibility="Collapsed">
                    <TextBlock Text="Server Validation" FontSize="24" FontWeight="Bold" Margin="0,0,0,20"/>
                    <Border Style="{StaticResource CardStyle}">
                        <StackPanel>
                             <TextBlock Text="Verify Server Configuration" FontWeight="SemiBold" Margin="0,0,0,10"/>
                             <TextBlock Text="Perform validation after a migration or major configuration change to ensure all services and roles are functioning correctly." TextWrapping="Wrap" Margin="0,0,0,15"/>
                            <TextBlock Text="Target Server for Validation:" Margin="0,0,0,2"/>
                            <TextBox x:Name="txtValidationServerName" Text="localhost" Margin="0,0,0,10"/>
                            
                            <TextBlock Text="Validation Options" FontWeight="SemiBold" Margin="0,10,0,10"/>
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                <StackPanel Grid.Column="0">
                                    <CheckBox x:Name="chkValidateBasics" Content="Basic Tests" IsChecked="True" Margin="0,5"/>
                                    <CheckBox x:Name="chkValidateAD" Content="Active Directory" IsChecked="True" Margin="0,5"/>
                                    <CheckBox x:Name="chkValidateDNS" Content="DNS Service" IsChecked="True" Margin="0,5"/>
                                    <CheckBox x:Name="chkValidateDHCP" Content="DHCP Service" IsChecked="True" Margin="0,5"/>
                                </StackPanel>
                                <StackPanel Grid.Column="1">
                                    <CheckBox x:Name="chkValidateRoles" Content="Server Roles Status" IsChecked="True" Margin="0,5"/>
                                    <CheckBox x:Name="chkValidateGPO" Content="Group Policies" IsChecked="True" Margin="0,5"/>
                                    <CheckBox x:Name="chkValidatePerf" Content="Performance Indicators" IsChecked="True" Margin="0,5"/>
                                    <CheckBox x:Name="chkValidateServicesState" Content="Key Services State" IsChecked="True" Margin="0,5"/>
                                </StackPanel>
                            </Grid>
                            <Button x:Name="btnStartValidation" Content="Start Validation" Style="{StaticResource StandardButton}" Margin="0,20,0,0"/>
                        </StackPanel>
                    </Border>
                    
                    <TextBlock x:Name="txtValidationStatus" Visibility="Collapsed" Margin="0,10,0,0"/>
                    <ProgressBar x:Name="progressValidation" Visibility="Collapsed" Height="20" Margin="0,10,0,0"/>
                    
                    <TabControl x:Name="tabValidationResults" Visibility="Collapsed" Margin="0,10,0,0">
                        <TabItem Header="Summary">
                            <ScrollViewer VerticalScrollBarVisibility="Auto" MaxHeight="300">
                                <DataGrid x:Name="dgValidationSummary" Margin="10"/>
                            </ScrollViewer>
                        </TabItem>
                        <TabItem Header="Active Directory">
                            <ScrollViewer VerticalScrollBarVisibility="Auto" MaxHeight="400">
                                <StackPanel Margin="10">
                                    <TextBlock Text="AD Services" FontWeight="Bold" Margin="0,0,0,5"/>
                                    <ScrollViewer VerticalScrollBarVisibility="Auto" MaxHeight="200">
                                        <DataGrid x:Name="dgValidationAD"/>
                                    </ScrollViewer>
                                    <TextBlock Text="AD Replication" FontWeight="Bold" Margin="0,10,0,5"/>
                                    <ScrollViewer VerticalScrollBarVisibility="Auto" MaxHeight="200">
                                        <DataGrid x:Name="dgValidationReplication"/>
                                    </ScrollViewer>
                                    <TextBlock Text="FSMO Roles" FontWeight="Bold" Margin="0,10,0,5"/>
                                    <ScrollViewer VerticalScrollBarVisibility="Auto" MaxHeight="200">
                                        <DataGrid x:Name="dgValidationFSMO"/>
                                    </ScrollViewer>
                                </StackPanel>
                            </ScrollViewer>
                        </TabItem>
                        <TabItem Header="DNS">
                             <ScrollViewer VerticalScrollBarVisibility="Auto" MaxHeight="400">
                                <StackPanel Margin="10">
                                    <TextBlock Text="DNS Service Status" FontWeight="Bold" Margin="0,0,0,5"/>
                                    <ScrollViewer VerticalScrollBarVisibility="Auto" MaxHeight="200">
                                        <DataGrid x:Name="dgValidationDNSService"/>
                                    </ScrollViewer>
                                    <TextBlock Text="DNS Zones" FontWeight="Bold" Margin="0,10,0,5"/>
                                    <ScrollViewer VerticalScrollBarVisibility="Auto" MaxHeight="200">
                                        <DataGrid x:Name="dgValidationDNSZones"/>
                                    </ScrollViewer>
                                    <TextBlock Text="DNS Resolution Tests" FontWeight="Bold" Margin="0,10,0,5"/>
                                    <ScrollViewer VerticalScrollBarVisibility="Auto" MaxHeight="200">
                                        <DataGrid x:Name="dgValidationDNSResolution"/>
                                    </ScrollViewer>
                                </StackPanel>
                            </ScrollViewer>
                        </TabItem>
                        <TabItem Header="DHCP">
                            <ScrollViewer VerticalScrollBarVisibility="Auto" MaxHeight="400">
                                <StackPanel Margin="10">
                                    <TextBlock Text="DHCP Service Status" FontWeight="Bold" Margin="0,0,0,5"/>
                                    <ScrollViewer VerticalScrollBarVisibility="Auto" MaxHeight="200">
                                        <DataGrid x:Name="dgValidationDHCPService"/>
                                    </ScrollViewer>
                                    <TextBlock Text="DHCP Scopes" FontWeight="Bold" Margin="0,10,0,5"/>
                                    <ScrollViewer VerticalScrollBarVisibility="Auto" MaxHeight="200">
                                        <DataGrid x:Name="dgValidationDHCPScopes"/>
                                    </ScrollViewer>
                                    <TextBlock Text="DHCP Failover" FontWeight="Bold" Margin="0,10,0,5"/>
                                    <ScrollViewer VerticalScrollBarVisibility="Auto" MaxHeight="200">
                                        <DataGrid x:Name="dgValidationDHCPFailover"/>
                                    </ScrollViewer>
                                </StackPanel>
                            </ScrollViewer>
                        </TabItem>
                        <TabItem Header="Services">
                            <ScrollViewer VerticalScrollBarVisibility="Auto" MaxHeight="300">
                                <DataGrid x:Name="dgValidationServices" Margin="10"/>
                            </ScrollViewer>
                        </TabItem>
                        <TabItem Header="Performance">
                             <ScrollViewer VerticalScrollBarVisibility="Auto" MaxHeight="400">
                                <StackPanel Margin="10">
                                    <TextBlock Text="Performance Metrics" FontWeight="Bold" Margin="0,0,0,5"/>
                                    <ScrollViewer VerticalScrollBarVisibility="Auto" MaxHeight="200">
                                        <DataGrid x:Name="dgValidationPerformance"/>
                                    </ScrollViewer>
                                    <TextBlock Text="Key Event Log Entries" FontWeight="Bold" Margin="0,10,0,5"/>
                                    <ScrollViewer VerticalScrollBarVisibility="Auto" MaxHeight="200">
                                        <DataGrid x:Name="dgValidationEventLogs"/>
                                    </ScrollViewer>
                                </StackPanel>
                            </ScrollViewer>
                        </TabItem>
                        <TabItem Header="Group Policies">
                            <ScrollViewer VerticalScrollBarVisibility="Auto" MaxHeight="400">
                                <StackPanel Margin="10">
                                    <TextBlock Text="GPO Replication" FontWeight="Bold" Margin="0,0,0,5"/>
                                    <ScrollViewer VerticalScrollBarVisibility="Auto" MaxHeight="200">
                                        <DataGrid x:Name="dgValidationGPO"/>
                                    </ScrollViewer>
                                    <TextBlock Text="SYSVOL Replication" FontWeight="Bold" Margin="0,10,0,5"/>
                                    <ScrollViewer VerticalScrollBarVisibility="Auto" MaxHeight="200">
                                        <DataGrid x:Name="dgValidationSYSVOL"/>
                                    </ScrollViewer>
                                </StackPanel>
                            </ScrollViewer>
                        </TabItem>
                        <TabItem Header="Validation Report">
                            <TextBox x:Name="txtValidationReport" Margin="10" IsReadOnly="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" Height="300"/>
                        </TabItem>
                    </TabControl>
                    
                    <StackPanel Orientation="Horizontal" Margin="0,20,0,0" Visibility="Collapsed" x:Name="validationExportButtons">
                        <Button x:Name="btnSaveValidationHTML" Content="Export HTML" Style="{StaticResource StandardButton}" Margin="0,0,10,0"/>
                        <Button x:Name="btnPrintValidation" Content="Print" Style="{StaticResource SecondaryButton}"/>
                    </StackPanel>
                </StackPanel>
                
                <!-- Server Audit Panel -->
                <StackPanel x:Name="panelAuditServer" Visibility="Collapsed">
                    <TextBlock Text="Server Audit" FontSize="24" FontWeight="Bold" Margin="0,0,0,20"/>
                    <Border Style="{StaticResource CardStyle}">
                        <StackPanel>
                            <TextBlock Text="Comprehensive Server Audit" FontWeight="SemiBold" Margin="0,0,0,10"/>
                            <TextBlock Text="Perform a detailed analysis of your Windows Server to identify configurations, security settings, and potential issues. This is helpful for inventory and before major changes." TextWrapping="Wrap" Margin="0,0,0,15"/>
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                <StackPanel Grid.Column="0" Margin="0,0,10,0">
                                    <CheckBox x:Name="chkAuditSystem" Content="System Information" IsChecked="True" Margin="0,5"/>
                                    <CheckBox x:Name="chkAuditNetworkConfig" Content="Network Configuration" IsChecked="True" Margin="0,5"/>
                                    <CheckBox x:Name="chkAuditSecuritySettings" Content="Security Settings" IsChecked="True" Margin="0,5"/>
                                    <CheckBox x:Name="chkAuditServicesAndRoles" Content="Services and Roles" IsChecked="True" Margin="0,5"/>
                                    <CheckBox x:Name="chkAuditActiveDirectoryState" Content="Active Directory Status" IsChecked="True" Margin="0,5"/>
                                </StackPanel>
                                <StackPanel Grid.Column="1" Margin="10,0,0,0">
                                    <CheckBox x:Name="chkAuditDNSConfig" Content="DNS Configuration" IsChecked="True" Margin="0,5"/>
                                    <CheckBox x:Name="chkAuditDHCPConfig" Content="DHCP Configuration" IsChecked="True" Margin="0,5"/>
                                    <CheckBox x:Name="chkAuditIISConfig" Content="IIS Web Server" IsChecked="False" Margin="0,5"/>
                                    <CheckBox x:Name="chkAuditHyperVConfig" Content="Hyper V" IsChecked="False" Margin="0,5"/>
                                    <CheckBox x:Name="chkAuditConnections" Content="Active Network Connections" IsChecked="True" Margin="0,5"/>
                                </StackPanel>
                            </Grid>
                            <Button x:Name="btnStartAudit" Content="Start Audit" Style="{StaticResource StandardButton}" Margin="0,20,0,0"/>
                        </StackPanel>
                    </Border>
                    <TextBlock x:Name="txtAuditStatus" Visibility="Collapsed" Margin="0,10,0,0"/>
                    <ProgressBar x:Name="progressAudit" Visibility="Collapsed" Height="20" Margin="0,10,0,0"/>
                    <TabControl x:Name="tabAuditResults" Visibility="Collapsed" Margin="0,10,0,0">
                        <TabItem Header="System">
                            <ScrollViewer VerticalScrollBarVisibility="Auto" MaxHeight="300"><TextBox x:Name="txtAuditSystem" IsReadOnly="True" TextWrapping="Wrap" FontFamily="Consolas" FontSize="10" Margin="5"/></ScrollViewer>
                        </TabItem>
                        <TabItem Header="Network">
                            <ScrollViewer VerticalScrollBarVisibility="Auto" MaxHeight="300"><TextBox x:Name="txtAuditNetwork" IsReadOnly="True" TextWrapping="Wrap" FontFamily="Consolas" FontSize="10" Margin="5"/></ScrollViewer>
                        </TabItem>
                        <TabItem Header="Security">
                            <ScrollViewer VerticalScrollBarVisibility="Auto" MaxHeight="300"><TextBox x:Name="txtAuditSecurity" IsReadOnly="True" TextWrapping="Wrap" FontFamily="Consolas" FontSize="10" Margin="5"/></ScrollViewer>
                        </TabItem>
                        <TabItem Header="Connections">
                            <ScrollViewer VerticalScrollBarVisibility="Auto" MaxHeight="300"><TextBox x:Name="txtAuditConnections" IsReadOnly="True" TextWrapping="Wrap" FontFamily="Consolas" FontSize="10" Margin="5"/></ScrollViewer>
                        </TabItem>
                        <TabItem Header="Summary">
                            <ScrollViewer VerticalScrollBarVisibility="Auto" MaxHeight="300"><TextBox x:Name="txtAuditSummary" IsReadOnly="True" TextWrapping="Wrap" FontFamily="Consolas" FontSize="10" Margin="5"/></ScrollViewer>
                        </TabItem>
                    </TabControl>
                    <StackPanel Orientation="Horizontal" Margin="0,20,0,0" Visibility="Collapsed" x:Name="auditExportButtons">
                        <Button x:Name="btnExportAuditHTML" Content="Export HTML" Style="{StaticResource StandardButton}" Margin="0,0,10,0"/>
                        <Button x:Name="btnExportAuditDrawIO" Content="Network Topology (Draw.IO)" Style="{StaticResource StandardButton}" Margin="0,0,10,0"/>
                        <Button x:Name="btnPrintAudit" Content="Print" Style="{StaticResource SecondaryButton}"/>
                    </StackPanel>
                </StackPanel>

                <!-- Role Audit Panel -->
                <StackPanel x:Name="panelAuditRoles" Visibility="Collapsed">
                    <TextBlock Text="Role Audit" FontSize="24" FontWeight="Bold" Margin="0,0,0,20"/>
                    <Border Style="{StaticResource CardStyle}">
                        <StackPanel>
                            <TextBlock Text="Server Role Configuration Audit" FontWeight="SemiBold" Margin="0,0,0,10"/>
                            <TextBlock Text="Audit the configuration of installed server roles to ensure they are set up according to best practices and identify potential misconfigurations." TextWrapping="Wrap" Margin="0,0,0,15"/>
                            <TextBlock Text="Target Server for Role Audit:" Margin="0,0,0,2"/>
                            <TextBox x:Name="txtRoleAuditServerName" Text="localhost" Margin="0,0,0,10"/>
                            <TextBlock Text="Select Roles to Audit:" Margin="0,10,0,5"/>
                            <ListBox x:Name="listRolesToAudit" Height="150" SelectionMode="Multiple" Margin="0,5,0,10">
                                <ListBoxItem Content="Active Directory Domain Services" Tag="ADDomainServices"/>
                                <ListBoxItem Content="DNS Server" Tag="DNSServer"/>
                                <ListBoxItem Content="DHCP Server" Tag="DHCPServer"/>
                                <ListBoxItem Content="File and Storage Services" Tag="FileServices"/>
                                <ListBoxItem Content="Web Server (IIS)" Tag="WebServerIIS"/>
                                <ListBoxItem Content="Windows Server Update Services" Tag="UpdateServices"/>
                                <ListBoxItem Content="Hyper V" Tag="HyperV"/>
                                <ListBoxItem Content="Remote Desktop Services" Tag="RDSServices"/>
                                <ListBoxItem Content="Print and Document Services" Tag="PrintServices"/>
                                <ListBoxItem Content="Network Policy and Access Services" Tag="NPAS"/>
                            </ListBox>
                            <Button x:Name="btnStartRoleAudit" Content="Start Role Audit" Style="{StaticResource StandardButton}" Margin="0,15,0,0"/>
                        </StackPanel>
                    </Border>
                    <TextBlock x:Name="txtRoleAuditStatus" Visibility="Collapsed" Margin="0,10,0,0"/>
                    <ProgressBar x:Name="progressRoleAudit" Visibility="Collapsed" Height="20" Margin="0,10,0,0"/>
                    <Border x:Name="roleAuditResultsCard" Style="{StaticResource CardStyle}" Visibility="Collapsed" Margin="0,10,0,0">
                        <StackPanel>
                            <TextBlock Text="Role Audit Results" FontWeight="SemiBold" Margin="0,0,0,10"/>
                            <DataGrid x:Name="dgRoleAuditResults" MaxHeight="300" VerticalScrollBarVisibility="Auto"/>
                             <TextBlock Text="Detailed findings will be presented here, specific to each audited role." FontStyle="Italic" Margin="0,10,0,0"/>
                        </StackPanel>
                    </Border>
                </StackPanel>

                <!-- Server Upgrade Panel -->
                <StackPanel x:Name="panelServerUpgrade" Visibility="Collapsed">
                    <TextBlock Text="Server Upgrade Strategies" FontSize="24" FontWeight="Bold" Margin="0,0,0,20"/>
                    <Border Style="{StaticResource CardStyle}">
                        <StackPanel>
                            <TextBlock Text="Windows Server Upgrade Options" FontWeight="SemiBold" Margin="0,0,0,10"/>
                            <TextBlock Text="Choose the appropriate upgrade strategy for your environment. Each method has pros and cons that must be carefully considered." TextWrapping="Wrap" Margin="0,0,0,15"/>
                            <TextBlock Text="Note that direct upgrades across multiple versions (e.g., 2012 R2 to 2022) are often not directly supported and may require intermediate steps (e.g., to 2019) or a side-by-side migration might be the better choice." TextWrapping="Wrap" FontStyle="Italic" Foreground="DarkRed"/>
                        </StackPanel>
                    </Border>
                    <Border Style="{StaticResource CardStyle}">
                        <StackPanel>
                            <TextBlock Text="In-Place Upgrade" FontWeight="SemiBold" Margin="0,0,0,5"/>
                            <TextBlock Text="Update the operating system on the same hardware. This method preserves existing settings, roles, and data, but carries higher risks and is not recommended or possible for all scenarios." TextWrapping="Wrap" Margin="0,0,0,10" Foreground="#444444"/>
                            <TextBlock Text="Possible direct paths (examples):" Margin="0,0,0,5"/>
                            <TextBlock Text="* Windows Server 2012 R2 to Windows Server 2019" FontSize="12" Margin="15,2,0,2"/>
                            <TextBlock Text="* Windows Server 2016 to Windows Server 2019 / 2022" FontSize="12" Margin="15,2,0,2"/>
                            <TextBlock Text="* Windows Server 2019 to Windows Server 2022" FontSize="12" Margin="15,2,0,2"/>
                            <Button x:Name="btnInPlaceUpgrade" Content="Prepare In-Place Upgrade" Style="{StaticResource StandardButton}" Margin="0,10,0,0"/>
                            <TextBlock Text="Note: An in-place upgrade is complex and requires careful planning, compatibility checks (e.g., with Setup.exe /ScanOnly), and full backups. Not recommended for Domain Controllers." Margin="0,10,0,0" FontStyle="Italic" FontSize="12" TextWrapping="Wrap"/>
                        </StackPanel>
                    </Border>
                    <Border Style="{StaticResource CardStyle}">
                        <StackPanel>
                             <TextBlock Text="Side-by-Side Migration (Swing Migration)" FontWeight="SemiBold" Margin="0,0,0,5"/>
                             <TextBlock Text="Migrate roles and data to a new server with a fresh operating system installation. This is often the safer, more flexible, and recommended method, especially for critical roles like Domain Controllers." TextWrapping="Wrap" Margin="0,0,0,10" Foreground="#444444"/>
                             <TextBlock Text="Advantages:" Margin="0,0,0,5"/>
                             <TextBlock Text="* Minimal downtime for the old server during preparation." FontSize="12" Margin="15,2,0,2"/>
                             <TextBlock Text="* Clean installation of the new OS, no legacy issues." FontSize="12" Margin="15,2,0,2"/>
                             <TextBlock Text="* Opportunity for hardware refresh and virtualization." FontSize="12" Margin="15,2,0,2"/>
                             <TextBlock Text="* Easier rollback in case of issues." FontSize="12" Margin="15,2,0,2"/>
                             <Button x:Name="btnSideBySideMigrationInfo" Content="Side-by-Side Migration Info" Style="{StaticResource StandardButton}" Margin="0,10,0,0"/>
                             <TextBlock Text="Use the 'Role Migrations' and 'Data Migration' tools for execution." Margin="0,10,0,0" FontStyle="Italic" FontSize="12" TextWrapping="Wrap"/>
                        </StackPanel>
                    </Border>
                </StackPanel>

                <!-- Role Migration Panel -->
                <StackPanel x:Name="panelRoleMigration" Visibility="Collapsed">
                    <TextBlock Text="Role Migrations" FontSize="24" FontWeight="Bold" Margin="0,0,0,20"/>
                    <Border Style="{StaticResource CardStyle}">
                        <TextBlock Text="Select a server role to view specific migration guides, tools, and considerations. Server role migration is typically done as part of a side-by-side migration to a new server." TextWrapping="Wrap" Margin="0,0,0,10"/>
                    </Border>
                    <TabControl x:Name="tabRoleMigration">
                        <TabItem Header="Active Directory">
                            <ScrollViewer VerticalScrollBarVisibility="Auto" MaxHeight="550">
                                <StackPanel Margin="10">
                                    <TextBlock Text="Active Directory Domain Services Migration" FontSize="18" FontWeight="SemiBold" Margin="0,0,0,15"/>
                                    <Border Style="{StaticResource CardStyle}">
                                        <StackPanel>
                                            <TextBlock Text="AD DS Migration Overview (Side-by-Side)" FontWeight="SemiBold" Margin="0,0,0,10"/>
                                            <TextBlock TextWrapping="Wrap" Margin="0,0,0,10">
                                                Migrating AD DS involves adding new Domain Controllers (DCs) with a newer Windows Server version to the existing domain, transferring FSMO roles, ensuring replication, and then demoting and removing the old DCs.
                                            </TextBlock>
                                            <TextBlock Text="Key Steps:" FontWeight="Normal" Margin="0,5,0,5"/>
                                            <TextBlock Text="1. Preparation: Ensure a stable AD/DNS environment. Run adprep /forestprep and adprep /domainprep if upgrading from very old environments (usually not needed for 2012R2 schema extensions by new DCs)." Margin="10,2,0,2" TextWrapping="Wrap"/>
                                            <TextBlock Text="2. New Server: Install OS, configure (IP, Name), join to the domain." Margin="10,2,0,2" TextWrapping="Wrap"/>
                                            <TextBlock Text="3. Install AD DS Role on the new server." Margin="10,2,0,2" TextWrapping="Wrap"/>
                                            <TextBlock Text="4. Promote the new server to an additional Domain Controller in the existing domain." Margin="10,2,0,2" TextWrapping="Wrap"/>
                                            <TextBlock Text="5. Verification: Check replication (`repadmin /showrepl`, `repadmin /replsummary`), health (`dcdiag /v`)." Margin="10,2,0,2" TextWrapping="Wrap"/>
                                            <TextBlock Text="6. FSMO Roles: Transfer all 5 FSMO roles to the new DC (or another suitable new DC)." Margin="10,2,0,2" TextWrapping="Wrap"/>
                                            <TextBlock Text="7. DNS Clients: Update DNS settings of clients and member servers to use the new DCs (often via DHCP)." Margin="10,2,0,2" TextWrapping="Wrap"/>
                                            <TextBlock Text="8. Global Catalog: Ensure the new DC is also a Global Catalog server (default)." Margin="10,2,0,2" TextWrapping="Wrap"/>
                                            <TextBlock Text="9. Observation Phase: Monitor the environment for a few days/weeks." Margin="10,2,0,2" TextWrapping="Wrap"/>
                                            <TextBlock Text="10. Old DCs: Demote and remove the old Domain Controllers from the domain." Margin="10,2,0,2" TextWrapping="Wrap"/>
                                            <TextBlock Text="11. Functional Levels: Raise domain and forest functional levels once all DCs meet minimum requirements." Margin="10,2,0,2" TextWrapping="Wrap"/>
                                            <Button x:Name="btnADSmigTool" Content="Windows Server Migration Tools (SMIG) for AD DS (Legacy)" Style="{StaticResource SecondaryButton}" Margin="0,10,0,0" IsEnabled="False" ToolTip="SMIG is no longer recommended for AD DS. Manual steps or PowerShell are preferred."/>
                                        </StackPanel>
                                    </Border>
                                </StackPanel>
                            </ScrollViewer>
                        </TabItem>
                        <TabItem Header="DNS">
                             <ScrollViewer VerticalScrollBarVisibility="Auto" MaxHeight="550">
                                <StackPanel Margin="10">
                                    <TextBlock Text="DNS Server Migration" FontSize="18" FontWeight="SemiBold" Margin="0,0,0,15"/>
                                    <Border Style="{StaticResource CardStyle}">
                                        <StackPanel>
                                            <TextBlock Text="DNS Migration Overview" FontWeight="SemiBold" Margin="0,0,0,10"/>
                                            <TextBlock TextWrapping="Wrap" Margin="0,0,0,10">
                                                If DNS is hosted on Domain Controllers (AD-integrated zones), DNS migrates automatically with AD DS migration. For standalone DNS servers or file-based zones, separate steps are required.
                                            </TextBlock>
                                            <TextBlock Text="AD-Integrated Zones:" FontWeight="Normal" Margin="0,5,0,5"/>
                                            <TextBlock Text="* Are automatically replicated to new DCs that are also DNS servers." Margin="10,2,0,2" TextWrapping="Wrap"/>
                                            <TextBlock Text="* Ensure new DCs are configured as DNS servers in client IP settings." Margin="10,2,0,2" TextWrapping="Wrap"/>
                                            <TextBlock Text="File-Based Zones (Primary/Secondary):" FontWeight="Normal" Margin="0,5,0,5"/>
                                            <TextBlock Text="1. New Server: Install DNS role." Margin="10,2,0,2" TextWrapping="Wrap"/>
                                            <TextBlock Text="2. Transfer Zones: Copy zone files (usually in `%windir%\System32\Dns`) from old to new server. Alternatively: configure as secondary, perform zone transfer, then change to primary." Margin="10,2,0,2" TextWrapping="Wrap"/>
                                            <TextBlock Text="3. DNS Server Settings (Forwarders, Root Hints, etc.): Manually configure on new server or export/import via PowerShell." Margin="10,2,0,2" TextWrapping="Wrap"/>
                                            <TextBlock Text="4. Change IP address of old DNS server to new DNS server, or update clients to new IP." Margin="10,2,0,2" TextWrapping="Wrap"/>
                                            <Button x:Name="btnDnsExportImport" Content="DNS Zones Export/Import (PowerShell)" Style="{StaticResource StandardButton}" Margin="0,10,0,0"/>
                                        </StackPanel>
                                    </Border>
                                </StackPanel>
                            </ScrollViewer>
                        </TabItem>
                        <TabItem Header="DHCP">
                            <ScrollViewer VerticalScrollBarVisibility="Auto" MaxHeight="550">
                                <StackPanel Margin="10">
                                    <TextBlock Text="DHCP Server Migration" FontSize="18" FontWeight="SemiBold" Margin="0,0,0,15"/>
                                    <Border Style="{StaticResource CardStyle}">
                                        <StackPanel>
                                            <TextBlock Text="DHCP Migration Overview" FontWeight="SemiBold" Margin="0,0,0,10"/>
                                            <TextBlock TextWrapping="Wrap" Margin="0,0,0,10">
                                                DHCP server configuration (scopes, reservations, options) can be exported and imported to a new server. Failover configurations require special attention.
                                            </TextBlock>
                                            <TextBlock Text="Migration Steps (Netsh or PowerShell):" FontWeight="Normal" Margin="0,5,0,5"/>
                                            <TextBlock Text="1. Old Server: Export DHCP configuration." Margin="10,2,0,2" TextWrapping="Wrap"/>
                                            <TextBlock Text="   `netsh dhcp server export C:\DHCPbackup\dhcpcfg.txt all`" FontFamily="Consolas" FontSize="11" Margin="20,0,0,0"/>
                                            <TextBlock Text="   `Export-DhcpServer -File C:\DHCPbackup\dhcpcfg.xml -ComputerName oldDHCPServer` (PS)" FontFamily="Consolas" FontSize="11" Margin="20,0,0,0"/>
                                            <TextBlock Text="2. New Server: Install DHCP role." Margin="10,2,0,2" TextWrapping="Wrap"/>
                                            <TextBlock Text="3. New Server: Stop DHCP service." Margin="10,2,0,2" TextWrapping="Wrap"/>
                                            <TextBlock Text="4. New Server: Import exported configuration." Margin="10,2,0,2" TextWrapping="Wrap"/>
                                            <TextBlock Text="   `netsh dhcp server import C:\DHCPbackup\dhcpcfg.txt all`" FontFamily="Consolas" FontSize="11" Margin="20,0,0,0"/>
                                            <TextBlock Text="   `Import-DhcpServer -File C:\DHCPbackup\dhcpcfg.xml -ComputerName newDHCPServer -BackupPath C:\DHCPbackup\backup\` (PS)" FontFamily="Consolas" FontSize="11" Margin="20,0,0,0"/>
                                            <TextBlock Text="5. New Server: Start DHCP service and authorize (if in AD environment)." Margin="10,2,0,2" TextWrapping="Wrap"/>
                                            <TextBlock Text="6. Old Server: De-authorize and stop/uninstall DHCP service." Margin="10,2,0,2" TextWrapping="Wrap"/>
                                            <TextBlock Text="7. DHCP Failover: If present, remove failover relationship on old server and reconfigure on new server." Margin="10,2,0,2" TextWrapping="Wrap"/>
                                            <Button x:Name="btnDhcpMigrate" Content="Start DHCP Migration with PowerShell" Style="{StaticResource StandardButton}" Margin="0,10,0,0"/>
                                        </StackPanel>
                                    </Border>
                                </StackPanel>
                            </ScrollViewer>
                        </TabItem>
                        <TabItem Header="File Services">
                             <ScrollViewer VerticalScrollBarVisibility="Auto" MaxHeight="550">
                                <StackPanel Margin="10">
                                    <TextBlock Text="File Services Migration" FontSize="18" FontWeight="SemiBold" Margin="0,0,0,15"/>
                                    <Border Style="{StaticResource CardStyle}">
                                        <StackPanel>
                                            <TextBlock Text="File Services Migration Overview" FontWeight="SemiBold" Margin="0,0,0,10"/>
                                            <TextBlock TextWrapping="Wrap" Margin="0,0,0,10">
                                                Migrating file services involves transferring files, folders, shares, and permissions. Storage Migration Service (SMS) is the recommended tool for Server 2012 R2 and newer. Robocopy is a robust alternative.
                                            </TextBlock>
                                            <TextBlock Text="Storage Migration Service (SMS):" FontWeight="Normal" Margin="0,5,0,5"/>
                                            <TextBlock Text="* Orchestrates migration from one or more source servers to new target servers (physical or virtual, including Azure)." Margin="10,2,0,2" TextWrapping="Wrap"/>
                                            <TextBlock Text="* Inventory, transfer, and cutover (source server identity takeover)." Margin="10,2,0,2" TextWrapping="Wrap"/>
                                            <TextBlock Text="* Requires an orchestration server (Win Server 2019/2022)." Margin="10,2,0,2" TextWrapping="Wrap"/>
                                            <Button x:Name="btnOpenStorageMigrationTool" Content="Storage Migration Service (in Menu)" Style="{StaticResource SecondaryButton}" Margin="0,10,0,5"/>

                                            <TextBlock Text="Robocopy:" FontWeight="Normal" Margin="0,10,0,5"/>
                                            <TextBlock Text="* Powerful command-line tool for copying files and folders while preserving permissions and attributes." Margin="10,2,0,2" TextWrapping="Wrap"/>
                                            <TextBlock Text="* Example: `robocopy \\SourceServer\Share D:\LocalPathOnTarget /E /COPYALL /R:3 /W:5 /LOG:C:\RoboLog.txt`" FontFamily="Consolas" FontSize="11" Margin="20,0,0,0" TextWrapping="Wrap"/>
                                            <TextBlock Text="* Shares must be manually recreated on the target server (Export/Import via Registry or PowerShell possible)." Margin="10,2,0,2" TextWrapping="Wrap"/>
                                            <Button x:Name="btnRobocopyHelper" Content="Robocopy Command Helper" Style="{StaticResource StandardButton}" Margin="0,10,0,0"/>
                                        </StackPanel>
                                    </Border>
                                </StackPanel>
                            </ScrollViewer>
                        </TabItem>
                        <TabItem Header="Web Server (IIS)">
                             <ScrollViewer VerticalScrollBarVisibility="Auto" MaxHeight="550">
                                <StackPanel Margin="10">
                                    <TextBlock Text="Web Server (IIS) Migration" FontSize="18" FontWeight="SemiBold" Margin="0,0,0,15"/>
                                    <Border Style="{StaticResource CardStyle}">
                                        <StackPanel>
                                            <TextBlock Text="IIS Migration Overview" FontWeight="SemiBold" Margin="0,0,0,10"/>
                                            <TextBlock TextWrapping="Wrap" Margin="0,0,0,10">
                                                Migrating IIS involves websites, application pools, configurations, certificates, and content. Web Deploy (MSDeploy.exe) is the primary tool.
                                            </TextBlock>
                                            <TextBlock Text="Web Deploy (MSDeploy):" FontWeight="Normal" Margin="0,5,0,5"/>
                                            <TextBlock Text="* Synchronizes or migrates IIS configurations and content between servers." Margin="10,2,0,2" TextWrapping="Wrap"/>
                                            <TextBlock Text="* Can be used via command line or IIS Manager UI (with Web Deploy installed)." Margin="10,2,0,2" TextWrapping="Wrap"/>
                                            <TextBlock Text="* Example (create package): `msdeploy -verb:sync -source:appHostConfig=&quot;Default Web Site&quot; -dest:package=c:\SitePackage.zip`" FontFamily="Consolas" FontSize="11" Margin="20,0,0,0" TextWrapping="Wrap"/>
                                            <TextBlock Text="* Example (deploy package): `msdeploy -verb:sync -source:package=c:\SitePackage.zip -dest:appHostConfig=&quot;Default Web Site&quot;`" FontFamily="Consolas" FontSize="11" Margin="20,0,0,0" TextWrapping="Wrap"/>
                                            <TextBlock Text="Shared Configuration:" FontWeight="Normal" Margin="0,10,0,5"/>
                                            <TextBlock Text="* IIS configuration files (applicationHost.config, administration.config) are stored on a file share and used by multiple IIS servers. Simplifies migration if already in use." Margin="10,2,0,2" TextWrapping="Wrap"/>
                                            <TextBlock Text="Manual Steps:" FontWeight="Normal" Margin="0,10,0,5"/>
                                            <TextBlock Text="* Copy website content." Margin="10,2,0,2"/>
                                            <TextBlock Text="* Export/Import SSL certificates." Margin="10,2,0,2"/>
                                            <TextBlock Text="* Manually recreate sites and AppPools (error-prone)." Margin="10,2,0,2"/>
                                            <Button x:Name="btnWebDeployHelper" Content="Web Deploy Command Helper" Style="{StaticResource StandardButton}" Margin="0,10,0,0"/>
                                        </StackPanel>
                                    </Border>
                                </StackPanel>
                            </ScrollViewer>
                        </TabItem>
                        <TabItem Header="Print Services">
                             <ScrollViewer VerticalScrollBarVisibility="Auto" MaxHeight="550">
                                <StackPanel Margin="10">
                                    <TextBlock Text="Print and Document Services Migration" FontSize="18" FontWeight="SemiBold" Margin="0,0,0,15"/>
                                    <Border Style="{StaticResource CardStyle}">
                                        <StackPanel>
                                            <TextBlock Text="Overview" FontWeight="SemiBold" Margin="0,0,0,10"/>
                                            <TextBlock Text="Use Print Management console (printbrm.exe) to export printers, drivers, and configurations from the source server and import them to the target server." TextWrapping="Wrap" Margin="0,0,0,10"/>
                                            <TextBlock Text="Key Steps:" Margin="0,5,0,5"/>
                                            <TextBlock Text="1. Source Server: Open Print Management, right-click server, select 'Export printers to a file'." Margin="10,2,0,2" TextWrapping="Wrap"/>
                                            <TextBlock Text="2. Target Server: Install Print and Document Services role." Margin="10,2,0,2" TextWrapping="Wrap"/>
                                            <TextBlock Text="3. Target Server: Open Print Management, right-click server, select 'Import printers from a file'." Margin="10,2,0,2" TextWrapping="Wrap"/>
                                            <TextBlock Text="4. Update clients (GPO, scripts) to use the new print server." Margin="10,2,0,2" TextWrapping="Wrap"/>
                                            <Button x:Name="btnOpenPrintManagement" Content="Open Print Management (Local)" Style="{StaticResource StandardButton}" Margin="0,10,0,0"/>
                                        </StackPanel>
                                    </Border>
                                </StackPanel>
                            </ScrollViewer>
                        </TabItem>
                        <TabItem Header="Remote Desktop Services">
                             <ScrollViewer VerticalScrollBarVisibility="Auto" MaxHeight="550">
                                <StackPanel Margin="10">
                                    <TextBlock Text="Remote Desktop Services (RDS) Migration" FontSize="18" FontWeight="SemiBold" Margin="0,0,0,15"/>
                                    <Border Style="{StaticResource CardStyle}">
                                        <StackPanel>
                                            <TextBlock Text="Overview" FontWeight="SemiBold" Margin="0,0,0,10"/>
                                            <TextBlock Text="RDS migration is complex and depends on the deployed roles (Connection Broker, Web Access, Gateway, Session Host, Licensing). Generally involves a side-by-side migration, adding new servers to the deployment, migrating roles, and then removing old servers." TextWrapping="Wrap" Margin="0,0,0,10"/>
                                            <TextBlock Text="Refer to official Microsoft documentation for specific RDS role migration." Margin="0,5,0,5"/>
                                            <TextBlock Text="Licensing: Migrate RDS CALs via the RD Licensing Manager." Margin="10,2,0,2" TextWrapping="Wrap"/>
                                        </StackPanel>
                                    </Border>
                                </StackPanel>
                            </ScrollViewer>
                        </TabItem>
                        <TabItem Header="WSUS">
                             <ScrollViewer VerticalScrollBarVisibility="Auto" MaxHeight="550">
                                <StackPanel Margin="10">
                                    <TextBlock Text="Windows Server Update Services (WSUS) Migration" FontSize="18" FontWeight="SemiBold" Margin="0,0,0,15"/>
                                    <Border Style="{StaticResource CardStyle}">
                                        <StackPanel>
                                            <TextBlock Text="Overview" FontWeight="SemiBold" Margin="0,0,0,10"/>
                                            <TextBlock Text="Migrate WSUS by setting up a new WSUS server and reconfiguring clients. Content can be redownloaded or copied. Database can be migrated if using SQL Server." TextWrapping="Wrap" Margin="0,0,0,10"/>
                                            <TextBlock Text="Key Steps:" Margin="0,5,0,5"/>
                                            <TextBlock Text="1. New Server: Install WSUS role." Margin="10,2,0,2" TextWrapping="Wrap"/>
                                            <TextBlock Text="2. Configure new WSUS (upstream, products, classifications)." Margin="10,2,0,2" TextWrapping="Wrap"/>
                                            <TextBlock Text="3. Content: Optionally copy WSUSContent folder or allow redownload." Margin="10,2,0,2" TextWrapping="Wrap"/>
                                            <TextBlock Text="4. Database: If using SQL, backup old DB and restore to new. If WID, consider starting fresh or more complex WID migration." Margin="10,2,0,2" TextWrapping="Wrap"/>
                                            <TextBlock Text="5. Update GPO to point clients to the new WSUS server." Margin="10,2,0,2" TextWrapping="Wrap"/>
                                            <TextBlock Text="6. Decommission old WSUS server after clients report to new one." Margin="10,2,0,2" TextWrapping="Wrap"/>
                                        </StackPanel>
                                    </Border>
                                </StackPanel>
                            </ScrollViewer>
                        </TabItem>
                        <TabItem Header="Hyper V">
                             <ScrollViewer VerticalScrollBarVisibility="Auto" MaxHeight="550">
                                <StackPanel Margin="10">
                                    <TextBlock Text="Hyper V Migration" FontSize="18" FontWeight="SemiBold" Margin="0,0,0,15"/>
                                    <Border Style="{StaticResource CardStyle}">
                                        <StackPanel>
                                            <TextBlock Text="Overview" FontWeight="SemiBold" Margin="0,0,0,10"/>
                                            <TextBlock Text="Migrating Hyper-V virtual machines can be done via Export/Import, Live Migration (if clustered or shared storage), or Storage Migration." TextWrapping="Wrap" Margin="0,0,0,10"/>
                                            <TextBlock Text="Methods:" Margin="0,5,0,5"/>
                                            <TextBlock Text="* Export/Import: Suitable for standalone hosts, involves downtime." Margin="10,2,0,2" TextWrapping="Wrap"/>
                                            <TextBlock Text="* Live Migration: Requires clustering or SMB 3.0 shared storage, minimal downtime." Margin="10,2,0,2" TextWrapping="Wrap"/>
                                            <TextBlock Text="* Shared Nothing Live Migration (Server 2012+): Migrate VMs between non-clustered hosts with only a network connection." Margin="10,2,0,2" TextWrapping="Wrap"/>
                                            <TextBlock Text="* Storage Live Migration: Move VM storage with no downtime." Margin="10,2,0,2" TextWrapping="Wrap"/>
                                            <TextBlock Text="Consider network configuration (virtual switches) and storage paths." Margin="0,5,0,5"/>
                                        </StackPanel>
                                    </Border>
                                </StackPanel>
                            </ScrollViewer>
                        </TabItem>
                        <TabItem Header="AD CS">
                             <ScrollViewer VerticalScrollBarVisibility="Auto" MaxHeight="550">
                                <StackPanel Margin="10">
                                    <TextBlock Text="Active Directory Certificate Services (AD CS) Migration" FontSize="18" FontWeight="SemiBold" Margin="0,0,0,15"/>
                                    <Border Style="{StaticResource CardStyle}">
                                        <StackPanel>
                                            <TextBlock Text="Overview" FontWeight="SemiBold" Margin="0,0,0,10"/>
                                            <TextBlock Text="AD CS migration is critical and must be planned carefully. It involves backing up the CA database and private key, restoring to a new server with the same name (or reconfiguring clients), and reissuing templates if needed." TextWrapping="Wrap" Margin="0,0,0,10"/>
                                            <TextBlock Text="Always consult official Microsoft documentation for AD CS migration. Errors can have significant impact." Margin="0,5,0,5" Foreground="DarkRed"/>
                                        </StackPanel>
                                    </Border>
                                </StackPanel>
                            </ScrollViewer>
                        </TabItem>
                        <!-- Add more role tabs here -->
                    </TabControl>
                </StackPanel>

                <!-- Cluster Upgrade Panel -->
                <StackPanel x:Name="panelClusterUpgrade" Visibility="Collapsed">
                    <TextBlock Text="Cluster OS Rolling Upgrade" FontSize="24" FontWeight="Bold" Margin="0,0,0,20"/>
                    <Border Style="{StaticResource CardStyle}">
                        <StackPanel>
                            <TextBlock Text="Failover Cluster Operating System Upgrade" FontWeight="SemiBold" Margin="0,0,0,10"/>
                            <TextBlock Text="Perform a rolling upgrade for your Windows Server Failover Cluster to update the OS of cluster nodes without downtime for cluster roles (e.g., Hyper-V or File Server)." TextWrapping="Wrap" Margin="0,0,0,15"/>
                            <TextBlock Text="Prerequisites:" Margin="0,0,0,5"/>
                            <TextBlock Text="* Cluster must be running Windows Server 2012 R2 or later." Margin="15,2,0,2"/>
                            <TextBlock Text="* Target OS must be Windows Server 2016 or later." Margin="15,2,0,2"/>
                            <TextBlock Text="* All nodes must have the same or compatible hardware." Margin="15,2,0,2"/>
                            <TextBlock Text="* Sufficient capacity in the cluster to take one node offline." Margin="15,2,0,2"/>
                        </StackPanel>
                    </Border>
                     <Border Style="{StaticResource CardStyle}">
                        <StackPanel>
                            <TextBlock Text="Upgrade Process (Simplified):" FontWeight="SemiBold" Margin="0,0,0,10"/>
                            <TextBlock Text="1. Prepare cluster for mixed-mode (Update-ClusterFunctionalLevel)." Margin="10,2,0,2"/>
                            <TextBlock Text="2. Remove one node from the cluster (pause, drain roles, evict)." Margin="10,2,0,2"/>
                            <TextBlock Text="3. Reinstall/upgrade OS of the removed node." Margin="10,2,0,2"/>
                            <TextBlock Text="4. Add the upgraded node back to the cluster." Margin="10,2,0,2"/>
                            <TextBlock Text="5. Repeat steps 2-4 for all nodes." Margin="10,2,0,2"/>
                            <TextBlock Text="6. Update cluster functional level (Update-ClusterFunctionalLevel)." Margin="10,2,0,2"/>

                            <CheckBox x:Name="chkValidateCluster" Content="Perform Cluster Validation before Upgrade" IsChecked="True" Margin="0,15,0,5"/>
                            <CheckBox x:Name="chkBackupClusterConfig" Content="Backup Cluster Configuration" IsChecked="True" Margin="0,5"/>
                            <Button x:Name="btnStartClusterUpgradeWizard" Content="Start Cluster Upgrade Wizard" Style="{StaticResource StandardButton}" Margin="0,15,0,0"/>
                        </StackPanel>
                    </Border>
                </StackPanel>

                <!-- Storage Migration Panel -->
                <StackPanel x:Name="panelStorageMigration" Visibility="Collapsed">
                    <TextBlock Text="Storage Migration Service" FontSize="24" FontWeight="Bold" Margin="0,0,0,20"/>
                     <Border Style="{StaticResource CardStyle}">
                        <StackPanel>
                            <TextBlock Text="File Server and Storage Migration" FontWeight="SemiBold" Margin="0,0,0,10"/>
                            <TextBlock Text="Use the Storage Migration Service (SMS) to migrate file servers and their data to newer Windows Servers or Azure VMs. SMS handles inventory, transfer, and optionally, cutover (identity switch)." TextWrapping="Wrap" Margin="0,0,0,15"/>
                            <TextBlock Text="SMS requires an orchestration server (Windows Server 2019 or later) and can be managed via Windows Admin Center or PowerShell." TextWrapping="Wrap" Margin="0,0,0,10"/>
                        </StackPanel>
                    </Border>
                    <Border Style="{StaticResource CardStyle}">
                        <StackPanel>
                            <TextBlock Text="Migration Parameters" FontWeight="SemiBold" Margin="0,0,0,10"/>
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
                                
                                <TextBlock Grid.Row="0" Grid.Column="0" Text="Orchestration Server:" Margin="0,5,10,5" VerticalAlignment="Center"/>
                                <TextBox Grid.Row="0" Grid.Column="1" x:Name="txtSmsOrchestrator" Text="localhost" Margin="0,5" ToolTip="The server running the Storage Migration Service Proxy."/>
                                
                                <TextBlock Grid.Row="1" Grid.Column="0" Text="Source Server(s):" Margin="0,5,10,5" VerticalAlignment="Center"/>
                                <TextBox Grid.Row="1" Grid.Column="1" x:Name="txtSmsSourceServers" Margin="0,5" ToolTip="Comma-separated list of source servers."/>
                                
                                <TextBlock Grid.Row="2" Grid.Column="0" Text="Target Server:" Margin="0,5,10,5" VerticalAlignment="Center"/>
                                <TextBox Grid.Row="2" Grid.Column="1" x:Name="txtSmsTargetServer" Margin="0,5"/>
                                
                                <Button Grid.Row="3" Grid.Column="1" x:Name="btnStartStorageMigrationService" Content="Start Storage Migration" Style="{StaticResource StandardButton}" Margin="0,15,0,0"/>
                            </Grid>
                             <TextBlock Text="Note: This tool can generate basic PowerShell commands for SMS. For complex scenarios, Windows Admin Center is recommended." Margin="0,10,0,0" FontStyle="Italic" FontSize="12" TextWrapping="Wrap"/>
                        </StackPanel>
                    </Border>
                </StackPanel>
                
                <!-- Decommission Panel -->
                <StackPanel x:Name="panelDecommission" Visibility="Collapsed">
                    <TextBlock Text="Server Decommissioning" FontSize="24" FontWeight="Bold" Margin="0,0,0,20"/>
                    <Border Style="{StaticResource CardStyle}">
                        <StackPanel>
                            <TextBlock Text="Secure Server Decommissioning" FontWeight="SemiBold" Margin="0,0,0,10"/>
                            <TextBlock Text="Checklist for proper decommissioning of old servers after successful migration of all services and data. Ensure all steps are carefully reviewed before a server is permanently taken offline." TextWrapping="Wrap" Margin="0,0,0,15"/>
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                <StackPanel Grid.Column="0">
                                    <CheckBox x:Name="chkBackupComplete" Content="Full final backup created" Margin="0,5"/>
                                    <CheckBox x:Name="chkRolesMigrated" Content="All roles successfully migrated/replaced" Margin="0,5"/>
                                    <CheckBox x:Name="chkDataMigrated" Content="All data successfully migrated/archived" Margin="0,5"/>
                                    <CheckBox x:Name="chkApplicationsMoved" Content="All applications migrated/replaced" Margin="0,5"/>
                                </StackPanel>
                                <StackPanel Grid.Column="1">
                                    <CheckBox x:Name="chkDNSUpdated" Content="DNS records updated/removed" Margin="0,5"/>
                                    <CheckBox x:Name="chkCertificatesMigrated" Content="Certificates migrated/revoked" Margin="0,5"/>
                                    <CheckBox x:Name="chkUsersNotified" Content="Users/Stakeholders informed" Margin="0,5"/>
                                    <CheckBox x:Name="chkMonitoringDisabled" Content="Monitoring/Alerting disabled" Margin="0,5"/>
                                </StackPanel>
                            </Grid>
                             <TextBlock Text="Server Name for Decommissioning:" Margin="0,15,0,2"/>
                             <TextBox x:Name="txtDecommissionServerName" Margin="0,0,0,10"/>
                            <Button x:Name="btnDecommissionServer" Content="Decommission Server (Simulated)" Style="{StaticResource StandardButton}" Margin="0,10,0,0" IsEnabled="False" ToolTip="Enabled when all checks are positive. Does not perform actual decommissioning actions but documents the process."/>
                            <TextBlock Text="Important: Actual decommissioning (shutdown, removal from AD/VMware, etc.) must be done manually and with caution." Margin="0,10,0,0" FontStyle="Italic" FontSize="12" TextWrapping="Wrap"/>
                        </StackPanel>
                    </Border>
                </StackPanel>

                <!-- Settings Panel -->
                <StackPanel x:Name="panelSettings" Visibility="Collapsed">
                    <TextBlock Text="Settings" FontSize="24" FontWeight="Bold" Margin="0,0,0,20"/>
                    <Border Style="{StaticResource CardStyle}">
                        <StackPanel>
                            <TextBlock Text="General Settings" FontWeight="SemiBold" Margin="0,0,0,10"/>
                            <CheckBox x:Name="chkEnableLogging" Content="Enable Logging" IsChecked="True" Margin="0,5"/>
                            <TextBlock Text="Log File Path:" Margin="0,5,0,2"/>
                            <TextBox x:Name="txtLogFilePath" Text="{Binding Path=LogFile, Mode=TwoWay}" Margin="0,0,0,10"/>
                            <Button x:Name="btnBrowseLogPath" Content="Browse..." Style="{StaticResource SecondaryButton}" HorizontalAlignment="Left"/>
                        </StackPanel>
                    </Border>
                    <Border Style="{StaticResource CardStyle}">
                        <StackPanel>
                            <TextBlock Text="Theme Settings" FontWeight="SemiBold" Margin="0,0,0,10"/>
                             <TextBlock Text="Accent Color (Hex):" Margin="0,5,0,2"/>
                             <TextBox x:Name="txtThemeColor" Text="{Binding Path=ThemeColor, Mode=TwoWay}" Margin="0,0,0,10"/>
                             <Button x:Name="btnApplyTheme" Content="Apply Theme" Style="{StaticResource StandardButton}" HorizontalAlignment="Left"/>
                        </StackPanel>
                    </Border>
                     <Border Style="{StaticResource CardStyle}">
                        <StackPanel>
                            <TextBlock Text="About This Tool" FontWeight="SemiBold" Margin="0,0,0,10"/>
                            <TextBlock x:Name="txtAppVersionInfo" Text="Version: 0.0.1 (Example)" Margin="0,5"/>
                            <TextBlock Text="Developed by: EasyIT" Margin="0,5"/>
                             <TextBlock Margin="0,15,0,0">
                                <Hyperlink NavigateUri="https://github.com/easyIT-Gruppe/easyWSMigrate" x:Name="linkGitHubRepo">
                                    GitHub Repository
                                </Hyperlink>
                            </TextBlock>
                        </StackPanel>
                    </Border>
                </StackPanel>


                 <!-- Reports Panel Placeholder -->
                <StackPanel x:Name="panelReports" Visibility="Collapsed">
                    <TextBlock Text="Reports" FontSize="24" FontWeight="Bold" Margin="0,0,0,20"/>
                    <Border Style="{StaticResource CardStyle}">
                        <StackPanel>
                            <TextBlock Text="Available Reports" FontWeight="SemiBold" Margin="0,0,0,10"/>
                            <TextBlock Text="This section will display generated reports from scans, audits, and migrations, and offer export options." TextWrapping="Wrap" Margin="0,0,0,15"/>
                            <TextBlock Text="Examples of Reports:" Margin="0,0,0,5"/>
                            <TextBlock Text="* System Analysis Report" Margin="15,2,0,2"/>
                            <TextBlock Text="* Server Audit Summary" Margin="15,2,0,2"/>
                            <TextBlock Text="* Migration Status Report" Margin="15,2,0,2"/>
                            <TextBlock Text="* Validation Report" Margin="15,2,0,2"/>
                            <TextBlock Text="Functionality will be implemented in a later version." FontStyle="Italic" Margin="0,10,0,0"/>
                        </StackPanel>
                    </Border>
                </StackPanel>

                <!-- Network Audit Panel Placeholder -->
                <StackPanel x:Name="panelAuditNetwork" Visibility="Collapsed">
                    <TextBlock Text="Network Audit" FontSize="24" FontWeight="Bold" Margin="0,0,0,20"/>
                    <Border Style="{StaticResource CardStyle}">
                        <StackPanel>
                            <TextBlock Text="Detailed Network Audit" FontWeight="SemiBold" Margin="0,0,0,10"/>
                            <TextBlock Text="This section will provide tools for a detailed network audit, including port scans, configuration analysis of firewalls and switches (if APIs are available), and bandwidth tests." TextWrapping="Wrap" Margin="0,0,0,15"/>
                            <TextBlock Text="Functionality will be implemented in a later version." FontStyle="Italic" Margin="0,10,0,0"/>
                        </StackPanel>
                    </Border>
                </StackPanel>

                <!-- Security Audit Panel Placeholder -->
                <StackPanel x:Name="panelAuditSecurity" Visibility="Collapsed">
                    <TextBlock Text="Security Audit" FontSize="24" FontWeight="Bold" Margin="0,0,0,20"/>
                    <Border Style="{StaticResource CardStyle}">
                        <StackPanel>
                            <TextBlock Text="Comprehensive Security Audit" FontWeight="SemiBold" Margin="0,0,0,10"/>
                            <TextBlock Text="This section will provide tools for a security audit, e.g., checking security settings, patch levels, user account policies, and potentially integrating with security baselines." TextWrapping="Wrap" Margin="0,0,0,15"/>
                            <TextBlock Text="Functionality will be implemented in a later version." FontStyle="Italic" Margin="0,10,0,0"/>
                        </StackPanel>
                    </Border>
                </StackPanel>

                <!-- Performance Audit Panel Placeholder -->
                <StackPanel x:Name="panelAuditPerformance" Visibility="Collapsed">
                    <TextBlock Text="Performance Audit" FontSize="24" FontWeight="Bold" Margin="0,0,0,20"/>
                    <Border Style="{StaticResource CardStyle}">
                        <StackPanel>
                            <TextBlock Text="Server Performance Analysis" FontWeight="SemiBold" Margin="0,0,0,10"/>
                            <TextBlock Text="This section will provide tools to analyze server performance, including monitoring CPU, RAM, disk I/O, network, and key performance indicators." TextWrapping="Wrap" Margin="0,0,0,15"/>
                            <TextBlock Text="Functionality will be implemented in a later version." FontStyle="Italic" Margin="0,10,0,0"/>
                        </StackPanel>
                    </Border>
                </StackPanel>

                <!-- Compliance Check Panel Placeholder -->
                <StackPanel x:Name="panelAuditCompliance" Visibility="Collapsed">
                    <TextBlock Text="Compliance Check" FontSize="24" FontWeight="Bold" Margin="0,0,0,20"/>
                    <Border Style="{StaticResource CardStyle}">
                        <StackPanel>
                            <TextBlock Text="Compliance Verification" FontWeight="SemiBold" Margin="0,0,0,10"/>
                            <TextBlock Text="This section will provide tools to check server configuration against defined compliance policies (e.g., CIS Benchmarks, internal standards)." TextWrapping="Wrap" Margin="0,0,0,15"/>
                            <TextBlock Text="Functionality will be implemented in a later version." FontStyle="Italic" Margin="0,10,0,0"/>
                        </StackPanel>
                    </Border>
                </StackPanel>


            </Grid>
            </ScrollViewer>
        </Grid>
        
        <!-- Footer -->
        <Border Grid.Row="2" Background="White" BorderBrush="#E0E0E0" BorderThickness="0,1,0,0">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                
                <TextBlock Grid.Column="0" x:Name="txtFooter" Text="{FooterText}" 
                           Foreground="#666666" VerticalAlignment="Center" Margin="24,0,0,0"/>
                
                <TextBlock Grid.Column="1" Margin="0,0,24,0" VerticalAlignment="Center">
                    <Hyperlink NavigateUri="{Binding Path=FooterWebsite, Mode=OneWay}" x:Name="linkFooterWebsite">
                         <TextBlock Text="{FooterWebsite}" Foreground="{StaticResource ThemeBrush}"/>
                    </Hyperlink>
                </TextBlock>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

# Platzhalter im XAML ersetzen
$XAML = $XAML -replace "{AppName}", $script:appName
$XAML = $XAML -replace "{ThemeColor}", $script:themeColor
$XAML = $XAML -replace "{FooterText}", $script:footerText
$XAML = $XAML -replace "{FooterWebsite}", $script:footerWebsite

# GUI in PowerShell laden
try {
    [xml]$xaml = $XAML
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
$listNavigation = $window.FindName("listNavigation")
$navDashboard = $window.FindName("navDashboard")
$navScan = $window.FindName("navScan")
$navAudit = $window.FindName("navAudit")
$navMigration = $window.FindName("navMigration")
$navReports = $window.FindName("navReports")
$navSettings = $window.FindName("navSettings")

# Migration Tools
$listMigrationTools = $window.FindName("listMigrationTools")
$toolRoleInstall = $window.FindName("toolRoleInstall")
$toolDataMigration = $window.FindName("toolDataMigration")
$toolClusterUpgrade = $window.FindName("toolClusterUpgrade")
$toolStorageMigration = $window.FindName("toolStorageMigration")

# Audit Tools
$listAuditTools = $window.FindName("listAuditTools")
$auditNetwork = $window.FindName("auditNetwork")
$auditSecurity = $window.FindName("auditSecurity")
$auditPerformance = $window.FindName("auditPerformance")
$auditCompliance = $window.FindName("auditCompliance")

# Header-Elemente
$txtCurrentPageTitle = $window.FindName("txtCurrentPageTitle")
$btnHelp = $window.FindName("btnHelp")
$btnSettings = $window.FindName("btnSettings")

# Inhaltsbereich
$mainContent = $window.FindName("mainContent")
$panelHome = $window.FindName("panelHome")
$panelDiscovery = $window.FindName("panelDiscovery")
$panelInstallation = $window.FindName("panelInstallation")
$panelMigration = $window.FindName("panelMigration")
$panelValidation = $window.FindName("panelValidation")
$panelDecommission = $window.FindName("panelDecommission")
$panelAudit = $window.FindName("panelAudit")
$panelMigrationTools = $window.FindName("panelMigrationTools")

# Migration-Elemente
$chkMigrateUserData = $window.FindName("chkMigrateUserData")
$chkMigrateShares = $window.FindName("chkMigrateShares")
$chkMigratePrinters = $window.FindName("chkMigratePrinters")
$chkMigrateRegistry = $window.FindName("chkMigrateRegistry")
$chkMigrateServices = $window.FindName("chkMigrateServices")
$txtMigrationSource = $window.FindName("txtMigrationSource")
$txtMigrationTarget = $window.FindName("txtMigrationTarget")
$btnStartMigration = $window.FindName("btnStartMigration")
$txtMigrationStatus = $window.FindName("txtMigrationStatus")
$progressMigration = $window.FindName("progressMigration")
$migrationResults = $window.FindName("migrationResults")
$dgMigrationResults = $window.FindName("dgMigrationResults")

# Discovery-Bereich
$btnStartDiscovery = $window.FindName("btnStartDiscovery")
$tabDiscovery = $window.FindName("tabDiscovery")
$dgDomainControllers = $window.FindName("dgDomainControllers")
$lvFSMORoles = $window.FindName("lvFSMORoles")
$dgServices = $window.FindName("dgServices")
$txtSummary = $window.FindName("txtSummary")

# Export-Buttons
$btnExportHTML = $window.FindName("btnExportHTML")
$btnPrintReport = $window.FindName("btnPrintReport")

# Installation
$txtInstallationStatus = $window.FindName("txtInstallationStatus")
$txtInstallationServer = $window.FindName("txtInstallationServer")
$listAvailableRoles = $window.FindName("listAvailableRoles")
$btnStartInstallation = $window.FindName("btnStartInstallation")
$progressInstallation = $window.FindName("progressInstallation")
$installationResults = $window.FindName("installationResults")
$dgInstallationResults = $window.FindName("dgInstallationResults")

# Validierungs-Elemente
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

# Audit-Elemente
$btnStartAudit = $window.FindName("btnStartAudit")
$txtAuditStatus = $window.FindName("txtAuditStatus")
$progressAudit = $window.FindName("progressAudit")
$tabAudit = $window.FindName("tabAudit")
$txtAuditSystem = $window.FindName("txtAuditSystem")
$txtAuditNetwork = $window.FindName("txtAuditNetwork")
$txtAuditSecurity = $window.FindName("txtAuditSecurity")
$txtAuditConnections = $window.FindName("txtAuditConnections")
$txtAuditSummary = $window.FindName("txtAuditSummary")

# Audit-Checkboxen
$chkAuditSystem = $window.FindName("chkAuditSystem")
$chkAuditNetwork = $window.FindName("chkAuditNetwork")
$chkAuditSecurity = $window.FindName("chkAuditSecurity")
$chkAuditServices = $window.FindName("chkAuditServices")
$chkAuditActiveDirectory = $window.FindName("chkAuditActiveDirectory")
$chkAuditDNS = $window.FindName("chkAuditDNS")
$chkAuditDHCP = $window.FindName("chkAuditDHCP")
$chkAuditIIS = $window.FindName("chkAuditIIS")
$chkAuditHyperV = $window.FindName("chkAuditHyperV")
$chkAuditConnections = $window.FindName("chkAuditConnections")

# Audit-Export-Buttons
$btnExportAuditHTML = $window.FindName("btnExportAuditHTML")
$btnExportAuditDrawIO = $window.FindName("btnExportAuditDrawIO")
$btnPrintAudit = $window.FindName("btnPrintAudit")

# Migration-Tool-Elemente
$btnInPlaceUpgrade = $window.FindName("btnInPlaceUpgrade")
$btnSideBySideMigration = $window.FindName("btnSideBySideMigration")
$btnClusterUpgrade = $window.FindName("btnClusterUpgrade")
$btnStorageMigration = $window.FindName("btnStorageMigration")
$txtSourceServer = $window.FindName("txtSourceServer")
$txtTargetServer = $window.FindName("txtTargetServer")

# Decommission-Elemente
$chkBackupComplete = $window.FindName("chkBackupComplete")
$chkRolesMigrated = $window.FindName("chkRolesMigrated")
$chkDataMigrated = $window.FindName("chkDataMigrated")
$chkDNSUpdated = $window.FindName("chkDNSUpdated")
$chkCertificatesMigrated = $window.FindName("chkCertificatesMigrated")
$chkUsersNotified = $window.FindName("chkUsersNotified")
$chkMonitoringDisabled = $window.FindName("chkMonitoringDisabled")
$btnDecommissionServer = $window.FindName("btnDecommissionServer")

# Footer
$txtFooter = $window.FindName("txtFooter")
$txtWebsite = $window.FindName("txtWebsite")
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
        $panelAudit.Visibility = "Collapsed"
        $panelMigrationTools.Visibility = "Collapsed"
        
        # Ausgewähltes Panel anzeigen
        switch ($ViewName) {
            "Home" { $panelHome.Visibility = "Visible" }
            "Discovery" { $panelDiscovery.Visibility = "Visible" }
            "Installation" { $panelInstallation.Visibility = "Visible" }
            "Migration" { $panelMigration.Visibility = "Visible" }
            "Validation" { $panelValidation.Visibility = "Visible" }
            "Decommission" { $panelDecommission.Visibility = "Visible" }
            "Audit" { $panelAudit.Visibility = "Visible" }
            "MigrationTools" { $panelMigrationTools.Visibility = "Visible" }
        }
        
        # Tab-Titel aktualisieren
        $txtCurrentPageTitle.Text = $TabTitle
        
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
$navDashboard.Add_Selected({ Switch-View -ViewName "Home" -TabTitle "Übersicht" })
$navScan.Add_Selected({ Switch-View -ViewName "Discovery" -TabTitle "Systemanalyse" })
$navAudit.Add_Selected({ Switch-View -ViewName "Audit" -TabTitle "Server Audit"; Initialize-AuditUI })
$navMigration.Add_Selected({ Switch-View -ViewName "MigrationTools" -TabTitle "Migration Tools" })
$navReports.Add_Selected({ Switch-View -ViewName "Validation" -TabTitle "Validierung"; Initialize-ValidationUI })
$navSettings.Add_Selected({ Switch-View -ViewName "Installation" -TabTitle "Installation" })

# Tools Event-Handler
$toolRoleInstall.Add_Selected({ Switch-View -ViewName "Installation" -TabTitle "Rolleninstallation" })
$toolDataMigration.Add_Selected({ Switch-View -ViewName "MigrationTools" -TabTitle "Datenmigration" })
$toolClusterUpgrade.Add_Selected({ Switch-View -ViewName "MigrationTools" -TabTitle "Cluster Upgrade" })
$toolStorageMigration.Add_Selected({ Switch-View -ViewName "MigrationTools" -TabTitle "Storage Migration" })

# Audit Tools Event-Handler
$auditNetwork.Add_Selected({ Switch-View -ViewName "Audit" -TabTitle "Netzwerk-Audit"; Initialize-AuditUI })
$auditSecurity.Add_Selected({ Switch-View -ViewName "Audit" -TabTitle "Sicherheits-Audit"; Initialize-AuditUI })
$auditPerformance.Add_Selected({ Switch-View -ViewName "Audit" -TabTitle "Performance-Audit"; Initialize-AuditUI })
$auditCompliance.Add_Selected({ Switch-View -ViewName "Audit" -TabTitle "Compliance-Check"; Initialize-AuditUI })

# Installation Event-Handler
$btnStartInstallation.Add_Click({
    try {
        $selectedRoles = @()
        foreach ($item in $listAvailableRoles.SelectedItems) {
            $selectedRoles += $item.Tag
        }
        
        if ($selectedRoles.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Bitte wählen Sie mindestens eine Rolle zur Installation aus.", "Keine Rolle ausgewählt", "OK", "Warning")
            return
        }
        
        $targetServer = $txtInstallationServer.Text.Trim()
        if ([string]::IsNullOrEmpty($targetServer) -or $targetServer -eq "localhost") {
            $targetServer = $env:COMPUTERNAME
        }
        
        [System.Windows.Forms.MessageBox]::Show("Rolleninstallation wird in einer zukünftigen Version implementiert.`n`nAusgewählte Rollen: $($selectedRoles -join ', ')`nZielserver: $targetServer", "Installation", "OK", "Information")
    }
    catch {
        Write-Log "Fehler bei der Rolleninstallation: $($_.Exception.Message)" -Level "ERROR"
        [System.Windows.Forms.MessageBox]::Show("Fehler bei der Rolleninstallation: $($_.Exception.Message)", "Installation Fehler", "OK", "Error")
    }
})

# Migration Event-Handler
$btnStartMigration.Add_Click({
    try {
        $sourceServer = $txtMigrationSource.Text.Trim()
        $targetServer = $txtMigrationTarget.Text.Trim()
        
        if ([string]::IsNullOrEmpty($sourceServer) -or [string]::IsNullOrEmpty($targetServer)) {
            [System.Windows.Forms.MessageBox]::Show("Bitte geben Sie sowohl Quell- als auch Ziel-Server an.", "Eingabe erforderlich", "OK", "Warning")
            return
        }
        
        $selectedMigrationTypes = @()
        if ($chkMigrateUserData.IsChecked) { $selectedMigrationTypes += "Benutzerdaten" }
        if ($chkMigrateShares.IsChecked) { $selectedMigrationTypes += "Dateifreigaben" }
        if ($chkMigratePrinters.IsChecked) { $selectedMigrationTypes += "Drucker" }
        if ($chkMigrateRegistry.IsChecked) { $selectedMigrationTypes += "Registry-Einstellungen" }
        if ($chkMigrateServices.IsChecked) { $selectedMigrationTypes += "Dienst-Konfigurationen" }
        
        if ($selectedMigrationTypes.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Bitte wählen Sie mindestens einen Migrationstyp aus.", "Keine Auswahl", "OK", "Warning")
            return
        }
        
        [System.Windows.Forms.MessageBox]::Show("Datenmigration wird in einer zukünftigen Version implementiert.`n`nQuell-Server: $sourceServer`nZiel-Server: $targetServer`nMigrationstypen: $($selectedMigrationTypes -join ', ')", "Migration", "OK", "Information")
    }
    catch {
        Write-Log "Fehler bei der Datenmigration: $($_.Exception.Message)" -Level "ERROR"
        [System.Windows.Forms.MessageBox]::Show("Fehler bei der Datenmigration: $($_.Exception.Message)", "Migration Fehler", "OK", "Error")
    }
})

# Header Event-Handler
$btnHelp.Add_Click({
    [System.Windows.Forms.MessageBox]::Show("Windows Server 2012 Migration Tool`nVersion 0.0.1`n`nDieses Tool unterstützt die Migration von Windows Server 2012 Domänencontrollern, DNS- und DHCP-Diensten auf neue Windows Server.", "Information", "OK", "Information")
})

$btnSettings.Add_Click({
    [System.Windows.Forms.MessageBox]::Show("Einstellungen sind in dieser Version noch nicht verfügbar.", "Einstellungen", "OK", "Information")
})

# Footer Event-Handler
$txtWebsite.Add_MouseLeftButtonUp({
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
            $validationExportButtons = $window.FindName("validationExportButtons")
            if ($validationExportButtons) {
                $validationExportButtons.Visibility = "Visible"
            }
            
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

# Audit Event-Handler
$btnStartAudit.Add_Click({
    try {
        Start-ServerAudit
    }
    catch {
        Write-Log "Fehler beim Starten des Audits: $($_.Exception.Message)" -Level "ERROR"
        [System.Windows.Forms.MessageBox]::Show("Fehler beim Starten des Audits: $($_.Exception.Message)", "Audit Fehler", "OK", "Error")
    }
})

$btnExportAuditHTML.Add_Click({
    try {
        $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveDialog.Filter = "HTML-Dateien (*.html)|*.html"
        $saveDialog.Title = "Audit-Ergebnisse als HTML speichern"
        $saveDialog.FileName = "ServerAudit_$env:COMPUTERNAME_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
        
        if ($saveDialog.ShowDialog() -eq "OK") {
            $success = Export-AuditToHTML -FilePath $saveDialog.FileName
            if ($success) {
                [System.Windows.Forms.MessageBox]::Show("Audit-Ergebnisse erfolgreich als HTML exportiert:`n$($saveDialog.FileName)", "Export erfolgreich", "OK", "Information")
                Start-Process $saveDialog.FileName
            } else {
                [System.Windows.Forms.MessageBox]::Show("Fehler beim HTML-Export der Audit-Ergebnisse.", "Export Fehler", "OK", "Error")
            }
        }
    }
    catch {
        Write-Log "Fehler beim HTML-Export: $($_.Exception.Message)" -Level "ERROR"
        [System.Windows.Forms.MessageBox]::Show("Fehler beim HTML-Export: $($_.Exception.Message)", "Export Fehler", "OK", "Error")
    }
})

$btnExportAuditDrawIO.Add_Click({
    [System.Windows.Forms.MessageBox]::Show("Draw.IO Netzwerk-Topologie Export wird in einer zukünftigen Version verfügbar sein.", "Feature in Entwicklung", "OK", "Information")
})

$btnPrintAudit.Add_Click({
    try {
        [System.Windows.Forms.MessageBox]::Show("Druck-Funktionalität wird in einer zukünftigen Version implementiert.", "Druck", "OK", "Information")
    }
    catch {
        Write-Log "Fehler beim Drucken: $($_.Exception.Message)" -Level "ERROR"
    }
})

# Migration Tools Event-Handler
$btnInPlaceUpgrade.Add_Click({
    [System.Windows.Forms.MessageBox]::Show("In-Place Upgrade Funktionalität wird in einer zukünftigen Version verfügbar sein.`n`nHinweis: In-Place Upgrades sollten nur nach gründlicher Planung und Backup durchgeführt werden.", "In-Place Upgrade", "OK", "Information")
})

$btnSideBySideMigration.Add_Click({
    [System.Windows.Forms.MessageBox]::Show("Side-by-Side Migration Funktionalität wird in einer zukünftigen Version verfügbar sein.`n`nDiese Methode ist die sicherste für kritische Produktionsumgebungen.", "Side-by-Side Migration", "OK", "Information")
})

$btnClusterUpgrade.Add_Click({
    [System.Windows.Forms.MessageBox]::Show("Cluster OS Rolling Upgrade Funktionalität wird in einer zukünftigen Version verfügbar sein.`n`nDiese Funktion ermöglicht Cluster-Upgrades ohne Downtime.", "Cluster Upgrade", "OK", "Information")
})

$btnStorageMigration.Add_Click({
    try {
        $sourceServer = $txtSourceServer.Text.Trim()
        $targetServer = $txtTargetServer.Text.Trim()
        
        if ([string]::IsNullOrEmpty($sourceServer) -or [string]::IsNullOrEmpty($targetServer)) {
            [System.Windows.Forms.MessageBox]::Show("Bitte geben Sie sowohl Quell- als auch Ziel-Server an.", "Eingabe erforderlich", "OK", "Warning")
            return
        }
        
        [System.Windows.Forms.MessageBox]::Show("Storage Migration Service wird gestartet...`n`nQuell-Server: $sourceServer`nZiel-Server: $targetServer`n`nDiese Funktionalität wird in einer zukünftigen Version vollständig implementiert.", "Storage Migration", "OK", "Information")
    }
    catch {
        Write-Log "Fehler bei Storage Migration: $($_.Exception.Message)" -Level "ERROR"
    }
})

# Hilfsfunktion für Decommission-Button
function Update-DecommissionButton {
    try {
        $allChecked = $chkBackupComplete.IsChecked -and $chkRolesMigrated.IsChecked -and $chkDataMigrated.IsChecked -and 
                     $chkDNSUpdated.IsChecked -and $chkCertificatesMigrated.IsChecked -and $chkUsersNotified.IsChecked -and 
                     $chkMonitoringDisabled.IsChecked
        
        $btnDecommissionServer.IsEnabled = $allChecked
    }
    catch {
        Write-Log "Fehler beim Aktualisieren des Decommission-Buttons: $($_.Exception.Message)" -Level "ERROR"
    }
}

# Decommission Event-Handler
$chkBackupComplete.Add_Checked({ Update-DecommissionButton })
$chkRolesMigrated.Add_Checked({ Update-DecommissionButton })
$chkDataMigrated.Add_Checked({ Update-DecommissionButton })
$chkDNSUpdated.Add_Checked({ Update-DecommissionButton })
$chkCertificatesMigrated.Add_Checked({ Update-DecommissionButton })
$chkUsersNotified.Add_Checked({ Update-DecommissionButton })
$chkMonitoringDisabled.Add_Checked({ Update-DecommissionButton })

$chkBackupComplete.Add_Unchecked({ Update-DecommissionButton })
$chkRolesMigrated.Add_Unchecked({ Update-DecommissionButton })
$chkDataMigrated.Add_Unchecked({ Update-DecommissionButton })
$chkDNSUpdated.Add_Unchecked({ Update-DecommissionButton })
$chkCertificatesMigrated.Add_Unchecked({ Update-DecommissionButton })
$chkUsersNotified.Add_Unchecked({ Update-DecommissionButton })
$chkMonitoringDisabled.Add_Unchecked({ Update-DecommissionButton })

$btnDecommissionServer.Add_Click({
    $result = [System.Windows.Forms.MessageBox]::Show("Sind Sie sicher, dass Sie den Server abschalten möchten?`n`nDieser Vorgang kann nicht rückgängig gemacht werden!", "Server Abschaltung bestätigen", "YesNo", "Warning")
    if ($result -eq "Yes") {
        [System.Windows.Forms.MessageBox]::Show("Server-Abschaltung wird in einer zukünftigen Version implementiert.`n`nBitte führen Sie die Abschaltung manuell durch.", "Abschaltung", "OK", "Information")
    }
})

# Zusätzliche Event-Handler können hier hinzugefügt werden
#endregion

#region Audit-Funktionen (aus easyWSAudit integriert)

# Globale Variablen für Audit-Ergebnisse
$script:auditResults = @{}
$script:connectionAuditResults = @{}

# Audit-Befehle definieren (vereinfacht aus easyWSAudit)
$script:auditCommands = @(
    # System-Informationen
    @{Name="Systeminformationen"; Command="Get-ComputerInfo"; Type="PowerShell"; Category="System"},
    @{Name="Betriebssystem Details"; Command="Get-CimInstance Win32_OperatingSystem"; Type="PowerShell"; Category="System"},
    @{Name="Hardware Informationen"; Command="Get-CimInstance Win32_ComputerSystem"; Type="PowerShell"; Category="Hardware"},
    @{Name="CPU Informationen"; Command="Get-CimInstance Win32_Processor"; Type="PowerShell"; Category="Hardware"},
    @{Name="Arbeitsspeicher Details"; Command="Get-CimInstance Win32_PhysicalMemory"; Type="PowerShell"; Category="Hardware"},
    @{Name="Festplatten Informationen"; Command="Get-CimInstance Win32_LogicalDisk"; Type="PowerShell"; Category="Storage"},
    @{Name="Installierte Features und Rollen"; Command="Get-WindowsFeature | Where-Object { `$_.Installed -eq `$true }"; Type="PowerShell"; Category="Features"},
    @{Name="Installierte Programme"; Command="Get-CimInstance Win32_Product | Select-Object Name, Version, Vendor | Sort-Object Name"; Type="PowerShell"; Category="Software"},
    @{Name="Windows Updates"; Command="Get-HotFix | Sort-Object InstalledOn -Descending | Select-Object -First 20"; Type="PowerShell"; Category="Updates"},
    
    # Netzwerk-Informationen
    @{Name="Netzwerkkonfiguration"; Command="Get-NetIPConfiguration"; Type="PowerShell"; Category="Network"},
    @{Name="Netzwerkadapter"; Command="Get-NetAdapter"; Type="PowerShell"; Category="Network"},
    @{Name="Aktive Netzwerkverbindungen"; Command="Get-NetTCPConnection | Where-Object State -eq 'Listen' | Select-Object LocalAddress, LocalPort, OwningProcess"; Type="PowerShell"; Category="Network"},
    @{Name="Firewall Regeln"; Command="Get-NetFirewallRule | Where-Object Enabled -eq 'True' | Select-Object DisplayName, Direction, Action | Sort-Object DisplayName"; Type="PowerShell"; Category="Security"},
    
    # Sicherheits-Informationen
    @{Name="Lokale Benutzer"; Command="Get-LocalUser | Select-Object Name, Enabled, LastLogon, PasswordRequired"; Type="PowerShell"; Category="Security"},
    @{Name="Lokale Gruppen"; Command="Get-LocalGroup | Select-Object Name, Description"; Type="PowerShell"; Category="Security"},
    @{Name="Services (Automatisch)"; Command="Get-Service | Where-Object StartType -eq 'Automatic' | Sort-Object Status, Name"; Type="PowerShell"; Category="Services"},
    @{Name="Services (Laufend)"; Command="Get-Service | Where-Object Status -eq 'Running' | Sort-Object Name"; Type="PowerShell"; Category="Services"},
    
    # Active Directory (falls verfügbar)
    @{Name="AD Domain Controller Status"; Command="Get-ADDomainController -Filter * | Select-Object Name, Site, IPv4Address, OperatingSystem, IsGlobalCatalog, IsReadOnly"; Type="PowerShell"; FeatureName="AD-Domain-Services"; Category="Active-Directory"},
    @{Name="AD Domain Informationen"; Command="Get-ADDomain | Select-Object Name, NetBIOSName, DomainMode, PDCEmulator, RIDMaster, InfrastructureMaster"; Type="PowerShell"; FeatureName="AD-Domain-Services"; Category="Active-Directory"},
    @{Name="AD Forest Informationen"; Command="Get-ADForest | Select-Object Name, ForestMode, DomainNamingMaster, SchemaMaster, Sites, Domains"; Type="PowerShell"; FeatureName="AD-Domain-Services"; Category="Active-Directory"},
    
    # DNS (falls verfügbar)
    @{Name="DNS Server Konfiguration"; Command="Get-DnsServer | Select-Object ComputerName, ZoneScavenging, EnableDnsSec, ServerSetting"; Type="PowerShell"; FeatureName="DNS"; Category="DNS"},
    @{Name="DNS Server Zonen"; Command="Get-DnsServerZone | Select-Object ZoneName, ZoneType, IsAutoCreated, IsDsIntegrated, IsReverseLookupZone"; Type="PowerShell"; FeatureName="DNS"; Category="DNS"},
    @{Name="DNS Forwarders"; Command="Get-DnsServerForwarder"; Type="PowerShell"; FeatureName="DNS"; Category="DNS"},
    
    # DHCP (falls verfügbar)
    @{Name="DHCP Server Konfiguration"; Command="Get-DhcpServerInDC"; Type="PowerShell"; FeatureName="DHCP"; Category="DHCP"},
    @{Name="DHCP IPv4 Bereiche"; Command="Get-DhcpServerv4Scope | Select-Object ScopeId, Name, StartRange, EndRange, SubnetMask, State, LeaseDuration"; Type="PowerShell"; FeatureName="DHCP"; Category="DHCP"},
    @{Name="DHCP Reservierungen"; Command="Get-DhcpServerv4Reservation | Select-Object ScopeId, IPAddress, ClientId, Name, Description"; Type="PowerShell"; FeatureName="DHCP"; Category="DHCP"},
    
    # IIS (falls verfügbar)
    @{Name="IIS Websites"; Command="Get-IISSite | Select-Object Name, Id, State, PhysicalPath"; Type="PowerShell"; FeatureName="Web-Server"; Category="IIS"},
    @{Name="IIS Application Pools"; Command="Get-IISAppPool | Select-Object Name, State, ProcessModel, Recycling"; Type="PowerShell"; FeatureName="Web-Server"; Category="IIS"},
    
    # Hyper-V (falls verfügbar)
    @{Name="Hyper-V Host Informationen"; Command="Get-VMHost | Select-Object ComputerName, LogicalProcessorCount, MemoryCapacity, VirtualMachinePath, VirtualHardDiskPath"; Type="PowerShell"; FeatureName="Hyper-V"; Category="Hyper-V"},
    @{Name="Hyper-V Virtuelle Maschinen"; Command="Get-VM | Select-Object Name, State, CPUUsage, MemoryAssigned, Uptime, Version, Generation"; Type="PowerShell"; FeatureName="Hyper-V"; Category="Hyper-V"}
)

# Verbindungsaudit-Befehle (vereinfacht)
$script:connectionAuditCommands = @(
    @{Name="Etablierte TCP-Verbindungen"; Command="Get-NetTCPConnection -State Established | Select-Object LocalAddress, LocalPort, RemoteAddress, RemotePort, OwningProcess"; Type="PowerShell"; Category="TCP-Connections"},
    @{Name="Lauschende Ports"; Command="Get-NetTCPConnection -State Listen | Select-Object LocalAddress, LocalPort, OwningProcess | Sort-Object LocalPort"; Type="PowerShell"; Category="TCP-Connections"},
    @{Name="UDP-Endpunkte"; Command="Get-NetUDPEndpoint | Select-Object LocalAddress, LocalPort, OwningProcess | Sort-Object LocalPort"; Type="PowerShell"; Category="UDP-Connections"},
    @{Name="Externe Verbindungen"; Command="Get-NetTCPConnection | Where-Object {`$_.RemoteAddress -notmatch '^127\.|^10\.|^192\.168\.|^172\.(1[6-9]|2[0-9]|3[0-1])\.|^::1|^fe80:' -and `$_.RemoteAddress -ne '0.0.0.0' -and `$_.RemoteAddress -ne '::'} | Select-Object LocalAddress, LocalPort, RemoteAddress, RemotePort, State, OwningProcess"; Type="PowerShell"; Category="External-Connections"},
    @{Name="ARP-Cache"; Command="Get-NetNeighbor | Where-Object State -ne 'Unreachable' | Select-Object IPAddress, MacAddress, State, InterfaceAlias | Sort-Object IPAddress"; Type="PowerShell"; Category="Local-Devices"},
    @{Name="DNS-Cache"; Command="Get-DnsClientCache | Where-Object { `$_.Type -eq 'A' } | Select-Object Name, Data, TTL, Section | Sort-Object Name"; Type="PowerShell"; Category="DNS-Info"},
    @{Name="Firewall-Status"; Command="Get-NetFirewallProfile | Select-Object Name, Enabled, DefaultInboundAction, DefaultOutboundAction"; Type="PowerShell"; Category="Firewall-Logs"},
    @{Name="Routing-Tabelle"; Command="Get-NetRoute | Select-Object DestinationPrefix, NextHop, InterfaceAlias, RouteMetric, Protocol | Sort-Object RouteMetric, DestinationPrefix"; Type="PowerShell"; Category="Network-Topology"}
)

# Funktion zum Ausführen eines Audit-Befehls
function Invoke-AuditCommand {
    param(
        [hashtable]$Command,
        [string]$ServerName = $env:COMPUTERNAME
    )
    
    try {
        Write-Log "Führe Audit-Befehl aus: $($Command.Name)" -Level "INFO"
        
        # Prüfe Feature-Abhängigkeiten
        if ($Command.FeatureName) {
            $feature = Get-WindowsFeature -Name $Command.FeatureName -ErrorAction SilentlyContinue
            if (-not $feature -or $feature.InstallState -ne "Installed") {
                return "Feature '$($Command.FeatureName)' ist nicht installiert."
            }
        }
        
        # Führe Befehl aus
        switch ($Command.Type) {
            "PowerShell" {
                $result = Invoke-Expression $Command.Command | Out-String
                return $result.Trim()
            }
            "CMD" {
                $result = cmd /c $Command.Command 2>&1 | Out-String
                return $result.Trim()
            }
            default {
                return "Unbekannter Befehlstyp: $($Command.Type)"
            }
        }
    }
    catch {
        Write-Log "Fehler beim Ausführen von '$($Command.Name)': $($_.Exception.Message)" -Level "ERROR"
        return "FEHLER: $($_.Exception.Message)"
    }
}

# Funktion zum Initialisieren der Audit-UI
function Initialize-AuditUI {
    try {
        Write-Log "Initialisiere Audit-UI" -Level "INFO"
        
        # Setze Standard-Werte
        $txtAuditStatus.Text = "Bereit für Audit-Start"
        $txtAuditStatus.Visibility = "Visible"
        $progressAudit.Value = 0
        $progressAudit.Visibility = "Collapsed"
        $tabAudit.Visibility = "Collapsed"
        
        # Leere vorherige Ergebnisse
        $txtAuditSystem.Text = ""
        $txtAuditNetwork.Text = ""
        $txtAuditSecurity.Text = ""
        $txtAuditConnections.Text = ""
        $txtAuditSummary.Text = ""
        
        Write-Log "Audit-UI erfolgreich initialisiert" -Level "INFO"
    }
    catch {
        Write-Log "Fehler beim Initialisieren der Audit-UI: $($_.Exception.Message)" -Level "ERROR"
    }
}

# Hauptfunktion für das Server-Audit
function Start-ServerAudit {
    try {
        Write-Log "Starte umfassendes Server-Audit" -Level "INFO"
        
        # UI-Updates
        $txtAuditStatus.Text = "Audit läuft..."
        $txtAuditStatus.Visibility = "Visible"
        $progressAudit.Visibility = "Visible"
        $progressAudit.Value = 0
        
        # Bestimme welche Kategorien geprüft werden sollen
        $categoriesToAudit = @()
        if ($chkAuditSystem.IsChecked) { $categoriesToAudit += "System", "Hardware", "Storage", "Software", "Updates" }
        if ($chkAuditNetwork.IsChecked) { $categoriesToAudit += "Network" }
        if ($chkAuditSecurity.IsChecked) { $categoriesToAudit += "Security", "Services" }
        if ($chkAuditActiveDirectory.IsChecked) { $categoriesToAudit += "Active-Directory" }
        if ($chkAuditDNS.IsChecked) { $categoriesToAudit += "DNS" }
        if ($chkAuditDHCP.IsChecked) { $categoriesToAudit += "DHCP" }
        if ($chkAuditIIS.IsChecked) { $categoriesToAudit += "IIS" }
        if ($chkAuditHyperV.IsChecked) { $categoriesToAudit += "Hyper-V" }
        
        # Filtere Befehle nach ausgewählten Kategorien
        $commandsToRun = $script:auditCommands | Where-Object { $_.Category -in $categoriesToAudit }
        
        # Führe Audit-Befehle aus
        $totalCommands = $commandsToRun.Count
        $currentCommand = 0
        $script:auditResults = @{}
        
        foreach ($command in $commandsToRun) {
            $currentCommand++
            $progressPercent = [math]::Round(($currentCommand / $totalCommands) * 80, 0) # 80% für normale Befehle
            $progressAudit.Value = $progressPercent
            
            $txtAuditStatus.Text = "Führe aus: $($command.Name) ($currentCommand/$totalCommands)"
            
            $result = Invoke-AuditCommand -Command $command
            $script:auditResults[$command.Name] = @{
                Category = $command.Category
                Result = $result
                Timestamp = Get-Date
            }
            
            # UI-Update für bessere Responsivität
            [System.Windows.Forms.Application]::DoEvents()
        }
        
        # Verbindungsaudit falls ausgewählt
        if ($chkAuditConnections.IsChecked) {
            $txtAuditStatus.Text = "Führe Netzwerk-Verbindungsaudit durch..."
            $progressAudit.Value = 85
            
            $script:connectionAuditResults = @{}
            foreach ($connCommand in $script:connectionAuditCommands) {
                $result = Invoke-AuditCommand -Command $connCommand
                $script:connectionAuditResults[$connCommand.Name] = $result
            }
        }
        
        # Ergebnisse in UI anzeigen
        $progressAudit.Value = 90
        $txtAuditStatus.Text = "Bereite Ergebnisse auf..."
        
        Show-AuditResults
        
        # Abschluss
        $progressAudit.Value = 100
        $txtAuditStatus.Text = "Audit erfolgreich abgeschlossen"
        $tabAudit.Visibility = "Visible"
        
        # Export-Buttons anzeigen
        $auditExportButtons = $window.FindName("auditExportButtons")
        if ($auditExportButtons) {
            $auditExportButtons.Visibility = "Visible"
        }
        
        Write-Log "Server-Audit erfolgreich abgeschlossen" -Level "SUCCESS"
    }
    catch {
        Write-Log "Fehler beim Server-Audit: $($_.Exception.Message)" -Level "ERROR"
        $txtAuditStatus.Text = "Fehler beim Audit: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Fehler beim Server-Audit: $($_.Exception.Message)", "Audit Fehler", "OK", "Error")
    }
    finally {
        $progressAudit.Visibility = "Collapsed"
    }
}

# Funktion zum Anzeigen der Audit-Ergebnisse
function Show-AuditResults {
    try {
        # System-Ergebnisse
        $systemCategories = @("System", "Hardware", "Storage", "Software", "Updates", "Features")
        $systemResults = ""
        foreach ($category in $systemCategories) {
            $categoryResults = $script:auditResults.GetEnumerator() | Where-Object { $_.Value.Category -eq $category }
            if ($categoryResults) {
                $systemResults += "=== $category ===`r`n"
                foreach ($result in $categoryResults) {
                    $systemResults += "$($result.Key):`r`n$($result.Value.Result)`r`n`r`n"
                }
            }
        }
        $txtAuditSystem.Text = $systemResults
        
        # Netzwerk-Ergebnisse
        $networkResults = ""
        $networkCategories = @("Network")
        foreach ($category in $networkCategories) {
            $categoryResults = $script:auditResults.GetEnumerator() | Where-Object { $_.Value.Category -eq $category }
            if ($categoryResults) {
                $networkResults += "=== $category ===`r`n"
                foreach ($result in $categoryResults) {
                    $networkResults += "$($result.Key):`r`n$($result.Value.Result)`r`n`r`n"
                }
            }
        }
        $txtAuditNetwork.Text = $networkResults
        
        # Sicherheits-Ergebnisse
        $securityResults = ""
        $securityCategories = @("Security", "Services")
        foreach ($category in $securityCategories) {
            $categoryResults = $script:auditResults.GetEnumerator() | Where-Object { $_.Value.Category -eq $category }
            if ($categoryResults) {
                $securityResults += "=== $category ===`r`n"
                foreach ($result in $categoryResults) {
                    $securityResults += "$($result.Key):`r`n$($result.Value.Result)`r`n`r`n"
                }
            }
        }
        $txtAuditSecurity.Text = $securityResults
        
        # Verbindungs-Ergebnisse
        $connectionResults = ""
        foreach ($connResult in $script:connectionAuditResults.GetEnumerator()) {
            $connectionResults += "$($connResult.Key):`r`n$($connResult.Value)`r`n`r`n"
        }
        $txtAuditConnections.Text = $connectionResults
        
        # Zusammenfassung erstellen
        $summary = "=== AUDIT ZUSAMMENFASSUNG ===`r`n"
        $summary += "Server: $env:COMPUTERNAME`r`n"
        $summary += "Zeitpunkt: $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')`r`n"
        $summary += "Geprüfte Kategorien: $($script:auditResults.Values.Category | Sort-Object -Unique | Join-String -Separator ', ')`r`n"
        $summary += "Anzahl Prüfungen: $($script:auditResults.Count)`r`n"
        $summary += "Verbindungsaudit: $(if ($script:connectionAuditResults.Count -gt 0) { 'Ja' } else { 'Nein' })`r`n`r`n"
        
        # Kritische Befunde
        $summary += "=== KRITISCHE BEFUNDE ===`r`n"
        $criticalFindings = @()
        
        # Prüfe auf kritische Sicherheitsprobleme
        foreach ($result in $script:auditResults.GetEnumerator()) {
            if ($result.Value.Result -match "FEHLER|ERROR|CRITICAL|Administrator.*enabled|Guest.*enabled") {
                $criticalFindings += "$($result.Key): Potentielles Sicherheitsproblem erkannt"
            }
        }
        
        if ($criticalFindings.Count -eq 0) {
            $summary += "Keine kritischen Befunde erkannt.`r`n"
        } else {
            $summary += ($criticalFindings -join "`r`n") + "`r`n"
        }
        
        $txtAuditSummary.Text = $summary
        
        Write-Log "Audit-Ergebnisse erfolgreich angezeigt" -Level "INFO"
    }
    catch {
        Write-Log "Fehler beim Anzeigen der Audit-Ergebnisse: $($_.Exception.Message)" -Level "ERROR"
    }
}

# Funktion zum Exportieren der Audit-Ergebnisse als HTML
function Export-AuditToHTML {
    param(
        [string]$FilePath
    )
    
    try {
        Write-Log "Exportiere Audit-Ergebnisse als HTML nach: $FilePath" -Level "INFO"
        
        $html = @"
<!DOCTYPE html>
<html>
<head>
    <title>Windows Server Audit - $env:COMPUTERNAME</title>
    <style>
        body { font-family: 'Segoe UI', Arial, sans-serif; margin: 20px; color: #333; }
        h1 { color: #007ACC; border-bottom: 2px solid #007ACC; padding-bottom: 10px; }
        h2 { color: #007ACC; margin-top: 30px; border-bottom: 1px solid #ccc; padding-bottom: 5px; }
        .summary { background-color: #f5f5f5; padding: 15px; border-left: 4px solid #007ACC; margin: 20px 0; }
        .category { margin: 20px 0; }
        .result { background-color: #f9f9f9; padding: 10px; margin: 10px 0; border-left: 3px solid #007ACC; }
        .result-title { font-weight: bold; color: #333; }
        .result-content { font-family: 'Consolas', monospace; font-size: 12px; white-space: pre-wrap; }
        .footer { margin-top: 50px; text-align: center; font-size: 12px; color: #666; }
    </style>
</head>
<body>
    <h1>Windows Server Audit Report</h1>
    <div class="summary">
        <h2>Zusammenfassung</h2>
        <p><strong>Server:</strong> $env:COMPUTERNAME</p>
        <p><strong>Audit-Zeitpunkt:</strong> $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')</p>
        <p><strong>Anzahl Prüfungen:</strong> $($script:auditResults.Count)</p>
        <p><strong>Verbindungsaudit:</strong> $(if ($script:connectionAuditResults.Count -gt 0) { 'Durchgeführt' } else { 'Nicht durchgeführt' })</p>
    </div>
"@
        
        # System-Kategorie
        $html += "<h2>System-Informationen</h2><div class='category'>"
        $systemCategories = @("System", "Hardware", "Storage", "Software", "Updates", "Features")
        foreach ($category in $systemCategories) {
            $categoryResults = $script:auditResults.GetEnumerator() | Where-Object { $_.Value.Category -eq $category }
            foreach ($result in $categoryResults) {
                $html += "<div class='result'><div class='result-title'>$($result.Key)</div><div class='result-content'>$([System.Web.HttpUtility]::HtmlEncode($result.Value.Result))</div></div>"
            }
        }
        $html += "</div>"
        
        # Netzwerk-Kategorie
        $html += "<h2>Netzwerk-Konfiguration</h2><div class='category'>"
        $networkResults = $script:auditResults.GetEnumerator() | Where-Object { $_.Value.Category -eq "Network" }
        foreach ($result in $networkResults) {
            $html += "<div class='result'><div class='result-title'>$($result.Key)</div><div class='result-content'>$([System.Web.HttpUtility]::HtmlEncode($result.Value.Result))</div></div>"
        }
        $html += "</div>"
        
        # Sicherheits-Kategorie
        $html += "<h2>Sicherheits-Einstellungen</h2><div class='category'>"
        $securityCategories = @("Security", "Services")
        foreach ($category in $securityCategories) {
            $categoryResults = $script:auditResults.GetEnumerator() | Where-Object { $_.Value.Category -eq $category }
            foreach ($result in $categoryResults) {
                $html += "<div class='result'><div class='result-title'>$($result.Key)</div><div class='result-content'>$([System.Web.HttpUtility]::HtmlEncode($result.Value.Result))</div></div>"
            }
        }
        $html += "</div>"
        
        # Verbindungsaudit
        if ($script:connectionAuditResults.Count -gt 0) {
            $html += "<h2>Netzwerk-Verbindungen</h2><div class='category'>"
            foreach ($connResult in $script:connectionAuditResults.GetEnumerator()) {
                $html += "<div class='result'><div class='result-title'>$($connResult.Key)</div><div class='result-content'>$([System.Web.HttpUtility]::HtmlEncode($connResult.Value))</div></div>"
            }
            $html += "</div>"
        }
        
        $html += @"
    <div class="footer">
        <p>Erstellt mit Easy Windows Server Migration Tool | $(Get-Date -Format 'yyyy')</p>
    </div>
</body>
</html>
"@
        
        # HTML-Datei speichern
        Add-Type -AssemblyName System.Web
        $html | Out-File -FilePath $FilePath -Encoding UTF8
        
        Write-Log "Audit-HTML-Export erfolgreich abgeschlossen" -Level "SUCCESS"
        return $true
    }
    catch {
        Write-Log "Fehler beim HTML-Export: $($_.Exception.Message)" -Level "ERROR"
        return $false
    }
}

#endregion

#region Hauptfunktionen für die Migration
# Funktion zur Analyse der AD-Umgebung
function Get-ADEnvironmentInfo {
    try {
        Write-Log "Starte AD-Umgebungsanalyse" -Level "INFO"
        
        $results = @{
            DomainControllers = @()
            FSMORoles = @()
            ServiceServers = @()
            ServerRoles = @()
            DnsConfig = @()
            DhcpConfig = @()
            ADStructure = @()
            Summary = ""
        }
        
        # Domänencontroller ermitteln
        try {
            if (Get-Module -ListAvailable -Name ActiveDirectory) {
                Import-Module ActiveDirectory -ErrorAction Stop
                
                $dcs = Get-ADDomainController -Filter * -ErrorAction Stop
                foreach ($dc in $dcs) {
                    $isOnline = Test-Connection -ComputerName $dc.Name -Count 1 -Quiet -ErrorAction SilentlyContinue
                    
                    $results.DomainControllers += [PSCustomObject]@{
                        Name = $dc.Name
                        Site = $dc.Site
                        IPAddress = $dc.IPv4Address
                        OperatingSystem = $dc.OperatingSystem
                        OSVersion = $dc.OperatingSystemVersion
                        IsGlobalCatalog = $dc.IsGlobalCatalog
                        IsOnline = $isOnline
                    }
                }
                
                # FSMO-Rollen ermitteln
                $forest = Get-ADForest -ErrorAction SilentlyContinue
                $domain = Get-ADDomain -ErrorAction SilentlyContinue
                
                if ($forest) {
                    $results.FSMORoles += [PSCustomObject]@{ RoleName = "Schema Master"; RoleOwner = $forest.SchemaMaster }
                    $results.FSMORoles += [PSCustomObject]@{ RoleName = "Domain Naming Master"; RoleOwner = $forest.DomainNamingMaster }
                }
                
                if ($domain) {
                    $results.FSMORoles += [PSCustomObject]@{ RoleName = "PDC Emulator"; RoleOwner = $domain.PDCEmulator }
                    $results.FSMORoles += [PSCustomObject]@{ RoleName = "RID Master"; RoleOwner = $domain.RIDMaster }
                    $results.FSMORoles += [PSCustomObject]@{ RoleName = "Infrastructure Master"; RoleOwner = $domain.InfrastructureMaster }
                }
            }
        }
        catch {
            Write-Log "Fehler bei der DC-Analyse: $($_.Exception.Message)" -Level "WARNING"
        }
        
        # DNS/DHCP-Server ermitteln
        try {
            foreach ($dc in $results.DomainControllers) {
                $dnsService = Get-Service -ComputerName $dc.Name -Name "DNS" -ErrorAction SilentlyContinue
                $dhcpService = Get-Service -ComputerName $dc.Name -Name "DHCPServer" -ErrorAction SilentlyContinue
                
                $results.ServiceServers += [PSCustomObject]@{
                    ServerName = $dc.Name
                    HasDNS = ($null -ne $dnsService)
                    HasDHCP = ($null -ne $dhcpService)
                    DNSStatus = if ($dnsService) { $dnsService.Status } else { "Not Installed" }
                    DHCPStatus = if ($dhcpService) { $dhcpService.Status } else { "Not Installed" }
                    OperatingSystem = $dc.OperatingSystem
                }
            }
        }
        catch {
            Write-Log "Fehler bei der Dienste-Analyse: $($_.Exception.Message)" -Level "WARNING"
        }
        
        # Zusammenfassung erstellen
        $dcCount = $results.DomainControllers.Count
        $onlineDCs = ($results.DomainControllers | Where-Object { $_.IsOnline }).Count
        $dnsServers = ($results.ServiceServers | Where-Object { $_.HasDNS }).Count
        $dhcpServers = ($results.ServiceServers | Where-Object { $_.HasDHCP }).Count
        
        $results.Summary = @"
Umgebungsanalyse Zusammenfassung:

Domänencontroller: $dcCount gefunden, $onlineDCs online
DNS-Server: $dnsServers
DHCP-Server: $dhcpServers
FSMO-Rollen: $($results.FSMORoles.Count) identifiziert

Analyse abgeschlossen am: $(Get-Date -Format "dd.MM.yyyy HH:mm:ss")
"@
        
        Write-Log "AD-Umgebungsanalyse erfolgreich abgeschlossen" -Level "SUCCESS"
        return $results
    }
    catch {
        Write-Log "Kritischer Fehler bei der AD-Umgebungsanalyse: $($_.Exception.Message)" -Level "ERROR"
        return $null
    }
}

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
        
        # Erweiterte Informationen (falls verfügbar)
        if ($AnalysisResults.ServerRoles) {
            Write-Log "Serverrollen-Daten verfügbar: $($AnalysisResults.ServerRoles.Count)" -Level "INFO"
        }
        if ($AnalysisResults.DnsConfig) {
            Write-Log "DNS-Konfigurationsdaten verfügbar: $($AnalysisResults.DnsConfig.Count)" -Level "INFO"
        }
        
        # Zusammenfassung aktualisieren
        $txtSummary.Text = $AnalysisResults.Summary
        
        # Tab-Steuerung anzeigen und Export-Button aktivieren
        $tabDiscovery.Visibility = "Visible"
        $exportButtons = $window.FindName("exportButtons")
        if ($exportButtons) {
            $exportButtons.Visibility = "Visible"
        }
        
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
                $htmlDnsConfig += "<tr><td>$($dnsServer.ServerName)</td><td>$($dnsServer.Zones)</td><td>$($dnsServer.Forwarders)</td><td>$($dnsServer.ZoneReplication)</td><td class='$recursionClass'>$($dnsServer.RecursionEnabled)</td></tr>"
            }
            $htmlDnsConfig += "</table>"
        } else {
            $htmlDnsConfig += "<p>Keine DNS-Konfiguration gefunden.</p>"
        }

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
finally {
    Write-Log "Anwendung wird beendet" -Level "INFO"
}
