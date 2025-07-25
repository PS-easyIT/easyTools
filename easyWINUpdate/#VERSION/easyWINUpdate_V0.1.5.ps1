<#
.SYNOPSIS
    easyWINUpdate - Windows Update Verwaltungstool mit GUI
.DESCRIPTION
    Dieses Script ermöglicht die Verwaltung von Windows Updates auf Windows 11 und Windows Server 2019-2022.
    Es bietet eine moderne XAML-GUI zur Anzeige, Installation und Deinstallation von Updates sowie zur Verwaltung
    der Update-Quellen und WSUS-Einstellungen.
.NOTES
    Version:        0.1.4
    Author:         easyIT
    Creation Date:  27.05.2025
#>

# Fehlercode-Mapping mit bekannten Loesungen
$errorCodes = @{
    "0x80240034" = @{
        "Beschreibung" = "WU_E_DOWNLOAD_FAILED - Update-Download ist fehlgeschlagen";
        "Loesungen" = @(
            "Windows Update-Dienste neu starten",
            "Temporaere Dateien und Cache loeschen",
            "Netzwerkverbindung ueberpruefen",
            "Windows Update-Komponenten zuruecksetzen"
        )
    };
    "0x8024001F" = @{
        "Beschreibung" = "WU_E_NO_CONNECTION - Keine Verbindung zum Update-Service";
        "Loesungen" = @(
            "Netzwerkverbindung ueberpruefen",
            "Firewall- und Proxy-Einstellungen pruefen",
            "Neustart des Netzwerkadapters",
            "Windows Update-Dienste neu starten"
        )
    };
    "0x80070020" = @{
        "Beschreibung" = "InstallFileLocked - Eine Datei, die aktualisiert werden soll, ist gesperrt";
        "Loesungen" = @(
            "Alle laufenden Programme schliessen",
            "Windows Update-Komponenten zuruecksetzen",
            "Temporaere Dateien und Cache loeschen",
            "System neustarten"
        )
    };
    "0x80240022" = @{
        "Beschreibung" = "WU_E_ALL_UPDATES_FAILED - Alle Updates konnten nicht installiert werden";
        "Loesungen" = @(
            "Windows Update-Komponenten zuruecksetzen",
            "System mit sauberem Boot starten",
            "Windows-Systemdateien reparieren",
            "Windows Update-Dienste neu starten"
        )
    }
}

#region Requires
# Module PSWindowsUpdate wird benötigt
if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
    try {
        Install-Module -Name PSWindowsUpdate -Force -Scope CurrentUser -ErrorAction Stop
        Import-Module PSWindowsUpdate
        Write-Host "PSWindowsUpdate-Modul wurde installiert und importiert." -ForegroundColor Green
    } catch {
        Write-Host "Fehler beim Installieren des PSWindowsUpdate-Moduls: $_" -ForegroundColor Red
        Write-Host "Bitte führen Sie 'Install-Module -Name PSWindowsUpdate -Force' manuell mit administrativen Rechten aus." -ForegroundColor Yellow
        exit
    }
} else {
    Import-Module PSWindowsUpdate
}
#endregion

# WPF-Assemblies laden
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

#region XAML GUI Definition
[xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="easyWINUpdate - Windows Update Management" Height="950" Width="1400" 
    WindowStartupLocation="CenterScreen" ResizeMode="CanResize" 
    Background="#F0F0F0">
    
    <Window.Resources>
        <Style x:Key="NavButtonStyle" TargetType="RadioButton">
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="RadioButton">
                        <Border x:Name="border" Background="Transparent" BorderThickness="4,0,0,0" BorderBrush="Transparent" Padding="{TemplateBinding Padding}">
                            <ContentPresenter VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#F0F0F0"/>
                            </Trigger>
                            <Trigger Property="IsChecked" Value="True">
                                <Setter TargetName="border" Property="BorderBrush" Value="#0078D7"/>
                                <Setter TargetName="border" Property="Background" Value="#F0F0F0"/>
                                <Setter Property="Foreground" Value="#0078D7"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>
    
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="60"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="40"/>
        </Grid.RowDefinitions>
        
        <!-- Header -->
        <Border Grid.Row="0" Background="#0078D7" Padding="15,0">
            <Grid>
                <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                    <TextBlock Text="easyWINUpdate" Foreground="White" FontSize="24" FontWeight="Bold" VerticalAlignment="Center"/>
                    <TextBlock Text="v0.1.4" Foreground="#CCFFFFFF" FontSize="14" Margin="10,0,0,0" VerticalAlignment="Center"/>
                </StackPanel>
                <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Center">
                    <TextBlock x:Name="computerNameLabel" Text="Computer: " Foreground="White" FontSize="14" VerticalAlignment="Center"/>
                    <TextBlock x:Name="computerName" Text="..." Foreground="White" FontSize="14" FontWeight="SemiBold" VerticalAlignment="Center"/>
                </StackPanel>
            </Grid>
        </Border>
        
        <!-- Content Area -->
        <Grid Grid.Row="1">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="250"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            
            <!-- Navigation Panel -->
            <Border Background="#F9F9F9" BorderBrush="#DDDDDD" BorderThickness="0,0,1,0">
                <StackPanel Margin="0,20,0,0">
                    <RadioButton x:Name="navUpdateStatus" Content="Update Status" GroupName="Navigation" 
                                IsChecked="True" Height="50" FontSize="14" Padding="15,0,0,0"
                                Foreground="#333333" Style="{StaticResource NavButtonStyle}"/>
                    
                    <RadioButton x:Name="navInstalledUpdates" Content="Installed Updates" GroupName="Navigation" 
                                Height="50" FontSize="14" Padding="15,0,0,0"
                                Foreground="#333333" Style="{StaticResource NavButtonStyle}"/>
                    
                    <RadioButton x:Name="navAvailableUpdates" Content="Available Updates" GroupName="Navigation" 
                                Height="50" FontSize="14" Padding="15,0,0,0"
                                Foreground="#333333" Style="{StaticResource NavButtonStyle}"/>
                    
                    <RadioButton x:Name="navWSUSSettings" Content="WSUS Settings" GroupName="Navigation" 
                                Height="50" FontSize="14" Padding="15,0,0,0"
                                Foreground="#333333" Style="{StaticResource NavButtonStyle}"/>
                    
                    <RadioButton x:Name="navTroubleshooting" Content="Troubleshooting" GroupName="Navigation" 
                                Height="50" FontSize="14" Padding="15,0,0,0"
                                Foreground="#333333" Style="{StaticResource NavButtonStyle}"/>
                </StackPanel>
            </Border>
            
            <!-- Content Pages -->
            <Grid Grid.Column="1" Margin="20">
                <!-- Update Status Page -->
                <Grid x:Name="updateStatusPage" Visibility="Visible">
                    <StackPanel>
                        <TextBlock Text="Windows Update Status" FontSize="22" FontWeight="SemiBold" Margin="0,0,0,20"/>
                        
                        <Grid Margin="0,0,0,20">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="Auto"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>
                            
                            <TextBlock Grid.Row="0" Grid.Column="0" Text="Windows Version:" FontWeight="SemiBold" Margin="0,5,10,5"/>
                            <TextBlock Grid.Row="0" Grid.Column="1" x:Name="txtWindowsVersion" Text="Loading..." Margin="0,5,0,5"/>
                            
                            <TextBlock Grid.Row="1" Grid.Column="0" Text="Update Source:" FontWeight="SemiBold" Margin="0,5,10,5"/>
                            <TextBlock Grid.Row="1" Grid.Column="1" x:Name="txtUpdateSource" Text="Loading..." Margin="0,5,0,5"/>
                            
                            <TextBlock Grid.Row="2" Grid.Column="0" Text="Last Update:" FontWeight="SemiBold" Margin="0,5,10,5"/>
                            <TextBlock Grid.Row="2" Grid.Column="1" x:Name="txtLastUpdate" Text="Loading..." Margin="0,5,0,5"/>
                            
                            <TextBlock Grid.Row="3" Grid.Column="0" Text="Update Status:" FontWeight="SemiBold" Margin="0,5,10,5"/>
                            <TextBlock Grid.Row="3" Grid.Column="1" x:Name="txtUpdateStatus" Text="Loading..." Margin="0,5,0,5"/>
                            
                            <TextBlock Grid.Row="4" Grid.Column="0" Text="Installed Updates:" FontWeight="SemiBold" Margin="0,5,10,5"/>
                            <TextBlock Grid.Row="4" Grid.Column="1" x:Name="txtInstalledUpdatesCount" Text="Loading..." Margin="0,5,0,5"/>
                        </Grid>
                        
                        <Button x:Name="btnCheckForUpdates" Content="Check for Updates" Padding="15,8" 
                                Background="#0078D7" Foreground="White" BorderThickness="0"
                                HorizontalAlignment="Left" Margin="0,10,0,20"/>
                        
                        <Border Background="#F5F5F5" BorderBrush="#DDDDDD" BorderThickness="1" Padding="15" CornerRadius="4">
                            <StackPanel>
                                <TextBlock Text="Windows Update Service Status" FontSize="16" FontWeight="SemiBold" Margin="0,0,0,10"/>
                                <Grid>
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="Auto"/>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
                                    <Grid.RowDefinitions>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                    </Grid.RowDefinitions>
                                    
                                    <TextBlock Grid.Row="0" Grid.Column="0" Text="Windows Update Service:" FontWeight="SemiBold" Margin="0,5,10,5"/>
                                    <TextBlock Grid.Row="0" Grid.Column="1" x:Name="txtUpdateService" Text="Loading..." Margin="0,5,0,5"/>
                                    <Button Grid.Row="0" Grid.Column="2" x:Name="btnRestartService" Content="Restart Service" Padding="10,5"/>
                                    
                                    <TextBlock Grid.Row="1" Grid.Column="0" Text="BITS Service:" FontWeight="SemiBold" Margin="0,5,10,5"/>
                                    <TextBlock Grid.Row="1" Grid.Column="1" x:Name="txtBITSService" Text="Loading..." Margin="0,5,0,5"/>
                                    <Button Grid.Row="1" Grid.Column="2" x:Name="btnRestartBITS" Content="Restart Service" Padding="10,5"/>
                                </Grid>
                            </StackPanel>
                        </Border>
                    </StackPanel>
                </Grid>
                
                <!-- Installed Updates Page -->
                <Grid x:Name="installedUpdatesPage" Visibility="Collapsed">
                    <StackPanel>
                        <TextBlock Text="Installed Windows Updates" FontSize="22" FontWeight="SemiBold" Margin="0,0,0,20"/>
                        
                        <Grid Margin="0,0,0,10">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            
                            <TextBox x:Name="txtSearchInstalled" Grid.Column="0" Padding="8" Margin="0,0,10,0" 
                                    BorderThickness="1" BorderBrush="#AAAAAA">
                                <TextBox.Style>
                                    <Style TargetType="TextBox">
                                        <Style.Resources>
                                            <VisualBrush x:Key="HintBrush" TileMode="None" Opacity="0.5" Stretch="None" AlignmentX="Left">
                                                <VisualBrush.Visual>
                                                    <TextBlock Text="Filter by KB number or title..." FontStyle="Italic" />
                                                </VisualBrush.Visual>
                                            </VisualBrush>
                                        </Style.Resources>
                                        <Style.Triggers>
                                            <Trigger Property="Text" Value="">
                                                <Setter Property="Background" Value="{StaticResource HintBrush}" />
                                            </Trigger>
                                            <Trigger Property="IsKeyboardFocused" Value="True">
                                                <Setter Property="Background" Value="White" />
                                            </Trigger>
                                        </Style.Triggers>
                                    </Style>
                                </TextBox.Style>
                            </TextBox>
                            <Button x:Name="btnSearchInstalled" Grid.Column="1" Content="Search" Padding="15,8" 
                                    Background="#0078D7" Foreground="White" BorderThickness="0"/>
                        </Grid>
                        
                        <DataGrid x:Name="dgInstalledUpdates" AutoGenerateColumns="False" HeadersVisibility="Column"
                                IsReadOnly="True" BorderThickness="1" BorderBrush="#DDDDDD"
                                Background="White" RowBackground="White" AlternatingRowBackground="#F8F8F8"
                                VerticalScrollBarVisibility="Auto" Height="400">
                            <DataGrid.Columns>
                                <DataGridTextColumn Header="KB Number" Binding="{Binding HotFixID}" Width="100"/>
                                <DataGridTextColumn Header="Description" Binding="{Binding Description}" Width="*"/>
                                <DataGridTextColumn Header="Installed On" Binding="{Binding InstalledOn}" Width="150"/>
                                <DataGridTemplateColumn Header="Actions" Width="100">
                                    <DataGridTemplateColumn.CellTemplate>
                                        <DataTemplate>
                                            <Button Content="Remove" Tag="{Binding HotFixID}" x:Name="btnRemoveUpdate"
                                                    Padding="10,5" Margin="5" Background="#E81123" Foreground="White" BorderThickness="0"/>
                                        </DataTemplate>
                                    </DataGridTemplateColumn.CellTemplate>
                                </DataGridTemplateColumn>
                            </DataGrid.Columns>
                        </DataGrid>
                        
                        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0">
                            <Button x:Name="btnRefreshInstalled" Content="Refresh" Padding="15,8" Margin="0,0,10,0"
                                    Background="#0078D7" Foreground="White" BorderThickness="0"/>
                            <Button x:Name="btnExportInstalled" Content="Export List" Padding="15,8" 
                                    Background="#107C10" Foreground="White" BorderThickness="0"/>
                        </StackPanel>
                    </StackPanel>
                </Grid>
                
                <!-- Available Updates Page -->
                <Grid x:Name="availableUpdatesPage" Visibility="Collapsed">
                    <StackPanel>
                        <TextBlock Text="Available Windows Updates" FontSize="22" FontWeight="SemiBold" Margin="0,0,0,20"/>
                        
                        <Grid Margin="0,0,0,10">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            
                            <StackPanel Grid.Column="0" Orientation="Horizontal">
                                <CheckBox x:Name="chkCriticalUpdates" Content="Critical Updates" IsChecked="True" Margin="0,0,15,0"/>
                                <CheckBox x:Name="chkSecurityUpdates" Content="Security Updates" IsChecked="True" Margin="0,0,15,0"/>
                                <CheckBox x:Name="chkDefinitionUpdates" Content="Definition Updates" IsChecked="True" Margin="0,0,15,0"/>
                                <CheckBox x:Name="chkFeatureUpdates" Content="Feature Updates" IsChecked="False"/>
                            </StackPanel>
                            
                            <Button x:Name="btnSearchAvailable" Grid.Column="1" Content="Search for Updates" Padding="15,8" Margin="10,0"
                                    Background="#0078D7" Foreground="White" BorderThickness="0"/>
                            <Button x:Name="btnInstallAll" Grid.Column="2" Content="Install All" Padding="15,8" 
                                    Background="#107C10" Foreground="White" BorderThickness="0"/>
                        </Grid>
                        
                        <ProgressBar x:Name="progressUpdates" Height="5" Margin="0,0,0,10" Visibility="Collapsed"/>
                        
                        <DataGrid x:Name="dgAvailableUpdates" AutoGenerateColumns="False" HeadersVisibility="Column"
                                IsReadOnly="False" BorderThickness="1" BorderBrush="#DDDDDD" 
                                Background="White" RowBackground="White" AlternatingRowBackground="#F8F8F8"
                                VerticalScrollBarVisibility="Auto" Height="400">
                            <DataGrid.Columns>
                                <DataGridCheckBoxColumn Header="Select" Binding="{Binding IsSelected}" Width="80"/>
                                <DataGridTextColumn Header="KB Number" Binding="{Binding KBArticleID}" Width="100"/>
                                <DataGridTextColumn Header="Title" Binding="{Binding Title}" Width="*"/>
                                <DataGridTextColumn Header="Category" Binding="{Binding Category}" Width="120"/>
                                <DataGridTextColumn Header="Size" Binding="{Binding Size}" Width="100"/>
                                <DataGridTemplateColumn Header="Actions" Width="120">
                                    <DataGridTemplateColumn.CellTemplate>
                                        <DataTemplate>
                                            <Button Content="Install" Tag="{Binding Identity}" x:Name="btnInstallUpdate"
                                                    Padding="10,5" Margin="5" Background="#107C10" Foreground="White" BorderThickness="0"/>
                                        </DataTemplate>
                                    </DataGridTemplateColumn.CellTemplate>
                                </DataGridTemplateColumn>
                            </DataGrid.Columns>
                        </DataGrid>
                        
                        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0">
                            <Button x:Name="btnInstallSelected" Content="Install Selected" Padding="15,8" Margin="0,0,10,0"
                                    Background="#107C10" Foreground="White" BorderThickness="0"/>
                            <Button x:Name="btnDownloadSelected" Content="Download Selected" Padding="15,8" 
                                    Background="#0078D7" Foreground="White" BorderThickness="0"/>
                        </StackPanel>
                    </StackPanel>
                </Grid>
                
                <!-- WSUS Settings Page -->
                <Grid x:Name="wsusSettingsPage" Visibility="Collapsed">
                    <StackPanel>
                        <TextBlock Text="WSUS Settings" FontSize="22" FontWeight="SemiBold" Margin="0,0,0,20"/>
                        
                        <Border Background="#F5F5F5" BorderBrush="#DDDDDD" BorderThickness="1" Padding="15" CornerRadius="4" Margin="0,0,0,20">
                            <StackPanel>
                                <TextBlock Text="Current WSUS Configuration" FontSize="16" FontWeight="SemiBold" Margin="0,0,0,10"/>
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
                                        <RowDefinition Height="Auto"/>
                                    </Grid.RowDefinitions>
                                    
                                    <TextBlock Grid.Row="0" Grid.Column="0" Text="WSUS Server:" FontWeight="SemiBold" Margin="0,5,10,5"/>
                                    <TextBlock Grid.Row="0" Grid.Column="1" x:Name="txtWSUSServer" Text="Loading..." Margin="0,5,0,5"/>
                                    
                                    <TextBlock Grid.Row="1" Grid.Column="0" Text="WSUS Status:" FontWeight="SemiBold" Margin="0,5,10,5"/>
                                    <TextBlock Grid.Row="1" Grid.Column="1" x:Name="txtWSUSStatus" Text="Loading..." Margin="0,5,0,5"/>
                                    
                                    <TextBlock Grid.Row="2" Grid.Column="0" Text="Target Group:" FontWeight="SemiBold" Margin="0,5,10,5"/>
                                    <TextBlock Grid.Row="2" Grid.Column="1" x:Name="txtWSUSTargetGroup" Text="Loading..." Margin="0,5,0,5"/>
                                    
                                    <TextBlock Grid.Row="3" Grid.Column="0" Text="Configuration Source:" FontWeight="SemiBold" Margin="0,5,10,5"/>
                                    <TextBlock Grid.Row="3" Grid.Column="1" x:Name="txtWSUSConfigSource" Text="Loading..." Margin="0,5,0,5"/>
                                    
                                    <TextBlock Grid.Row="4" Grid.Column="0" Text="Last Check:" FontWeight="SemiBold" Margin="0,5,10,5"/>
                                    <TextBlock Grid.Row="4" Grid.Column="1" x:Name="txtWSUSLastCheck" Text="Loading..." Margin="0,5,0,5"/>
                                </Grid>
                            </StackPanel>
                        </Border>
                        
                        <Border Background="#F0F7FF" BorderBrush="#99CCF9" BorderThickness="1" Padding="15" CornerRadius="4" Margin="0,0,0,20">
                            <StackPanel>
                                <TextBlock Text="Reset WSUS Settings" FontSize="16" FontWeight="SemiBold" Margin="0,0,0,10"/>
                                <TextBlock TextWrapping="Wrap" Margin="0,0,0,15">
                                    Resetting the WSUS settings will cause the client to receive updates directly from Microsoft again.
                                    The existing WSUS configuration will be removed and the Windows Update services will be restarted.
                                </TextBlock>
                                <Button x:Name="btnResetWSUS" Content="Reset WSUS Settings" Padding="15,8" 
                                        Background="#E81123" Foreground="White" BorderThickness="0" HorizontalAlignment="Left"/>
                            </StackPanel>
                        </Border>
                        
                        <Border Background="#F5FFF0" BorderBrush="#99F9CC" BorderThickness="1" Padding="15" CornerRadius="4">
                            <StackPanel>
                                <TextBlock Text="Check WSUS Connection" FontSize="16" FontWeight="SemiBold" Margin="0,0,0,10"/>
                                <TextBlock TextWrapping="Wrap" Margin="0,0,0,15">
                                    Checks the connection to the configured WSUS server and synchronizes client settings.
                                </TextBlock>
                                <StackPanel Orientation="Horizontal">
                                    <Button x:Name="btnCheckWSUSConn" Content="Check WSUS Connection" Padding="15,8" 
                                            Background="#0078D7" Foreground="White" BorderThickness="0" Margin="0,0,10,0"/>
                                    <Button x:Name="btnDetectNow" Content="Start Update Detection" Padding="15,8" 
                                            Background="#107C10" Foreground="White" BorderThickness="0"/>
                                </StackPanel>
                            </StackPanel>
                        </Border>
                    </StackPanel>
                </Grid>
                
                <!-- Troubleshooting Page -->
                <Grid x:Name="troubleshootingPage" Visibility="Collapsed">
                    <ScrollViewer MaxHeight="750" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
                        <StackPanel Margin="15,0">
                            <TextBlock Text="Windows Update Troubleshooting" FontSize="20" FontWeight="SemiBold" Margin="0,0,0,10"/>
                            
                            <!-- Warning -->
                            <Border Background="#FFF4F4" BorderBrush="#FFCECE" BorderThickness="1" Padding="10" CornerRadius="4" Margin="0,0,0,10">
                                <StackPanel>
                                    <TextBlock TextWrapping="Wrap" Margin="0,0,0,5">
                                        <Bold>Warning:</Bold> The following actions can reset or repair Windows Update components. It is recommended to create a system restore point first.
                                    </TextBlock>
                                    <Button x:Name="btnCreateRestorePoint" Content="Create System Restore Point" Padding="10,5" 
                                            Background="#0078D7" Foreground="White" BorderThickness="0"
                                            HorizontalAlignment="Right" Margin="0,5,0,0"/>
                                </StackPanel>
                            </Border>
                            
                            <!-- Grid for standard functions -->
                            <TextBlock Text="Standard Troubleshooting" FontSize="16" FontWeight="SemiBold" Margin="0,5,0,10"/>
                            <Grid Margin="0,0,0,10">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>
                                
                                <!-- Reset components -->
                                <Border Grid.Column="0" Grid.Row="0" Background="#F5F5F5" BorderBrush="#DDDDDD" BorderThickness="1" Padding="10" CornerRadius="4" Margin="0,0,5,10">
                                    <StackPanel>
                                        <TextBlock Text="Reset Components" FontSize="14" FontWeight="SemiBold" Margin="0,0,0,5"/>
                                        <TextBlock TextWrapping="Wrap" Margin="0,0,0,5" Height="50">
                                            Resets Windows Update components and services to their default state.
                                        </TextBlock>
                                        <Button x:Name="btnResetComponents" Content="Reset" Padding="10,5" 
                                                Background="#E81123" Foreground="White" BorderThickness="0"
                                                HorizontalAlignment="Right"/>
                                    </StackPanel>
                                </Border>
                                
                                <!-- Check system files -->
                                <Border Grid.Column="1" Grid.Row="0" Background="#F5F5F5" BorderBrush="#DDDDDD" BorderThickness="1" Padding="10" CornerRadius="4" Margin="5,0,5,10">
                                    <StackPanel>
                                        <TextBlock Text="Check System Files" FontSize="14" FontWeight="SemiBold" Margin="0,0,0,5"/>
                                        <TextBlock TextWrapping="Wrap" Margin="0,0,0,5" Height="50">
                                            Runs SFC and DISM tools to scan and repair corrupted system files that may affect updates.
                                        </TextBlock>
                                        <Button x:Name="btnCheckSystemFiles" Content="Check" Padding="10,5" 
                                                Background="#0078D7" Foreground="White" BorderThickness="0"
                                                HorizontalAlignment="Right"/>
                                    </StackPanel>
                                </Border>
                                
                                <!-- Clear update history -->
                                <Border Grid.Column="2" Grid.Row="0" Background="#F5F5F5" BorderBrush="#DDDDDD" BorderThickness="1" Padding="10" CornerRadius="4" Margin="5,0,5,10">
                                    <StackPanel>
                                        <TextBlock Text="Clear Update History" FontSize="14" FontWeight="SemiBold" Margin="0,0,0,5"/>
                                        <TextBlock TextWrapping="Wrap" Margin="0,0,0,5" Height="50">
                                            Clears update history, caches, and temporary files to resolve update installation issues.
                                        </TextBlock>
                                        <Button x:Name="btnClearUpdateHistory" Content="Clear" Padding="10,5" 
                                                Background="#0078D7" Foreground="White" BorderThickness="0"
                                                HorizontalAlignment="Right"/>
                                    </StackPanel>
                                </Border>
                                
                                <!-- Re-register DLLs -->
                                <Border Grid.Column="3" Grid.Row="0" Background="#F5F5F5" BorderBrush="#DDDDDD" BorderThickness="1" Padding="10" CornerRadius="4" Margin="5,0,0,10">
                                    <StackPanel>
                                        <TextBlock Text="Re-register DLLs" FontSize="14" FontWeight="SemiBold" Margin="0,0,0,5"/>
                                        <TextBlock TextWrapping="Wrap" Margin="0,0,0,5" Height="50">
                                            Re-registers critical Windows Update DLL files to fix component registration issues.
                                        </TextBlock>
                                        <Button x:Name="btnRegisterDLLs" Content="Register" Padding="10,5" 
                                                Background="#0078D7" Foreground="White" BorderThickness="0"
                                                HorizontalAlignment="Right"/>
                                    </StackPanel>
                                </Border>
                                
                                <!-- Remove stuck updates -->
                                <Border Grid.Column="0" Grid.Row="1" Background="#F5F5F5" BorderBrush="#DDDDDD" BorderThickness="1" Padding="10" CornerRadius="4" Margin="0,0,5,10">
                                    <StackPanel>
                                        <TextBlock Text="Stuck Updates" FontSize="14" FontWeight="SemiBold" Margin="0,0,0,5"/>
                                        <TextBlock TextWrapping="Wrap" Margin="0,0,0,5" Height="50">
                                            Removes stuck BITS jobs and pending updates that may be blocking the update process.
                                        </TextBlock>
                                        <Button x:Name="btnClearStuckUpdates" Content="Remove" Padding="10,5" 
                                                Background="#0078D7" Foreground="White" BorderThickness="0"
                                                HorizontalAlignment="Right"/>
                                    </StackPanel>
                                </Border>

                                <!-- Additional box 1 -->
                                <Border Grid.Column="1" Grid.Row="1" Background="#F5F5F5" BorderBrush="#DDDDDD" BorderThickness="1" Padding="10" CornerRadius="4" Margin="5,0,5,10">
                                    <StackPanel>
                                        <TextBlock Text="Clean Temp Files" FontSize="14" FontWeight="SemiBold" Margin="0,0,0,5"/>
                                        <TextBlock TextWrapping="Wrap" Margin="0,0,0,5" Height="50">
                                            Removes temporary files that might interfere with the update process.
                                        </TextBlock>
                                        <Button x:Name="btnCleanTemp" Content="Clean" Padding="10,5" 
                                                Background="#0078D7" Foreground="White" BorderThickness="0"
                                                HorizontalAlignment="Right"/>
                                    </StackPanel>
                                </Border>

                                <!-- Additional box 2 -->
                                <Border Grid.Column="2" Grid.Row="1" Background="#F5F5F5" BorderBrush="#DDDDDD" BorderThickness="1" Padding="10" CornerRadius="4" Margin="5,0,5,10">
                                    <StackPanel>
                                        <TextBlock Text="Reset WinSock" FontSize="14" FontWeight="SemiBold" Margin="0,0,0,5"/>
                                        <TextBlock TextWrapping="Wrap" Margin="0,0,0,5" Height="50">
                                            Resets network configuration to fix network-related update issues.
                                        </TextBlock>
                                        <Button x:Name="btnResetWinsock" Content="Reset" Padding="10,5" 
                                                Background="#0078D7" Foreground="White" BorderThickness="0"
                                                HorizontalAlignment="Right"/>
                                    </StackPanel>
                                </Border>

                                <!-- Additional box 3 -->
                                <Border Grid.Column="3" Grid.Row="1" Background="#F5F5F5" BorderBrush="#DDDDDD" BorderThickness="1" Padding="10" CornerRadius="4" Margin="5,0,0,10">
                                    <StackPanel>
                                        <TextBlock Text="Check Disk" FontSize="14" FontWeight="SemiBold" Margin="0,0,0,5"/>
                                        <TextBlock TextWrapping="Wrap" Margin="0,0,0,5" Height="50">
                                            Checks disk for errors that might cause update installation failures.
                                        </TextBlock>
                                        <Button x:Name="btnCheckDisk" Content="Check" Padding="10,5" 
                                                Background="#0078D7" Foreground="White" BorderThickness="0"
                                                HorizontalAlignment="Right"/>
                                    </StackPanel>
                                </Border>
                            </Grid>
                            
                            <!-- Grid for advanced functions -->
                            <TextBlock Text="Advanced Troubleshooting Options" FontSize="16" FontWeight="SemiBold" Margin="0,5,0,10"/>
                            <Grid>
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
                                
                                <!-- Error diagnosis -->
                                <Border Grid.Column="0" Grid.Row="0" Background="#F5F5F5" BorderBrush="#DDDDDD" BorderThickness="1" Padding="10" CornerRadius="4" Margin="0,0,5,10">
                                    <StackPanel>
                                        <TextBlock Text="Error Diagnosis" FontSize="14" FontWeight="SemiBold" Margin="0,0,0,5"/>
                                        <TextBlock TextWrapping="Wrap" Margin="0,0,0,5" Height="50">
                                            Analyzes Windows Update error codes and provides targeted solutions for specific issues.
                                        </TextBlock>
                                        <Button x:Name="btnDiagnoseError" Content="Diagnose" Padding="10,5" 
                                                Background="#0078D7" Foreground="White" BorderThickness="0"
                                                HorizontalAlignment="Right"/>
                                    </StackPanel>
                                </Border>
                                
                                <!-- Repair BITS -->
                                <Border Grid.Column="1" Grid.Row="0" Background="#F5F5F5" BorderBrush="#DDDDDD" BorderThickness="1" Padding="10" CornerRadius="4" Margin="5,0,5,10">
                                    <StackPanel>
                                        <TextBlock Text="Repair BITS Service" FontSize="14" FontWeight="SemiBold" Margin="0,0,0,5"/>
                                        <TextBlock TextWrapping="Wrap" Margin="0,0,0,5" Height="50">
                                            Repairs the Background Intelligent Transfer Service which is essential for downloading updates.
                                        </TextBlock>
                                        <Button x:Name="btnRepairBITS" Content="Repair" Padding="10,5" 
                                                Background="#0078D7" Foreground="White" BorderThickness="0"
                                                HorizontalAlignment="Right"/>
                                    </StackPanel>
                                </Border>
                                
                                <!-- Check WaaSMedic -->
                                <Border Grid.Column="2" Grid.Row="0" Background="#F5F5F5" BorderBrush="#DDDDDD" BorderThickness="1" Padding="10" CornerRadius="4" Margin="5,0,5,10">
                                    <StackPanel>
                                        <TextBlock Text="Check WaaSMedic" FontSize="14" FontWeight="SemiBold" Margin="0,0,0,5"/>
                                        <TextBlock TextWrapping="Wrap" Margin="0,0,0,5" Height="50">
                                            Checks and repairs Windows Update Medic Service which helps maintain update functionality.
                                        </TextBlock>
                                        <Button x:Name="btnWaaSMedic" Content="Check" Padding="10,5" 
                                                Background="#0078D7" Foreground="White" BorderThickness="0"
                                                HorizontalAlignment="Right"/>
                                    </StackPanel>
                                </Border>
                                
                                <!-- Analyze update logs -->
                                <Border Grid.Column="3" Grid.Row="0" Background="#F5F5F5" BorderBrush="#DDDDDD" BorderThickness="1" Padding="10" CornerRadius="4" Margin="5,0,0,10">
                                    <StackPanel>
                                        <TextBlock Text="Analyze Logs" FontSize="14" FontWeight="SemiBold" Margin="0,0,0,5"/>
                                        <TextBlock TextWrapping="Wrap" Margin="0,0,0,5" Height="50">
                                            Analyzes Windows Update logs to identify specific errors and troubleshooting information.
                                        </TextBlock>
                                        <Button x:Name="btnAnalyzeLogs" Content="Analyze" Padding="10,5" 
                                                Background="#0078D7" Foreground="White" BorderThickness="0"
                                                HorizontalAlignment="Right"/>
                                    </StackPanel>
                                </Border>
                                
                                <!-- Repair registry -->
                                <Border Grid.Column="0" Grid.Row="1" Background="#F5F5F5" BorderBrush="#DDDDDD" BorderThickness="1" Padding="10" CornerRadius="4" Margin="0,0,5,10">
                                    <StackPanel>
                                        <TextBlock Text="Repair Registry" FontSize="14" FontWeight="SemiBold" Margin="0,0,0,5"/>
                                        <TextBlock TextWrapping="Wrap" Margin="0,0,0,5" Height="50">
                                            Repairs registry settings related to Windows Update to restore proper functionality.
                                        </TextBlock>
                                        <Button x:Name="btnRepairRegistry" Content="Repair" Padding="10,5" 
                                                Background="#0078D7" Foreground="White" BorderThickness="0"
                                                HorizontalAlignment="Right"/>
                                    </StackPanel>
                                </Border>
                                
                                <!-- Repair database -->
                                <Border Grid.Column="1" Grid.Row="1" Background="#F5F5F5" BorderBrush="#DDDDDD" BorderThickness="1" Padding="10" CornerRadius="4" Margin="5,0,5,10">
                                    <StackPanel>
                                        <TextBlock Text="Repair Database" FontSize="14" FontWeight="SemiBold" Margin="0,0,0,5"/>
                                        <TextBlock TextWrapping="Wrap" Margin="0,0,0,5" Height="50">
                                            Repairs the Windows Update database to fix corruption issues that prevent updates.
                                        </TextBlock>
                                        <Button x:Name="btnRepairDatabase" Content="Repair" Padding="10,5" 
                                                Background="#0078D7" Foreground="White" BorderThickness="0"
                                                HorizontalAlignment="Right"/>
                                    </StackPanel>
                                </Border>
                                
                                <!-- Toggle debug mode -->
                                <Border Grid.Column="2" Grid.Row="1" Background="#F5F5F5" BorderBrush="#DDDDDD" BorderThickness="1" Padding="10" CornerRadius="4" Margin="5,0,5,10">
                                    <StackPanel>
                                        <TextBlock Text="Debug Mode" FontSize="14" FontWeight="SemiBold" Margin="0,0,0,5"/>
                                        <TextBlock TextWrapping="Wrap" Margin="0,0,0,5" Height="50">
                                            Toggles enhanced logging and diagnostic features for advanced troubleshooting.
                                        </TextBlock>
                                        <Button x:Name="btnToggleDebug" Content="Toggle" Padding="10,5" 
                                                Background="#0078D7" Foreground="White" BorderThickness="0"
                                                HorizontalAlignment="Right"/>
                                    </StackPanel>
                                </Border>
                                
                                <!-- Auto-Fix -->
                                <Border Grid.Column="3" Grid.Row="1" Background="#F5F5F5" BorderBrush="#DDDDDD" BorderThickness="1" Padding="10" CornerRadius="4" Margin="5,0,0,10">
                                    <StackPanel>
                                        <TextBlock Text="Auto-Fix" FontSize="14" FontWeight="SemiBold" Margin="0,0,0,5"/>
                                        <TextBlock TextWrapping="Wrap" Margin="0,0,0,5" Height="50">
                                            Runs comprehensive automatic troubleshooting to identify and fix common update issues.
                                        </TextBlock>
                                        <Button x:Name="btnAutoFix" Content="Start" Padding="10,5" 
                                                Background="#107C10" Foreground="White" BorderThickness="0"
                                                HorizontalAlignment="Right"/>
                                    </StackPanel>
                                </Border>
                            </Grid>
                        </StackPanel>
                    </ScrollViewer>
                </Grid>
            </Grid>
        </Grid>
        
        <!-- Footer -->
        <Border Grid.Row="2" Background="#F0F0F0" BorderBrush="#DDDDDD" BorderThickness="0,1,0,0">
            <Grid Margin="20,0">
                <TextBlock x:Name="statusText" Text="Bereit" VerticalAlignment="Center"/>
                <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Center">
                    <TextBlock x:Name="versionText" Text="easyWINUpdate v0.1.4" FontStyle="Italic" Foreground="#555555"/>
                </StackPanel>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

# XAML-Reader erstellen und GUI laden
try {
    $reader = New-Object System.Xml.XmlNodeReader $xaml
    try {
        $window = [Windows.Markup.XamlReader]::Load($reader)
        # Erfolgreiche Meldung
        Write-Host "GUI wurde erfolgreich geladen." -ForegroundColor Green
    } catch {
        Write-Error "Fehler beim Laden des XAML: $($_.Exception.Message)"
        if ($_.Exception.Message -like '*PlaceholderText*') {
            Write-Host "Die Eigenschaft 'PlaceholderText' wird in WPF nicht unterstützt." -ForegroundColor Yellow
            Write-Host "Bitte ersetzen Sie diese durch eine WPF-kompatible Lösung, wie z.B. eine TextBox mit VisualBrush." -ForegroundColor Yellow
        }
        exit
    }
} catch {
    Write-Error "Fehler beim Erstellen des XmlNodeReader: $($_.Exception.Message)"
    exit
}

# Hauptfunktionen des Scripts
#region Hilfsfunktionen

# Funktion zum Aktualisieren des Status-Textes
function Update-StatusText {
    param (
        [string]$Text,
        [string]$Color = "Black"
    )
    
    $statusText = $window.FindName("statusText")
    $statusText.Text = $Text
    $statusText.Foreground = $Color
}

# Funktion zum Überprüfen der Update-Quelle
function Get-UpdateSource {
    $wsusSettings = Get-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate" -ErrorAction SilentlyContinue
    $wsusUrl = $wsusSettings.WUServer
    
    if ($wsusUrl) {
        return "WSUS ($wsusUrl)"
    }
    
    $intuneRegistered = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Enrollments\*" -ErrorAction SilentlyContinue | 
                        Where-Object { $_.ProviderID -eq "MS DM Server" -and $_.EnrollmentState -eq 1 }
    
    if ($intuneRegistered) {
        return "Microsoft Intune"
    }
    
    return "Windows Update (Microsoft)"
}

# Funktion zum Abrufen des letzten Update-Datums
function Get-LastUpdateDate {
    $lastUpdate = Get-HotFix | Sort-Object -Property InstalledOn -Descending | Select-Object -First 1 -ExpandProperty InstalledOn
    
    if ($lastUpdate) {
        return $lastUpdate.ToString("dd.MM.yyyy HH:mm")
    } else {
        return "Keine Updates gefunden"
    }
}
# Funktion zur Diagnose von Windows Update-Fehlercodes
# Funktion zur Diagnose von Windows Update-Fehlercodes
function Diagnose-WindowsUpdateError {
    param (
        [string]$ErrorCode = ""
    )
    
    try {
        Update-StatusText -Text "Analysiere Windows Update-Fehler..." -Color "Blue"
        
        # Wenn kein Fehlercode uebergeben wurde, versuche den letzten Fehlercode aus dem Ereignisprotokoll zu holen
        if ([string]::IsNullOrEmpty($ErrorCode)) {
            $lastError = Get-WinEvent -LogName "Microsoft-Windows-WindowsUpdateClient/Operational" -MaxEvents 50 | 
                Where-Object { $_.LevelDisplayName -eq "Fehler" -or $_.Message -like "*0x8*" } | 
                Select-Object -First 1
                
            if ($lastError) {
                # Extrahiere den Fehlercode aus der Nachricht (Format 0x8xxxxxxx)
                $matches = [regex]::Matches($lastError.Message, "0x8[0-9A-Fa-f]{7}")
                if ($matches.Count -gt 0) {
                    $ErrorCode = $matches[0].Value
                }
            }
        }
        
        # Verwende die globale Fehlercode-Tabelle
        # Wenn ein bekannter Fehlercode gefunden wurde
        if ($ErrorCode -and $errorCodes.ContainsKey($ErrorCode)) {
            $errorInfo = $errorCodes[$ErrorCode]
            
            # Erstelle die Ausgabe
            $output = "Windows Update-Fehlercode: $ErrorCode`r`n"
            $output += "Beschreibung: $($errorInfo.Beschreibung)`r`n"
            $output += "Empfohlene Loesungen:`r`n"
            
            foreach ($loesung in $errorInfo.Loesungen) {
                $output += "- $loesung`r`n"
            }
            
            Update-StatusText -Text "Fehlercode-Analyse abgeschlossen." -Color "Green"
            return $output
        }
        elseif ([string]::IsNullOrEmpty($ErrorCode)) {
            Update-StatusText -Text "Kein Windows Update-Fehlercode gefunden." -Color "Yellow"
            return "Es wurde kein Windows Update-Fehlercode gefunden."
        }
        else {
            Update-StatusText -Text "Unbekannter Fehlercode: $ErrorCode" -Color "Yellow"
            return "Unbekannter Windows Update-Fehlercode: $ErrorCode"
        }
    }
    catch {
        Update-StatusText -Text "Fehler bei der Analyse des Windows Update-Fehlers: $_" -Color "Red"
        return "Fehler bei der Analyse: $_"
    }
}

# Funktion zur erweiterten BITS-Reparatur
function Repair-BITSService {
    try {
        Update-StatusText -Text "Führe erweiterte BITS-Reparatur durch..." -Color "Blue"
        
        # BITS-Dienst stoppen
        Stop-Service -Name BITS -Force
        
        # BITS-Warteschlange löschen
        Get-BitsTransfer -AllUsers | Remove-BitsTransfer
        
        # BITS-Verzeichnisse bereinigen
        $qmgrFiles = @(
            "$env:ALLUSERSPROFILE\Microsoft\Network\Downloader\qmgr*.dat",
            "$env:ALLUSERSPROFILE\Application Data\Microsoft\Network\Downloader\qmgr*.dat"
        )
        
        foreach ($file in $qmgrFiles) {
            if (Test-Path $file) {
                Remove-Item -Path $file -Force
            }
        }
        
        # Registry-Einträge für BITS zurücksetzen
        $bitsSvcKey = "HKLM:\SYSTEM\CurrentControlSet\Services\BITS"
        
        # Dienststarttyp auf Automatisch setzen
        Set-ItemProperty -Path $bitsSvcKey -Name "Start" -Value 2
        
        # Sicherheitsbeschreibungen zurücksetzen
        $securityCommand = "sc.exe sdset bits D:(A;;CCLCSWRPWPDTLOCRRC;;;SY)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWLOCRRC;;;AU)(A;;CCLCSWRPWPDTLOCRRC;;;PU)"
        Invoke-Expression -Command $securityCommand
        
        # BITS-Dienst neu starten
        Start-Service -Name BITS
        
        Update-StatusText -Text "BITS-Dienst wurde erfolgreich repariert." -Color "Green"
        return $true
    } catch {
        Update-StatusText -Text "Fehler bei der BITS-Reparatur: $_" -Color "Red"
        return $false
    }
}

# Funktion zur Prüfung und Reparatur des WaaSMedic-Dienstes
function Repair-WaaSMedicService {
    try {
        Update-StatusText -Text "Prüfe Windows Update Medic Service..." -Color "Blue"
        
        # Prüfen, ob der Dienst existiert (nur auf neueren Windows-Versionen)
        $waasService = Get-Service -Name WaaSMedicSvc -ErrorAction SilentlyContinue
        
        if (-not $waasService) {
            Update-StatusText -Text "Windows Update Medic Service (WaaSMedicSvc) ist auf diesem System nicht verfügbar." -Color "Yellow"
            return $false
        }
        
        # Prüfen und korrigieren des Dienststarttyps
        if ($waasService.StartType -ne "Automatic") {
            Set-Service -Name WaaSMedicSvc -StartupType Automatic
            Update-StatusText -Text "WaaSMedicSvc-Starttyp auf Automatisch gesetzt." -Color "Blue"
        }
        
        # Dienst starten, falls er nicht läuft
        if ($waasService.Status -ne "Running") {
            Start-Service -Name WaaSMedicSvc
            Update-StatusText -Text "WaaSMedicSvc-Dienst gestartet." -Color "Blue"
        }
        
        # Überprüfen der Registry-Einstellungen
        $waasKey = "HKLM:\SYSTEM\CurrentControlSet\Services\WaaSMedicSvc"
        if (Test-Path $waasKey) {
            $startValue = (Get-ItemProperty -Path $waasKey -Name "Start").Start
            if ($startValue -ne 2) {
                Set-ItemProperty -Path $waasKey -Name "Start" -Value 2
                Update-StatusText -Text "WaaSMedicSvc Registry-Einstellungen korrigiert." -Color "Blue"
            }
        }
        
        Update-StatusText -Text "Windows Update Medic Service wurde überprüft und repariert." -Color "Green"
        return $true
    } catch {
        Update-StatusText -Text "Fehler bei der Überprüfung des Windows Update Medic Service: $_" -Color "Red"
        return $false
    }
}
# Funktion zur Analyse der Windows Update-Logs
function Analyze-WindowsUpdateLogs {
    try {
        Update-StatusText -Text "Analysiere Windows Update-Logs..." -Color "Blue"
        
        # Pfad zum Windows Update Log je nach Windows-Version
        $osInfo = Get-CimInstance -ClassName Win32_OperatingSystem
        $osVersion = [Version]$osInfo.Version
        
        if ($osVersion -ge [Version]"10.0") {
            # Windows 10/11: Logs konvertieren
            $logPath = "$env:TEMP\WindowsUpdate.log"
            Get-WindowsUpdateLog -LogPath $logPath | Out-Null
        } else {
            # Ältere Windows-Versionen
            $logPath = "$env:SystemRoot\WindowsUpdate.log"
        }
        
        if (-not (Test-Path $logPath)) {
            Update-StatusText -Text "Windows Update-Log wurde nicht gefunden." -Color "Yellow"
            return "Windows Update-Log konnte nicht gefunden werden."
        }
        
        # Log-Dateigröße prüfen
        $logFile = Get-Item $logPath
        if ($logFile.Length -gt 5MB) {
            Update-StatusText -Text "Windows Update-Log ist sehr groß ($(($logFile.Length / 1MB).ToString('0.00')) MB)." -Color "Yellow"
        }
        
        # Fehler im Log suchen
        $errorKeywords = @(
            "Fehler", "Error", "0x8", "FATAL", "WARNING", "Failed", "Fehlgeschlagen"
        )
        
        $logContent = Get-Content -Path $logPath -Tail 1000 -ErrorAction SilentlyContinue
        $errors = $logContent | Where-Object { 
            $line = $_
            $foundKeyword = $false
            foreach ($keyword in $errorKeywords) {
                if ($line -match $keyword) {
                    $foundKeyword = $true
                    break
                }
            }
            $foundKeyword
        } | Select-Object -Last 20
        
        if ($errors.Count -eq 0) {
            Update-StatusText -Text "Keine offensichtlichen Fehler in den Windows Update-Logs gefunden." -Color "Green"
            return "Keine eindeutigen Fehler in den letzten 1000 Log-Einträgen gefunden."
        } else {
            Update-StatusText -Text "Fehler in Windows Update-Logs gefunden." -Color "Yellow"
            
            $output = "Die letzten 20 Fehlermeldungen aus den Windows Update-Logs:`r`n"
            foreach ($error in $errors) {
                $output += "- $error`r`n"
            }
            
            return $output
        }
    } catch {
        Update-StatusText -Text "Fehler bei der Analyse der Windows Update-Logs: $_" -Color "Red"
        return "Fehler bei der Analyse: $_"
    }
}
# Funktion zur Behebung häufiger Registry-Probleme mit Windows Update
function Repair-WindowsUpdateRegistry {
    try {
        Update-StatusText -Text "Repariere Windows Update-Registry-Einstellungen..." -Color "Blue"
        
        # Wichtige Windows Update-Registry-Pfade
        $registryPaths = @(
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate",
            "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate",
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update"
        )
        
        # Problematische Registry-Werte prüfen und korrigieren
        foreach ($path in $registryPaths) {
            # Prüfen, ob der Pfad existiert
            if (-not (Test-Path $path)) {
                New-Item -Path $path -Force | Out-Null
                Update-StatusText -Text "Registry-Pfad erstellt: $path" -Color "Blue"
            }
            
            # Überprüfung und Korrektur spezifischer Schlüsselwerte
            if ($path -eq "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update") {
                # AUOptions prüfen (falls auf deaktiviert)
                $auOptionsPath = Join-Path $path "AUOptions"
                if (Test-Path $auOptionsPath) {
                    $auOptions = (Get-ItemProperty -Path $path).AUOptions
                    if ($auOptions -eq 1) {
                        Set-ItemProperty -Path $path -Name "AUOptions" -Value 3 # Auf "Herunterladen aber nicht installieren" setzen
                        Update-StatusText -Text "Automatische Updates waren deaktiviert - auf 'Herunterladen aber nicht installieren' gesetzt." -Color "Blue"
                    }
                }
            }
        }
        
        # Windows Update-Dienste auf den richtigen Starttyp setzen
        $updateServices = @(
            @{Name = "wuauserv"; StartupType = "Automatic"; DisplayName = "Windows Update"},
            @{Name = "bits"; StartupType = "Manual"; DisplayName = "Background Intelligent Transfer Service"},
            @{Name = "cryptsvc"; StartupType = "Automatic"; DisplayName = "Cryptographic Services"},
            @{Name = "trustedinstaller"; StartupType = "Manual"; DisplayName = "Windows Modules Installer"}
        )
        
        foreach ($service in $updateServices) {
            $svc = Get-Service -Name $service.Name -ErrorAction SilentlyContinue
            if ($svc) {
                if ($svc.StartType -ne $service.StartupType) {
                    Set-Service -Name $service.Name -StartupType $service.StartupType
                    Update-StatusText -Text "$($service.DisplayName)-Dienst auf $($service.StartupType) gesetzt." -Color "Blue"
                }
            }
        }
        
        # Debugging-Modus für Windows Update ausschalten (falls aktiviert)
        $wuTraceKey = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Trace"
        if (Test-Path $wuTraceKey) {
            $flags = (Get-ItemProperty -Path $wuTraceKey -ErrorAction SilentlyContinue).Flags
            if ($flags -eq 1) {
                Set-ItemProperty -Path $wuTraceKey -Name "Flags" -Value 0
                Update-StatusText -Text "Windows Update-Debugging deaktiviert." -Color "Blue"
            }
        }
        
        Update-StatusText -Text "Windows Update-Registry-Einstellungen wurden repariert." -Color "Green"
        return $true
    } catch {
        Update-StatusText -Text "Fehler bei der Reparatur der Registry-Einstellungen: $_" -Color "Red"
        return $false
    }
}
# Funktion zur Reparatur der Windows Update-Datenbank
function Repair-WindowsUpdateDatabase {
    try {
        Update-StatusText -Text "Repariere Windows Update-Datenbank..." -Color "Blue"
        
        # Windows Update-Dienste stoppen
        $services = @('wuauserv', 'cryptSvc', 'bits', 'msiserver', 'appidsvc')
        foreach ($service in $services) {
            Stop-Service -Name $service -Force -ErrorAction SilentlyContinue
        }
        
        # SoftwareDistribution-Ordner umbenennen (nicht löschen)
        $sdFolder = "$env:SystemRoot\SoftwareDistribution"
        $sdBackup = "$env:SystemRoot\SoftwareDistribution.bak"
        $catroot2Folder = "$env:SystemRoot\System32\catroot2"
        $catroot2Backup = "$env:SystemRoot\System32\catroot2.bak"
        
        if (Test-Path $sdFolder) {
            # Altes Backup entfernen, falls vorhanden
            if (Test-Path $sdBackup) {
                Remove-Item -Path $sdBackup -Recurse -Force -ErrorAction SilentlyContinue
            }
            
            # Umbenennen statt löschen
            Rename-Item -Path $sdFolder -NewName "SoftwareDistribution.bak" -Force -ErrorAction SilentlyContinue
            New-Item -Path $sdFolder -ItemType Directory -Force | Out-Null
        }
        
        if (Test-Path $catroot2Folder) {
            # Altes Backup entfernen, falls vorhanden
            if (Test-Path $catroot2Backup) {
                Remove-Item -Path $catroot2Backup -Recurse -Force -ErrorAction SilentlyContinue
            }
            
            # Umbenennen statt löschen
            Rename-Item -Path $catroot2Folder -NewName "catroot2.bak" -Force -ErrorAction SilentlyContinue
            # catroot2 wird automatisch neu erstellt, daher kein New-Item notwendig
        }
        
        # Datenbankdateien zurücksetzen
        $dbFiles = @(
            "$env:SystemRoot\WindowsUpdate.log",
            "$env:ALLUSERSPROFILE\Application Data\Microsoft\Network\Downloader\qmgr*.dat",
            "$env:ALLUSERSPROFILE\Microsoft\Network\Downloader\qmgr*.dat"
        )
        
        foreach ($file in $dbFiles) {
            if (Test-Path $file) {
                Remove-Item -Path $file -Force -ErrorAction SilentlyContinue
            }
        }
        
        # WSUS Client-Einstellungen zurücksetzen (wenn vorhanden)
        $wsusClientKey = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\WSUS\Client"
        if (Test-Path $wsusClientKey) {
            Remove-Item -Path $wsusClientKey -Recurse -Force -ErrorAction SilentlyContinue
        }
        
        # Dienste wieder starten
        foreach ($service in $services) {
            Start-Service -Name $service -ErrorAction SilentlyContinue
        }
        
        # Update-Erkennung erzwingen
        wuauclt.exe /resetauthorization /detectnow
        
        Update-StatusText -Text "Windows Update-Datenbank wurde repariert. System sollte neu gestartet werden." -Color "Green"
        return $true
    } catch {
        Update-StatusText -Text "Fehler bei der Reparatur der Windows Update-Datenbank: $_" -Color "Red"
        return $false
    }
}
# Funktion zum Aktivieren/Deaktivieren des Fehlersuchmodus für Windows Update
function Toggle-WindowsUpdateDebugMode {
    param (
        [bool]$Enable = $true
    )
    
    try {
        if ($Enable) {
            Update-StatusText -Text "Aktiviere Windows Update-Fehlersuchmodus..." -Color "Blue"
            
            # Registry-Pfad für Windows Update-Trace erstellen oder aktualisieren
            $wuTraceKey = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Trace"
            if (-not (Test-Path $wuTraceKey)) {
                New-Item -Path $wuTraceKey -Force | Out-Null
            }
            
            # Trace-Flag setzen
            New-ItemProperty -Path $wuTraceKey -Name "Flags" -Value 1 -PropertyType DWord -Force | Out-Null
            
            # Detaillierte Protokollierung aktivieren
            $regPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Setup"
            if (-not (Test-Path $regPath)) {
                New-Item -Path $regPath -Force | Out-Null
            }
            New-ItemProperty -Path $regPath -Name "EnableVerboseLogging" -Value 1 -PropertyType DWord -Force | Out-Null
            
            Update-StatusText -Text "Windows Update-Fehlersuchmodus wurde aktiviert. Versuchen Sie jetzt die Update-Operation erneut und prüfen Sie die Logs." -Color "Green"
            return $true
        } else {
            Update-StatusText -Text "Deaktiviere Windows Update-Fehlersuchmodus..." -Color "Blue"
            
            # Registry-Pfad für Windows Update-Trace aktualisieren
            $wuTraceKey = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Trace"
            if (Test-Path $wuTraceKey) {
                Set-ItemProperty -Path $wuTraceKey -Name "Flags" -Value 0 -Force | Out-Null
            }
            
            # Detaillierte Protokollierung deaktivieren
            $regPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Setup"
            if (Test-Path $regPath) {
                Set-ItemProperty -Path $regPath -Name "EnableVerboseLogging" -Value 0 -Force | Out-Null
            }
            
            Update-StatusText -Text "Windows Update-Fehlersuchmodus wurde deaktiviert." -Color "Green"
            return $true
        }
    } catch {
        Update-StatusText -Text "Fehler beim Aendern des Windows Update-Fehlersuchmodus: $_" -Color "Red"
        return $false
    }
}


# Funktion zum Abrufen der WSUS-Einstellungen
function Get-WSUSSettings {
    $wsusSettings = @{
        Server = $null
        Status = "Nicht konfiguriert"
        TargetGroup = "Keine"
        ConfigSource = "Nicht konfiguriert"
        LastCheck = "Unbekannt"
    }
    
    $wuPolicies = Get-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate" -ErrorAction SilentlyContinue
    $wuSettings = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate" -ErrorAction SilentlyContinue
    
    if ($wuPolicies) {
        $wsusSettings.Server = $wuPolicies.WUServer
        $wsusSettings.Status = if ($wsusSettings.Server) { "Aktiv" } else { "Nicht konfiguriert" }
        $wsusSettings.TargetGroup = if ($wuPolicies.TargetGroupEnabled -eq 1) { $wuPolicies.TargetGroup } else { "Keine" }
        
        if ($wsusSettings.Server) {
            if (Test-Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate") {
                $wsusSettings.ConfigSource = "Gruppenrichtlinie"
            } elseif (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\WSUS") {
                $wsusSettings.ConfigSource = "Manuell konfiguriert"
            } else {
                $wsusSettings.ConfigSource = "Unbekannt"
            }
        }
    }
    
    if ($wuSettings -and $wuSettings.LastWUAutoupdate) {
        try {
            $lastCheckTime = [DateTime]::FromFileTime($wuSettings.LastWUAutoupdate)
            $wsusSettings.LastCheck = $lastCheckTime.ToString("dd.MM.yyyy HH:mm")
        } catch {
            $wsusSettings.LastCheck = "Fehler beim Konvertieren des Datums"
        }
    }
    
    return $wsusSettings
}
# Funktion zur automatischen Analyse und Behebung von Windows Update-Problemen
function Auto-FixWindowsUpdateIssues {
    try {
        Update-StatusText -Text "Starte automatische Behebung von Windows Update-Problemen..." -Color "Blue"
        
        # Schritt 1: Windows Update-Dienst prüfen
        $wuauserv = Get-Service -Name wuauserv
        if ($wuauserv.Status -ne "Running" -or $wuauserv.StartType -ne "Automatic") {
            Update-StatusText -Text "Windows Update-Dienst ist nicht korrekt konfiguriert. Wird korrigiert..." -Color "Yellow"
            Set-Service -Name wuauserv -StartupType Automatic
            Start-Service -Name wuauserv
        }
        
        # Schritt 2: Vorhandene Fehler analysieren
        $lastError = Get-WinEvent -LogName "Microsoft-Windows-WindowsUpdateClient/Operational" -MaxEvents 50 | 
            Where-Object { $_.LevelDisplayName -eq "Fehler" -or $_.Message -like "*0x8*" } | 
            Select-Object -First 1
        
        $errorCode = ""
        if ($lastError) {
            $matches = [regex]::Matches($lastError.Message, "0x8[0-9A-Fa-f]{7}")
            if ($matches.Count -gt 0) {
                $errorCode = $matches[0].Value
                Update-StatusText -Text "Letzter Windows Update-Fehlercode: $errorCode" -Color "Yellow"
            }
        }
        
        # Schritt 3: Spezifische Fehlerbehebung basierend auf dem Fehlercode
        $specificFix = $false
        
        switch ($errorCode) {
            "0x80240034" { # Download-Fehler
                Update-StatusText -Text "Download-Fehler erkannt. Behebe Netzwerkprobleme..." -Color "Yellow"
                Clear-WindowsUpdateHistory
                Repair-BITSService
                $specificFix = $true
            }
            "0x8024001F" { # Keine Verbindung
                Update-StatusText -Text "Verbindungsproblem erkannt. Prüfe Netzwerkeinstellungen..." -Color "Yellow"
                Repair-BITSService
                Reset-WindowsUpdateComponents
                $specificFix = $true
            }
            "0x80070020" { # Datei gesperrt
                Update-StatusText -Text "Dateisperrproblem erkannt. Bereinige Windows Update-Datenbank..." -Color "Yellow"
                Repair-WindowsUpdateDatabase
                $specificFix = $true
            }
            "0x80240022" { # Alle Updates fehlgeschlagen
                Update-StatusText -Text "Kritisches Update-Problem erkannt. Führe vollständige Reparatur durch..." -Color "Yellow"
                Reset-WindowsUpdateComponents
                Repair-WindowsSystemFiles
                $specificFix = $true
            }
            "0x80070422" { # Dienst deaktiviert
                Update-StatusText -Text "Windows Update-Dienst deaktiviert. Aktiviere Dienste..." -Color "Yellow"
                Repair-WindowsUpdateRegistry
                $specificFix = $true
            }
            "0x8024402C" { # DNS-Auflösungsproblem
                Update-StatusText -Text "DNS-Problem erkannt. Setze Netzwerkeinstellungen zurück..." -Color "Yellow"
                # Netzwerkeinstellungen zurücksetzen
                ipconfig /flushdns
                ipconfig /registerdns
                $specificFix = $true
            }
        }
        
        # Schritt 4: Allgemeine Reparatur, wenn keine spezifische Behebung angewendet wurde
        if (-not $specificFix) {
            Update-StatusText -Text "Führe allgemeine Reparaturmaßnahmen durch..." -Color "Blue"
            
            # Windows Update-Komponenten zurücksetzen
            Reset-WindowsUpdateComponents
            
            # WaaSMedic-Dienst reparieren (auf neueren Windows-Versionen)
            $osInfo = Get-CimInstance -ClassName Win32_OperatingSystem
            $osVersion = [Version]$osInfo.Version
            if ($osVersion -ge [Version]"10.0.17134") {
                Repair-WaaSMedicService
            }
            
            # Registry-Einstellungen reparieren
            Repair-WindowsUpdateRegistry
        }
        
        # Schritt 5: Update-Dienste neu starten
        Update-StatusText -Text "Starte Windows Update-Dienste neu..." -Color "Blue"
        Restart-Service -Name wuauserv -Force
        Restart-Service -Name bits -Force
        
        # Schritt 6: Update-Erkennung erzwingen
        Update-StatusText -Text "Erzwinge Update-Erkennung..." -Color "Blue"
        wuauclt.exe /resetauthorization /detectnow
        
        Update-StatusText -Text "Automatische Fehlerbehebung abgeschlossen. Es wird empfohlen, den Computer neu zu starten." -Color "Green"
        return $true
    } catch {
        Update-StatusText -Text "Fehler bei der automatischen Fehlerbehebung: $_" -Color "Red"
        return $false
    }
}
# Funktion zum Zurücksetzen der WSUS-Einstellungen
function Reset-WSUSSettings {
    try {
        # WSUS-Einstellungen aus der Registry entfernen
        $regKeys = @(
            "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate",
            "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"
        )
        
        foreach ($key in $regKeys) {
            if (Test-Path $key) {
                Remove-Item -Path $key -Force -Recurse -ErrorAction SilentlyContinue
            }
        }
        
        # Windows Update Dienste neustarten
        Restart-Service -Name wuauserv -Force
        Restart-Service -Name bits -Force
        
        # WindowsUpdate-Dienst-Client ID zurücksetzen
        Stop-Service -Name wuauserv
        Remove-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate" -Name "SusClientId" -ErrorAction SilentlyContinue
        Remove-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate" -Name "SusClientIDValidation" -ErrorAction SilentlyContinue
        Start-Service -Name wuauserv
        
        # Client-Aktualisierung erzwingen
        wuauclt.exe /resetauthorization /detectnow
        
        return $true
    } catch {
        Write-Error "Fehler beim Zurücksetzen der WSUS-Einstellungen: $_"
        return $false
    }
}

# Funktion zum Abrufen installierter Updates
function Get-InstalledWindowsUpdates {
    try {
        $hotfixes = Get-HotFix | Select-Object @{Name="HotFixID"; Expression={$_.HotFixID}}, 
                                                @{Name="Description"; Expression={$_.Description}}, 
                                                @{Name="InstalledOn"; Expression={if($_.InstalledOn){$_.InstalledOn.ToString("dd.MM.yyyy")}else{"Unbekannt"}}}
        
        return $hotfixes
    } catch {
        Write-Error "Fehler beim Abrufen der installierten Updates: $_"
        return $null
    }
}

# Funktion zum Abrufen verfügbarer Updates
function Get-AvailableWindowsUpdates {
    param (
        [switch]$IncludeCritical = $true,
        [switch]$IncludeSecurity = $true,
        [switch]$IncludeDefinition = $true,
        [switch]$IncludeFeature = $false
    )
    
    try {
        Write-Host "Suche nach Windows Updates..." -ForegroundColor Cyan
        
        # Kategorien basierend auf Parametern bestimmen
        $categoryFilter = @()
        if ($IncludeCritical) { $categoryFilter += "Critical Updates" }
        if ($IncludeSecurity) { $categoryFilter += "Security Updates" }
        if ($IncludeDefinition) { $categoryFilter += "Definition Updates" }
        if ($IncludeFeature) { $categoryFilter += "Feature Packs" }
        
        Write-Host "Suche in Kategorien: $($categoryFilter -join ", ")" -ForegroundColor Cyan
        
        # Windows Update COM-Objekt erstellen
        $updateSession = New-Object -ComObject Microsoft.Update.Session
        $updateSearcher = $updateSession.CreateUpdateSearcher()
        
        # Suchkriterien definieren
        $searchCriteria = "IsInstalled=0"
        
        # Nach Updates suchen
        $searchResult = $updateSearcher.Search($searchCriteria)
        
        # Updates filtern und formatieren
        $updates = @()
        foreach ($update in $searchResult.Updates) {
            # Kategorie bestimmen
            $category = "Andere"
            if ($update.MsrcSeverity -eq "Critical") { $category = "Critical Updates" }
            elseif ($update.Type -like "*Security*") { $category = "Security Updates" }
            elseif ($update.Title -like "*Definition*") { $category = "Definition Updates" }
            elseif ($update.Title -like "*Feature*") { $category = "Feature Packs" }
            
            # Prüfen, ob Update der ausgewählten Kategorie entspricht
            if ($categoryFilter.Count -eq 0 -or $categoryFilter -contains $category) {
                # KB-Nummer extrahieren
                $kbMatch = [regex]::Match($update.Title, 'KB\d+')
                $kbNumber = if ($kbMatch.Success) { $kbMatch.Value } else { "Unbekannt" }
                
                $updates += [PSCustomObject]@{
                    IsSelected = $false
                    KBArticleID = $kbNumber
                    Title = $update.Title
                    Category = $category
                    Size = if ($update.MaxDownloadSize -gt 0) { "$([Math]::Round($update.MaxDownloadSize / 1MB, 2)) MB" } else { "Unbekannt" }
                    Identity = $update.Identity
                }
            }
        }
        
        if ($updates.Count -gt 0) {
            Write-Host "$($updates.Count) Updates gefunden." -ForegroundColor Green
        } else {
            Write-Host "Keine Updates gefunden." -ForegroundColor Yellow
        }
        
        return $updates
    } catch {
        $errorMessage = $_.Exception.Message
        Write-Error "Fehler beim Abrufen verfügbarer Updates: $errorMessage"
        return $null
    }
}

# Funktion zum Entfernen eines Windows Updates
function Remove-WindowsUpdateKB {
    param (
        [string]$KBNumber
    )
    
    try {
        if (-not $KBNumber.StartsWith("KB")) {
            $KBNumber = "KB$KBNumber"
        }
        
        $kbID = $KBNumber.Replace('KB', '')
        $result = Start-Process -FilePath "wusa.exe" -ArgumentList "/uninstall", "/kb:$kbID", "/quiet", "/norestart" -Wait -PassThru
        
        # Warten auf Abschluss der Deinstallation
        Start-Sleep -Seconds 5
        
        return $true
    } catch {
        $errorMessage = $_.Exception.Message
        Write-Error "Fehler beim Entfernen des Updates $KBNumber $($errorMessage)"
        return $false
    }
}

# Funktion zum Abrufen des Dienststatus
function Get-ServiceStatusInfo {
    param (
        [string]$ServiceName
    )
    
    try {
        $service = Get-Service -Name $ServiceName -ErrorAction Stop
        
        return @{
            Status = $service.Status
            DisplayName = $service.DisplayName
            StatusText = if ($service.Status -eq "Running") { "Aktiv" } else { "Gestoppt" }
            StatusColor = if ($service.Status -eq "Running") { "Green" } else { "Red" }
        }
    } catch {
        return @{
            Status = "Unknown"
            DisplayName = "Dienst nicht gefunden"
            StatusText = "Fehler"
            StatusColor = "Red"
        }
    }
}

# Funktion zum Testen der WSUS-Verbindung
function Test-WSUSConnection {
    param (
        [string]$WSUSServer
    )
    
    try {
        if (-not $WSUSServer -or $WSUSServer -eq "Nicht konfiguriert") {
            return $false
        }
        
        # Server-URL extrahieren
        $serverUrl = $WSUSServer
        if ($serverUrl -match "http[s]?://([^/:]+)(?::[0-9]+)?/?") {
            $serverName = $matches[1]
        } else {
            $serverName = $serverUrl
        }
        
        # Testen, ob der Server erreichbar ist
        $ping = Test-Connection -ComputerName $serverName -Count 2 -Quiet
        
        if (-not $ping) {
            return $false
        }
        
        # Versuchen, die WSUS-Konfiguration zu aktualisieren
        try {
            # PowerShell-Befehl zum Aktualisieren der WSUS-Konfiguration
            wuauclt.exe /detectnow
            return $true
        } catch {
            return $false
        }
    } catch {
        return $false
    }
}

# Funktion zum Abrufen der Windows-Version
function Get-WindowsVersionInfo {
    try {
        $osInfo = Get-CimInstance -ClassName Win32_OperatingSystem
        $versionCaption = $osInfo.Caption
        $versionBuild = $osInfo.BuildNumber
        $versionInfo = "$versionCaption (Build $versionBuild)"
        
        return $versionInfo
    } catch {
        Write-Error "Fehler beim Abrufen der Windows-Version: $_"
        return "Unbekannt"
    }
}

# Funktion zum Erstellen eines Systemwiederherstellungspunkts
function New-SystemRestorePoint {
    try {
        # Systemwiederherstellungspunkt erstellen
        $description = "Vor Windows Update Troubleshooting"
        Enable-ComputerRestore -Drive "$env:SystemDrive" -ErrorAction SilentlyContinue
        Checkpoint-Computer -Description $description -RestorePointType "MODIFY_SETTINGS" -ErrorAction Stop
        return $true
    } catch {
        Write-Error "Fehler beim Erstellen des Systemwiederherstellungspunkts: $_"
        return $false
    }
}

# Funktion zum Zurücksetzen der Windows Update-Komponenten
function Reset-WindowsUpdateComponents {
    try {
        Update-StatusText -Text "Setze Windows Update-Komponenten zurück..." -Color "Blue"
        
        # Dienste stoppen
        $services = @('wuauserv', 'cryptSvc', 'bits', 'msiserver', 'appidsvc', 'trustedinstaller')
        foreach ($service in $services) {
            Stop-Service -Name $service -Force -ErrorAction SilentlyContinue
        }
        
        # Windows Update-Datenbank und Cache-Ordner löschen
        $folders = @("$env:SystemRoot\SoftwareDistribution", "$env:SystemRoot\System32\catroot2")
        foreach ($folder in $folders) {
            if (Test-Path $folder) {
                Remove-Item -Path $folder -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
        
        # QMGR-Dateien entfernen
        Remove-Item -Path "$env:ALLUSERSPROFILE\Application Data\Microsoft\Network\Downloader\qmgr*.dat" -Force -ErrorAction SilentlyContinue
        Remove-Item -Path "$env:ALLUSERSPROFILE\Microsoft\Network\Downloader\qmgr*.dat" -Force -ErrorAction SilentlyContinue
        
        # Sicherstellen, dass Windows Update-Ordner wieder erstellt werden
        New-Item -Path "$env:SystemRoot\SoftwareDistribution" -ItemType Directory -Force | Out-Null
        
        # Alle BITS-Jobs löschen
        Get-BitsTransfer -AllUsers | Remove-BitsTransfer -ErrorAction SilentlyContinue
        
        # Dienste wieder starten
        foreach ($service in $services) {
            Start-Service -Name $service -ErrorAction SilentlyContinue
        }
        
        # Update-Erkennung erzwingen
        wuauclt.exe /resetauthorization /detectnow
        
        Update-StatusText -Text "Windows Update-Komponenten wurden zurückgesetzt. System sollte neu gestartet werden." -Color "Green"
        return $true
    } catch {
        Update-StatusText -Text "Fehler beim Zurücksetzen der Windows Update-Komponenten: $_" -Color "Red"
        return $false
    }
}

# Funktion zum Überprüfen und Reparieren von Windows-Systemdateien
function Repair-WindowsSystemFiles {
    try {
        Update-StatusText -Text "Überprüfe Windows-Systemdateien..." -Color "Blue"
        
        # DISM zum Reparieren des Windows-Images verwenden
        Start-Process -FilePath "DISM.exe" -ArgumentList "/Online", "/Cleanup-Image", "/RestoreHealth" -Wait -NoNewWindow
        
        # SFC zum Überprüfen und Reparieren von Systemdateien verwenden
        Start-Process -FilePath "sfc.exe" -ArgumentList "/scannow" -Wait -NoNewWindow
        
        Update-StatusText -Text "Windows-Systemdateien wurden überprüft und repariert." -Color "Green"
        return $true
    } catch {
        Update-StatusText -Text "Fehler bei der Überprüfung der Windows-Systemdateien: $_" -Color "Red"
        return $false
    }
}

# Funktion zum Löschen des Windows Update-Verlaufs
function Clear-WindowsUpdateHistory {
    try {
        Update-StatusText -Text "Lösche Windows Update-Verlauf..." -Color "Blue"
        
        # Dienste stoppen
        Stop-Service -Name wuauserv -Force
        
        # SoftwareDistribution-Ordner löschen
        if (Test-Path "$env:SystemRoot\SoftwareDistribution") {
            Remove-Item -Path "$env:SystemRoot\SoftwareDistribution" -Recurse -Force
        }
        
        # Windows Update-Log löschen
        if (Test-Path "$env:SystemRoot\WindowsUpdate.log") {
            Remove-Item -Path "$env:SystemRoot\WindowsUpdate.log" -Force
        }
        
        # SoftwareDistribution-Ordner neu erstellen
        New-Item -Path "$env:SystemRoot\SoftwareDistribution" -ItemType Directory -Force | Out-Null
        
        # Dienst wieder starten
        Start-Service -Name wuauserv
        
        Update-StatusText -Text "Windows Update-Verlauf wurde gelöscht." -Color "Green"
        return $true
    } catch {
        Update-StatusText -Text "Fehler beim Löschen des Windows Update-Verlaufs: $_" -Color "Red"
        return $false
    }
}

# Funktion zum Neuregistrieren der Windows Update-DLLs
function Register-WindowsUpdateDLLs {
    try {
        Update-StatusText -Text "Registriere Windows Update-DLLs neu..." -Color "Blue"
        
        # Stoppen der Windows Update-Dienste
        Stop-Service -Name wuauserv -Force
        
        # Wichtige Windows Update-DLLs registrieren
        $dllFiles = @(
            "atl.dll", "urlmon.dll", "mshtml.dll", "shdocvw.dll", "browseui.dll", "jscript.dll", 
            "vbscript.dll", "scrrun.dll", "msxml.dll", "msxml3.dll", "msxml6.dll", "actxprxy.dll", 
            "softpub.dll", "wintrust.dll", "dssenh.dll", "rsaenh.dll", "gpkcsp.dll", "sccbase.dll", 
            "slbcsp.dll", "cryptdlg.dll", "oleaut32.dll", "ole32.dll", "shell32.dll", "wuapi.dll", 
            "wuaueng.dll", "wuaueng1.dll", "wucltui.dll", "wups.dll", "wups2.dll", "wuweb.dll", 
            "qmgr.dll", "qmgrprxy.dll", "wucltux.dll", "muweb.dll", "wuwebv.dll"
        )
        
        foreach ($dll in $dllFiles) {
            $dllPath = Join-Path $env:SystemRoot "System32\$dll"
            if (Test-Path $dllPath) {
                Start-Process -FilePath "regsvr32.exe" -ArgumentList "/s", $dllPath -Wait
            }
        }
        
        # Windows Update-Dienst neu starten
        Start-Service -Name wuauserv
        
        Update-StatusText -Text "Windows Update-DLLs wurden neu registriert." -Color "Green"
        return $true
    } catch {
        Update-StatusText -Text "Fehler beim Neuregistrieren der Windows Update-DLLs: $_" -Color "Red"
        return $false
    }
}

# Funktion zum Entfernen von hängenden Updates und BITS-Jobs
function Clear-StuckUpdatesAndBITSJobs {
    try {
        Update-StatusText -Text "Entferne hängende Updates und BITS-Jobs..." -Color "Blue"
        
        # Alle BITS-Jobs löschen
        Get-BitsTransfer -AllUsers | Remove-BitsTransfer
        
        # Gestagte Windows Update-Pakete entfernen
        $stagedUpdates = Get-WindowsPackage -Online | Where-Object { $_.PackageState -eq 'Staged' }
        if ($stagedUpdates) {
            foreach ($update in $stagedUpdates) {
                Remove-WindowsPackage -PackageName $update.PackageName -Online -NoRestart
            }
        }
        
        # Windows Update und BITS-Dienste neu starten
        Restart-Service -Name wuauserv -Force
        Restart-Service -Name bits -Force
        
        # Update-Erkennung erzwingen
        wuauclt.exe /detectnow
        
        Update-StatusText -Text "Hängende Updates und BITS-Jobs wurden entfernt." -Color "Green"
        return $true
    } catch {
        Update-StatusText -Text "Fehler beim Entfernen von hängenden Updates: $_" -Color "Red"
        return $false
    }
}
#endregion

#region Event-Handler und GUI-Logik

# Hauptfunktion zum Laden der Update-Status-Seite
function Load-UpdateStatusPage {
    # Windows-Version anzeigen
    $txtWindowsVersion = $window.FindName("txtWindowsVersion")
    $txtWindowsVersion.Text = Get-WindowsVersionInfo
    
    # Update-Quelle anzeigen
    $txtUpdateSource = $window.FindName("txtUpdateSource")
    $txtUpdateSource.Text = Get-UpdateSource
    
    # Letztes Update-Datum anzeigen
    $txtLastUpdate = $window.FindName("txtLastUpdate")
    $txtLastUpdate.Text = Get-LastUpdateDate
    
    # Update-Status anzeigen
    $txtUpdateStatus = $window.FindName("txtUpdateStatus")
    $wuaStatus = Get-ServiceStatusInfo -ServiceName "wuauserv"
    $txtUpdateStatus.Text = "Windows Update Dienst ist $($wuaStatus.StatusText)"
    
    # Anzahl installierter Updates anzeigen
    $txtInstalledUpdatesCount = $window.FindName("txtInstalledUpdatesCount")
    $hotfixes = Get-InstalledWindowsUpdates
    $txtInstalledUpdatesCount.Text = if ($hotfixes) { "$($hotfixes.Count) Updates installiert" } else { "Keine Updates gefunden" }
    
    # Dienststatus anzeigen
    $txtUpdateService = $window.FindName("txtUpdateService")
    $txtUpdateService.Text = $wuaStatus.StatusText
    $txtUpdateService.Foreground = $wuaStatus.StatusColor
    
    $txtBITSService = $window.FindName("txtBITSService")
    $bitsStatus = Get-ServiceStatusInfo -ServiceName "bits"
    $txtBITSService.Text = $bitsStatus.StatusText
    $txtBITSService.Foreground = $bitsStatus.StatusColor
    
    Update-StatusText -Text "Update-Status wurde aktualisiert." -Color "Green"
}

# Funktion zum Laden der installierten Updates-Seite
function Load-InstalledUpdatesPage {
    $dgInstalledUpdates = $window.FindName("dgInstalledUpdates")
    $hotfixes = Get-InstalledWindowsUpdates
    
    if ($hotfixes) {
        $dgInstalledUpdates.ItemsSource = $hotfixes
        Update-StatusText -Text "$($hotfixes.Count) installierte Updates geladen." -Color "Green"
    } else {
        Update-StatusText -Text "Keine installierten Updates gefunden." -Color "Red"
    }
}

# Funktion zum Laden der verfügbaren Updates-Seite
function Load-AvailableUpdatesPage {
    $chkCriticalUpdates = $window.FindName("chkCriticalUpdates")
    $chkSecurityUpdates = $window.FindName("chkSecurityUpdates")
    $chkDefinitionUpdates = $window.FindName("chkDefinitionUpdates")
    $chkFeatureUpdates = $window.FindName("chkFeatureUpdates")
    
    $progressUpdates = $window.FindName("progressUpdates")
    $progressUpdates.Visibility = "Visible"
    Update-StatusText -Text "Suche nach verfügbaren Updates..." -Color "Blue"
    
    # Updates mit ausgewählten Kategorien suchen
    # Vereinfachter Aufruf ohne Parameter
    $updates = Get-AvailableWindowsUpdates
    
    $dgAvailableUpdates = $window.FindName("dgAvailableUpdates")
    
    if ($updates) {
        $dgAvailableUpdates.ItemsSource = $updates
        Update-StatusText -Text "$($updates.Count) verfügbare Updates gefunden." -Color "Green"
    } else {
        $dgAvailableUpdates.ItemsSource = @()
        Update-StatusText -Text "Keine verfügbaren Updates gefunden." -Color "Blue"
    }
    
    $progressUpdates.Visibility = "Collapsed"
}

# Funktion zum Laden der WSUS-Einstellungen-Seite
function Load-WSUSSettingsPage {
    $wsusSettings = Get-WSUSSettings
    
    $txtWSUSServer = $window.FindName("txtWSUSServer")
    $txtWSUSServer.Text = if ($wsusSettings.Server) { $wsusSettings.Server } else { "Nicht konfiguriert" }
    
    $txtWSUSStatus = $window.FindName("txtWSUSStatus")
    $txtWSUSStatus.Text = $wsusSettings.Status
    
    $txtWSUSTargetGroup = $window.FindName("txtWSUSTargetGroup")
    $txtWSUSTargetGroup.Text = $wsusSettings.TargetGroup
    
    $txtWSUSConfigSource = $window.FindName("txtWSUSConfigSource")
    $txtWSUSConfigSource.Text = $wsusSettings.ConfigSource
    
    $txtWSUSLastCheck = $window.FindName("txtWSUSLastCheck")
    $txtWSUSLastCheck.Text = $wsusSettings.LastCheck
    
    Update-StatusText -Text "WSUS-Einstellungen wurden geladen." -Color "Green"
}

# Funktion zum Umschalten zwischen den Seiten
function Switch-Page {
    param (
        [string]$PageName
    )
    
    $updateStatusPage = $window.FindName("updateStatusPage")
    $installedUpdatesPage = $window.FindName("installedUpdatesPage")
    $availableUpdatesPage = $window.FindName("availableUpdatesPage")
    $wsusSettingsPage = $window.FindName("wsusSettingsPage")
    $troubleshootingPage = $window.FindName("troubleshootingPage")
    
    $updateStatusPage.Visibility = "Collapsed"
    $installedUpdatesPage.Visibility = "Collapsed"
    $availableUpdatesPage.Visibility = "Collapsed"
    $wsusSettingsPage.Visibility = "Collapsed"
    $troubleshootingPage.Visibility = "Collapsed"
    
    switch ($PageName) {
        "UpdateStatus" {
            $updateStatusPage.Visibility = "Visible"
            Load-UpdateStatusPage
        }
        "InstalledUpdates" {
            $installedUpdatesPage.Visibility = "Visible"
            Load-InstalledUpdatesPage
        }
        "AvailableUpdates" {
            $availableUpdatesPage.Visibility = "Visible"
            Load-AvailableUpdatesPage
        }
        "WSUSSettings" {
            $wsusSettingsPage.Visibility = "Visible"
            Load-WSUSSettingsPage
        }
        "Troubleshooting" {
            $troubleshootingPage.Visibility = "Visible"
            Update-StatusText -Text "Troubleshooting-Bereich bereit." -Color "Blue"
        }
    }
}

# Event-Handler für Seitennavigation
$navUpdateStatus = $window.FindName("navUpdateStatus")
$navUpdateStatus.Add_Checked({
    Switch-Page -PageName "UpdateStatus"
})

$navInstalledUpdates = $window.FindName("navInstalledUpdates")
$navInstalledUpdates.Add_Checked({
    Switch-Page -PageName "InstalledUpdates"
})

$navAvailableUpdates = $window.FindName("navAvailableUpdates")
$navAvailableUpdates.Add_Checked({
    Switch-Page -PageName "AvailableUpdates"
})

$navWSUSSettings = $window.FindName("navWSUSSettings")
$navWSUSSettings.Add_Checked({
    Switch-Page -PageName "WSUSSettings"
})

$navTroubleshooting = $window.FindName("navTroubleshooting")
$navTroubleshooting.Add_Checked({
    Switch-Page -PageName "Troubleshooting"
})

# Event-Handler für Update-Status-Seite
$btnCheckForUpdates = $window.FindName("btnCheckForUpdates")
$btnCheckForUpdates.Add_Click({
    Update-StatusText -Text "Suche nach Updates..." -Color "Blue"
    wuauclt.exe /detectnow
    Start-Sleep -Seconds 2
    Load-UpdateStatusPage
})

$btnRestartService = $window.FindName("btnRestartService")
$btnRestartService.Add_Click({
    Update-StatusText -Text "Windows Update Dienst wird neu gestartet..." -Color "Blue"
    Restart-Service -Name wuauserv -Force
    Start-Sleep -Seconds 2
    Load-UpdateStatusPage
})

$btnRestartBITS = $window.FindName("btnRestartBITS")
$btnRestartBITS.Add_Click({
    Update-StatusText -Text "BITS Dienst wird neu gestartet..." -Color "Blue"
    Restart-Service -Name bits -Force
    Start-Sleep -Seconds 2
    Load-UpdateStatusPage
})

# Event-Handler für installierte Updates-Seite
$btnSearchInstalled = $window.FindName("btnSearchInstalled")
$btnSearchInstalled.Add_Click({
    $txtSearchInstalled = $window.FindName("txtSearchInstalled")
    $searchText = $txtSearchInstalled.Text.ToLower()
    
    if (-not [string]::IsNullOrEmpty($searchText)) {
        $hotfixes = Get-InstalledWindowsUpdates
        $filteredHotfixes = $hotfixes | Where-Object { 
            $_.HotFixID.ToLower().Contains($searchText) -or 
            $_.Description.ToLower().Contains($searchText) 
        }
        
        $dgInstalledUpdates = $window.FindName("dgInstalledUpdates")
        $dgInstalledUpdates.ItemsSource = $filteredHotfixes
        
        Update-StatusText -Text "$($filteredHotfixes.Count) Updates gefunden, die '$searchText' enthalten." -Color "Green"
    } else {
        Load-InstalledUpdatesPage
    }
})

$btnRefreshInstalled = $window.FindName("btnRefreshInstalled")
$btnRefreshInstalled.Add_Click({
    Load-InstalledUpdatesPage
})

$btnExportInstalled = $window.FindName("btnExportInstalled")
$btnExportInstalled.Add_Click({
    $savePath = "$env:USERPROFILE\Desktop\Installierte_Updates_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $hotfixes = Get-InstalledWindowsUpdates
    
    if ($hotfixes) {
        $hotfixes | Export-Csv -Path $savePath -NoTypeInformation -Delimiter ";"
        Update-StatusText -Text "Updates wurden nach '$savePath' exportiert." -Color "Green"
    } else {
        Update-StatusText -Text "Keine Updates zum Exportieren vorhanden." -Color "Red"
    }
})

# Event-Handler für verfügbare Updates-Seite
$btnSearchAvailable = $window.FindName("btnSearchAvailable")
$btnSearchAvailable.Add_Click({
    Load-AvailableUpdatesPage
})

$btnInstallAll = $window.FindName("btnInstallAll")
$btnInstallAll.Add_Click({
    $dgAvailableUpdates = $window.FindName("dgAvailableUpdates")
    $updates = $dgAvailableUpdates.ItemsSource
    
    if ($updates -and $updates.Count -gt 0) {
        $result = [System.Windows.MessageBox]::Show(
            "Möchten Sie alle $($updates.Count) Updates installieren? Der Computer wird möglicherweise neu gestartet.",
            "Updates installieren",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Question
        )
        
        if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
            Update-StatusText -Text "Installiere alle Updates..." -Color "Blue"
            
            try {
                $progressUpdates = $window.FindName("progressUpdates")
                $progressUpdates.Visibility = "Visible"
                
                # Alle verfügbaren Updates installieren
                Get-WindowsUpdate -WindowsUpdate -Install -AcceptAll -IgnoreReboot | Out-Null
                
                $progressUpdates.Visibility = "Collapsed"
                Update-StatusText -Text "Alle Updates wurden installiert." -Color "Green"
                
                # Nach der Installation neu laden
                Load-AvailableUpdatesPage
            } catch {
                Update-StatusText -Text "Fehler bei der Installation: $_" -Color "Red"
                $progressUpdates.Visibility = "Collapsed"
            }
        }
    } else {
        Update-StatusText -Text "Keine Updates zum Installieren verfügbar." -Color "Red"
    }
})

$btnInstallSelected = $window.FindName("btnInstallSelected")
$btnInstallSelected.Add_Click({
    $dgAvailableUpdates = $window.FindName("dgAvailableUpdates")
    $updates = $dgAvailableUpdates.ItemsSource | Where-Object { $_.IsSelected }
    
    if ($updates -and $updates.Count -gt 0) {
        $result = [System.Windows.MessageBox]::Show(
            "Möchten Sie die ausgewählten $($updates.Count) Updates installieren? Der Computer wird möglicherweise neu gestartet.",
            "Updates installieren",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Question
        )
        
        if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
            Update-StatusText -Text "Installiere ausgewählte Updates..." -Color "Blue"
            
            try {
                $progressUpdates = $window.FindName("progressUpdates")
                $progressUpdates.Visibility = "Visible"
                
                # Ausgewählte Updates installieren
                foreach ($update in $updates) {
                    Get-WindowsUpdate -WindowsUpdate -KBArticleID $update.KBArticleID.Replace("KB", "") -Install -AcceptAll -IgnoreReboot | Out-Null
                }
                
                $progressUpdates.Visibility = "Collapsed"
                Update-StatusText -Text "Ausgewählte Updates wurden installiert." -Color "Green"
                
                # Nach der Installation neu laden
                Load-AvailableUpdatesPage
            } catch {
                Update-StatusText -Text "Fehler bei der Installation: $_" -Color "Red"
                $progressUpdates.Visibility = "Collapsed"
            }
        }
    } else {
        Update-StatusText -Text "Keine Updates ausgewählt." -Color "Red"
    }
})

$btnDownloadSelected = $window.FindName("btnDownloadSelected")
$btnDownloadSelected.Add_Click({
    $dgAvailableUpdates = $window.FindName("dgAvailableUpdates")
    $updates = $dgAvailableUpdates.ItemsSource | Where-Object { $_.IsSelected }
    
    if ($updates -and $updates.Count -gt 0) {
        Update-StatusText -Text "Lade ausgewählte Updates herunter..." -Color "Blue"
        
        try {
            $progressUpdates = $window.FindName("progressUpdates")
            $progressUpdates.Visibility = "Visible"
            
            # Ausgewählte Updates herunterladen
            foreach ($update in $updates) {
                Get-WindowsUpdate -WindowsUpdate -KBArticleID $update.KBArticleID.Replace("KB", "") -Download -AcceptAll | Out-Null
            }
            
            $progressUpdates.Visibility = "Collapsed"
            Update-StatusText -Text "Ausgewählte Updates wurden heruntergeladen." -Color "Green"
        } catch {
            Update-StatusText -Text "Fehler beim Herunterladen: $_" -Color "Red"
            $progressUpdates.Visibility = "Collapsed"
        }
    } else {
        Update-StatusText -Text "Keine Updates ausgewählt." -Color "Red"
    }
})

# Event-Handler für WSUS-Einstellungen-Seite
$btnResetWSUS = $window.FindName("btnResetWSUS")
$btnResetWSUS.Add_Click({
    $result = [System.Windows.MessageBox]::Show(
        "Möchten Sie die WSUS-Einstellungen wirklich zurücksetzen? Der Computer wird danach Updates direkt von Microsoft beziehen.",
        "WSUS-Einstellungen zurücksetzen",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Warning
    )
    
    if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
        Update-StatusText -Text "WSUS-Einstellungen werden zurückgesetzt..." -Color "Blue"
        
        if (Reset-WSUSSettings) {
            Update-StatusText -Text "WSUS-Einstellungen wurden erfolgreich zurückgesetzt." -Color "Green"
            Load-WSUSSettingsPage
        } else {
            Update-StatusText -Text "Fehler beim Zurücksetzen der WSUS-Einstellungen." -Color "Red"
        }
    }
})

$btnCheckWSUSConn = $window.FindName("btnCheckWSUSConn")
$btnCheckWSUSConn.Add_Click({
    $wsusSettings = Get-WSUSSettings
    
    if ($wsusSettings.Server) {
        Update-StatusText -Text "Überprüfe Verbindung zu WSUS-Server..." -Color "Blue"
        
        if (Test-WSUSConnection -WSUSServer $wsusSettings.Server) {
            Update-StatusText -Text "Verbindung zum WSUS-Server erfolgreich hergestellt." -Color "Green"
        } else {
            Update-StatusText -Text "Verbindung zum WSUS-Server fehlgeschlagen." -Color "Red"
        }
    } else {
        Update-StatusText -Text "Kein WSUS-Server konfiguriert." -Color "Red"
    }
})

$btnDetectNow = $window.FindName("btnDetectNow")
$btnDetectNow.Add_Click({
    Update-StatusText -Text "Starte Updateerkennung..." -Color "Blue"
    
    # Update-Erkennung erzwingen
    wuauclt.exe /detectnow
    
    Update-StatusText -Text "Updateerkennung wurde gestartet." -Color "Green"
})

# Event-Handler für die Entfernen-Buttons in den DataGrids
# Da diese dynamisch generiert werden, müssen wir die Click-Events für die gesamte DataGrid abfangen
$dgInstalledUpdates = $window.FindName("dgInstalledUpdates")
$dgInstalledUpdates.Add_LoadingRow({
    param($sender, $e)
    
    $row = $e.Row
    $btnRemoveUpdate = $row.FindName("btnRemoveUpdate")
    
    if ($btnRemoveUpdate) {
        $btnRemoveUpdate.Add_Click({
            $kbNumber = $btnRemoveUpdate.Tag
            
            $result = [System.Windows.MessageBox]::Show(
                "Möchten Sie das Update $kbNumber wirklich entfernen?",
                "Update entfernen",
                [System.Windows.MessageBoxButton]::YesNo,
                [System.Windows.MessageBoxImage]::Warning
            )
            
            if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
                Update-StatusText -Text "Entferne Update $kbNumber..." -Color "Blue"
                
                if (Remove-WindowsUpdateKB -KBNumber $kbNumber) {
                    Update-StatusText -Text "Update $kbNumber wurde entfernt. Computer muss möglicherweise neu gestartet werden." -Color "Green"
                    # Liste aktualisieren
                    Load-InstalledUpdatesPage
                } else {
                    Update-StatusText -Text "Fehler beim Entfernen des Updates $kbNumber." -Color "Red"
                }
            }
        })
    }
})

# Event-Handler für Troubleshooting-Seite
$btnCreateRestorePoint = $window.FindName("btnCreateRestorePoint")
$btnCreateRestorePoint.Add_Click({
    Update-StatusText -Text "Erstelle Systemwiederherstellungspunkt..." -Color "Blue"
    
    if (New-SystemRestorePoint) {
        Update-StatusText -Text "Systemwiederherstellungspunkt wurde erfolgreich erstellt." -Color "Green"
    } else {
        Update-StatusText -Text "Fehler beim Erstellen des Systemwiederherstellungspunkts." -Color "Red"
    }
})

$btnResetComponents = $window.FindName("btnResetComponents")
$btnResetComponents.Add_Click({
    $result = [System.Windows.MessageBox]::Show(
        "Möchten Sie die Windows Update-Komponenten wirklich zurücksetzen? Dies stoppt die Update-Dienste und löscht alle Update-Caches.",
        "Windows Update-Komponenten zurücksetzen",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Warning
    )
    
    if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
        Reset-WindowsUpdateComponents
    }
})

$btnCheckSystemFiles = $window.FindName("btnCheckSystemFiles")
$btnCheckSystemFiles.Add_Click({
    $result = [System.Windows.MessageBox]::Show(
        "Möchten Sie die Windows-Systemdateien überprüfen und reparieren? Dieser Vorgang kann einige Zeit dauern.",
        "Systemdateien überprüfen",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Question
    )
    
    if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
        Repair-WindowsSystemFiles
    }
})

$btnClearUpdateHistory = $window.FindName("btnClearUpdateHistory")
$btnClearUpdateHistory.Add_Click({
    $result = [System.Windows.MessageBox]::Show(
        "Möchten Sie den Windows Update-Verlauf und alle Caches löschen? Dies entfernt alle Informationen über frühere Update-Versuche.",
        "Update-Verlauf löschen",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Warning
    )
    
    if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
        Clear-WindowsUpdateHistory
    }
})

$btnRegisterDLLs = $window.FindName("btnRegisterDLLs")
$btnRegisterDLLs.Add_Click({
    $result = [System.Windows.MessageBox]::Show(
        "Möchten Sie alle Windows Update-DLLs neu registrieren? Dies kann bei Problemen mit Update-Komponenten helfen.",
        "Update-DLLs neu registrieren",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Question
    )
    
    if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
        Register-WindowsUpdateDLLs
    }
})

$btnClearStuckUpdates = $window.FindName("btnClearStuckUpdates")
$btnClearStuckUpdates.Add_Click({
    $result = [System.Windows.MessageBox]::Show(
        "Möchten Sie hängende Updates und BITS-Jobs entfernen? Dies kann bei Problemen mit hängenden Downloads helfen.",
        "Hängende Updates entfernen",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Question
    )
    
    if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
        Clear-StuckUpdatesAndBITSJobs
    }
})
# Event-Handler für die neuen Troubleshooting-Funktionen
$btnDiagnoseError = $window.FindName("btnDiagnoseError")
$btnDiagnoseError.Add_Click({
    $result = Diagnose-WindowsUpdateError
    [System.Windows.MessageBox]::Show($result, "Windows Update-Fehlerdiagnose", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
})

$btnRepairBITS = $window.FindName("btnRepairBITS")
$btnRepairBITS.Add_Click({
    $result = [System.Windows.MessageBox]::Show(
        "Möchten Sie eine erweiterte Reparatur des BITS-Dienstes durchführen?",
        "BITS-Dienst reparieren",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Question
    )
    
    if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
        Repair-BITSService
    }
})
# WaaSMedic-Dienst
$btnWaaSMedic = $window.FindName("btnWaaSMedic")
$btnWaaSMedic.Add_Click({
    $result = [System.Windows.MessageBox]::Show(
        "Möchten Sie den Windows Update Medic Service prüfen und reparieren? Dies kann bei Problemen mit dem Update-Prozess helfen.",
        "WaaSMedic-Dienst reparieren",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Question
    )
    
    if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
        Repair-WaaSMedicService
    }
})

# Windows Update-Log-Analyse
$btnAnalyzeLogs = $window.FindName("btnAnalyzeLogs")
$btnAnalyzeLogs.Add_Click({
    $result = Analyze-WindowsUpdateLogs
    [System.Windows.MessageBox]::Show($result, "Windows Update-Log-Analyse", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
})

# Registry-Reparatur
$btnRepairRegistry = $window.FindName("btnRepairRegistry")
$btnRepairRegistry.Add_Click({
    $result = [System.Windows.MessageBox]::Show(
        "Möchten Sie die Windows Update-Registry-Einstellungen reparieren? Dies kann bei Problemen mit Konfigurationseinstellungen helfen.",
        "Registry-Einstellungen reparieren",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Question
    )
    
    if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
        Repair-WindowsUpdateRegistry
    }
})

# Datenbank-Reparatur
$btnRepairDatabase = $window.FindName("btnRepairDatabase")
$btnRepairDatabase.Add_Click({
    $result = [System.Windows.MessageBox]::Show(
        "Möchten Sie die Windows Update-Datenbank reparieren? Dies erstellt Sicherungen der aktuellen Update-Datenbank und setzt sie zurück.",
        "Update-Datenbank reparieren",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Question
    )
    
    if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
        Repair-WindowsUpdateDatabase
    }
})

# Debug-Modus umschalten
$btnToggleDebug = $window.FindName("btnToggleDebug")
$btnToggleDebug.Add_Click({
    # Prüfen, ob der Debug-Modus aktiv ist
    $wuTraceKey = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Trace"
    $debugActive = $false
    
    if (Test-Path $wuTraceKey) {
        $flags = (Get-ItemProperty -Path $wuTraceKey -ErrorAction SilentlyContinue).Flags
        if ($flags -eq 1) {
            $debugActive = $true
        }
    }
    
    $message = if ($debugActive) {
        "Der Windows Update-Fehlersuchmodus ist derzeit aktiv. Möchten Sie ihn deaktivieren?"
    } else {
        "Möchten Sie den Windows Update-Fehlersuchmodus aktivieren? Dies hilft bei der Diagnose von schwerwiegenden Update-Problemen."
    }
    
    $result = [System.Windows.MessageBox]::Show(
        $message,
        "Fehlersuchmodus umschalten",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Question
    )
    
    if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
        Toggle-WindowsUpdateDebugMode -Enable (-not $debugActive)
    }
})

# Auto-Fix
$btnAutoFix = $window.FindName("btnAutoFix")
$btnAutoFix.Add_Click({
    $result = [System.Windows.MessageBox]::Show(
        "Möchten Sie eine automatische Fehlerbehebung für Windows Update-Probleme durchführen? Das System versucht, automatisch Fehler zu erkennen und zu beheben.",
        "Automatische Fehlerbehebung",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Question
    )
    
    if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
        Auto-FixWindowsUpdateIssues
    }
})

$dgAvailableUpdates = $window.FindName("dgAvailableUpdates")
$dgAvailableUpdates.Add_LoadingRow({
    param($sender, $e)
    
    $row = $e.Row
    $btnInstallUpdate = $row.FindName("btnInstallUpdate")
    
    if ($btnInstallUpdate) {
        $btnInstallUpdate.Add_Click({
            $identity = $btnInstallUpdate.Tag
            $kbNumber = ($dgAvailableUpdates.ItemsSource | Where-Object { $_.Identity -eq $identity }).KBArticleID
            
            $result = [System.Windows.MessageBox]::Show(
                "Möchten Sie das Update $kbNumber installieren?",
                "Update installieren",
                [System.Windows.MessageBoxButton]::YesNo,
                [System.Windows.MessageBoxImage]::Question
            )
            
            if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
                Update-StatusText -Text "Installiere Update $kbNumber..." -Color "Blue"
                
                try {
                    $progressUpdates = $window.FindName("progressUpdates")
                    $progressUpdates.Visibility = "Visible"
                    
                    # Update installieren
                    Get-WindowsUpdate -WindowsUpdate -KBArticleID $kbNumber.Replace("KB", "") -Install -AcceptAll -IgnoreReboot | Out-Null
                    
                    $progressUpdates.Visibility = "Collapsed"
                    Update-StatusText -Text "Update $kbNumber wurde installiert." -Color "Green"
                    
                    # Nach der Installation neu laden
                    Load-AvailableUpdatesPage
                } catch {
                    Update-StatusText -Text "Fehler bei der Installation: $_" -Color "Red"
                    $progressUpdates.Visibility = "Collapsed"
                }
            }
        })
    }
})
#endregion

#region Hauptausführung

# Computer-Namen anzeigen
$computerName = $window.FindName("computerName")
$computerName.Text = $env:COMPUTERNAME

# Initialisierung
try {
    # Event-Handler für Schließen des Fensters
    $window.Add_Closing({
        # Aufräumarbeiten hier
    })
    
    # Erstmalige Anzeige der Update-Status-Seite
    Load-UpdateStatusPage
    
    # GUI anzeigen
    $window.ShowDialog() | Out-Null
} catch {
    Write-Error "Fehler bei der Ausfuehrung des Scripts: $_"
}
#endregion
# SIG # Begin signature block
# MIIbywYJKoZIhvcNAQcCoIIbvDCCG7gCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC708hWVDi/4HMD
# CTTaHt9/3wk7NxlIgvEPFls/gzfZXqCCFhcwggMQMIIB+KADAgECAhB3jzsyX9Cg
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
# DQEJBDEiBCDUJL1zUZQPFcpt27kXi/l6OY814gLV3dtPua00fJu6pTANBgkqhkiG
# 9w0BAQEFAASCAQBNcc/GMzqWMwEZOxhREUZ4Qs3skju9/YCYFpO/fhQJ6FUc6r3f
# AkBcBAV8JnQaT8pnOgvnIFYRja+7ux/S6BumjhGFIw/7R4HUikyvksdZ/9ITRzQc
# OkpHVn5/EYRXKrv7fHxHuDsYfu+VscTOHTDlwhB3V7i4sHpyadhxzWEabz3Hy3ED
# /jX7FFGWhuVxV2UF27nk8AwStU4DihEkvUxgkGrKxy+k+3fEA2gt8VUlmRg7UUh1
# SVa8zbJqPgqR4wsqxt+5zI8QaL3k5ng1uJ3ilU0zkLJAyHWMpostbK9xBXJi+JbD
# 2G8giE1kNH0nF8ei9xK3++/9W7YJUzC3KsdfoYIDIDCCAxwGCSqGSIb3DQEJBjGC
# Aw0wggMJAgEBMHcwYzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJ
# bmMuMTswOQYDVQQDEzJEaWdpQ2VydCBUcnVzdGVkIEc0IFJTQTQwOTYgU0hBMjU2
# IFRpbWVTdGFtcGluZyBDQQIQC65mvFq6f5WHxvnpBOMzBDANBglghkgBZQMEAgEF
# AKBpMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTI1
# MDcwNTEwMTIxMFowLwYJKoZIhvcNAQkEMSIEIDbRo4ux4Q/OpmNYG4xLdU9tmYUT
# hTBl/LNlE/Pew5R/MA0GCSqGSIb3DQEBAQUABIICAINo3OyF/4bywhdhQ97Fkkft
# cy2AuowPve0ySi9zuBkXHEAcOnNdARPAxth9naZEe2Tb4ZWOHxErOukxpgWmeSc9
# XJtM1fMP4Wyf1HXMJPQH92aVaQ4XUu6PIx4D03kv/Cz+z40sP7xN2/G4iONwUOdF
# NeLVV7SmV2lNEwp1Te0J9+BInHJmmIRWavzBg+mqFfKablAILV8vs8Se/W1VVI+Y
# kOkAf2YjNQ96UEpdV97EogJKxLa5kAzqvvf+oL9CgnNh+Ngdik3cK5L8EKuPYyB2
# ctH0lfuwjC4WZGg+vOHCutKdY5mP7qvLOtzNL52fH3CeThWvYx0n+wT2q0Y+aoEc
# 5tHAZ/M0xQkDy0fag+GfCvW3Q/T7jEcKkdVBQi9y5fPtBVGmKrzpTy8ajfRaiUDR
# zOhJs+m7I9HwiFb0mD8hrP2reUNQ0diXdUrjtWlTVsqrL28eHgtcUCZrlnOPJWP8
# n0elQKuiCIEMkgn1o8bXjIOMMAf/6JS8If3Ax5LZccci6Sat4nuw2CpP2xThkkW9
# J8ietTHDqRn4h4B0BzgDeNbu4vT9oJWdja0Z5K0xIG3hSPSOLjXNIYFnse51WM44
# 4Jw8N+BAclZp8q9fDsIGq3M9gL79tVzOSdm/+3oZ4ccZlesOKGWRhDJHU1+RKX4r
# gt5AMaQFxjsXXnsh4pxN
# SIG # End signature block
