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
                    <TextBlock Text="v0.0.1" Foreground="#CCFFFFFF" FontSize="14" Margin="10,0,0,0" VerticalAlignment="Center"/>
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
                    <StackPanel>
                        <TextBlock Text="Windows Update Troubleshooting" FontSize="22" FontWeight="SemiBold" Margin="0,0,0,15"/>
                        
                        <Border Background="#FFF4F4" BorderBrush="#FFCECE" BorderThickness="1" Padding="15" CornerRadius="4" Margin="0,0,0,15">
                            <StackPanel>
                                <TextBlock TextWrapping="Wrap" Margin="0,0,0,10">
                                    <Bold>Attention:</Bold> The following actions can reset or repair Windows Update components. It is recommended to create a system restore point before performing these actions.
                                </TextBlock>
                                <Button x:Name="btnCreateRestorePoint" Content="Create System Restore Point" Padding="15,8" 
                                        Background="#0078D7" Foreground="White" BorderThickness="0"
                                        HorizontalAlignment="Left" Margin="0,5,0,0"/>
                            </StackPanel>
                        </Border>
                        
                        <!-- Two-column layout for troubleshooting options -->
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>
                            
                            <!-- Left Column, First Row -->
                            <Border Grid.Column="0" Grid.Row="0" Background="#F5F5F5" BorderBrush="#DDDDDD" BorderThickness="1" Padding="15" CornerRadius="4" Margin="0,0,7,15">
                                <StackPanel>
                                    <TextBlock Text="Reset Windows Update Components" FontSize="16" FontWeight="SemiBold" Margin="0,0,0,10"/>
                                    <TextBlock TextWrapping="Wrap" Margin="0,0,0,10" Height="60">
                                        Resets all Windows Update components to their default settings. Stops services, clears caches, and restarts services.
                                    </TextBlock>
                                    <Button x:Name="btnResetComponents" Content="Reset Components" Padding="15,8" 
                                            Background="#E81123" Foreground="White" BorderThickness="0"
                                            HorizontalAlignment="Left" Margin="0,5,0,0"/>
                                </StackPanel>
                            </Border>
                            
                            <!-- Right Column, First Row -->
                            <Border Grid.Column="1" Grid.Row="0" Background="#F5F5F5" BorderBrush="#DDDDDD" BorderThickness="1" Padding="15" CornerRadius="4" Margin="7,0,0,15">
                                <StackPanel>
                                    <TextBlock Text="Repair Windows System Files" FontSize="16" FontWeight="SemiBold" Margin="0,0,0,10"/>
                                    <TextBlock TextWrapping="Wrap" Margin="0,0,0,10" Height="60">
                                        Runs SFC and DISM to repair corrupted Windows files. This process may take some time.
                                    </TextBlock>
                                    <Button x:Name="btnCheckSystemFiles" Content="Check System Files" Padding="15,8" 
                                            Background="#0078D7" Foreground="White" BorderThickness="0"
                                            HorizontalAlignment="Left" Margin="0,5,0,0"/>
                                </StackPanel>
                            </Border>
                            
                            <!-- Left Column, Second Row -->
                            <Border Grid.Column="0" Grid.Row="1" Background="#F5F5F5" BorderBrush="#DDDDDD" BorderThickness="1" Padding="15" CornerRadius="4" Margin="0,0,7,15">
                                <StackPanel>
                                    <TextBlock Text="Clear Update History and Caches" FontSize="16" FontWeight="SemiBold" Margin="0,0,0,10"/>
                                    <TextBlock TextWrapping="Wrap" Margin="0,0,0,10" Height="60">
                                        Clears the Windows Update history and all caches. Windows Update will create a new history during the next scan.
                                    </TextBlock>
                                    <Button x:Name="btnClearUpdateHistory" Content="Clear Update History" Padding="15,8" 
                                            Background="#0078D7" Foreground="White" BorderThickness="0"
                                            HorizontalAlignment="Left" Margin="0,5,0,0"/>
                                </StackPanel>
                            </Border>
                            
                            <!-- Right Column, Second Row -->
                            <Border Grid.Column="1" Grid.Row="1" Background="#F5F5F5" BorderBrush="#DDDDDD" BorderThickness="1" Padding="15" CornerRadius="4" Margin="7,0,0,15">
                                <StackPanel>
                                    <TextBlock Text="Re-register Update DLLs" FontSize="16" FontWeight="SemiBold" Margin="0,0,0,10"/>
                                    <TextBlock TextWrapping="Wrap" Margin="0,0,0,10" Height="60">
                                        Re-registers all important Windows Update DLLs. Helps with issues related to update components.
                                    </TextBlock>
                                    <Button x:Name="btnRegisterDLLs" Content="Re-register DLLs" Padding="15,8" 
                                            Background="#0078D7" Foreground="White" BorderThickness="0"
                                            HorizontalAlignment="Left" Margin="0,5,0,0"/>
                                </StackPanel>
                            </Border>
                            
                            <!-- Left Column, Third Row -->
                            <Border Grid.Column="0" Grid.Row="2" Background="#F5F5F5" BorderBrush="#DDDDDD" BorderThickness="1" Padding="15" CornerRadius="4" Margin="0,0,7,0">
                                <StackPanel>
                                    <TextBlock Text="Remove Stuck Updates and BITS Jobs" FontSize="16" FontWeight="SemiBold" Margin="0,0,0,10"/>
                                    <TextBlock TextWrapping="Wrap" Margin="0,0,0,10" Height="60">
                                        Removes all stuck BITS jobs and updates that might be blocking the update process.
                                    </TextBlock>
                                    <Button x:Name="btnClearStuckUpdates" Content="Remove Stuck Updates" Padding="15,8" 
                                            Background="#0078D7" Foreground="White" BorderThickness="0"
                                            HorizontalAlignment="Left" Margin="0,5,0,0"/>
                                </StackPanel>
                            </Border>
                        </Grid>
                    </StackPanel>
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
        if ($serverUrl -match "http[s]?://([^/:]+)(?::[0-9]+)?") {
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
    Write-Error "Fehler bei der Ausführung des Scripts: $_"
}
#endregion