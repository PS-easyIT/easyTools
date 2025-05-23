# EasyADFinder - Advanced Active Directory Query Tool
# Version: 0.0.4 - Optimiert für Performance und Stabilität

# Modul-Imports
try {
    Import-Module ActiveDirectory -ErrorAction Stop
} catch {
    Write-Host "Das Active Directory PowerShell-Modul konnte nicht geladen werden.`nBitte installieren Sie die RSAT-Tools (Remoteserver-Verwaltungstools)." -ForegroundColor Red
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show("Das ActiveDirectory-Modul ist nicht verfügbar.`nBitte installieren Sie die RSAT-Tools.", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    exit
}

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.DirectoryServices

# XAML für die WPF GUI
[xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="AD-Finder" Height="750" Width="1120"
    WindowStartupLocation="CenterScreen" Background="#f7f7f7">
    
    <Window.Resources>
        <!-- Windows 11 optimierte Stile -->
        <Style x:Key="Win11Button" TargetType="Button">
            <Setter Property="Background" Value="#0078d4"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="Padding" Value="16,8"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" 
                                CornerRadius="5" 
                                Padding="{TemplateBinding Padding}"
                                Effect="{Binding RelativeSource={RelativeSource TemplatedParent}, Path=(TextElement.Effect)}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#106ebe"/>
                    <Setter Property="Effect">
                        <Setter.Value>
                            <DropShadowEffect BlurRadius="5" ShadowDepth="1" Direction="270" Opacity="0.2" Color="#000000"/>
                        </Setter.Value>
                    </Setter>
                </Trigger>
                <Trigger Property="IsPressed" Value="True">
                    <Setter Property="Background" Value="#005a9e"/>
                </Trigger>
                <Trigger Property="IsEnabled" Value="False">
                    <Setter Property="Background" Value="#cccccc"/>
                    <Setter Property="Foreground" Value="#666666"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        
        <Style x:Key="IconButton" TargetType="Button">
            <Setter Property="Padding" Value="8"/>
            <Setter Property="Width" Value="36"/>
            <Setter Property="Height" Value="36"/>
            <Setter Property="Margin" Value="6,0"/>
            <Setter Property="Background" Value="#ffffff"/>
            <Setter Property="Foreground" Value="#0078d4"/>
            <Setter Property="BorderBrush" Value="#e6e6e6"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" 
                                CornerRadius="5"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                BorderBrush="{TemplateBinding BorderBrush}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#f5f5f5"/>
                    <Setter Property="BorderBrush" Value="#d1d1d1"/>
                </Trigger>
                <Trigger Property="IsPressed" Value="True">
                    <Setter Property="Background" Value="#ebebeb"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        
        <Style x:Key="Win11ComboBox" TargetType="ComboBox">
            <Setter Property="Padding" Value="10,6"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="BorderBrush" Value="#d1d1d1"/>
            <Setter Property="Background" Value="White"/>
            <Setter Property="FontSize" Value="13"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="BorderBrush" Value="#0078d4"/>
                </Trigger>
                <Trigger Property="IsDropDownOpen" Value="True">
                    <Setter Property="BorderBrush" Value="#0078d4"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        
        <Style x:Key="Win11TextBox" TargetType="TextBox">
            <Setter Property="Padding" Value="10,6"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="BorderBrush" Value="#d1d1d1"/>
            <Setter Property="Background" Value="White"/>
            <Setter Property="FontSize" Value="13"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="BorderBrush" Value="#0078d4"/>
                </Trigger>
                <Trigger Property="IsKeyboardFocused" Value="True">
                    <Setter Property="BorderBrush" Value="#0078d4"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        
        <Style x:Key="Win11CheckBox" TargetType="CheckBox">
            <Setter Property="Margin" Value="0,6"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
        </Style>

        <Style x:Key="Win11TabItem" TargetType="TabItem">
            <Setter Property="Padding" Value="14,8"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="HeaderTemplate">
                <Setter.Value>
                    <DataTemplate>
                        <TextBlock Text="{Binding}" FontWeight="SemiBold"/>
                    </DataTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="BorderBrush" Value="#0078d4"/>
                    <Setter Property="BorderThickness" Value="0,0,0,2"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        
        <Style x:Key="Win11DataGrid" TargetType="DataGrid">
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="BorderBrush" Value="#e0e0e0"/>
            <Setter Property="GridLinesVisibility" Value="Horizontal"/>
            <Setter Property="HorizontalGridLinesBrush" Value="#f0f0f0"/>
            <Setter Property="RowHeaderWidth" Value="0"/>
            <Setter Property="AlternatingRowBackground" Value="#f9f9f9"/>
            <Setter Property="Background" Value="White"/>
            <Setter Property="FontSize" Value="13"/>
        </Style>
        
        <Style x:Key="Win11DataGridColumnHeader" TargetType="DataGridColumnHeader">
            <Setter Property="Background" Value="White"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Padding" Value="10,8"/>
            <Setter Property="BorderThickness" Value="0,0,1,1"/>
            <Setter Property="BorderBrush" Value="#e0e0e0"/>
        </Style>
        
        <Style x:Key="Win11GroupBox" TargetType="GroupBox">
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="BorderBrush" Value="#e0e0e0"/>
            <Setter Property="Padding" Value="12"/>
        </Style>
        
        <Style x:Key="Win11Card" TargetType="Border">
            <Setter Property="Background" Value="White"/>
            <Setter Property="CornerRadius" Value="8"/>
            <Setter Property="Padding" Value="16"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="BorderBrush" Value="#e0e0e0"/>
            <Setter Property="Effect">
                <Setter.Value>
                    <DropShadowEffect BlurRadius="10" ShadowDepth="1" Direction="270" Opacity="0.1" Color="#000000"/>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>
    
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <!-- Header mit Filtern -->
        <Border Grid.Row="0" Style="{StaticResource Win11Card}" Margin="0,0,0,16">
            <Grid>
                <TabControl BorderThickness="0" Background="Transparent">
                    <TabItem Header="Hauptfilter" Style="{StaticResource Win11TabItem}">
                        <Grid Margin="0,12,0,0">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            
                            <StackPanel Grid.Column="0">
                                <TextBlock Text="AD-Finder" FontSize="22" FontWeight="SemiBold" Margin="0,0,0,16"/>
                                
                                <Grid>
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="Auto"/>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                        <ColumnDefinition Width="*"/>
                                    </Grid.ColumnDefinitions>
                                    <Grid.RowDefinitions>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                    </Grid.RowDefinitions>
                                    
                                    <!-- Zeile 1 -->
                                    <TextBlock Grid.Row="0" Grid.Column="0" Text="Objekttyp:" VerticalAlignment="Center" Margin="0,0,10,12"/>
                                    <ComboBox x:Name="cmbObjectType" Grid.Row="0" Grid.Column="1" Style="{StaticResource Win11ComboBox}" Margin="0,0,20,12">
                                        <ComboBoxItem Content="Benutzer" IsSelected="True"/>
                                        <ComboBoxItem Content="Computer"/>
                                        <ComboBoxItem Content="Gruppen"/>
                                        <ComboBoxItem Content="OUs"/>
                                    </ComboBox>
                                    
                                    <TextBlock Grid.Row="0" Grid.Column="2" Text="Betriebssystem:" VerticalAlignment="Center" Margin="0,0,10,12"/>
                                    <ComboBox x:Name="cmbOS" Grid.Row="0" Grid.Column="3" Style="{StaticResource Win11ComboBox}" Margin="0,0,0,12" IsEditable="True" />
                                    
                                    <!-- Zeile 2 -->
                                    <TextBlock Grid.Row="1" Grid.Column="0" Text="Name enthält:" VerticalAlignment="Center" Margin="0,0,10,12"/>
                                    <TextBox x:Name="txtName" Grid.Row="1" Grid.Column="1" Style="{StaticResource Win11TextBox}" Margin="0,0,20,12"/>
                                    
                                    <TextBlock Grid.Row="1" Grid.Column="2" Text="Letztes Logon:" VerticalAlignment="Center" Margin="0,0,10,12"/>
                                    <Grid Grid.Row="1" Grid.Column="3" Margin="0,0,0,12">
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="Auto"/>
                                        </Grid.ColumnDefinitions>
                                        <ComboBox x:Name="cmbLastLogon" Grid.Column="0" Style="{StaticResource Win11ComboBox}" IsEditable="True">
                                            <ComboBoxItem Content="Alle"/>
                                            <ComboBoxItem Content="30"/>
                                            <ComboBoxItem Content="60"/>
                                            <ComboBoxItem Content="90"/>
                                            <ComboBoxItem Content="180"/>
                                            <ComboBoxItem Content="365"/>
                                        </ComboBox>
                                        <TextBlock Grid.Column="1" Text=" Tage" VerticalAlignment="Center" Margin="6,0,0,0"/>
                                    </Grid>
                                    
                                    <!-- Zeile 3 -->
                                    <TextBlock Grid.Row="2" Grid.Column="0" Text="Beschreibung:" VerticalAlignment="Center" Margin="0,0,10,12"/>
                                    <TextBox x:Name="txtDescription" Grid.Row="2" Grid.Column="1" Style="{StaticResource Win11TextBox}" Margin="0,0,20,12"/>
                                    
                                    <TextBlock Grid.Row="2" Grid.Column="2" Text="Status:" VerticalAlignment="Center" Margin="0,0,10,12"/>
                                    <ComboBox x:Name="cmbStatus" Grid.Row="2" Grid.Column="3" Style="{StaticResource Win11ComboBox}" Margin="0,0,0,12">
                                        <ComboBoxItem Content="Alle" IsSelected="True"/>
                                        <ComboBoxItem Content="Aktiviert"/>
                                        <ComboBoxItem Content="Deaktiviert"/>
                                        <ComboBoxItem Content="Gesperrt"/>
                                    </ComboBox>
                                    
                                    <!-- Zeile 4 -->
                                    <TextBlock Grid.Row="3" Grid.Column="0" Text="OU enthält:" VerticalAlignment="Center" Margin="0,0,10,0"/>
                                    <ComboBox x:Name="cmbOU" Grid.Row="3" Grid.Column="1" Style="{StaticResource Win11ComboBox}" Margin="0,0,20,0" IsEditable="True"/>
                                    
                                    <TextBlock Grid.Row="3" Grid.Column="2" Text="Ergebnislimit:" VerticalAlignment="Center" Margin="0,0,10,0"/>
                                    <Grid Grid.Row="3" Grid.Column="3">
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="Auto"/>
                                        </Grid.ColumnDefinitions>
                                        <ComboBox x:Name="cmbResultLimit" Grid.Column="0" Style="{StaticResource Win11ComboBox}" IsEditable="True">
                                            <ComboBoxItem Content="100"/>
                                            <ComboBoxItem Content="500"/>
                                            <ComboBoxItem Content="1000" IsSelected="True"/>
                                            <ComboBoxItem Content="2000"/>
                                            <ComboBoxItem Content="5000"/>
                                            <ComboBoxItem Content="Alle"/>
                                        </ComboBox>
                                        <CheckBox x:Name="chkExtendedAttributes" Grid.Column="1" Style="{StaticResource Win11CheckBox}" 
                                                Content="Erweiterte Attribute" VerticalAlignment="Center" Margin="12,0,0,0"/>
                                    </Grid>
                                </Grid>
                            </StackPanel>
                            
                            <Button x:Name="btnSearch" Grid.Column="1" Content="Suchen" Style="{StaticResource Win11Button}" 
                                    VerticalAlignment="Bottom" Width="140" Margin="20,0,0,0"/>
                        </Grid>
                    </TabItem>
                    <TabItem Header="Erweiterte Filter" Style="{StaticResource Win11TabItem}">
                        <Grid Margin="0,12,0,0">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            
                            <StackPanel Grid.Column="0">
                                <TextBlock Text="Erweiterte Suchparameter" FontSize="20" FontWeight="SemiBold" Margin="0,0,0,16"/>
                                
                                <Grid>
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="Auto"/>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
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
                                    
                                    <!-- Zeile 1 - Benutzer -->
                                    <TextBlock Grid.Row="0" Grid.Column="0" Text="Abteilung:" VerticalAlignment="Center" Margin="0,0,10,12"/>
                                    <ComboBox x:Name="cmbDepartment" Grid.Row="0" Grid.Column="1" Style="{StaticResource Win11ComboBox}" Margin="0,0,20,12" IsEditable="True"/>
                                    
                                    <TextBlock Grid.Row="0" Grid.Column="2" Text="E-Mail enthält:" VerticalAlignment="Center" Margin="0,0,10,12"/>
                                    <TextBox x:Name="txtEmail" Grid.Row="0" Grid.Column="3" Style="{StaticResource Win11TextBox}" Margin="0,0,0,12"/>
                                    
                                    <!-- Zeile 2 - Datum Filter -->
                                    <TextBlock Grid.Row="1" Grid.Column="0" Text="Erstellt nach:" VerticalAlignment="Center" Margin="0,0,10,12"/>
                                    <DatePicker x:Name="dtCreatedAfter" Grid.Row="1" Grid.Column="1" Margin="0,0,20,12" Height="32" Padding="8,0"/>
                                    
                                    <TextBlock Grid.Row="1" Grid.Column="2" Text="Erstellt vor:" VerticalAlignment="Center" Margin="0,0,10,12"/>
                                    <DatePicker x:Name="dtCreatedBefore" Grid.Row="1" Grid.Column="3" Margin="0,0,0,12" Height="32" Padding="8,0"/>
                                    
                                    <!-- Zeile 3 - Passwort-Status -->
                                    <TextBlock Grid.Row="2" Grid.Column="0" Text="Passwort-Status:" VerticalAlignment="Center" Margin="0,0,10,12"/>
                                    <ComboBox x:Name="cmbPasswordStatus" Grid.Row="2" Grid.Column="1" Style="{StaticResource Win11ComboBox}" Margin="0,0,20,12">
                                        <ComboBoxItem Content="Alle" IsSelected="True"/>
                                        <ComboBoxItem Content="Abgelaufen"/>
                                        <ComboBoxItem Content="Läuft nie ab"/>
                                        <ComboBoxItem Content="Muss geändert werden"/>
                                    </ComboBox>
                                    
                                    <TextBlock Grid.Row="2" Grid.Column="2" Text="Passwort zuletzt geändert:" VerticalAlignment="Center" Margin="0,0,10,12"/>
                                    <Grid Grid.Row="2" Grid.Column="3" Margin="0,0,0,12">
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="Auto"/>
                                        </Grid.ColumnDefinitions>
                                        <ComboBox x:Name="cmbPwdLastSet" Grid.Column="0" Style="{StaticResource Win11ComboBox}" IsEditable="True">
                                            <ComboBoxItem Content="Alle"/>
                                            <ComboBoxItem Content="30"/>
                                            <ComboBoxItem Content="60"/>
                                            <ComboBoxItem Content="90"/>
                                            <ComboBoxItem Content="180"/>
                                        </ComboBox>
                                        <TextBlock Grid.Column="1" Text=" Tage" VerticalAlignment="Center" Margin="6,0,0,0"/>
                                    </Grid>
                                    
                                    <!-- Zeile 4 - Gruppenmitgliedschaft -->
                                    <TextBlock Grid.Row="3" Grid.Column="0" Text="Mitglied von Gruppe:" VerticalAlignment="Center" Margin="0,0,10,12"/>
                                    <ComboBox x:Name="cmbMemberOf" Grid.Row="3" Grid.Column="1" Style="{StaticResource Win11ComboBox}" Margin="0,0,20,12" IsEditable="True"/>
                                    
                                    <!-- Zeile 5 - Standort/Land -->
                                    <TextBlock Grid.Row="3" Grid.Column="2" Text="Standort/Büro:" VerticalAlignment="Center" Margin="0,0,10,12"/>
                                    <ComboBox x:Name="cmbLocation" Grid.Row="3" Grid.Column="3" Style="{StaticResource Win11ComboBox}" Margin="0,0,0,12" IsEditable="True"/>
                                    
                                    <!-- Zeile 5 - Benutzerdefinierte Attribute -->
                                    <TextBlock Grid.Row="4" Grid.Column="0" Text="Attribut:" VerticalAlignment="Center" Margin="0,0,10,12"/>
                                    <ComboBox x:Name="cmbCustomAttribute" Grid.Row="4" Grid.Column="1" Style="{StaticResource Win11ComboBox}" Margin="0,0,20,12" IsEditable="True">
                                        <ComboBoxItem Content="department"/>
                                        <ComboBoxItem Content="title"/>
                                        <ComboBoxItem Content="manager"/>
                                        <ComboBoxItem Content="homeDirectory"/>
                                        <ComboBoxItem Content="telephoneNumber"/>
                                        <ComboBoxItem Content="mobile"/>
                                        <ComboBoxItem Content="extensionAttribute1"/>
                                        <ComboBoxItem Content="servicePrincipalName"/>
                                    </ComboBox>
                                    
                                    <TextBlock Grid.Row="4" Grid.Column="2" Text="Wert enthält:" VerticalAlignment="Center" Margin="0,0,10,12"/>
                                    <TextBox x:Name="txtCustomAttributeValue" Grid.Row="4" Grid.Column="3" Style="{StaticResource Win11TextBox}" Margin="0,0,0,12"/>
                                    
                                    <!-- Zeile 6 - LDAP-Filter -->
                                    <TextBlock Grid.Row="5" Grid.Column="0" Text="LDAP-Filter:" VerticalAlignment="Center" Margin="0,0,10,0"/>
                                    <TextBox x:Name="txtLdapFilter" Grid.Row="5" Grid.Column="1" Grid.ColumnSpan="3" Style="{StaticResource Win11TextBox}" Margin="0,0,0,0"/>
                                </Grid>
                            </StackPanel>
                            
                            <Button x:Name="btnAdvancedSearch" Grid.Column="1" Content="Erweitert Suchen" Style="{StaticResource Win11Button}" 
                                    VerticalAlignment="Bottom" Width="140" Margin="20,0,0,0"/>
                        </Grid>
                    </TabItem>
                    <TabItem Header="Computer-Filter" Style="{StaticResource Win11TabItem}">
                        <Grid Margin="0,12,0,0">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            
                            <StackPanel Grid.Column="0">
                                <TextBlock Text="Erweiterte Computer-Filter" FontSize="20" FontWeight="SemiBold" Margin="0,0,0,16"/>
                                
                                <Grid>
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="Auto"/>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                        <ColumnDefinition Width="*"/>
                                    </Grid.ColumnDefinitions>
                                    <Grid.RowDefinitions>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                    </Grid.RowDefinitions>
                                    
                                    <!-- Zeile 1 - Betriebssystem-Version -->
                                    <TextBlock Grid.Row="0" Grid.Column="0" Text="OS-Version:" VerticalAlignment="Center" Margin="0,0,10,12"/>
                                    <ComboBox x:Name="cmbOSVersion" Grid.Row="0" Grid.Column="1" Style="{StaticResource Win11ComboBox}" Margin="0,0,20,12" IsEditable="True">
                                        <ComboBoxItem Content="Alle"/>
                                        <ComboBoxItem Content="Windows 10"/>
                                        <ComboBoxItem Content="Windows 11"/>
                                        <ComboBoxItem Content="Windows Server 2016"/>
                                        <ComboBoxItem Content="Windows Server 2019"/>
                                        <ComboBoxItem Content="Windows Server 2022"/>
                                    </ComboBox>
                                    
                                    <TextBlock Grid.Row="0" Grid.Column="2" Text="Service Pack:" VerticalAlignment="Center" Margin="0,0,10,12"/>
                                    <ComboBox x:Name="cmbServicePack" Grid.Row="0" Grid.Column="3" Style="{StaticResource Win11ComboBox}" Margin="0,0,0,12" IsEditable="True">
                                        <ComboBoxItem Content="Alle"/>
                                        <ComboBoxItem Content="SP1"/>
                                        <ComboBoxItem Content="SP2"/>
                                        <ComboBoxItem Content="SP3"/>
                                    </ComboBox>
                                    
                                    <!-- Zeile 2 - IP-Bereich -->
                                    <TextBlock Grid.Row="1" Grid.Column="0" Text="IP-Adresse beginnt mit:" VerticalAlignment="Center" Margin="0,0,10,12"/>
                                    <TextBox x:Name="txtIPRange" Grid.Row="1" Grid.Column="1" Style="{StaticResource Win11TextBox}" Margin="0,0,20,12"/>
                                    
                                    <TextBlock Grid.Row="1" Grid.Column="2" Text="DNS-Name enthält:" VerticalAlignment="Center" Margin="0,0,10,12"/>
                                    <TextBox x:Name="txtDNSName" Grid.Row="1" Grid.Column="3" Style="{StaticResource Win11TextBox}" Margin="0,0,0,12"/>
                                    
                                    <!-- Zeile 3 - Inaktivität -->
                                    <TextBlock Grid.Row="2" Grid.Column="0" Text="Inaktiv seit:" VerticalAlignment="Center" Margin="0,0,10,12"/>
                                    <Grid Grid.Row="2" Grid.Column="1" Margin="0,0,20,12">
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="Auto"/>
                                        </Grid.ColumnDefinitions>
                                        <ComboBox x:Name="cmbInactiveDays" Grid.Column="0" Style="{StaticResource Win11ComboBox}" IsEditable="True">
                                            <ComboBoxItem Content="Alle"/>
                                            <ComboBoxItem Content="30"/>
                                            <ComboBoxItem Content="60"/>
                                            <ComboBoxItem Content="90"/>
                                            <ComboBoxItem Content="180"/>
                                            <ComboBoxItem Content="365"/>
                                        </ComboBox>
                                        <TextBlock Grid.Column="1" Text=" Tage" VerticalAlignment="Center" Margin="6,0,0,0"/>
                                    </Grid>
                                    
                                    <TextBlock Grid.Row="2" Grid.Column="2" Text="Verwaltet von:" VerticalAlignment="Center" Margin="0,0,10,12"/>
                                    <TextBox x:Name="txtManagedBy" Grid.Row="2" Grid.Column="3" Style="{StaticResource Win11TextBox}" Margin="0,0,0,12"/>
                                    
                                    <!-- Zeile 4 - Weitere Computer-Filter -->
                                    <TextBlock Grid.Row="3" Grid.Column="0" Text="Computer-Rolle:" VerticalAlignment="Center" Margin="0,0,10,0"/>
                                    <ComboBox x:Name="cmbComputerRole" Grid.Row="3" Grid.Column="1" Style="{StaticResource Win11ComboBox}" Margin="0,0,20,0" IsEditable="True">
                                        <ComboBoxItem Content="Alle"/>
                                        <ComboBoxItem Content="Workstation"/>
                                        <ComboBoxItem Content="Server"/>
                                        <ComboBoxItem Content="Domain Controller"/>
                                    </ComboBox>
                                </Grid>
                            </StackPanel>
                            
                            <Button x:Name="btnComputerSearch" Grid.Column="1" Content="Computer Suchen" Style="{StaticResource Win11Button}" 
                                    VerticalAlignment="Bottom" Width="140" Margin="20,0,0,0"/>
                        </Grid>
                    </TabItem>
                    <TabItem Header="Gespeicherte Suchen" Style="{StaticResource Win11TabItem}">
                        <Grid Margin="0,12,0,0">
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>
                            
                            <TextBlock Grid.Row="0" Text="Gespeicherte Suchanfragen" FontSize="20" FontWeight="SemiBold" Margin="0,0,0,16"/>
                            
                            <ListBox x:Name="lstSavedSearches" Grid.Row="1" Margin="0,0,0,16" BorderThickness="1" BorderBrush="#e0e0e0" Background="White" Padding="4"/>
                            
                            <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right">
                                <Button x:Name="btnLoadSearch" Content="Laden" Style="{StaticResource Win11Button}" Width="110" Margin="0,0,10,0"/>
                                <Button x:Name="btnSaveSearch" Content="Aktuelle Suche speichern" Style="{StaticResource Win11Button}" Width="210" Margin="0,0,10,0"/>
                                <Button x:Name="btnDeleteSearch" Content="Löschen" Style="{StaticResource Win11Button}" Width="110" Background="#d83b01"/>
                            </StackPanel>
                        </Grid>
                    </TabItem>
                </TabControl>
            </Grid>
        </Border>
        
        <!-- Inhalt - DataGrid für Suchergebnisse -->
        <Grid Grid.Row="1" Margin="0,0,0,16">
            <TabControl x:Name="tabResults" BorderThickness="0">
                <TabItem Header="Ergebnisse" Style="{StaticResource Win11TabItem}">
                    <Border Style="{StaticResource Win11Card}">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>
                            
                            <Grid Grid.Row="0" Margin="0,0,0,12">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <TextBlock Grid.Column="0" Text="Suchergebnisse" FontSize="18" FontWeight="SemiBold" VerticalAlignment="Center"/>
                                <TextBox x:Name="txtFilter" Grid.Column="1" Style="{StaticResource Win11TextBox}" Margin="16,0" 
                                        ToolTip="In Ergebnissen filtern"/>
                                <StackPanel Grid.Column="2" Orientation="Horizontal">
                                    <Button x:Name="btnRefresh" Style="{StaticResource IconButton}" ToolTip="Aktualisieren" Content="↻" FontSize="15"/>
                                    <Button x:Name="btnSelectColumns" Style="{StaticResource IconButton}" ToolTip="Spalten auswählen" Content="≡" FontSize="16"/>
                                    <Button x:Name="btnExportGrid" Style="{StaticResource IconButton}" ToolTip="Ergebnisse exportieren" Content="↓" FontSize="16"/>
                                </StackPanel>
                            </Grid>
                            
                            <DataGrid x:Name="dgResults" Grid.Row="1" Style="{StaticResource Win11DataGrid}" 
                                      AutoGenerateColumns="False" IsReadOnly="True" VerticalScrollBarVisibility="Auto">
                                <DataGrid.Resources>
                                    <Style TargetType="DataGridColumnHeader" BasedOn="{StaticResource Win11DataGridColumnHeader}"/>
                                </DataGrid.Resources>
                                <DataGrid.ContextMenu>
                                    <ContextMenu>
                                        <MenuItem x:Name="menuDetails" Header="Details anzeigen">
                                            <MenuItem.Icon>
                                                <TextBlock Text="🔍" FontSize="13"/>
                                            </MenuItem.Icon>
                                        </MenuItem>
                                        <MenuItem x:Name="menuReset" Header="Passwort zurücksetzen">
                                            <MenuItem.Icon>
                                                <TextBlock Text="🔑" FontSize="13"/>
                                            </MenuItem.Icon>
                                        </MenuItem> 
                                        <MenuItem x:Name="menuDisable" Header="Deaktivieren">
                                            <MenuItem.Icon>
                                                <TextBlock Text="🔒" FontSize="13"/>
                                            </MenuItem.Icon>
                                        </MenuItem>
                                        <MenuItem x:Name="menuEnable" Header="Aktivieren">
                                            <MenuItem.Icon>
                                                <TextBlock Text="🔓" FontSize="13"/>
                                            </MenuItem.Icon>
                                        </MenuItem>
                                        <MenuItem x:Name="menuUnlock" Header="Konto entsperren">
                                            <MenuItem.Icon>
                                                <TextBlock Text="✓" FontSize="13"/>
                                            </MenuItem.Icon>
                                        </MenuItem>
                                        <Separator/>
                                        <MenuItem x:Name="menuMove" Header="In andere OU verschieben">
                                            <MenuItem.Icon>
                                                <TextBlock Text="📁" FontSize="13"/>
                                            </MenuItem.Icon>
                                        </MenuItem>
                                        <MenuItem x:Name="menuAddToGroup" Header="Zu Gruppe hinzufügen">
                                            <MenuItem.Icon>
                                                <TextBlock Text="➕" FontSize="13"/>
                                            </MenuItem.Icon>
                                        </MenuItem>
                                        <MenuItem x:Name="menuRemoveFromGroup" Header="Aus Gruppe entfernen">
                                            <MenuItem.Icon>
                                                <TextBlock Text="➖" FontSize="13"/>
                                            </MenuItem.Icon>
                                        </MenuItem>
                                        <Separator/>
                                        <MenuItem x:Name="menuRemote" Header="Remote-Verbindung herstellen">
                                            <MenuItem.Icon>
                                                <TextBlock Text="🖧" FontSize="13"/>
                                            </MenuItem.Icon>
                                        </MenuItem>
                                        <MenuItem x:Name="menuPSRemote" Header="PowerShell Remote-Sitzung">
                                            <MenuItem.Icon>
                                                <TextBlock Text="💻" FontSize="13"/>
                                            </MenuItem.Icon>
                                        </MenuItem>
                                        <Separator/>
                                        <MenuItem x:Name="menuExportSelected" Header="Ausgewählte exportieren">
                                            <MenuItem.Icon>
                                                <TextBlock Text="📄" FontSize="13"/>
                                            </MenuItem.Icon>
                                        </MenuItem>
                                    </ContextMenu>
                                </DataGrid.ContextMenu>
                                <DataGrid.Columns>
                                    <DataGridTextColumn Header="Name" Binding="{Binding Name}" Width="150"/>
                                    <DataGridTextColumn Header="Distinguished Name" Binding="{Binding DistinguishedName}" Width="250"/>
                                    <DataGridCheckBoxColumn Header="Aktiviert" Binding="{Binding Enabled}" Width="80"/>
                                    <DataGridTextColumn Header="Beschreibung" Binding="{Binding Description}" Width="200"/>
                                    <DataGridTextColumn Header="Betriebssystem" Binding="{Binding OperatingSystem}" Width="180"/>
                                    <DataGridTextColumn Header="Letzter Login" Binding="{Binding LastLogonDate}" Width="150"/>
                                </DataGrid.Columns>
                            </DataGrid>
                        </Grid>
                    </Border>
                </TabItem>
                <TabItem Header="Detailansicht" Style="{StaticResource Win11TabItem}">
                    <Border Style="{StaticResource Win11Card}">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>
                            
                            <Grid Grid.Row="0" Margin="0,0,0,12">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <Button x:Name="btnBackToResults" Grid.Column="0" Style="{StaticResource IconButton}" Content="←" ToolTip="Zurück zur Ergebnisliste" FontSize="16"/>
                                <TextBlock Grid.Column="1" x:Name="txtSelectedObjectName" Text="Objektdetails" FontSize="18" FontWeight="SemiBold" VerticalAlignment="Center" Margin="12,0"/>
                                <StackPanel Grid.Column="2" Orientation="Horizontal">
                                    <Button x:Name="btnEditObject" Style="{StaticResource IconButton}" ToolTip="Objekt bearbeiten" Content="✎" FontSize="16"/>
                                    <Button x:Name="btnShowInADUC" Style="{StaticResource IconButton}" ToolTip="In AD-Benutzer und -Computer anzeigen" Content="🔍" FontSize="16"/>
                                </StackPanel>
                            </Grid>
                            
                            <TabControl Grid.Row="1" BorderThickness="0" Background="Transparent">
                                <TabItem Header="Allgemeine Info" Style="{StaticResource Win11TabItem}">
                                    <ScrollViewer VerticalScrollBarVisibility="Auto">
                                        <Grid x:Name="gridGeneralInfo" Margin="10">
                                            <!-- Wird dynamisch mit den Objekteigenschaften gefüllt -->
                                        </Grid>
                                    </ScrollViewer>
                                </TabItem>
                                <TabItem Header="Attribute" Style="{StaticResource Win11TabItem}">
                                    <Grid>
                                        <Grid.RowDefinitions>
                                            <RowDefinition Height="Auto"/>
                                            <RowDefinition Height="*"/>
                                        </Grid.RowDefinitions>
                                        <TextBox Grid.Row="0" x:Name="txtAttributeFilter" Style="{StaticResource Win11TextBox}" 
                                                Margin="0,0,0,12" ToolTip="Nach Attributen filtern" />
                                        <DataGrid Grid.Row="1" x:Name="dgAttributes" Style="{StaticResource Win11DataGrid}"
                                                AutoGenerateColumns="False" IsReadOnly="True" VerticalScrollBarVisibility="Auto">
                                            <DataGrid.Resources>
                                                <Style TargetType="DataGridColumnHeader" BasedOn="{StaticResource Win11DataGridColumnHeader}"/>
                                            </DataGrid.Resources>
                                            <DataGrid.Columns>
                                                <DataGridTextColumn Header="Attribut" Binding="{Binding Name}" Width="220"/>
                                                <DataGridTextColumn Header="Wert" Binding="{Binding Value}" Width="*"/>
                                            </DataGrid.Columns>
                                        </DataGrid>
                                    </Grid>
                                </TabItem>
                                <TabItem Header="Mitgliedschaften" Style="{StaticResource Win11TabItem}" x:Name="tabMembership">
                                    <Grid>
                                        <Grid.RowDefinitions>
                                            <RowDefinition Height="*"/>
                                            <RowDefinition Height="Auto"/>
                                            <RowDefinition Height="*"/>
                                        </Grid.RowDefinitions>
                                        
                                        <Grid Grid.Row="0">
                                            <Grid.RowDefinitions>
                                                <RowDefinition Height="Auto"/>
                                                <RowDefinition Height="*"/>
                                            </Grid.RowDefinitions>
                                            <TextBlock Grid.Row="0" Text="Mitglied von" FontWeight="SemiBold" Margin="0,0,0,6"/>
                                            <ListBox Grid.Row="1" x:Name="lbMemberOf" Background="White" BorderBrush="#e0e0e0" BorderThickness="1"/>
                                        </Grid>
                                        
                                        <GridSplitter Grid.Row="1" Height="6" HorizontalAlignment="Stretch" Background="#e8e8e8" Margin="0,12"/>
                                        
                                        <Grid Grid.Row="2">
                                            <Grid.RowDefinitions>
                                                <RowDefinition Height="Auto"/>
                                                <RowDefinition Height="*"/>
                                            </Grid.RowDefinitions>
                                            <TextBlock Grid.Row="0" Text="Mitglieder (für Gruppen)" FontWeight="SemiBold" Margin="0,0,0,6"/>
                                            <ListBox Grid.Row="1" x:Name="lbMembers" Background="White" BorderBrush="#e0e0e0" BorderThickness="1"/>
                                        </Grid>
                                    </Grid>
                                </TabItem>
                            </TabControl>
                        </Grid>
                    </Border>
                </TabItem>
            </TabControl>
        </Grid>
        
        <!-- Footer mit Status und Aktionsknöpfen -->
        <Border Grid.Row="2" Style="{StaticResource Win11Card}" Padding="12">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                
                <StackPanel Grid.Column="0" Orientation="Horizontal">
                    <TextBlock x:Name="txtStatus" Text="Bereit" VerticalAlignment="Center" FontSize="13"/>
                    <TextBlock x:Name="txtResultCount" Text=" | 0 Ergebnisse" VerticalAlignment="Center" Margin="16,0,0,0" FontSize="13"/>
                </StackPanel>
                
                <StackPanel Grid.Column="1" Orientation="Horizontal">
                    <Button x:Name="btnExport" Content="Exportieren" Style="{StaticResource Win11Button}" Margin="0,0,12,0"/>
                    <Button x:Name="btnSettings" Content="Einstellungen" Style="{StaticResource Win11Button}"/>
                </StackPanel>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

# XAML laden und in WPF Objekt umwandeln
try {
    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $window = [Windows.Markup.XamlReader]::Load($reader)
} catch {
    [System.Windows.MessageBox]::Show("XAML konnte nicht geladen werden: $_", "XAML Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    exit
}

# UI-Elemente auslesen und als Variablen bereitstellen
$xaml.SelectNodes("//*[@*[contains(translate(name(.),'x','X'),'Name')]]") | ForEach-Object {
    $name = $_.Name
    $variable = Get-Variable -Name $name -ErrorAction SilentlyContinue
    if ($variable) {
        Set-Variable -Name $name -Value $window.FindName($name)
    } else {
        New-Variable -Name $name -Value $window.FindName($name) -Scope Script
    }
}

# Globale Variablen
$script:adSearchResults = @()
$script:adDomain = New-Object System.DirectoryServices.DirectoryEntry
$script:searchCompleted = $false
$script:exportPath = "$env:USERPROFILE\Documents"
$script:runspacePool = $null
$script:searchJob = $null
$script:cancelSearch = $false

# Runspace-Pool für Asynchrone Operationen initialisieren
function Initialize-RunspacePool {
    $sessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    $runspacePool = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, 2, $sessionState, $Host)
    $runspacePool.Open()
    return $runspacePool
}

# Event-Handler für Abbruch-Button
$btnSearch.Add_Click({
    if ($script:searchJob) {
        # Wenn bereits eine Suche läuft, als Abbruch-Button nutzen
        $script:cancelSearch = $true
        $txtStatus.Text = "Suche wird abgebrochen..."
        $btnSearch.IsEnabled = $false
        return
    }
    
    # Sonst normale Suche starten
    Search-AD
})

# Funktion zum Laden der Comboboxen mit AD-Daten
function Initialize-ComboBoxes {
    try {
        # Status-Update
        $txtStatus.Text = "Lade AD-Daten..."
        $btnSearch.IsEnabled = $false
        $btnAdvancedSearch.IsEnabled = $false
        $btnComputerSearch.IsEnabled = $false
        
        # Asynchrones Laden der AD-Daten
        $script:runspacePool = Initialize-RunspacePool
        
        $scriptBlock = {
            param($syncHash)
            
            try {
                # Betriebssysteme laden (max. 100 für Performance)
                $osList = Get-ADComputer -Filter * -Properties OperatingSystem -ResultSetSize 100 | 
                          Where-Object { $null -ne $_.OperatingSystem } | 
                          Select-Object -ExpandProperty OperatingSystem -Unique | 
                          Sort-Object
                $syncHash.osList = $osList
                
                # OUs laden (mit Limit für bessere Performance)
                $ouList = Get-ADOrganizationalUnit -Filter * -ResultSetSize 200 | 
                          Select-Object -ExpandProperty Name -Unique |
                          Sort-Object
                $syncHash.ouList = $ouList
                
                # Abteilungen laden (mit Limit für bessere Performance)
                $deptList = Get-ADUser -Filter * -Properties Department -ResultSetSize 200 | 
                           Where-Object { $null -ne $_.Department } | 
                           Select-Object -ExpandProperty Department -Unique |
                           Sort-Object
                $syncHash.deptList = $deptList
                
                # Gruppen für Mitgliedschaften laden (nur die wichtigsten für bessere Performance)
                $groupList = Get-ADGroup -Filter "member -like '*'" -ResultSetSize 100 | 
                            Select-Object -ExpandProperty Name |
                            Sort-Object
                $syncHash.groupList = $groupList
                
                # Standorte laden (mit Limit für bessere Performance)
                $locationList = Get-ADUser -Filter * -Properties physicalDeliveryOfficeName -ResultSetSize 100 | 
                                Where-Object { $null -ne $_.physicalDeliveryOfficeName } | 
                                Select-Object -ExpandProperty physicalDeliveryOfficeName -Unique |
                                Sort-Object
                $syncHash.locationList = $locationList
                
                $syncHash.success = $true
            } catch {
                $syncHash.success = $false
                $syncHash.error = $_
            }
        }
        
        $syncHash = [hashtable]::Synchronized(@{
            osList = @()
            ouList = @()
            deptList = @()
            groupList = @()
            locationList = @()
            success = $false
            error = $null
        })
        
        $powershell = [powershell]::Create().AddScript($scriptBlock).AddArgument($syncHash)
        $powershell.RunspacePool = $script:runspacePool
        $handle = $powershell.BeginInvoke()
        
        # Timer für Statusabfrage
        $timer = New-Object System.Windows.Threading.DispatcherTimer
        $timer.Interval = [TimeSpan]::FromMilliseconds(500)
        $timer.Add_Tick({
            if ($handle.IsCompleted) {
                $timer.Stop()
                
                if ($syncHash.success) {
                    # Comboboxen mit den geladenen Daten füllen
                    foreach ($os in $syncHash.osList) {
                        [void]$cmbOS.Items.Add($os)
                    }
                    
                    foreach ($ou in $syncHash.ouList) {
                        [void]$cmbOU.Items.Add($ou)
                    }
                    
                    foreach ($dept in $syncHash.deptList) {
                        [void]$cmbDepartment.Items.Add($dept)
                    }
                    
                    foreach ($group in $syncHash.groupList) {
                        [void]$cmbMemberOf.Items.Add($group)
                    }
                    
                    foreach ($location in $syncHash.locationList) {
                        [void]$cmbLocation.Items.Add($location)
                    }
                    
                    # Status-Update
                    $txtStatus.Text = "AD-Daten erfolgreich geladen"
                } else {
                    [System.Windows.MessageBox]::Show("Fehler beim Laden der AD-Daten: $($syncHash.error)", "Fehler", 
                        [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                    $txtStatus.Text = "Fehler beim Laden der AD-Daten"
                }
                
                # Buttons wieder aktivieren
                $btnSearch.IsEnabled = $true
                $btnAdvancedSearch.IsEnabled = $true
                $btnComputerSearch.IsEnabled = $true
                
                # Ressourcen aufräumen
                $powershell.EndInvoke($handle)
                $powershell.Dispose()
            }
        })
        $timer.Start()
    } catch {
        $btnSearch.IsEnabled = $true
        $btnAdvancedSearch.IsEnabled = $true
        $btnComputerSearch.IsEnabled = $true
        [System.Windows.MessageBox]::Show("Fehler beim Laden der AD-Daten: $_", "Fehler", 
            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        $txtStatus.Text = "Fehler beim Laden der AD-Daten"
    }
}

# Funktion zum Escapen von LDAP-Sonderzeichen
function Escape-LdapFilterString {
    param (
        [string]$InputString
    )
    
    if ([string]::IsNullOrEmpty($InputString)) {
        return $InputString
    }
    
    return $InputString.Replace('(', '\28').Replace(')', '\29').Replace('*', '\2A').
                       Replace('/', '\2F').Replace('\', '\5C').Replace([char]0, '\00')
}

# Funktion zum Erstellen des LDAP-Filters
function Build-LDAPFilter {
    try {
        $objectType = ($cmbObjectType.SelectedItem).Content
        $name = $txtName.Text.Trim()
        $description = $txtDescription.Text.Trim()
        $ouFilter = $cmbOU.Text.Trim()
        $osFilter = $cmbOS.Text.Trim()
        $statusFilter = ($cmbStatus.SelectedItem).Content
        
        # Basis-Filter nach Objekttyp
        switch ($objectType) {
            "Benutzer" { $baseFilter = "(objectClass=user)(objectCategory=person)" }
            "Computer" { $baseFilter = "(objectClass=computer)" }
            "Gruppen" { $baseFilter = "(objectClass=group)" }
            "OUs" { $baseFilter = "(objectClass=organizationalUnit)" }
            default { $baseFilter = "(objectClass=user)(objectCategory=person)" }
        }
        
        # Filter kombinieren
        $filter = "(&$baseFilter"
        
        if ($name) { 
            # Escapen von Sonderzeichen im Filter
            $escapedName = Escape-LdapFilterString -InputString $name
            $filter += "(name=*$escapedName*)" 
        }
        if ($description) { 
            $escapedDesc = Escape-LdapFilterString -InputString $description
            $filter += "(description=*$escapedDesc*)" 
        }
        
        # Status-Filter
        if ($statusFilter -eq "Aktiviert") { $filter += "(!userAccountControl:1.2.840.113556.1.4.803:=2)" }
        elseif ($statusFilter -eq "Deaktiviert") { $filter += "(userAccountControl:1.2.840.113556.1.4.803:=2)" }
        elseif ($statusFilter -eq "Gesperrt") { $filter += "(lockoutTime>=1)" }
        
        # Betriebssystem-Filter (für Computer)
        if ($osFilter -and $objectType -eq "Computer") { 
            $escapedOS = Escape-LdapFilterString -InputString $osFilter
            $filter += "(operatingSystem=*$escapedOS*)" 
        }
        
        # OU-Filter
        if ($ouFilter) { 
            $escapedOU = Escape-LdapFilterString -InputString $ouFilter
            $filter += "(|(ou=*$escapedOU*)(distinguishedName=*$escapedOU*))" 
        }
        
        # Erweiterte Filter hinzufügen (Tab: Erweiterte Filter)
        if ($objectType -eq "Benutzer") {
            $email = $txtEmail.Text.Trim()
            $department = $cmbDepartment.Text.Trim()
            $memberOf = $cmbMemberOf.Text.Trim()
            
            if ($email) { 
                $escapedEmail = Escape-LdapFilterString -InputString $email
                $filter += "(mail=*$escapedEmail*)" 
            }
            if ($department) { 
                $escapedDept = Escape-LdapFilterString -InputString $department
                $filter += "(department=*$escapedDept*)" 
            }
            if ($memberOf) {
                # Gruppensuche verbessern - nach DN suchen
                try {
                    $group = Get-ADGroup -Filter "Name -eq '$memberOf'" -ErrorAction SilentlyContinue
                    if ($group) {
                        $filter += "(memberOf=$($group.DistinguishedName))"
                    } else {
                        $escapedMember = Escape-LdapFilterString -InputString $memberOf
                        $filter += "(memberOf=*$escapedMember*)"
                    }
                } catch {
                    # Fallback bei Fehler
                    $escapedMember = Escape-LdapFilterString -InputString $memberOf
                    $filter += "(memberOf=*$escapedMember*)"
                }
            }
        }
        
        # Computer-spezifische Filter (Tab: Computer-Filter)
        if ($objectType -eq "Computer") {
            $osVersion = $cmbOSVersion.Text.Trim()
            $ipRange = $txtIPRange.Text.Trim()
            $dnsName = $txtDNSName.Text.Trim()
            
            if ($osVersion -and $osVersion -ne "Alle") { 
                $escapedOSV = Escape-LdapFilterString -InputString $osVersion
                $filter += "(operatingSystemVersion=*$escapedOSV*)" 
            }
            if ($ipRange) { 
                $escapedIP = Escape-LdapFilterString -InputString $ipRange
                $filter += "(networkAddress=*$escapedIP*)" 
            }
            if ($dnsName) { 
                $escapedDNS = Escape-LdapFilterString -InputString $dnsName
                $filter += "(dNSHostName=*$escapedDNS*)" 
            }
        }
        
        # Benutzerdefiniertes Attribut
        $customAttr = $cmbCustomAttribute.Text.Trim()
        $customValue = $txtCustomAttributeValue.Text.Trim()
        if ($customAttr -and $customValue) {
            $escapedAttr = Escape-LdapFilterString -InputString $customAttr
            $escapedVal = Escape-LdapFilterString -InputString $customValue
            $filter += "($escapedAttr=*$escapedVal*)"
        }
        
        # Manuellen LDAP-Filter hinzufügen, falls vorhanden
        $manualFilter = $txtLdapFilter.Text.Trim()
        if ($manualFilter) {
            # Validieren des manuellen LDAP-Filters
            try {
                $testFilter = "(&(objectClass=*)$manualFilter)"
                $testSearcher = New-Object System.DirectoryServices.DirectorySearcher
                $testSearcher.Filter = $testFilter
                $testSearcher.FindOne() | Out-Null
                
                # Filter ist gültig, kann verwendet werden
                if ($manualFilter.StartsWith("(") -and $manualFilter.EndsWith(")")) {
                    $filter += $manualFilter
                } else {
                    $filter += "($manualFilter)"
                }
            } catch {
                [System.Windows.MessageBox]::Show("Der manuelle LDAP-Filter ist ungültig und wurde nicht angewendet.", "Ungültiger Filter", 
                    [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            }
        }
        
        $filter += ")"
        return $filter
    } catch {
        [System.Windows.MessageBox]::Show("Fehler beim Erstellen des LDAP-Filters: $_", "Fehler", 
            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        return "(&(objectClass=user)(objectCategory=person)(name=ERROR_IN_FILTER))"
    }
}

# Funktion zur AD-Suche mit asynchroner Verarbeitung
function Search-AD {
    # UI-Elemente sperren während der Suche
    $btnSearch.Content = "Abbrechen"
    $btnAdvancedSearch.IsEnabled = $false
    $btnComputerSearch.IsEnabled = $false
    $btnExport.IsEnabled = $false
    $script:cancelSearch = $false
    
    # Status-Update
    $txtStatus.Text = "Suche läuft..."
    
    try {
        $ldapFilter = Build-LDAPFilter
        $objectType = ($cmbObjectType.SelectedItem).Content
        $resultLimit = ($cmbResultLimit.SelectedItem).Content
        
        if ($resultLimit -eq "Alle") { $resultLimit = 0 }
        else { $resultLimit = [int]$resultLimit }
        
        # Asynchrone Suche starten
        $script:runspacePool = Initialize-RunspacePool
        
        $scriptBlock = {
            param($ldapFilter, $resultLimit, $objectType, $extendedAttributes, $syncHash)
            
            try {
                # AD-Searcher konfigurieren
                $searchRoot = New-Object System.DirectoryServices.DirectoryEntry
                $searcher = New-Object System.DirectoryServices.DirectorySearcher($searchRoot)
                $searcher.Filter = $ldapFilter
                $searcher.PageSize = 1000
                
                if ($resultLimit -gt 0) { $searcher.SizeLimit = $resultLimit }
                
                # Attribute festlegen, die geladen werden sollen
                $propertiesToLoad = @("name", "distinguishedName", "description", "userAccountControl", "operatingSystem", "lastLogonTimestamp")
                
                # Erweiterte Attribute bei Bedarf laden
                if ($extendedAttributes) {
                    $propertiesToLoad += @("mail", "department", "title", "telephoneNumber", "mobile", "manager", "memberOf")
                    if ($objectType -eq "Computer") {
                        $propertiesToLoad += @("operatingSystemVersion", "dNSHostName", "servicePrincipalName")
                    }
                }
                
                $searcher.PropertiesToLoad.AddRange($propertiesToLoad)
                
                # Suche ausführen
                $searchResults = $searcher.FindAll()
                
                # Ergebnisse konvertieren
                $adObjects = @()
                foreach ($result in $searchResults) {
                    if ($syncHash.cancelSearch) {
                        break
                    }
                    
                    $adObject = New-Object PSObject
                    $enabled = $true
                    if ($result.Properties["userAccountControl"]) {
                        $uac = $result.Properties["userAccountControl"][0]
                        $enabled = -not [bool]($uac -band 2)
                    }
                    
                    $lastLogon = $null
                    if ($result.Properties["lastLogonTimestamp"]) {
                        try {
                            $lastLogonTime = [DateTime]::FromFileTime($result.Properties["lastLogonTimestamp"][0])
                            $lastLogon = $lastLogonTime.ToString("yyyy-MM-dd HH:mm:ss")
                        } catch { 
                            $lastLogon = "Niemals" 
                        }
                    }
                    
                    # Eigenschaften zum Objekt hinzufügen
                    $adObject | Add-Member -Type NoteProperty -Name "Name" -Value $result.Properties["name"][0]
                    $adObject | Add-Member -Type NoteProperty -Name "DistinguishedName" -Value $result.Properties["distinguishedname"][0]
                    $adObject | Add-Member -Type NoteProperty -Name "Enabled" -Value $enabled
                    
                    if ($result.Properties["description"]) {
                        $adObject | Add-Member -Type NoteProperty -Name "Description" -Value $result.Properties["description"][0]
                    } else {
                        $adObject | Add-Member -Type NoteProperty -Name "Description" -Value ""
                    }
                    
                    if ($result.Properties["operatingsystem"]) {
                        $adObject | Add-Member -Type NoteProperty -Name "OperatingSystem" -Value $result.Properties["operatingsystem"][0]
                    } else {
                        $adObject | Add-Member -Type NoteProperty -Name "OperatingSystem" -Value ""
                    }
                    
                    $adObject | Add-Member -Type NoteProperty -Name "LastLogonDate" -Value $lastLogon
                    
                    # DirectoryEntry-Objekt für spätere Nutzung speichern
                    $adObject | Add-Member -Type NoteProperty -Name "ADObject" -Value $result.GetDirectoryEntry()
                    
                    $adObjects += $adObject
                    
                    # Status aktualisieren (alle 100 Objekte)
                    if ($adObjects.Count % 100 -eq 0) {
                        $syncHash.status = "Suche läuft... $($adObjects.Count) Objekte gefunden"
                    }
                }
                
                $searchResults.Dispose()
                $searcher.Dispose()
                
                $syncHash.adObjects = $adObjects
                $syncHash.success = $true
            } catch {
                $syncHash.success = $false
                $syncHash.error = $_
            }
        }
        
        $syncHash = [hashtable]::Synchronized(@{
            adObjects = @()
            success = $false
            error = $null
            status = "Suche läuft..."
            cancelSearch = $false
        })
        
        $powershell = [powershell]::Create().AddScript($scriptBlock).AddArgument($ldapFilter).AddArgument($resultLimit).AddArgument($objectType).AddArgument($chkExtendedAttributes.IsChecked).AddArgument($syncHash)
        $powershell.RunspacePool = $script:runspacePool
        $handle = $powershell.BeginInvoke()
        
        $script:searchJob = @{
            Powershell = $powershell
            Handle = $handle
            SyncHash = $syncHash
        }
        
        # Timer für Statusabfrage
        $timer = New-Object System.Windows.Threading.DispatcherTimer
        $timer.Interval = [TimeSpan]::FromMilliseconds(500)
        $timer.Add_Tick({
            # Status aktualisieren
            $txtStatus.Text = $script:searchJob.SyncHash.status
            
            if ($script:cancelSearch) {
                $script:searchJob.SyncHash.cancelSearch = $true
            }
            
            if ($script:searchJob.Handle.IsCompleted) {
                $timer.Stop()
                
                # UI-Elemente entsperren
                $btnSearch.Content = "Suchen"
                $btnSearch.IsEnabled = $true
                $btnAdvancedSearch.IsEnabled = $true
                $btnComputerSearch.IsEnabled = $true
                $btnExport.IsEnabled = $true
                
                if ($script:searchJob.SyncHash.success) {
                    # Ergebnisse anzeigen
                    $script:adSearchResults = $script:searchJob.SyncHash.adObjects
                    $dgResults.ItemsSource = $script:searchJob.SyncHash.adObjects
                    $txtResultCount.Text = " | " + $script:searchJob.SyncHash.adObjects.Count + " Ergebnisse"
                    $script:searchCompleted = $true
                    $txtStatus.Text = "Suche abgeschlossen"
                } else {
                    if ($script:cancelSearch) {
                        $txtStatus.Text = "Suche wurde abgebrochen"
                    } else {
                        [System.Windows.MessageBox]::Show("Fehler bei der AD-Suche: $($script:searchJob.SyncHash.error)", "Fehler", 
                            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                        $txtStatus.Text = "Fehler bei der Suche"
                    }
                }
                
                # Ressourcen aufräumen
                $script:searchJob.Powershell.EndInvoke($script:searchJob.Handle)
                $script:searchJob.Powershell.Dispose()
                $script:searchJob = $null
            }
        })
        $timer.Start()
    } catch {
        # Bei Fehler UI wiederherstellen
        $btnSearch.Content = "Suchen"
        $btnSearch.IsEnabled = $true
        $btnAdvancedSearch.IsEnabled = $true
        $btnComputerSearch.IsEnabled = $true
        $btnExport.IsEnabled = $true
        
        [System.Windows.MessageBox]::Show("Fehler beim Starten der AD-Suche: $_", "Fehler", 
            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        $txtStatus.Text = "Fehler beim Starten der Suche"
    }
}

# Event-Handler für erweiterte Suche
$btnAdvancedSearch.Add_Click({
    Search-AD
})

# Event-Handler für Computer-Suche
$btnComputerSearch.Add_Click({
    if ($cmbObjectType.SelectedIndex -ne 1) {
        $cmbObjectType.SelectedIndex = 1  # Computer auswählen
    }
    Search-AD
})

# Event-Handler für Export-Button
$btnExport.Add_Click({
    if (-not $script:searchCompleted -or $script:adSearchResults.Count -eq 0) {
        [System.Windows.MessageBox]::Show("Es sind keine Suchergebnisse zum Exportieren vorhanden.", "Information", 
            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        return
    }
    
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "CSV-Dateien (*.csv)|*.csv|Excel-Dateien (*.xlsx)|*.xlsx|Alle Dateien (*.*)|*.*"
    $saveDialog.Title = "Suchergebnisse exportieren"
    $saveDialog.InitialDirectory = $script:exportPath
    
    if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            $filePath = $saveDialog.FileName
            $script:exportPath = [System.IO.Path]::GetDirectoryName($filePath)
            
            if ($filePath.EndsWith(".csv")) {
                # Exportierte Objekte bereinigen - ADObject-Eigenschaft entfernen
                $exportObjects = $script:adSearchResults | Select-Object -Property * -ExcludeProperty ADObject
                $exportObjects | Export-Csv -Path $filePath -NoTypeInformation -Encoding UTF8
            }
            elseif ($filePath.EndsWith(".xlsx")) {
                # Excel-Export durch Installation des ImportExcel-Moduls
                if (Get-Module -ListAvailable ImportExcel) {
                    $exportObjects = $script:adSearchResults | Select-Object -Property * -ExcludeProperty ADObject
                    $exportObjects | Export-Excel -Path $filePath -AutoSize -WorksheetName "AD-Suchergebnisse"
                }
                else {
                    [System.Windows.MessageBox]::Show("Das ImportExcel-Modul ist nicht installiert. Bitte installieren Sie es mit 'Install-Module ImportExcel'", 
                        "Modul fehlt", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                    return
                }
            }
            
            [System.Windows.MessageBox]::Show("Suchergebnisse wurden erfolgreich nach '$filePath' exportiert.", 
                "Export erfolgreich", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        }
        catch {
            [System.Windows.MessageBox]::Show("Fehler beim Exportieren: $_", "Fehler", 
                [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        }
    }
})

# Event-Handler für Details anzeigen (Kontextmenü)
$menuDetails.Add_Click({
    if ($dgResults.SelectedItem) {
        ShowObjectDetails $dgResults.SelectedItem
    }
})

# Funktion: Objektdetails anzeigen
function ShowObjectDetails($adObject) {
    try {
        # Zum Details-Tab wechseln
        $tabResults.SelectedIndex = 1
        
        # Objektname anzeigen
        $txtSelectedObjectName.Text = $adObject.Name
        
        # Allgemeine Informationen anzeigen
        PopulateGeneralInfo $adObject
        
        # Attribute laden
        LoadObjectAttributes $adObject
        
        # Mitgliedschaften laden (für Benutzer und Computer)
        $lbMemberOf.Items.Clear()
        if ($adObject.ADObject.Properties["memberof"]) {
            foreach ($group in $adObject.ADObject.Properties["memberof"]) {
                try {
                    $groupName = $group.Split(",")[0].Replace("CN=", "")
                    [void]$lbMemberOf.Items.Add($groupName)
                } catch {
                    # Fehlerhafte DN überspringen
                    continue
                }
            }
        }
        
        # Für Gruppen: Mitglieder anzeigen
        if ($adObject.ADObject.Properties["objectClass"] -contains "group") {
            $tabMembership.Visibility = "Visible"
            
            try {
                # Asynchrones Laden der Gruppenmitglieder
                $lbMembers.Items.Clear()
                [void]$lbMembers.Items.Add("Lade Gruppenmitglieder...")
                
                $script:runspacePool = Initialize-RunspacePool
                
                $scriptBlock = {
                    param($distinguishedName)
                    
                    try {
                        $groupMembers = Get-ADGroupMember -Identity $distinguishedName -ErrorAction Stop
                        return $groupMembers | Select-Object Name
                    } catch {
                        return @{Error = $true; Message = $_.Exception.Message}
                    }
                }
                
                $powershell = [powershell]::Create().AddScript($scriptBlock).AddArgument($adObject.DistinguishedName)
                $powershell.RunspacePool = $script:runspacePool
                $handle = $powershell.BeginInvoke()
                
                # Timer für Statusabfrage
                $timer = New-Object System.Windows.Threading.DispatcherTimer
                $timer.Interval = [TimeSpan]::FromMilliseconds(500)
                $timer.Add_Tick({
                    if ($handle.IsCompleted) {
                        $timer.Stop()
                        
                        $result = $powershell.EndInvoke($handle)
                        $powershell.Dispose()
                        
                        $lbMembers.Items.Clear()
                        
                        if ($result -and $result[0] -and $result[0].Error) {
                            [void]$lbMembers.Items.Add("Fehler beim Laden: $($result[0].Message)")
                        } else {
                            foreach ($member in $result) {
                                [void]$lbMembers.Items.Add($member.Name)
                            }
                            
                            if ($lbMembers.Items.Count -eq 0) {
                                [void]$lbMembers.Items.Add("Keine Mitglieder")
                            }
                        }
                    }
                })
                $timer.Start()
            }
            catch {
                $lbMembers.Items.Clear()
                [void]$lbMembers.Items.Add("Fehler beim Laden der Gruppenmitglieder")
            }
        }
        else {
            $tabMembership.Visibility = "Collapsed"
        }
    } catch {
        [System.Windows.MessageBox]::Show("Fehler beim Anzeigen der Objektdetails: $_", "Fehler", 
            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
}

# Funktion: Allgemeine Informationen anzeigen
function PopulateGeneralInfo($adObject) {
    try {
        # Grid für Allgemeine Informationen leeren
        $gridGeneralInfo.Children.Clear()
        $gridGeneralInfo.RowDefinitions.Clear()
        
        # Grundlegende Eigenschaften für die Anzeige definieren
        $properties = @(
            @{Name="Name"; Property="Name"},
            @{Name="Distinguished Name"; Property="DistinguishedName"},
            @{Name="Beschreibung"; Property="Description"},
            @{Name="Status"; Expression={if ($_.Enabled) {"Aktiviert"} else {"Deaktiviert"}}},
            @{Name="Betriebssystem"; Property="OperatingSystem"},
            @{Name="Letzter Login"; Property="LastLogonDate"}
        )
        
        # Zeilen für das Grid definieren
        for ($i = 0; $i -lt $properties.Count; $i++) {
            $gridGeneralInfo.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition))
        }
        
        # Spalten für das Grid definieren
        $gridGeneralInfo.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{Width = "Auto"}))
        $gridGeneralInfo.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition))
        
        # Eigenschaften ins Grid einfügen
        for ($i = 0; $i -lt $properties.Count; $i++) {
            $prop = $properties[$i]
            
            # Label für die Eigenschaft
            $label = New-Object System.Windows.Controls.TextBlock
            $label.Text = $prop.Name + ":"
            $label.FontWeight = "SemiBold"
            $label.Margin = "0,5,10,5"
            $label.VerticalAlignment = "Center"
            
            # Wert der Eigenschaft
            $value = New-Object System.Windows.Controls.TextBlock
            
            if ($prop.Expression) {
                # Verwenden einer Berechnungsvorschrift
                $value.Text = & $prop.Expression $adObject
            } else {
                # Direkter Zugriff auf die Eigenschaft
                $value.Text = $adObject.($prop.Property)
            }
            
            $value.Margin = "0,5,0,5"
            $value.VerticalAlignment = "Center"
            $value.TextWrapping = "Wrap"
            
            # Elemente zum Grid hinzufügen
            [System.Windows.Controls.Grid]::SetRow($label, $i)
            [System.Windows.Controls.Grid]::SetColumn($label, 0)
            [System.Windows.Controls.Grid]::SetRow($value, $i)
            [System.Windows.Controls.Grid]::SetColumn($value, 1)
            
            $gridGeneralInfo.Children.Add($label)
            $gridGeneralInfo.Children.Add($value)
        }
    } catch {
        [System.Windows.MessageBox]::Show("Fehler beim Anzeigen der allgemeinen Informationen: $_", "Fehler", 
            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
}

# Funktion: Objektattribute laden
function LoadObjectAttributes($adObject) {
    try {
        $attributes = @()
        
        # Alle Eigenschaften des Objekts durchlaufen
        foreach ($property in $adObject.ADObject.Properties.PropertyNames) {
            $value = $adObject.ADObject.Properties[$property]
            
            # Wenn es mehrere Werte gibt, diese mit Komma trennen
            if ($value -is [System.DirectoryServices.ResultPropertyValueCollection] -and $value.Count -gt 1) {
                $valueString = $value -join ", "
            } else {
                # Spezielle Formatierung für bestimmte Attributtypen
                switch ($property) {
                    "lastLogonTimestamp" { 
                        try {
                            $valueString = [DateTime]::FromFileTime($value[0]).ToString("yyyy-MM-dd HH:mm:ss")
                        } catch {
                            $valueString = $value
                        }
                    }
                    "pwdLastSet" {
                        try {
                            $valueString = [DateTime]::FromFileTime($value[0]).ToString("yyyy-MM-dd HH:mm:ss")
                        } catch {
                            $valueString = $value
                        }
                    }
                    "userAccountControl" {
                        $uacValue = $value[0]
                        $valueString = "$uacValue - " + (ConvertFrom-UAC $uacValue)
                    }
                    default {
                        $valueString = $value
                    }
                }
            }
            
            $attributes += [PSCustomObject]@{
                Name = $property
                Value = $valueString
            }
        }
        
        # DataGrid mit Attributen füllen
        $dgAttributes.ItemsSource = $attributes
        
        # Filter-Ereignis für die Textbox einrichten, falls noch nicht vorhanden
        if (-not $txtAttributeFilter.Tag) {
            $txtAttributeFilter.Tag = $true
            $txtAttributeFilter.Add_TextChanged({
                $filter = $txtAttributeFilter.Text
                if (-not [string]::IsNullOrWhiteSpace($filter)) {
                    $view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($dgAttributes.ItemsSource)
                    $view.Filter = {
                        $item = $_
                        ($item.Name -like "*$filter*") -or ($item.Value -like "*$filter*")
                    }
                } else {
                    $view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($dgAttributes.ItemsSource)
                    $view.Filter = $null
                }
            })
        }
    }
    catch {
        [System.Windows.MessageBox]::Show("Fehler beim Laden der Attribute: $_", "Fehler", 
            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
}

# Hilfsfunktion zum Konvertieren der userAccountControl-Werte
function ConvertFrom-UAC {
    param (
        [int]$UAC
    )
    
    $flags = @()
    
    if ($UAC -band 0x0001) { $flags += "SCRIPT" }
    if ($UAC -band 0x0002) { $flags += "ACCOUNTDISABLE" }
    if ($UAC -band 0x0008) { $flags += "HOMEDIR_REQUIRED" }
    if ($UAC -band 0x0010) { $flags += "LOCKOUT" }
    if ($UAC -band 0x0020) { $flags += "PASSWD_NOTREQD" }
    if ($UAC -band 0x0040) { $flags += "PASSWD_CANT_CHANGE" }
    if ($UAC -band 0x0080) { $flags += "ENCRYPTED_TEXT_PWD_ALLOWED" }
    if ($UAC -band 0x0100) { $flags += "TEMP_DUPLICATE_ACCOUNT" }
    if ($UAC -band 0x0200) { $flags += "NORMAL_ACCOUNT" }
    if ($UAC -band 0x0800) { $flags += "INTERDOMAIN_TRUST_ACCOUNT" }
    if ($UAC -band 0x1000) { $flags += "WORKSTATION_TRUST_ACCOUNT" }
    if ($UAC -band 0x2000) { $flags += "SERVER_TRUST_ACCOUNT" }
    if ($UAC -band 0x10000) { $flags += "DONT_EXPIRE_PASSWORD" }
    if ($UAC -band 0x20000) { $flags += "MNS_LOGON_ACCOUNT" }
    if ($UAC -band 0x40000) { $flags += "SMARTCARD_REQUIRED" }
    if ($UAC -band 0x80000) { $flags += "TRUSTED_FOR_DELEGATION" }
    if ($UAC -band 0x100000) { $flags += "NOT_DELEGATED" }
    if ($UAC -band 0x200000) { $flags += "USE_DES_KEY_ONLY" }
    if ($UAC -band 0x400000) { $flags += "DONT_REQ_PREAUTH" }
    if ($UAC -band 0x800000) { $flags += "PASSWORD_EXPIRED" }
    if ($UAC -band 0x1000000) { $flags += "TRUSTED_TO_AUTH_FOR_DELEGATION" }
    if ($UAC -band 0x04000000) { $flags += "PARTIAL_SECRETS_ACCOUNT" }
    
    return $flags -join ", "
}

# Event-Handler für Zurück-Button in Detailansicht
$btnBackToResults.Add_Click({
    $tabResults.SelectedIndex = 0
})

# Event-Handler für Aktualisieren-Button
$btnRefresh.Add_Click({
    Search-AD
})

# Event-Handler für Doppelklick auf ein Ergebnis
$dgResults.Add_MouseDoubleClick({
    if ($dgResults.SelectedItem) {
        ShowObjectDetails $dgResults.SelectedItem
    }
})

# Event-Handler für Filterung der Ergebnisse
$txtFilter.Add_TextChanged({
    if (-not $script:searchCompleted) { return }
    
    $filter = $txtFilter.Text
    if ([string]::IsNullOrWhiteSpace($filter)) {
        $dgResults.ItemsSource = $script:adSearchResults
        $txtResultCount.Text = " | " + $script:adSearchResults.Count + " Ergebnisse"
        return
    }
    
    $filteredResults = $script:adSearchResults | Where-Object {
        $_.Name -like "*$filter*" -or
        $_.DistinguishedName -like "*$filter*" -or
        $_.Description -like "*$filter*" -or
        $_.OperatingSystem -like "*$filter*"
    }
    
    $dgResults.ItemsSource = $filteredResults
    $txtResultCount.Text = " | " + $filteredResults.Count + " Ergebnisse (gefiltert)"
})

# Initialisierung der Anwendung
try {
    # Standardwerte für Filter setzen
    $cmbObjectType.SelectedIndex = 0
    $cmbStatus.SelectedIndex = 0
    $cmbLastLogon.SelectedIndex = 0
    $cmbResultLimit.SelectedIndex = 2
    
    # Initialisierung der Comboboxen
    Initialize-ComboBoxes
    
    # Hauptfenster anzeigen
    $window.ShowDialog()
    
    # Aufräumen beim Schließen
    if ($script:runspacePool) {
        $script:runspacePool.Close()
        $script:runspacePool.Dispose()
    }
} catch {
    [System.Windows.MessageBox]::Show("Fehler bei der Initialisierung: $_", "Fehler", 
        [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
}