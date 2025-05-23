# EasyADFinder - Advanced Active Directory Query Tool

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.DirectoryServices

# Konverterklassen für die im XAML verwendeten Konverter definieren
Add-Type -ReferencedAssemblies @(
    'System.dll',
    'System.Core.dll',
    'WindowsBase.dll', 
    'PresentationCore.dll', 
    'PresentationFramework.dll'
) -TypeDefinition @"
using System;
using System.Windows;
using System.Windows.Data;
using System.Globalization;

namespace EasyADFinder.Converters
{
    public class UserVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value != null && value.ToString().ToLower() == "user")
                return Visibility.Visible;
            return Visibility.Collapsed;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    public class ComputerVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value != null && value.ToString().ToLower() == "computer")
                return Visibility.Visible;
            return Visibility.Collapsed;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
"@

# XAML für die WPF GUI anpassen mit korrektem Namespace für die Konverter
[xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:local="clr-namespace:System;assembly=mscorlib"
    xmlns:converters="clr-namespace:EasyADFinder.Converters;assembly=EasyADFinder.Converters"
    Title="AD-Finder" Height="700" Width="1050"
    WindowStartupLocation="CenterScreen" Background="#f0f0f0">
    
    <Window.Resources>
        <!-- Konverter definieren -->
        <converters:UserVisibilityConverter x:Key="UserVisibilityConverter" />
        <converters:ComputerVisibilityConverter x:Key="ComputerVisibilityConverter" />
        
        <!-- Erweiterte Stile für Windows 11 Look & Feel -->
        <Style x:Key="ModernButton" TargetType="Button">
            <Setter Property="Background" Value="#0078d4"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="Padding" Value="15,8"/>
            <Setter Property="BorderThickness" Value="0"/>
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
                    <Setter Property="Background" Value="#106ebe"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        
        <Style x:Key="IconButton" TargetType="Button" BasedOn="{StaticResource ModernButton}">
            <Setter Property="Padding" Value="8"/>
            <Setter Property="Width" Value="34"/>
            <Setter Property="Height" Value="34"/>
            <Setter Property="Margin" Value="5,0"/>
            <Setter Property="Background" Value="#f0f0f0"/>
            <Setter Property="Foreground" Value="#0078d4"/>
            <Setter Property="BorderBrush" Value="#d1d1d1"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#e5e5e5"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        
        <Style x:Key="ModernComboBox" TargetType="ComboBox">
            <Setter Property="Padding" Value="8,5"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="BorderBrush" Value="#d1d1d1"/>
            <Setter Property="Background" Value="White"/>
        </Style>
        
        <Style x:Key="ModernTextBox" TargetType="TextBox">
            <Setter Property="Padding" Value="8,5"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="BorderBrush" Value="#d1d1d1"/>
        </Style>
        
        <Style x:Key="ModernCheckBox" TargetType="CheckBox">
            <Setter Property="Margin" Value="0,5"/>
        </Style>

        <Style x:Key="ModernTabItem" TargetType="TabItem">
            <Setter Property="Padding" Value="12,5"/>
            <Setter Property="HeaderTemplate">
                <Setter.Value>
                    <DataTemplate>
                        <TextBlock Text="{Binding}" FontWeight="SemiBold"/>
                    </DataTemplate>
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
        <Border Grid.Row="0" Background="White" CornerRadius="8" Padding="15" Margin="0,0,0,15">
            <Grid>
                <TabControl BorderThickness="0" Background="Transparent">
                    <TabItem Header="Hauptfilter" Style="{StaticResource ModernTabItem}">
                        <Grid Margin="0,10,0,0">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            
                            <StackPanel Grid.Column="0">
                                <TextBlock Text="AD-Finder" FontSize="20" FontWeight="SemiBold" Margin="0,0,0,15"/>
                                
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
                                    <TextBlock Grid.Row="0" Grid.Column="0" Text="Objekttyp:" VerticalAlignment="Center" Margin="0,0,10,10"/>
                                    <ComboBox x:Name="cmbObjectType" Grid.Row="0" Grid.Column="1" Style="{StaticResource ModernComboBox}" Margin="0,0,20,10">
                                        <ComboBoxItem Content="Benutzer" IsSelected="True"/>
                                        <ComboBoxItem Content="Computer"/>
                                        <ComboBoxItem Content="Gruppen"/>
                                        <ComboBoxItem Content="OUs"/>
                                    </ComboBox>
                                    
                                    <TextBlock Grid.Row="0" Grid.Column="2" Text="Betriebssystem:" VerticalAlignment="Center" Margin="0,0,10,10"/>
                                    <TextBox x:Name="txtOS" Grid.Row="0" Grid.Column="3" Style="{StaticResource ModernTextBox}" Margin="0,0,0,10"/>
                                    
                                    <!-- Zeile 2 -->
                                    <TextBlock Grid.Row="1" Grid.Column="0" Text="Name enthält:" VerticalAlignment="Center" Margin="0,0,10,10"/>
                                    <TextBox x:Name="txtName" Grid.Row="1" Grid.Column="1" Style="{StaticResource ModernTextBox}" Margin="0,0,20,10"/>
                                    
                                    <TextBlock Grid.Row="1" Grid.Column="2" Text="Letztes Logon:" VerticalAlignment="Center" Margin="0,0,10,10"/>
                                    <Grid Grid.Row="1" Grid.Column="3" Margin="0,0,0,10">
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="Auto"/>
                                        </Grid.ColumnDefinitions>
                                        <TextBox x:Name="txtLastLogon" Grid.Column="0" Style="{StaticResource ModernTextBox}"/>
                                        <TextBlock Grid.Column="1" Text=" Tage" VerticalAlignment="Center" Margin="5,0,0,0"/>
                                    </Grid>
                                    
                                    <!-- Zeile 3 -->
                                    <TextBlock Grid.Row="2" Grid.Column="0" Text="Beschreibung:" VerticalAlignment="Center" Margin="0,0,10,10"/>
                                    <TextBox x:Name="txtDescription" Grid.Row="2" Grid.Column="1" Style="{StaticResource ModernTextBox}" Margin="0,0,20,10"/>
                                    
                                    <TextBlock Grid.Row="2" Grid.Column="2" Text="Status:" VerticalAlignment="Center" Margin="0,0,10,10"/>
                                    <ComboBox x:Name="cmbStatus" Grid.Row="2" Grid.Column="3" Style="{StaticResource ModernComboBox}" Margin="0,0,0,10">
                                        <ComboBoxItem Content="Alle" IsSelected="True"/>
                                        <ComboBoxItem Content="Aktiviert"/>
                                        <ComboBoxItem Content="Deaktiviert"/>
                                    </ComboBox>
                                    
                                    <!-- Zeile 4 -->
                                    <TextBlock Grid.Row="3" Grid.Column="0" Text="OU enthält:" VerticalAlignment="Center" Margin="0,0,10,10"/>
                                    <TextBox x:Name="txtOU" Grid.Row="3" Grid.Column="1" Style="{StaticResource ModernTextBox}" Margin="0,0,20,10"/>
                                    
                                    <TextBlock Grid.Row="3" Grid.Column="2" Text="Ergebnislimit:" VerticalAlignment="Center" Margin="0,0,10,0"/>
                                    <Grid Grid.Row="3" Grid.Column="3">
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="Auto"/>
                                            <ColumnDefinition Width="Auto"/>
                                        </Grid.ColumnDefinitions>
                                        <TextBox x:Name="txtResultLimit" Grid.Column="0" Style="{StaticResource ModernTextBox}" Text="1000"/>
                                        <CheckBox x:Name="chkExtendedAttributes" Grid.Column="1" Style="{StaticResource ModernCheckBox}" 
                                                Content="Erweiterte Attribute" VerticalAlignment="Center" Margin="10,0,0,0"/>
                                    </Grid>
                                </Grid>
                            </StackPanel>
                            
                            <Button x:Name="btnSearch" Grid.Column="1" Content="Suchen" Style="{StaticResource ModernButton}" 
                                    VerticalAlignment="Bottom" Width="120" Margin="20,0,0,0"/>
                        </Grid>
                    </TabItem>
                    <TabItem Header="Erweiterte Filter" Style="{StaticResource ModernTabItem}">
                        <Grid Margin="0,10,0,0">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            
                            <StackPanel Grid.Column="0">
                                <TextBlock Text="Erweiterte Suchparameter" FontSize="18" FontWeight="SemiBold" Margin="0,0,0,15"/>
                                
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
                                    
                                    <!-- Zeile 1 - Benutzer -->
                                    <TextBlock Grid.Row="0" Grid.Column="0" Text="Abteilung:" VerticalAlignment="Center" Margin="0,0,10,10"/>
                                    <TextBox x:Name="txtDepartment" Grid.Row="0" Grid.Column="1" Style="{StaticResource ModernTextBox}" Margin="0,0,20,10"/>
                                    
                                    <TextBlock Grid.Row="0" Grid.Column="2" Text="E-Mail enthält:" VerticalAlignment="Center" Margin="0,0,10,10"/>
                                    <TextBox x:Name="txtEmail" Grid.Row="0" Grid.Column="3" Style="{StaticResource ModernTextBox}" Margin="0,0,0,10"/>
                                    
                                    <!-- Zeile 2 - Benutzer/Computer -->
                                    <TextBlock Grid.Row="1" Grid.Column="0" Text="Erstellt nach:" VerticalAlignment="Center" Margin="0,0,10,10"/>
                                    <DatePicker x:Name="dtCreatedAfter" Grid.Row="1" Grid.Column="1" Margin="0,0,20,10"/>
                                    
                                    <TextBlock Grid.Row="1" Grid.Column="2" Text="Erstellt vor:" VerticalAlignment="Center" Margin="0,0,10,10"/>
                                    <DatePicker x:Name="dtCreatedBefore" Grid.Row="1" Grid.Column="3" Margin="0,0,0,10"/>
                                    
                                    <!-- Zeile 3 - Gruppen -->
                                    <TextBlock Grid.Row="2" Grid.Column="0" Text="Gruppenkategorie:" VerticalAlignment="Center" Margin="0,0,10,10"/>
                                    <ComboBox x:Name="cmbGroupCategory" Grid.Row="2" Grid.Column="1" Style="{StaticResource ModernComboBox}" Margin="0,0,20,10">
                                        <ComboBoxItem Content="Alle" IsSelected="True"/>
                                        <ComboBoxItem Content="Sicherheit"/>
                                        <ComboBoxItem Content="Distribution"/>
                                    </ComboBox>
                                    
                                    <TextBlock Grid.Row="2" Grid.Column="2" Text="Gruppenbereich:" VerticalAlignment="Center" Margin="0,0,10,10"/>
                                    <ComboBox x:Name="cmbGroupScope" Grid.Row="2" Grid.Column="3" Style="{StaticResource ModernComboBox}" Margin="0,0,0,10">
                                        <ComboBoxItem Content="Alle" IsSelected="True"/>
                                        <ComboBoxItem Content="Global"/>
                                        <ComboBoxItem Content="Domain Local"/>
                                        <ComboBoxItem Content="Universal"/>
                                    </ComboBox>
                                    
                                    <!-- Zeile 4 - LDAP-Filter -->
                                    <TextBlock Grid.Row="3" Grid.Column="0" Text="LDAP-Filter:" VerticalAlignment="Center" Margin="0,0,10,10"/>
                                    <TextBox x:Name="txtLdapFilter" Grid.Row="3" Grid.Column="1" Grid.ColumnSpan="3" Style="{StaticResource ModernTextBox}" Margin="0,0,0,10"/>
                                </Grid>
                            </StackPanel>
                            
                            <Button x:Name="btnAdvancedSearch" Grid.Column="1" Content="Erweitert Suchen" Style="{StaticResource ModernButton}" 
                                    VerticalAlignment="Bottom" Width="120" Margin="20,0,0,0"/>
                        </Grid>
                    </TabItem>
                    <TabItem Header="Gespeicherte Suchen" Style="{StaticResource ModernTabItem}">
                        <Grid Margin="0,10,0,0">
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>
                            
                            <TextBlock Grid.Row="0" Text="Gespeicherte Suchanfragen" FontSize="18" FontWeight="SemiBold" Margin="0,0,0,15"/>
                            
                            <ListBox x:Name="lstSavedSearches" Grid.Row="1" Margin="0,0,0,10"/>
                            
                            <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right">
                                <Button x:Name="btnLoadSearch" Content="Laden" Style="{StaticResource ModernButton}" Width="100" Margin="0,0,10,0"/>
                                <Button x:Name="btnSaveSearch" Content="Aktuelle Suche speichern" Style="{StaticResource ModernButton}" Width="180" Margin="0,0,10,0"/>
                                <Button x:Name="btnDeleteSearch" Content="Löschen" Style="{StaticResource ModernButton}" Width="100" Background="#d83b01"/>
                            </StackPanel>
                        </Grid>
                    </TabItem>
                </TabControl>
            </Grid>
        </Border>
        
        <!-- Inhalt - DataGrid für Suchergebnisse -->
        <Grid Grid.Row="1" Margin="0,0,0,15">
            <TabControl x:Name="tabResults" BorderThickness="0">
                <TabItem Header="Ergebnisse" Style="{StaticResource ModernTabItem}">
                    <Border Background="White" CornerRadius="8" Padding="15">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>
                            
                            <Grid Grid.Row="0" Margin="0,0,0,10">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <TextBlock Grid.Column="0" Text="Suchergebnisse" FontSize="16" FontWeight="SemiBold"/>
                                <TextBox x:Name="txtFilter" Grid.Column="1" Style="{StaticResource ModernTextBox}" Margin="15,0" 
                                        ToolTip="In Ergebnissen filtern"/>
                                <StackPanel Grid.Column="2" Orientation="Horizontal">
                                    <Button x:Name="btnRefresh" Style="{StaticResource IconButton}" ToolTip="Aktualisieren" Content="↻"/>
                                    <Button x:Name="btnSelectColumns" Style="{StaticResource IconButton}" ToolTip="Spalten auswählen" Content="≡"/>
                                </StackPanel>
                            </Grid>
                            
                            <DataGrid x:Name="dgResults" Grid.Row="1" AutoGenerateColumns="False" IsReadOnly="True" 
                                    AlternatingRowBackground="#f9f9f9" VerticalScrollBarVisibility="Auto"
                                    BorderThickness="1" BorderBrush="#e0e0e0">
                                <DataGrid.ContextMenu>
                                    <ContextMenu>
                                        <MenuItem x:Name="menuDetails" Header="Details anzeigen"/>
                                        <MenuItem x:Name="menuReset" Header="Passwort zurücksetzen"/> 
                                        <MenuItem x:Name="menuDisable" Header="Deaktivieren"/>
                                        <MenuItem x:Name="menuEnable" Header="Aktivieren"/>
                                        <Separator/>
                                        <MenuItem x:Name="menuMove" Header="In andere OU verschieben"/>
                                        <MenuItem x:Name="menuAddToGroup" Header="Zu Gruppe hinzufügen"/>
                                        <Separator/>
                                        <MenuItem x:Name="menuRemote" Header="Remote-Verbindung herstellen"/>
                                        <Separator/>
                                        <MenuItem x:Name="menuExportSelected" Header="Ausgewählte exportieren"/>
                                    </ContextMenu>
                                </DataGrid.ContextMenu>
                                <DataGrid.Columns>
                                    <DataGridTextColumn Header="Name" Binding="{Binding Name}" Width="150"/>
                                    <DataGridTextColumn Header="Distinguished Name" Binding="{Binding DistinguishedName}" Width="250"/>
                                    <DataGridCheckBoxColumn Header="Aktiviert" Binding="{Binding Enabled}" Width="70"/>
                                    <DataGridTextColumn Header="Beschreibung" Binding="{Binding Description}" Width="200"/>
                                    <DataGridTextColumn Header="Betriebssystem" Binding="{Binding OperatingSystem}" Width="150"/>
                                    <DataGridTextColumn Header="Letzter Login" Binding="{Binding LastLogonDate}" Width="150"/>
                                </DataGrid.Columns>
                            </DataGrid>
                        </Grid>
                    </Border>
                </TabItem>
                <TabItem Header="Detailansicht" Style="{StaticResource ModernTabItem}">
                    <Border Background="White" CornerRadius="8" Padding="15">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>
                            
                            <Grid Grid.Row="0" Margin="0,0,0,10">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <Button x:Name="btnBackToResults" Grid.Column="0" Style="{StaticResource IconButton}" Content="←" ToolTip="Zurück zur Ergebnisliste"/>
                                <TextBlock Grid.Column="1" x:Name="txtSelectedObjectName" Text="Objektdetails" FontSize="16" FontWeight="SemiBold" VerticalAlignment="Center" Margin="10,0"/>
                                <StackPanel Grid.Column="2" Orientation="Horizontal">
                                    <Button x:Name="btnEditObject" Style="{StaticResource IconButton}" ToolTip="Objekt bearbeiten" Content="✎"/>
                                    <Button x:Name="btnShowInADUC" Style="{StaticResource IconButton}" ToolTip="In AD-Benutzer und -Computer anzeigen" Content="🔍"/>
                                </StackPanel>
                            </Grid>
                            
                            <TabControl Grid.Row="1" BorderThickness="0">
                                <TabItem Header="Allgemeine Info" Style="{StaticResource ModernTabItem}">
                                    <ScrollViewer VerticalScrollBarVisibility="Auto">
                                        <Grid x:Name="gridGeneralInfo" Margin="10">
                                            <!-- Wird dynamisch mit den Objekteigenschaften gefüllt -->
                                        </Grid>
                                    </ScrollViewer>
                                </TabItem>
                                <TabItem Header="Attribute" Style="{StaticResource ModernTabItem}">
                                    <Grid>
                                        <Grid.RowDefinitions>
                                            <RowDefinition Height="Auto"/>
                                            <RowDefinition Height="*"/>
                                        </Grid.RowDefinitions>
                                        <TextBox Grid.Row="0" x:Name="txtAttributeFilter" Style="{StaticResource ModernTextBox}" 
                                                Margin="0,0,0,10" ToolTip="Nach Attributen filtern"/>
                                        <DataGrid Grid.Row="1" x:Name="dgAttributes" AutoGenerateColumns="False" IsReadOnly="True" 
                                                AlternatingRowBackground="#f9f9f9" VerticalScrollBarVisibility="Auto"
                                                BorderThickness="1" BorderBrush="#e0e0e0">
                                            <DataGrid.Columns>
                                                <DataGridTextColumn Header="Attribut" Binding="{Binding Name}" Width="200"/>
                                                <DataGridTextColumn Header="Wert" Binding="{Binding Value}" Width="*"/>
                                            </DataGrid.Columns>
                                        </DataGrid>
                                    </Grid>
                                </TabItem>
                                <TabItem Header="Mitgliedschaften" Style="{StaticResource ModernTabItem}" x:Name="tabMembership">
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
                                            <TextBlock Grid.Row="0" Text="Mitglied von" FontWeight="SemiBold" Margin="0,0,0,5"/>
                                            <ListBox Grid.Row="1" x:Name="lbMemberOf"/>
                                        </Grid>
                                        
                                        <GridSplitter Grid.Row="1" Height="5" HorizontalAlignment="Stretch" Background="#e0e0e0" Margin="0,10"/>
                                        
                                        <Grid Grid.Row="2">
                                            <Grid.RowDefinitions>
                                                <RowDefinition Height="Auto"/>
                                                <RowDefinition Height="*"/>
                                            </Grid.RowDefinitions>
                                            <TextBlock Grid.Row="0" Text="Mitglieder" FontWeight="SemiBold" Margin="0,0,0,5"/>
                                            <ListBox Grid.Row="1" x:Name="lbMembers"/>
                                        </Grid>
                                    </Grid>
                                </TabItem>
                            </TabControl>
                        </Grid>
                    </Border>
                </TabItem>
            </TabControl>
        </Grid>
        
        <!-- Footer mit Buttons -->
        <Border Grid.Row="2" Background="White" CornerRadius="8" Padding="15">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                
                <TextBlock x:Name="lblStatus" Text="Bereit" VerticalAlignment="Center"/>
                <Button x:Name="btnExport" Grid.Column="1" Content="CSV Export" Style="{StaticResource ModernButton}" 
                        Margin="0,0,15,0" Width="120"/>
                <Button x:Name="btnClose" Grid.Column="2" Content="Schließen" Style="{StaticResource ModernButton}" 
                        Background="#d83b01" Width="120"/>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

# XAML laden
$reader = New-Object System.Xml.XmlNodeReader $xaml

# Das Namespace-Problem beheben
$NWin = [Xml.XmlNamespaceManager]::new($xaml.NameTable)
$NWin.AddNamespace("x", "http://schemas.microsoft.com/winfx/2006/xaml")
$NWin.AddNamespace("local", "clr-namespace:System;assembly=mscorlib")

# WPF-Fenster laden
try {
    $window = [Windows.Markup.XamlReader]::Load($reader)
    
    # UI-Elemente abrufen
    $btnSearch = $window.FindName("btnSearch")
    $btnExport = $window.FindName("btnExport")
    $btnClose = $window.FindName("btnClose")
    $cmbObjectType = $window.FindName("cmbObjectType")
    $txtOS = $window.FindName("txtOS")
    $txtName = $window.FindName("txtName")
    $txtLastLogon = $window.FindName("txtLastLogon")
    $txtDescription = $window.FindName("txtDescription")
    $cmbStatus = $window.FindName("cmbStatus")
    $txtOU = $window.FindName("txtOU")
    $chkExtendedAttributes = $window.FindName("chkExtendedAttributes")
    $dgResults = $window.FindName("dgResults")
    $lblStatus = $window.FindName("lblStatus")
    $menuReset = $window.FindName("menuReset")
    $menuRemote = $window.FindName("menuRemote")

    # Kontextmenüelemente basierend auf dem Objekttyp anzeigen/ausblenden
    $dgResults.Add_SelectionChanged({
        $selectedItem = $dgResults.SelectedItem
        if ($selectedItem -ne $null) {
            # Menüelemente basierend auf dem Objekttyp anzeigen/ausblenden
            if ($selectedItem.ObjectClass -eq "user") {
                $menuReset.Visibility = "Visible"
                $menuRemote.Visibility = "Collapsed"
            } 
            elseif ($selectedItem.ObjectClass -eq "computer") {
                $menuReset.Visibility = "Collapsed"
                $menuRemote.Visibility = "Visible"
            }
            else {
                $menuReset.Visibility = "Collapsed"
                $menuRemote.Visibility = "Collapsed"
            }
        }
    })

    # Globale Variable für Suchergebnisse
    $global:searchResults = @()

    # Funktion zum Suchen von AD-Objekten
    function Search-ADObjects {
        $objectType = $cmbObjectType.SelectedItem.Content
        $os = $txtOS.Text
        $name = $txtName.Text
        $lastLogon = $txtLastLogon.Text 
        $description = $txtDescription.Text
        $status = $cmbStatus.SelectedItem.Content
        $ou = $txtOU.Text
        $extendedAttributes = $chkExtendedAttributes.IsChecked

        $lblStatus.Text = "Suche läuft..."
        
        # Properties für die Suche festlegen
        $properties = @("Name", "DistinguishedName", "Enabled", "Description", "OperatingSystem", "LastLogonDate", "Created", "ObjectClass")
        
        if ($extendedAttributes) {
            if ($objectType -eq "Benutzer") {
                $properties += @("Title", "Department", "Office", "EmailAddress", "PasswordLastSet", "BadPwdCount")
            } elseif ($objectType -eq "Computer") {
                $properties += @("IPv4Address", "DNSHostName", "ServicePrincipalNames", "ManagedBy")
            }
        }
        
        # Filter erstellen
        $filter = "*"
        if ($name) { $filter = "Name -like '*$name*'" }
        
        try {
            # AD-Objekte abrufen
            if ($objectType -eq "Benutzer") {
                $adObjects = Get-ADUser -Filter $filter -Properties $properties
            } else {
                $adObjects = Get-ADComputer -Filter $filter -Properties $properties
            }
            
            # Weitere Filterbedingungen anwenden
            $filteredObjects = $adObjects | Where-Object {
                $result = $true
                
                if ($os) { $result = $result -and ($_.OperatingSystem -like "*$os*") }
                if ($description) { $result = $result -and ($_.Description -like "*$description*") }
                if ($ou) { $result = $result -and ($_.DistinguishedName -like "*$ou*") }
                if ($lastLogon -and [int]::TryParse($lastLogon, [ref]$null)) { 
                    $result = $result -and ($_.LastLogonDate -gt (Get-Date).AddDays(-[int]$lastLogon)) 
                }
                if ($status -ne "Alle") { 
                    $result = $result -and (($status -eq "Aktiviert" -and $_.Enabled -eq $true) -or 
                                          ($status -eq "Deaktiviert" -and $_.Enabled -eq $false))
                }
                
                return $result
            }
            
            # Ergebnisse speichern
            $global:searchResults = $filteredObjects
            
            # DataGrid befüllen
            $dgResults.ItemsSource = $filteredObjects
            
            $lblStatus.Text = "$($filteredObjects.Count) $objectType gefunden"
        }
        catch {
            $lblStatus.Text = "Fehler: $($_.Exception.Message)"
        }
    }

    # Funktion zum Exportieren als CSV
    function Export-ResultsToCSV {
        if ($global:searchResults.Count -eq 0) {
            $lblStatus.Text = "Keine Daten zum Exportieren vorhanden!"
            return
        }
        
        $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveDialog.Filter = "CSV Dateien (*.csv)|*.csv"
        $saveDialog.Title = "CSV-Datei speichern"
        $saveDialog.FileName = "AD_Export_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        
        if ($saveDialog.ShowDialog() -eq 'OK') {
            try {
                $global:searchResults | Export-Csv -Path $saveDialog.FileName -NoTypeInformation -Encoding UTF8
                $lblStatus.Text = "Daten wurden nach $($saveDialog.FileName) exportiert."
            }
            catch {
                $lblStatus.Text = "Fehler beim Export: $($_.Exception.Message)"
            }
        }
    }

    # Event-Handler für Buttons
    $btnSearch.Add_Click({ Search-ADObjects })
    $btnExport.Add_Click({ Export-ResultsToCSV })
    $btnClose.Add_Click({ $window.Close() })

    # GUI anzeigen
    [void]$window.ShowDialog()
}
catch {
    [System.Windows.MessageBox]::Show("Fehler beim Laden des UI: $($_.Exception.Message)", "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
}
