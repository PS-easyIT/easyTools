﻿# ============================================ 
# Version:     0.1.0 
# Autor:       Andreas Hepp (unterstützt durch Cascade AI) 
# ============================================ 
 
#requires -Version 5.0 
#requires -Modules ActiveDirectory 
 
<# 
.SYNOPSIS 
    Erstellt Active Directory Dummy-Benutzer basierend auf einer CSV-Datei. 
.DESCRIPTION 
    Ein PowerShell-Skript mit WPF-GUI zum Einlesen einer CSV-Datei und zum Erstellen von AD-Benutzern. 
    Ermöglicht die Auswahl der zu verwendenden Attribute und bietet Platzhalter für weitere AD-Verwaltungsfunktionen.
    
    Features:
    - WPF-GUI für benutzerfreundliche Bedienung
    - CSV-Import mit konfigurierbaren Attributzuordnungen
    - Automatische UPN-Generierung
    - OU-Dropdown mit AD-Integration
    - Fortschrittsanzeige für längere Operationen
    - Umfassende Logging-Funktionalität
    - Sicherheitsfunktionen für kritische AD-Operationen 
.NOTES 
    Stellen Sie sicher, dass das ActiveDirectory PowerShell-Modul installiert ist und Sie über die erforderlichen Berechtigungen verfügen. 
#> 
 
#region Globale Variablen und Konfiguration 
Add-Type -AssemblyName PresentationFramework 
Add-Type -AssemblyName PresentationCore 
Add-Type -AssemblyName WindowsBase 
 
$Global:AppConfig = @{ 
    AppName                = "Easy AD Dummy User Creator" 
    ScriptVersion          = "0.1.0" 
    Author                 = "Andreas Hepp (Cascade AI)" 
    DefaultCsvPath         = ".\AD_DummyUserList.csv" # Standardpfad, kann im GUI geändert werden 
    DefaultPassword        = "P@sswOrd123!"            # Standardpasswort für neue User 
    EnablePasswordChange   = $true                     # Benutzer muss Passwort bei nächster Anmeldung ändern
    LogFilePath           = ".\easyADDummyUserCreator.log" # Pfad für Logdatei
} 
 
$Global:Window = $null 
$Global:GuiControls = @{} 
$Global:CsvData = $null 
$Global:CreatedUsers = [System.Collections.Generic.List[object]]::new() 
#endregion 
 
#region XAML Definition 
$XAML = @" 
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" 
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" 
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008" 
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" 
        mc:Ignorable="d" 
        Title="$($Global:AppConfig.AppName) v$($Global:AppConfig.ScriptVersion)" Height="800" Width="1200" 
        WindowStartupLocation="CenterScreen" FontFamily="Segoe UI"> 
    <Window.Resources> 
        <Style x:Key="ModernButton" TargetType="Button"> 
            <Setter Property="Background" Value="#0078D4" /> 
            <Setter Property="Foreground" Value="White" /> 
            <Setter Property="BorderThickness" Value="0" /> 
            <Setter Property="Padding" Value="10,5" /> 
            <Setter Property="Margin" Value="5" /> 
            <Setter Property="FontWeight" Value="SemiBold" /> 
            <Setter Property="Cursor" Value="Hand" /> 
            <Setter Property="MinWidth" Value="120" /> 
            <Setter Property="Template"> 
                <Setter.Value> 
                    <ControlTemplate TargetType="Button"> 
                        <Border x:Name="border" Background="{TemplateBinding Background}" CornerRadius="4"> 
                            <ContentPresenter Margin="{TemplateBinding Padding}" HorizontalAlignment="Center" VerticalAlignment="Center" /> 
                        </Border> 
                        <ControlTemplate.Triggers> 
                            <Trigger Property="IsMouseOver" Value="True"> 
                                <Setter TargetName="border" Property="Background" Value="#106EBE" /> 
                            </Trigger> 
                            <Trigger Property="IsPressed" Value="True"> 
                                <Setter TargetName="border" Property="Background" Value="#005A9E" /> 
                            </Trigger> 
                            <Trigger Property="IsEnabled" Value="False"> 
                                <Setter TargetName="border" Property="Background" Value="#A0A0A0" /> 
                                <Setter Property="Foreground" Value="#D0D0D0" /> 
                            </Trigger> 
                        </ControlTemplate.Triggers> 
                    </ControlTemplate> 
                </Setter.Value> 
            </Setter> 
        </Style> 
        <Style TargetType="Label"> 
            <Setter Property="Margin" Value="5,0,0,0"/> 
            <Setter Property="VerticalAlignment" Value="Center"/> 
        </Style> 
        <Style TargetType="TextBox"> 
            <Setter Property="Margin" Value="5"/> 
            <Setter Property="Padding" Value="3"/> 
            <Setter Property="VerticalAlignment" Value="Center"/> 
        </Style> 
        <Style TargetType="CheckBox"> 
            <Setter Property="Margin" Value="5"/> 
            <Setter Property="VerticalAlignment" Value="Center"/> 
        </Style> 
        <Style TargetType="ComboBox"> 
            <Setter Property="Margin" Value="5"/> 
            <Setter Property="Padding" Value="3"/> 
            <Setter Property="VerticalAlignment" Value="Center"/> 
        </Style> 
    </Window.Resources> 
    <Grid> 
        <Grid.RowDefinitions> 
            <RowDefinition Height="Auto"/> <!-- Header --> 
            <RowDefinition Height="*"/>    <!-- Main Content --> 
            <RowDefinition Height="Auto"/> <!-- Footer / Status --> 
        </Grid.RowDefinitions> 
 
        <!-- Header --> 
        <Border Grid.Row="0" Background="#FF1C323C" Padding="10"> 
            <TextBlock Text="$($Global:AppConfig.AppName)" Foreground="White" FontSize="18" VerticalAlignment="Center" HorizontalAlignment="Center"/> 
        </Border> 
 
        <!-- Main Content --> 
        <Grid Grid.Row="1" Margin="10"> 
            <Grid.ColumnDefinitions> 
                <ColumnDefinition Width="3*"/> 
                <ColumnDefinition Width="2*"/> 
            </Grid.ColumnDefinitions> 
 
            <!-- Left Panel: CSV Data and User Creation --> 
            <Grid Grid.Column="0" Margin="0,0,10,0"> 
                <Grid.RowDefinitions> 
                    <RowDefinition Height="Auto"/> <!-- CSV Load --> 
                    <RowDefinition Height="*"/>    <!-- CSV DataGrid --> 
                    <RowDefinition Height="Auto"/> <!-- User Creation Controls --> 
                </Grid.RowDefinitions> 
 
                <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,10"> 
                    <Label Content="CSV-Datei Pfad:"/> 
                    <TextBox x:Name="TextBoxCsvPath" Width="250" Text="$($Global:AppConfig.DefaultCsvPath)"/> 
                    <Button x:Name="ButtonBrowseCsv" Content="Durchsuchen..." Style="{StaticResource ModernButton}" Margin="5,5,0,5"/>
                    <Button x:Name="ButtonLoadCsv" Content="CSV laden" Style="{StaticResource ModernButton}"/> 
                </StackPanel> 
 
                <DataGrid x:Name="DataGridCsvContent" Grid.Row="1" Margin="0,5,0,5" IsReadOnly="True" AutoGenerateColumns="True"/> 
 
                <GroupBox Grid.Row="2" Header="AD Benutzer erstellen" Padding="10" Margin="0,10,0,0"> 
                    <Grid> 
                        <Grid.RowDefinitions> 
                            <RowDefinition Height="Auto"/> 
                            <RowDefinition Height="Auto"/> 
                            <RowDefinition Height="Auto"/> 
                            <RowDefinition Height="*"/> 
                            <RowDefinition Height="Auto"/> 
                            <RowDefinition Height="Auto"/> 
                        </Grid.RowDefinitions> 
                        <Grid.ColumnDefinitions> 
                            <ColumnDefinition Width="Auto"/> 
                            <ColumnDefinition Width="*"/> 
                        </Grid.ColumnDefinitions> 
 
                        <Label Grid.Row="0" Grid.Column="0" Content="Ziel OU (DN):"/> 
                        <ComboBox x:Name="ComboBoxTargetOU" Grid.Row="0" Grid.Column="1" IsEditable="True" Text="OU=DummyUsers,DC=example,DC=com" ToolTip="Wählen Sie eine OU aus der Liste oder geben Sie einen benutzerdefinierten DN ein"/>  
 
                        <Label Grid.Row="1" Grid.Column="0" Content="Anzahl User (leer=alle):"/> 
                        <TextBox x:Name="TextBoxNumUsersToCreate" Grid.Row="1" Grid.Column="1"/> 
 
                        <Label Grid.Row="2" Grid.Column="0" Content="Attribute für Neuanlage:" VerticalAlignment="Top" Margin="5,10,0,0"/> 
                        <ScrollViewer Grid.Row="2" Grid.Column="1" Grid.RowSpan="2" MaxHeight="150" VerticalScrollBarVisibility="Auto" Margin="5,5,0,5" BorderBrush="LightGray" BorderThickness="1"> 
                            <StackPanel x:Name="StackPanelAttributeSelection" Orientation="Vertical"/> 
                        </ScrollViewer> 
 
                        <Button x:Name="ButtonCreateADUsers" Grid.Row="4" Grid.ColumnSpan="2" Content="AD User erstellen" Style="{StaticResource ModernButton}" HorizontalAlignment="Right"/> 
                        
                        <!-- Fortschrittsanzeige --> 
                        <Grid Grid.Row="5" Grid.ColumnSpan="2" Margin="0,10,0,0" x:Name="GridProgressArea" Visibility="Collapsed"> 
                            <Grid.RowDefinitions> 
                                <RowDefinition Height="Auto"/> 
                                <RowDefinition Height="Auto"/> 
                            </Grid.RowDefinitions> 
                            <Label Grid.Row="0" Content="Fortschritt:" x:Name="LabelProgress"/> 
                            <ProgressBar Grid.Row="1" Height="20" x:Name="ProgressBarMain" Minimum="0" Maximum="100" Value="0"/> 
                        </Grid> 
                    </Grid> 
                </GroupBox> 
            </Grid> 
 
            <!-- Right Panel: Additional Actions --> 
            <GroupBox Grid.Column="1" Header="Weitere Aktionen (auf zuvor erstellte User)" Padding="10"> 
                <StackPanel Orientation="Vertical"> 
                    <StackPanel Orientation="Horizontal" Margin="0,0,0,10"> 
                        <Label Content="Anzahl User für Aktion:"/> 
                        <TextBox x:Name="TextBoxNumUsersForAction" Width="50"/> 
                    </StackPanel> 
 
                    <Button x:Name="ButtonFillAttributes" Content="Attribute befüllen" Style="{StaticResource ModernButton}" ToolTip="Aktualisiert Attribute der erstellten Benutzer basierend auf CSV-Daten"/> 
                    <Button x:Name="ButtonDisableUsers" Content="User deaktivieren" Style="{StaticResource ModernButton}" ToolTip="Deaktiviert die erstellten Benutzer"/>
                    <Button x:Name="ButtonDeleteUsers" Content="User löschen" Style="{StaticResource ModernButton}" ToolTip="Löscht die erstellten Benutzer (VORSICHT!)"/>
                    <Button x:Name="ButtonExportCreatedUsers" Content="Erstellte User exportieren" Style="{StaticResource ModernButton}" ToolTip="Exportiert Liste der erstellten Benutzer in CSV"/> 
                     
                    <GroupBox Header="Kritische Sicherheitseinstellungen" Margin="0,10,0,0" Padding="5"> 
                        <StackPanel> 
                            <CheckBox x:Name="CheckBoxPwdNotExpires" Content="Passwort läuft nicht ab"/> 
                            <CheckBox x:Name="CheckBoxRevPwd" Content="Umgekehrte Verschlüsselung erlauben"/> 
                            <Button x:Name="ButtonApplySecuritySettings" Content="Sicherheitseinstellungen anwenden" Style="{StaticResource ModernButton}" ToolTip="WARNUNG: Kritische Sicherheitseinstellungen!"/> 
                        </StackPanel> 
                    </GroupBox> 
 
                    <GroupBox Header="In kritische Gruppen zuweisen" Margin="0,10,0,0" Padding="5"> 
                        <StackPanel> 
                            <ComboBox x:Name="ComboBoxCriticalGroups"> 
                                <ComboBoxItem Content="Domain Admins"/> 
                                <ComboBoxItem Content="Enterprise Admins"/> 
                                <ComboBoxItem Content="Schema Admins"/> 
                            </ComboBox> 
                            <Button x:Name="ButtonAssignToGroups" Content="Gruppenzuweisung" Style="{StaticResource ModernButton}" ToolTip="WARNUNG: Höchste Privilegien!"/> 
                        </StackPanel> 
                    </GroupBox> 
                </StackPanel> 
            </GroupBox> 
        </Grid> 
 
        <!-- Footer / Status --> 
        <StatusBar Grid.Row="2"> 
            <StatusBarItem> 
                <TextBlock x:Name="TextBlockStatus" Text="Bereit."/> 
            </StatusBarItem> 
            <StatusBarItem HorizontalAlignment="Right"> 
                <TextBlock Text="$($Global:AppConfig.Author)"/> 
            </StatusBarItem> 
        </StatusBar> 
    </Grid> 
</Window> 
"@ 
#endregion 
 
#region Logging und Hilfsfunktionen 
function Write-Log { 
    param( 
        [string]$Message, 
        [string]$Level = "INFO" # INFO, WARN, ERROR, DEBUG, SUCCESS 
    ) 
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss" 
    $logEntry = "[$timestamp] [$Level] $Message" 
    Write-Host $logEntry # Für direkte Konsolenausgabe während der Entwicklung 
    
    if ($Global:Window -and $Global:GuiControls.TextBlockStatus) { 
        $Global:GuiControls.TextBlockStatus.Text = $Message 
    } 
    
    # In Datei loggen
    try {
        $logEntry | Out-File -FilePath $Global:AppConfig.LogFilePath -Append -Encoding UTF8
    } catch {
        # Fehler beim Logging nicht weiterwerfen, um Hauptfunktionalität nicht zu beeinträchtigen
        Write-Host "Fehler beim Schreiben in Logdatei: $($_.Exception.Message)" -ForegroundColor Yellow
    }
} 

function Normalize-SamAccountName {
    <#
    .SYNOPSIS
    Normalisiert einen SamAccountName für Active Directory-Kompatibilität.
    
    .DESCRIPTION
    Diese Funktion bereinigt und normalisiert einen SamAccountName, um sicherzustellen,
    dass er den AD-Anforderungen entspricht:
    - Keine Umlaute/Akzente (ä→a, ö→o, ü→u, ß→ss, é→e, etc.)
    - Keine Leerzeichen (ersetzt durch Unterstriche)
    - Keine Sonderzeichen (entfernt oder ersetzt)
    - Nur ASCII-Zeichen (a-z, A-Z, 0-9, -, _)
    - Maximale Länge von 20 Zeichen
    - Eindeutigkeit durch Counter bei Duplikaten
    
    .PARAMETER SamAccountName
    Der ursprüngliche SamAccountName aus der CSV-Datei
    
    .PARAMETER ExistingNames
    HashSet bestehender Namen zur Duplikatsvermeidung
    
    .EXAMPLE
    Normalize-SamAccountName -SamAccountName "müller.johann" -ExistingNames $existingNames
    # Gibt zurück: "mueller.johann"
    
    .EXAMPLE
    Normalize-SamAccountName -SamAccountName "mauch schlauchin" -ExistingNames $existingNames
    # Gibt zurück: "mauch_schlauchin"
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$SamAccountName,
        
        [Parameter(Mandatory=$false)]
        [System.Collections.Generic.HashSet[string]]$ExistingNames
    )
    
    if ([string]::IsNullOrWhiteSpace($SamAccountName)) {
        Write-Log "Normalize-SamAccountName: Leerer SamAccountName übergeben" -Level "ERROR"
        return "user" + (Get-Random -Minimum 1000 -Maximum 9999)
    }
    
    # 1. Zu Lowercase konvertieren
    $normalized = $SamAccountName.ToLower()
    
    # 2. Umlaute und Akzente ersetzen
    $replacements = @{
        'ä' = 'ae'; 'ö' = 'oe'; 'ü' = 'ue'; 'ß' = 'ss'
        'à' = 'a'; 'á' = 'a'; 'â' = 'a'; 'ã' = 'a'; 'å' = 'a'; 'æ' = 'ae'
        'è' = 'e'; 'é' = 'e'; 'ê' = 'e'; 'ë' = 'e'
        'ì' = 'i'; 'í' = 'i'; 'î' = 'i'; 'ï' = 'i'
        'ò' = 'o'; 'ó' = 'o'; 'ô' = 'o'; 'õ' = 'o'; 'ø' = 'o'
        'ù' = 'u'; 'ú' = 'u'; 'û' = 'u'
        'ý' = 'y'; 'ÿ' = 'y'
        'ñ' = 'n'; 'ç' = 'c'
        'ą' = 'a'; 'ć' = 'c'; 'ę' = 'e'; 'ł' = 'l'; 'ń' = 'n'; 'ś' = 's'; 'ź' = 'z'; 'ż' = 'z'
    }
    
    foreach ($char in $replacements.Keys) {
        $normalized = $normalized -replace [regex]::Escape($char), $replacements[$char]
    }
    
    # 3. Leerzeichen durch Unterstriche ersetzen
    $normalized = $normalized -replace '\s+', '_'
    
    # 4. Sonderzeichen entfernen oder ersetzen
    $normalized = $normalized -replace "'", ''  # Apostrophe entfernen
    $normalized = $normalized -replace '[^\w.-]', ''  # Nur Buchstaben, Zahlen, Punkte, Bindestriche, Unterstriche
    
    # 5. Mehrfache Punkte/Unterstriche/Bindestriche zusammenfassen
    $normalized = $normalized -replace '\.{2,}', '.'
    $normalized = $normalized -replace '_{2,}', '_'
    $normalized = $normalized -replace '-{2,}', '-'
    
    # 6. Führende/nachfolgende Sonderzeichen entfernen
    $normalized = $normalized.Trim('.', '_', '-')
    
    # 7. Nicht-ASCII-Zeichen durch Platzhalter ersetzen (für chinesische, japanische, koreanische, thai Zeichen)
    $normalized = $normalized -replace '[^\x00-\x7F]', 'x'
    
    # 8. Falls leer, Fallback verwenden
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        $normalized = "user" + (Get-Random -Minimum 1000 -Maximum 9999)
        Write-Log "SamAccountName war nach Normalisierung leer, verwende Fallback: $normalized" -Level "WARN"
    }
    
    # 9. Auf 20 Zeichen kürzen (AD-Limit)
    if ($normalized.Length -gt 20) {
        $normalized = $normalized.Substring(0, 17) + (Get-Random -Minimum 100 -Maximum 999)
        Write-Log "SamAccountName zu lang, gekürzt auf: $normalized" -Level "DEBUG"
    }
    
    # 10. Eindeutigkeit sicherstellen
    if ($ExistingNames) {
        $originalNormalized = $normalized
        $counter = 1
        
        while ($ExistingNames.Contains($normalized)) {
            $suffix = $counter.ToString()
            if (($originalNormalized.Length + $suffix.Length) -le 20) {
                $normalized = $originalNormalized + $suffix
            } else {
                $baseLength = 20 - $suffix.Length
                $normalized = $originalNormalized.Substring(0, $baseLength) + $suffix
            }
            $counter++
            
            # Sicherheitsbremse gegen Endlosschleife
            if ($counter -gt 999) {
                $normalized = "user" + (Get-Random -Minimum 1000 -Maximum 9999)
                break
            }
        }
        
        # Zur Liste hinzufügen
        $ExistingNames.Add($normalized) | Out-Null
        
        if ($normalized -ne $originalNormalized) {
            Write-Log "SamAccountName Duplikat vermieden: '$originalNormalized' → '$normalized'" -Level "DEBUG"
        }
    }
    
    Write-Log "SamAccountName normalisiert: '$SamAccountName' → '$normalized'" -Level "DEBUG"
    return $normalized
}
#endregion 
 
#region Event Handler und GUI Logik 
 
function Import-CsvData { 
    $csvPath = $Global:GuiControls.TextBoxCsvPath.Text 
    if (-not (Test-Path $csvPath -PathType Leaf)) { 
        Write-Log "CSV-Datei nicht gefunden: $csvPath" -Level "ERROR" 
        [System.Windows.MessageBox]::Show("Die angegebene CSV-Datei wurde nicht gefunden: `n$csvPath", "Fehler", "OK", "Error") 
        return 
    } 
    try { 
        $Global:CsvData = Import-Csv -Path $csvPath 
        $Global:GuiControls.DataGridCsvContent.ItemsSource = $Global:CsvData 
        Write-Log "CSV-Datei erfolgreich geladen: $csvPath ($($Global:CsvData.Count) Einträge)" -Level "SUCCESS" 
 
        # Attribute für Auswahl-Checkboxes populieren 
        $Global:GuiControls.StackPanelAttributeSelection.Children.Clear() 
        $csvHeaders = $Global:CsvData | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name 
 
        foreach ($attrInfo in $Global:SelectableADAttributes) { 
            if ($csvHeaders -contains $attrInfo.CsvHeader) { 
                $checkBox = New-Object System.Windows.Controls.CheckBox 
                $checkBox.Content = "$($attrInfo.Name) (aus $($attrInfo.CsvHeader))" 
                $checkBox.Tag = $attrInfo # Store the attribute info object 
                if ($attrInfo.Required -or @("GivenName", "Surname", "DisplayName", "UserPrincipalName") -contains $attrInfo.Name) { 
                    $checkBox.IsChecked = $true # Pre-select required and common attributes 
                } 
                $Global:GuiControls.StackPanelAttributeSelection.Children.Add($checkBox) | Out-Null 
            } 
        } 
    } catch { 
        Write-Log "Fehler beim Laden der CSV-Datei: $($_.Exception.Message)" -Level "ERROR" 
        [System.Windows.MessageBox]::Show("Fehler beim Laden der CSV-Datei:`n$($_.Exception.Message)", "Fehler", "OK", "Error") 
    } 
} 
 
function New-ADUsersFromCsv { 
    if (-not $Global:CsvData) { 
        Write-Log "Bitte zuerst eine CSV-Datei laden." -Level "WARN" 
        [System.Windows.MessageBox]::Show("Bitte laden Sie zuerst eine CSV-Datei.", "Hinweis", "OK", "Warning") 
        return 
    } 
 
    # AD-Verbindung prüfen
    if (-not (Test-ADConnection)) {
        return
    }

    $targetOU = $Global:GuiControls.ComboBoxTargetOU.Text 
    if (-not $targetOU) { 
        Write-Log "Bitte geben Sie eine Ziel-OU an." -Level "WARN" 
        [System.Windows.MessageBox]::Show("Bitte geben Sie eine gültige Ziel-OU (Distinguished Name) an.", "Hinweis", "OK", "Warning") 
        return 
    } 
    
    # OU validieren
    if (-not (Validate-TargetOU -TargetOU $targetOU)) {
        return
    } 
 
    $numUsersToCreateText = $Global:GuiControls.TextBoxNumUsersToCreate.Text 
    $numUsersToCreate = if ([string]::IsNullOrWhiteSpace($numUsersToCreateText)) { $Global:CsvData.Count } else { [int]$numUsersToCreateText } 
 
    if ($numUsersToCreate -le 0) { 
        Write-Log "Anzahl der zu erstellenden Benutzer muss größer als 0 sein." -Level "WARN" 
        return 
    } 
 
    $selectedAttributesCheckboxes = $Global:GuiControls.StackPanelAttributeSelection.Children | Where-Object {$_.IsChecked -eq $true} 
    if ($selectedAttributesCheckboxes.Count -eq 0) { 
        Write-Log "Bitte wählen Sie mindestens ein Attribut für die Benutzererstellung aus." -Level "WARN" 
        [System.Windows.MessageBox]::Show("Bitte wählen Sie mindestens ein Attribut für die Benutzererstellung aus.", "Hinweis", "OK", "Warning") 
        return 
    } 
 
    # Sicherstellen, dass SamAccountName ausgewählt ist, wenn es in der CSV existiert 
    $samAttrInfo = $Global:SelectableADAttributes | Where-Object {$_.Name -eq "SamAccountName"} 
    $samSelected = $selectedAttributesCheckboxes | ForEach-Object { $_.Tag } | Where-Object {$_.Name -eq "SamAccountName"} 
    if (($Global:CsvData[0].PSObject.Properties.Name -contains $samAttrInfo.CsvHeader) -and (-not $samSelected)) { 
         Write-Log "Das Attribut 'SamAccountName' muss ausgewählt sein, wenn es in der CSV existiert und für die Erstellung verwendet werden soll." -Level "WARN" 
        [System.Windows.MessageBox]::Show("Das Attribut 'SamAccountName' (aus CSV-Spalte '$($samAttrInfo.CsvHeader)') ist für die Benutzererstellung erforderlich und muss ausgewählt werden.", "Hinweis", "OK", "Warning") 
        return 
    } 

    # Fortschrittsanzeige aktivieren
    Show-Progress -Text "Starte Benutzererstellung..." -Value 0
 
    Write-Log "Starte Erstellung von bis zu $numUsersToCreate AD-Benutzern..." -Level "INFO" 
    $Global:CreatedUsers.Clear() 
    $createdCount = 0 
    $errorCount = 0 
    
    # HashSet für SamAccountName-Duplikate zur Laufzeit
    $usedSamAccountNames = New-Object System.Collections.Generic.HashSet[string] 
 
    for ($i = 0; $i -lt $Global:CsvData.Count; $i++) { 
        if ($createdCount -ge $numUsersToCreate) { break } 
 
        $csvRow = $Global:CsvData[$i] 
        $userParams = @{ 
            Path    = $targetOU 
            Enabled = $true
            ChangePasswordAtLogon = $Global:AppConfig.EnablePasswordChange # Wert aus globaler Konfiguration verwenden
            AccountPassword = (ConvertTo-SecureString $Global:AppConfig.DefaultPassword -AsPlainText -Force) 
        } 
 
        # Standard Name, falls nicht anders spezifiziert 
        $userParams.Name = if ($csvRow.PSObject.Properties["DisplayName"]) { $csvRow.DisplayName } elseif ($csvRow.PSObject.Properties["FirstName"] -and $csvRow.PSObject.Properties["LastName"]) { "$($csvRow.FirstName) $($csvRow.LastName)" } else { $csvRow.SamAccountName } 
         
        foreach ($checkBox in $selectedAttributesCheckboxes) { 
            $attrInfo = $checkBox.Tag 
            if ($csvRow.PSObject.Properties[$attrInfo.CsvHeader]) { 
                $value = $csvRow.($attrInfo.CsvHeader) 
                if (-not [string]::IsNullOrWhiteSpace($value)) { 
                    # Spezielle Behandlung für bestimmte Attribute 
                    if ($attrInfo.Name -eq "Country") { 
                        # Hier könnte eine Konvertierung zu 2-Buchstaben-Code erfolgen, falls nötig 
                        # Fürs Erste wird der Wert direkt übernommen 
                    } 
                    if ($attrInfo.Name -eq "proxyAddresses") { 
                        $value = $value -split ";" # Annahme: Semikolon-getrennte Liste 
                    } 
                    $userParams[$attrInfo.Name] = $value 
                } 
            } 
        } 
 
        # SamAccountName ist zwingend erforderlich 
        if (-not $userParams.ContainsKey("SamAccountName") -or [string]::IsNullOrWhiteSpace($userParams.SamAccountName)) { 
            Write-Log "Fehler: SamAccountName fehlt oder ist leer für Zeile $($i+1). Benutzer wird übersprungen." -Level "ERROR" 
            $errorCount++ 
            continue 
        }

        # SamAccountName für AD-Kompatibilität normalisieren
        $originalSamAccountName = $userParams.SamAccountName
        $userParams.SamAccountName = Normalize-SamAccountName -SamAccountName $originalSamAccountName -ExistingNames $usedSamAccountNames
        
        if ($originalSamAccountName -ne $userParams.SamAccountName) {
            Write-Log "SamAccountName normalisiert für Zeile $($i+1): '$originalSamAccountName' → '$($userParams.SamAccountName)'" -Level "INFO"
        }

        # Fortschritt aktualisieren
        $progressText = if ($originalSamAccountName -ne $userParams.SamAccountName) {
            "Erstelle Benutzer ($($originalSamAccountName) → $($userParams.SamAccountName))"
        } else {
            "Erstelle Benutzer ($($userParams.SamAccountName))"
        }
        Update-Progress -Current ($i + 1) -Total $numUsersToCreate -Text $progressText

        # Automatische UPN-Generierung, falls nicht bereits in CSV vorhanden
        if (-not $userParams.ContainsKey("UserPrincipalName") -or [string]::IsNullOrWhiteSpace($userParams.UserPrincipalName)) {
            $userParams.UserPrincipalName = Generate-UPN -SamAccountName $userParams.SamAccountName
            Write-Log "UPN automatisch generiert: $($userParams.UserPrincipalName)" -Level "DEBUG"
        }
 
        try { 
            Write-Log "Erstelle Benutzer: $($userParams.SamAccountName)" -Level "DEBUG" 
            # Write-Host ("DEBUG: New-ADUser @userParams: " + ($userParams | Format-List | Out-String)) # Zum Debuggen der Parameter 
            $newUser = New-ADUser @userParams -PassThru -ErrorAction Stop 
            Write-Log "Benutzer '$($newUser.SamAccountName)' erfolgreich erstellt." -Level "SUCCESS" 
            $Global:CreatedUsers.Add($newUser.SamAccountName) # Speichere nur SamAccountName für spätere Aktionen
            $createdCount++ 
        } catch { 
            Write-Log "Fehler beim Erstellen des Benutzers '$($userParams.SamAccountName)': $($_.Exception.Message)" -Level "ERROR" 
            $errorCount++ 
        } 
    }

    # Fortschrittsanzeige ausblenden
    Hide-Progress
    
    Write-Log "Benutzererstellung abgeschlossen. $createdCount erfolgreich erstellt, $errorCount Fehler." -Level "INFO" 
    [System.Windows.MessageBox]::Show("Benutzererstellung abgeschlossen.`nErfolgreich: $createdCount`nFehler: $errorCount", "Ergebnis", "OK", "Information") 
} 
 
 
function Set-ADUserAttributesFromCsv {
    Write-Log "Aktion 'Attribute befüllen' gestartet..."
    if (-not $Global:CreatedUsers -or $Global:CreatedUsers.Count -eq 0) {
        Write-Log "Keine Benutzer zuvor erstellt oder Liste ist leer. Aktion abgebrochen." -Level "WARN"
        [System.Windows.MessageBox]::Show("Es wurden noch keine Benutzer mit diesem Tool erstellt oder die Liste ist leer. Bitte erstellen Sie zuerst Benutzer.", "Hinweis", "OK", "Warning")
        return
    }

    if (-not $Global:CsvData) {
        Write-Log "Keine CSV-Daten geladen. Aktion abgebrochen." -Level "WARN"
        [System.Windows.MessageBox]::Show("Bitte laden Sie zuerst eine CSV-Datei, die die neuen Attributwerte enthält.", "Hinweis", "OK", "Warning")
        return
    }

    $numUsersForActionText = $Global:GuiControls.TextBoxNumUsersForAction.Text
    $numUsersToUpdate = 0
    if ([string]::IsNullOrWhiteSpace($numUsersForActionText)) {
        $numUsersToUpdate = $Global:CreatedUsers.Count # Default to all created users
        Write-Log "Keine Anzahl für Aktion angegeben, verwende alle $($Global:CreatedUsers.Count) erstellten Benutzer."
    } elseif ($numUsersForActionText -match '^\d+$') {
        $numUsersToUpdate = [int]$numUsersForActionText
    } else {
        Write-Log "Ungültige Eingabe für 'Anzahl Benutzer für Aktion': '$numUsersForActionText'. Muss eine Zahl sein." -Level "WARN"
        [System.Windows.MessageBox]::Show("Die Eingabe für 'Anzahl Benutzer für Aktion' ist ungültig. Bitte geben Sie eine Zahl ein.", "Fehler", "OK", "Error")
        return
    }

    if ($numUsersToUpdate -le 0) {
        Write-Log "Anzahl der zu aktualisierenden Benutzer muss größer als 0 sein." -Level "WARN"
        [System.Windows.MessageBox]::Show("Die Anzahl der Benutzer für diese Aktion muss größer als 0 sein.", "Hinweis", "OK", "Warning")
        return
    }
    
    if ($numUsersToUpdate -gt $Global:CreatedUsers.Count) {
        Write-Log "Angeforderte Anzahl ($numUsersToUpdate) übersteigt Anzahl erstellter Benutzer ($($Global:CreatedUsers.Count)). Reduziere auf Maximum." -Level "WARN"
        $numUsersToUpdate = $Global:CreatedUsers.Count
    }
    
    if ($numUsersToUpdate -gt $Global:CsvData.Count) {
        Write-Log "Angeforderte Anzahl ($numUsersToUpdate) übersteigt Anzahl der CSV-Einträge ($($Global:CsvData.Count)). Aktion nicht möglich für alle angeforderten Benutzer." -Level "ERROR"
        [System.Windows.MessageBox]::Show("Nicht genügend Einträge in der CSV-Datei vorhanden, um $numUsersToUpdate Benutzer zu aktualisieren (nur $($Global:CsvData.Count) CSV-Zeilen).", "Fehler", "OK", "Error")
        return
    }

    $selectedAttributesToSetInfo = @{}
    $anyAttributeSelected = $false
    foreach ($checkBox in $Global:GuiControls.StackPanelAttributeSelection.Children) {
        if ($checkBox.IsChecked -eq $true) {
            $attrInfo = $checkBox.Tag
            if ($attrInfo) {
                # SamAccountName ist der Identifikator, nicht zum Setzen via -Replace hier
                if ($attrInfo.Name -ne "SamAccountName") {
                     $selectedAttributesToSetInfo[$attrInfo.Name] = $attrInfo.CsvHeader
                     $anyAttributeSelected = $true
                }
            }
        }
    }

    if (-not $anyAttributeSelected) {
        Write-Log "Keine Attribute zum Befüllen ausgewählt." -Level "WARN"
        [System.Windows.MessageBox]::Show("Bitte wählen Sie mindestens ein Attribut aus, das befüllt werden soll.", "Hinweis", "OK", "Warning")
        return
    }
    Write-Log "Folgende Attribute wurden zum Befüllen ausgewählt: $($selectedAttributesToSetInfo.Keys -join ', ')"

    # Fortschrittsanzeige aktivieren
    Show-Progress -Text "Starte Attribut-Update..." -Value 0

    $usersProcessed = 0
    $usersFailed = 0

    for ($i = 0; $i -lt $numUsersToUpdate; $i++) {
        $samAccountName = $Global:CreatedUsers[$i]
        $csvRow = $Global:CsvData[$i]

        # Fortschritt aktualisieren
        Update-Progress -Current ($i + 1) -Total $numUsersToUpdate -Text "Aktualisiere Benutzerattribute"

        Write-Log "Verarbeite Benutzer '$samAccountName' (Datensatz $($i+1) aus CSV)..."
        $attributesForReplaceCmd = @{}

        foreach ($attrNameKey in $selectedAttributesToSetInfo.Keys) {
            $csvHeaderName = $selectedAttributesToSetInfo[$attrNameKey]
            if ($csvRow.PSObject.Properties[$csvHeaderName]) {
                $valueFromCsv = $csvRow.$($csvHeaderName)
                # Nur nicht-leere Werte aus der CSV verwenden, um Attribute zu setzen/überschreiben
                # Leere Werte in der CSV führen nicht zum Löschen des Attributs mit dieser Logik
                if (-not [string]::IsNullOrEmpty($valueFromCsv)) {
                    if ($attrNameKey -eq "proxyAddresses") {
                        $proxyAddressesArray = $valueFromCsv -split ';' | ForEach-Object {$_.Trim()} | Where-Object {$_}
                        $attributesForReplaceCmd[$attrNameKey] = $proxyAddressesArray
                    } else {
                        $attributesForReplaceCmd[$attrNameKey] = $valueFromCsv
                    }
                    Write-Log "  Setze '$attrNameKey' auf '$($attributesForReplaceCmd[$attrNameKey])' (aus CSV-Spalte '$csvHeaderName')"
                } else {
                    Write-Log "  Wert für '$attrNameKey' (aus CSV '$csvHeaderName') ist leer. Attribut wird für '$samAccountName' nicht geändert." -Level "INFO"
                }
            } else {
                Write-Log "  CSV-Spalte '$csvHeaderName' für Attribut '$attrNameKey' nicht im aktuellen CSV-Datensatz gefunden." -Level "WARN"
            }
        }

        if ($attributesForReplaceCmd.Count -gt 0) {
            try {
                Set-ADUser -Identity $samAccountName -Replace $attributesForReplaceCmd -ErrorAction Stop
                Write-Log "Benutzer '$samAccountName' erfolgreich aktualisiert." -Level "SUCCESS"
                $usersProcessed++
            } catch {
                Write-Log "Fehler beim Aktualisieren von Benutzer '$samAccountName': $($_.Exception.Message)" -Level "ERROR"
                $usersFailed++
            }
        } else {
            Write-Log "Keine gültigen Attributwerte zum Aktualisieren für Benutzer '$samAccountName' gefunden. Übersprungen." -Level "INFO"
        }
    }

    # Fortschrittsanzeige ausblenden
    Hide-Progress

    $summaryMessage = "Aktion 'Attribute befüllen' abgeschlossen. $usersProcessed Benutzer erfolgreich aktualisiert, $usersFailed Fehler."
    Write-Log $summaryMessage -Level "INFO"
    [System.Windows.MessageBox]::Show($summaryMessage, "Aktion abgeschlossen", "OK", "Information")
}

function Disable-CreatedADUsers {
    Write-Log "Aktion 'User deaktivieren' gestartet..."
    if (-not $Global:CreatedUsers -or $Global:CreatedUsers.Count -eq 0) {
        Write-Log "Keine Benutzer zuvor erstellt. Aktion abgebrochen." -Level "WARN"
        [System.Windows.MessageBox]::Show("Es wurden noch keine Benutzer mit diesem Tool erstellt.", "Hinweis", "OK", "Warning")
        return
    }

    $result = [System.Windows.MessageBox]::Show("Sind Sie sicher, dass Sie die erstellten Benutzer deaktivieren möchten?", "Bestätigung", "YesNo", "Question")
    if ($result -ne "Yes") {
        Write-Log "Benutzer-Deaktivierung abgebrochen." -Level "INFO"
        return
    }

    $numUsersForActionText = $Global:GuiControls.TextBoxNumUsersForAction.Text
    $numUsersToDisable = if ([string]::IsNullOrWhiteSpace($numUsersForActionText)) { $Global:CreatedUsers.Count } else { [int]$numUsersForActionText }

    $usersProcessed = 0
    $usersFailed = 0

    for ($i = 0; $i -lt [Math]::Min($numUsersToDisable, $Global:CreatedUsers.Count); $i++) {
        $samAccountName = $Global:CreatedUsers[$i]
        try {
            Disable-ADAccount -Identity $samAccountName -ErrorAction Stop
            Write-Log "Benutzer '$samAccountName' erfolgreich deaktiviert." -Level "SUCCESS"
            $usersProcessed++
        } catch {
            Write-Log "Fehler beim Deaktivieren von Benutzer '$samAccountName': $($_.Exception.Message)" -Level "ERROR"
            $usersFailed++
        }
    }

    $summaryMessage = "Aktion 'User deaktivieren' abgeschlossen. $usersProcessed Benutzer deaktiviert, $usersFailed Fehler."
    Write-Log $summaryMessage -Level "INFO"
    [System.Windows.MessageBox]::Show($summaryMessage, "Aktion abgeschlossen", "OK", "Information")
}

function Remove-CreatedADUsers {
    Write-Log "Aktion 'User löschen' gestartet..."
    if (-not $Global:CreatedUsers -or $Global:CreatedUsers.Count -eq 0) {
        Write-Log "Keine Benutzer zuvor erstellt. Aktion abgebrochen." -Level "WARN"
        [System.Windows.MessageBox]::Show("Es wurden noch keine Benutzer mit diesem Tool erstellt.", "Hinweis", "OK", "Warning")
        return
    }

    $result = [System.Windows.MessageBox]::Show("WARNUNG: Diese Aktion löscht die erstellten Benutzer UNWIDERRUFLICH!`n`nSind Sie sicher?", "VORSICHT - Benutzer löschen", "YesNo", "Warning")
    if ($result -ne "Yes") {
        Write-Log "Benutzer-Löschung abgebrochen." -Level "INFO"
        return
    }

    $numUsersForActionText = $Global:GuiControls.TextBoxNumUsersForAction.Text
    $numUsersToDelete = if ([string]::IsNullOrWhiteSpace($numUsersForActionText)) { $Global:CreatedUsers.Count } else { [int]$numUsersForActionText }

    $usersProcessed = 0
    $usersFailed = 0

    for ($i = 0; $i -lt [Math]::Min($numUsersToDelete, $Global:CreatedUsers.Count); $i++) {
        $samAccountName = $Global:CreatedUsers[$i]
        try {
            Remove-ADUser -Identity $samAccountName -Confirm:$false -ErrorAction Stop
            Write-Log "Benutzer '$samAccountName' erfolgreich gelöscht." -Level "SUCCESS"
            $usersProcessed++
        } catch {
            Write-Log "Fehler beim Löschen von Benutzer '$samAccountName': $($_.Exception.Message)" -Level "ERROR"
            $usersFailed++
        }
    }

    # Erfolgreich gelöschte Benutzer aus der Liste entfernen
    for ($i = $usersProcessed - 1; $i -ge 0; $i--) {
        $Global:CreatedUsers.RemoveAt($i)
    }

    $summaryMessage = "Aktion 'User löschen' abgeschlossen. $usersProcessed Benutzer gelöscht, $usersFailed Fehler."
    Write-Log $summaryMessage -Level "INFO"
    [System.Windows.MessageBox]::Show($summaryMessage, "Aktion abgeschlossen", "OK", "Information")
}

function Export-CreatedUsers {
    Write-Log "Export der erstellten Benutzer gestartet..."
    if (-not $Global:CreatedUsers -or $Global:CreatedUsers.Count -eq 0) {
        Write-Log "Keine Benutzer zum Exportieren vorhanden." -Level "WARN"
        [System.Windows.MessageBox]::Show("Es wurden noch keine Benutzer erstellt oder alle wurden bereits gelöscht.", "Hinweis", "OK", "Warning")
        return
    }

    try {
        $exportPath = ".\Created_Users_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        $exportData = @()
        
        foreach ($samAccountName in $Global:CreatedUsers) {
            try {
                $user = Get-ADUser -Identity $samAccountName -Properties * -ErrorAction Stop
                $exportData += [PSCustomObject]@{
                    SamAccountName = $user.SamAccountName
                    DisplayName = $user.DisplayName
                    GivenName = $user.GivenName
                    Surname = $user.Surname
                    UserPrincipalName = $user.UserPrincipalName
                    EmailAddress = $user.EmailAddress
                    Enabled = $user.Enabled
                    Created = $user.Created
                    DistinguishedName = $user.DistinguishedName
                }
            } catch {
                Write-Log "Fehler beim Abrufen der Details für Benutzer '$samAccountName': $($_.Exception.Message)" -Level "WARN"
                $exportData += [PSCustomObject]@{
                    SamAccountName = $samAccountName
                    DisplayName = "Fehler beim Abrufen"
                    GivenName = ""
                    Surname = ""
                    UserPrincipalName = ""
                    EmailAddress = ""
                    Enabled = "Unbekannt"
                    Created = ""
                    DistinguishedName = ""
                }
            }
        }

        $exportData | Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8
        Write-Log "Export erfolgreich: $exportPath" -Level "SUCCESS"
        [System.Windows.MessageBox]::Show("Benutzerliste erfolgreich exportiert nach:`n$exportPath", "Export abgeschlossen", "OK", "Information")
    } catch {
        Write-Log "Fehler beim Export: $($_.Exception.Message)" -Level "ERROR"
        [System.Windows.MessageBox]::Show("Fehler beim Exportieren der Benutzerliste:`n$($_.Exception.Message)", "Fehler", "OK", "Error")
    }
}

function Show-CsvBrowser {
    try {
        Add-Type -AssemblyName System.Windows.Forms
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
        $openFileDialog.Title = "CSV-Datei auswählen"
        $openFileDialog.InitialDirectory = Split-Path $Global:GuiControls.TextBoxCsvPath.Text -Parent

        if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $Global:GuiControls.TextBoxCsvPath.Text = $openFileDialog.FileName
            Write-Log "CSV-Pfad ausgewählt: $($openFileDialog.FileName)" -Level "INFO"
        }
    } catch {
        Write-Log "Fehler beim Öffnen des Datei-Browsers: $($_.Exception.Message)" -Level "ERROR"
        [System.Windows.MessageBox]::Show("Fehler beim Öffnen des Datei-Browsers. Bitte geben Sie den Pfad manuell ein.", "Fehler", "OK", "Error")
    }
}

function Apply-SecuritySettings {
    Write-Log "Anwenden kritischer Sicherheitseinstellungen gestartet..."
    if (-not $Global:CreatedUsers -or $Global:CreatedUsers.Count -eq 0) {
        Write-Log "Keine Benutzer zuvor erstellt. Aktion abgebrochen." -Level "WARN"
        [System.Windows.MessageBox]::Show("Es wurden noch keine Benutzer mit diesem Tool erstellt.", "Hinweis", "OK", "Warning")
        return
    }

    $result = [System.Windows.MessageBox]::Show("WARNUNG: Sie sind dabei, kritische Sicherheitseinstellungen zu ändern!`n`nDies kann Ihre Domäne gefährden. Sind Sie sicher?", "KRITISCHE SICHERHEITSWARNUNG", "YesNo", "Warning")
    if ($result -ne "Yes") {
        Write-Log "Anwendung kritischer Sicherheitseinstellungen abgebrochen." -Level "INFO"
        return
    }

    $numUsersForActionText = $Global:GuiControls.TextBoxNumUsersForAction.Text
    $numUsersToModify = if ([string]::IsNullOrWhiteSpace($numUsersForActionText)) { $Global:CreatedUsers.Count } else { [int]$numUsersForActionText }

    $usersProcessed = 0
    $usersFailed = 0

    $pwdNotExpires = $Global:GuiControls.CheckBoxPwdNotExpires.IsChecked
    $revPwd = $Global:GuiControls.CheckBoxRevPwd.IsChecked

    for ($i = 0; $i -lt [Math]::Min($numUsersToModify, $Global:CreatedUsers.Count); $i++) {
        $samAccountName = $Global:CreatedUsers[$i]
        try {
            $params = @{}
            
            if ($pwdNotExpires) {
                $params.PasswordNeverExpires = $true
                Write-Log "Setze für '$samAccountName': Passwort läuft nie ab" -Level "WARN"
            }
            
            if ($revPwd) {
                $params.AllowReversiblePasswordEncryption = $true
                Write-Log "Setze für '$samAccountName': Umgekehrte Verschlüsselung aktiviert" -Level "WARN"
            }

            if ($params.Count -gt 0) {
                Set-ADUser -Identity $samAccountName @params -ErrorAction Stop
                Write-Log "Sicherheitseinstellungen für Benutzer '$samAccountName' erfolgreich angewendet." -Level "SUCCESS"
                $usersProcessed++
            }
        } catch {
            Write-Log "Fehler beim Anwenden der Sicherheitseinstellungen für Benutzer '$samAccountName': $($_.Exception.Message)" -Level "ERROR"
            $usersFailed++
        }
    }

    $summaryMessage = "Anwendung kritischer Sicherheitseinstellungen abgeschlossen. $usersProcessed Benutzer bearbeitet, $usersFailed Fehler."
    Write-Log $summaryMessage -Level "INFO"
    [System.Windows.MessageBox]::Show($summaryMessage, "Aktion abgeschlossen", "OK", "Information")
}

function Add-ToGroup {
    Write-Log "Gruppenzuweisung gestartet..."
    if (-not $Global:CreatedUsers -or $Global:CreatedUsers.Count -eq 0) {
        Write-Log "Keine Benutzer zuvor erstellt. Aktion abgebrochen." -Level "WARN"
        [System.Windows.MessageBox]::Show("Es wurden noch keine Benutzer mit diesem Tool erstellt.", "Hinweis", "OK", "Warning")
        return
    }

    $selectedGroup = $Global:GuiControls.ComboBoxCriticalGroups.SelectedItem
    if (-not $selectedGroup) {
        Write-Log "Keine Gruppe ausgewählt." -Level "WARN"
        [System.Windows.MessageBox]::Show("Bitte wählen Sie eine Gruppe aus der Liste aus.", "Hinweis", "OK", "Warning")
        return
    }

    $groupName = $selectedGroup.Content
    $result = [System.Windows.MessageBox]::Show("EXTREM GEFAEHRLICH: Sie sind dabei, Benutzer zur Gruppe '$groupName' hinzuzufuegen!`n`nDies gibt diesen Benutzern hoechste Privilegien. Sind Sie ABSOLUT sicher?", "KRITISCHE SICHERHEITSWARNUNG", "YesNo", "Warning")
    if ($result -ne "Yes") {
        Write-Log "Gruppenzuweisung zu '$groupName' abgebrochen." -Level "INFO"
        return
    }

    $numUsersForActionText = $Global:GuiControls.TextBoxNumUsersForAction.Text
    $numUsersToAdd = if ([string]::IsNullOrWhiteSpace($numUsersForActionText)) { $Global:CreatedUsers.Count } else { [int]$numUsersForActionText }

    $usersProcessed = 0
    $usersFailed = 0

    for ($i = 0; $i -lt [Math]::Min($numUsersToAdd, $Global:CreatedUsers.Count); $i++) {
        $samAccountName = $Global:CreatedUsers[$i]
        try {
            Add-ADGroupMember -Identity $groupName -Members $samAccountName -ErrorAction Stop
            Write-Log "Benutzer '$samAccountName' erfolgreich zur Gruppe '$groupName' hinzugefügt." -Level "SUCCESS"
            $usersProcessed++
        } catch {
            Write-Log "Fehler beim Hinzufügen von Benutzer '$samAccountName' zur Gruppe '$groupName': $($_.Exception.Message)" -Level "ERROR"
            $usersFailed++
        }
    }

    $summaryMessage = "Gruppenzuweisung zu '$groupName' abgeschlossen. $usersProcessed Benutzer hinzugefügt, $usersFailed Fehler."
    Write-Log $summaryMessage -Level "INFO"
    [System.Windows.MessageBox]::Show($summaryMessage, "Aktion abgeschlossen", "OK", "Information")
}

function Test-ADConnection {
    try {
        Write-Log "Prüfe Active Directory Verbindung..."
        $domain = Get-ADDomain -ErrorAction Stop
        Write-Log "AD-Verbindung erfolgreich. Domäne: $($domain.Name)" -Level "SUCCESS"
        return $true
    } catch {
        Write-Log "Fehler bei AD-Verbindung: $($_.Exception.Message)" -Level "ERROR"
        [System.Windows.MessageBox]::Show("Active Directory Verbindung fehlgeschlagen:`n$($_.Exception.Message)`n`nBitte stellen Sie sicher, dass:`n- Das ActiveDirectory-Modul installiert ist`n- Sie mit einer Domäne verbunden sind`n- Sie über ausreichende Berechtigungen verfügen", "AD-Verbindungsfehler", "OK", "Error")
        return $false
    }
}

function Validate-TargetOU {
    param([string]$TargetOU)
    
    if (-not $TargetOU) {
        return $false
    }
    
    try {
        $ou = Get-ADOrganizationalUnit -Identity $TargetOU -ErrorAction Stop
        Write-Log "Ziel-OU validiert: $($ou.DistinguishedName)" -Level "SUCCESS"
        return $true
    } catch {
        Write-Log "Ziel-OU nicht gefunden oder ungültig: $TargetOU - $($_.Exception.Message)" -Level "ERROR"
        [System.Windows.MessageBox]::Show("Die angegebene Ziel-OU ist ungültig oder nicht gefunden:`n$TargetOU`n`nFehler: $($_.Exception.Message)", "OU-Validierungsfehler", "OK", "Error")
        return $false
    }
}

function Register-EventHandlers { 
    $Global:GuiControls.ButtonBrowseCsv.Add_Click({ Show-CsvBrowser })
    $Global:GuiControls.ButtonLoadCsv.Add_Click({ Import-CsvData }) 
    $Global:GuiControls.ButtonCreateADUsers.Add_Click({ New-ADUsersFromCsv }) 
    $Global:GuiControls.ButtonFillAttributes.Add_Click({ Set-ADUserAttributesFromCsv }) 
    $Global:GuiControls.ButtonDisableUsers.Add_Click({ Disable-CreatedADUsers })
    $Global:GuiControls.ButtonDeleteUsers.Add_Click({ Remove-CreatedADUsers })
    $Global:GuiControls.ButtonExportCreatedUsers.Add_Click({ Export-CreatedUsers })
    $Global:GuiControls.ButtonApplySecuritySettings.Add_Click({ Apply-SecuritySettings })
    $Global:GuiControls.ButtonAssignToGroups.Add_Click({ Add-ToGroup })
 
    Write-Log "Event Handler registriert." 
} 
 
#endregion 
 
#region Initialisierung und Start 
 
function Initialize-Application { 
    Write-Log "Initialisiere Anwendung..." 
    try { 
        # XAML korrekt laden 
        $stringReader = New-Object System.IO.StringReader -ArgumentList $XAML 
        $xmlTextReader = [System.Xml.XmlTextReader]::new($stringReader) 
        $Global:Window = [Windows.Markup.XamlReader]::Load($xmlTextReader) 
 
        if (-not $Global:Window) { 
            Throw "Kritischer Fehler: Das XAML-Fenster konnte nicht geladen werden." 
        } 
        Write-Log "XAML-Fenster erfolgreich geladen." 
 
        Get-GuiControls # Ruft die korrigierte Funktion auf, die $Global:Window verwendet 
        Register-EventHandlers 
 
        # Standard CSV laden, falls vorhanden, ansonsten Beispiel erstellen
        $Global:GuiControls.TextBoxCsvPath.Text = $Global:AppConfig.DefaultCsvPath
        if (Test-Path $Global:AppConfig.DefaultCsvPath -PathType Leaf) { 
            Import-CsvData 
        } else { 
            Write-Log "Standard CSV-Datei nicht gefunden: $($Global:AppConfig.DefaultCsvPath)" -Level "WARN" 
            Create-SampleCsv
        } 

        # AD-Verbindung testen (nicht blockierend)
        if (-not (Test-ADConnection)) {
            Write-Log "AD-Verbindung fehlgeschlagen - Funktion eingeschränkt verfügbar." -Level "WARN"
        } else {
            # OU-Dropdown initialisieren wenn AD-Verbindung besteht
            Initialize-OUDropdown
        }

        Write-Log "Anwendung erfolgreich initialisiert." -Level "SUCCESS" 
    } catch { 
        Write-Log "Fehler während der Anwendungsinitialisierung: $($_.Exception.ToString())" -Level "ERROR" 
        # Den Fehler weiterwerfen, damit er vom äußeren try-catch in Start-Application behandelt wird 
        throw $_ 
    } 
}

function Create-SampleCsv {
    try {
        $sampleData = @(
            [PSCustomObject]@{
                SamAccountName = "test.user1"
                FirstName = "Test"
                LastName = "User1"
                DisplayName = "Test User 1"
                Email = "test.user1@example.com"
                Department = "IT"
                Title = "Test Engineer"
                Company = "Example Corp"
                UserPrincipalName = "test.user1@example.com"
            },
            [PSCustomObject]@{
                SamAccountName = "test.user2"
                FirstName = "Test"
                LastName = "User2"
                DisplayName = "Test User 2"
                Email = "test.user2@example.com"
                Department = "HR"
                Title = "Test Manager"
                Company = "Example Corp"
                UserPrincipalName = "test.user2@example.com"
            }
        )
        
        $sampleData | Export-Csv -Path $Global:AppConfig.DefaultCsvPath -NoTypeInformation -Encoding UTF8
        Write-Log "Beispiel-CSV erstellt: $($Global:AppConfig.DefaultCsvPath)" -Level "INFO"
        Import-CsvData
    } catch {
        Write-Log "Fehler beim Erstellen der Beispiel-CSV: $($_.Exception.Message)" -Level "ERROR"
    }
}

#region GUI Verbesserungen - OU Dropdown, Progress, UPN Generation

function Get-ADOrganizationalUnits {
    <#
    .SYNOPSIS
    Ruft alle verfügbaren Organizational Units aus Active Directory ab
    #>
    try {
        Write-Log "Lade OUs aus Active Directory..."
        $ous = Get-ADOrganizationalUnit -Filter * | Sort-Object DistinguishedName
        Write-Log "$(($ous | Measure-Object).Count) OUs gefunden" -Level "SUCCESS"
        return $ous
    } catch {
        Write-Log "Fehler beim Laden der OUs: $($_.Exception.Message)" -Level "ERROR"
        return @()
    }
}

function Initialize-OUDropdown {
    <#
    .SYNOPSIS
    Befüllt die OU-ComboBox mit verfügbaren OUs
    #>
    try {
        $comboBox = $Global:GuiControls.ComboBoxTargetOU
        $comboBox.Items.Clear()
        
        # Standard-OU hinzufügen
        $comboBox.Items.Add("OU=DummyUsers,DC=example,DC=com")
        
        # AD OUs laden
        $ous = Get-ADOrganizationalUnits
        foreach ($ou in $ous) {
            if ($ou.DistinguishedName -notlike "*CN=*") {  # Nur echte OUs, keine Container
                $comboBox.Items.Add($ou.DistinguishedName)
            }
        }
        
        # Ersten Eintrag auswählen
        if ($comboBox.Items.Count -gt 0) {
            $comboBox.SelectedIndex = 0
        }
        
        Write-Log "OU-Dropdown initialisiert mit $(($comboBox.Items | Measure-Object).Count) Einträgen" -Level "SUCCESS"
    } catch {
        Write-Log "Fehler beim Initialisieren des OU-Dropdowns: $($_.Exception.Message)" -Level "ERROR"
    }
}

function Generate-UPN {
    <#
    .SYNOPSIS
    Generiert automatisch User Principal Names basierend auf SamAccountName
    #>
    param(
        [string]$SamAccountName,
        [string]$DefaultDomain = $null
    )
    
    try {
        if (-not $DefaultDomain) {
            # Versuche die Standard-Domain aus AD zu ermitteln
            $domain = Get-ADDomain -ErrorAction SilentlyContinue
            if ($domain) {
                $DefaultDomain = $domain.DNSRoot
            } else {
                $DefaultDomain = "example.com"
            }
        }
        
        $upn = "$SamAccountName@$DefaultDomain"
        Write-Log "UPN generiert: $upn" -Level "DEBUG"
        return $upn
    } catch {
        Write-Log "Fehler bei UPN-Generierung für '$SamAccountName': $($_.Exception.Message)" -Level "ERROR"
        return "$SamAccountName@example.com"
    }
}

function Show-Progress {
    <#
    .SYNOPSIS
    Zeigt die Fortschrittsanzeige an
    #>
    param(
        [string]$Text = "Vorgang läuft...",
        [int]$Value = 0
    )
    
    try {
        $Global:GuiControls.GridProgressArea.Visibility = [System.Windows.Visibility]::Visible
        $Global:GuiControls.LabelProgress.Content = $Text
        $Global:GuiControls.ProgressBarMain.Value = $Value
        
        # UI aktualisieren
        $Global:Window.Dispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Render, [action]{})
    } catch {
        Write-Log "Fehler beim Anzeigen des Fortschritts: $($_.Exception.Message)" -Level "ERROR"
    }
}

function Hide-Progress {
    <#
    .SYNOPSIS
    Versteckt die Fortschrittsanzeige
    #>
    try {
        $Global:GuiControls.GridProgressArea.Visibility = [System.Windows.Visibility]::Collapsed
        $Global:GuiControls.ProgressBarMain.Value = 0
    } catch {
        Write-Log "Fehler beim Verstecken des Fortschritts: $($_.Exception.Message)" -Level "ERROR"
    }
}

function Update-Progress {
    <#
    .SYNOPSIS
    Aktualisiert den Fortschrittsbalken
    #>
    param(
        [int]$Current,
        [int]$Total,
        [string]$Text = $null
    )
    
    try {
        $percentage = [math]::Round(($Current / $Total) * 100, 0)
        
        if ($Text) {
            Show-Progress -Text "$Text ($Current von $Total)" -Value $percentage
        } else {
            $Global:GuiControls.ProgressBarMain.Value = $percentage
        }
    } catch {
        Write-Log "Fehler beim Aktualisieren des Fortschritts: $($_.Exception.Message)" -Level "ERROR"
    }
}

#endregion

function Start-Application {
    try { 
        Initialize-Application 
        Write-Log "Starte GUI..." 
        $null = $Global:Window.ShowDialog() 
    } catch { 
        Write-Log "Kritischer Fehler in der Anwendung: $($_.Exception.Message)" -Level "ERROR" 
        [System.Windows.MessageBox]::Show("Ein kritischer Fehler ist aufgetreten:`n$($_.Exception.Message)", "Kritischer Fehler", "OK", "Error") 
    } finally { 
        Write-Log "Anwendung beendet." 
    } 
} 
 
#endregion 
 
# ============================================================================ 
# SCRIPT ENTRY POINT 
# ============================================================================ 
Start-Application

# SIG # Begin signature block
# MIIbywYJKoZIhvcNAQcCoIIbvDCCG7gCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCjW5aX0fQVADE0
# B11BUDmGqC1Odu8gH5zGvbxL/rPnwaCCFhcwggMQMIIB+KADAgECAhB3jzsyX9Cg
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
# DQEJBDEiBCA8YQLHMSQztE/jtPsQDa0pcpxSWXSrewWqnl13lVdH4TANBgkqhkiG
# 9w0BAQEFAASCAQA+aOsctTVBHDZrxPMapAaivFHoEw34Kk2b6sxQj4JPYT2NfRJp
# OuwOvyEy7g9XtF89JMLmnViHKFlWe64O2QgabddvOb0FW2K6OuylEw5gfTU4wwBC
# PgN7zPZsskNOHuNdGxGGIOWHczFdqgCDtyQUEX6/qwYpio3lxthtPXeUKROH8Xcj
# 5UgivfTvQ+HIAug62gIb4pAbT9pNrpPYscbQ2tb7XRIRZi3Qk5GBKtKpfIBHbAub
# DH5jrOAkpLNUk8ioTVZWOS5c764Yg1U+ZhUysYtGJdMJ1ssGfrcE5mB5Z7/nEkLs
# 5n0DK2ivv7QNH9YdjeEJFcqH3jCXC593Bt0doYIDIDCCAxwGCSqGSIb3DQEJBjGC
# Aw0wggMJAgEBMHcwYzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJ
# bmMuMTswOQYDVQQDEzJEaWdpQ2VydCBUcnVzdGVkIEc0IFJTQTQwOTYgU0hBMjU2
# IFRpbWVTdGFtcGluZyBDQQIQC65mvFq6f5WHxvnpBOMzBDANBglghkgBZQMEAgEF
# AKBpMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTI1
# MDcwNTEwMTIwOFowLwYJKoZIhvcNAQkEMSIEIBcUdaS7ElidwSM6xQgPfAVpJau/
# AfbP2UrFnlsasg63MA0GCSqGSIb3DQEBAQUABIICAIXCNnKZ09Ao8JBKlulyYann
# lBGuuXG+8a/f8iuoTn9EFnN7sI1enkioayoncMIpfPr2NVcKPFFztTI0hxc26vo/
# xyoXt9UaITJCBEPDNKuWcG3q3y5bCkfCC3e8uYBcX9fKfZwAnG+hiaqu+WZoOvDo
# VkuS2+K7Jdf/SeAHS5yK1hze+6CQiqkM/2obK1BueCnIl1xpzHSnyr5j+UaVa20f
# RcS2L1bUkIRGu7t7aiQ/Qv+RuZon/tC9CwUQgbTaG8b+F+M0dV0h+HESgsWrIK9G
# zFtjXhG2LqKKCqED7s3SSNZ47rkBd53aaD4dHUZ5JaXBMjPrGPi4OO3AoAbVEAeu
# uTqLRTbZH+RNLewk3hpvUd2S9DEetH+SASBskBmaDNFfAfgrdS2kXhCRrpJWGDWQ
# /Gy0Jb2lNG+cDij3KMHK79DGfNyScB+RCE2QOyfzLA2Lob+VgO9qbW+CqIfxRX4O
# CThE979fexXgIo7VsorB/L0TvqJTbNwTrsIdSP3D5I/fzr36orIp3eG0v9Bav5+R
# 0HGtMhkYb71lcREqGZ+QP+++vndnK0T6kEI7NcxqU4yiLOXRioM2OSVlsbG51b44
# baA7ts63uLDpOCpZICTdVs0hYmQgh77hk0ZUFR5OHkQEPzlX706+BWev7E56EStn
# o2zWFnqdKbumL5qPcM2S
# SIG # End signature block
