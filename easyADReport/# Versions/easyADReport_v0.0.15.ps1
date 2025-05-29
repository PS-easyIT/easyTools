[xml]$Global:XAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        Title="easyADReport v0.0.3" Height="800" Width="1400"
        WindowStartupLocation="CenterScreen" ResizeMode="CanResizeWithGrip">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Obere Leiste: Filter und Attributauswahl -->
        <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="10">
            <GroupBox Header="Filter" Margin="5">
                <StackPanel Orientation="Horizontal">
                    <Label Content="Filter Attribut:" VerticalAlignment="Center"/>
                    <ComboBox x:Name="ComboBoxFilterAttribute" Width="150" Margin="5,0" VerticalAlignment="Center"/>
                    <Label Content="Filter Wert:" VerticalAlignment="Center"/>
                    <TextBox x:Name="TextBoxFilterValue" Width="200" Margin="5,0" VerticalAlignment="Center"/>
                </StackPanel>
            </GroupBox>
            <GroupBox Header="Zu exportierende Attribute" Margin="5">
                <!-- Hier kommt die Attributauswahl, z.B. eine ListBox mit CheckBoxen -->
                <ListBox x:Name="ListBoxSelectAttributes" Width="300" Height="80" SelectionMode="Multiple">
                    <!-- Beispielattribute, werden später dynamisch gefüllt -->
                    <ListBoxItem Content="DisplayName"/>
                    <ListBoxItem Content="SamAccountName"/>
                    <ListBoxItem Content="GivenName"/>
                    <ListBoxItem Content="Surname"/>
                    <ListBoxItem Content="mail"/>
                    <ListBoxItem Content="Department"/>
                    <ListBoxItem Content="Title"/>
                    <ListBoxItem Content="Enabled"/>
                </ListBox>
            </GroupBox>
            <Button x:Name="ButtonQueryAD" Content="Abfragen" Width="100" Margin="10,5" VerticalAlignment="Center"/>
        </StackPanel>

        <!-- Mittlerer Bereich: Quick Bereich und Ergebnisse -->
        <Grid Grid.Row="1" Margin="10">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="200"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

            <GroupBox Header="Quick Reports" Grid.Column="0" Margin="5">
                <StackPanel>
                    <Button x:Name="ButtonQuickAllUsers" Content="Alle Benutzer" Margin="5"/>
                    <Button x:Name="ButtonQuickDisabledUsers" Content="Deaktivierte Benutzer" Margin="5"/>
                    <Button x:Name="ButtonQuickLockedUsers" Content="Gesperrte Benutzer" Margin="5"/>
                    <Button x:Name="ButtonQuickGroups" Content="Alle Gruppen" Margin="5"/>
                    <!-- Weitere Quick Reports hier -->
                </StackPanel>
            </GroupBox>

            <GroupBox Header="Ergebnisse" Grid.Column="1" Margin="5">
                <DataGrid x:Name="DataGridResults" AutoGenerateColumns="True" IsReadOnly="True"/>
            </GroupBox>
        </Grid>

        <!-- Untere Leiste: Export und Status -->
        <StatusBar Grid.Row="2">
            <StatusBarItem>
                <TextBlock x:Name="TextBlockStatus" Text="Bereit"/>
            </StatusBarItem>
            <StatusBarItem HorizontalAlignment="Right">
                <StackPanel Orientation="Horizontal">
                    <Button x:Name="ButtonExportCSV" Content="Export CSV" Width="100" Margin="5"/>
                    <Button x:Name="ButtonExportHTML" Content="Export HTML" Width="100" Margin="5"/>
                </StackPanel>
            </StatusBarItem>
        </StatusBar>
    </Grid>
</Window>
"@

# Assembly für WPF laden
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms # Für SaveFileDialog

# --- Log-Funktion für konsistente Fehlerausgabe ---
Function Write-ADReportLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [ValidateSet('Info', 'Warning', 'Error')]
        [string]$Type = 'Info',
        
        [Parameter(Mandatory=$false)]
        [switch]$GUI,
        
        [Parameter(Mandatory=$false)]
        [switch]$Terminal
    )
    
    # Standardmäßig sowohl GUI als auch Terminal, wenn nicht explizit angegeben
    if (-not $GUI -and -not $Terminal) {
        $GUI = $true
        $Terminal = $true
    }
    
    # Ausgabe in der GUI
    if ($GUI -and $Global:TextBlockStatus) {
        $Global:TextBlockStatus.Text = $Message
    }
    
    # Ausgabe im Terminal
    if ($Terminal) {
        switch ($Type) {
            'Info'    { Write-Host $Message }
            'Warning' { Write-Warning $Message }
            'Error'   { Write-Error $Message }
        }
    }
}

# --- Funktion zum Abrufen von AD-Daten ---
Function Get-ADReportData {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false)]
        [string]$FilterAttribute,
        [Parameter(Mandatory=$false)]
        [string]$FilterValue,
        [Parameter(Mandatory=$true)]
        [System.Collections.IList]$SelectedAttributes,
        [Parameter(Mandatory=$false)]
        [string]$CustomFilter
    )

    # Überprüfen, ob das AD-Modul verfügbar ist
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        Write-ADReportLog -Message "Fehler: Active Directory Modul nicht gefunden." -Type Error
        return $null
    }

    try {
        # Konvertiere SelectedAttributes zu String-Array und filtere leere/null Werte
        $PropertiesToLoad = @(
            $SelectedAttributes | ForEach-Object {
                if ($null -ne $_) {
                    # Wenn es sich um ListBoxItem handelt, Content verwenden
                    if ($_ -is [System.Windows.Controls.ListBoxItem]) {
                        $_.Content.ToString()
                    } else {
                        $_.ToString()
                    }
                }
            } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        )

        # Sicherstellen, dass immer einige Basiseigenschaften geladen werden
        if ($FilterAttribute -and (-not ($PropertiesToLoad -contains $FilterAttribute))) {
            $PropertiesToLoad += $FilterAttribute
        }
        if ($PropertiesToLoad.Count -eq 0) {
            # Standardeigenschaften, falls nichts ausgewählt wurde
            $PropertiesToLoad = @("DisplayName", "SamAccountName") 
        }

        $Filter = "*"
        if ($CustomFilter) {
            $Filter = $CustomFilter
        } elseif ($FilterAttribute -and $FilterValue) {
            $Filter = "$($FilterAttribute) -like '$($FilterValue)*'"
        } elseif ($FilterValue -and (-not $FilterAttribute)) {
            $Filter = "DisplayName -like '$($FilterValue)*'"
        }
        
        $Users = $null # Initialize Users

        # Special handling for LockedOut accounts
        if ($CustomFilter -and $CustomFilter.Trim() -eq "LockedOut -eq `$true") {
            Write-Host "Führe Search-ADAccount -LockedOut -UsersOnly aus, um gesperrte Benutzer zu finden."
            # Ensure AD module is loaded; this is checked at the beginning of the function.
            $LockedOutAccounts = Search-ADAccount -LockedOut -UsersOnly -ErrorAction Stop 
            
            if ($LockedOutAccounts) {
                Write-Host "$($LockedOutAccounts.Count) gesperrte(s) Konto/Konten gefunden. Rufe Details für ausgewählte Attribute ab..."
                
                $Users = foreach ($Account in $LockedOutAccounts) {
                    try {
                        Get-ADUser -Identity $Account.SamAccountName -Properties $PropertiesToLoad -ErrorAction SilentlyContinue
                    } catch {
                        Write-Warning "Konnte Details für Benutzer $($Account.SamAccountName) nicht abrufen: $($_.Exception.Message)"
                        $null
                    }
                }
                # Filter out any null results from failed Get-ADUser calls
                $Users = $Users | Where-Object {$_ -ne $null}

            } else {
                Write-Host "Keine gesperrten Benutzerkonten gefunden via Search-ADAccount."
            }
        } else {
            # Standard AD User Abfrage für andere Filter
            Write-Host "Führe Get-ADUser mit Filter '$Filter' und Eigenschaften '$($PropertiesToLoad -join ', ')' aus"
            # Sicherstellen, dass Get-ADUser immer ein Array zurückgibt, auch bei nur einem Ergebnis
            $Users = @(Get-ADUser -Filter $Filter -Properties $PropertiesToLoad -ErrorAction Stop)
            Write-ADReportLog -Message "Gefundene Benutzer: $($Users.Count) - Typ: $($Users.GetType().FullName)" -Type Info -Terminal
        }

        if ($Users) {
            # Erstelle Array mit den bereinigten Attributnamen für Select-Object
            $SelectAttributes = @(
                $SelectedAttributes | ForEach-Object {
                    if ($null -ne $_) {
                        if ($_ -is [System.Windows.Controls.ListBoxItem]) {
                            $_.Content.ToString()
                        } else {
                            $_.ToString()
                        }
                    }
                } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            )
            
            # Verwende Select-Object, um ein Array von Objekten zu erstellen
            # Dies stellt sicher, dass wir eine IEnumerable-Sammlung zurückgeben
            $Output = $Users | Select-Object $SelectAttributes
            return $Output
        } else {
            Write-ADReportLog -Message "Keine Benutzer für den angegebenen Filter gefunden." -Type Warning
            return $null
        }
    } catch {
        $ErrorMessage = "Fehler bei der AD-Abfrage: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return $null
    }
}

# --- Funktion zum Abrufen von AD-Gruppendaten ---
Function Get-ADGroupReportData {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [System.Collections.IList]$SelectedAttributes,
        [Parameter(Mandatory=$false)]
        [string]$CustomFilter
    )

    # Überprüfen, ob das AD-Modul verfügbar ist
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        Write-ADReportLog -Message "Fehler: Active Directory Modul nicht gefunden." -Type Error
        return $null
    }

    try {
        # Konvertiere SelectedAttributes zu String-Array und filtere leere/null Werte
        $PropertiesToLoad = @(
            $SelectedAttributes | ForEach-Object {
                if ($null -ne $_) {
                    # Wenn es sich um ListBoxItem handelt, Content verwenden
                    if ($_ -is [System.Windows.Controls.ListBoxItem]) {
                        $_.Content.ToString()
                    } else {
                        $_.ToString()
                    }
                }
            } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        )

        if ($PropertiesToLoad.Count -eq 0) {
            $PropertiesToLoad = @("Name", "SamAccountName", "GroupCategory", "GroupScope") 
        }

        $FilterToUse = "*"
        if ($CustomFilter) {
            $FilterToUse = $CustomFilter
        }
        
        Write-Host "Führe Get-ADGroup mit Filter '$FilterToUse' und Eigenschaften '$($PropertiesToLoad -join ', ')' aus"
        $Groups = Get-ADGroup -Filter $FilterToUse -Properties $PropertiesToLoad -ErrorAction Stop

        if ($Groups) {
            # Erstelle Array mit den bereinigten Attributnamen für Select-Object
            $SelectAttributes = @(
                $SelectedAttributes | ForEach-Object {
                    if ($null -ne $_) {
                        if ($_ -is [System.Windows.Controls.ListBoxItem]) {
                            $_.Content.ToString()
                        } else {
                            $_.ToString()
                        }
                    }
                } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            )
            
            # Verwende Select-Object, um ein Array von Objekten zu erstellen
            # Dies stellt sicher, dass wir eine IEnumerable-Sammlung zurückgeben
            $Output = $Groups | Select-Object $SelectAttributes
            return $Output
        } else {
            Write-ADReportLog -Message "Keine Gruppen für den angegebenen Filter gefunden." -Type Warning
            return $null
        }
    } catch {
        $ErrorMessage = "Fehler bei der AD-Gruppenabfrage: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return $null
    }
}

# --- Globale Variablen für UI Elemente --- 
Function Initialize-ADReportForm {
    param($XAMLContent)
    $reader = New-Object System.Xml.XmlNodeReader $XAMLContent
    $Global:Window = [Windows.Markup.XamlReader]::Load( $reader )

    # --- UI Elemente zu globalen Variablen zuweisen --- 
    # Filter und Attribute
    $Global:ComboBoxFilterAttribute = $Window.FindName("ComboBoxFilterAttribute")
    $Global:TextBoxFilterValue = $Window.FindName("TextBoxFilterValue")
    $Global:ListBoxSelectAttributes = $Window.FindName("ListBoxSelectAttributes")
    $Global:ButtonQueryAD = $Window.FindName("ButtonQueryAD")

    # Quick Reports
    $Global:ButtonQuickAllUsers = $Window.FindName("ButtonQuickAllUsers")
    $Global:ButtonQuickDisabledUsers = $Window.FindName("ButtonQuickDisabledUsers")
    $Global:ButtonQuickLockedUsers = $Window.FindName("ButtonQuickLockedUsers")
    $Global:ButtonQuickGroups = $Window.FindName("ButtonQuickGroups")

    # Ergebnisse
    $Global:DataGridResults = $Window.FindName("DataGridResults")

    # Status und Export
    $Global:TextBlockStatus = $Window.FindName("TextBlockStatus")
    $Global:ButtonExportCSV = $Window.FindName("ButtonExportCSV")
    $Global:ButtonExportHTML = $Window.FindName("ButtonExportHTML")

    # --- UI Elemente füllen ---
    # Standardattribute für die Auswahl-ListBox füllen
    $DefaultAttributes = @("DisplayName", "SamAccountName", "GivenName", "Surname", "mail", "Department", "Title", "Enabled", "LastLogonDate", "whenCreated", "MemberOf")
    $DefaultAttributes | ForEach-Object { $Global:ListBoxSelectAttributes.Items.Add($_) }

    # Standardattribute für die Filter-ComboBox füllen
    $FilterableAttributes = @("DisplayName", "SamAccountName", "GivenName", "Surname", "mail", "Department", "Title")
    $FilterableAttributes | ForEach-Object { $Global:ComboBoxFilterAttribute.Items.Add($_) }
    if ($Global:ComboBoxFilterAttribute.Items.Count -gt 0) { $Global:ComboBoxFilterAttribute.SelectedIndex = 0 }

    # --- Event Handler zuweisen --- 
    $ButtonQueryAD.add_Click({
        Write-ADReportLog -Message "Führe Abfrage aus..." -Type Info
        try {
            $SelectedFilterAttribute = $Global:ComboBoxFilterAttribute.SelectedItem.ToString()
            $FilterValue = $Global:TextBoxFilterValue.Text
            $SelectedExportAttributes = $Global:ListBoxSelectAttributes.SelectedItems

            if ($SelectedExportAttributes.Count -eq 0) {
                Write-ADReportLog -Message "Bitte wählen Sie mindestens ein Attribut für den Export aus." -Type Warning
                return
            }

            # AD-Abfrage durchführen
            $ReportData = Get-ADReportData -FilterAttribute $SelectedFilterAttribute -FilterValue $FilterValue -SelectedAttributes $SelectedExportAttributes
            
            if ($ReportData) {
                try {
                    # Debug-Informationen
                    Write-ADReportLog -Message "ReportData Typ: $($ReportData.GetType().FullName)" -Type Info -Terminal
                    
                    # Wir brauchen sicherzustellen, dass wir immer eine Liste/Sammlung haben, auch bei einzelnen Objekten
                    # Verwende @() um es als Array zu erzwingen
                    $ReportCollection = @($ReportData)
                    
                    # Direkte Zuweisung an DataGrid
                    $Global:DataGridResults.ItemsSource = $ReportCollection
                    
                    # Zähle die Anzahl der Ergebnisse
                    $Count = $ReportCollection.Count
                    Write-ADReportLog -Message "Abfrage abgeschlossen. $Count Ergebnis(se) gefunden." -Type Info
                } catch {
                    $ErrorMessage = "Fehler beim Verarbeiten des Abfrageergebnisses: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    $Global:DataGridResults.ItemsSource = $null
                }
            } else {
                $Global:DataGridResults.ItemsSource = $null # DataGrid leeren
                # Status wird in Get-ADReportData gesetzt oder bleibt auf "Keine Benutzer gefunden"
            }
        } catch {
            $ErrorMessage = "Fehler im Abfrageprozess: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
        }
    })

    $ButtonQuickAllUsers.add_Click({
        Write-ADReportLog -Message "Lade alle Benutzer..." -Type Info
        try {
            $QuickReportAttributes = @("DisplayName", "SamAccountName", "Enabled", "LastLogonDate")
            $ReportData = Get-ADReportData -CustomFilter "*" -SelectedAttributes $QuickReportAttributes
            if ($ReportData) {
                $Global:DataGridResults.ItemsSource = $ReportData
                Write-ADReportLog -Message "Alle Benutzer geladen. $($ReportData.Count) Ergebnis(se) gefunden." -Type Info
            } else {
                $Global:DataGridResults.ItemsSource = $null
                # Status wird in Get-ADReportData gesetzt
            }
        } catch {
            $ErrorMessage = "Fehler beim Laden aller Benutzer: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
        }
    })

    $ButtonQuickDisabledUsers.add_Click({
        Write-ADReportLog -Message "Lade deaktivierte Benutzer..." -Type Info
        try {
            $QuickReportAttributes = @("DisplayName", "SamAccountName", "Enabled", "LastLogonDate")
            # Der Filter für deaktivierte Benutzer ist 'Enabled -eq $false'
            $ReportData = Get-ADReportData -CustomFilter "Enabled -eq `$false" -SelectedAttributes $QuickReportAttributes
            if ($ReportData) {
                $Global:DataGridResults.ItemsSource = $ReportData
                Write-ADReportLog -Message "Deaktivierte Benutzer geladen. $($ReportData.Count) Ergebnis(se) gefunden." -Type Info
            } else {
                $Global:DataGridResults.ItemsSource = $null
                # Status wird in Get-ADReportData gesetzt
            }
        } catch {
            $ErrorMessage = "Fehler beim Laden deaktivierter Benutzer: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
        }
    })

    $ButtonQuickLockedUsers.add_Click({
        Write-ADReportLog -Message "Lade gesperrte Benutzer..." -Type Info
        try {
            $QuickReportAttributes = @("DisplayName", "SamAccountName", "Enabled", "LockedOut", "LastLogonDate")
            # Der Filter für gesperrte Benutzer ist 'LockedOut -eq $true'
            # Wichtig: Get-ADUser muss -Properties LockedOut explizit anfordern, damit das Attribut vorhanden ist.
            $ReportData = Get-ADReportData -CustomFilter "LockedOut -eq `$true" -SelectedAttributes $QuickReportAttributes
            if ($ReportData) {
                $Global:DataGridResults.ItemsSource = $ReportData
                Write-ADReportLog -Message "Gesperrte Benutzer geladen. $($ReportData.Count) Ergebnis(se) gefunden." -Type Info
            } else {
                $Global:DataGridResults.ItemsSource = $null
                # Status wird in Get-ADReportData gesetzt
            }
        } catch {
            $ErrorMessage = "Fehler beim Laden gesperrter Benutzer: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
        }
    })

    $ButtonQuickGroups.add_Click({
        Write-ADReportLog -Message "Lade alle Gruppen..." -Type Info
        try {
            $QuickReportAttributes = @("Name", "SamAccountName", "GroupCategory", "GroupScope")
            $ReportData = Get-ADGroupReportData -CustomFilter "*" -SelectedAttributes $QuickReportAttributes
            if ($ReportData) {
                $Global:DataGridResults.ItemsSource = $ReportData
                Write-ADReportLog -Message "Alle Gruppen geladen. $($ReportData.Count) Ergebnis(se) gefunden." -Type Info
            } else {
                $Global:DataGridResults.ItemsSource = $null
                # Status wird in Get-ADGroupReportData gesetzt
            }
        } catch {
            $ErrorMessage = "Fehler beim Laden aller Gruppen: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
        }
    })

    $ButtonExportCSV.add_Click({
        Write-ADReportLog -Message "CSV Export wird vorbereitet..." -Type Info
        if ($null -eq $Global:DataGridResults.ItemsSource -or $Global:DataGridResults.Items.Count -eq 0) {
            Write-ADReportLog -Message "Keine Daten zum Exportieren vorhanden." -Type Warning
            [System.Windows.Forms.MessageBox]::Show("Es sind keine Daten zum Exportieren vorhanden.", "Hinweis", "OK", "Information")
            return
        }

        $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $SaveFileDialog.Filter = "CSV (Comma delimited) (*.csv)|*.csv"
        $SaveFileDialog.Title = "CSV-Datei speichern unter"
        $SaveFileDialog.InitialDirectory = [Environment]::GetFolderPath("MyDocuments")

        if ($SaveFileDialog.ShowDialog() -eq "OK") {
            $FilePath = $SaveFileDialog.FileName
            try {
                $Global:DataGridResults.ItemsSource | Export-Csv -Path $FilePath -NoTypeInformation -Encoding UTF8 -Delimiter ';'
                $Global:TextBlockStatus.Text = "Daten erfolgreich nach $FilePath exportiert."
                [System.Windows.Forms.MessageBox]::Show("Daten erfolgreich exportiert!", "CSV Export", "OK", "Information")
            } catch {
                $ErrorMessage = "Fehler beim CSV-Export: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                [System.Windows.Forms.MessageBox]::Show("Fehler beim Exportieren der Daten: $($_.Exception.Message)", "Export Fehler", "OK", "Error")
            }
        } else {
            Write-ADReportLog -Message "CSV-Export vom Benutzer abgebrochen." -Type Info
        }
    })

    $ButtonExportHTML.add_Click({
        Write-ADReportLog -Message "HTML Export wird vorbereitet..." -Type Info
        if ($null -eq $Global:DataGridResults.ItemsSource -or $Global:DataGridResults.Items.Count -eq 0) {
            $Global:TextBlockStatus.Text = "Keine Daten zum Exportieren vorhanden."
            [System.Windows.Forms.MessageBox]::Show("Es sind keine Daten zum Exportieren vorhanden.", "Hinweis", "OK", "Information")
            return
        }

        $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $SaveFileDialog.Filter = "HTML-Datei (*.html;*.htm)|*.html;*.htm"
        $SaveFileDialog.Title = "HTML-Datei speichern unter"
        $SaveFileDialog.InitialDirectory = [Environment]::GetFolderPath("MyDocuments")

        if ($SaveFileDialog.ShowDialog() -eq "OK") {
            $FilePath = $SaveFileDialog.FileName
            try {
                # Erstelle einen ansprechenderen HTML-Header
                $HtmlHead = @"
<meta http-equiv='Content-Type' content='text/html; charset=utf-8' />
<title>Active Directory Report</title>
<style>
  body { font-family: Segoe UI, Arial, sans-serif; margin: 20px; }
  table { border-collapse: collapse; width: 90%; margin: 20px auto; box-shadow: 0 0 10px rgba(0,0,0,0.1); }
  th, td { border: 1px solid #ddd; padding: 10px; text-align: left; }
  th { background-color: #0078D4; color: white; }
  tr:nth-child(even) { background-color: #f2f2f2; }
  h1 { text-align: center; color: #333; }
</style>
"@
                $DateTimeNow = Get-Date -Format "dd.MM.yyyy HH:mm:ss"
                $ReportTitle = "Active Directory Report - Erstellt am $DateTimeNow"

                $Global:DataGridResults.ItemsSource | ConvertTo-Html -Head $HtmlHead -Body "<h1>$ReportTitle</h1>" | Out-File -FilePath $FilePath -Encoding UTF8
                $Global:TextBlockStatus.Text = "Daten erfolgreich nach $FilePath exportiert."
                [System.Windows.Forms.MessageBox]::Show("Daten erfolgreich exportiert!", "HTML Export", "OK", "Information")
            } catch {
                $ErrorMessage = "Fehler beim HTML-Export: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                [System.Windows.Forms.MessageBox]::Show("Fehler beim Exportieren der Daten: $($_.Exception.Message)", "Export Fehler", "OK", "Error")
            }
        } else {
            Write-ADReportLog -Message "HTML-Export vom Benutzer abgebrochen." -Type Info
        }
    })

    # Fenster anzeigen
    $null = $Window.ShowDialog()
}

# --- Hauptlogik --- 
Function Start-ADReportGUI {
    # Ruft Initialize-ADReportForm auf, welche die UI lädt, Elemente zuweist und füllt.
    Initialize-ADReportForm -XAMLContent $Global:XAML
}

# --- Skriptstart --- 
Start-ADReportGUI