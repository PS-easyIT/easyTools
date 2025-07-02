[xml]$Global:XAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        Title="easyADReport v0.0.17" Height="850" Width="1200"
        WindowStartupLocation="CenterScreen" ResizeMode="CanResizeWithGrip"
        Background="#F9F9F9" FontFamily="Segoe UI">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="60"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="50"/>
        </Grid.RowDefinitions>

        <!-- Header -->
        <Border Grid.Row="0" Background="#0078D7" BorderThickness="0">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                
                <!-- Title und Version -->
                <StackPanel Grid.Column="0" Orientation="Horizontal" Margin="20,0">
                    <TextBlock Text="easyADReport" Foreground="White" FontSize="22" FontWeight="SemiBold" VerticalAlignment="Center"/>
                    <TextBlock Text="v0.0.19" Foreground="#E1E1E1" FontSize="14" VerticalAlignment="Center" Margin="10,0,0,0"/>
                </StackPanel>
                
                <!-- Ergebnisanzahl-Anzeige -->
                <Border Grid.Column="1" Background="#0063B1" Width="300" Height="50" CornerRadius="4" Margin="0,5,20,5">
                    <Grid x:Name="ResultCountGrid" Margin="5">
                        <!-- Gesamtergebnis-Anzeige -->
                        <Border Background="#0063B1" CornerRadius="2" Margin="2">
                            <Grid>
                                <TextBlock Text="Ergebnisse gesamt" Foreground="White" FontSize="14" HorizontalAlignment="Center" VerticalAlignment="Top" Margin="0,5,0,0"/>
                                <TextBlock x:Name="TotalResultCountText" Text="0" Foreground="White" FontSize="20" FontWeight="Bold" HorizontalAlignment="Center" VerticalAlignment="Center" Margin="0,10,0,0"/>
                            </Grid>
                        </Border>
                    </Grid>
                </Border>
            </Grid>
        </Border>

        <!-- Upper Section: Filter and Attribute Selection -->
        <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="20,15">
            <Border Background="White" CornerRadius="4" BorderBrush="#DDDDDD" BorderThickness="1" Margin="0,0,15,0">
                <GroupBox Header="Filter" Margin="5" BorderThickness="0">
                    <StackPanel Orientation="Horizontal">
                        <Label Content="Filter Attribute:" VerticalAlignment="Center"/>
                        <ComboBox x:Name="ComboBoxFilterAttribute" Width="150" Margin="5,0" VerticalAlignment="Center" BorderThickness="1" BorderBrush="#CCCCCC"/>
                        <Label Content="Filter Value:" VerticalAlignment="Center"/>
                        <TextBox x:Name="TextBoxFilterValue" Width="200" Margin="5,0" VerticalAlignment="Center" BorderThickness="1" BorderBrush="#CCCCCC"/>
                    </StackPanel>
                </GroupBox>
            </Border>
            <Border Background="White" CornerRadius="4" BorderBrush="#DDDDDD" BorderThickness="1" Margin="0,0,15,0">
                <GroupBox Header="Attributes to Export" Margin="5" BorderThickness="0">
                    <ListBox x:Name="ListBoxSelectAttributes" Width="300" Height="80" SelectionMode="Multiple" BorderThickness="0">
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
            </Border>
            <Button x:Name="ButtonQueryAD" Content="Query" Width="100" Height="36" Margin="0,5" VerticalAlignment="Center" 
                    Background="#0078D7" Foreground="White" BorderThickness="0">
                <Button.Resources>
                    <Style TargetType="Border">
                        <Setter Property="CornerRadius" Value="4"/>
                    </Style>
                </Button.Resources>
            </Button>
        </StackPanel>

        <!-- Middle Section: Quick Reports and Results -->
        <Grid Grid.Row="2" Margin="20,0,20,15">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="220"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

            <Border Background="White" CornerRadius="4" BorderBrush="#DDDDDD" BorderThickness="1" Margin="0,0,15,0">
                <GroupBox Header="Quick Reports" BorderThickness="0" Margin="5,0,0,0">
                    <StackPanel>
                        <GroupBox Header="Users" BorderThickness="0" Margin="5,10,0,0">
                            <StackPanel>
                                <Button x:Name="ButtonQuickAllUsers" Content="All Users" Margin="0,2" Height="28" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                <Button x:Name="ButtonQuickDisabledUsers" Content="Disabled Users" Margin="0,2" Height="28" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                <Button x:Name="ButtonQuickLockedUsers" Content="Locked Users" Margin="0,2" Height="28" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                <Button x:Name="ButtonQuickNeverExpire" Content="Password Never Expires" Margin="0,2" Height="28" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                <Button x:Name="ButtonQuickInactiveUsers" Content="Inactive Users (90+ days)" Margin="0,2" Height="28" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                <Button x:Name="ButtonQuickAdminUsers" Content="Administrators" Margin="0,2" Height="28" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                            </StackPanel>
                        </GroupBox>
                        <GroupBox Header="Groups" BorderThickness="0" Margin="5,15,0,0">
                            <StackPanel>
                                <Button x:Name="ButtonQuickGroups" Content="All Groups" Margin="0,2" Height="28" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                <Button x:Name="ButtonQuickSecurityGroups" Content="Security Groups" Margin="0,2" Height="28" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                <Button x:Name="ButtonQuickDistributionGroups" Content="Distribution Lists" Margin="0,2" Height="28" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                            </StackPanel>
                        </GroupBox>
                        <GroupBox Header="Computers" BorderThickness="0" Margin="5,15,0,0">
                            <StackPanel>
                                <Button x:Name="ButtonQuickComputers" Content="All Computers" Margin="0,2" Height="28" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                <Button x:Name="ButtonQuickInactiveComputers" Content="Inactive Computers (90+ days)" Margin="0,2" Height="28" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                            </StackPanel>
                        </GroupBox>
                    </StackPanel>
                </GroupBox>
            </Border>

            <Border Grid.Column="1" Background="White" CornerRadius="4" BorderBrush="#DDDDDD" BorderThickness="1">
                <GroupBox BorderThickness="0" Margin="5">
                    <DataGrid x:Name="DataGridResults" AutoGenerateColumns="True" IsReadOnly="True" BorderThickness="0"
                              Background="Transparent" GridLinesVisibility="All" RowBackground="White" AlternatingRowBackground="#F5F5F5"/>
                </GroupBox>
            </Border>
        </Grid>

        <!-- Footer -->
        <Border Grid.Row="3" Background="#F0F0F0" BorderThickness="0,1,0,0" BorderBrush="#DDDDDD">
            <Grid Margin="20,0">
                <TextBlock x:Name="TextBlockStatus" Text="Ready" VerticalAlignment="Center"/>
                <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                    <Button x:Name="ButtonExportCSV" Content="Export CSV" Width="100" Height="28" Margin="5,6" Background="#F0F0F0" BorderBrush="#CCCCCC">
                        <Button.Resources>
                            <Style TargetType="Border">
                                <Setter Property="CornerRadius" Value="3"/>
                            </Style>
                        </Button.Resources>
                    </Button>
                    <Button x:Name="ButtonExportHTML" Content="Export HTML" Width="100" Height="28" Margin="5,6" Background="#F0F0F0" BorderBrush="#CCCCCC">
                        <Button.Resources>
                            <Style TargetType="Border">
                                <Setter Property="CornerRadius" Value="3"/>
                            </Style>
                        </Button.Resources>
                    </Button>
                </StackPanel>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

# Setze die Ausgabekodierung auf UTF-8, um Probleme mit Umlauten zu vermeiden
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

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
        Write-ADReportLog -Message "Error: Active Directory module not found." -Type Error
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
            # Default properties if nothing was selected
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
            Write-Host "Executing Get-ADUser with filter '$Filter' and properties '$($PropertiesToLoad -join ', ')'"
            # Sicherstellen, dass Get-ADUser immer ein Array zurückgibt, auch bei nur einem Ergebnis
            $Users = @(Get-ADUser -Filter $Filter -Properties $PropertiesToLoad -ErrorAction Stop)
            Write-ADReportLog -Message "Users found: $($Users.Count) - Type: $($Users.GetType().FullName)" -Type Info -Terminal
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
            Write-ADReportLog -Message "No users found for the specified filter." -Type Warning
            return $null
        }
    } catch {
        $ErrorMessage = "Error in AD query: $($_.Exception.Message)"
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
            Write-ADReportLog -Message "No groups found for the specified filter." -Type Warning
            return $null
        }
    } catch {
        $ErrorMessage = "Error in AD group query: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return $null
    }
}

# --- Globale Variablen für UI Elemente --- 
Function Initialize-ADReportForm {
    param($XAMLContent)
    # Überprüfen, ob das Window-Objekt bereits existiert und zurücksetzen
    if ($Global:Window) {
        Remove-Variable -Name Window -Scope Global -ErrorAction SilentlyContinue
    }
    
    $reader = New-Object System.Xml.XmlNodeReader $XAMLContent
    $Global:Window = [Windows.Markup.XamlReader]::Load( $reader )

    # --- UI Elemente zu globalen Variablen zuweisen --- 
    # Filter und Attribute
    $Global:ComboBoxFilterAttribute = $Window.FindName("ComboBoxFilterAttribute")
    $Global:TextBoxFilterValue = $Window.FindName("TextBoxFilterValue")
    $Global:ListBoxSelectAttributes = $Window.FindName("ListBoxSelectAttributes")
    $Global:ButtonQueryAD = $Window.FindName("ButtonQueryAD")

    # Quick Reports - Benutzer
    $Global:ButtonQuickAllUsers = $Window.FindName("ButtonQuickAllUsers")
    $Global:ButtonQuickDisabledUsers = $Window.FindName("ButtonQuickDisabledUsers")
    $Global:ButtonQuickLockedUsers = $Window.FindName("ButtonQuickLockedUsers")
    $Global:ButtonQuickNeverExpire = $Window.FindName("ButtonQuickNeverExpire")
    $Global:ButtonQuickInactiveUsers = $Window.FindName("ButtonQuickInactiveUsers")
    $Global:ButtonQuickAdminUsers = $Window.FindName("ButtonQuickAdminUsers")
    
    # Quick Reports - Gruppen
    $Global:ButtonQuickGroups = $Window.FindName("ButtonQuickGroups")
    $Global:ButtonQuickSecurityGroups = $Window.FindName("ButtonQuickSecurityGroups")
    $Global:ButtonQuickDistributionGroups = $Window.FindName("ButtonQuickDistributionGroups")
    
    # Quick Reports - Computer
    $Global:ButtonQuickComputers = $Window.FindName("ButtonQuickComputers")
    $Global:ButtonQuickInactiveComputers = $Window.FindName("ButtonQuickInactiveComputers")

    # Ergebnisanzahl-Anzeige
    $Global:ResultCountGrid = $Window.FindName("ResultCountGrid")
    $Global:TotalResultCountText = $Window.FindName("TotalResultCountText")
    $Global:UserCountText = $Window.FindName("UserCountText")
    $Global:ComputerCountText = $Window.FindName("ComputerCountText")
    $Global:GroupCountText = $Window.FindName("GroupCountText")
    
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

    # Funktion zur Aktualisierung der Ergebniszähler im Header
Function Update-ResultCounters {
    param (
        [Parameter(Mandatory=$true)]
        $Results
    )
    
    Write-ADReportLog -Message "Aktualisiere Ergebnisanzeige..." -Type Info -Terminal
    
    try {
        # Ergebnisanzahl ermitteln
        $totalCount = 0
        
        if ($null -ne $Results) {
            $totalCount = $Results.Count
            
            # Objekttyp für das Log bestimmen (nur für die Protokollierung)
            if ($totalCount -gt 0) {
                $firstObject = $Results[0]
                $objectType = "Unbekannt"
                
                if ($firstObject.PSObject.Properties.Name -contains "ObjectClass") {
                    switch ($firstObject.ObjectClass) {
                        "user" { $objectType = "Benutzer" }
                        "computer" { $objectType = "Computer" }
                        "group" { $objectType = "Gruppen" }
                    }
                }
                
                Write-ADReportLog -Message "$objectType-Ergebnisse: $totalCount" -Type Info -Terminal
            } else {
                Write-ADReportLog -Message "Keine Ergebnisse gefunden" -Type Info -Terminal
            }
        } else {
            Write-ADReportLog -Message "Keine Ergebnisse (Null)" -Type Info -Terminal
        }
        
        # Überprüfen, ob die UI-Elemente existieren, bevor wir versuchen, sie zu aktualisieren
        if ($null -ne $Global:TotalResultCountText) {
            $Global:TotalResultCountText.Text = $totalCount.ToString()
        }
        
        # Aktualisiere den Status im Status-Textblock
        if ($null -ne $Global:TextBlockStatus) {
            $Global:TextBlockStatus.Text = "Abfrage abgeschlossen. $totalCount Ergebnis(se) gefunden."
        }
    }
    catch {
        Write-ADReportLog -Message "Fehler beim Aktualisieren der Ergebnisanzeige: $($_.Exception.Message)" -Type Error
    }
}

# Funktion zum Initialisieren der Ergebnisanzeige
Function Initialize-ResultCounters {
    # Gesamtergebniszähler auf 0 setzen
    if ($null -ne $Global:TotalResultCountText) {
        $Global:TotalResultCountText.Text = "0"
    }
    
    # Sicherstellen, dass alle Zähler zurückgesetzt werden
    if ($null -ne $Global:UserCountText) {
        $Global:UserCountText.Text = "0"
    }
    
    if ($null -ne $Global:ComputerCountText) {
        $Global:ComputerCountText.Text = "0"
    }
    
    if ($null -ne $Global:GroupCountText) {
        $Global:GroupCountText.Text = "0"
    }
    
    # Status zurücksetzen
    if ($null -ne $Global:TextBlockStatus) {
        $Global:TextBlockStatus.Text = "Bereit für Abfrage..."
    }
    
    # DataGrid leeren
    if ($null -ne $Global:DataGridResults) {
        $Global:DataGridResults.ItemsSource = $null
    }
}

# Funktion zur universellen Aktualisierung der Ergebnisanzeige und DataGrid
Function Update-ADReportResults {
    param (
        [Parameter(Mandatory=$false)]
        $Results = $null,
        
        [Parameter(Mandatory=$false)]
        [string]$StatusMessage = ""
    )
    
    # DataGrid aktualisieren
    if ($null -ne $Global:DataGridResults) {
        $Global:DataGridResults.ItemsSource = $Results
    }
    
    # Ergebniszähler aktualisieren
    Update-ResultCounters -Results $Results
    
    # Statusmeldung anzeigen, wenn vorhanden
    if (-not [string]::IsNullOrWhiteSpace($StatusMessage) -and $null -ne $Global:TextBlockStatus) {
        $Global:TextBlockStatus.Text = $StatusMessage
    }
}

    # --- Event Handler zuweisen --- 
    $ButtonQueryAD.add_Click({
        Write-ADReportLog -Message "Executing query..." -Type Info
        try {
            $SelectedFilterAttribute = $Global:ComboBoxFilterAttribute.SelectedItem.ToString()
            $FilterValue = $Global:TextBoxFilterValue.Text
            $SelectedExportAttributes = $Global:ListBoxSelectAttributes.SelectedItems

            if ($SelectedExportAttributes.Count -eq 0) {
                Write-ADReportLog -Message "Please select at least one attribute for export." -Type Warning
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
                    Write-ADReportLog -Message "Query completed. $Count result(s) found." -Type Info
                    
                    # Ergebniszähler im Header aktualisieren
                    Update-ResultCounters -Results $ReportCollection
                } catch {
                    $ErrorMessage = "Error processing query result: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    $Global:DataGridResults.ItemsSource = $null
                }
            } else {
                $Global:DataGridResults.ItemsSource = $null # DataGrid leeren
                # Status wird in Get-ADReportData gesetzt oder bleibt auf "Keine Benutzer gefunden"
            }
        } catch {
            $ErrorMessage = "Error in query process: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
        }
    })

    $ButtonQuickAllUsers.add_Click({
        Write-ADReportLog -Message "Loading all users..." -Type Info
        try {
            $QuickReportAttributes = @("DisplayName", "SamAccountName", "Enabled", "LastLogonDate")
            $ReportData = Get-ADReportData -CustomFilter "*" -SelectedAttributes $QuickReportAttributes
            
            # Verwende die neue universelle Funktion zur Aktualisierung der Ergebnisse
            if ($ReportData) {
                $StatusMessage = "All users loaded. $($ReportData.Count) result(s) found."
            } else {
                $StatusMessage = "Keine Benutzer gefunden."
                $ReportData = @() # Leeres Array bei keinen Ergebnissen
            }
            
            Write-ADReportLog -Message $StatusMessage -Type Info
            Update-ADReportResults -Results $ReportData -StatusMessage $StatusMessage
        } catch {
            $ErrorMessage = "Error loading all users: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            Update-ADReportResults -Results @() -StatusMessage $ErrorMessage
        }
    })

    $ButtonQuickDisabledUsers.add_Click({
        Write-ADReportLog -Message "Loading disabled users..." -Type Info
        try {
            $QuickReportAttributes = @("DisplayName", "SamAccountName", "Enabled", "LastLogonDate")
            # Der Filter für deaktivierte Benutzer ist 'Enabled -eq $false'
            $ReportData = Get-ADReportData -CustomFilter "Enabled -eq `$false" -SelectedAttributes $QuickReportAttributes
            
            # Verwende die neue universelle Funktion zur Aktualisierung der Ergebnisse
            if ($ReportData) {
                $StatusMessage = "Disabled users loaded. $($ReportData.Count) result(s) found."
            } else {
                $StatusMessage = "Keine deaktivierten Benutzer gefunden."
                $ReportData = @() # Leeres Array bei keinen Ergebnissen
            }
            
            Write-ADReportLog -Message $StatusMessage -Type Info
            Update-ADReportResults -Results $ReportData -StatusMessage $StatusMessage
        } catch {
            $ErrorMessage = "Error loading disabled users: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            Update-ADReportResults -Results @() -StatusMessage $ErrorMessage
        }
    })

    $ButtonQuickLockedUsers.add_Click({
        Write-ADReportLog -Message "Loading locked users..." -Type Info
        try {
            # Verwende Search-ADAccount, da es zuverlässiger für gesperrte Konten ist
            Write-ADReportLog -Message "Executing Search-ADAccount -LockedOut -UsersOnly to find locked users." -Type Info -Terminal
            $LockedUsers = Search-ADAccount -LockedOut -UsersOnly
            
            if ($LockedUsers -and $LockedUsers.Count -gt 0) {
                # Hole zusätzliche Eigenschaften für die gefundenen Benutzer
                $LockedUsersData = $LockedUsers | ForEach-Object {
                    Get-ADUser -Identity $_.SamAccountName -Properties DisplayName, SamAccountName, Enabled, LockedOut, LastLogonDate
                } | Select-Object DisplayName, SamAccountName, Enabled, LockedOut, LastLogonDate
                
                $Global:DataGridResults.ItemsSource = $LockedUsersData
                Write-ADReportLog -Message "Locked users loaded. $($LockedUsersData.Count) result(s) found." -Type Info
                Update-ResultCounters -Results $LockedUsersData
            } else {
                # Versuche es mit der Filter-Methode als Fallback
                $QuickReportAttributes = @("DisplayName", "SamAccountName", "Enabled", "LockedOut", "LastLogonDate")
                $ReportData = Get-ADReportData -CustomFilter "LockedOut -eq `$true" -SelectedAttributes $QuickReportAttributes
                
                if ($ReportData -and $ReportData.Count -gt 0) {
                    $Global:DataGridResults.ItemsSource = $ReportData
                    Write-ADReportLog -Message "Locked users loaded. $($ReportData.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $ReportData
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No locked user accounts found." -Type Warning
                }
            }
        } catch {
            $ErrorMessage = "Error loading locked users: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
        }
    })

    $ButtonQuickGroups.add_Click({
        Write-ADReportLog -Message "Loading all groups..." -Type Info
        try {
            $QuickReportAttributes = @("Name", "SamAccountName", "GroupCategory", "GroupScope")
            $ReportData = Get-ADGroupReportData -CustomFilter "*" -SelectedAttributes $QuickReportAttributes
            if ($ReportData) {
                $Global:DataGridResults.ItemsSource = $ReportData
                Write-ADReportLog -Message "All groups loaded. $($ReportData.Count) result(s) found." -Type Info
                Update-ResultCounters -Results $ReportData
            } else {
                $Global:DataGridResults.ItemsSource = $null
                # Status wird in Get-ADGroupReportData gesetzt
            }
        } catch {
            $ErrorMessage = "Error loading all groups: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
        }
    })

    $ButtonQuickSecurityGroups.add_Click({
        Write-ADReportLog -Message "Loading security groups..." -Type Info
        try {
            $QuickReportAttributes = @("Name", "SamAccountName", "GroupCategory", "GroupScope")
            $ReportData = Get-ADGroupReportData -CustomFilter "GroupCategory -eq 'Security'" -SelectedAttributes $QuickReportAttributes
            if ($ReportData) {
                $Global:DataGridResults.ItemsSource = $ReportData
                Write-ADReportLog -Message "Security groups loaded. $($ReportData.Count) result(s) found." -Type Info
                Update-ResultCounters -Results $ReportData
            } else {
                $Global:DataGridResults.ItemsSource = $null
            }
        } catch {
            $ErrorMessage = "Error loading security groups: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
        }
    })

    $ButtonQuickDistributionGroups.add_Click({
        Write-ADReportLog -Message "Loading distribution lists..." -Type Info
        try {
            $QuickReportAttributes = @("Name", "SamAccountName", "GroupCategory", "GroupScope")
            $ReportData = Get-ADGroupReportData -CustomFilter "GroupCategory -eq 'Distribution'" -SelectedAttributes $QuickReportAttributes
            if ($ReportData) {
                $Global:DataGridResults.ItemsSource = $ReportData
                Write-ADReportLog -Message "Distribution lists loaded. $($ReportData.Count) result(s) found." -Type Info
                Update-ResultCounters -Results $ReportData
            } else {
                $Global:DataGridResults.ItemsSource = $null
            }
        } catch {
            $ErrorMessage = "Error loading distribution lists: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
        }
    })

    # Neue Funktionen für Benutzer
    $ButtonQuickNeverExpire.add_Click({
        Write-ADReportLog -Message "Loading users with non-expiring passwords..." -Type Info
        try {
            $QuickReportAttributes = @("DisplayName", "SamAccountName", "Enabled", "PasswordNeverExpires", "LastLogonDate")
            $ReportData = Get-ADReportData -CustomFilter "PasswordNeverExpires -eq `$true" -SelectedAttributes $QuickReportAttributes
            if ($ReportData) {
                $Global:DataGridResults.ItemsSource = $ReportData
                Write-ADReportLog -Message "Users with non-expiring passwords loaded. $($ReportData.Count) result(s) found." -Type Info
                Update-ResultCounters -Results $ReportData
            } else {
                $Global:DataGridResults.ItemsSource = $null
            }
        } catch {
            $ErrorMessage = "Error loading users with non-expiring passwords: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
        }
    })

    $ButtonQuickInactiveUsers.add_Click({
        Write-ADReportLog -Message "Loading inactive users (>90 days)..." -Type Info
        try {
            $QuickReportAttributes = @("DisplayName", "SamAccountName", "Enabled", "LastLogonDate")
            $InactiveCutoffDate = (Get-Date).AddDays(-90)
            $ReportData = Get-ADReportData -CustomFilter "(LastLogonTimeStamp -lt '$($InactiveCutoffDate.ToFileTime())') -or (LastLogonTimeStamp -notlike '*')" -SelectedAttributes $QuickReportAttributes
            if ($ReportData) {
                # Nachbearbeitung: Nur wirklich inaktive Benutzer behalten
                $FilteredData = $ReportData | Where-Object { 
                    !$_.LastLogonDate -or ($_.LastLogonDate -lt $InactiveCutoffDate)
                }
                $Global:DataGridResults.ItemsSource = $FilteredData
                Write-ADReportLog -Message "Inactive users loaded. $($FilteredData.Count) result(s) found." -Type Info
                Update-ResultCounters -Results $FilteredData
            } else {
                $Global:DataGridResults.ItemsSource = $null
            }
        } catch {
            $ErrorMessage = "Error loading inactive users: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
        }
    })

    $ButtonQuickAdminUsers.add_Click({
        Write-ADReportLog -Message "Loading administrators..." -Type Info
        try {
            # Verbesserte Methode zum Finden von Admin-Benutzern
            # Zuerst alle Benutzer laden
            $AllUsers = Get-ADUser -Filter * -Properties DisplayName, SamAccountName, Enabled, LastLogonDate
            
            # Bekannte Admin-Gruppenbezeichnungen (deutsch und englisch)
            $AdminGroups = @(
                "Domain Admins", "Domänen-Admins",
                "Enterprise Admins", "Organisations-Admins",
                "Administrators", "Administratoren", 
                "Schema Admins", "Schema-Admins",
                "AAD DC Administrators", "Azure AD-DC-Administratoren",
                "Server Operators", "Server-Operatoren",
                "Account Operators", "Konten-Operatoren",
                "Backup Operators", "Sicherungs-Operatoren"
            )
            
            $AdminUsers = @()
            # Prüfe Benutzer auf Admin-Rechte - erst SIDs der Admin-Gruppen ermitteln
            $AdminGroupSIDs = @()
            foreach ($groupName in $AdminGroups) {
                try {
                    $group = Get-ADGroup -Filter "Name -eq '$groupName'" -ErrorAction SilentlyContinue
                    if ($group) {
                        $AdminGroupSIDs += $group.SID.Value
                    }
                } catch {
                    # Ignoriere Fehler bei nicht existierenden Gruppen
                }
            }
            
            # Nur fortfahren, wenn wir Admin-Gruppen gefunden haben
            if ($AdminGroupSIDs.Count -gt 0) {
                # Für jeden Benutzer die Gruppenmitgliedschaften prüfen
                foreach ($user in $AllUsers) {
                    $memberOf = Get-ADPrincipalGroupMembership -Identity $user.SamAccountName -ErrorAction SilentlyContinue
                    foreach ($group in $memberOf) {
                        if ($AdminGroupSIDs -contains $group.SID.Value) {
                            # Benutzer ist Mitglied einer Admin-Gruppe
                            $adminUser = $user | Select-Object DisplayName, SamAccountName, Enabled, LastLogonDate, 
                                @{Name='AdminGroups'; Expression={
                                    ($memberOf | Where-Object { $AdminGroupSIDs -contains $_.SID.Value } | Select-Object -ExpandProperty Name) -join ", "
                                }}
                            $AdminUsers += $adminUser
                            break # Genug, wenn wir eine Admin-Gruppe gefunden haben
                        }
                    }
                }
            }
            
            if ($AdminUsers.Count -gt 0) {
                $Global:DataGridResults.ItemsSource = $AdminUsers
                Write-ADReportLog -Message "Administrators loaded. $($AdminUsers.Count) result(s) found." -Type Info
                Update-ResultCounters -Results $AdminUsers
            } else {
                $Global:DataGridResults.ItemsSource = $null
                Write-ADReportLog -Message "No administrator accounts found." -Type Warning
            }
        } catch {
            $ErrorMessage = "Error loading administrators: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
        }
    })

    # Neue Funktionen für Computer
    $ButtonQuickComputers.add_Click({
        Write-ADReportLog -Message "Loading all computers..." -Type Info
        try {
            $QuickReportAttributes = @("Name", "DNSHostName", "OperatingSystem", "Enabled", "LastLogonDate")
            $ReportData = Get-ADComputer -Filter * -Properties $QuickReportAttributes | Select-Object $QuickReportAttributes
            if ($ReportData) {
                $Global:DataGridResults.ItemsSource = $ReportData
                Write-ADReportLog -Message "All computers loaded. $($ReportData.Count) result(s) found." -Type Info
                Update-ResultCounters -Results $ReportData
            } else {
                $Global:DataGridResults.ItemsSource = $null
                Write-ADReportLog -Message "No computers found." -Type Warning
            }
        } catch {
            $ErrorMessage = "Error loading all computers: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
        }
    })

    $ButtonQuickInactiveComputers.add_Click({
        Write-ADReportLog -Message "Loading inactive computers (>90 days)..." -Type Info
        try {
            $QuickReportAttributes = @("Name", "DNSHostName", "OperatingSystem", "Enabled", "LastLogonDate", "PasswordLastSet")
            
            # Direktes Anfordern der spezifischen Computer-Eigenschaften, die wir brauchen
            $AllComputers = Get-ADComputer -Filter * -Properties Name, DNSHostName, OperatingSystem, Enabled, LastLogonDate, LastLogonTimeStamp, PasswordLastSet
            
            # Filtern nach inaktiven Computern basierend auf LastLogonDate und LastLogonTimeStamp
            $InactiveComputers = $AllComputers | Where-Object {
                ($_.LastLogonDate -lt ((Get-Date).AddDays(-90)) -or $_.LastLogonDate -eq $null) -and
                ($_.LastLogonTimeStamp -lt ((Get-Date).AddDays(-90)).ToFileTime() -or $_.LastLogonTimeStamp -eq $null)
            } | Select-Object Name, DNSHostName, OperatingSystem, Enabled, LastLogonDate, PasswordLastSet,
              @{Name='InactiveDays'; Expression={
                if ($_.LastLogonDate) {
                    [math]::Round((New-TimeSpan -Start $_.LastLogonDate -End (Get-Date)).TotalDays)
                } else { "N/A" }
              }}
            
            if ($InactiveComputers -and $InactiveComputers.Count -gt 0) {
                $Global:DataGridResults.ItemsSource = $InactiveComputers
                Write-ADReportLog -Message "Inactive computers loaded. $($InactiveComputers.Count) result(s) found." -Type Info
                Update-ResultCounters -Results $InactiveComputers
            } else {
                $Global:DataGridResults.ItemsSource = $null
                Write-ADReportLog -Message "No inactive computers found." -Type Warning
            }
        } catch {
            $ErrorMessage = "Error loading inactive computers: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
        }
    })

    $ButtonExportCSV.add_Click({
        Write-ADReportLog -Message "Preparing CSV export..." -Type Info
        if ($null -eq $Global:DataGridResults.ItemsSource -or $Global:DataGridResults.Items.Count -eq 0) {
            Write-ADReportLog -Message "No data available for export." -Type Warning
            [System.Windows.Forms.MessageBox]::Show("Es sind keine Daten zum Exportieren vorhanden.", "Hinweis", "OK", "Information")
            return
        }

        $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $SaveFileDialog.Filter = "CSV (Comma delimited) (*.csv)|*.csv"
        $SaveFileDialog.Title = "Save CSV file as"
        $SaveFileDialog.InitialDirectory = [Environment]::GetFolderPath("MyDocuments")

        if ($SaveFileDialog.ShowDialog() -eq "OK") {
            $FilePath = $SaveFileDialog.FileName
            try {
                $Global:DataGridResults.ItemsSource | Export-Csv -Path $FilePath -NoTypeInformation -Encoding UTF8 -Delimiter ';'
                $Global:TextBlockStatus.Text = "Daten erfolgreich nach $FilePath exportiert."
                [System.Windows.Forms.MessageBox]::Show("Data successfully exported!", "CSV Export", "OK", "Information")
            } catch {
                $ErrorMessage = "Error in CSV export: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                [System.Windows.Forms.MessageBox]::Show("Fehler beim Exportieren der Daten: $($_.Exception.Message)", "Export Fehler", "OK", "Error")
            }
        } else {
            Write-ADReportLog -Message "CSV export canceled by user." -Type Info
        }
    })

    $ButtonExportHTML.add_Click({
        Write-ADReportLog -Message "Preparing HTML export..." -Type Info
        if ($null -eq $Global:DataGridResults.ItemsSource -or $Global:DataGridResults.Items.Count -eq 0) {
            $Global:TextBlockStatus.Text = "No data available for export."
            [System.Windows.Forms.MessageBox]::Show("Es sind keine Daten zum Exportieren vorhanden.", "Hinweis", "OK", "Information")
            return
        }

        $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $SaveFileDialog.Filter = "HTML File (*.html;*.htm)|*.html;*.htm"
        $SaveFileDialog.Title = "Save HTML file as"
        $SaveFileDialog.InitialDirectory = [Environment]::GetFolderPath("MyDocuments")

        if ($SaveFileDialog.ShowDialog() -eq "OK") {
            $FilePath = $SaveFileDialog.FileName
            try {
                # Create a more attractive HTML header
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
                $ReportTitle = "Active Directory Report - Created on $DateTimeNow"

                $Global:DataGridResults.ItemsSource | ConvertTo-Html -Head $HtmlHead -Body "<h1>$ReportTitle</h1>" | Out-File -FilePath $FilePath -Encoding UTF8
                $Global:TextBlockStatus.Text = "Daten erfolgreich nach $FilePath exportiert."
                [System.Windows.Forms.MessageBox]::Show("Data successfully exported!", "HTML Export", "OK", "Information")
            } catch {
                $ErrorMessage = "Error in HTML export: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                [System.Windows.Forms.MessageBox]::Show("Fehler beim Exportieren der Daten: $($_.Exception.Message)", "Export Fehler", "OK", "Error")
            }
        } else {
            Write-ADReportLog -Message "HTML export canceled by user." -Type Info
        }
    })

    # Fenster anzeigen
    $null = $Window.ShowDialog()
}

# --- Visualisierungsfunktionen ---
# Funktion zum Erstellen von Mini-Donut-Charts in Canvas-Elementen
Function New-MiniDonutChart {
    param (
        [Parameter(Mandatory=$true)]
        [System.Windows.Controls.Canvas]$Canvas,
        
        [Parameter(Mandatory=$true)]
        [hashtable]$Data,  # z.B. @{"Enabled" = 25; "Disabled" = 5}
        
        [Parameter(Mandatory=$false)]
        [hashtable]$Colors = @{"Enabled" = "#4CAF50"; "Disabled" = "#F44336"; "Default" = "#2196F3"}
    )
    
    # Canvas leeren
    $Canvas.Children.Clear()
    
    # Dimensionen berechnen
    $canvasWidth = $Canvas.Width
    if (-not $canvasWidth -or $canvasWidth -eq 0) {
        $canvasWidth = 80  # Standardbreite
    }
    $canvasHeight = $Canvas.Height
    if (-not $canvasHeight -or $canvasHeight -eq 0) {
        $canvasHeight = 25  # Standardhöhe
    }
    
    $center = New-Object System.Windows.Point($canvasWidth/2, $canvasHeight/2)
    $radius = [Math]::Min($canvasWidth, $canvasHeight) / 2 - 2
    
    # Gesamtsumme berechnen
    $total = 0
    foreach ($value in $Data.Values) {
        $total += $value
    }
    
    if ($total -eq 0) {
        # Wenn keine Daten vorhanden sind, zeige einen grauen Kreis
        $ellipse = New-Object System.Windows.Shapes.Ellipse
        $ellipse.Width = $radius * 2
        $ellipse.Height = $radius * 2
        $ellipse.Fill = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(200, 200, 200))
        $Canvas.Children.Add($ellipse) | Out-Null
        [System.Windows.Controls.Canvas]::SetLeft($ellipse, $center.X - $radius)
        [System.Windows.Controls.Canvas]::SetTop($ellipse, $center.Y - $radius)
        
        # Textanzeige für "No Data"
        $textBlock = New-Object System.Windows.Controls.TextBlock
        $textBlock.Text = "$total"
        $textBlock.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Colors]::White)
        $textBlock.FontSize = 8
        $textBlock.FontWeight = [System.Windows.FontWeights]::Bold
        $textBlock.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Center
        $textBlock.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
        $Canvas.Children.Add($textBlock) | Out-Null
        [System.Windows.Controls.Canvas]::SetLeft($textBlock, $center.X - 8)
        [System.Windows.Controls.Canvas]::SetTop($textBlock, $center.Y - 8)
        
        return
    }
    
    # Startwinkel für die Segmente
    $startAngle = 0
    
    # Für jeden Datenpunkt ein Segment zeichnen
    foreach ($key in $Data.Keys) {
        $value = $Data[$key]
        $sweepAngle = 360 * ($value / $total)
        
        # Segment nur zeichnen, wenn der Wert > 0 ist
        if ($value -gt 0) {
            # Farbe ermitteln
            $colorBrush = $null
            if ($Colors.ContainsKey($key)) {
                $colorCode = $Colors[$key]
                $colorBrush = New-Object System.Windows.Media.SolidColorBrush((ConvertFrom-HexColor $colorCode))
            }
            else {
                # Standardfarbe verwenden
                $colorCode = $Colors["Default"]
                $colorBrush = New-Object System.Windows.Media.SolidColorBrush((ConvertFrom-HexColor $colorCode))
            }
            
            # Segment zeichnen
            $segment = New-PieSegment -CenterX $center.X -CenterY $center.Y -Radius $radius -StartAngle $startAngle -SweepAngle $sweepAngle -Fill $colorBrush
            $Canvas.Children.Add($segment) | Out-Null
            
            # Startwinkel für das nächste Segment aktualisieren
            $startAngle += $sweepAngle
        }
    }
    
    # Inneres Loch für Donut-Effekt
    $innerRadius = $radius * 0.6
    $innerEllipse = New-Object System.Windows.Shapes.Ellipse
    $innerEllipse.Width = $innerRadius * 2
    $innerEllipse.Height = $innerRadius * 2
    $innerEllipse.Fill = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Colors]::Transparent)
    $innerEllipse.Fill = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(0, 99, 177)) # #0063B1
    $Canvas.Children.Add($innerEllipse) | Out-Null
    [System.Windows.Controls.Canvas]::SetLeft($innerEllipse, $center.X - $innerRadius)
    [System.Windows.Controls.Canvas]::SetTop($innerEllipse, $center.Y - $innerRadius)
    
    # Textanzeige für Gesamtzahl
    $textBlock = New-Object System.Windows.Controls.TextBlock
    $textBlock.Text = "$total"
    $textBlock.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Colors]::White)
    $textBlock.FontSize = 8
    $textBlock.FontWeight = [System.Windows.FontWeights]::Bold
    $textBlock.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Center
    $textBlock.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
    $Canvas.Children.Add($textBlock) | Out-Null
    [System.Windows.Controls.Canvas]::SetLeft($textBlock, $center.X - 8)
    [System.Windows.Controls.Canvas]::SetTop($textBlock, $center.Y - 8)
}

# Hilfsfunktion zum Konvertieren von Hex-Farbcodes in Color-Objekte
Function ConvertFrom-HexColor {
    param(
        [Parameter(Mandatory=$true)]
        [string]$HexColor
    )
    
    $HexColor = $HexColor.TrimStart('#')
    if ($HexColor.Length -eq 6) {
        $r = [Convert]::ToInt32($HexColor.Substring(0, 2), 16)
        $g = [Convert]::ToInt32($HexColor.Substring(2, 2), 16)
        $b = [Convert]::ToInt32($HexColor.Substring(4, 2), 16)
        return [System.Windows.Media.Color]::FromRgb($r, $g, $b)
    }
    return [System.Windows.Media.Colors]::Black  # Fallback
}

# Hilfsfunktion zum Erstellen eines Kreissegments
Function New-PieSegment {
    param (
        [double]$CenterX,
        [double]$CenterY,
        [double]$Radius,
        [double]$StartAngle,
        [double]$SweepAngle,
        [System.Windows.Media.Brush]$Fill
    )
    
    # Winkel in Radianten umrechnen
    $startAngleRad = $StartAngle * [Math]::PI / 180
    $endAngleRad = ($StartAngle + $SweepAngle) * [Math]::PI / 180
    
    # Groß-Kreis Flag (true wenn SweepAngle > 180°)
    $isLargeArc = $SweepAngle -gt 180
    
    # Punkte für den Pfad berechnen
    $startPoint = New-Object System.Windows.Point($CenterX + $Radius * [Math]::Cos($startAngleRad), $CenterY + $Radius * [Math]::Sin($startAngleRad))
    $endPoint = New-Object System.Windows.Point($CenterX + $Radius * [Math]::Cos($endAngleRad), $CenterY + $Radius * [Math]::Sin($endAngleRad))
    
    # PathFigure erstellen
    $pathFigure = New-Object System.Windows.Media.PathFigure
    $pathFigure.StartPoint = $CenterX, $CenterY
    $pathFigure.IsClosed = $true
    
    # Line zum Startpunkt
    $lineSegment1 = New-Object System.Windows.Media.LineSegment
    $lineSegment1.Point = $startPoint
    $pathFigure.Segments.Add($lineSegment1) | Out-Null
    
    # Kreisbogen zum Endpunkt
    $arcSegment = New-Object System.Windows.Media.ArcSegment
    $arcSegment.Point = $endPoint
    $arcSegment.Size = New-Object System.Windows.Size($Radius, $Radius)
    $arcSegment.IsLargeArc = $isLargeArc
    $arcSegment.SweepDirection = [System.Windows.Media.SweepDirection]::Clockwise
    $pathFigure.Segments.Add($arcSegment) | Out-Null
    
    # Line zurück zum Zentrum
    $lineSegment2 = New-Object System.Windows.Media.LineSegment
    $lineSegment2.Point = $CenterX, $CenterY
    $pathFigure.Segments.Add($lineSegment2) | Out-Null
    
    # PathGeometry erstellen
    $pathGeometry = New-Object System.Windows.Media.PathGeometry
    $pathGeometry.Figures.Add($pathFigure) | Out-Null
    
    # Path erstellen
    $path = New-Object System.Windows.Shapes.Path
    $path.Data = $pathGeometry
    $path.Fill = $Fill
    
    return $path
}

# Funktion zum Aktualisieren der Ergebnisvisualisierung
Function Update-ResultVisualization {
    param (
        [Parameter(Mandatory=$true)]
        [array]$Results
    )
    
    try {
        # Analysieren der Suchergebnisse basierend auf dem Objekttyp
        # Standardwerte für die Visualisierungen
        $userData = @{"Total" = 0}
        $computerData = @{"Total" = 0}
        $groupData = @{"Total" = 0}
        
        # Farben für die Diagramme
        $userColors = @{
            "Enabled" = "#4CAF50";
            "Disabled" = "#F44336";
            "Locked" = "#FFC107";
            "Admin" = "#2196F3";
            "Total" = "#607D8B";
            "Default" = "#9E9E9E"
        }
        
        $computerColors = @{
            "Enabled" = "#4CAF50";
            "Disabled" = "#F44336";
            "Inactive" = "#FFC107";
            "Total" = "#607D8B";
            "Default" = "#9E9E9E"
        }
        
        $groupColors = @{
            "Security" = "#FF9800";
            "Distribution" = "#9C27B0";
            "Total" = "#607D8B";
            "Default" = "#9E9E9E"
        }
        
        # Objekttyp bestimmen basierend auf Eigenschaften des ersten Elements, falls vorhanden
        if ($Results.Count -gt 0) {
            $firstObject = $Results[0]
            
            # Überprüfen auf Benutzer, Computer oder Gruppe basierend auf ObjectClass Eigenschaft
            if ($firstObject.PSObject.Properties.Name -contains "ObjectClass") {
                $objectClass = $firstObject.ObjectClass
                
                if ($objectClass -eq "user") {
                    # Benutzerdetails analysieren
                    $enabledUsers = $Results | Where-Object {$_.Enabled -eq $true}
                    $disabledUsers = $Results | Where-Object {$_.Enabled -eq $false}
                    
                    $userData = @{
                        "Total" = $Results.Count
                    }
                    
                    if ($enabledUsers.Count -gt 0) {
                        $userData["Enabled"] = $enabledUsers.Count
                    }
                    
                    if ($disabledUsers.Count -gt 0) {
                        $userData["Disabled"] = $disabledUsers.Count
                    }
                    
                    # Wenn die Eigenschaften LockedOut oder AdminCount vorhanden sind, diese auch berücksichtigen
                    if ($firstObject.PSObject.Properties.Name -contains "LockedOut") {
                        $lockedUsers = $Results | Where-Object {$_.LockedOut -eq $true}
                        if ($lockedUsers.Count -gt 0) {
                            $userData["Locked"] = $lockedUsers.Count
                        }
                    }
                    
                    if ($firstObject.PSObject.Properties.Name -contains "AdminCount") {
                        $adminUsers = $Results | Where-Object {$_.AdminCount -eq 1}
                        if ($adminUsers.Count -gt 0) {
                            $userData["Admin"] = $adminUsers.Count
                        }
                    }
                }
                elseif ($objectClass -eq "computer") {
                    # Computerdetails analysieren
                    $enabledComputers = $Results | Where-Object {$_.Enabled -eq $true}
                    $disabledComputers = $Results | Where-Object {$_.Enabled -eq $false}
                    
                    $computerData = @{
                        "Total" = $Results.Count
                    }
                    
                    if ($enabledComputers.Count -gt 0) {
                        $computerData["Enabled"] = $enabledComputers.Count
                    }
                    
                    if ($disabledComputers.Count -gt 0) {
                        $computerData["Disabled"] = $disabledComputers.Count
                    }
                    
                    # Wenn InactiveDays vorhanden ist, auch inaktive Computer anzeigen
                    if ($firstObject.PSObject.Properties.Name -contains "InactiveDays") {
                        $inactiveComputers = $Results | Where-Object {$_.InactiveDays -gt 30}
                        if ($inactiveComputers.Count -gt 0) {
                            $computerData["Inactive"] = $inactiveComputers.Count
                        }
                    }
                }
                elseif ($objectClass -eq "group") {
                    # Gruppendetails analysieren
                    if ($firstObject.PSObject.Properties.Name -contains "GroupCategory") {
                        $securityGroups = $Results | Where-Object {$_.GroupCategory -eq "Security"}
                        $distributionGroups = $Results | Where-Object {$_.GroupCategory -eq "Distribution"}
                        
                        $groupData = @{
                            "Total" = $Results.Count
                        }
                        
                        if ($securityGroups.Count -gt 0) {
                            $groupData["Security"] = $securityGroups.Count
                        }
                        
                        if ($distributionGroups.Count -gt 0) {
                            $groupData["Distribution"] = $distributionGroups.Count
                        }
                    }
                    else {
                        # Wenn GroupCategory nicht verfügbar ist, nur Gesamtzahl anzeigen
                        $groupData = @{
                            "Total" = $Results.Count
                        }
                    }
                }
                else {
                    # Für andere Objekttypen nur Gesamtzahl anzeigen
                    $userData = @{"Total" = $Results.Count}
                }
            }
            else {
                # Wenn keine ObjectClass-Eigenschaft vorhanden ist, zeige nur die Gesamtzahl an
                $userData = @{"Total" = $Results.Count}
            }
        }
        
        # Visualisierungen erstellen
        New-MiniDonutChart -Canvas $Global:UserStatusCanvas -Data $userData -Colors $userColors
        New-MiniDonutChart -Canvas $Global:ComputerStatusCanvas -Data $computerData -Colors $computerColors
        New-MiniDonutChart -Canvas $Global:GroupsStatusCanvas -Data $groupData -Colors $groupColors
        
        Write-ADReportLog -Message "Result visualization updated successfully." -Type Info -Terminal
    }
    catch {
        Write-ADReportLog -Message "Error updating result visualization: $($_.Exception.Message)" -Type Error
    }
}

# Platzhalter-Funktion für initiale Visualisierung beim Start
Function Initialize-ResultVisualization {
    [CmdletBinding()]
    param()
    
    # Erstellt eine leere Visualisierung, bis Suchergebnisse vorliegen
    $emptyData = @{"Total" = 0}
    
    # Visualisierungen mit leeren Daten initialisieren
    New-MiniDonutChart -Canvas $Global:UserStatusCanvas -Data $emptyData
    New-MiniDonutChart -Canvas $Global:ComputerStatusCanvas -Data $emptyData
    New-MiniDonutChart -Canvas $Global:GroupsStatusCanvas -Data $emptyData
}

# --- Hauptlogik ---
Function Start-ADReportGUI {
    # Bereinige alle alten globalen UI-Variablen vor dem Start
    $UiVariables = @("Window", "ComboBoxFilterAttribute", "TextBoxFilterValue", "ListBoxSelectAttributes",
                  "ButtonQueryAD", "ButtonQuickAllUsers", "ButtonQuickDisabledUsers", "ButtonQuickLockedUsers",
                  "ButtonQuickGroups", "ButtonQuickSecurityGroups", "ButtonQuickDistributionGroups",
                  "ButtonQuickNeverExpire", "ButtonQuickInactiveUsers", "ButtonQuickAdminUsers",
                  "ButtonQuickComputers", "ButtonQuickInactiveComputers", "DataGridResults",
                  "TextBlockStatus", "ButtonExportCSV", "ButtonExportHTML",
                  "ResultCountGrid", "UserCountText", "ComputerCountText", "GroupCountText")
    
    foreach ($var in $UiVariables) {
        Remove-Variable -Name $var -Scope Global -ErrorAction SilentlyContinue
    }

    # Ruft Initialize-ADReportForm auf, welche die UI lädt, Elemente zuweist und füllt.
    Initialize-ADReportForm -XAMLContent $Global:XAML
}

# --- Skriptstart ---
Start-ADReportGUI