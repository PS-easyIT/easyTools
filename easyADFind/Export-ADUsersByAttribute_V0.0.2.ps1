[CmdletBinding()]
param(
    [ValidateSet('OU','Department','Company')]
    [string]$FilterType,

    [ValidateSet('HTML','CSV','Both')]
    [string]$OutputType,

    [string]$OutputPath, # Standardwert wird unten dynamisch gesetzt

    [switch]$IncludeDisabled,
    [switch]$Silent
    # Parameter LogPath, LogLevel, UseTranscript und DebugMode entfernt
)

# Dynamically determine script directory for default paths
$ScriptDirectory = $null
try {
    # Try to get script path via MyInvocation (more reliable in some PS 5.1 contexts)
    $InvocationInfo = (Get-Variable MyInvocation -Scope 1 -ErrorAction Stop).Value
    if ($InvocationInfo.MyCommand.Path) {
        $ScriptDirectory = Split-Path -Path $InvocationInfo.MyCommand.Path -Parent
    }
}
catch {
    # MyInvocation might not be available or suitable (e.g., in runspace, interactive debugging)
    # Fall back to PSScriptRoot if MyInvocation failed
    if (-not ([string]::IsNullOrEmpty($PSScriptRoot))) {
        $ScriptDirectory = $PSScriptRoot
    }
}

# Resolve OutputPath default if not provided by user
if (-not $PSBoundParameters.ContainsKey('OutputPath')) {
    if (-not ([string]::IsNullOrEmpty($ScriptDirectory))) {
        $OutputPath = Join-Path $ScriptDirectory 'ADUsersExport'
    } else {
        $OutputPath = Join-Path -Path (Get-Location).Path -ChildPath 'ADUsersExport' # Explicitly CWD
        Write-Warning "WARNUNG: Skript-Stammverzeichnis konnte nicht zuverlässig ermittelt werden. OutputPath ('$OutputPath') wird im aktuellen Arbeitsverzeichnis erstellt: '$((Get-Location).Path)'"
    }
}

# Prüfen der PowerShell Ausführungsrichtlinie
# Diese Prüfung erfolgt nach den Parametern und vor der Hauptlogik.
$currentExecutionPolicy = Get-ExecutionPolicy
if ($currentExecutionPolicy -eq [Microsoft.PowerShell.ExecutionPolicy]::Restricted) {
    Write-Host ""
    Write-Host "---------------------------------------------------------------------------------" -ForegroundColor Yellow
    Write-Host " WICHTIGER HINWEIS ZUR AUSFÜHRUNGSRICHTLINIE" -ForegroundColor Yellow
    Write-Host "---------------------------------------------------------------------------------" -ForegroundColor Yellow
    Write-Host ""
    Write-Host " Die aktuelle PowerShell Ausführungsrichtlinie ist auf 'Restricted' gesetzt." -ForegroundColor Red
    Write-Host " Dies bedeutet, dass das Ausführen von PowerShell-Skripten auf diesem System"
    Write-Host " generell deaktiviert ist."
    Write-Host ""
    Write-Host " UM DIESES SKRIPT AUSZUFÜHREN, MUSS DIE RICHTLINIE GEÄNDERT WERDEN." -ForegroundColor White
    Write-Host ""
    Write-Host " Mögliche Lösungen:"
    Write-Host " 1. Für die aktuelle PowerShell-Sitzung (keine Administratorrechte erforderlich, wirkt nur temporär):"
    Write-Host "    Führen Sie in Ihrer PowerShell-Konsole folgenden Befehl aus und starten Sie"
    Write-Host "    das Skript danach erneut in derselben Konsole:"
    Write-Host ""
    Write-Host "    Set-ExecutionPolicy RemoteSigned -Scope Process -Force" -ForegroundColor Cyan
    Write-Host ""
    Write-Host " 2. Für den aktuellen Benutzer (Administratorrechte sind ggf. einmalig erforderlich,"
    Write-Host "    wirkt dauerhaft für Ihren Benutzer, sofern nicht durch GPO überschrieben):"
    Write-Host "    Öffnen Sie PowerShell als Administrator und führen Sie folgenden Befehl aus:"
    Write-Host ""
    Write-Host "    Set-ExecutionPolicy RemoteSigned -Scope CurrentUser -Force" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "    Hinweis: Wenn die Ausführungsrichtlinie durch eine zentrale Gruppenrichtlinie (GPO)"
    Write-Host "    Ihres Unternehmens festgelegt wurde, können diese Befehle möglicherweise nicht"
    Write-Host "    ausreichend sein oder fehlschlagen. Wenden Sie sich in diesem Fall bitte an Ihre IT-Abteilung."
    Write-Host ""
    Write-Host " Weitere Informationen zu PowerShell Ausführungsrichtlinien finden Sie hier:"
    Write-Host " Get-Help about_Execution_Policies" -ForegroundColor Green
    Write-Host " oder online: https://go.microsoft.com/fwlink/?LinkID=135170"
    Write-Host ""
    Write-Host " Das Skript kann unter der aktuellen 'Restricted'-Richtlinie nicht ausgeführt werden und wird nun beendet." -ForegroundColor Red
    Write-Host "---------------------------------------------------------------------------------" -ForegroundColor Yellow
    Write-Host ""
    
    if ($Host.UI.SupportsInteractiveHost) {
        try {
            Read-Host -Prompt "Drücken Sie die Eingabetaste, um das Skript zu beenden"
        } catch {
            Start-Sleep -Seconds 7
        }
    } else {
        Start-Sleep -Seconds 2
    }
    exit 1
}

#region ─── Helfer ───────────────────────────────────────────────────────────────
function Ensure-ADModule {
    if (-not (Get-Module -Name ActiveDirectory -ListAvailable)) { 
        throw 'ActiveDirectory‑Modul nicht installiert.' 
    }
    Import-Module -Name ActiveDirectory -ErrorAction Stop
    Write-Debug 'ActiveDirectory‑Modul geladen.'
}

function Choose-FilterType { param($Current) if ($Silent -or $Current){return $Current}
    Write-Host "Filtertyp wählen:\`n1) OU\`n2) Department\`n3) Company" -ForegroundColor Cyan
    switch (Read-Host 'Ihre Wahl'){ 1{'OU'} 2{'Department'} 3{'Company'} default{Choose-FilterType} }
}

function Choose-OutputType { 
    param(
        [Parameter()]
        [string]$Current
    )
    
    if ($Silent -or $Current) { 
        return $Current 
    }
    
    Write-Host "Exportformat wählen:\`n1) HTML\`n2) CSV\`n3) Both" -ForegroundColor Cyan
    switch (Read-Host 'Ihre Wahl') { 
        1 { 'HTML' }
        2 { 'CSV' }
        3 { 'Both' }
        default { Choose-OutputType }
    }
}

function Select-Values {
    param($Type)
    switch ($Type){
        'OU'        { Get-ADOrganizationalUnit -Filter * | Sort DistinguishedName | Select -Expand DistinguishedName | Out-GridView -Title 'OU auswählen' -PassThru }
        'Department'{ Get-ADUser -Filter * -Properties Department | Where Department | Select -Expand Department -Unique | Sort | Out-GridView -Title 'Department auswählen' -PassThru }
        'Company'   { Get-ADUser -Filter * -Properties Company    | Where Company    | Select -Expand Company    -Unique | Sort | Out-GridView -Title 'Company auswählen' -PassThru }
    }
}

function Build-LdapFilter {
    param(
        $Type,
        $Values,
        [System.Nullable[bool]]$FilterByEnabledState = $false, # Ob der Enabled-Status im Filter berücksichtigt werden soll
        [System.Nullable[bool]]$UsersShouldBeEnabled = $true   # Wenn FilterByEnabledState true ist, gibt dies an, ob Benutzer aktiviert oder deaktiviert sein sollen
    )
    $orClauses = switch($Type){
        'OU'        { $Values | ForEach-Object {"(distinguishedName=*$_*)"} }
        Default     { $Values | ForEach-Object {"($Type=$($_))"} }
    }
    $mainFilterClause = "(|$($orClauses -join ''))"

    $finalFilterClauses = @($mainFilterClause)

    if ($FilterByEnabledState -eq $true) {
        if ($UsersShouldBeEnabled -eq $true) { # Nur aktivierte Benutzer
            $finalFilterClauses += "(!(userAccountControl:1.2.840.113556.1.4.803:=2))"
        } else { # Nur deaktivierte Benutzer (selten für Vorschau, aber möglich)
            $finalFilterClauses += "(userAccountControl:1.2.840.113556.1.4.803:=2)"
        }
    }
    
    return "(&$($finalFilterClauses -join ''))"
}

function Get-SelectableADAttributes {
    $defaultAttributes = @(
        'SamAccountName',
        'Name',
        'Enabled',
        'mail',
        'Department',
        'Company',
        'DistinguishedName',
        'DisplayName',
        'GivenName',
        'Surname',
        'Title',
        'Office',
        'TelephoneNumber',
        'MobilePhone',
        'Manager',
        'LastLogonDate',
        'Created',
        'Modified',
        'Description',
        'EmployeeID',
        'EmployeeNumber',
        'EmployeeType',
        'Division',
        'OfficePhone',
        'HomePhone',
        'Pager',
        'Fax',
        'StreetAddress',
        'City',
        'State',
        'PostalCode',
        'Country',
        'Notes'
    )
    return $defaultAttributes
}

function Select-ExportAttributes {
    param($Current)
    if ($Silent -or $Current) { return $Current }
    
    $availableAttributes = Get-SelectableADAttributes
    $selectedAttributes = $availableAttributes | 
        ForEach-Object {
            [PSCustomObject]@{
                Attribute = $_
                Selected = $true # Default to selected
            }
        } | 
        Out-GridView -Title 'AD-Attribute für Export auswählen' -PassThru
    
    if ($selectedAttributes) {
        return $selectedAttributes | Where-Object Selected | Select-Object -ExpandProperty Attribute
    }
    return $null
}

function Export-Data {
    param(
        $Collection,
        [ValidateSet('CSV','HTML')]
        [string]$Type,
        [string]$Path,
        [string[]]$Attributes,
        [string]$SortBy,
        [string]$ExcludePattern,
        [string]$RequiredAttribute
    )
    try {
        Write-Debug "Starte Export ($Type) nach $Path"
        
        $Path = [System.IO.Path]::GetFullPath($Path)
        
        if (-not (Test-Path $Path)) {
            Write-Debug "Erstelle Export-Verzeichnis: $Path"
            New-Item $Path -ItemType Directory -Force | Out-Null
        }
        
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $file = Join-Path $Path "ADUsers_$timestamp.$($Type.ToLower())"
        Write-Debug "Exportiere nach: $file"
        
        $exportData = $Collection | Select-Object $Attributes
        
        if ($RequiredAttribute) {
            Write-Debug "Filtere nach Pflichtattribut: $RequiredAttribute"
            $exportData = $exportData | Where-Object { $_.$RequiredAttribute }
        }
        
        if ($SortBy) {
            Write-Debug "Sortiere nach: $SortBy"
            $exportData = $exportData | Sort-Object $SortBy
        }
        
        if ($ExcludePattern) {
            Write-Debug "Filtere Benutzer mit Pattern: $ExcludePattern"
            $exportData = $exportData | Where-Object { $_.Name -notlike "*$ExcludePattern*" }
        }
        
        switch ($Type) {
            'CSV' { 
                Write-Debug "Exportiere CSV mit $($exportData.Count) Datensätzen"
                $exportData | Export-Csv -Path $file -NoType -Encoding UTF8 -ErrorAction Stop
            }
            'HTML' { 
                Write-Debug "Exportiere HTML mit $($exportData.Count) Datensätzen"
                
                $totalUsers = $exportData.Count
                $companyStats = $exportData | Group-Object Company | Sort-Object Count -Descending | 
                    ForEach-Object { 
                        [PSCustomObject]@{
                            Company = if ([string]::IsNullOrEmpty($_.Name)) { '(Leer)' } else { $_.Name }
                            Count = $_.Count
                            Percentage = if ($totalUsers -gt 0) { [math]::Round(($_.Count / $totalUsers) * 100, 1) } else { 0 }
                        }
                    }
                
                $companyStatRows = $companyStats | ForEach-Object {
                    "<tr>
                        <td>$($_.Company)</td>
                        <td style='text-align: right;'>$($_.Count)</td>
                        <td style='text-align: right;'>$($_.Percentage)%</td>
                    </tr>"
                }
                $companyStatsTableContent = $companyStatRows -join "\`n"

                $statsHtml = @"
<div style="margin: 20px 0; padding: 15px; background-color: #f8f9fa; border-radius: 5px;">
    <h2>Statistik</h2>
    <p>Gesamtanzahl Benutzer: <strong>$totalUsers</strong></p>
    <h3>Verteilung nach Unternehmen:</h3>
    <table style="width: 100%; margin-top: 10px;">
        <tr>
            <th style="text-align: left;">Unternehmen</th>
            <th style="text-align: right;">Anzahl</th>
            <th style="text-align: right;">Prozent</th>
        </tr>
        $companyStatsTableContent
    </table>
</div>
"@
                
                $css = @'
<style>
    body { font-family: Segoe UI, Arial; margin: 20px; }
    table { border-collapse: collapse; width: 100%; margin-top: 20px; }
    th, td { padding: 8px; border: 1px solid #ddd; }
    th { background: #f0f0f0; }
    tr:nth-child(even) { background-color: #f9f9f9; }
    tr:hover { background-color: #f5f5f5; }
    .enabled { color: green; }
    .disabled { color: red; }
</style>
'@
                $tableHtml = $exportData | ConvertTo-Html -Title 'AD User Export' -PreContent $css | Out-String
                
                $html = @"
<!DOCTYPE html>
<html>
<head>
    <title>AD User Export</title>
    $css
</head>
<body>
    <h1>AD User Export</h1>
    <p>Exportiert am: $(Get-Date -Format "dd.MM.yyyy HH:mm:ss")</p>
    $statsHtml
    $tableHtml
    <footer style="margin-top: 50px; padding: 20px; text-align: center; border-top: 1px solid #ddd;">
        <p>© $(Get-Date -Format "yyyy") EasyADFinder - AD User Export</p>
        <p style="font-size: 0.8em; color: #666;">Version 0.0.5 | Andreas Hepp</p>
    </footer>
</body>
</html>
"@
                $html | Out-File -FilePath $file -Encoding UTF8 -ErrorAction Stop
            }
        }
        
        if (Test-Path $file) {
            Write-Verbose "Export erfolgreich: $file"
            return $file
        } else {
            throw "Export-Datei wurde nicht erstellt: $file"
        }
    }
    catch {
        Write-Error "Fehler beim Export ($Type): $_"
        throw
    }
}
#endregion Helfer

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms # Für FolderBrowserDialog

# Globale Variablen für das Hauptfenster und XAML-Elemente
$Global:MainWindow = $null
$Global:XamlControls = @{}
$Global:SelectedFilterValuesFromListBox = @() # Hält die Strings der ausgewählten Werte aus der ListBox
$Global:SelectedExportAttributes = $null # Wird die vom Benutzer ausgewählten Attribute halten
$Global:UserCountPreviewJob = $null # Hält den aktuellen Zähl-Job

# Pfad zur XAML-Datei
# Stelle sicher, dass $PSScriptRoot korrekt ermittelt wird, bevor diese Zeile ausgeführt wird.
# Wenn $PSScriptRoot leer ist, wird die XAML-Datei im aktuellen Arbeitsverzeichnis gesucht.
$xamlPath = ""
if ($PSScriptRoot) {
    $xamlPath = Join-Path $PSScriptRoot "MainWindow.xaml"
} else {
    # Fallback, falls PSScriptRoot nicht verfügbar ist (z.B. direkt in ISE ausgeführt ohne Speichern)
    $xamlPath = Join-Path (Get-Location).Path "MainWindow.xaml"
    Write-Warning "WARNUNG: PSScriptRoot ist nicht gesetzt. MainWindow.xaml wird im aktuellen Verzeichnis gesucht: $(Get-Location). Speichern Sie das Skript für eine zuverlässige Pfadermittlung."
}

function Load-XamlFile {
    param (
        [string]$XamlFilePath
    )
    try {
        [xml]$xaml = Get-Content -Path $XamlFilePath -ErrorAction Stop
        $reader = (New-Object System.Xml.XmlNodeReader $xaml)
        $Global:MainWindow = [Windows.Markup.XamlReader]::Load($reader)

        # Steuerelemente aus XAML per Name sammeln
        $xaml.SelectNodes("//*[@Name]") | ForEach-Object {
            $Global:XamlControls[$_.Name] = $Global:MainWindow.FindName($_.Name)
        }

        Write-Debug "XAML-Datei '$XamlFilePath' erfolgreich geladen."
        return $true
    }
    catch {
        Write-Error "Fehler beim Laden der XAML-Datei '$XamlFilePath': $_"
        # Fallback-Log oder detailliertere Fehlermeldung für den Benutzer
        # Hier könnte eine MessageBox angezeigt werden, wenn PresentationFramework geladen ist.
        # Zum Beispiel: [System.Windows.MessageBox]::Show("Fehler beim Laden der GUI: $($_.Exception.Message)", "Kritischer Fehler", "OK", "Error")
        return $false
    }
}

# Hauptlogik des Skripts, gekapselt in einer Funktion, die von der GUI aufgerufen wird
function Start-ExportProcess {
    param(
        # Diese Parameter werden von den GUI-Handler-Funktionen übergeben
        [string]$GuiFilterType,
        [string]$GuiOutputType,
        [string]$GuiOutputPath,
        [bool]$GuiIncludeDisabled,
        [string[]]$GuiSelectedValues,
        [string[]]$GuiSelectedAttributes,
        [string]$GuiRequiredAttribute,
        [string]$GuiSortBy,
        [string]$GuiExcludePattern
    )
    try{
        $Global:XamlControls.TextBlockStatus.Text = "Starte Exportprozess... Validierung der Eingaben."
        $Global:XamlControls.ProgressBarStatus.IsIndeterminate = $true
        Write-Verbose 'Exportprozess gestartet via GUI'
        Write-Debug "Parameter von GUI: FilterType=$GuiFilterType, OutputType=$GuiOutputType, OutputPath=$GuiOutputPath, IncludeDisabled=$GuiIncludeDisabled"
        Write-Debug "Gewählte Werte: $($GuiSelectedValues -join ', ')"
        Write-Debug "Gewählte Attribute: $($GuiSelectedAttributes -join ', ')"
        Write-Debug "Pflichtattribut: $GuiRequiredAttribute, Sortierung: $GuiSortBy, Ausschluss: $GuiExcludePattern"

        $Global:XamlControls.TextBlockStatus.Text = "Prüfe Active Directory Modul..."
        Ensure-ADModule
        
        # Validierung hier erneut, auch wenn schon in Start-ExportProcessGui, für den Fall direkten Aufrufs
        if(-not $GuiSelectedValues){
            $msg = 'Keine Werte für den Filter ausgewählt – Abbruch.'
            Write-Warning $msg
            $Global:XamlControls.TextBlockStatus.Text = $msg
            [System.Windows.MessageBox]::Show($msg, "Validierungsfehler", "OK", "Warning")
            return
        }
        
        if (-not $GuiSelectedAttributes) {
            $msg = 'Keine Attribute für den Export ausgewählt – Abbruch.'
            Write-Warning $msg
            $Global:XamlControls.TextBlockStatus.Text = $msg
            [System.Windows.MessageBox]::Show($msg, "Validierungsfehler", "OK", "Warning")
            return
        }

        $coreProcessingAttributes = @('Name', 'Enabled', 'Company') # Für interne Logik benötigt
        $attributesToFetch = ($GuiSelectedAttributes + $coreProcessingAttributes) | Select-Object -Unique
        Write-Debug "Attribute für Get-ADUser: $($attributesToFetch -join ', ')"

        $ldap = Build-LdapFilter $GuiFilterType $GuiSelectedValues -FilterByEnabledState (-not $GuiIncludeDisabled) -UsersShouldBeEnabled $true
        Write-Debug "LDAP-Filter für Abfrage (inkl. Enabled-Status wenn nötig): $ldap"

        $Global:XamlControls.TextBlockStatus.Text = "Starte AD-Abfrage für Benutzer..."
        # Optional: $Global:XamlControls.UserCountLabel.Content = "0 Benutzer gefunden" # Falls separates Label vorhanden

        $users = [System.Collections.ArrayList]::new()
        $c=0
        
        # Da Get-ADUser potenziell lange laufen kann, hier eine Indeterminate ProgressBar
        # Bei sehr großen Umgebungen wäre ein Background-Job mit Fortschritts-Updates an die GUI besser.
        # Für den Moment: Direkte Iteration mit gelegentlichen Debug-Ausgaben.

        # Wir verwenden den LDAP-Filter, den Build-LdapFilter erzeugt, der bereits den Enabled-Status berücksichtigt,
        # WENN $GuiIncludeDisabled $false ist. Wenn $GuiIncludeDisabled $true ist, wird der Enabled-Status im Filter ignoriert.
        $allUsersFromAd = Get-ADUser -LDAPFilter $ldap -Properties $attributesToFetch -ErrorAction Stop

        if ($GuiIncludeDisabled) {
            # Wenn deaktivierte Benutzer eingeschlossen werden sollen, wurden bereits alle durch den LDAP-Filter geholt.
            $users.AddRange($allUsersFromAd)
            $c = $users.Count
            Write-Debug "AD-Abfrage (inkl. Deaktivierte): $($c) Benutzer gefunden."
        } else {
            # Wenn nur aktivierte Benutzer exportiert werden sollen, filtert der LDAP-Filter bereits vor.
            # $allUsersFromAd enthält also nur aktivierte Benutzer.
            $users.AddRange($allUsersFromAd)
            $c = $users.Count
            Write-Debug "AD-Abfrage (nur Aktivierte): $($c) Benutzer gefunden."
        }

        # Alte Schleife zur manuellen Filterung von `Enabled` ist nicht mehr nötig,
        # da Build-LdapFilter dies nun handhabt.
        # foreach($u in Get-ADUser -LDAPFilter $ldap -Properties $attributesToFetch){\n        #     if(-not $GuiIncludeDisabled -and !$u.Enabled){continue} # Diese Logik ist jetzt in Build-LdapFilter\n        #     $users.Add($u) | Out-Null\n        #     $c++\n        #     if($c%50 -eq 0){ \n        #         Write-Debug \"Bisher gefundene Benutzer: $c\"\n        #         # GUI Update (muss im UI-Thread erfolgen, falls asynchron)\n        #         # $Global:MainWindow.Dispatcher.Invoke([action]{\n        #         #    $Global:XamlControls.UserCountLabel.Content = \"$c Benutzer gefunden\"\n        #         # })\n        #     }\n        # }\n

        # $Global:XamlControls.UserCountLabel.Content = \"$($users.Count) Benutzer gefunden\" # Falls separates Label vorhanden
        $Global:XamlControls.TextBlockStatus.Text = "AD-Abfrage abgeschlossen. $($users.Count) Benutzer gefunden. Starte Export..."
        Write-Debug "AD-Abfrage abgeschlossen. Gefundene Benutzer: $($users.Count)"
        
        if($users.Count -eq 0){
            $msg = 'Keine passenden Benutzer für den Export gefunden.'
            Write-Warning $msg
            $Global:XamlControls.TextBlockStatus.Text = $msg
            [System.Windows.MessageBox]::Show($msg, "Keine Ergebnisse", "OK", "Information")
            return
        }

        $csvPath = $null
        $htmlPath = $null
        $exportSuccess = $false

        switch($GuiOutputType){
            'CSV' { 
                $Global:XamlControls.TextBlockStatus.Text = "Exportiere als CSV... ($($users.Count) Benutzer)"
                Write-Debug "Exportiere als CSV..."
                $csvPath = Export-Data $users 'CSV' $GuiOutputPath $GuiSelectedAttributes $GuiSortBy $GuiExcludePattern $GuiRequiredAttribute
                if ($csvPath) { $exportSuccess = $true }
            }
            'HTML'{ 
                $Global:XamlControls.TextBlockStatus.Text = "Exportiere als HTML... ($($users.Count) Benutzer)"
                Write-Debug "Exportiere als HTML..."
                $htmlPath = Export-Data $users 'HTML' $GuiOutputPath $GuiSelectedAttributes $GuiSortBy $GuiExcludePattern $GuiRequiredAttribute
                if ($htmlPath) { $exportSuccess = $true }
            }
            'Both'{ 
                $Global:XamlControls.TextBlockStatus.Text = "Exportiere als CSV... ($($users.Count) Benutzer)"
                Write-Debug "Exportiere als CSV und HTML..."
                $csvPath = Export-Data $users 'CSV' $GuiOutputPath $GuiSelectedAttributes $GuiSortBy $GuiExcludePattern $GuiRequiredAttribute
                
                $Global:XamlControls.TextBlockStatus.Text = "Exportiere als HTML... ($($users.Count) Benutzer)"
                $htmlPath = Export-Data $users 'HTML' $GuiOutputPath $GuiSelectedAttributes $GuiSortBy $GuiExcludePattern $GuiRequiredAttribute
                if ($csvPath -or $htmlPath) { $exportSuccess = $true } # Erfolg, wenn mindestens einer klappt
            }
        }
        Write-Verbose "Export abgeschlossen."
        $message = "Export abgeschlossen.`n"
        if($csvPath){ $message += "CSV erstellt: $csvPath`n" }
        if($htmlPath){ $message += "HTML erstellt: $htmlPath" }
        
        if ($exportSuccess) {
            $Global:XamlControls.TextBlockStatus.Text = "Export erfolgreich abgeschlossen."
            [System.Windows.MessageBox]::Show($message, "Erfolg", "OK", "Information")
        } else {
            # Fehler wurden bereits in Export-Data behandelt und geloggt, hier nur eine generelle Meldung.
            # Die genauere Fehlermeldung sollte bereits in der Konsole/Log stehen.
            $msg = "Export fehlgeschlagen. Bitte prüfen Sie die Logs/Konsolenausgaben für Details."
            $Global:XamlControls.TextBlockStatus.Text = $msg
            [System.Windows.MessageBox]::Show($msg, "Exportfehler", "OK", "Error")
        }
    }
    catch{
        $errorMessage = "Fehler im Exportprozess: $($_.Exception.Message)"
        Write-Error $errorMessage
        $Global:XamlControls.TextBlockStatus.Text = "Fehler: $errorMessage"
        try {
             [System.Windows.MessageBox]::Show($errorMessage, "Fehler im Exportprozess", "OK", "Error")
        } catch {
            Write-Warning "Konnte MessageBox für Fehler nicht anzeigen: $($_.Exception.Message)"
        }
    }
    finally{
        Write-Debug "Exportprozess GUI-Funktion beendet."
        if ($Global:XamlControls.ContainsKey('ProgressBarStatus')) {
            $Global:XamlControls.ProgressBarStatus.IsIndeterminate = $false
        }
    }
}

# Async Funktion zur Aktualisierung der Benutzeranzahl-Vorschau
function Update-UserCountPreviewAsync {
    try {
        # Alten Job abbrechen, falls vorhanden und noch läuft
        if ($Global:UserCountPreviewJob -and $Global:UserCountPreviewJob.State -eq 'Running') {
            Write-Debug "Bestehender UserCountPreviewJob wird gestoppt."
            Stop-Job -Job $Global:UserCountPreviewJob | Out-Null
            Remove-Job -Job $Global:UserCountPreviewJob -Force | Out-Null # Aufräumen
            # Wichtig: Auch den Event-Handler deregistrieren, wenn der Job gestoppt wird!
            Get-EventSubscriber -SourceIdentifier "UserCountJobChanged" | Unregister-Event -Force | Out-Null
        }

        $Global:MainWindow.Dispatcher.Invoke([action]{
            $Global:XamlControls.TextBlockUserCountPreview.Text = "Zähle Benutzer..."
            $Global:XamlControls.ProgressBarStatus.IsIndeterminate = $true
            $Global:XamlControls.TextBlockStatus.Text = "Aktualisiere Benutzeranzahl-Vorschau..."
        })

        $filterType = $Global:XamlControls.ComboBoxFilterType.SelectedItem
        $selectedValues = $Global:SelectedFilterValuesFromListBox
        $includeDisabledUsersInCount = $Global:XamlControls.CheckBoxIncludeDisabled.IsChecked -eq $true

        if (-not $filterType -or -not $selectedValues) {
            $Global:MainWindow.Dispatcher.Invoke([action]{
                $Global:XamlControls.TextBlockUserCountPreview.Text = "(Filter unvollständig)"
                $Global:XamlControls.ProgressBarStatus.IsIndeterminate = $false
                $Global:XamlControls.TextBlockStatus.Text = "Filterkriterien für Vorschau unvollständig."
            })
            return
        }

        Ensure-ADModule # Stellt sicher, dass das Modul geladen ist

        # LDAP-Filter für die Zählung erstellen
        # Wenn $includeDisabledUsersInCount $true ist, wollen wir keinen Filter auf Enabled/Disabled setzen (alle Benutzer)
        # Wenn $includeDisabledUsersInCount $false ist, wollen wir nur aktivierte Benutzer ($UsersShouldBeEnabled = $true)
        $applyEnabledFilterForCount = (-not $includeDisabledUsersInCount)
        $ldapFilterForCount = Build-LdapFilter -Type $filterType -Values $selectedValues -FilterByEnabledState $applyEnabledFilterForCount -UsersShouldBeEnabled $true

        Write-Debug "Update-UserCountPreviewAsync: LDAP Filter für Zählung = $ldapFilterForCount"

        $scriptBlock = {
            param($LdapFilterParam)
            try {
                Import-Module ActiveDirectory -ErrorAction Stop
                # Zählt Benutzer mit minimalen abgerufenen Daten
                $count = (Get-ADUser -LDAPFilter $LdapFilterParam -Properties 'PrimaryGroupID' -ErrorAction Stop).Count
                return $count
            }
            catch {
                Write-Warning "Fehler im Hintergrundjob (UserCountPreview): $($_.Exception.Message)"
                return -1 # Fehlerindikator
            }
        }
        
        $Global:UserCountPreviewJob = Start-Job -ScriptBlock $scriptBlock -ArgumentList $ldapFilterForCount
        
        Register-ObjectEvent -InputObject $Global:UserCountPreviewJob -EventName StateChanged -SourceIdentifier "UserCountJobChanged" -Action {
            param($Sender, $EventArgs)
            try {
                $job = $Sender
                if ($job.State -in ('Completed', 'Failed', 'Stopped')) {
                    $countResult = Receive-Job -Job $job -Keep
                    Remove-Job -Job $job -Force | Out-Null # Job aufräumen
                    Unregister-Event -SourceIdentifier "UserCountJobChanged" # Event-Registrierung aufräumen

                    $Global:MainWindow.Dispatcher.Invoke([action]{
                        if ($job.State -eq 'Failed' -or $countResult -is [array] -or $countResult -eq -1) { # Fehler oder ungültiges Ergebnis
                            $Global:XamlControls.TextBlockUserCountPreview.Text = "Fehler bei Zählung"
                            $errorMsg = "Fehler bei der Benutzerzählung-Vorschau."
                            if ($job.ChildJobs[0].Error) {
                                $errorMsg += " Details: " + ($job.ChildJobs[0].Error | Out-String)
                            }
                            $Global:XamlControls.TextBlockStatus.Text = $errorMsg
                        } elseif ($countResult -eq $null) {
                             $Global:XamlControls.TextBlockUserCountPreview.Text = "0 Benutzer (Vorschau)" # oder "Keine Daten"
                        } else {
                            $Global:XamlControls.TextBlockUserCountPreview.Text = "$($countResult) Benutzer (Vorschau)"
                            $Global:XamlControls.TextBlockStatus.Text = "Benutzeranzahl-Vorschau aktualisiert."
                        }
                        $Global:XamlControls.ProgressBarStatus.IsIndeterminate = $false
                    })
                    $Global:UserCountPreviewJob = $null # Job-Referenz löschen
                }
            }
            catch {
                Write-Error "Fehler im UserCountJob StateChanged Event Handler: $($_.Exception.Message)"
                 $Global:MainWindow.Dispatcher.Invoke([action]{
                    $Global:XamlControls.TextBlockUserCountPreview.Text = "Fehler (Event Zählung)"
                    $Global:XamlControls.ProgressBarStatus.IsIndeterminate = $false
                })
            }
        } | Out-Null
    }
    catch {
        $errorMessage = "Fehler in Update-UserCountPreviewAsync Hauptfunktion: $($_.Exception.Message)"
        Write-Error $errorMessage
        $Global:MainWindow.Dispatcher.Invoke([action]{
            $Global:XamlControls.TextBlockUserCountPreview.Text = "Fehler bei Zählung"
            $Global:XamlControls.TextBlockStatus.Text = $errorMessage
            $Global:XamlControls.ProgressBarStatus.IsIndeterminate = $false
        })
    }
}

# Initialisierung der GUI und Zuweisung von Event-Handlern
function Initialize-Gui {
    if (-not (Load-XamlFile -XamlFilePath $xamlPath)) {
        # Da PresentationFramework möglicherweise noch nicht geladen ist, wenn XAML fehlt,
        # verwenden wir eine einfache Write-Error und Exit.
        Write-Error "Kritischer Fehler: Die XAML-Datei '$xamlPath' konnte nicht geladen werden. Bitte stellen Sie sicher, dass die Datei existiert und korrekt formatiert ist."
        # Optional: Eine robustere Fehlerbehandlung für den Fall, dass $Host.UI.SupportsInteractiveHost false ist.
        if ($Host.UI.SupportsInteractiveHost) {
            Read-Host "Drücken Sie die Eingabetaste zum Beenden."
        }
        exit 1 # Beendet das Skript
    }

    # Standardwerte und ComboBoxen füllen
    try {
        # Pfad zuerst setzen, da er von Parametern abhängen kann, die vor GUI Initialisierung ausgewertet werden
        if ($Global:XamlControls.ContainsKey("TextBoxOutputPath")) {
            $Global:XamlControls.TextBoxOutputPath.Text = $OutputPath # $OutputPath wird am Skriptanfang gesetzt
        } else { Write-Warning "XAML Element 'TextBoxOutputPath' nicht gefunden." }

        if ($Global:XamlControls.ContainsKey("ComboBoxFilterType")) {
            $Global:XamlControls.ComboBoxFilterType.ItemsSource = @("OU", "Department", "Company")
            $Global:XamlControls.ComboBoxFilterType.SelectedIndex = 0 # Standard: OU
        } else { Write-Warning "XAML Element 'ComboBoxFilterType' nicht gefunden." }
        
        if ($Global:XamlControls.ContainsKey("ComboBoxOutputType")) {
            $Global:XamlControls.ComboBoxOutputType.ItemsSource = @("HTML", "CSV", "Both")
            $Global:XamlControls.ComboBoxOutputType.SelectedIndex = 0 # Standard: HTML
        } else { Write-Warning "XAML Element 'ComboBoxOutputType' nicht gefunden." }

        # Initial ausgewählte Attribute Text setzen
        if ($Global:XamlControls.ContainsKey("TextBlockSelectedAttributesCount")) {
            $Global:XamlControls.TextBlockSelectedAttributesCount.Text = "Ausgewählte Attribute: (Bitte auswählen)"
        } else { Write-Warning "XAML Element 'TextBlockSelectedAttributesCount' nicht gefunden." }

        # Standardtext für Statusleiste setzen
        if ($Global:XamlControls.ContainsKey("TextBlockStatus")) {
            $Global:XamlControls.TextBlockStatus.Text = "Bereit. Bitte Filterkriterien auswählen."
        } else { Write-Warning "XAML Element 'TextBlockStatus' nicht gefunden." }
        
        # Standardtext für Benutzeranzahl-Vorschau
        if ($Global:XamlControls.ContainsKey("TextBlockUserCountPreview")) {
            $Global:XamlControls.TextBlockUserCountPreview.Text = "Benutzer (Vorschau): -"
        } else { Write-Warning "XAML Element 'TextBlockUserCountPreview' nicht gefunden." }

    }
    catch {
        $errMsg = "Fehler beim Initialisieren der GUI-Steuerelemente: $($_.Exception.Message)"
        Write-Error $errMsg
        # Versuche, eine MessageBox anzuzeigen, wenn möglich
        try { [System.Windows.MessageBox]::Show($errMsg, "GUI Initialisierungsfehler", "OK", "Error") } catch {}
    }

    # Event-Handler zuweisen
    # Alle Zuweisungen in try-catch Blöcke, um Fehler bei fehlenden Elementen abzufangen
    try {
        if ($Global:XamlControls.ButtonBrowseOutputPath) {
            $Global:XamlControls.ButtonBrowseOutputPath.add_Click({
                try {
                    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
                    $folderBrowser.Description = "Wählen Sie einen Ordner für den Export"
                    $folderBrowser.ShowNewFolderButton = $true
                    if ($Global:XamlControls.TextBoxOutputPath.Text -and (Test-Path $Global:XamlControls.TextBoxOutputPath.Text)) {
                        $folderBrowser.SelectedPath = $Global:XamlControls.TextBoxOutputPath.Text
                    }
                    if ($folderBrowser.ShowDialog((New-Object System.Windows.Forms.NativeWindow)) -eq [System.Windows.Forms.DialogResult]::OK) {
                        $Global:XamlControls.TextBoxOutputPath.Text = $folderBrowser.SelectedPath
                        $Global:XamlControls.TextBlockStatus.Text = "Exportpfad ausgewählt: $($folderBrowser.SelectedPath)"
                    }
                }
                catch {
                    Write-Error "Fehler beim Anzeigen des Ordnerauswahldialogs: $_"
                    $Global:XamlControls.TextBlockStatus.Text = "Fehler (Ordnerdialog): $($_.Exception.Message)"
                    try { [System.Windows.MessageBox]::Show("Fehler beim Ordnerauswahldialog: $($_.Exception.Message)", "Dialogfehler", "OK", "Error") } catch {}
                }
            })
        } else { Write-Warning "XAML Element 'ButtonBrowseOutputPath' nicht gefunden." }

        if ($Global:XamlControls.ComboBoxFilterType) {
            $Global:XamlControls.ComboBoxFilterType.add_SelectionChanged({
                try {
                    $Global:XamlControls.ListBoxFilterValues.ItemsSource = @() 
                    $Global:XamlControls.TextBlockSelectedFilterValues.Text = "Ausgewählte Werte: (Keine)"
                    $Global:SelectedFilterValuesFromListBox = @()
                    Update-UserCountPreviewAsync 
                    $Global:XamlControls.TextBlockStatus.Text = "Filtertyp geändert. Bitte auf 'Werte anzeigen/aktualisieren' klicken."
                } 
                catch {
                    Write-Error "Fehler im ComboBoxFilterType.SelectionChanged: $_"
                    $Global:XamlControls.TextBlockStatus.Text = "Fehler (Filtertyp Auswahl): $($_.Exception.Message)"
                }
            })
        } else { Write-Warning "XAML Element 'ComboBoxFilterType' nicht gefunden." }

        if ($Global:XamlControls.ButtonPopulateFilterValues) {
            $Global:XamlControls.ButtonPopulateFilterValues.add_Click({
                try {
                    Populate-FilterValuesListBox -ListBoxElement $Global:XamlControls.ListBoxFilterValues
                } 
                catch {
                    Write-Error "Fehler im ButtonPopulateFilterValues.Click: $_"
                    $Global:XamlControls.TextBlockStatus.Text = "Fehler (Werte laden): $($_.Exception.Message)"
                }
            })
        } else { Write-Warning "XAML Element 'ButtonPopulateFilterValues' nicht gefunden." }

        if ($Global:XamlControls.ListBoxFilterValues) {
            $Global:XamlControls.ListBoxFilterValues.add_MouseLeftButtonUp({param($s,$e) 
                try {
                    Handle-ListBoxFilterValueClick -sender $s -e $e 
                }
                catch {
                    Write-Error "Fehler im ListBoxFilterValues.MouseLeftButtonUp: $_"
                    $Global:XamlControls.TextBlockStatus.Text = "Fehler (Filterwert Auswahl): $($_.Exception.Message)"
                }
            })
        } else { Write-Warning "XAML Element 'ListBoxFilterValues' nicht gefunden." }

        if ($Global:XamlControls.CheckBoxIncludeDisabled) {
            $Global:XamlControls.CheckBoxIncludeDisabled.add_Click({
                try {
                    # Stelle sicher, dass die globale Liste der ausgewählten Elemente aktualisiert wird
                    Get-SelectedItemsFromFilterValuesListBox -ListBoxElement $Global:XamlControls.ListBoxFilterValues
                    Update-UserCountPreviewAsync 
                }
                catch {
                    Write-Error "Fehler im CheckBoxIncludeDisabled.Click: $_"
                    $Global:XamlControls.TextBlockStatus.Text = "Fehler (Checkbox Deaktivierte): $($_.Exception.Message)"
                }
            })
        } else { Write-Warning "XAML Element 'CheckBoxIncludeDisabled' nicht gefunden." }

        if ($Global:XamlControls.ButtonSelectExportAttributes) {
            $Global:XamlControls.ButtonSelectExportAttributes.add_Click({
                try {
                    # Diese Funktion muss noch implementiert werden
                    Select-ExportAttributesGui
                }
                catch {
                    Write-Error "Fehler im ButtonSelectExportAttributes.Click: $_"
                    $Global:XamlControls.TextBlockStatus.Text = "Fehler (Attributauswahl): $($_.Exception.Message)"
                }
            })
        } else { Write-Warning "XAML Element 'ButtonSelectExportAttributes' nicht gefunden." }

        if ($Global:XamlControls.ButtonStartExport) {
            $Global:XamlControls.ButtonStartExport.add_Click({
                try {
                    # Diese Funktion wird die Parameter aus der GUI sammeln und Start-ExportProcess aufrufen
                    Start-ExportProcessGui
                }
                catch {
                    Write-Error "Fehler im ButtonStartExport.Click: $_"
                    $Global:XamlControls.TextBlockStatus.Text = "Fehler (Exportstart): $($_.Exception.Message)"
                }
            })
        } else { Write-Warning "XAML Element 'ButtonStartExport' nicht gefunden." }
        
        if ($Global:XamlControls.ButtonClose) {
            $Global:XamlControls.ButtonClose.add_Click({
                try {
                    $Global:MainWindow.Close()
                }
                catch {
                    Write-Error "Fehler im ButtonClose.Click: $_"
                    # Fallback, falls Hauptfenster nicht schließbar ist.
                }
            })
        } else { Write-Warning "XAML Element 'ButtonClose' nicht gefunden." }

    }
    catch {
        $errMsg = "Fehler beim Zuweisen der GUI Event-Handler: $($_.Exception.Message)"
        Write-Error $errMsg
        try { [System.Windows.MessageBox]::Show($errMsg, "GUI Event-Handler Fehler", "OK", "Error") } catch {}
    }


    # GUI anzeigen
    Write-Host "Starte GUI..." # Debug-Ausgabe für Konsole
    try {
        # Wichtig: Die ShowDialog()-Methode blockiert, bis das Fenster geschlossen wird.
        # Alle Initialisierungen müssen vorher abgeschlossen sein.
        $null = $Global:MainWindow.ShowDialog() # $null unterdrückt die Ausgabe des DialogResult (True/False)
    }
    catch {
        $errorMessage = "Kritischer Fehler beim Anzeigen der GUI: $($_.Exception.Message)"
        Write-Error $errorMessage
        # Versuche, eine MessageBox anzuzeigen, wenn PresentationFramework geladen wurde
        try { [System.Windows.MessageBox]::Show($errorMessage, "GUI Anzeigefehler", "OK", "Error") } catch {}
        # Skript beenden, wenn die GUI nicht angezeigt werden kann
        exit 1
    }
}

# Funktion zum Anzeigen des Attributauswahl-Dialogs
function Select-ExportAttributesGui {
    try {
        $dialogXamlPath = Join-Path $PSScriptRoot "AttributeSelectionDialog.xaml"
        if (-not (Test-Path $dialogXamlPath)) {
            [System.Windows.MessageBox]::Show("Die Dialogdatei 'AttributeSelectionDialog.xaml' wurde nicht gefunden.", "Fehler", "OK", "Error")
            return
        }

        [xml]$dialogXaml = Get-Content -Path $dialogXamlPath -ErrorAction Stop
        $reader = (New-Object System.Xml.XmlNodeReader $dialogXaml)
        $dialogWindow = [Windows.Markup.XamlReader]::Load($reader)
        $dialogWindow.Owner = $Global:MainWindow # Setzt das Hauptfenster als Besitzer

        # XAML-Elemente des Dialogs holen
        $xamlDialogControls = @{}
        $dialogXaml.SelectNodes("//*[@Name]") | ForEach-Object {
            $xamlDialogControls[$_.Name] = $dialogWindow.FindName($_.Name)
        }

        $availableAttributes = Get-SelectableADAttributes | Sort-Object
        $attributeObjects = $availableAttributes | ForEach-Object {
            [PSCustomObject]@{
                AttributeName = $_
                IsSelected = ($Global:SelectedExportAttributes -contains $_) # Vorselektieren
            }
        }
        
        # Ursprüngliche, ungefilterte Liste für die Filterung speichern
        $xamlDialogControls.ListBoxAttributes.Tag = $attributeObjects 
        $xamlDialogControls.ListBoxAttributes.ItemsSource = $attributeObjects

        # Event-Handler für die Filter-TextBox
        $xamlDialogControls.TextBoxAttributeFilter.add_TextChanged({
            param($sender, $e)
            try {
                $filterText = $sender.Text
                $originalList = $xamlDialogControls.ListBoxAttributes.Tag # Zugriff auf die gespeicherte Originalliste
                if ([string]::IsNullOrWhiteSpace($filterText)) {
                    $xamlDialogControls.ListBoxAttributes.ItemsSource = $originalList
                } else {
                    $filteredList = $originalList | Where-Object {$_.AttributeName -like "*$filterText*"}
                    $xamlDialogControls.ListBoxAttributes.ItemsSource = $filteredList
                }
            }
            catch {
                Write-Warning "Fehler beim Filtern der Attributliste: $_"
            }
        })

        # Event-Handler für OK-Button
        $xamlDialogControls.ButtonOK.add_Click({
            param($sender, $e)
            # Ausgewählte Attribute sammeln
            $selectedAttrs = $xamlDialogControls.ListBoxAttributes.ItemsSource | Where-Object {$_.IsSelected} | Select-Object -ExpandProperty AttributeName
            $Global:SelectedExportAttributes = $selectedAttrs # Globale Variable aktualisieren
            
            if ($Global:SelectedExportAttributes -and $Global:SelectedExportAttributes.Count -gt 0) {
                $Global:XamlControls.TextBlockSelectedAttributesCount.Text = "Ausgewählte Attribute: $($Global:SelectedExportAttributes.Count)"
                $Global:XamlControls.TextBlockStatus.Text = "$($Global:SelectedExportAttributes.Count) Attribute für den Export ausgewählt."
            } else {
                $Global:XamlControls.TextBlockSelectedAttributesCount.Text = "Ausgewählte Attribute: (Keine ausgewählt)"
                $Global:XamlControls.TextBlockStatus.Text = "Keine Attribute für den Export ausgewählt."
            }
            $dialogWindow.DialogResult = $true # Setzt das Ergebnis und schließt den Dialog
            $dialogWindow.Close()
        })

        # Event-Handler für Cancel-Button (schließt den Dialog)
        $xamlDialogControls.ButtonCancel.add_Click({
            $dialogWindow.DialogResult = $false
            $dialogWindow.Close()
        })

        # Dialog modal anzeigen
        $result = $dialogWindow.ShowDialog()

        # Hier könnte man $result auswerten, wenn nötig (true für OK, false/null für Cancel/Schließen)
        # Die Aktualisierung von $Global:SelectedExportAttributes erfolgt bereits im OK-Click-Handler.

    }
    catch {
        $errMsg = "Fehler in Select-ExportAttributesGui: $($_.Exception.Message)"
        Write-Error $errMsg
        $Global:XamlControls.TextBlockStatus.Text = "Fehler bei Attributauswahl: $errMsg"
        try {[System.Windows.MessageBox]::Show($errMsg, "Attributauswahl Fehler", "OK", "Error")} catch {}
    }
}

function Start-ExportProcessGui {
    # Diese Funktion sammelt alle Parameter aus den GUI-Elementen
    # und ruft dann die Kernfunktion Start-ExportProcess auf.
    try {
        $Global:XamlControls.TextBlockStatus.Text = "Sammle Exportparameter aus GUI..."
        $Global:XamlControls.ProgressBarStatus.IsIndeterminate = $true

        $guiFilterType = $Global:XamlControls.ComboBoxFilterType.SelectedItem
        $guiOutputType = $Global:XamlControls.ComboBoxOutputType.SelectedItem
        $guiOutputPath = $Global:XamlControls.TextBoxOutputPath.Text
        $guiIncludeDisabled = $Global:XamlControls.CheckBoxIncludeDisabled.IsChecked -eq $true # Explizit zu [bool] casten
        
        # Ausgewählte Werte aus der ListBox holen (stellt sicher, dass $Global:SelectedFilterValuesFromListBox aktuell ist)
        Get-SelectedItemsFromFilterValuesListBox -ListBoxElement $Global:XamlControls.ListBoxFilterValues
        $guiSelectedValues = $Global:SelectedFilterValuesFromListBox
        
        # Ausgewählte Attribute (sollten bereits in $Global:SelectedExportAttributes sein)
        $guiSelectedAttributes = $Global:SelectedExportAttributes
        
        $guiRequiredAttribute = $Global:XamlControls.TextBoxRequiredAttribute.Text
        $guiSortBy = $Global:XamlControls.TextBoxSortBy.Text
        $guiExcludePattern = $Global:XamlControls.TextBoxExcludePattern.Text

        # Validierungen
        if (-not $guiFilterType) {
            [System.Windows.MessageBox]::Show("Bitte wählen Sie einen Filtertyp aus.", "Validierungsfehler", "OK", "Warning")
            $Global:XamlControls.TextBlockStatus.Text = "Validierungsfehler: Filtertyp fehlt."
            $Global:XamlControls.ProgressBarStatus.IsIndeterminate = $false
            return
        }
        if (-not $guiSelectedValues -or $guiSelectedValues.Count -eq 0) {
            [System.Windows.MessageBox]::Show("Bitte wählen Sie mindestens einen Filterwert aus (über 'Werte anzeigen/aktualisieren').", "Validierungsfehler", "OK", "Warning")
            $Global:XamlControls.TextBlockStatus.Text = "Validierungsfehler: Filterwerte fehlen."
            $Global:XamlControls.ProgressBarStatus.IsIndeterminate = $false
            return
        }
        if (-not $guiSelectedAttributes -or $guiSelectedAttributes.Count -eq 0) {
            [System.Windows.MessageBox]::Show("Bitte wählen Sie mindestens ein Attribut für den Export aus (über 'Attribute auswählen...').", "Validierungsfehler", "OK", "Warning")
            $Global:XamlControls.TextBlockStatus.Text = "Validierungsfehler: Exportattribute fehlen."
            $Global:XamlControls.ProgressBarStatus.IsIndeterminate = $false
            return
        }
        if (-not $guiOutputPath -or -not (Test-Path (Split-Path $guiOutputPath -Parent))) {
             [System.Windows.MessageBox]::Show("Der übergeordnete Ordner des Exportpfads '$guiOutputPath' existiert nicht oder ist ungültig. Bitte wählen Sie einen gültigen Pfad.", "Validierungsfehler", "OK", "Warning")
            $Global:XamlControls.TextBlockStatus.Text = "Validierungsfehler: Exportpfad ungültig."
            $Global:XamlControls.ProgressBarStatus.IsIndeterminate = $false
            return
        }


        $Global:XamlControls.TextBlockStatus.Text = "Starte Exportprozess..."
        # Die eigentliche Exportfunktion aufrufen. Diese sollte für GUI-Feedback angepasst werden.
        Start-ExportProcess -GuiFilterType $guiFilterType `
                            -GuiOutputType $guiOutputType `
                            -GuiOutputPath $guiOutputPath `
                            -GuiIncludeDisabled $guiIncludeDisabled `
                            -GuiSelectedValues $guiSelectedValues `
                            -GuiSelectedAttributes $guiSelectedAttributes `
                            -GuiRequiredAttribute $guiRequiredAttribute `
                            -GuiSortBy $guiSortBy `
                            -GuiExcludePattern $guiExcludePattern
        
        # Feedback nach Abschluss (Erfolg/Fehler wird in Start-ExportProcess behandelt und dort via MessageBox angezeigt)
        # $Global:XamlControls.TextBlockStatus.Text = "Exportvorgang beendet." # Wird von Start-ExportProcess aktualisiert

    }
    catch {
        $errMsg = "Schwerwiegender Fehler in Start-ExportProcessGui: $($_.Exception.Message)"
        Write-Error $errMsg
        $Global:XamlControls.TextBlockStatus.Text = "Fehler: $errMsg"
        try {[System.Windows.MessageBox]::Show($errMsg, "Exportfehler", "OK", "Error")} catch {}
    }
    finally {
        $Global:XamlControls.ProgressBarStatus.IsIndeterminate = $false
    }
}


function Choose-FilterTypeFromGui {
    # Stattdessen: Wert aus $Global:XamlControls.ComboBoxFilterType.SelectedItem auslesen
    if ($Global:XamlControls.ContainsKey("ComboBoxFilterType")) {
        return $Global:XamlControls.ComboBoxFilterType.SelectedItem.ToString() # oder .Tag, je nach Befüllung
    }
    throw "ComboBoxFilterType nicht in GUI gefunden oder nicht initialisiert."
}

function Choose-OutputTypeFromGui { # Bleibt vorerst für Start-ExportProcess relevant
    # Liest den Wert aus dem entsprechenden GUI-Element
    if ($Global:XamlControls.ContainsKey("ComboBoxOutputType")) {
        return $Global:XamlControls.ComboBoxOutputType.SelectedItem.ToString() # oder .Tag
    }
    throw "ComboBoxOutputType nicht in GUI gefunden oder nicht initialisiert."
}

function Get-RawFilterValues {
    param($Type)
    # Diese Funktion liefert die Rohdaten für die ListBox, ohne Out-GridView
    Write-Debug "Get-RawFilterValues: Typ = $Type"
    try {
        switch ($Type) {
            'OU'         { Get-ADOrganizationalUnit -Filter * | Sort-Object DistinguishedName | Select-Object -ExpandProperty DistinguishedName }
            'Department' { Get-ADUser -Filter * -Properties Department | Where-Object {$_.Department} | Select-Object -ExpandProperty Department -Unique | Sort-Object }
            'Company'    { Get-ADUser -Filter * -Properties Company    | Where-Object {$_.Company}    | Select-Object -ExpandProperty Company    -Unique | Sort-Object }
            default      { Write-Warning "Unbekannter Filtertyp in Get-RawFilterValues: $Type"; return @() }
        }
    }
    catch {
        Write-Error "Fehler in Get-RawFilterValues beim Abrufen von '$Type': $_"
        [System.Windows.MessageBox]::Show("Fehler beim Abrufen der Filterwerte für '$Type': $($_.Exception.Message)", "Datenabruffehler", "OK", "Error")
        return @()
    }
}

function Populate-FilterValuesListBox {
    param ([System.Windows.Controls.ListBox]$ListBoxElement)
    try {
        $selectedFilterType = $Global:XamlControls.ComboBoxFilterType.SelectedItem
        if (-not $selectedFilterType) {
            $ListBoxElement.ItemsSource = @()
            $Global:XamlControls.TextBlockStatus.Text = "Bitte zuerst einen Filtertyp auswählen."
            return
        }

        $Global:XamlControls.TextBlockStatus.Text = "Lade Werte für Filtertyp '$selectedFilterType'..."
        $Global:XamlControls.ProgressBarStatus.IsIndeterminate = $true

        $rawValues = Get-RawFilterValues -Type $selectedFilterType
        
        $listBoxItems = foreach ($value in $rawValues) {
            [PSCustomObject]@{ DisplayName = $value; IsSelected = $false }
        }
        
        $ListBoxElement.ItemsSource = $listBoxItems
        $Global:XamlControls.TextBlockSelectedFilterValues.Text = "Ausgewählte Werte: (Keine)" # Zurücksetzen bei Neubefüllung
        $Global:SelectedFilterValuesFromListBox = @() # Zurücksetzen bei Neubefüllung
        Update-UserCountPreviewAsync # Vorschau aktualisieren, da sich die Auswahlmöglichkeiten geändert haben (potenziell keine Auswahl mehr)

        if ($listBoxItems.Count -eq 0) {
            $Global:XamlControls.TextBlockStatus.Text = "Keine Werte für Filtertyp '$selectedFilterType' gefunden."
        } else {
            $Global:XamlControls.TextBlockStatus.Text = "Werte für '$selectedFilterType' geladen. Bitte auswählen."
        }
    }
    catch {
        Write-Error "Fehler in Populate-FilterValuesListBox: $_"
        $Global:XamlControls.TextBlockStatus.Text = "Fehler beim Laden der Filterwerte: $($_.Exception.Message)"
        [System.Windows.MessageBox]::Show("Fehler beim Laden der Filterwerte-Liste: $($_.Exception.Message)", "Listenfehler", "OK", "Error")
    }
    finally {
        $Global:XamlControls.ProgressBarStatus.IsIndeterminate = $false
    }
}

function Get-SelectedItemsFromFilterValuesListBox {
    param ([System.Windows.Controls.ListBox]$ListBoxElement)
    $selectedItems = @()
    if ($ListBoxElement.ItemsSource) {
        foreach ($item in $ListBoxElement.ItemsSource) {
            if ($item.IsSelected) {
                $selectedItems += $item.DisplayName
            }
        }
    }
    $Global:SelectedFilterValuesFromListBox = $selectedItems # Aktualisiere die globale Variable
    $Global:XamlControls.TextBlockSelectedFilterValues.Text = "Ausgewählte Werte: $(if($selectedItems.Count -gt 0) {$selectedItems -join ', '} else {'(Keine)'})"
    return $selectedItems
}

# Event-Handler für die CheckBoxen innerhalb der ListBoxFilterValues (wenn ein Item geklickt wird)
# Dies erfordert, dass wir den Event Handler an jede CheckBox dynamisch binden, was komplexer ist.
# Eine Alternative ist ein Button "Auswahl übernehmen" oder die Auswertung beim Klick auf "Vorschau aktualisieren" / "Export starten".
# Für eine direktere Reaktion können wir den SelectionChanged der ListBox nutzen, obwohl der für CheckBoxen nicht ideal ist.
# Oder wir binden einen Click-Handler an die CheckBoxen, wenn sie generiert werden. Das ist der robusteste Ansatz.
# Für den Moment: Aktualisierung der Vorschau und der Textanzeige nach Neubefüllung und beim Klick auf CheckBoxIncludeDisabled.
# Der Benutzer muss nach Auswahl in der ListBox ggf. manuell eine Aktion auslösen, die die Vorschau aktualisiert, oder wir lösen es bei `ButtonPopulateFilterValues` aus.

# Wir fügen einen allgemeinen Event-Handler für Klicks auf die ListBox hinzu
# und versuchen, die CheckBox zu identifizieren.
function Handle-ListBoxFilterValueClick {
    # Diese Funktion wird aufgerufen, wenn die ListBox selbst oder ein Element darin geklickt wird.
    # Wir müssen prüfen, ob das OriginalSource eine CheckBox ist und dann deren Status umschalten und die Logik ausführen.
    param($sender, $e)

    try {
        $clickedElement = $e.OriginalSource
        if ($clickedElement -is [System.Windows.Controls.CheckBox]) {
            # Der IsChecked-Status der CheckBox wird durch die Bindung im XAML automatisch aktualisiert,
            # wenn IsSelected im gebundenen Objekt geändert wird.
            # Wenn der Benutzer direkt auf die CheckBox klickt, müssen wir sicherstellen,
            # dass das gebundene IsSelected-Property des Datenobjekts aktualisiert wird.
            # Da wir eine OneWay-Bindung von IsSelected zu CheckBox.IsChecked haben könnten (je nach XAML),
            # ist es sicherer, das DataContext des geklickten Elements zu holen und dessen Property zu setzen.
            
            $dataContext = $clickedElement.DataContext
            if ($dataContext -and $dataContext.PSObject.Properties["IsSelected"]) {
                 # Den Wert umschalten. Das DataTemplate {Binding IsSelected} sollte TwoWay sein oder hier manuell gesetzt werden.
                 # Im aktuellen XAML ({Binding IsSelected}) ist es standardmäßig OneWay von Source zu Target.
                 # Für eine Checkbox ist ein TwoWay Binding üblicher: IsChecked="{Binding IsSelected, Mode=TwoWay}"
                 # Da es nicht TwoWay ist, setzen wir es hier manuell, falls der Klick von der Checkbox kommt.
                 # Dies ist ein Workaround, wenn das XAML nicht Mode=TwoWay für die CheckBox verwendet.
                 # $dataContext.IsSelected = $clickedElement.IsChecked # Aktualisiert das Quellobjekt

                 # Da das XAML <CheckBox Content="{Binding DisplayName}" IsChecked="{Binding IsSelected}" Margin="2"/>
                 # verwendet (implizit OneWay für IsChecked, wenn nicht anders spezifiziert),
                 # und der Klick auf die CheckBox deren IsChecked-Zustand direkt ändert,
                 # müssen wir IsSelected im Objekt aktualisieren.
                 # Es ist besser, die `Get-SelectedItemsFromFilterValuesListBox` aufzurufen,
                 # die die `IsSelected` Properties basierend auf den `ItemsSource` der ListBox neu setzt.
                 # Der Klick auf die CheckBox hat deren Zustand geändert. `Get-SelectedItemsFromFilterValuesListBox`
                 # liest diesen Zustand und aktualisiert $Global:SelectedFilterValuesFromListBox.
            }
            
            # Nachdem der Klick verarbeitet wurde (und die CheckBox ihren Zustand geändert hat),
            # rufen wir die Funktionen auf, um die globale Liste und die Vorschau zu aktualisieren.
            Get-SelectedItemsFromFilterValuesListBox -ListBoxElement $Global:XamlControls.ListBoxFilterValues
            Update-UserCountPreviewAsync
        }
        # Falls auf das ListBoxItem selbst geklickt wird (nicht die CheckBox), könnten wir hier auch das IsSelected-Property des DataContext umschalten.
        # elseif ($clickedElement -is [System.Windows.Controls.ListBoxItem]) {
        #    $dataContext = $clickedElement.DataContext
        #    if ($dataContext -and $dataContext.PSObject.Properties["IsSelected"]) {
        #        $dataContext.IsSelected = -not $dataContext.IsSelected
        #        Get-SelectedItemsFromFilterValuesListBox -ListBoxElement $Global:XamlControls.ListBoxFilterValues
        #        Update-UserCountPreviewAsync
        #    }
        # }
    }
    catch {
        $errMsg = "Fehler in Handle-ListBoxFilterValueClick: $($_.Exception.Message)"
        Write-Error $errMsg
        $Global:XamlControls.TextBlockStatus.Text = "Fehler bei ListBox-Klick: $errMsg"
    }
}

function Select-ValuesFromGui {
    param($Type)
    # Diese Funktion wird NICHT MEHR VERWENDET, da die Auswahl direkt in der ListBox erfolgt.
    # Die Logik zum Abrufen der Daten ist jetzt in Get-RawFilterValues.
    Write-Warning "Select-ValuesFromGui ist veraltet und sollte nicht mehr aufgerufen werden."
    return $null
}

function Get-SelectableADAttributes { # Diese Funktion bleibt unverändert
    $defaultAttributes = @(
        'SamAccountName',
        'Name',
        'Enabled',
        'mail',
        'Department',
        'Company',
        'DistinguishedName',
        'DisplayName',
        'GivenName',
        'Surname',
        'Title',
        'Office',
        'TelephoneNumber',
        'MobilePhone',
        'Manager',
        'LastLogonDate',
        'Created',
        'Modified',
        'Description',
        'EmployeeID',
        'EmployeeNumber',
        'EmployeeType',
        'Division',
        'OfficePhone',
        'HomePhone',
        'Pager',
        'Fax',
        'StreetAddress',
        'City',
        'State',
        'PostalCode',
        'Country',
        'Notes'
    )
    return $defaultAttributes
}