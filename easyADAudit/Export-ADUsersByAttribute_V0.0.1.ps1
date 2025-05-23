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
    param($Type,$Values)
    $or = switch($Type){
        'OU'        { $Values|ForEach{"(distinguishedName=*$_*)"} }
        Default     { $Values|ForEach{"($Type=$($_))"} }
    }
    "(&(|$($or -join '')))"
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

#region ─── Haupt ───────────────────────────────────────────────────────────────
try{
    Write-Verbose 'Skriptstart'
    Write-Debug "Parameter: FilterType=$FilterType, OutputType=$OutputType, Silent=$Silent, IncludeDisabled=$IncludeDisabled"
    
    Write-Debug "Lade ActiveDirectory-Modul..."
    Ensure-ADModule
    Write-Debug "ActiveDirectory-Modul erfolgreich geladen"

    Write-Debug "Wähle Filtertyp..."
    $FilterType = Choose-FilterType $FilterType
    Write-Debug "Gewählter Filtertyp: $FilterType"

    Write-Debug "Wähle Ausgabetyp..."
    $OutputType = Choose-OutputType $OutputType
    Write-Debug "Gewählter Ausgabetyp: $OutputType"

    Write-Debug "Wähle Werte..."
    $Values = Select-Values $FilterType
    Write-Debug "Gewählte Werte: $($Values -join ', ')"
    
    if(-not $Values){
        Write-Warning 'Keine Werte gewählt – Abbruch.'
        return
    }

    Write-Debug "Wähle Export-Attribute..."
    $selectedAttributes = Select-ExportAttributes
    Write-Debug "Gewählte Attribute: $($selectedAttributes -join ', ')"
    
    if (-not $selectedAttributes) {
        Write-Warning 'Keine Attribute ausgewählt – Abbruch.'
        return
    }

    $coreProcessingAttributes = @('Name', 'Enabled', 'Company')
    $attributesToFetch = ($selectedAttributes + $coreProcessingAttributes) | Select-Object -Unique
    Write-Debug "Attribute für Get-ADUser: $($attributesToFetch -join ', ')"

    $RequiredAttribute = $null
    $SortBy = $null
    $ExcludePattern = $null

    if (-not $Silent) {
        Write-Host "\`nSoll ein Pflichtattribut für den Export festgelegt werden?" -ForegroundColor Cyan
        Write-Host "Verfügbare Attribute: $($selectedAttributes -join ', ')"
        $UserInput = Read-Host "Pflichtattribut (Enter für keine Einschränkung)"
        if ($UserInput -eq "-join") {
            Write-Warning "Ungültiger Wert '-join' für Pflichtattribut erhalten. Wird ignoriert."
            Write-Host "Warnung: Der eingegebene Wert '-join' für das Pflichtattribut ist ungültig und wird ignoriert." -ForegroundColor Yellow
            $RequiredAttribute = $null
        } else {
            $RequiredAttribute = $UserInput
        }
    }

    if (-not $Silent) {
        Write-Host "\`nNach welchem Attribut soll sortiert werden?" -ForegroundColor Cyan
        Write-Host "Verfügbare Attribute: $($selectedAttributes -join ', ')"
        $UserInput = Read-Host "Attribut (Enter für keine Sortierung)"
        if ($UserInput -eq "-join") {
            Write-Warning "Ungültiger Wert '-join' für Sortierung erhalten. Wird ignoriert."
            Write-Host "Warnung: Der eingegebene Wert '-join' für die Sortierung ist ungültig und wird ignoriert." -ForegroundColor Yellow
            $SortBy = $null
        } else {
            $SortBy = $UserInput
        }
    }

    if (-not $Silent) {
        Write-Host "\`nWelche Benutzer sollen ausgeschlossen werden?" -ForegroundColor Cyan
        Write-Host "Beispiel: 'test' würde alle Benutzer mit 'test' im Namen ausschließen"
        $UserInput = Read-Host "Suchmuster (Enter für keine Ausnahmen)"
        if ($UserInput -eq "-join") {
            Write-Warning "Ungültiger Wert '-join' für Ausschlussmuster erhalten. Wird ignoriert."
            Write-Host "Warnung: Der eingegebene Wert '-join' für das Ausschlussmuster ist ungültig und wird ignoriert." -ForegroundColor Yellow
            $ExcludePattern = $null
        } else {
            $ExcludePattern = $UserInput
        }
    }

    Write-Debug "Erstelle LDAP-Filter..."
    $ldap = Build-LdapFilter $FilterType $Values
    Write-Debug "LDAP-Filter: $ldap"

    Write-Debug "Starte AD-Abfrage..."
    $users = [System.Collections.ArrayList]::new()
    $c=0
    Write-Progress -Activity 'AD-Abfrage' -Status '0 Benutzer' -PercentComplete 0
    foreach($u in Get-ADUser -LDAPFilter $ldap -Properties $attributesToFetch){
        if(!$IncludeDisabled -and !$u.Enabled){continue}
        $users.Add($u) | Out-Null
        $c++
        if($c%200 -eq 0){
            Write-Progress -Activity 'AD-Abfrage' -Status "$c Benutzer" -PercentComplete 0
            Write-Debug "Bisher gefundene Benutzer: $c"
        }
    }
    Write-Progress -Activity 'AD-Abfrage' -Completed
    Write-Debug "AD-Abfrage abgeschlossen. Gefundene Benutzer: $($users.Count)"
    
    if($users.Count -eq 0){
        Write-Warning 'Keine Benutzer gefunden.'
        return
    }

    Write-Debug "Starte Export..."
    $csv=$html=$null
    switch($OutputType){
        'CSV' { 
            Write-Debug "Exportiere als CSV..."
            $csv=Export-Data $users 'CSV' $OutputPath $selectedAttributes $SortBy $ExcludePattern $RequiredAttribute
        }
        'HTML'{ 
            Write-Debug "Exportiere als HTML..."
            $html=Export-Data $users 'HTML' $OutputPath $selectedAttributes $SortBy $ExcludePattern $RequiredAttribute
        }
        'Both'{ 
            Write-Debug "Exportiere als CSV und HTML..."
            $csv=Export-Data $users 'CSV' $OutputPath $selectedAttributes $SortBy $ExcludePattern $RequiredAttribute
            $html=Export-Data $users 'HTML' $OutputPath $selectedAttributes $SortBy $ExcludePattern $RequiredAttribute
        }
    }
    Write-Verbose "Export abgeschlossen."
    if($csv){Write-Host "CSV : $csv" -ForegroundColor Green}
    if($html){Write-Host "HTML: $html" -ForegroundColor Green}
}
catch{
    Write-Error "Fehler aufgetreten: $_"
    throw # Fehler weiterwerfen, damit das Skript mit Fehlercode beendet wird
}
finally{
    Write-Debug "Skript beendet."
}
#endregion Haupt 