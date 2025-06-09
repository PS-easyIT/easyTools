# ==============================================
# Skript: AD-Testuser mit Menü und Zufalls-Aktionen
# Version: 1.2
#
# ANWENDUNG:
# Das Skript kann direkt ausgeführt werden. Die Basis-OU (Distinguished Name)
# wird dann interaktiv abgefragt, falls sie nicht beim Aufruf als Parameter übergeben wurde.
#
# Beispielaufruf mit BaseOU-Parameter:
#   .\Create_75DummyUsers1OU.ps1 -BaseOU "OU=MeineTestBenutzer,DC=beispiel,DC=com"
#
# Ohne Parameter wird die BaseOU abgefragt:
#   .\Create_75DummyUsers1OU.ps1
#   > Bitte geben Sie den Distinguished Name der Basis-OU ein (z.B. OU=TestUsers,DC=contoso,DC=com):
#
# Weitere optionale Parameter (wie -DomainDNS, -DefaultPasswordPrefix, -NewUserCount)
# können ebenfalls beim Aufruf mitgegeben werden.
# ==============================================

param (
    [Parameter(Mandatory=$false)]
    # Tragen Sie hier Ihre Standard-BaseOU ein
    [string]$BaseOU = "OU=Mitarbeiter,OU=USER,DC=phinit,DC=de", 

    [Parameter(Mandatory=$false)]
    [string]$DomainDNS = (Get-ADDomain -ErrorAction SilentlyContinue).DNSRoot,

    [Parameter(Mandatory=$false)]
    [string]$DefaultPasswordPrefix = "P@ssw0rd!",

    [Parameter(Mandatory=$false)]
    [int]$NewUserCount = 75,   # Option 1

    [Parameter(Mandatory=$false)]
    [int]$AttrAssignCount = 20,   # Option 2

    [Parameter(Mandatory=$false)]
    [int]$DisableCount = 10,   # Option 3

    [Parameter(Mandatory=$false)]
    [int]$SecSettingCount = 5,    # Option 4

    [Parameter(Mandatory=$false)]
    [int]$CriticalCount = 5     # Option 5
)

# Globale Namenslisten
$firstNamesJapan        = @("Akira","Yuki","Hiro","Mai","Kenta","Sakura","Taro","Yuna","Ryo","Emi")
$lastNamesJapan         = @("Tanaka","Suzuki","Takahashi","Sato","Yamamoto","Nakamura","Kobayashi","Kato","Yoshida","Ito")
$firstNamesChina        = @("Wei","Fang","Li","Zhang","Jun","Xiao","Ming","Lei","Hua","Tao")
$lastNamesChina         = @("Wang","Li","Zhang","Liu","Chen","Yang","Huang","Zhao","Wu","Zhou")
$firstNamesEngland      = @("James","Mary","John","Patricia","Robert","Jennifer","Michael","Linda","William","Elizabeth")
$lastNamesEngland       = @("Smith","Jones","Taylor","Brown","Williams","Wilson","Davies","Evans","Thomas","Johnson")
$firstNamesDeutschland  = @("Max","Anna","Lukas","Laura","Tim","Julia","Leon","Sophie","Felix","Marie")
$lastNamesDeutschland   = @("Müller","Schmidt","Schneider","Fischer","Weber","Meyer","Wagner","Becker","Schulz","Hoffmann")
$firstNamesSpanien      = @("Alejandro","María","Juan","Lucía","Carlos","Carmen","Miguel","Ana","José","Laura")
$lastNamesSpanien       = @("García","Martínez","López","Sánchez","Pérez","Gómez","Martín","Jiménez","Ruiz","Hernández")
$firstNamesFrankreich   = @("Jean","Marie","Pierre","Sophie","Michel","Claire","François","Camille","Philippe","Chloé")
$lastNamesFrankreich    = @("Dupont","Martin","Bernard","Dubois","Thomas","Robert","Richard","Petit","Durand","Leroy")
$firstNamesItalien      = @("Marco","Giulia","Alessandro","Francesca","Luca","Martina","Andrea","Laura","Davide","Elena")
$lastNamesItalien       = @("Rossi","Russo","Ferrari","Esposito","Bianchi","Romano","Colombo","Ricci","Marino","Greco")
$firstNamesBrasilien    = @("João","Maria","Pedro","Ana","Lucas","Beatriz","Gabriel","Juliana","Rafael","Camila")
$lastNamesBrasilien     = @("Silva","Santos","Oliveira","Souza","Rodrigues","Almeida","Costa","Pereira","Lima","Gomes")
$firstNamesRussland     = @("Ivan","Olga","Dmitry","Anna","Sergey","Elena","Alexey","Natalia","Vladimir","Marina")
$lastNamesRussland      = @("Ivanov","Smirnov","Kuznetsov","Popov","Sokolov","Lebedev","Kozlov","Novikov","Morozov","Petrov")
$firstNamesIndien       = @("Rahul","Priya","Amit","Neha","Rohan","Anjali","Vikram","Sneha","Arjun","Pooja")
$lastNamesIndien        = @("Patel","Kumar","Shah","Singh","Sharma","Gupta","Mehta","Khan","Das","Rao")

$firstNamesAll = $firstNamesJapan + $firstNamesChina + $firstNamesEngland + $firstNamesDeutschland + `
                 $firstNamesSpanien + $firstNamesFrankreich + $firstNamesItalien + $firstNamesBrasilien + `
                 $firstNamesRussland + $firstNamesIndien
$lastNamesAll  = $lastNamesJapan  + $lastNamesChina  + $lastNamesEngland  + $lastNamesDeutschland + `
                 $lastNamesSpanien + $lastNamesFrankreich + $lastNamesItalien + $lastNamesBrasilien + `
                 $lastNamesRussland + $lastNamesIndien

$Global:NameLists = @{
    Japan = @{ First = $firstNamesJapan; Last = $lastNamesJapan }
    China = @{ First = $firstNamesChina; Last = $lastNamesChina }
    English = @{ First = $firstNamesEngland; Last = $lastNamesEngland }
    German = @{ First = $firstNamesDeutschland; Last = $lastNamesDeutschland }
    Spanish = @{ First = $firstNamesSpanien; Last = $lastNamesSpanien }
    French = @{ First = $firstNamesFrankreich; Last = $lastNamesFrankreich }
    Italian = @{ First = $firstNamesItalien; Last = $lastNamesItalien }
    Brazilian = @{ First = $firstNamesBrasilien; Last = $lastNamesBrasilien }
    Russian = @{ First = $firstNamesRussland; Last = $lastNamesRussland }
    Indian = @{ First = $firstNamesIndien; Last = $lastNamesIndien }
    All = @{ First = $firstNamesAll; Last = $lastNamesAll }
}

$Global:Departments    = @("IT","HR","Vertrieb","Support","Entwicklung")
$Global:Titles         = @("Administrator","Manager","Consultant","Mitarbeiter","Teamleiter")
$Global:CriticalGroups = @("Domain Admins","Enterprise Admins","Schema Admins","Account Operators","Backup Operators")

# --- Funktionen ---

function Get-CountrySpecificNames {
    param (
        [string]$CountryChoice
    )
    switch ($CountryChoice) {
        '1' { $countryKey = 'Japan';     $countryDisplay = 'Japan' } 
        '2' { $countryKey = 'China';     $countryDisplay = 'China' }
        '3' { $countryKey = 'English';   $countryDisplay = 'England' }
        '4' { $countryKey = 'German';    $countryDisplay = 'Deutschland' }
        '5' { $countryKey = 'Spanish';   $countryDisplay = 'Spanien' }
        '6' { $countryKey = 'French';    $countryDisplay = 'Frankreich' }
        '7' { $countryKey = 'Italian';   $countryDisplay = 'Italien' }
        '8' { $countryKey = 'Brazilian'; $countryDisplay = 'Brasilien' }
        '9' { $countryKey = 'Russian';   $countryDisplay = 'Russland' }
        '10' { $countryKey = 'Indian';   $countryDisplay = 'Indien' }
        '11' { $countryKey = 'All';      $countryDisplay = 'Alle (gemischt)' }
        default {
            Write-Warning "Ungültige Länderauswahl. Standard: Alle (gemischt) wird verwendet."
            $countryKey = 'All'; $countryDisplay = 'Alle (gemischt)'
        }
    }
    return @{ 
        FirstNames = $Global:NameLists[$countryKey].First;
        LastNames  = $Global:NameLists[$countryKey].Last;
        CountryName = $countryDisplay
    }
}

function Get-ChildOUs {
    param (
        [string]$ParentOUPath,
        [switch]$Recursive
    )
    $searchScope = if ($Recursive) { 'Subtree' } else { 'OneLevel' }
    try {
        $ous = Get-ADOrganizationalUnit -Filter * -SearchBase $ParentOUPath -SearchScope $searchScope -ErrorAction Stop | 
               Select-Object DistinguishedName, Name | Sort-Object Name
        return $ous
    }
    catch {
        Write-Warning "Fehler beim Auslesen der Kind-OUs unter '$ParentOUPath': $($_.Exception.Message)"
        return $null
    }
}

function New-ADTestUserInternal {
    param (
        [string]$FirstName,
        [string]$LastName,
        [string]$SamAccountNameBase,
        [string]$UserPrincipalNameDomain,
        [string]$PathOU,
        [string]$Password,
        [string]$Department,
        [string]$Title,
        [string]$OfficeName,
        [int]$UserIndex # Für eindeutiges Passwort
    )
    
    # Eindeutigen sAMAccountName ermitteln
    $counter = 0
    $sam = ''
    do {
        $counter++
        $sam = if ($counter -gt 1) { "$SamAccountNameBase$counter" } else { $SamAccountNameBase }
        $existingUser = Get-ADUser -Filter "SamAccountName -eq '$sam'" -SearchBase $PathOU -ErrorAction SilentlyContinue
    } while ($existingUser)

    $upn = "$sam@$UserPrincipalNameDomain"
    $logonFT = (Get-Date).AddDays(- (Get-Random -Minimum 0 -Maximum 366)).ToFileTime()
    $pwFT    = (Get-Date).AddDays(- (Get-Random -Minimum 0 -Maximum 181)).ToFileTime()

    $userParams = @{
        Name                  = "$FirstName $LastName"
        GivenName             = $FirstName
        Surname               = $LastName
        SamAccountName        = $sam
        UserPrincipalName     = $upn
        Path                  = $PathOU
        AccountPassword       = (ConvertTo-SecureString "$($Password)$UserIndex" -AsPlainText -Force)
        Enabled               = $true
        ChangePasswordAtLogon = $true
        Department            = $Department
        Title                 = $Title
        Office                = $OfficeName
        EmailAddress          = $upn
        Description           = "Test-User automatisch erstellt durch Skript"
        ErrorAction           = 'Stop'
    }
    try {
        $newUser = New-ADUser @userParams
        if ($newUser) {
            Set-ADObject -Identity $newUser.DistinguishedName -Replace @{ lastLogon = $logonFT; lastLogonTimestamp = $logonFT } -ErrorAction SilentlyContinue
            Set-ADUser -Identity $newUser.SamAccountName -Replace @{ pwdLastSet = $pwFT } -ErrorAction SilentlyContinue
            Write-Host ("→ Benutzer '{0}' erfolgreich in OU '{1}' angelegt." -f $sam, $PathOU) -ForegroundColor Green
            $Global:SessionCreatedUsers.Add($newUser) # Benutzer zur globalen Liste hinzufügen
            return $newUser
        } else {
            # Dieser Fall sollte eigentlich nicht eintreten, wenn New-ADUser fehlschlägt, da dann der Catch-Block greift.
            # Aber zur Sicherheit, falls New-ADUser $null zurückgibt ohne Fehler zu werfen.
            Write-Warning ("New-ADUser hat für '$sam' kein Objekt zurückgegeben, obwohl kein Fehler abgefangen wurde.")
            return $null
        }
    }
    catch {
        Write-Warning ("Fehler beim Erstellen des Benutzers '$sam' in OU '$PathOU': {0}" -f $_.Exception.Message)
        return $null
    }
}

#region Hilfsfunktionen für Menüoptionen 2-5

function Set-RandomADUserAttributes {
    param (
        [Parameter(Mandatory=$true)]
        [Microsoft.ActiveDirectory.Management.ADUser]$User # BaseOU wird nicht mehr benötigt, da User direkt übergeben wird
    )
    process {
        Write-Host ("Ändere zufällige Attribute für Benutzer '{0}'..." -f $User.SamAccountName) -ForegroundColor Yellow
        $attributesToSet = @{
        }
        $possibleAttributes = @(
            @{ Name = "Description"; Value = "Zufällige Beschreibung $$(Get-Random -Minimum 1000 -Maximum 9999)" },
            @{ Name = "Office"; Value = ($Global:NameLists.GetEnumerator() | Get-Random).Value.Last | Get-Random }, # Nimmt einen zufälligen Nachnamen als Office
            @{ Name = "Department"; Value = ($Global:Departments | Get-Random) },
            @{ Name = "Title"; Value = ($Global:Titles | Get-Random) },
            @{ Name = "telephoneNumber"; Value = "0{0}-{1}" -f (Get-Random -Minimum 100 -Maximum 999), (Get-Random -Minimum 1000000 -Maximum 9999999) },
            @{ Name = "mobile"; Value = "01{0} {1}" -f (Get-Random -Minimum 50 -Maximum 79), (Get-Random -Minimum 1000000 -Maximum 9999999) },
            @{ Name = "streetAddress"; Value = "Teststraße $$(Get-Random -Minimum 1 -Maximum 200)" },
            @{ Name = "l"; Value = "Teststadt" }, # City (l)
            @{ Name = "postalCode"; Value = "$$(Get-Random -Minimum 10000 -Maximum 99999)" },
            @{ Name = "st"; Value = "Bundesland $$(Get-Random -Minimum 1 -Maximum 16)" }, # State (st)
            @{ Name = "co"; Value = ($Global:NameLists.GetEnumerator() | Get-Random).Name } # Country (co)
        )

        # Wähle eine zufällige Anzahl von Attributen zum Setzen (mindestens 1)
        $numAttributesToChange = Get-Random -Minimum 1 -Maximum $possibleAttributes.Count
        $selectedAttributes = $possibleAttributes | Get-Random -Count $numAttributesToChange

        foreach ($attr in $selectedAttributes) {
            $attributesToSet[$attr.Name] = $attr.Value
            Write-Host ("  Setze '{0}' zu '{1}'" -f $attr.Name, $attr.Value) -ForegroundColor Gray
        }

        try {
            Set-ADUser -Identity $User -Replace $attributesToSet -ErrorAction Stop
            Write-Host ("→ Attribute für '{0}' erfolgreich aktualisiert." -f $User.SamAccountName) -ForegroundColor Green
        }
        catch {
            Write-Warning ("Fehler beim Aktualisieren der Attribute für '{0}': {1}" -f $User.SamAccountName, $_.Exception.Message)
        }
    }
}

function Set-RandomADUserSecurityFlags {
    param (
        [Parameter(Mandatory=$true)]
        [Microsoft.ActiveDirectory.Management.ADUser]$User
    )
    process {
        Write-Host ("Ändere zufällige Sicherheitseinstellungen für Benutzer '{0}'..." -f $User.SamAccountName) -ForegroundColor Yellow
        $flagsToChange = @(
            @{ Name = "UserCannotChangePassword"; Value = (Get-Random -Minimum 0 -Maximum 1 | ConvertTo-Bool) },
            @{ Name = "PasswordNeverExpires"; Value = (Get-Random -Minimum 0 -Maximum 1 | ConvertTo-Bool) },
            @{ Name = "PasswordNotRequired"; Value = $false } # Sollte i.d.R. false sein
        )
        # AccountExpires - komplexer, da Datumswert, erstmal weglassen oder als separates Flag

        $selectedFlag = $flagsToChange | Get-Random

        try {
            Invoke-Command -ScriptBlock (New-Object System.Management.Automation.ScriptBlock -ArgumentList ("Set-ADUser -Identity '{0}' -{1} ${2}" -f $User.DistinguishedName, $selectedFlag.Name, $selectedFlag.Value))
            # Set-ADUser -Identity $User -Replace @{ $selectedFlag.Name = $selectedFlag.Value } # Alternative, aber manche Flags sind keine Replace-Attribute
            Write-Host ("  Setze '{0}' zu '{1}'" -f $selectedFlag.Name, $selectedFlag.Value) -ForegroundColor Gray
            Write-Host ("→ Sicherheitseinstellung für '{0}' erfolgreich aktualisiert." -f $User.SamAccountName) -ForegroundColor Green
        }
        catch {
            Write-Warning ("Fehler beim Setzen der Sicherheitseinstellung '{0}' für '{1}': {2}" -f $selectedFlag.Name, $User.SamAccountName, $_.Exception.Message)
        }
    }
}

function Add-RandomADUserToCriticalGroup {
    param (
        [Parameter(Mandatory=$true)]
        [Microsoft.ActiveDirectory.Management.ADUser]$User,
        [string]$GroupName = "Test_Critical_Group"
    )
    process {
        Write-Host ("Füge Benutzer '{0}' zur Gruppe '{1}' hinzu..." -f $User.SamAccountName, $GroupName) -ForegroundColor Yellow
        try {
            $group = Get-ADGroup -Identity $GroupName -ErrorAction SilentlyContinue
            if (-not $group) {
                Write-Host "Gruppe '$GroupName' nicht gefunden, erstelle sie..." -ForegroundColor DarkYellow
                $group = New-ADGroup -Name $GroupName -GroupScope Global -GroupCategory Security -Path $BaseOU -PassThru -ErrorAction Stop
                Write-Host "Gruppe '$GroupName' erfolgreich in '$BaseOU' erstellt." -ForegroundColor Green
            }
            Add-ADGroupMember -Identity $group -Members $User -ErrorAction Stop
            Write-Host ("→ Benutzer '{0}' erfolgreich zur Gruppe '{1}' hinzugefügt." -f $User.SamAccountName, $GroupName) -ForegroundColor Green
        }
        catch {
            Write-Warning ("Fehler beim Hinzufügen von Benutzer '{0}' zur Gruppe '{1}': {2}" -f $User.SamAccountName, $GroupName, $_.Exception.Message)
        }
    }
}

#endregion

# --- Hauptlogik ---

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    Write-Verbose "ActiveDirectory Modul erfolgreich importiert."
}
catch {
    Write-Error "Kritischer Fehler: Das ActiveDirectory-Modul konnte nicht geladen werden. $($_.Exception.Message)"
    Write-Warning "Stellen Sie sicher, dass die RSAT-Tools für Active Directory installiert sind..."
    exit 1
}

# Überprüfen der BaseOU Existenz beim Start
if (-not (Get-ADOrganizationalUnit -Filter "DistinguishedName -eq '$BaseOU'" -ErrorAction SilentlyContinue)) {
    Write-Error "Die angegebene Basis-OU '$BaseOU' existiert nicht oder konnte nicht gelesen werden. Skript wird beendet."
    exit 1
}

# DomainDNS und DefaultPasswordPrefix initial abfragen/bestätigen, falls nötig
if (-not $DomainDNS) {
    $DomainDNS = Read-Host "Domänen-DNS-Name konnte nicht automatisch ermittelt werden. Bitte eingeben (z.B. contoso.com)"
    if (-not $DomainDNS) { Write-Error "Kein Domänenname angegeben. Skript wird beendet."; exit 1}
} else {
    $userConfirmation = Read-Host "Soll der Domänen-DNS-Name '$DomainDNS' verwendet werden? (J/N)"
    if ($userConfirmation -ne 'j') {
        $DomainDNS = Read-Host "Bitte den korrekten Domänen-DNS-Namen eingeben"
        if (-not $DomainDNS) { Write-Error "Kein Domänenname angegeben. Skript wird beendet."; exit 1}
    }
}

$userConfirmationPass = Read-Host "Soll das Passwort-Präfix '$DefaultPasswordPrefix' verwendet werden? (J/N)"
if ($userConfirmationPass -ne 'j') {
    $DefaultPasswordPrefix = Read-Host "Bitte das gewünschte Passwort-Präfix eingeben"
    if (-not $DefaultPasswordPrefix) { Write-Warning "Kein Passwort-Präfix angegeben, verwende '$($Global:DefaultPasswordPrefix)'."; $DefaultPasswordPrefix = $Global:DefaultPasswordPrefix }
}


do {
    Clear-Host
    Write-Host "================================== Create AD-User ==================================="
    Write-Host "Basis-OU:" -ForegroundColor Cyan
    Write-Host "  $BaseOU" -ForegroundColor Yellow
    Write-Host "Zieldomäne:" -ForegroundColor Cyan
    Write-Host "  $DomainDNS" -ForegroundColor Yellow
    Write-Host "Passwort:" -ForegroundColor Cyan
    Write-Host "  $DefaultPasswordPrefix" -ForegroundColor Yellow
    Write-Host "-------------------------------------------------------------------------------------"
    Write-Host "1. Anzugebene Anzahl an Benutzern anlegen"
    Write-Host "2. -> in dieser Sitzung erstellte Benutzer -> Zufällige Attribute zuweisen"
    Write-Host "3. -> in dieser Sitzung erstellte Benutzer -> deaktivieren"
    Write-Host "4. -> in dieser Sitzung erstellte Benutzer -> sichere Flags setzen"
    Write-Host "5. -> in dieser Sitzung erstellte Benutzer -> kritische Gruppen/Settings zuweisen"
    Write-Host "-------------------------------------------------------------------------------------"
    Write-Host "0. Beenden"
    Write-Host "====================================================================================="
    $choice = Read-Host "Option von oben wählen"
    Write-Host "-------------------------------------------------------------------------------------"

    switch ($choice) {
        '1' {
            Write-Host "`nOption 1: Neue Benutzer anlegen" -ForegroundColor Magenta

            # Anzahl abfragen (bleibt wie in v1.1)
            $currentUserCountInput = Read-Host "Wieviele User sollen angelegt werden? (Standard: $NewUserCount)"
            if (-not [string]::IsNullOrWhiteSpace($currentUserCountInput) -and $currentUserCountInput -match '^\d+$') {
                 $NewUserCount = [int]$currentUserCountInput
            } elseif (-not [string]::IsNullOrWhiteSpace($currentUserCountInput)) {
                 Write-Warning "Ungültige Eingabe für Anzahl. Verwende Standard: $NewUserCount"
            }
            
            # OU Auswahl Logik
            $childOUs = Get-ChildOUs -ParentOUPath $BaseOU -Recursive:$false # Direkte Kinder für Auswahl
            $targetOUsForCreation = [System.Collections.Generic.List[string]]::new()
            $selectedOUDescription = ""

            if ($childOUs -and $childOUs.Count -gt 0) {
                # Sub-OUs vorhanden, Auswahl anbieten
                Write-Host "`nGefundene direkte Sub-OUs unter '$BaseOU':" -ForegroundColor Yellow
                $ouSelectionOptions = @{}
                $i = 1
                $childOUs | ForEach-Object { Write-Host ("  {0}) {1} ({2})" -f $i, $_.Name, $_.DistinguishedName); $ouSelectionOptions[$i.ToString()] = $_.DistinguishedName; $i++ }
                
                Write-Host "`nWo sollen die Benutzer angelegt werden?" -ForegroundColor Yellow
                Write-Host "  S) In einer spezifischen Sub-OU aus der Liste oben"
                Write-Host "  V) Verteilt auf ALLE oben gelisteten direkten Sub-OUs"
                Write-Host "  B) Direkt in der Basis-OU '$BaseOU'"
                Write-Host "  A) Abbrechen und zurück zum Hauptmenü"
                $ouCreationChoice = Read-Host "Ihre Wahl (S/V/B/A)"

                switch ($ouCreationChoice.ToUpper()) {
                    'S' {
                        $selectedOUIndex = Read-Host "Bitte Nummer der Ziel-Sub-OU eingeben"
                        if ($ouSelectionOptions.ContainsKey($selectedOUIndex)) {
                            $targetOUsForCreation.Add($ouSelectionOptions[$selectedOUIndex])
                            $selectedOUDescription = "in spezifischer Sub-OU '" + ($childOUs | Where-Object {$_.DistinguishedName -eq $ouSelectionOptions[$selectedOUIndex]}).Name + "'"
                        } else {
                            Write-Warning "Ungültige Auswahl. Breche Benutzererstellung ab."
                            Read-Host "Enter..."; continue
                        }
                    }
                    'V' {
                        $childOUs.DistinguishedName | ForEach-Object { $targetOUsForCreation.Add($_) }
                        $selectedOUDescription = "verteilt auf alle $($childOUs.Count) direkten Sub-OUs"
                    }
                    'B' {
                        $targetOUsForCreation.Add($BaseOU)
                        $selectedOUDescription = "direkt in Basis-OU '$BaseOU'"
                    }
                    'A' { Write-Host "Benutzererstellung abgebrochen."; Read-Host "Enter..."; continue }
                    default { Write-Warning "Ungültige Auswahl. Breche Benutzererstellung ab."; Read-Host "Enter..."; continue }
                }
            } else {
                # Keine Sub-OUs gefunden, BaseOU automatisch verwenden
                Write-Host "`nKeine direkten Sub-OUs unter '$BaseOU' gefunden. Benutzer werden in '$BaseOU' erstellt." -ForegroundColor Yellow
                $targetOUsForCreation.Add($BaseOU)
                $selectedOUDescription = "direkt in Basis-OU '$BaseOU' (keine Sub-OUs gefunden)"
            }

            if ($targetOUsForCreation.Count -eq 0) {
                 Write-Warning "Keine Ziel-OU(s) ausgewählt. Breche Benutzererstellung ab."
                 Read-Host "Enter..."; continue
            }

            # Land abfragen
            Write-Host "`nFür welches Land sollen die Namen generiert werden?" -ForegroundColor Yellow
            $countryOptions = @{
                '1' = 'Japan'
                '2' = 'China'
                '3' = 'England (Englische Namen)'
                '4' = 'Deutschland (Deutsche Namen)'
                '5' = 'Spanien (Spanische Namen)'
                '6' = 'Frankreich (Französische Namen)'
                '7' = 'Italien (Italienische Namen)'
                '8' = 'Brasilien (Brasilianische Namen)'
                '9' = 'Russland (Russische Namen)'
                '10' = 'Indien (Indische Namen)'
                '11' = 'Alle (gemischte Namen)'
            }
            $countryOptions.GetEnumerator() | ForEach-Object { Write-Host ("  {0}) {1}" -f $_.Name, $_.Value) }
            
            $countryChoice = ''
            do {
                $countryChoice = Read-Host "Ihre Wahl (1-11)"
                if (-not $countryOptions.ContainsKey($countryChoice)) {
                    Write-Warning "Ungültige Auswahl. Bitte geben Sie eine Zahl von 1 bis 11 ein."
                }
            } while (-not $countryOptions.ContainsKey($countryChoice))

            $nameData = Get-CountrySpecificNames -CountryChoice $countryChoice
            $selectedCountryDisplay = $nameData.CountryName

            Write-Host ("Erstelle $NewUserCount Benutzer mit Namen für '$selectedCountryDisplay' $selectedOUDescription...") -ForegroundColor Cyan
            $usersCreatedSuccessfully = 0
            $ouIndexForDistribution = 0

            1..$NewUserCount | ForEach-Object {
                $userIndex = $_ 
                $firstName = $nameData.FirstNames | Get-Random
                $lastName  = $nameData.LastNames  | Get-Random
                $samBase   = "{0}.{1}" -f $firstName.ToLower(), $lastName.ToLower() -replace '[^a-z0-9.]',''
                
                $currentPathOU = $targetOUsForCreation[0] # Standard für einzelne OU
                if ($targetOUsForCreation.Count -gt 1) { # Verteilungslogik
                    $currentPathOU = $targetOUsForCreation[$ouIndexForDistribution]
                    $ouIndexForDistribution = ($ouIndexForDistribution + 1) % $targetOUsForCreation.Count
                }

                $office = if ($selectedCountryDisplay -eq 'Deutschland') { 'Deutschland' } elseif ($selectedCountryDisplay -eq 'Japan') { 'Tokio' } else { 'Headquarters' }

                $createdADUser = New-ADTestUserInternal `
                    -FirstName $firstName `
                    -LastName $lastName `
                    -SamAccountNameBase $samBase `
                    -UserPrincipalNameDomain $DomainDNS `
                    -PathOU $currentPathOU `
                    -Password $DefaultPasswordPrefix `
                    -Department ($Global:Departments | Get-Random) `
                    -Title ($Global:Titles | Get-Random) `
                    -OfficeName $office `
                    -UserIndex $userIndex
                if ($createdADUser) {
                    $usersCreatedSuccessfully++
                    $Global:SessionCreatedUsers.Add($createdADUser)
                }
            }
            Write-Host "`n$usersCreatedSuccessfully von $NewUserCount Benutzern erfolgreich erstellt." -ForegroundColor Green
            Write-Warning "WICHTIG: Die generierten Passwörter (`$($DefaultPasswordPrefix)1`, etc.) müssen den Passwortrichtlinien Ihrer Domäne entsprechen!"
            Read-Host "Fertig mit Option 1. Drücken Sie Enter, um zum Hauptmenü zurückzukehren."
        }

        '2' {
            Write-Host "`nOption 2: Zufällige Attribute für Benutzer setzen" -ForegroundColor Magenta
            if (-not ($Global:SessionCreatedUsers -and $Global:SessionCreatedUsers.Count -gt 0)) {
                Write-Warning "Es wurden in dieser Sitzung noch keine Benutzer mit Option 1 erstellt."
                Read-Host "Drücken Sie Enter..."; break
            }
            $usersToModify = $Global:SessionCreatedUsers | Get-Random -Count ([System.Math]::Min($AttrAssignCount, $Global:SessionCreatedUsers.Count))
            if (-not $usersToModify) {
                Write-Warning "Keine Benutzer zum Bearbeiten ausgewählt (möglicherweise ist $AttrAssignCount zu klein oder die Liste ist leer)."
                Read-Host "Drücken Sie Enter..."; break
            }
            Write-Host ("Bearbeite Attribute für {0} Benutzer..." -f $usersToModify.Count) -ForegroundColor Yellow
            foreach ($user in $usersToModify) {
                Set-RandomADUserAttributes -User $user # BaseOU wird in der Funktion nicht mehr benötigt, wenn User direkt übergeben wird
            }
            Read-Host "Fertig mit Option 2. Drücken Sie Enter..."
            break 
        }
        '3' {
            Write-Host "`nOption 3: Zufällige Benutzer deaktivieren" -ForegroundColor Magenta
            if (-not ($Global:SessionCreatedUsers -and $Global:SessionCreatedUsers.Count -gt 0)) {
                Write-Warning "Es wurden in dieser Sitzung noch keine Benutzer mit Option 1 erstellt."
                Read-Host "Drücken Sie Enter..."; break
            }
            $usersToDisable = $Global:SessionCreatedUsers | Get-Random -Count ([System.Math]::Min($DisableCount, $Global:SessionCreatedUsers.Count))
            if (-not $usersToDisable) {
                Write-Warning "Keine Benutzer zum Deaktivieren ausgewählt (möglicherweise ist $DisableCount zu klein oder die Liste ist leer)."
                Read-Host "Drücken Sie Enter..."; break
            }
            Write-Host ("Deaktiviere {0} Benutzer..." -f $usersToDisable.Count) -ForegroundColor Yellow
            foreach ($user in $usersToDisable) {
                try {
                    Disable-ADAccount -Identity $user -ErrorAction Stop
                    Write-Host ("→ Benutzer '{0}' erfolgreich deaktiviert." -f $user.SamAccountName) -ForegroundColor Green
                }
                catch {
                    Write-Warning ("Fehler beim Deaktivieren von Benutzer '{0}': {1}" -f $user.SamAccountName, $_.Exception.Message)
                }
            }
            Read-Host "Fertig mit Option 3. Drücken Sie Enter..."
            break 
        }
        '4' {
            Write-Host "`nOption 4: Zufällige Sicherheitseinstellungen für Benutzer ändern" -ForegroundColor Magenta
            if (-not ($Global:SessionCreatedUsers -and $Global:SessionCreatedUsers.Count -gt 0)) {
                Write-Warning "Es wurden in dieser Sitzung noch keine Benutzer mit Option 1 erstellt."
                Read-Host "Drücken Sie Enter..."; break
            }
            $usersToModify = $Global:SessionCreatedUsers | Get-Random -Count ([System.Math]::Min($SecSettingCount, $Global:SessionCreatedUsers.Count))
            if (-not $usersToModify) {
                Write-Warning "Keine Benutzer zum Bearbeiten ausgewählt (möglicherweise ist $SecSettingCount zu klein oder die Liste ist leer)."
                Read-Host "Drücken Sie Enter..."; break
            }
            Write-Host ("Bearbeite Sicherheitseinstellungen für {0} Benutzer..." -f $usersToModify.Count) -ForegroundColor Yellow
            foreach ($user in $usersToModify) {
                Set-RandomADUserSecurityFlags -User $user
            }
            Read-Host "Fertig mit Option 4. Drücken Sie Enter..."
            break 
        }
        '5' {
            Write-Host "`nOption 5: Zufällige Benutzer zu kritischer Test-Gruppe hinzufügen" -ForegroundColor Magenta
            if (-not ($Global:SessionCreatedUsers -and $Global:SessionCreatedUsers.Count -gt 0)) {
                Write-Warning "Es wurden in dieser Sitzung noch keine Benutzer mit Option 1 erstellt."
                Read-Host "Drücken Sie Enter..."; break
            }
            $usersToModify = $Global:SessionCreatedUsers | Get-Random -Count ([System.Math]::Min($CriticalCount, $Global:SessionCreatedUsers.Count))
            if (-not $usersToModify) {
                Write-Warning "Keine Benutzer zum Bearbeiten ausgewählt (möglicherweise ist $CriticalCount zu klein oder die Liste ist leer)."
                Read-Host "Drücken Sie Enter..."; break
            }
            Write-Host ("Füge {0} Benutzer zu kritischer Gruppe hinzu..." -f $usersToModify.Count) -ForegroundColor Yellow
            $criticalGroupName = Read-Host "Name der kritischen Test-Gruppe (Standard: Test_Critical_Group)"
            if ([string]::IsNullOrWhiteSpace($criticalGroupName)) { $criticalGroupName = "Test_Critical_Group" }
            foreach ($user in $usersToModify) {
                Add-RandomADUserToCriticalGroup -User $user -GroupName $criticalGroupName
            }
            Read-Host "Fertig mit Option 5. Drücken Sie Enter..."
            break 
        }

        '0' { Write-Host "Skript wird beendet."; break }
        default { Write-Warning "Ungültige Auswahl!"; Start-Sleep -Seconds 2 }
    }
} while ($choice -ne '0')

Write-Host "`nSkript beendet." -ForegroundColor Green
