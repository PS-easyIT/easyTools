## Module

### Write-Log.psm1
Bietet eine einheitliche Logging-Funktionalität für alle Module.

**Hauptfunktion:**
- `Write-Log` - Schreibt formatierte Protokollnachrichten mit Zeitstempel und Schweregrad

### Test-ServerBasics.psm1
Überprüft grundlegende Server-Funktionalität und -Ressourcen.

**Hauptfunktion:**
- `Test-ServerBasics` - Führt grundlegende Tests für einen Server durch (Ping, Dienste, Ressourcennutzung)

### Test-RequiredAssets.psm1
Überprüft das Vorhandensein erforderlicher Ressourcen für GUI-Anwendungen.

**Hauptfunktion:**
- `Test-RequiredAssets` - Überprüft, ob die erforderlichen Assets (Icons, etc.) vorhanden sind

### Test-ADServices.psm1
Testet Active Directory-bezogene Dienste auf einem Server.

**Hauptfunktion:**
- `Test-ADServices` - Überprüft AD-Dienste auf einem Server (NTDS, Kerberos, etc.)

### Test-ADReplication.psm1
Überprüft den AD-Replikationsstatus zwischen Domänencontrollern.

**Hauptfunktion:**
- `Test-ADReplication` - Testet den Active Directory-Replikationsstatus für einen DC

### Switch-View.psm1
GUI-Hilfsfunktion zum Umschalten zwischen verschiedenen Ansichten in WPF-basierten Oberflächen.

**Hauptfunktion:**
- `Switch-View` - Wechselt zwischen verschiedenen Ansichten in einer WPF-GUI

### Open-URL.psm1
Einfache Funktion zum Öffnen von URLs im Standardbrowser.

**Hauptfunktion:**
- `Open-URL` - Öffnet die angegebene URL im Standardbrowser

### Get-DomainServers.psm1
Ruft Informationen über Windows-Server in der aktuellen Active Directory-Domäne ab.

**Hauptfunktion:**
- `Get-DomainServers` - Ruft alle Windows-Server aus der aktuellen AD-Domäne ab und prüft deren Online-Status


### Serverstatus prüfen
```powershell
$serverTests = Test-ServerBasics -ServerName "SERVER01"
$serverTests | Format-Table TestName, Result, Details -AutoSize
```

### AD-Replikation prüfen
```powershell
$replStatus = Test-ADReplication -ServerName "DC01"
if ($replStatus | Where-Object { $_.Status -eq "FEHLER" }) {
    Write-Host "Replikationsprobleme gefunden!" -ForegroundColor Red
}
```

### Alle Domänenserver auflisten
```powershell
$servers = Get-DomainServers
$servers | Where-Object { $_.Online -eq $true } | Format-Table Name, OS, Version -AutoSize
```
