# easyADReport

![easyADReport Screenshot](https://github.com/PS-easyIT/easyTools/blob/main/easyADReport/%23%20Screenshots/Screenshot_V0.4.2_PasswordNeverExpires-Report.jpg)

## 📊 Overview

**easyADReport** is a comprehensive PowerShell-based Active Directory reporting tool with a modern WPF GUI. It provides over 80 pre-built reports for auditing, security analysis, and compliance monitoring of your Active Directory environment.

### ✨ Key Features

- **80+ Pre-built Reports**: Comprehensive coverage of AD objects and security configurations
- **Modern WPF Interface**: Clean, intuitive Windows 11-style UI
- **Real-time Analysis**: Instant report generation without database requirements
- **Export Capabilities**: Export results to CSV, Excel, HTML, or PDF
- **Advanced Filtering**: Powerful search and filter options for all reports
- **Security Focused**: Extensive security analysis and vulnerability detection
- **No Dependencies**: Works with standard PowerShell and AD modules

## 🚀 Quick Start

### Prerequisites

- Windows 10/11 or Windows Server 2016+
- PowerShell 5.1 or higher
- Active Directory PowerShell module
- Domain Administrator or appropriate read permissions
- .NET Framework 4.7.2 or higher

### Installation

1. Download the latest release from the repository
2. Extract to your preferred location
3. Run PowerShell as Administrator
4. Navigate to the script directory
5. Execute: `.\easyADReport_v0.4.2.ps1`

### First Run

```powershell
# Run with elevated privileges
.\easyADReport_v0.4.2.ps1

# Or with specific domain controller
.\easyADReport_v0.4.2.ps1 -DomainController "DC01.domain.local"
```

## 📋 Report Categories

### 👤 Users Reports
- All Users
- Disabled/Locked/Inactive Users
- Password Status (Expiring, Never Expire, Stale)
- Administrative Accounts
- Never Logged On Users
- Organization Structure (by Department, Manager)

### 👥 Groups Reports
- Security/Distribution Groups
- Empty Groups
- Nested Groups & Circular References
- Mail-Enabled Groups
- Groups without Owners

### 💻 Computers Reports
- All Computers
- Inactive Computers
- Operating System Summary
- BitLocker Status
- Duplicate Names

### 🔐 Security Analysis
- Privileged Account Analysis
- Kerberoastable/ASREPRoastable Accounts
- Weak Password Policies
- Permission Analysis
- DCSync Rights
- SID History Abuse Detection

### 📋 Infrastructure Reports
- FSMO Role Holders
- Domain Controller Status
- Replication Health
- GPO Analysis
- Trust Relationships

## 🛠️ Usage Examples

### Basic Report Generation

1. Launch the application
2. Select a report from the sidebar
3. Click the report button
4. View results in the main grid
5. Export if needed

### Advanced Filtering

```
# Filter examples in the search box:
Name -like "*admin*"
Department -eq "IT"
LastLogonDate -lt (Get-Date).AddDays(-90)
```

### Bulk Operations

The tool supports multi-select for:
- Exporting multiple objects
- Comparing user/group memberships
- Batch analysis

## 📤 Export Options

- **CSV**: For Excel analysis
- **HTML**: For web viewing and sharing

## 🐛 Troubleshooting

### Common Issues

1. **"Access Denied" Error**
   - Ensure you have appropriate AD read permissions
   - Run PowerShell as Administrator

2. **"Module Not Found" Error**
   - Install RSAT tools: `Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0`

3. **Performance Issues**
   - Limit result size in settings
   - Use specific OUs instead of entire domain

## 📄 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

---

# easyADReport (Deutsch)

## 📊 Übersicht

**easyADReport** ist ein umfassendes PowerShell-basiertes Active Directory Reporting-Tool mit moderner WPF-Oberfläche. Es bietet über 80 vorgefertigte Berichte für Auditing, Sicherheitsanalyse und Compliance-Überwachung Ihrer Active Directory-Umgebung.

### ✨ Hauptfunktionen

- **80+ Vorgefertigte Berichte**: Umfassende Abdeckung von AD-Objekten und Sicherheitskonfigurationen
- **Moderne WPF-Oberfläche**: Saubere, intuitive Windows 11-Style Benutzeroberfläche
- **Echtzeit-Analyse**: Sofortige Berichtserstellung ohne Datenbankanforderungen
- **Export-Funktionen**: Export der Ergebnisse nach CSV, Excel, HTML oder PDF
- **Erweiterte Filterung**: Leistungsstarke Such- und Filteroptionen für alle Berichte
- **Sicherheitsfokussiert**: Umfangreiche Sicherheitsanalyse und Schwachstellenerkennung
- **Keine Abhängigkeiten**: Funktioniert mit Standard PowerShell und AD-Modulen

## 🚀 Schnellstart

### Voraussetzungen

- Windows 10/11 oder Windows Server 2016+
- PowerShell 5.1 oder höher
- Active Directory PowerShell-Modul
- Domänenadministrator oder entsprechende Leseberechtigungen
- .NET Framework 4.7.2 oder höher

### Installation

1. Laden Sie die neueste Version aus dem Repository herunter
2. Entpacken Sie an Ihren bevorzugten Speicherort
3. Führen Sie PowerShell als Administrator aus
4. Navigieren Sie zum Skriptverzeichnis
5. Ausführen: `.\easyADReport_v0.4.2.ps1`

### Erster Start

```powershell
# Mit erhöhten Rechten ausführen
.\easyADReport_v0.4.2.ps1

# Oder mit spezifischem Domänencontroller
.\easyADReport_v0.4.2.ps1 -DomainController "DC01.domain.local"
```

## 📋 Berichtskategorien

### 👤 Benutzerberichte
- Alle Benutzer
- Deaktivierte/Gesperrte/Inaktive Benutzer
- Passwortstatus (Ablaufend, Nie ablaufend, Veraltet)
- Administrative Konten
- Nie angemeldete Benutzer
- Organisationsstruktur (nach Abteilung, Manager)

### 👥 Gruppenberichte
- Sicherheits-/Verteilergruppen
- Leere Gruppen
- Verschachtelte Gruppen & Zirkelbezüge
- Mail-aktivierte Gruppen
- Gruppen ohne Besitzer

### 💻 Computerberichte
- Alle Computer
- Inaktive Computer
- Betriebssystem-Zusammenfassung
- BitLocker-Status
- Doppelte Namen

### 🔐 Sicherheitsanalyse
- Privilegierte Kontenanalyse
- Kerberoastbare/ASREPRoastbare Konten
- Schwache Passwortrichtlinien
- Berechtigungsanalyse
- DCSync-Rechte
- SID-Verlauf Missbrauchserkennung

### 📋 Infrastrukturberichte
- FSMO-Rolleninhaber
- Domänencontroller-Status
- Replikationsintegrität
- GPO-Analyse
- Vertrauensstellungen

## 🛠️ Verwendungsbeispiele

### Grundlegende Berichtserstellung

1. Starten Sie die Anwendung
2. Wählen Sie einen Bericht aus der Seitenleiste
3. Klicken Sie auf die Berichtsschaltfläche
4. Ergebnisse im Hauptraster anzeigen
5. Bei Bedarf exportieren

### Erweiterte Filterung

```
# Filterbeispiele im Suchfeld:
Name -like "*admin*"
Department -eq "IT"
LastLogonDate -lt (Get-Date).AddDays(-90)
```

### Massenoperationen

Das Tool unterstützt Mehrfachauswahl für:
- Export mehrerer Objekte
- Vergleich von Benutzer-/Gruppenmitgliedschaften
- Batch-Analyse

## 📤 Exportoptionen

- **CSV**: Für Excel-Analyse
- **HTML**: Für Webanzeige und Freigabe

## 🐛 Fehlerbehebung

### Häufige Probleme

1. **"Zugriff verweigert" Fehler**
   - Stellen Sie sicher, dass Sie entsprechende AD-Leseberechtigungen haben
   - PowerShell als Administrator ausführen

2. **"Modul nicht gefunden" Fehler**
   - RSAT-Tools installieren: `Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0`

3. **Leistungsprobleme**
   - Ergebnisgröße in den Einstellungen begrenzen
   - Spezifische OUs anstelle der gesamten Domäne verwenden

## 📄 Lizenz

Dieses Projekt ist unter der MIT-Lizenz lizenziert - siehe LICENSE-Datei für Details.

## 🤝 Mitwirken

Beiträge sind willkommen! Bitte zögern Sie nicht, einen Pull Request einzureichen.