# easyADReport

![easyADReport Screenshot](https://github.com/PS-easyIT/easyTools/blob/main/easyADReport/%23%20Screenshots/Screenshot_V0.4.1_ComputerByOSReport.jpg)

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

## 🆕 Latest Version

**Current Stable**: v0.5.3-FINAL

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
5. Execute: `.\easyADReport_v0.5.3-FINAL.ps1`

### First Run

```powershell
# Run stable version with elevated privileges
.\easyADReport_v0.5.3-FINAL.ps1

# With specific domain controller
.\easyADReport_v0.5.3-FINAL.ps1 -DomainController "DC01.domain.local"
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
- Compromise Indicators detection
- Authentication Protocol usage analysis
- Failed Authentication Patterns
- Advanced HoneyToken/Canary Account detection
- Service Account risk assessment

### 📋 Infrastructure Reports
- FSMO Role Holders
- Domain Controller Status
- Replication Health
- GPO Analysis
- Trust Relationships

## 📈 Version History

### v0.5.X Series
- Added comprehensive Security Dashboard with risk metrics
- Implemented Compromise Indicators detection
- Added Authentication Protocol Analysis
- Introduced HoneyToken detection capabilities
- Enhanced Failed Authentication Pattern analysis
- Added Service Account security assessment
- Improved risk scoring algorithms
- Performance optimizations and stability improvements

### v0.4.X Series
- Stabilized all existing features
- Performance improvements
- Bug fixes and UI enhancements
- Enhanced export functionality
- Added Computer by OS report
- Improved UI responsiveness
- Enhanced filtering capabilities

### v0.3.x Series
- Added GPO analysis reports
- Enhanced group membership analysis
- Added circular group detection
- Improved performance for large domains

### v0.2.x Series
- Added infrastructure reports
- Introduced export to CSV/HTML

### v0.1.x Series
- Added security analysis reports
- Enhanced user reports
- Added computer reports
- Initial group reports

### v0.0.x Series (Early Development)
- Added advanced filtering
- Implemented export functionality
- Enhanced UI design
- Added basic group reports
- Implemented user reports
- Initial release with basic functionality

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

## 🆕 Neueste Version

**Aktuelle Stable**: v0.5.3-FINAL

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
5. Ausführen: `.\easyADReport_v0.5.3-FINAL.ps1`

### Erster Start

```powershell
# Stabile Version mit erhöhten Rechten ausführen
.\easyADReport_v0.5.3-FINAL.ps1

# Mit spezifischem Domänencontroller
.\easyADReport_v0.5.3-FINAL.ps1 -DomainController "DC01.domain.local"
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
- Security Dashboard mit Bewertungssystem
- Erkennung von Kompromittierungsindikatoren
- Analyse der Authentifizierungsprotokoll-Nutzung
- Fehlgeschlagene Authentifizierungsmuster
- Erweiterte HoneyToken/Canary-Konto-Erkennung
- Risikobewertung von Dienstkonten

### 📋 Infrastrukturberichte
- FSMO-Rolleninhaber
- Domänencontroller-Status
- Replikationsintegrität
- GPO-Analyse
- Vertrauensstellungen

## 📈 Versionshistorie

### v0.5.X Serie
- Umfassendes Security Dashboard mit Risikometriken hinzugefügt
- Erkennung von Kompromittierungsindikatoren implementiert
- Authentifizierungsprotokoll-Analyse hinzugefügt
- HoneyToken-Erkennungsfunktionen eingeführt
- Erweiterte Analyse fehlgeschlagener Authentifizierungsmuster
- Sicherheitsbewertung von Dienstkonten hinzugefügt
- Verbesserte Risikobewertungsalgorithmen
- Leistungsoptimierungen und Stabilitätsverbesserungen

### v0.4.X Serie
- Alle bestehenden Funktionen stabilisiert
- Leistungsverbesserungen
- Fehlerbehebungen und UI-Verbesserungen
- Erweiterte Exportfunktionalität
- Computer nach OS-Bericht hinzugefügt
- Verbesserte UI-Reaktionsfähigkeit
- Erweiterte Filterfunktionen

### v0.3.x Serie
- GPO-Analyseberichte hinzugefügt
- Erweiterte Gruppenmitgliedschaftsanalyse
- Zirkuläre Gruppenerkennung hinzugefügt
- Verbesserte Leistung für große Domänen

### v0.2.x Serie
- Infrastrukturberichte hinzugefügt
- Export nach CSV/HTML eingeführt

### v0.1.x Serie
- Sicherheitsanalyseberichte hinzugefügt
- Erweiterte Benutzerberichte
- Computerberichte hinzugefügt
- Erste Gruppenberichte

### v0.0.x Serie
- Erweiterte Filterung hinzugefügt
- Exportfunktionalität implementiert
- UI-Design verbessert
- Grundlegende Gruppenberichte hinzugefügt
- Benutzerberichte implementiert
- Erstveröffentlichung mit Grundfunktionalität

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

---

## 📊 Quick Report Reference

### Overview of all available reports / Übersicht aller verfügbaren Berichte

| REPORT | ENG Description | DE Beschreibung |
|--------|-----------------|-----------------|
| **USER REPORTS** | | |
| All Users | Lists all user accounts in the domain | Listet alle Benutzerkonten in der Domäne auf |
| Disabled Users | Shows all disabled user accounts | Zeigt alle deaktivierten Benutzerkonten an |
| Locked Users | Displays users with locked accounts | Zeigt Benutzer mit gesperrten Konten an |
| Inactive Users | Lists users inactive for specified period | Listet inaktive Benutzer für einen bestimmten Zeitraum auf |
| Password Expiring | Shows users with expiring passwords | Zeigt Benutzer mit ablaufenden Passwörtern an |
| Password Never Expires | Lists users with non-expiring passwords | Listet Benutzer mit nicht ablaufenden Passwörtern auf |
| Stale Passwords | Shows users who haven't changed passwords recently | Zeigt Benutzer, die ihre Passwörter lange nicht geändert haben |
| Administrative Accounts | Lists all administrative user accounts | Listet alle administrativen Benutzerkonten auf |
| Never Logged On | Shows users who have never logged in | Zeigt Benutzer, die sich noch nie angemeldet haben |
| Users by Department | Groups users by their department | Gruppiert Benutzer nach ihrer Abteilung |
| Users by Manager | Shows organizational hierarchy by manager | Zeigt Organisationshierarchie nach Manager |
| **GROUP REPORTS** | | |
| All Groups | Lists all groups in the domain | Listet alle Gruppen in der Domäne auf |
| Security Groups | Shows all security groups | Zeigt alle Sicherheitsgruppen an |
| Distribution Groups | Lists all distribution groups | Listet alle Verteilergruppen auf |
| Empty Groups | Shows groups without members | Zeigt Gruppen ohne Mitglieder an |
| Nested Groups | Displays group nesting hierarchy | Zeigt verschachtelte Gruppenhierarchie an |
| Circular Group References | Detects circular group memberships | Erkennt zirkuläre Gruppenmitgliedschaften |
| Mail-Enabled Groups | Lists groups with email addresses | Listet Gruppen mit E-Mail-Adressen auf |
| Groups without Owners | Shows groups lacking designated owners | Zeigt Gruppen ohne zugewiesene Besitzer |
| **COMPUTER REPORTS** | | |
| All Computers | Lists all computer accounts | Listet alle Computerkonten auf |
| Inactive Computers | Shows computers inactive for specified period | Zeigt inaktive Computer für einen bestimmten Zeitraum |
| Operating System Summary | Groups computers by OS version | Gruppiert Computer nach Betriebssystemversion |
| BitLocker Status | Shows BitLocker encryption status | Zeigt BitLocker-Verschlüsselungsstatus an |
| Duplicate Computer Names | Detects duplicate computer names | Erkennt doppelte Computernamen |
| Computers by Location | Groups computers by physical location | Gruppiert Computer nach physischem Standort |
| **SECURITY ANALYSIS** | | |
| Privileged Accounts | Analyzes accounts with elevated privileges | Analysiert Konten mit erhöhten Berechtigungen |
| Kerberoastable Accounts | Identifies accounts vulnerable to Kerberoasting | Identifiziert für Kerberoasting anfällige Konten |
| ASREPRoastable Accounts | Lists accounts vulnerable to AS-REP roasting | Listet für AS-REP Roasting anfällige Konten auf |
| Weak Password Policies | Analyzes password policy weaknesses | Analysiert Schwächen in Passwortrichtlinien |
| Permission Analysis | Reviews object permissions and delegations | Überprüft Objektberechtigungen und Delegierungen |
| DCSync Rights | Shows accounts with DCSync permissions | Zeigt Konten mit DCSync-Berechtigungen |
| SID History Abuse | Detects potential SID history abuse | Erkennt potenziellen SID-Verlauf-Missbrauch |
| Compromise Indicators | Identifies signs of potential compromise | Identifiziert Anzeichen möglicher Kompromittierung |
| Authentication Protocols | Analyzes authentication protocol usage | Analysiert Nutzung von Authentifizierungsprotokollen |
| Failed Authentication | Shows failed authentication patterns | Zeigt fehlgeschlagene Authentifizierungsmuster |
| HoneyToken Detection | Detects canary/honeypot accounts | Erkennt Canary-/Honeypot-Konten |
| Service Account Security | Assesses service account security risks | Bewertet Sicherheitsrisiken von Dienstkonten |
| **INFRASTRUCTURE REPORTS** | | |
| FSMO Role Holders | Shows Flexible Single Master Operation roles | Zeigt Flexible Single Master Operation Rollen |
| Domain Controllers | Lists all domain controllers and status | Listet alle Domänencontroller und Status auf |
| Replication Health | Shows AD replication status and issues | Zeigt AD-Replikationsstatus und Probleme |
| GPO Analysis | Analyzes Group Policy Objects | Analysiert Gruppenrichtlinienobjekte |
| Trust Relationships | Shows domain trust configurations | Zeigt Domänen-Vertrauensstellungen |
| Sites and Subnets | Lists AD sites and subnet configurations | Listet AD-Standorte und Subnetzkonfigurationen |
| DNS Zones | Shows integrated DNS zones | Zeigt integrierte DNS-Zonen |
| **SECURITY DASHBOARD** | | |
| Risk Overview | Comprehensive security risk assessment | Umfassende Sicherheitsrisikobewertung |
| Security Score | Overall domain security score calculation | Berechnung des Gesamtsicherheitswerts der Domäne |
| Critical Findings | Highlights critical security issues | Hebt kritische Sicherheitsprobleme hervor |
| Remediation Priority | Prioritized list of security fixes | Priorisierte Liste von Sicherheitskorrekturen |
| Compliance Status | Shows compliance with security standards | Zeigt Compliance mit Sicherheitsstandards |

### Report Categories Legend / Berichtskategorien-Legende

- 👤 **User Reports** / **Benutzerberichte**: Focus on user account management / Fokus auf Benutzerkontenverwaltung
- 👥 **Group Reports** / **Gruppenberichte**: Group membership and structure analysis / Gruppenmitgliedschaft und Strukturanalyse
- 💻 **Computer Reports** / **Computerberichte**: Computer account and system information / Computerkonten und Systeminformationen
- 🔐 **Security Analysis** / **Sicherheitsanalyse**: Security vulnerabilities and risks / Sicherheitslücken und Risiken
- 📋 **Infrastructure** / **Infrastruktur**: AD infrastructure health and configuration / AD-Infrastruktur-Zustand und Konfiguration
- 🛡️ **Security Dashboard** / **Sicherheits-Dashboard**: Comprehensive security overview / Umfassende Sicherheitsübersicht