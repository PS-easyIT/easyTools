## 🚀 easyWSAudit v0.2.3 - Windows Server Audit

### ✅ Über 120 Audit-Befehle

#### 🌐 Verbindungsaudit (v0.2.3)
*   **50+ spezialisierte Netzwerk-Audit-Befehle** für erweiterte Verbindungsanalysen
*   **Echtzeitanalyse aktiver TCP/UDP-Verbindungen** mit Prozess-Zuordnung
*   **Netzwerk-Topologie-Mapping** mit automatischer Geräte-Erkennung
*   **DRAW.IO-Export** für professionelle Netzwerk-Diagramme
*   **Sicherheitsanalyse** mit Blacklist-Checking und Bedrohungserkennung
*   **ARP-Cache und MAC-Adressen-Analyse** für lokale Netzwerk-Discovery


#### 📊 System-Audits (v0.0.12)
*   Lokale Administratoren & Guest Accounts
*   Shared Folders & User Profiles
*   Windows Firewall & Power Management
*   Credential Manager & Audit Policy
*   Group Policy Results & Environment Variables
*   Installierte Software (Registry-basiert)

#### 🔐 Active Directory (v0.0.12)
*   **ADFS (5 Audits):** Relying Party Trusts, Claims Provider, Certificates, Endpoints
*   **ADLDS (2 Audits):** Instances & Configuration
*   **ADRMS (3 Audits):** Cluster Info, Server Info, Templates

#### 🩺 AD Health Check (v0.1.9)
*   AD DHCP Server Discovery in Active Directory
*   AD Service Dependencies Health (DNS, DFS, Kerberos, etc.)
*   AD DC Diagnostics (dcdiag)
*   AD Time Sync Status & Sysvol Replication

#### 🛡️ Sicherheit & Compliance (v0.0.12)
*   **Device Health Attestation (1 Audit):** Encryption/Signing Certificates
*   **Volume Activation (2 Audits):** KMS Server/Client Status
*   **Windows Defender (3 Audits):** Status, Preferences, Threat Detection
*   **Host Guardian Service (2 Audits):** HGS Info & Attestation Policies

#### 💾 Backup & Storage (v0.1.9)
*   **Windows Server Backup (3 Audits):** Policies, Jobs, Disks
*   **Storage Management (4 Audits):** Storage Pools, Virtual Disks, Storage Spaces
*   **Windows Internal Database (2 Audits):** SQL Express Instances

#### 🌐 Netzwerk & Zugriff (v0.1.9)
*   **NPAS/NPS (4 Audits):** Network Policies, RADIUS Clients, Connection Policies
*   **Remote Access (3 Audits):** DirectAccess, VPN, Routing Table

#### ⚙️ System Services (v0.0.12)
*   **Windows Process Activation Service (2 Audits):** WAS Status & App Pools
*   **Windows Search Service (2 Audits):** Search Status & Indexer
*   **Windows Server Essentials (2 Audits):** Dashboard & Backup
*   **Migration Services (1 Audit):** Migration Tools
*   **Windows Identity Foundation (1 Audit):** Identity Features


### 🎯 Kernfunktionen im Überblick
*   **Über 120 verschiedene Audit-Befehle:** Umfassende Abdeckung für tiefgreifende Analysen.
*   **Automatische Rollen-Erkennung:** Identifiziert zuverlässig installierte Windows Server Rollen.
*   **20+ Audit-Kategorien:** Sorgt für eine übersichtliche Struktur und Navigation in der GUI.
*   **CMD + PowerShell Unterstützung:** Bietet Flexibilität bei der Ausführung von Audit-Befehlen.
*   **Professionelle HTML-Reports:** Ermöglicht eine einfache Analyse durch Tab-Navigation und klare Darstellung.
*   **Live-Fortschrittsanzeige:** Erlaubt die Verfolgung des Audit-Prozesses für alle Befehle in Echtzeit.


## 📋 Systemanforderungen
- **Windows Server 2016/2019/2022**
- **PowerShell 5.1** oder höher
- **Administrator-Rechte** für vollständige Audit-Funktionalität
- **.NET Framework 4.7.2** oder höher (für WPF-GUI)

## 🚀 Installation & Verwendung

### Schnellstart
```powershell
# 1. Script herunterladen
# 2. PowerShell als Administrator starten
# 3. Ausführungsrichtlinie temporär ändern (falls erforderlich)
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process

# 4. Script ausführen
.\easyWSAudit_V0.0.3.ps1


### Tipps für beste Ergebnisse
- **Führen Sie das Tool als Domain Administrator aus** für vollständige AD-Audits
- **Aktivieren Sie den Debug-Modus** bei Problemen (`$DEBUG = $true` in Zeile 5)
- **Verwenden Sie "Vollständiges Audit"** für eine komplette Systemanalyse
- **Exportieren Sie Berichte als HTML** für bessere Lesbarkeit und Archivierung
