====================================================================================================
# ANDREAS HEPP - easySPODesign
====================================================================================================

- **Version**         : 0.1.4
- **Letztes Update**  : 23.03.2025
- **Autor**           : ANDREAS HEPP


# SharePoint Online Design-Export-Tool
Dieses PowerShell-Skript ermöglicht den Export und die Übertragung von SharePoint Online Site-Designs zwischen verschiedenen Sites oder Tenants.

## Funktionen
- Export von Site-Design-Elementen von einer Quell-Site
- Übertragung auf eine einzelne Ziel-Site oder mehrere Sites mit gemeinsamen Präfix
- Unterstützung für MFA-Authentifizierung
- Grafische Benutzeroberfläche zur einfachen Konfiguration

## Wichtiger Hinweis zu Einschränkungen

**Dieses Tool überträgt nur Site-Design-Elemente**, darunter:
- Navigation
- Themes und Branding
- Seitenlayouts und -strukturen
- Header- und Footer-Konfiguration

Es überträgt **keine** Inhalte wie:
- Listen und Bibliotheken mit ihren Daten
- Dokumente und Dateien
- Benutzerberechtigungen
- Workflows, Power Automate-Flows oder Apps


## Voraussetzungen
- PowerShell 5.1 oder höher
- Installierte PnP.PowerShell-Modul
- Administrator-Rechte für SharePoint Online
- Berechtigungen zum Erstellen von Site-Designs

## Installation
1. Stellen Sie sicher, dass PowerShell installiert ist
2. Das Skript wird beim ersten Start das PnP.PowerShell-Modul automatisch installieren, falls es fehlt

## Verwendung
1. Führen Sie das Skript `exportSPODesign.ps1` aus
2. Geben Sie die Admin-URL und die Quell-Site-URL ein
3. Wählen Sie zwischen Export zu einer einzelnen Ziel-Site oder zu mehreren Sites
4. Bei Bedarf können erweiterte Export-Optionen aktiviert werden
5. Klicken Sie auf "OK", um den Export-Prozess zu starten

## Technische Hinweise
Die Site-Design-Übertragung basiert auf der PnP-Funktionalität von Microsoft. Dieses Skript verwendet die folgenden Befehle:

- `Get-PnPSiteScriptFromWeb` - Exportiert das Design einer bestehenden Site
- `Add-PnPSiteScript` - Erstellt ein neues Site-Script in der Ziel-Umgebung
- `Add-PnPSiteDesign` - Erstellt ein neues Site-Design mit dem erstellten Script
- `Invoke-PnPSiteDesign` - Wendet das Design auf die Ziel-Site(s) an

## Fehlerbehandlung
- Stellen Sie sicher, dass Sie über ausreichende Berechtigungen verfügen
- Überprüfen Sie, dass das PnP.PowerShell-Modul installiert ist
- Bei Verbindungsproblemen versuchen Sie eine manuelle Anmeldung vor dem Scriptstart

## Weitere Ressourcen
- [SharePoint PnP PowerShell Dokumentation](https://pnp.github.io/powershell/)
- [Microsoft SharePoint Site Design und Site Script Übersicht](https://learn.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-overview)
