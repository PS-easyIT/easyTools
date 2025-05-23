# easyWusaWSUS - Windows Update Management Tool

Ein leistungsstarkes PowerShell-basiertes Tool zur Verwaltung von Windows Updates, mit benutzerfreundlicher GUI.

## Funktionen

- **Dashboard**: Systemübersicht und Update-Status
- **Update-Verlauf**: Anzeige und Export der Update-Historie
- **Dienst-Monitor**: Überwachung und Steuerung relevanter Windows Update-Dienste
- **Systemressourcen**: Überwachung von CPU-, RAM- und Festplattenauslastung
- **Log-Viewer**: Anzeige verschiedener Windows Update-relevanter Logs
- **Toolbox**: Zugriff auf häufig verwendete System-Tools
- **Fehlersuche**: Datenbank mit Windows Update-Fehlercodes und Lösungen

## Installation und Ausführung

1. Laden Sie alle Dateien in ein Verzeichnis herunter
2. Starten Sie das Tool über `easyWINUpdate.ps1`

### Systemvoraussetzungen

- Windows 10/11 oder Windows Server 2016/2019/2022
- PowerShell 5.1 oder höher
- Administratorrechte für einige Funktionen

## Fehlerbehebung

### Problem: Fehler bei der Ausführung von Befehlen oder Reparaturfunktionen

**Lösung**: Viele Funktionen des Tools erfordern Administratorrechte. Starten Sie PowerShell mit "Als Administrator ausführen".

## Hinweise

- Für volle Funktionalität sollte das Tool mit Administratorrechten ausgeführt werden
- Einige Funktionen (z.B. DISM, SFC) erfordern zwingend Administratorrechte
- Die Log-Dateien werden im `Logs`-Unterverzeichnis gespeichert

## Kontakt und Support

Bei Problemen oder für Support erstellen Sie bitte ein Issue im Repository oder kontaktieren Sie den Entwickler.

## Lizenz

Dieses Tool steht unter interner Lizenz und darf nur innerhalb der Organisation verwendet werden.