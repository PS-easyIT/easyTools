# easyTools

**easyTools** ist eine Sammlung nützlicher PowerShell-Skripte zur Administration von Windows Server, Microsoft 365 und verwandten Umgebungen. Die Skripte decken verschiedene Aufgaben ab, darunter Systemanalysen, Modulerkennung, Sicherheitsüberprüfungen und Automatisierung typischer Verwaltungsaufgaben.

---

## ▶️ Ausführen der Skripte

Um die PowerShell-Skripte korrekt ausführen zu können, kann es erforderlich sein, die **Execution Policy** temporär oder dauerhaft anzupassen. Standardmäßig blockiert Windows eventuell nicht signierte Skripte.

### Ausführungsrichtlinien in PowerShell

| Richtlinie       | Beschreibung |
|------------------|--------------|
| `Restricted`     | Keine Skriptausführung erlaubt (Standard unter Windows) |
| `AllSigned`      | Nur signierte Skripte erlaubt |
| `RemoteSigned`   | Lokale Skripte erlaubt, Remote-Skripte müssen signiert sein |
| `Unrestricted`   | Alle Skripte können ausgeführt werden (mit Warnung) |
| `Bypass`         | Keine Einschränkungen, keine Warnungen |

### Anzeige der aktuellen Richtlinie

Get-ExecutionPolicy -List

## Ändern der aktuellen Richtlinie

    Temporäre Änderung für die aktuelle Sitzung:

	Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass

    Dauerhafte Änderung für den aktuellen Benutzer:

	Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned


## ⚠️ Hinweis zur Sicherheit

Verwende diese Skripte nur in kontrollierten Umgebungen und prüfe den Quellcode vor Ausführung.
Passe sie ggf. an deine Umgebung und Sicherheitsanforderungen an.


## 📄 Lizenz

Diese Skripte werden unter der MIT License veröffentlicht. Nutzung auf eigene Verantwortung.


## 🤝 Mitwirken

Pull Requests und Feature-Vorschläge sind willkommen! 