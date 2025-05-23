<#
.SYNOPSIS
    Launcher für das SharePoint Online Design Export Tool
.DESCRIPTION
    Dieses Skript prüft die Ausführungsrichtlinien und startet das SharePoint Online Design Export Tool
.NOTES
    Version: 1.0.0
    Autor: Andreas Hepp
#>

# Aktuelle Ausführungsrichtlinie speichern
$originalExecutionPolicy = Get-ExecutionPolicy -Scope Process

Write-Host "SharePoint Design Export Tool - Launcher" -ForegroundColor Cyan
Write-Host "=========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Aktuelle Ausführungsrichtlinie: $originalExecutionPolicy" -ForegroundColor Yellow

# Pfad zum Hauptskript
$mainScriptPath = Join-Path -Path $PSScriptRoot -ChildPath "exportSPODesign_V0.2.0.ps1"

if (-not (Test-Path $mainScriptPath)) {
    Write-Host "FEHLER: Das Hauptskript wurde nicht gefunden unter: $mainScriptPath" -ForegroundColor Red
    Write-Host "Bitte stellen Sie sicher, dass sich das Skript im selben Verzeichnis befindet." -ForegroundColor Red
    exit 1
}

Write-Host "Hauptskript gefunden: $mainScriptPath" -ForegroundColor Green
Write-Host ""
Write-Host "Dieses Launcher-Skript wird die Ausführungsrichtlinie temporär ändern, um das Hauptskript auszuführen." -ForegroundColor Yellow
Write-Host "Die ursprüngliche Ausführungsrichtlinie wird nach Beendigung wiederhergestellt." -ForegroundColor Yellow
Write-Host ""

$confirm = Read-Host "Möchten Sie fortfahren? (J/N)"
if ($confirm -ne "J" -and $confirm -ne "j") {
    Write-Host "Vorgang abgebrochen." -ForegroundColor Red
    exit
}

try {
    # Temporär Ausführungsrichtlinie auf RemoteSigned setzen (nur für den aktuellen Prozess)
    Write-Host "Setze Ausführungsrichtlinie temporär auf RemoteSigned..." -ForegroundColor Cyan
    Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process -Force
    
    # Ausführungsrichtlinie nach Änderung prüfen
    $newPolicy = Get-ExecutionPolicy -Scope Process
    Write-Host "Neue Ausführungsrichtlinie: $newPolicy" -ForegroundColor Green
    
    # Hauptskript ausführen
    Write-Host "Starte Hauptskript..." -ForegroundColor Cyan
    & $mainScriptPath
    
    Write-Host "Hauptskript wurde beendet." -ForegroundColor Green
}
catch {
    Write-Host "FEHLER beim Ausführen des Skripts:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
}
finally {
    # Optional: Ursprüngliche Ausführungsrichtlinie wiederherstellen
    if ($originalExecutionPolicy -ne (Get-ExecutionPolicy -Scope Process)) {
        Write-Host "Stelle ursprüngliche Ausführungsrichtlinie wieder her..." -ForegroundColor Cyan
        Set-ExecutionPolicy -ExecutionPolicy $originalExecutionPolicy -Scope Process -Force
        Write-Host "Ausführungsrichtlinie wurde zurückgesetzt auf: $originalExecutionPolicy" -ForegroundColor Green
    }
}
