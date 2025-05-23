# Dieses Skript erstellt einfache Icons für easyWSUS

$iconsPath = Split-Path -Parent $MyInvocation.MyCommand.Path

# Prüfen, ob System.Drawing.Common verfügbar ist
try {
    Add-Type -AssemblyName System.Drawing
}
catch {
    Write-Error "System.Drawing konnte nicht geladen werden. Icons werden nicht erstellt."
    exit 1
}

function Create-Icon {
    param (
        [string]$filename,
        [string]$text,
        [System.Drawing.Color]$backgroundColor,
        [System.Drawing.Color]$textColor
    )

    $bitmap = New-Object System.Drawing.Bitmap(32, 32)
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $graphics.Clear($backgroundColor)

    $font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
    $brush = New-Object System.Drawing.SolidBrush($textColor)
    
    # Zentrieren des Texts
    $textSize = $graphics.MeasureString($text, $font)
    $x = ($bitmap.Width - $textSize.Width) / 2
    $y = ($bitmap.Height - $textSize.Height) / 2
    
    $graphics.DrawString($text, $font, $brush, $x, $y)
    $graphics.Dispose()
    
    # Speichern als PNG
    $bitmap.Save("$iconsPath\$filename", [System.Drawing.Imaging.ImageFormat]::Png)
    $bitmap.Dispose()
}

# Icons erstellen
Create-Icon -filename "info.png" -text "i" -backgroundColor ([System.Drawing.Color]::SteelBlue) -textColor ([System.Drawing.Color]::White)
Create-Icon -filename "settings.png" -text "⚙" -backgroundColor ([System.Drawing.Color]::DarkGray) -textColor ([System.Drawing.Color]::White)
Create-Icon -filename "close.png" -text "X" -backgroundColor ([System.Drawing.Color]::IndianRed) -textColor ([System.Drawing.Color]::White)
Create-Icon -filename "logo.png" -text "WU" -backgroundColor ([System.Drawing.Color]::FromArgb(0, 120, 215)) -textColor ([System.Drawing.Color]::White)

Write-Output "Icons wurden erstellt in: $iconsPath"
