Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles() | Out-Null

# Formular
$form = New-Object System.Windows.Forms.Form
$form.Text = 'SharePoint URL Decoder'
$form.Size = New-Object System.Drawing.Size(520, 220)
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'
$form.Topmost = $true

# Eingabe-Label
$label = New-Object System.Windows.Forms.Label
$label.Text = 'SharePoint Site URL (encoded):'
$label.Location = New-Object System.Drawing.Point(10, 20)
$label.Size = New-Object System.Drawing.Size(480, 20)
$form.Controls.Add($label)

# Eingabefeld
$textBoxInput = New-Object System.Windows.Forms.TextBox
$textBoxInput.Location = New-Object System.Drawing.Point(10, 45)
$textBoxInput.Size = New-Object System.Drawing.Size(480, 25)
$form.Controls.Add($textBoxInput)

# Ausgabe-Feld
$textBoxOutput = New-Object System.Windows.Forms.TextBox
$textBoxOutput.Location = New-Object System.Drawing.Point(10, 85)
$textBoxOutput.Size = New-Object System.Drawing.Size(480, 25)
$textBoxOutput.ReadOnly = $true
$form.Controls.Add($textBoxOutput)

# Button: Dekodieren
$btnDecode = New-Object System.Windows.Forms.Button
$btnDecode.Text = 'Dekodieren'
$btnDecode.Location = New-Object System.Drawing.Point(150, 125)
$btnDecode.Size = New-Object System.Drawing.Size(100, 30)
$btnDecode.DialogResult = [System.Windows.Forms.DialogResult]::None
$form.Controls.Add($btnDecode)

# Button: Kopieren
$btnCopy = New-Object System.Windows.Forms.Button
$btnCopy.Text = 'Kopieren'
$btnCopy.Location = New-Object System.Drawing.Point(270, 125)
$btnCopy.Size = New-Object System.Drawing.Size(100, 30)
$btnCopy.DialogResult = [System.Windows.Forms.DialogResult]::None
$btnCopy.Enabled = $false
$form.Controls.Add($btnCopy)

# Event: Dekodieren
$btnDecode.Add_Click({
    $encoded = $textBoxInput.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($encoded)) {
        [System.Windows.Forms.MessageBox]::Show('Bitte gib eine gültige URL ein.','Fehler',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning)
    } else {
        try {
            $decoded = [uri]::UnescapeDataString($encoded)
            $textBoxOutput.Text = $decoded
            $btnCopy.Enabled = $true
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Fehler beim Dekodieren:`n$($_.Exception.Message)",'Fehler',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
})

# Event: Kopieren
$btnCopy.Add_Click({
    if (-not [string]::IsNullOrWhiteSpace($textBoxOutput.Text)) {
        Set-Clipboard -Value $textBoxOutput.Text
        [System.Windows.Forms.MessageBox]::Show('Dekodierte URL wurde in die Zwischenablage kopiert.','Erfolg',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
    }
})

# Zeige das Formular
$form.Add_Shown({ $textBoxInput.Select() })
[void]$form.ShowDialog()
$form.Dispose()
