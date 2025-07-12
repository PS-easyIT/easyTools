# Assembly-Loading für Windows Forms und Drawing
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Globale Variablen
$global:currentPolicies = @{}
$global:backupPolicies = @{}
$global:autoBackupPath = "$env:APPDATA\easyExecutionPolicy\AutoBackups"
$policyOptions = @("Restricted", "AllSigned", "RemoteSigned", "Unrestricted", "Bypass", "Undefined")

# Hilfsfunktion zum Erstellen des Backup-Verzeichnisses
function Initialize-BackupDirectory {
    if (-not (Test-Path $global:autoBackupPath)) {
        try {
            New-Item -Path $global:autoBackupPath -ItemType Directory -Force | Out-Null
            Write-Host "Backup-Verzeichnis erstellt: $global:autoBackupPath"
        } catch {
            Write-Warning "Konnte Backup-Verzeichnis nicht erstellen: $($_.Exception.Message)"
            $global:autoBackupPath = $env:TEMP
        }
    }
}

# Funktion zur Überprüfung und Anforderung von Administratorrechten
function Ensure-Admin {
    if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        $dialogResult = [System.Windows.Forms.MessageBox]::Show("Dieses Skript erfordert Administratorrechte. Möchten Sie es als Administrator neu starten?", "Administratorrechte erforderlich", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)

        if ($dialogResult -eq [System.Windows.Forms.DialogResult]::Yes) {
            try {
                $powershellPath = (Get-Command PowerShell).Source
                $argumentList = "-File `"$($MyInvocation.MyCommand.Path)`""
                Start-Process -FilePath $powershellPath -Verb RunAs -ArgumentList $argumentList -ErrorAction Stop
                Write-Host "Die nicht-administrative Instanz wird beendet, da eine neue administrative Instanz gestartet wurde."
                exit
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Fehler beim Neustart des Skripts als Administrator:`n$($_.Exception.Message)", "Fehler beim Neustart", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                Write-Host "Fehler beim Neustart als Administrator. Das Skript wird beendet."
                exit
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show("Das Skript kann ohne Administratorrechte nicht fortfahren und wird nun beendet.", "Administratorrechte verweigert", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            Write-Host "Benutzer hat den Neustart als Administrator abgelehnt. Das Skript wird beendet."
            exit
        }
    }
}

# Function to get current Execution Policies
function Get-CurrentExecutionPolicies {
    $scopes = @("MachinePolicy", "UserPolicy", "Process", "CurrentUser", "LocalMachine")
    foreach ($scope in $scopes) {
        try {
            $policy = Get-ExecutionPolicy -Scope $scope -ErrorAction SilentlyContinue 
            if ($null -eq $policy) {
                $global:currentPolicies[$scope] = "Undefined"
            } else {
                $global:currentPolicies[$scope] = $policy.ToString()
            }
        } catch {
            $global:currentPolicies[$scope] = "Error"
        }
    }
    try {
        $global:currentPolicies["Effective"] = (Get-ExecutionPolicy).ToString()
    } catch {
        $global:currentPolicies["Effective"] = "Error"
    }
}

# Funktion zum Erstellen eines Backups der aktuellen Policies
function Backup-CurrentPolicies {
    param(
        [switch]$SaveToFile,
        [string]$CustomPath
    )
    
    $global:backupPolicies = $global:currentPolicies.Clone()
    Write-Host "Backup der aktuellen Execution Policies erstellt (Arbeitsspeicher)."
    
    if ($SaveToFile -or $CustomPath) {
        try {
            Initialize-BackupDirectory
            
            $backupData = @{
                BackupDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                ComputerName = $env:COMPUTERNAME
                Username = $env:USERNAME
                BackupType = if ($CustomPath) { "Manual" } else { "Automatic" }
                Policies = $global:currentPolicies
            }
            
            $filePath = if ($CustomPath) { 
                $CustomPath 
            } else { 
                "$global:autoBackupPath\ExecutionPolicyBackup_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
            }
            
            $backupData | ConvertTo-Json -Depth 3 | Out-File -FilePath $filePath -Encoding UTF8
            Write-Host "Backup gespeichert unter: $filePath"
            
            if (-not $CustomPath) {
                $autoBackups = Get-ChildItem -Path $global:autoBackupPath -Filter "ExecutionPolicyBackup_*.json" | 
                              Sort-Object CreationTime -Descending
                if ($autoBackups.Count -gt 10) {
                    $autoBackups | Select-Object -Skip 10 | Remove-Item -Force
                    Write-Host "Alte Backups bereinigt. Behalten: 10"
                }
            }
            
            return $filePath
        } catch {
            Write-Warning "Fehler beim Speichern des Backups: $($_.Exception.Message)"
            [System.Windows.Forms.MessageBox]::Show("Backup konnte nicht gespeichert werden:`n$($_.Exception.Message)", "Backup Warnung", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    }
    
    return $null
}

# Funktion zum Wiederherstellen der Backup-Policies
function Restore-BackupPolicies {
    param(
        [switch]$FromFile,
        [string]$BackupFilePath
    )
    
    if ($FromFile -or $BackupFilePath) {
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.Filter = "JSON Backup files (*.json)|*.json|All files (*.*)|*.*"
        $openFileDialog.Title = "Backup-Datei zum Wiederherstellen auswählen"
        $openFileDialog.InitialDirectory = $global:autoBackupPath
        
        $filePath = if ($BackupFilePath) {
            $BackupFilePath
        } else {
            if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $openFileDialog.FileName
            } else {
                return $false
            }
        }
        
        try {
            $backupContent = Get-Content -Path $filePath -Raw | ConvertFrom-Json
            $policiesToRestore = $backupContent.Policies
            
            $confirmText = "Backup-Details:`n"
            $confirmText += "Datum: $($backupContent.BackupDate)`n"
            $confirmText += "Computer: $($backupContent.ComputerName)`n"
            $confirmText += "Typ: $($backupContent.BackupType)`n`n"
            $confirmText += "Möchten Sie diese Policies wiederherstellen?"
            
            $dialogResult = [System.Windows.Forms.MessageBox]::Show($confirmText, "Backup wiederherstellen", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Fehler beim Lesen der Backup-Datei:`n$($_.Exception.Message)", "Restore Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return $false
        }
    } else {
        if ($global:backupPolicies.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Kein Backup im Arbeitsspeicher verfügbar!`n`nVerwenden Sie 'Aus Datei wiederherstellen' um ein gespeichertes Backup zu laden.", "Restore Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return $false
        }
        
        $policiesToRestore = $global:backupPolicies
        $dialogResult = [System.Windows.Forms.MessageBox]::Show("Möchten Sie wirklich alle Execution Policies auf den Backup-Stand (Arbeitsspeicher) zurücksetzen?", "Restore bestätigen", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    }
    
    if ($dialogResult -eq [System.Windows.Forms.DialogResult]::Yes) {
        try {
            $settableScopes = @("Process", "CurrentUser", "LocalMachine")
            foreach ($scope in $settableScopes) {
                if ($policiesToRestore.PSObject.Properties.Name -contains $scope) {
                    $backupPolicy = $policiesToRestore.$scope
                    if ($backupPolicy -and $backupPolicy -ne "Undefined" -and $backupPolicy -ne "Error") {
                        Set-ExecutionPolicy -ExecutionPolicy $backupPolicy -Scope $scope -Force -ErrorAction Stop
                    }
                    Write-Host "Wiederhergestellt: $scope = $backupPolicy"
                }
            }
            Update-PolicyDisplay
            return $true
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Fehler beim Wiederherstellen: $($_.Exception.Message)", "Restore Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return $false
        }
    }
    return $false
}

# Funktion für manuelles Backup mit Dateiauswahl
function Create-ManualBackup {
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "JSON Backup files (*.json)|*.json|All files (*.*)|*.*"
    $saveFileDialog.Title = "Backup speichern unter..."
    $saveFileDialog.FileName = "ExecutionPolicyBackup_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
    $saveFileDialog.InitialDirectory = $global:autoBackupPath
    
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $savedPath = Backup-CurrentPolicies -SaveToFile -CustomPath $saveFileDialog.FileName
        if ($savedPath) {
            [System.Windows.Forms.MessageBox]::Show("Backup erfolgreich gespeichert unter:`n$savedPath", "Backup erstellt", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            return $true
        }
    }
    return $false
}

# Validierungsfunktion für Policy-Änderungen
function Validate-PolicyChange {
    param($Scope, $NewPolicy, $CurrentPolicy)
    
    if ($NewPolicy -eq "Unrestricted" -or $NewPolicy -eq "Bypass") {
        $result = [System.Windows.Forms.MessageBox]::Show("WARNUNG: Die Policy '$NewPolicy' für Scope '$Scope' kann Sicherheitsrisiken darstellen.`n`nMöchten Sie trotzdem fortfahren?", "Sicherheitswarnung", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return ($result -eq [System.Windows.Forms.DialogResult]::Yes)
    }
    
    return $true
}

Ensure-Admin

# GUI Creation
$mainForm = New-Object System.Windows.Forms.Form
$mainForm.Text = "Execution Policy Manager v0.0.3"
$mainForm.Size = New-Object System.Drawing.Size(820, 450)
$mainForm.StartPosition = "CenterScreen"
$mainForm.FormBorderStyle = 'FixedDialog'
$mainForm.MaximizeBox = $false
$mainForm.MinimizeBox = $false
$mainForm.BackColor = [System.Drawing.Color]::White

# Define Fonts and Colors
$uiFontFamily = "Segoe UI"
$uiFontSizeNormal = 12
$uiFontSizeLarge = 20

$fontNormal = New-Object System.Drawing.Font($uiFontFamily, $uiFontSizeNormal)
$fontBold = New-Object System.Drawing.Font($uiFontFamily, $uiFontSizeNormal, [System.Drawing.FontStyle]::Bold)
$fontHeader = New-Object System.Drawing.Font($uiFontFamily, $uiFontSizeLarge, [System.Drawing.FontStyle]::Bold)

$colorDarkText = [System.Drawing.Color]::FromArgb(28, 28, 28)
$colorAccent = [System.Drawing.Color]::FromArgb(0, 120, 212)
$colorSuccess = [System.Drawing.Color]::FromArgb(16, 124, 16)
$colorWarning = [System.Drawing.Color]::FromArgb(255, 140, 0)
$colorGray = [System.Drawing.Color]::FromArgb(108, 117, 125)

# Menü hinzufügen
$menuStrip = New-Object System.Windows.Forms.MenuStrip
$fileMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$fileMenu.Text = "Datei"

$exportMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$exportMenuItem.Text = "Konfiguration exportieren..."
$exportMenuItem.Add_Click({ 
    # Export function inline für bessere Referenz
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
    $saveFileDialog.Title = "Export Execution Policy Konfiguration"
    $saveFileDialog.FileName = "ExecutionPolicy_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
    
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            $exportData = @{
                ExportDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                ComputerName = $env:COMPUTERNAME
                Username = $env:USERNAME
                Policies = $global:currentPolicies
            }
            $exportData | ConvertTo-Json -Depth 3 | Out-File -FilePath $saveFileDialog.FileName -Encoding UTF8
            [System.Windows.Forms.MessageBox]::Show("Konfiguration erfolgreich exportiert nach:`n$($saveFileDialog.FileName)", "Export erfolgreich", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Fehler beim Export: $($_.Exception.Message)", "Export Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
})

$importMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$importMenuItem.Text = "Konfiguration importieren..."
$importMenuItem.Add_Click({ 
    # Import function inline
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
    $openFileDialog.Title = "Import Execution Policy Konfiguration"
    
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            $importContent = Get-Content -Path $openFileDialog.FileName -Raw | ConvertFrom-Json
            $result = [System.Windows.Forms.MessageBox]::Show("Import von '$($importContent.ComputerName)' ($($importContent.ExportDate))?`n`nNur änderbare Scopes werden importiert.", "Import bestätigen", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
            
            if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
                $settableScopes = @("Process", "CurrentUser", "LocalMachine")
                foreach ($scope in $settableScopes) {
                    if ($importContent.Policies.PSObject.Properties.Name -contains $scope) {
                        $policyValue = $importContent.Policies.$scope
                        if ($policyValue -and $policyValue -ne "Undefined" -and $policyValue -ne "Error") {
                            Set-ExecutionPolicy -ExecutionPolicy $policyValue -Scope $scope -Force
                        }
                    }
                }
                Update-PolicyDisplay
                [System.Windows.Forms.MessageBox]::Show("Konfiguration erfolgreich importiert!", "Import erfolgreich", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Fehler beim Import: $($_.Exception.Message)", "Import Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
})

$separatorMenuItem = New-Object System.Windows.Forms.ToolStripSeparator
$exitMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$exitMenuItem.Text = "Beenden"
$exitMenuItem.Add_Click({ $mainForm.Close() })

$fileMenu.DropDownItems.AddRange(@($exportMenuItem, $importMenuItem, $separatorMenuItem, $exitMenuItem))

$toolsMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$toolsMenu.Text = "Tools"

$backupMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$backupMenuItem.Text = "Backup erstellen (Datei)"
$backupMenuItem.Add_Click({ 
    if (Create-ManualBackup) {
        $statusLabel.Text = "Manuelles Backup erfolgreich erstellt."
        $statusLabel.ForeColor = $colorSuccess
    }
})

$restoreMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$restoreMenuItem.Text = "Backup wiederherstellen"
$restoreMenuItem.Add_Click({ 
    if (Restore-BackupPolicies) {
        $statusLabel.Text = "Backup erfolgreich wiederhergestellt."
        $statusLabel.ForeColor = $colorSuccess
    }
})

$restoreFromFileMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$restoreFromFileMenuItem.Text = "Aus Datei wiederherstellen..."
$restoreFromFileMenuItem.Add_Click({ 
    if (Restore-BackupPolicies -FromFile) {
        $statusLabel.Text = "Backup aus Datei erfolgreich wiederhergestellt."
        $statusLabel.ForeColor = $colorSuccess
    }
})

$toolsMenu.DropDownItems.AddRange(@($backupMenuItem, $restoreMenuItem, $restoreFromFileMenuItem))
$menuStrip.Items.AddRange(@($fileMenu, $toolsMenu))
$mainForm.MainMenuStrip = $menuStrip
$mainForm.Controls.Add($menuStrip)

# Header Panel
$headerPanel = New-Object System.Windows.Forms.Panel
$headerPanel.Size = New-Object System.Drawing.Size($mainForm.ClientSize.Width, 60)
$headerPanel.Location = New-Object System.Drawing.Point(0, $menuStrip.Height)
$headerPanel.BackColor = [System.Drawing.Color]::FromArgb(28, 50, 60)
$headerPanel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$mainForm.Controls.Add($headerPanel)

$headerLabel = New-Object System.Windows.Forms.Label
$headerLabel.Text = "Execution Policy Manager"
$headerLabel.Font = $fontHeader
$headerLabel.ForeColor = [System.Drawing.Color]::FromArgb(226, 226, 226)
$headerLabel.AutoSize = $true
$headerLabel.Location = New-Object System.Drawing.Point(20, 17)
$headerPanel.Controls.Add($headerLabel)

# Content Panel
$contentPanel = New-Object System.Windows.Forms.Panel
$contentPanel.Location = New-Object System.Drawing.Point(0, ($headerPanel.Location.Y + $headerPanel.Height))
$contentPanel.Size = New-Object System.Drawing.Size($mainForm.ClientSize.Width, ($mainForm.ClientSize.Height - $headerPanel.Height - $menuStrip.Height - 60))
$contentPanel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$contentPanel.AutoScroll = $true
$mainForm.Controls.Add($contentPanel)

# Footer Panel
$footerPanel = New-Object System.Windows.Forms.Panel
$footerPanel.Size = New-Object System.Drawing.Size($mainForm.ClientSize.Width, 60)
$footerPanel.Location = New-Object System.Drawing.Point(0, ($mainForm.ClientSize.Height - $footerPanel.Height))
$footerPanel.BackColor = [System.Drawing.Color]::White
$footerPanel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$mainForm.Controls.Add($footerPanel)

# ToolTip Object
$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.AutoPopDelay = 10000
$toolTip.InitialDelay = 500

# Define Tooltip Texts
$scopeTooltips = @{
    "MachinePolicy" = "Richtlinie für alle Benutzer dieses Computers, typischerweise per GPO gesetzt."
    "UserPolicy"    = "Richtlinie für den aktuellen Benutzer, typischerweise per GPO gesetzt."
    "Process"       = "Richtlinie nur für die aktuelle PowerShell-Sitzung. Geht beim Schließen verloren."
    "CurrentUser"   = "Richtlinie für den aktuellen Benutzer, in der Registry gespeichert (HKCU)."
    "LocalMachine"  = "Standardrichtlinie für alle Benutzer dieses Computers, wenn keine spezifischere Richtlinie greift (HKLM)."
}
$policyExplanationTooltips = @{
    "Restricted"    = "Keine Skripte dürfen ausgeführt werden. PowerShell kann nur im interaktiven Modus verwendet werden."
    "AllSigned"     = "Nur Skripte, die von einem vertrauenswürdigen Herausgeber signiert wurden, können ausgeführt werden."
    "RemoteSigned"  = "Lokal erstellte Skripte können ausgeführt werden. Aus dem Internet heruntergeladene Skripte müssen von einem vertrauenswürdigen Herausgeber signiert sein."
    "Unrestricted"  = "Alle Skripte können ausgeführt werden. Vorsicht: Dies kann unsicher sein."
    "Bypass"        = "Nichts wird blockiert und es gibt keine Warnungen oder Eingabeaufforderungen."
    "Undefined"     = "Es ist keine Ausführungsrichtlinie in diesem Bereich festgelegt. Die Richtlinie eines übergeordneten Bereichs oder die Standardrichtlinie greift."
    "loading..."    = "Die aktuelle Richtlinie wird geladen..."
}

# Create GUI elements - Labels und ComboBoxes
$yPos = 25
$leftMargin = 20
$column1Width = 350
$column2Width = 180
$column3Width = 200
$spacing = 25

$column2XPos = $leftMargin + $column1Width + $spacing
$column3XPos = $column2XPos + $column2Width + $spacing

$scopesToManage = @(
    @{Name="MachinePolicy"; Display="MachinePolicy (GPO - All users)"; Settable=$false},
    @{Name="UserPolicy"; Display="UserPolicy (GPO - Current user)"; Settable=$false},
    @{Name="Process"; Display="Process (Current session only)"; Settable=$true},
    @{Name="CurrentUser"; Display="CurrentUser (Registry HKCU)"; Settable=$true},
    @{Name="LocalMachine"; Display="LocalMachine (Registry HKLM)"; Settable=$true}
)

foreach ($scopeInfo in $scopesToManage) {
    $scopeName = $scopeInfo.Name
    $scopeDisplay = $scopeInfo.Display
    $isSettable = $scopeInfo.Settable

    # Label for Scope name
    $labelScope = New-Object System.Windows.Forms.Label
    $labelScope.Text = "${scopeDisplay}:"
    $labelScope.Location = New-Object System.Drawing.Point($leftMargin, $yPos)
    $labelScope.Size = New-Object System.Drawing.Size($column1Width, 20)
    $labelScope.Font = $fontNormal
    $labelScope.ForeColor = $colorDarkText
    $labelScope.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $contentPanel.Controls.Add($labelScope)
    if ($scopeTooltips.ContainsKey($scopeName)) {
        $toolTip.SetToolTip($labelScope, $scopeTooltips[$scopeName])
    }

    # Label for current Policy
    $lblCurrentPolicy = New-Object System.Windows.Forms.Label
    $lblCurrentPolicy.Name = "lblCurrent$scopeName"
    $lblCurrentPolicy.Text = "loading..."
    $lblCurrentPolicy.Location = New-Object System.Drawing.Point($column2XPos, $yPos)
    $lblCurrentPolicy.Size = New-Object System.Drawing.Size($column2Width, 20)
    $lblCurrentPolicy.Font = $fontBold
    $lblCurrentPolicy.ForeColor = $colorDarkText
    $lblCurrentPolicy.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $contentPanel.Controls.Add($lblCurrentPolicy)

    # ComboBox or placeholder
    if ($isSettable) {
        $cbPolicy = New-Object System.Windows.Forms.ComboBox
        $cbPolicy.Name = "cb$scopeName"
        $cbPolicy.Location = New-Object System.Drawing.Point($column3XPos, ($yPos - 3))
        $cbPolicy.Size = New-Object System.Drawing.Size($column3Width, 26)
        $cbPolicy.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
        $cbPolicy.Font = $fontNormal
        $policyOptions | ForEach-Object { $cbPolicy.Items.Add($_) } | Out-Null
        $contentPanel.Controls.Add($cbPolicy)
        $toolTip.SetToolTip($cbPolicy, "Wählen Sie hier die neue Richtlinie für '$scopeDisplay'.")
    } else {
        $lblNotSettable = New-Object System.Windows.Forms.Label
        $lblNotSettable.Text = "(Controlled by GPO)"
        $lblNotSettable.Location = New-Object System.Drawing.Point($column3XPos, $yPos)
        $lblNotSettable.Size = New-Object System.Drawing.Size($column3Width, 20)
        $lblNotSettable.Font = $fontNormal
        $lblNotSettable.ForeColor = $colorGray
        $lblNotSettable.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
        $contentPanel.Controls.Add($lblNotSettable)
        $toolTip.SetToolTip($lblNotSettable, "Diese Richtlinie wird durch GPO verwaltet.")
    }
    $yPos += 31
}

# Separator Line
$yPos += 10
$separatorLine = New-Object System.Windows.Forms.Label
$separatorLine.Location = New-Object System.Drawing.Point($leftMargin, $yPos)
$separatorLine.Size = New-Object System.Drawing.Size(($column1Width + $column2Width + $column3Width + 2 * $spacing), 1)
$separatorLine.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
$contentPanel.Controls.Add($separatorLine)
$yPos += 10

# Label for effective Execution Policy
$lblEffectivePolicy = New-Object System.Windows.Forms.Label
$lblEffectivePolicy.Text = "Effective Policy (currently active):"
$lblEffectivePolicy.Location = New-Object System.Drawing.Point($leftMargin, $yPos)
$lblEffectivePolicy.Size = New-Object System.Drawing.Size($column1Width, 20)
$lblEffectivePolicy.Font = $fontNormal
$lblEffectivePolicy.ForeColor = $colorDarkText
$lblEffectivePolicy.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$contentPanel.Controls.Add($lblEffectivePolicy)
$toolTip.SetToolTip($lblEffectivePolicy, "Die effektive Ausführungsrichtlinie für die aktuelle PowerShell-Sitzung.")

$lblEffectivePolicyValue = New-Object System.Windows.Forms.Label
$lblEffectivePolicyValue.Name = "lblEffectivePolicyValue"
$lblEffectivePolicyValue.Text = "loading..."
$lblEffectivePolicyValue.Location = New-Object System.Drawing.Point($column2XPos, $yPos)
$lblEffectivePolicyValue.Size = New-Object System.Drawing.Size(($column2Width + $spacing + $column3Width), 20)
$lblEffectivePolicyValue.Font = $fontBold
$lblEffectivePolicyValue.ForeColor = $colorAccent
$lblEffectivePolicyValue.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$contentPanel.Controls.Add($lblEffectivePolicyValue)

# Status Label
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Name = "statusLabel"
$statusLabel.Text = "Ready."
$statusLabel.Location = New-Object System.Drawing.Point($leftMargin, 5)
$statusLabel.Size = New-Object System.Drawing.Size(($footerPanel.Width - 2 * $leftMargin), 15)
$statusLabel.Font = $fontNormal
$statusLabel.ForeColor = $colorSuccess
$statusLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$footerPanel.Controls.Add($statusLabel)

# Buttons
$applyButton = New-Object System.Windows.Forms.Button
$applyButton.Text = "APPLY Selected Policies"
$applyButton.Location = New-Object System.Drawing.Point(($footerPanel.Width - 250 - $leftMargin), 25)
$applyButton.Size = New-Object System.Drawing.Size(250, 28)
$applyButton.Font = $fontBold
$applyButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$applyButton.FlatAppearance.BorderSize = 0
$applyButton.BackColor = $colorAccent
$applyButton.ForeColor = [System.Drawing.Color]::White
$applyButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
$footerPanel.Controls.Add($applyButton)

$refreshButton = New-Object System.Windows.Forms.Button
$refreshButton.Text = "Refresh"
$refreshButton.Location = New-Object System.Drawing.Point($leftMargin, 25)
$refreshButton.Size = New-Object System.Drawing.Size(80, 28)
$refreshButton.Font = $fontNormal
$refreshButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$refreshButton.FlatAppearance.BorderSize = 0
$refreshButton.BackColor = $colorGray
$refreshButton.ForeColor = [System.Drawing.Color]::White
$refreshButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$footerPanel.Controls.Add($refreshButton)

$backupButton = New-Object System.Windows.Forms.Button
$backupButton.Text = "Backup"
$backupButton.Location = New-Object System.Drawing.Point(($leftMargin + 90), 25)
$backupButton.Size = New-Object System.Drawing.Size(80, 28)
$backupButton.Font = $fontNormal
$backupButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$backupButton.FlatAppearance.BorderSize = 0
$backupButton.BackColor = $colorGray
$backupButton.ForeColor = [System.Drawing.Color]::White
$backupButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$footerPanel.Controls.Add($backupButton)

$restoreButton = New-Object System.Windows.Forms.Button
$restoreButton.Text = "Restore"
$restoreButton.Location = New-Object System.Drawing.Point(($leftMargin + 180), 25)
$restoreButton.Size = New-Object System.Drawing.Size(80, 28)
$restoreButton.Font = $fontNormal
$restoreButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$restoreButton.FlatAppearance.BorderSize = 0
$restoreButton.BackColor = $colorWarning
$restoreButton.ForeColor = [System.Drawing.Color]::White
$restoreButton.Enabled = $false
$restoreButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$footerPanel.Controls.Add($restoreButton)

# ToolTips für Buttons
$toolTip.SetToolTip($applyButton, "Wenden Sie die ausgewählten Richtlinienänderungen an.")
$toolTip.SetToolTip($refreshButton, "Aktualisieren Sie die angezeigten Richtlinienwerte.")
$toolTip.SetToolTip($backupButton, "Erstellt ein Backup der aktuellen Execution Policies.")
$toolTip.SetToolTip($restoreButton, "Stellt das letzte Backup wieder her.")

# Function to update GUI display
function Update-PolicyDisplay {
    Get-CurrentExecutionPolicies
    $scopes = @("MachinePolicy", "UserPolicy", "Process", "CurrentUser", "LocalMachine")
    for ($i = 0; $i -lt $scopes.Length; $i++) {
        $scopeName = $scopes[$i]
        $label = $contentPanel.Controls.Find("lblCurrent$scopeName", $true)[0]
        $comboBox = $contentPanel.Controls.Find("cb$scopeName", $true)[0]

        if ($label) {
            $currentPolicyValue = $global:currentPolicies[$scopeName]
            $label.Text = $currentPolicyValue
            $label.ForeColor = $colorDarkText
            if ($policyExplanationTooltips.ContainsKey($currentPolicyValue)) {
                $toolTip.SetToolTip($label, $policyExplanationTooltips[$currentPolicyValue])
            }
        }
        if ($comboBox) {
            $policyValueToSelect = $global:currentPolicies[$scopeName]
            $selectedIndex = $policyOptions.IndexOf($policyValueToSelect)
            if ($selectedIndex -eq -1) {
                $selectedIndex = $policyOptions.IndexOf("Undefined")
            }
            $comboBox.SelectedIndex = $selectedIndex
        }
    }
    try {
        $effectivePolicy = Get-ExecutionPolicy
        $lblEffectivePolicyValue = $contentPanel.Controls.Find("lblEffectivePolicyValue", $true)[0]
        if ($lblEffectivePolicyValue) {
            $lblEffectivePolicyValue.Text = $effectivePolicy.ToString()
            $lblEffectivePolicyValue.ForeColor = $colorAccent
            if ($policyExplanationTooltips.ContainsKey($effectivePolicy.ToString())) {
                $toolTip.SetToolTip($lblEffectivePolicyValue, $policyExplanationTooltips[$effectivePolicy.ToString()])
            }
        }
    } catch {
        $lblEffectivePolicyValue = $contentPanel.Controls.Find("lblEffectivePolicyValue", $true)[0]
        if ($lblEffectivePolicyValue) {
            $lblEffectivePolicyValue.Text = "Error retrieving"
            $lblEffectivePolicyValue.ForeColor = [System.Drawing.Color]::Red
        }
    }
}

# Event Handlers
$refreshButton.Add_Click({
    $statusLabel.Text = "Refreshing policies..."
    $statusLabel.ForeColor = $colorWarning
    Update-PolicyDisplay
    $statusLabel.Text = "Policies refreshed."
    $statusLabel.ForeColor = $colorSuccess
})

$backupButton.Add_Click({
    Backup-CurrentPolicies
    $statusLabel.Text = "Backup der aktuellen Policies erstellt."
    $statusLabel.ForeColor = $colorSuccess
    $restoreButton.Enabled = $true
})

$restoreButton.Add_Click({
    if (Restore-BackupPolicies) {
        $statusLabel.Text = "Backup erfolgreich wiederhergestellt."
        $statusLabel.ForeColor = $colorSuccess
    }
})

$applyButton.Add_Click({
    $statusLabel.Text = "Validating and applying selected policies..."
    $statusLabel.ForeColor = $colorWarning

    $changesMade = $false
    $errorOccurred = $false

    $scopesToManage | Where-Object {$_.Settable} | ForEach-Object {
        $scopeName = $_.Name
        $comboBox = $contentPanel.Controls.Find("cb$scopeName", $true)[0]
        if ($comboBox -and $comboBox.SelectedItem) {
            $selectedPolicy = $comboBox.SelectedItem.ToString()
            $currentPolicy = $global:currentPolicies[$scopeName] 

            if ($selectedPolicy -ne $currentPolicy) {
                if (-not (Validate-PolicyChange -Scope $scopeName -NewPolicy $selectedPolicy -CurrentPolicy $currentPolicy)) {
                    continue
                }
                
                try {
                    Write-Host "Setting ExecutionPolicy for Scope '$scopeName' to '$selectedPolicy'..."
                    if ($selectedPolicy -eq "Undefined") {
                        if ($scopeName -eq "Process") {
                            Set-ExecutionPolicy -ExecutionPolicy Undefined -Scope $scopeName -Force -ErrorAction Stop
                        } else {
                            $registryPath = switch ($scopeName) {
                                "CurrentUser" { "HKCU:\SOFTWARE\Microsoft\PowerShell\1\ShellIds\Microsoft.PowerShell" }
                                "LocalMachine" { "HKLM:\SOFTWARE\Microsoft\PowerShell\1\ShellIds\Microsoft.PowerShell" }
                                default { "" }
                            }
                            
                            if ($registryPath -and (Test-Path $registryPath)) {
                                if (Get-ItemProperty -Path $registryPath -Name "ExecutionPolicy" -ErrorAction SilentlyContinue) {
                                    Remove-ItemProperty -Path $registryPath -Name "ExecutionPolicy" -Force -ErrorAction Stop
                                    Write-Host "Removed ExecutionPolicy registry value for Scope '$scopeName'."
                                }
                            }
                        }
                    } else {
                        Set-ExecutionPolicy -ExecutionPolicy $selectedPolicy -Scope $scopeName -Force -ErrorAction Stop
                    }
                    $changesMade = $true
                    Write-Host "Successfully processed Scope '$scopeName' for policy '$selectedPolicy'."
                } catch {
                    $errorMessage = "Error for Scope ${scopeName}: $($_.Exception.Message.Split([Environment]::NewLine)[0])"
                    $statusLabel.Text = $errorMessage 
                    $statusLabel.ForeColor = [System.Drawing.Color]::Red
                    $errorOccurred = $true
                    Write-Warning "Error setting ExecutionPolicy for Scope '$scopeName' to '$selectedPolicy': $($_.Exception.Message)"
                }
            }
        }
    }

    Update-PolicyDisplay

    if ($errorOccurred) {
        # Status bereits gesetzt im Catch-Block
    } elseif ($changesMade) {
        $statusLabel.Text = "Execution policies successfully applied."
        $statusLabel.ForeColor = $colorSuccess
    } else {
        $statusLabel.Text = "No changes were made to execution policies."
        $statusLabel.ForeColor = $colorAccent
    }
})

# On Form Load
$mainForm.Add_Load({
    $mainForm.BeginInvoke([Action]{ 
        Update-PolicyDisplay
        Backup-CurrentPolicies
        $restoreButton.Enabled = $true
        $statusLabel.Text = "Ready. Initial backup created."
        $statusLabel.ForeColor = $colorSuccess
    })
})

# Show GUI
$mainForm.ShowDialog() | Out-Null

# Clean up
$mainForm.Dispose()
# SIG # Begin signature block
# MIIcCAYJKoZIhvcNAQcCoIIb+TCCG/UCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC+b1YCeJgSrvKS
# Sn9RLLvMEsQiK+4wHZoM9RRfTU9aAqCCFk4wggMQMIIB+KADAgECAhB3jzsyX9Cg
# jEi+sBC2rBMTMA0GCSqGSIb3DQEBCwUAMCAxHjAcBgNVBAMMFVBoaW5JVC1QU3Nj
# cmlwdHNfU2lnbjAeFw0yNTA3MDUwODI4MTZaFw0yNzA3MDUwODM4MTZaMCAxHjAc
# BgNVBAMMFVBoaW5JVC1QU3NjcmlwdHNfU2lnbjCCASIwDQYJKoZIhvcNAQEBBQAD
# ggEPADCCAQoCggEBALmz3o//iDA5MvAndTjGX7/AvzTSACClfuUR9WYK0f6Ut2dI
# mPxn+Y9pZlLjXIpZT0H2Lvxq5aSI+aYeFtuJ8/0lULYNCVT31Bf+HxervRBKsUyi
# W9+4PH6STxo3Pl4l56UNQMcWLPNjDORWRPWHn0f99iNtjI+L4tUC/LoWSs3obzxN
# 3uTypzlaPBxis2qFSTR5SWqFdZdRkcuI5LNsJjyc/QWdTYRrfmVqp0QrvcxzCv8u
# EiVuni6jkXfiE6wz+oeI3L2iR+ywmU6CUX4tPWoS9VTtmm7AhEpasRTmrrnSg20Q
# jiBa1eH5TyLAH3TcYMxhfMbN9a2xDX5pzM65EJUCAwEAAaNGMEQwDgYDVR0PAQH/
# BAQDAgeAMBMGA1UdJQQMMAoGCCsGAQUFBwMDMB0GA1UdDgQWBBQO7XOqiE/EYi+n
# IaR6YO5M2MUuVTANBgkqhkiG9w0BAQsFAAOCAQEAjYOKIwBu1pfbdvEFFaR/uY88
# peKPk0NnvNEc3dpGdOv+Fsgbz27JPvItITFd6AKMoN1W48YjQLaU22M2jdhjGN5i
# FSobznP5KgQCDkRsuoDKiIOTiKAAknjhoBaCCEZGw8SZgKJtWzbST36Thsdd/won
# ihLsuoLxfcFnmBfrXh3rTIvTwvfujob68s0Sf5derHP/F+nphTymlg+y4VTEAijk
# g2dhy8RAsbS2JYZT7K5aEJpPXMiOLBqd7oTGfM7y5sLk2LIM4cT8hzgz3v5yPMkF
# H2MdR//K403e1EKH9MsGuGAJZddVN8ppaiESoPLoXrgnw2SY5KCmhYw1xRFdjTCC
# BY0wggR1oAMCAQICEA6bGI750C3n79tQ4ghAGFowDQYJKoZIhvcNAQEMBQAwZTEL
# MAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3
# LmRpZ2ljZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBSb290
# IENBMB4XDTIyMDgwMTAwMDAwMFoXDTMxMTEwOTIzNTk1OVowYjELMAkGA1UEBhMC
# VVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0
# LmNvbTEhMB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MIICIjANBgkq
# hkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAv+aQc2jeu+RdSjwwIjBpM+zCpyUuySE9
# 8orYWcLhKac9WKt2ms2uexuEDcQwH/MbpDgW61bGl20dq7J58soR0uRf1gU8Ug9S
# H8aeFaV+vp+pVxZZVXKvaJNwwrK6dZlqczKU0RBEEC7fgvMHhOZ0O21x4i0MG+4g
# 1ckgHWMpLc7sXk7Ik/ghYZs06wXGXuxbGrzryc/NrDRAX7F6Zu53yEioZldXn1RY
# jgwrt0+nMNlW7sp7XeOtyU9e5TXnMcvak17cjo+A2raRmECQecN4x7axxLVqGDgD
# EI3Y1DekLgV9iPWCPhCRcKtVgkEy19sEcypukQF8IUzUvK4bA3VdeGbZOjFEmjNA
# vwjXWkmkwuapoGfdpCe8oU85tRFYF/ckXEaPZPfBaYh2mHY9WV1CdoeJl2l6SPDg
# ohIbZpp0yt5LHucOY67m1O+SkjqePdwA5EUlibaaRBkrfsCUtNJhbesz2cXfSwQA
# zH0clcOP9yGyshG3u3/y1YxwLEFgqrFjGESVGnZifvaAsPvoZKYz0YkH4b235kOk
# GLimdwHhD5QMIR2yVCkliWzlDlJRR3S+Jqy2QXXeeqxfjT/JvNNBERJb5RBQ6zHF
# ynIWIgnffEx1P2PsIV/EIFFrb7GrhotPwtZFX50g/KEexcCPorF+CiaZ9eRpL5gd
# LfXZqbId5RsCAwEAAaOCATowggE2MA8GA1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYE
# FOzX44LScV1kTN8uZz/nupiuHA9PMB8GA1UdIwQYMBaAFEXroq/0ksuCMS1Ri6en
# IZ3zbcgPMA4GA1UdDwEB/wQEAwIBhjB5BggrBgEFBQcBAQRtMGswJAYIKwYBBQUH
# MAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3aHR0cDov
# L2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNy
# dDBFBgNVHR8EPjA8MDqgOKA2hjRodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGln
# aUNlcnRBc3N1cmVkSURSb290Q0EuY3JsMBEGA1UdIAQKMAgwBgYEVR0gADANBgkq
# hkiG9w0BAQwFAAOCAQEAcKC/Q1xV5zhfoKN0Gz22Ftf3v1cHvZqsoYcs7IVeqRq7
# IviHGmlUIu2kiHdtvRoU9BNKei8ttzjv9P+Aufih9/Jy3iS8UgPITtAq3votVs/5
# 9PesMHqai7Je1M/RQ0SbQyHrlnKhSLSZy51PpwYDE3cnRNTnf+hZqPC/Lwum6fI0
# POz3A8eHqNJMQBk1RmppVLC4oVaO7KTVPeix3P0c2PR3WlxUjG/voVA9/HYJaISf
# b8rbII01YBwCA8sgsKxYoA5AY8WYIsGyWfVVa88nq2x2zm8jLfR+cWojayL/ErhU
# LSd+2DrZ8LaHlv1b0VysGMNNn3O3AamfV6peKOK5lDCCBrQwggScoAMCAQICEA3H
# rFcF/yGZLkBDIgw6SYYwDQYJKoZIhvcNAQELBQAwYjELMAkGA1UEBhMCVVMxFTAT
# BgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEh
# MB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MB4XDTI1MDUwNzAwMDAw
# MFoXDTM4MDExNDIzNTk1OVowaTELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lD
# ZXJ0LCBJbmMuMUEwPwYDVQQDEzhEaWdpQ2VydCBUcnVzdGVkIEc0IFRpbWVTdGFt
# cGluZyBSU0E0MDk2IFNIQTI1NiAyMDI1IENBMTCCAiIwDQYJKoZIhvcNAQEBBQAD
# ggIPADCCAgoCggIBALR4MdMKmEFyvjxGwBysddujRmh0tFEXnU2tjQ2UtZmWgyxU
# 7UNqEY81FzJsQqr5G7A6c+Gh/qm8Xi4aPCOo2N8S9SLrC6Kbltqn7SWCWgzbNfiR
# +2fkHUiljNOqnIVD/gG3SYDEAd4dg2dDGpeZGKe+42DFUF0mR/vtLa4+gKPsYfwE
# u7EEbkC9+0F2w4QJLVSTEG8yAR2CQWIM1iI5PHg62IVwxKSpO0XaF9DPfNBKS7Za
# zch8NF5vp7eaZ2CVNxpqumzTCNSOxm+SAWSuIr21Qomb+zzQWKhxKTVVgtmUPAW3
# 5xUUFREmDrMxSNlr/NsJyUXzdtFUUt4aS4CEeIY8y9IaaGBpPNXKFifinT7zL2gd
# FpBP9qh8SdLnEut/GcalNeJQ55IuwnKCgs+nrpuQNfVmUB5KlCX3ZA4x5HHKS+rq
# BvKWxdCyQEEGcbLe1b8Aw4wJkhU1JrPsFfxW1gaou30yZ46t4Y9F20HHfIY4/6vH
# espYMQmUiote8ladjS/nJ0+k6MvqzfpzPDOy5y6gqztiT96Fv/9bH7mQyogxG9QE
# PHrPV6/7umw052AkyiLA6tQbZl1KhBtTasySkuJDpsZGKdlsjg4u70EwgWbVRSX1
# Wd4+zoFpp4Ra+MlKM2baoD6x0VR4RjSpWM8o5a6D8bpfm4CLKczsG7ZrIGNTAgMB
# AAGjggFdMIIBWTASBgNVHRMBAf8ECDAGAQH/AgEAMB0GA1UdDgQWBBTvb1NK6eQG
# fHrK4pBW9i/USezLTjAfBgNVHSMEGDAWgBTs1+OC0nFdZEzfLmc/57qYrhwPTzAO
# BgNVHQ8BAf8EBAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwgwdwYIKwYBBQUHAQEE
# azBpMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQQYIKwYB
# BQUHMAKGNWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0
# ZWRSb290RzQuY3J0MEMGA1UdHwQ8MDowOKA2oDSGMmh0dHA6Ly9jcmwzLmRpZ2lj
# ZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRSb290RzQuY3JsMCAGA1UdIAQZMBcwCAYG
# Z4EMAQQCMAsGCWCGSAGG/WwHATANBgkqhkiG9w0BAQsFAAOCAgEAF877FoAc/gc9
# EXZxML2+C8i1NKZ/zdCHxYgaMH9Pw5tcBnPw6O6FTGNpoV2V4wzSUGvI9NAzaoQk
# 97frPBtIj+ZLzdp+yXdhOP4hCFATuNT+ReOPK0mCefSG+tXqGpYZ3essBS3q8nL2
# UwM+NMvEuBd/2vmdYxDCvwzJv2sRUoKEfJ+nN57mQfQXwcAEGCvRR2qKtntujB71
# WPYAgwPyWLKu6RnaID/B0ba2H3LUiwDRAXx1Neq9ydOal95CHfmTnM4I+ZI2rVQf
# jXQA1WSjjf4J2a7jLzWGNqNX+DF0SQzHU0pTi4dBwp9nEC8EAqoxW6q17r0z0noD
# js6+BFo+z7bKSBwZXTRNivYuve3L2oiKNqetRHdqfMTCW/NmKLJ9M+MtucVGyOxi
# Df06VXxyKkOirv6o02OoXN4bFzK0vlNMsvhlqgF2puE6FndlENSmE+9JGYxOGLS/
# D284NHNboDGcmWXfwXRy4kbu4QFhOm0xJuF2EZAOk5eCkhSxZON3rGlHqhpB/8Ml
# uDezooIs8CVnrpHMiD2wL40mm53+/j7tFaxYKIqL0Q4ssd8xHZnIn/7GELH3IdvG
# 2XlM9q7WP/UwgOkw/HQtyRN62JK4S1C8uw3PdBunvAZapsiI5YKdvlarEvf8EA+8
# hcpSM9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAKgO8YS43xBYLR
# xHanlXRoMA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5E
# aWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1l
# U3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjUwNjA0MDAwMDAw
# WhcNMzYwOTAzMjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNl
# cnQsIEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVz
# dGFtcCBSZXNwb25kZXIgMjAyNSAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIIC
# CgKCAgEA0EasLRLGntDqrmBWsytXum9R/4ZwCgHfyjfMGUIwYzKomd8U1nH7C8Dr
# 0cVMF3BsfAFI54um8+dnxk36+jx0Tb+k+87H9WPxNyFPJIDZHhAqlUPt281mHrBb
# ZHqRK71Em3/hCGC5KyyneqiZ7syvFXJ9A72wzHpkBaMUNg7MOLxI6E9RaUueHTQK
# WXymOtRwJXcrcTTPPT2V1D/+cFllESviH8YjoPFvZSjKs3SKO1QNUdFd2adw44wD
# cKgH+JRJE5Qg0NP3yiSyi5MxgU6cehGHr7zou1znOM8odbkqoK+lJ25LCHBSai25
# CFyD23DZgPfDrJJJK77epTwMP6eKA0kWa3osAe8fcpK40uhktzUd/Yk0xUvhDU6l
# vJukx7jphx40DQt82yepyekl4i0r8OEps/FNO4ahfvAk12hE5FVs9HVVWcO5J4dV
# mVzix4A77p3awLbr89A90/nWGjXMGn7FQhmSlIUDy9Z2hSgctaepZTd0ILIUbWuh
# KuAeNIeWrzHKYueMJtItnj2Q+aTyLLKLM0MheP/9w6CtjuuVHJOVoIJ/DtpJRE7C
# e7vMRHoRon4CWIvuiNN1Lk9Y+xZ66lazs2kKFSTnnkrT3pXWETTJkhd76CIDBbTR
# ofOsNyEhzZtCGmnQigpFHti58CSmvEyJcAlDVcKacJ+A9/z7eacCAwEAAaOCAZUw
# ggGRMAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFOQ7/PIx7f391/ORcWMZUEPPYYzo
# MB8GA1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIH
# gDAWBgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYI
# KwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZR
# aHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGlt
# ZVN0YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBS
# oFCGTmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRp
# bWVTdGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgG
# BmeBDAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAGUqrfEcJwS5
# rmBB7NEIRJ5jQHIh+OT2Ik/bNYulCrVvhREafBYF0RkP2AGr181o2YWPoSHz9iZE
# N/FPsLSTwVQWo2H62yGBvg7ouCODwrx6ULj6hYKqdT8wv2UV+Kbz/3ImZlJ7YXwB
# D9R0oU62PtgxOao872bOySCILdBghQ/ZLcdC8cbUUO75ZSpbh1oipOhcUT8lD8QA
# GB9lctZTTOJM3pHfKBAEcxQFoHlt2s9sXoxFizTeHihsQyfFg5fxUFEp7W42fNBV
# N4ueLaceRf9Cq9ec1v5iQMWTFQa0xNqItH3CPFTG7aEQJmmrJTV3Qhtfparz+BW6
# 0OiMEgV5GWoBy4RVPRwqxv7Mk0Sy4QHs7v9y69NBqycz0BZwhB9WOfOu/CIJnzkQ
# TwtSSpGGhLdjnQ4eBpjtP+XB3pQCtv4E5UCSDag6+iX8MmB10nfldPF9SVD7weCC
# 3yXZi/uuhqdwkgVxuiMFzGVFwYbQsiGnoa9F5AaAyBjFBtXVLcKtapnMG3VH3EmA
# p/jsJ3FVF3+d1SVDTmjFjLbNFZUWMXuZyvgLfgyPehwJVxwC+UpX2MSey2ueIu9T
# HFVkT+um1vshETaWyQo8gmBto/m3acaP9QsuLj3FNwFlTxq25+T4QwX9xa6ILs84
# ZPvmpovq90K8eWyG2N01c4IhSOxqt81nMYIFEDCCBQwCAQEwNDAgMR4wHAYDVQQD
# DBVQaGluSVQtUFNzY3JpcHRzX1NpZ24CEHePOzJf0KCMSL6wELasExMwDQYJYIZI
# AWUDBAIBBQCggYQwGAYKKwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkqhkiG9w0B
# CQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAv
# BgkqhkiG9w0BCQQxIgQgdfHcRIpCRzCI5+kcYnTWwZ/fV89XB0elUADFiIFwqOEw
# DQYJKoZIhvcNAQEBBQAEggEAb+aUl74OKMdcIV9gSrAOyUkemln+UDKoTLB9UYpe
# oHSic3pL4l4law3ejBNx76dS14yyA54h/5sgCdjsKElWsaVXod7oEw+cFmw/9svw
# 2KHvhsa2Pzwuhx+1I8XLGHq8xE0qun0PZZAockDfK3AgUyeVTV2RvkGSDS6D0gk5
# XS+T8ZCMZYei+gUGMfw8zsm5R2GhigYqILtM40N9gSjaxbZWrRInvSA20uLcO8CR
# iG9hzz3HmvMtEJeWhSJajmlfPE0O9rITbEUAxQv3eL8HBXc8n7/Ck0M4qVM4xWgK
# EB4kmmePaKvXI+bi/1p7h6kN6vWE52f69DxqANUlUowha6GCAyYwggMiBgkqhkiG
# 9w0BCQYxggMTMIIDDwIBATB9MGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgw
# DQYJYIZIAWUDBAIBBQCgaTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqG
# SIb3DQEJBTEPFw0yNTA3MTIwODAwMjRaMC8GCSqGSIb3DQEJBDEiBCBa620zOOYi
# PjGM/ir6FyaAshH03KoAWLwY5meeb9s3UjANBgkqhkiG9w0BAQEFAASCAgBF74pg
# 6eZUA5PqTD2jJulYw+nIW4owwOlqiDB86aizX7Ks1biQLronQkWHeu5DygN/68IO
# hXiMkwFLQPEML+KxvCZGhBSrAE9S4KxbbkSiwbvuRWudi9DGBEASVSDMeOTNGQDl
# y77yL1YJhrvmDovCa01iuVrQ22XV9K3VzDF9oc4IY4gZHW8I5/FKrvDIvGZ3gyZl
# +yxiwWRZAxHx6/ttg39TDUXVvnsFKElehrddxiOY3/v3GuEPc2FVPDwc2g7oBvw1
# rL9l70C5ZZqj0e+Kp22BOvHYMpa+Fm7Nj4G/Zc7iXfRVuOTGUYCZSN7eWft4zwAW
# jNbXncXfi4RNARHKAK7VILKGEJQ3u16xXi+uPZ0QV1n7AveuEreQqS40QE0O9FTU
# fWEotWwHiqKEVcPhIT9tedBbULMVIbihX7+M/wUAepxrtL/MwxkRMQMhtZ8Zxxdk
# 63+fBqTB/fUL2gch/ahQFe9wzNIdg/whvpWoEsCyuhsn2gft7xzbhYSIRfoRcytp
# 6JhduRkYMInxnLKcqoNekhD6IhmC41IqfkJiWyVBWw8KjEdSOI16vavi5aCjQc8K
# n0836GqbL5xGN3BOJRnnyRoxOUdYaGFomKiIiw+gh68Ztp/tnNqyv6xY7HHPLCKZ
# dYb5/bTERM38CfDkrOf2K0s68LF7OScPe+pluA==
# SIG # End signature block
