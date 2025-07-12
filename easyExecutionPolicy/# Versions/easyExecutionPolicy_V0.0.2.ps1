# Assembly-Loading für Windows Forms und Drawing
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Funktion zur Überprüfung und Anforderung von Administratorrechten
function Ensure-Admin {
    # Prüfen, ob das Skript bereits mit Administratorrechten ausgeführt wird
    if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        # Dialogfenster anzeigen, um den Benutzer zu fragen, ob das Skript als Administrator neu gestartet werden soll
        $dialogResult = [System.Windows.Forms.MessageBox]::Show("Dieses Skript erfordert Administratorrechte. Möchten Sie es als Administrator neu starten?", "Administratorrechte erforderlich", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)

        if ($dialogResult -eq [System.Windows.Forms.DialogResult]::Yes) {
            # Versuch, das Skript mit Administratorrechten neu zu starten
            try {
                # Parameter für den Neustart vorbereiten
                $powershellPath = (Get-Command PowerShell).Source
                $argumentList = "-File `"$($MyInvocation.MyCommand.Path)`""
                
                # Prozess als Administrator starten
                Start-Process -FilePath $powershellPath -Verb RunAs -ArgumentList $argumentList -ErrorAction Stop
                
                # Aktuelle (nicht-administrative) Instanz beenden
                Write-Host "Die nicht-administrative Instanz wird beendet, da eine neue administrative Instanz gestartet wurde."
                exit
            } catch {
                # Fehlermeldung anzeigen, wenn der Neustart fehlschlägt, und Skript beenden
                [System.Windows.Forms.MessageBox]::Show("Fehler beim Neustart des Skripts als Administrator:`n$($_.Exception.Message)", "Fehler beim Neustart", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                Write-Host "Fehler beim Neustart als Administrator. Das Skript wird beendet."
                exit
            }
        } else {
            # Informationsmeldung anzeigen, wenn der Benutzer "Nein" wählt, und Skript beenden
            [System.Windows.Forms.MessageBox]::Show("Das Skript kann ohne Administratorrechte nicht fortfahren und wird nun beendet.", "Administratorrechte verweigert", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            Write-Host "Benutzer hat den Neustart als Administrator abgelehnt. Das Skript wird beendet."
            exit
        }
    }
}

Ensure-Admin

# GUI Creation
$mainForm = New-Object System.Windows.Forms.Form
$mainForm.Text = "Execution Policy Manager"
$mainForm.Size = New-Object System.Drawing.Size(820, 400) # Reduziert von 500 auf 400
$mainForm.StartPosition = "CenterScreen"
$mainForm.FormBorderStyle = 'FixedDialog'
$mainForm.MaximizeBox = $false
$mainForm.MinimizeBox = $false
$mainForm.BackColor = [System.Drawing.Color]::White

# Define Fonts based on XAML
$uiFontFamily = "Segoe UI"
$uiFontSizeNormal = 12
$uiFontSizeLarge = 20

$fontNormal = New-Object System.Drawing.Font($uiFontFamily, $uiFontSizeNormal)
$fontBold = New-Object System.Drawing.Font($uiFontFamily, $uiFontSizeNormal, [System.Drawing.FontStyle]::Bold)
$fontHeader = New-Object System.Drawing.Font($uiFontFamily, $uiFontSizeLarge, [System.Drawing.FontStyle]::Bold)

# Farbkonstanten definieren
$colorDarkText = [System.Drawing.Color]::FromArgb(28, 28, 28)
$colorAccent = [System.Drawing.Color]::FromArgb(0, 120, 212)
$colorSuccess = [System.Drawing.Color]::FromArgb(16, 124, 16)
$colorWarning = [System.Drawing.Color]::FromArgb(255, 140, 0)
$colorGray = [System.Drawing.Color]::FromArgb(108, 117, 125)

# Header Panel
$headerPanel = New-Object System.Windows.Forms.Panel
$headerPanel.Size = New-Object System.Drawing.Size($mainForm.ClientSize.Width, 60)
$headerPanel.Location = New-Object System.Drawing.Point(0, 0)
$headerPanel.BackColor = [System.Drawing.Color]::FromArgb(255, 28, 50, 60) # Ersetzt ColorTranslator
$headerPanel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$mainForm.Controls.Add($headerPanel)

$headerLabel = New-Object System.Windows.Forms.Label
$headerLabel.Text = "Execution Policy Manager"
$headerLabel.Font = $fontHeader
$headerLabel.ForeColor = [System.Drawing.Color]::FromArgb(226, 226, 226) # Ersetzt ColorTranslator
$headerLabel.AutoSize = $true
$headerLabel.Location = New-Object System.Drawing.Point(20, [Math]::Round(($headerPanel.Height - 25) / 2)) # Height durch geschätzte Label-Höhe
$headerPanel.Controls.Add($headerLabel)

# Content Panel (um die Y-Positionierung zu vereinfachen)
$contentPanel = New-Object System.Windows.Forms.Panel
$contentPanel.Location = New-Object System.Drawing.Point(0, $headerPanel.Height)
$contentPanel.Size = New-Object System.Drawing.Size($mainForm.ClientSize.Width, ($mainForm.ClientSize.Height - $headerPanel.Height - 50)) # Reduziert von 70 auf 50
$contentPanel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$contentPanel.AutoScroll = $true # Falls Inhalt größer wird
$mainForm.Controls.Add($contentPanel)

# Footer Panel
$footerPanel = New-Object System.Windows.Forms.Panel
$footerPanel.Size = New-Object System.Drawing.Size($mainForm.ClientSize.Width, 50) # Reduziert von 70 auf 50
$footerPanel.Location = New-Object System.Drawing.Point(0, ($mainForm.ClientSize.Height - $footerPanel.Height))
$footerPanel.BackColor = [System.Drawing.Color]::White
$footerPanel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$mainForm.Controls.Add($footerPanel)

# ToolTip Object for providing hints
$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.AutoPopDelay = 10000 # Längere Anzeigedauer für Tooltips
$toolTip.InitialDelay = 500   # Kürzere Verzögerung bis zur Anzeige

# Define Tooltip Texts (Beispielhaft, bitte anpassen/erweitern)
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


$global:currentPolicies = @{}
$policyOptions = @("Restricted", "AllSigned", "RemoteSigned", "Unrestricted", "Bypass", "Undefined")

# Function to get current Execution Policies
function Get-CurrentExecutionPolicies {
    $scopes = @("MachinePolicy", "UserPolicy", "Process", "CurrentUser", "LocalMachine")
    foreach ($scope in $scopes) {
        try {
            $policy = Get-ExecutionPolicy -Scope $scope -ErrorAction SilentlyContinue 
            if ($null -eq $policy) { # Handles cases where policy is explicitly 'Undefined' or not set, Get-ExecutionPolicy might return $null
                $global:currentPolicies[$scope] = "Undefined"
            } else {
                $global:currentPolicies[$scope] = $policy.ToString()
            }
        } catch {
            $global:currentPolicies[$scope] = "Error" # Error retrieving specific scope
        }
    }
    # Get effective policy for the current session (overall)
    try {
        $global:currentPolicies["Effective"] = (Get-ExecutionPolicy).ToString()
    } catch {
        $global:currentPolicies["Effective"] = "Error"
    }
}


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
            $label.ForeColor = [System.Drawing.Color]::FromArgb(28, 28, 28) # Direkte Farbe statt Variable
            if ($policyExplanationTooltips.ContainsKey($currentPolicyValue)) {
                $toolTip.SetToolTip($label, $policyExplanationTooltips[$currentPolicyValue])
            } else {
                $toolTip.SetToolTip($label, "Keine spezifische Erklärung für diese Richtlinie verfügbar.")
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
            $lblEffectivePolicyValue.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 212) # Direkte Farbe
            if ($policyExplanationTooltips.ContainsKey($effectivePolicy.ToString())) {
                $toolTip.SetToolTip($lblEffectivePolicyValue, $policyExplanationTooltips[$effectivePolicy.ToString()])
            } else {
                $toolTip.SetToolTip($lblEffectivePolicyValue, "Der Wert der aktuell wirksamen Ausführungsrichtlinie.")
            }
        }
    } catch {
        $lblEffectivePolicyValue = $contentPanel.Controls.Find("lblEffectivePolicyValue", $true)[0]
        if ($lblEffectivePolicyValue) {
            $lblEffectivePolicyValue.Text = "Error retrieving"
            $lblEffectivePolicyValue.ForeColor = [System.Drawing.Color]::Red
            $toolTip.SetToolTip($lblEffectivePolicyValue, "Fehler beim Abrufen der effektiven Richtlinie.")
        }
    }
}

# Create GUI elements - Labels und ComboBoxes mit optimiertem Layout
$yPos = 25
$leftMargin = 20
$column1Width = 350  # Reduziert für bessere Proportion
$column2Width = 180  # Reduziert
$column3Width = 200  # Neue Spalte für ComboBox
$spacing = 25       # Konsistenter Abstand

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

    # Label for Scope name (Spalte 1)
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

    # Label for current Policy (Spalte 2)
    $lblCurrentPolicy = New-Object System.Windows.Forms.Label
    $lblCurrentPolicy.Name = "lblCurrent$scopeName"
    $lblCurrentPolicy.Text = "loading..."
    $lblCurrentPolicy.Location = New-Object System.Drawing.Point($column2XPos, $yPos)
    $lblCurrentPolicy.Size = New-Object System.Drawing.Size($column2Width, 20)
    $lblCurrentPolicy.Font = $fontBold
    $lblCurrentPolicy.ForeColor = $colorDarkText
    $lblCurrentPolicy.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $contentPanel.Controls.Add($lblCurrentPolicy)

    # ComboBox or placeholder (Spalte 3)
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
    $yPos += 31  # Reduziert von 35 auf 25
}

# Separator Line - minimale Abstände
$yPos += 10  # Reduziert von 5 auf 2
$separatorLine = New-Object System.Windows.Forms.Label
$separatorLine.Location = New-Object System.Drawing.Point($leftMargin, $yPos)
$separatorLine.Size = New-Object System.Drawing.Size(($column1Width + $column2Width + $column3Width + 2 * $spacing), 1)
$separatorLine.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
$contentPanel.Controls.Add($separatorLine)
$yPos += 10  # Reduziert von 10 auf 5

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

# Status Label - korrekt positioniert mit genug Platz
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Name = "statusLabel"
$statusLabel.Text = "Ready."
$statusLabel.Location = New-Object System.Drawing.Point($leftMargin, 1)
$statusLabel.Size = New-Object System.Drawing.Size(($footerPanel.Width - 2 * $leftMargin), 13) # Feste Höhe von 16px
$statusLabel.Font = $fontNormal
$statusLabel.ForeColor = $colorSuccess
$statusLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$footerPanel.Controls.Add($statusLabel)

# Button to apply changes - unter dem Status-Text
$applyButton = New-Object System.Windows.Forms.Button
$applyButton.Text = "APPLY Selected Policies"
$applyButton.Location = New-Object System.Drawing.Point(($footerPanel.Width - 250 - $leftMargin), 22) # Y-Position angepasst
$applyButton.Size = New-Object System.Drawing.Size(250, 24) # Höhe reduziert
$applyButton.Font = $fontBold
$applyButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$applyButton.FlatAppearance.BorderSize = 0
$applyButton.BackColor = $colorAccent
$applyButton.ForeColor = [System.Drawing.Color]::White
$applyButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
$footerPanel.Controls.Add($applyButton)

# Button to refresh - unter dem Status-Text
$refreshButton = New-Object System.Windows.Forms.Button
$refreshButton.Text = "Refresh"
$refreshButton.Location = New-Object System.Drawing.Point($leftMargin, 22) # Y-Position angepasst
$refreshButton.Size = New-Object System.Drawing.Size(120, 24) # Höhe reduziert
$refreshButton.Font = $fontNormal
$refreshButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$refreshButton.FlatAppearance.BorderSize = 0
$refreshButton.BackColor = $colorGray
$refreshButton.ForeColor = [System.Drawing.Color]::White
$refreshButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$footerPanel.Controls.Add($refreshButton)

# ToolTip für Footer-Buttons
$toolTip.SetToolTip($applyButton, "Wenden Sie die ausgewählten Richtlinienänderungen an.")
$toolTip.SetToolTip($refreshButton, "Aktualisieren Sie die angezeigten Richtlinienwerte.")

# Event Handler für Refresh Button
$refreshButton.Add_Click({
    $statusLabel.Text = "Refreshing policies..."
    $statusLabel.ForeColor = $colorWarning
    Update-PolicyDisplay
    $statusLabel.Text = "Policies refreshed."
    $statusLabel.ForeColor = $colorSuccess
})

# Event Handler für Apply Button
$applyButton.Add_Click({
    $statusLabel.Text = "Applying selected policies..."
    $statusLabel.ForeColor = $colorWarning

    $changesMade = $false
    $errorOccurred = $false
    $errorMessagesList = [System.Collections.Generic.List[string]]::new()

    $scopesToManage | Where-Object {$_.Settable} | ForEach-Object {
        $scopeName = $_.Name
        $comboBox = $contentPanel.Controls.Find("cb$scopeName", $true)[0]
        if ($comboBox -and $comboBox.SelectedItem) {
            $selectedPolicy = $comboBox.SelectedItem.ToString()
            $currentPolicy = $global:currentPolicies[$scopeName] 

            if ($selectedPolicy -ne $currentPolicy) {
                try {
                    Write-Host "Setting ExecutionPolicy for Scope '$scopeName' to '$selectedPolicy'..."
                    if ($selectedPolicy -eq "Undefined") {
                        # Optimierte Undefined-Behandlung
                        if ($scopeName -eq "Process") {
                            Set-ExecutionPolicy -ExecutionPolicy Undefined -Scope $scopeName -Force -ErrorAction Stop
                        } else {
                            # Registry-basierte Behandlung für CurrentUser/LocalMachine
                            $registryPath = switch ($scopeName) {
                                "CurrentUser" { "HKCU:\SOFTWARE\Microsoft\PowerShell\1\ShellIds\Microsoft.PowerShell" }
                                "LocalMachine" { "HKLM:\SOFTWARE\Microsoft\PowerShell\1\ShellIds\Microsoft.PowerShell" }
                                default { "" }
                            }
                            
                            if ($registryPath -and (Test-Path $registryPath)) {
                                if (Get-ItemProperty -Path $registryPath -Name "ExecutionPolicy" -ErrorAction SilentlyContinue) {
                                    Remove-ItemProperty -Path $registryPath -Name "ExecutionPolicy" -Force -ErrorAction Stop
                                    Write-Host "Removed ExecutionPolicy registry value for Scope '$scopeName'."
                                } else {
                                    Write-Host "No ExecutionPolicy registry value to remove for Scope '$scopeName'."
                                }
                            }
                        }
                    } else {
                        Set-ExecutionPolicy -ExecutionPolicy $selectedPolicy -Scope $scopeName -Force -ErrorAction Stop
                    }
                    $changesMade = $true
                    Write-Host "Successfully processed Scope '$scopeName' for policy '$selectedPolicy'."
                } catch {
                    $errorMessage = "Error for Scope ${scopeName} ('$selectedPolicy'): $($_.Exception.Message.Split([Environment]::NewLine)[0])"
                    $errorMessagesList.Add($errorMessage)
                    if (-not $errorOccurred) {
                        $statusLabel.Text = $errorMessage 
                        $statusLabel.ForeColor = [System.Drawing.Color]::Red
                    }
                    $errorOccurred = $true
                    Write-Warning "Error setting ExecutionPolicy for Scope '$scopeName' to '$selectedPolicy': $($_.Exception.Message)"
                }
            }
        }
    }

    Update-PolicyDisplay

    if ($errorOccurred) {
        if ($errorMessagesList.Count -gt 1) {
            $statusLabel.Text = "Multiple errors occurred. Check console for details."
        }
        $statusLabel.ForeColor = [System.Drawing.Color]::Red
    } elseif ($changesMade) {
        $statusLabel.Text = "Execution policies successfully applied."
        $statusLabel.ForeColor = $colorSuccess
    } else {
        $statusLabel.Text = "No changes were made to execution policies."
        $statusLabel.ForeColor = $colorAccent
    }
})

# Function to update GUI display - verschoben nach GUI-Erstellung
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
            } else {
                $toolTip.SetToolTip($label, "Keine spezifische Erklärung verfügbar.")
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
            } else {
                $toolTip.SetToolTip($lblEffectivePolicyValue, "Aktuell wirksame Ausführungsrichtlinie.")
            }
        }
    } catch {
        $lblEffectivePolicyValue = $contentPanel.Controls.Find("lblEffectivePolicyValue", $true)[0]
        if ($lblEffectivePolicyValue) {
            $lblEffectivePolicyValue.Text = "Error retrieving"
            $lblEffectivePolicyValue.ForeColor = [System.Drawing.Color]::Red
            $toolTip.SetToolTip($lblEffectivePolicyValue, "Fehler beim Abrufen der Richtlinie.")
        }
    }
}

# On Form Load
$mainForm.Add_Load({
    $mainForm.BeginInvoke([Action]{ Update-PolicyDisplay })
})

# Show GUI
$mainForm.ShowDialog() | Out-Null

# Clean up
$mainForm.Dispose()
# SIG # Begin signature block
# MIIbywYJKoZIhvcNAQcCoIIbvDCCG7gCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBAqzPoeixvSHie
# I7XuAwiqFGvqkuLHI9XUqDo9i0meLqCCFhcwggMQMIIB+KADAgECAhB3jzsyX9Cg
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
# LSd+2DrZ8LaHlv1b0VysGMNNn3O3AamfV6peKOK5lDCCBq4wggSWoAMCAQICEAc2
# N7ckVHzYR6z9KGYqXlswDQYJKoZIhvcNAQELBQAwYjELMAkGA1UEBhMCVVMxFTAT
# BgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEh
# MB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MB4XDTIyMDMyMzAwMDAw
# MFoXDTM3MDMyMjIzNTk1OVowYzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lD
# ZXJ0LCBJbmMuMTswOQYDVQQDEzJEaWdpQ2VydCBUcnVzdGVkIEc0IFJTQTQwOTYg
# U0hBMjU2IFRpbWVTdGFtcGluZyBDQTCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCC
# AgoCggIBAMaGNQZJs8E9cklRVcclA8TykTepl1Gh1tKD0Z5Mom2gsMyD+Vr2EaFE
# FUJfpIjzaPp985yJC3+dH54PMx9QEwsmc5Zt+FeoAn39Q7SE2hHxc7Gz7iuAhIoi
# GN/r2j3EF3+rGSs+QtxnjupRPfDWVtTnKC3r07G1decfBmWNlCnT2exp39mQh0YA
# e9tEQYncfGpXevA3eZ9drMvohGS0UvJ2R/dhgxndX7RUCyFobjchu0CsX7LeSn3O
# 9TkSZ+8OpWNs5KbFHc02DVzV5huowWR0QKfAcsW6Th+xtVhNef7Xj3OTrCw54qVI
# 1vCwMROpVymWJy71h6aPTnYVVSZwmCZ/oBpHIEPjQ2OAe3VuJyWQmDo4EbP29p7m
# O1vsgd4iFNmCKseSv6De4z6ic/rnH1pslPJSlRErWHRAKKtzQ87fSqEcazjFKfPK
# qpZzQmiftkaznTqj1QPgv/CiPMpC3BhIfxQ0z9JMq++bPf4OuGQq+nUoJEHtQr8F
# nGZJUlD0UfM2SU2LINIsVzV5K6jzRWC8I41Y99xh3pP+OcD5sjClTNfpmEpYPtMD
# iP6zj9NeS3YSUZPJjAw7W4oiqMEmCPkUEBIDfV8ju2TjY+Cm4T72wnSyPx4Jduyr
# XUZ14mCjWAkBKAAOhFTuzuldyF4wEr1GnrXTdrnSDmuZDNIztM2xAgMBAAGjggFd
# MIIBWTASBgNVHRMBAf8ECDAGAQH/AgEAMB0GA1UdDgQWBBS6FtltTYUvcyl2mi91
# jGogj57IbzAfBgNVHSMEGDAWgBTs1+OC0nFdZEzfLmc/57qYrhwPTzAOBgNVHQ8B
# Af8EBAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwgwdwYIKwYBBQUHAQEEazBpMCQG
# CCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQQYIKwYBBQUHMAKG
# NWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRSb290
# RzQuY3J0MEMGA1UdHwQ8MDowOKA2oDSGMmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNv
# bS9EaWdpQ2VydFRydXN0ZWRSb290RzQuY3JsMCAGA1UdIAQZMBcwCAYGZ4EMAQQC
# MAsGCWCGSAGG/WwHATANBgkqhkiG9w0BAQsFAAOCAgEAfVmOwJO2b5ipRCIBfmbW
# 2CFC4bAYLhBNE88wU86/GPvHUF3iSyn7cIoNqilp/GnBzx0H6T5gyNgL5Vxb122H
# +oQgJTQxZ822EpZvxFBMYh0MCIKoFr2pVs8Vc40BIiXOlWk/R3f7cnQU1/+rT4os
# equFzUNf7WC2qk+RZp4snuCKrOX9jLxkJodskr2dfNBwCnzvqLx1T7pa96kQsl3p
# /yhUifDVinF2ZdrM8HKjI/rAJ4JErpknG6skHibBt94q6/aesXmZgaNWhqsKRcnf
# xI2g55j7+6adcq/Ex8HBanHZxhOACcS2n82HhyS7T6NJuXdmkfFynOlLAlKnN36T
# U6w7HQhJD5TNOXrd/yVjmScsPT9rp/Fmw0HNT7ZAmyEhQNC3EyTN3B14OuSereU0
# cZLXJmvkOHOrpgFPvT87eK1MrfvElXvtCl8zOYdBeHo46Zzh3SP9HSjTx/no8Zhf
# +yvYfvJGnXUsHicsJttvFXseGYs2uJPU5vIXmVnKcPA3v5gA3yAWTyf7YGcWoWa6
# 3VXAOimGsJigK+2VQbc61RWYMbRiCQ8KvYHZE/6/pNHzV9m8BPqC3jLfBInwAM1d
# wvnQI38AC+R2AibZ8GV2QqYphwlHK+Z/GqSFD/yYlvZVVCsfgPrA8g4r5db7qS9E
# FUrnEw4d2zc4GqEr9u3WfPwwgga8MIIEpKADAgECAhALrma8Wrp/lYfG+ekE4zME
# MA0GCSqGSIb3DQEBCwUAMGMxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2Vy
# dCwgSW5jLjE7MDkGA1UEAxMyRGlnaUNlcnQgVHJ1c3RlZCBHNCBSU0E0MDk2IFNI
# QTI1NiBUaW1lU3RhbXBpbmcgQ0EwHhcNMjQwOTI2MDAwMDAwWhcNMzUxMTI1MjM1
# OTU5WjBCMQswCQYDVQQGEwJVUzERMA8GA1UEChMIRGlnaUNlcnQxIDAeBgNVBAMT
# F0RpZ2lDZXJ0IFRpbWVzdGFtcCAyMDI0MIICIjANBgkqhkiG9w0BAQEFAAOCAg8A
# MIICCgKCAgEAvmpzn/aVIauWMLpbbeZZo7Xo/ZEfGMSIO2qZ46XB/QowIEMSvgjE
# dEZ3v4vrrTHleW1JWGErrjOL0J4L0HqVR1czSzvUQ5xF7z4IQmn7dHY7yijvoQ7u
# jm0u6yXF2v1CrzZopykD07/9fpAT4BxpT9vJoJqAsP8YuhRvflJ9YeHjes4fduks
# THulntq9WelRWY++TFPxzZrbILRYynyEy7rS1lHQKFpXvo2GePfsMRhNf1F41nyE
# g5h7iOXv+vjX0K8RhUisfqw3TTLHj1uhS66YX2LZPxS4oaf33rp9HlfqSBePejlY
# eEdU740GKQM7SaVSH3TbBL8R6HwX9QVpGnXPlKdE4fBIn5BBFnV+KwPxRNUNK6lY
# k2y1WSKour4hJN0SMkoaNV8hyyADiX1xuTxKaXN12HgR+8WulU2d6zhzXomJ2Ple
# I9V2yfmfXSPGYanGgxzqI+ShoOGLomMd3mJt92nm7Mheng/TBeSA2z4I78JpwGpT
# RHiT7yHqBiV2ngUIyCtd0pZ8zg3S7bk4QC4RrcnKJ3FbjyPAGogmoiZ33c1HG93V
# p6lJ415ERcC7bFQMRbxqrMVANiav1k425zYyFMyLNyE1QulQSgDpW9rtvVcIH7Wv
# G9sqYup9j8z9J1XqbBZPJ5XLln8mS8wWmdDLnBHXgYly/p1DhoQo5fkCAwEAAaOC
# AYswggGHMA4GA1UdDwEB/wQEAwIHgDAMBgNVHRMBAf8EAjAAMBYGA1UdJQEB/wQM
# MAoGCCsGAQUFBwMIMCAGA1UdIAQZMBcwCAYGZ4EMAQQCMAsGCWCGSAGG/WwHATAf
# BgNVHSMEGDAWgBS6FtltTYUvcyl2mi91jGogj57IbzAdBgNVHQ4EFgQUn1csA3cO
# KBWQZqVjXu5Pkh92oFswWgYDVR0fBFMwUTBPoE2gS4ZJaHR0cDovL2NybDMuZGln
# aWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0UlNBNDA5NlNIQTI1NlRpbWVTdGFt
# cGluZ0NBLmNybDCBkAYIKwYBBQUHAQEEgYMwgYAwJAYIKwYBBQUHMAGGGGh0dHA6
# Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBYBggrBgEFBQcwAoZMaHR0cDovL2NhY2VydHMu
# ZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0UlNBNDA5NlNIQTI1NlRpbWVT
# dGFtcGluZ0NBLmNydDANBgkqhkiG9w0BAQsFAAOCAgEAPa0eH3aZW+M4hBJH2UOR
# 9hHbm04IHdEoT8/T3HuBSyZeq3jSi5GXeWP7xCKhVireKCnCs+8GZl2uVYFvQe+p
# PTScVJeCZSsMo1JCoZN2mMew/L4tpqVNbSpWO9QGFwfMEy60HofN6V51sMLMXNTL
# fhVqs+e8haupWiArSozyAmGH/6oMQAh078qRh6wvJNU6gnh5OruCP1QUAvVSu4kq
# VOcJVozZR5RRb/zPd++PGE3qF1P3xWvYViUJLsxtvge/mzA75oBfFZSbdakHJe2B
# VDGIGVNVjOp8sNt70+kEoMF+T6tptMUNlehSR7vM+C13v9+9ZOUKzfRUAYSyyEmY
# tsnpltD/GWX8eM70ls1V6QG/ZOB6b6Yum1HvIiulqJ1Elesj5TMHq8CWT/xrW7tw
# ipXTJ5/i5pkU5E16RSBAdOp12aw8IQhhA/vEbFkEiF2abhuFixUDobZaA0VhqAsM
# HOmaT3XThZDNi5U2zHKhUs5uHHdG6BoQau75KiNbh0c+hatSF+02kULkftARjsyE
# pHKsF7u5zKRbt5oK5YGwFvgc4pEVUNytmB3BpIiowOIIuDgP5M9WArHYSAR16gc0
# dP2XdkMEP5eBsX7bf/MGN4K3HP50v/01ZHo/Z5lGLvNwQ7XHBx1yomzLP8lx4Q1z
# ZKDyHcp4VQJLu2kWTsKsOqQxggUKMIIFBgIBATA0MCAxHjAcBgNVBAMMFVBoaW5J
# VC1QU3NjcmlwdHNfU2lnbgIQd487Ml/QoIxIvrAQtqwTEzANBglghkgBZQMEAgEF
# AKCBhDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgor
# BgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3
# DQEJBDEiBCAgG6qokiqCxpNRgz8N700wvf6/fVzkeCxjnDuuVSOSSDANBgkqhkiG
# 9w0BAQEFAASCAQCNhdQsyigJa5bJtm55KibDSHx+IGq5HgnHRorhUvX8IVTJgzgW
# UQoIHXuFHZS0DUpwtDFNQpJis+0WgHjhJAOCc5BAavZhdDf/RNhIxOVKFyW8tCZQ
# r8zr5ZnLlrNs0hvy8v2FPhIDssPcxnhl00ZXKEL9gck0Mc1Z6SaDJMB0V7IsMN/v
# 5TITWvQhaFG4+jUqOCWjMpiz5hbLfzqtUCp9EpW4dPQWRaJ3y7HE31XT1Z0L72xf
# ySEW51vj3C1f3HGu3pbust2DdZ7JyUXFQ4xSXRaMeA0rdl3YDOG+PRfcMsj4kAQI
# UFLZLHyZrz8MOXGCgGrqw7x8s84Q0RPAz6k1oYIDIDCCAxwGCSqGSIb3DQEJBjGC
# Aw0wggMJAgEBMHcwYzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJ
# bmMuMTswOQYDVQQDEzJEaWdpQ2VydCBUcnVzdGVkIEc0IFJTQTQwOTYgU0hBMjU2
# IFRpbWVTdGFtcGluZyBDQQIQC65mvFq6f5WHxvnpBOMzBDANBglghkgBZQMEAgEF
# AKBpMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTI1
# MDcwNTEwMTIwOVowLwYJKoZIhvcNAQkEMSIEIH1yJAOmZpOxzJ2YGzHWSME91VkU
# 74FGhH1hSg4+nk8WMA0GCSqGSIb3DQEBAQUABIICALfUEcjfKB27EKIj1x5Gmft/
# wvflJCqZP9rwCaiR64GacS4AtWTT3rWe76BswGWRrHH7vSzTuFZwGBBGNkeRYDD+
# s8rtfe7/7xohuTfNDnDKpVhxXJ/zMbKqXcGL46FwNoUgOG29Yb7iUfPuJivxe+cg
# W5mMq9PvFm/zgyunSDQSudZeLV8FmTBnbLHGAWhnqneA5KCRHknXmSsjRsAiDF+w
# PZp+nGKLOEiN8XWHchlsrqhk/LIZvvS7a03Q4pvH67HERXh7VOpmfHHIDL+wJR7n
# GKn0K/FUsLD1fbcqku8/0qJmk1cqDvyWRVrLsmCcbBnLtY8YONsx0Y8eYnz3jrfl
# zOGmmuYBsy/YuSJx6D3JS9f4iASVEMpHRmezv63ICo90w/gVvsbX4RbTjX12vZcb
# 0rzErPBTDHRPKIyLZ8SgZaF8s45k7vbfKySkTRqO8E2m/sJ5TgUkONxElpkd7RCd
# VmyNarPVHKUUKW4B0/JeYSKA0mPkkOn3Dzqd3tqj2Gi84Ch9mYvut5aHP4Va7Ihu
# 4WI1HHsPDLE//Hq/mpLP9/wWB8ZEFactnCsxqln8FwQPY1YYGj201JCjtB3tdWjI
# 3ndC/igZibDcPTnVW+M8Ere2G7jU22Uk3p595XSveoXOckPqUpUye0RSV6NYgHdi
# 8fe2IwPdKrieHV3U4yyt
# SIG # End signature block
