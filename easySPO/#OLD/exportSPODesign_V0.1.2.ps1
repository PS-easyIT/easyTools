Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Get current execution policy for use in the form
$epCurrent = Get-ExecutionPolicy -Scope Process -ErrorAction SilentlyContinue

# Check if module is installed for use in the form
$moduleName = "PnP.PowerShell"
$moduleInstalled = Get-Module -ListAvailable -Name $moduleName

function Show-ConfigForm {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "SharePoint Design Settings"
    $form.StartPosition = "CenterScreen"
    $form.Size = New-Object System.Drawing.Size(700, 620)  # Increased width for side-by-side sections
    $form.Topmost = $true
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false

    # Execution Policy Section
    $groupExecPolicy = New-Object System.Windows.Forms.GroupBox
    $groupExecPolicy.Text = "Execution Policy"
    $groupExecPolicy.Location = New-Object System.Drawing.Point(20, 20)
    $groupExecPolicy.Size = New-Object System.Drawing.Size(660, 60)
    $form.Controls.Add($groupExecPolicy)

    $lblCurrentPolicy = New-Object System.Windows.Forms.Label
    $lblCurrentPolicy.Text = "Current Policy: $epCurrent"
    $lblCurrentPolicy.AutoSize = $true
    $lblCurrentPolicy.Location = New-Object System.Drawing.Point(10, 25)
    $groupExecPolicy.Controls.Add($lblCurrentPolicy)

    $chkRemoteSigned = New-Object System.Windows.Forms.CheckBox
    $chkRemoteSigned.Text = "Set RemoteSigned"
    $chkRemoteSigned.Checked = ($epCurrent -eq "RemoteSigned")
    $chkRemoteSigned.Location = New-Object System.Drawing.Point(300, 25)
    $chkRemoteSigned.AutoSize = $true
    $groupExecPolicy.Controls.Add($chkRemoteSigned)

    # Left Panel - Module Installation Section
    $groupModule = New-Object System.Windows.Forms.GroupBox
    $groupModule.Text = "PnP.PowerShell module"
    $groupModule.Location = New-Object System.Drawing.Point(20, 90)
    $groupModule.Size = New-Object System.Drawing.Size(330, 100)
    $form.Controls.Add($groupModule)

    $lblModuleStatus = New-Object System.Windows.Forms.Label
    if ($moduleInstalled) {
        $lblModuleStatus.Text = "MODULE INSTALLED"
        $lblModuleStatus.ForeColor = [System.Drawing.Color]::Green
    } else {
        $lblModuleStatus.Text = "PLEASE INSTALL MODULE!!!"
        $lblModuleStatus.ForeColor = [System.Drawing.Color]::Red
    }
    $lblModuleStatus.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
    $lblModuleStatus.AutoSize = $true
    $lblModuleStatus.Location = New-Object System.Drawing.Point(10, 25)
    $groupModule.Controls.Add($lblModuleStatus)

    $btnInstallModule = New-Object System.Windows.Forms.Button
    $btnInstallModule.Text = "Install Module"
    $btnInstallModule.Location = New-Object System.Drawing.Point(190, 55)
    $btnInstallModule.Width = 130
    $btnInstallModule.Height = 30
    $btnInstallModule.Enabled = !$moduleInstalled
    $groupModule.Controls.Add($btnInstallModule)

    # Right Panel - Login with MFA
    $groupMFA = New-Object System.Windows.Forms.GroupBox
    $groupMFA.Text = "Login with MFA"
    $groupMFA.Location = New-Object System.Drawing.Point(360, 90)
    $groupMFA.Size = New-Object System.Drawing.Size(320, 100)
    $form.Controls.Add($groupMFA)

    $lblMFAStatus = New-Object System.Windows.Forms.Label
    $lblMFAStatus.Text = "Not logged in"
    $lblMFAStatus.ForeColor = [System.Drawing.Color]::Red
    $lblMFAStatus.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
    $lblMFAStatus.AutoSize = $true
    $lblMFAStatus.Location = New-Object System.Drawing.Point(10, 25)
    $groupMFA.Controls.Add($lblMFAStatus)

    $btnMFALogin = New-Object System.Windows.Forms.Button
    $btnMFALogin.Text = "Login to MS365 Admin"
    $btnMFALogin.Location = New-Object System.Drawing.Point(170, 55)
    $btnMFALogin.Width = 140
    $btnMFALogin.Height = 30
    $groupMFA.Controls.Add($btnMFALogin)

    # Repositioned controls below the top panels
    $lblAdminUrl = New-Object System.Windows.Forms.Label
    $lblAdminUrl.Text = "Admin Center URL (e.g. https://YourTenant-admin.sharepoint.com)"
    $lblAdminUrl.AutoSize = $true
    $lblAdminUrl.Location = New-Object System.Drawing.Point(20, 200)
    $form.Controls.Add($lblAdminUrl)

    $txtAdminUrl = New-Object System.Windows.Forms.TextBox
    $txtAdminUrl.Location = New-Object System.Drawing.Point(20, 225)
    $txtAdminUrl.Width = 660
    $form.Controls.Add($txtAdminUrl)

    $lblSourceSite = New-Object System.Windows.Forms.Label
    $lblSourceSite.Text = "Source Site URL (e.g. https://YourTenant.sharepoint.com/sites/Source)"
    $lblSourceSite.AutoSize = $true
    $lblSourceSite.Location = New-Object System.Drawing.Point(20, 255)
    $form.Controls.Add($lblSourceSite)

    $txtSourceSite = New-Object System.Windows.Forms.TextBox
    $txtSourceSite.Location = New-Object System.Drawing.Point(20, 280)
    $txtSourceSite.Width = 660
    $form.Controls.Add($txtSourceSite)

    $lblTargetSite = New-Object System.Windows.Forms.Label
    $lblTargetSite.Text = "Target Site URL (e.g. https://YourTenant.sharepoint.com/sites/Target)"
    $lblTargetSite.AutoSize = $true
    $lblTargetSite.Location = New-Object System.Drawing.Point(20, 310)
    $form.Controls.Add($lblTargetSite)

    $txtTargetSite = New-Object System.Windows.Forms.TextBox
    $txtTargetSite.Location = New-Object System.Drawing.Point(20, 335)
    $txtTargetSite.Width = 660
    $form.Controls.Add($txtTargetSite)

    $lblPrefix = New-Object System.Windows.Forms.Label
    $lblPrefix.Text = "Prefix (optional, for bulk updates):"
    $lblPrefix.AutoSize = $true
    $lblPrefix.Location = New-Object System.Drawing.Point(20, 365)
    $form.Controls.Add($lblPrefix)

    $txtPrefix = New-Object System.Windows.Forms.TextBox
    $txtPrefix.Location = New-Object System.Drawing.Point(20, 390)
    $txtPrefix.Width = 200
    $form.Controls.Add($txtPrefix)

    $SiteDesignTitleDefault = "ExportedDesign"
    $SiteScriptTitleDefault = "ExportedScript"

    $lblSiteDesign = New-Object System.Windows.Forms.Label
    $lblSiteDesign.Text = "Title for Site Design (Default: $SiteDesignTitleDefault)"
    $lblSiteDesign.AutoSize = $true
    $lblSiteDesign.Location = New-Object System.Drawing.Point(20, 425)
    $form.Controls.Add($lblSiteDesign)

    $txtSiteDesign = New-Object System.Windows.Forms.TextBox
    $txtSiteDesign.Location = New-Object System.Drawing.Point(20, 450)
    $txtSiteDesign.Width = 300
    $form.Controls.Add($txtSiteDesign)

    $lblSiteScript = New-Object System.Windows.Forms.Label
    $lblSiteScript.Text = "Title for Site Script (Default: $SiteScriptTitleDefault)"
    $lblSiteScript.AutoSize = $true
    $lblSiteScript.Location = New-Object System.Drawing.Point(20, 480)
    $form.Controls.Add($lblSiteScript)

    $txtSiteScript = New-Object System.Windows.Forms.TextBox
    $txtSiteScript.Location = New-Object System.Drawing.Point(20, 505)
    $txtSiteScript.Width = 300
    $form.Controls.Add($txtSiteScript)

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = "OK"
    $btnOk.Width = 80
    $btnOk.Height = 30
    $btnOk.Location = New-Object System.Drawing.Point(490, 545)
    $form.Controls.Add($btnOk)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Width = 80
    $btnCancel.Height = 30
    $btnCancel.Location = New-Object System.Drawing.Point(580, 545)
    $form.Controls.Add($btnCancel)

    $inputs = [PSCustomObject]@{
        AdminUrl         = $null
        SourceSiteUrl    = $null
        TargetSiteUrl    = $null
        Prefix           = $null
        SiteDesignTitle  = $null
        SiteScriptTitle  = $null
        SetRemoteSigned  = $false
        MFALoggedIn      = $false
    }

    # Install module event handler
    $btnInstallModule.Add_Click({
        try {
            $btnInstallModule.Enabled = $false
            $lblModuleStatus.Text = "Installing..."
            $lblModuleStatus.ForeColor = [System.Drawing.Color]::Blue
            $form.Refresh()

            Install-PackageProvider NuGet -Force
            Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted
            Install-Module $moduleName -Force
            
            $lblModuleStatus.Text = "MODULE INSTALLED"
            $lblModuleStatus.ForeColor = [System.Drawing.Color]::Green
            $btnInstallModule.Enabled = $false
            [System.Windows.Forms.MessageBox]::Show("'$moduleName' was installed.", "Installation successful", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
        catch {
            $lblModuleStatus.Text = "INSTALLATION FAILED!!!"
            $lblModuleStatus.ForeColor = [System.Drawing.Color]::Red
            $btnInstallModule.Enabled = $true
            [System.Windows.Forms.MessageBox]::Show("ERROR: Could not install module '$moduleName'`n`n$($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })

    # MFA login event handler
    $btnMFALogin.Add_Click({
        try {
            $adminUrl = $txtAdminUrl.Text.Trim()
            
            if ([string]::IsNullOrWhiteSpace($adminUrl)) {
                [System.Windows.Forms.MessageBox]::Show("Please enter Admin Center URL first.", "Input Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }
            
            $lblMFAStatus.Text = "Connecting..."
            $lblMFAStatus.ForeColor = [System.Drawing.Color]::Blue
            $form.Refresh()
            
            Connect-PnPOnline -Url $adminUrl -Interactive -ErrorAction Stop
            
            $lblMFAStatus.Text = "LOGGED IN"
            $lblMFAStatus.ForeColor = [System.Drawing.Color]::Green
            $inputs.MFALoggedIn = $true
            
            [System.Windows.Forms.MessageBox]::Show("Successfully logged in to Admin Center.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
        catch {
            $lblMFAStatus.Text = "LOGIN FAILED!"
            $lblMFAStatus.ForeColor = [System.Drawing.Color]::Red
            $inputs.MFALoggedIn = $false
            [System.Windows.Forms.MessageBox]::Show("ERROR: Could not connect to Admin Center`n`n$($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })

    $btnOk.Add_Click({
        $inputs.AdminUrl = $txtAdminUrl.Text
        $inputs.SourceSiteUrl = $txtSourceSite.Text
        $inputs.TargetSiteUrl = $txtTargetSite.Text
        $inputs.Prefix = $txtPrefix.Text
        $inputs.SetRemoteSigned = $chkRemoteSigned.Checked
        if ([string]::IsNullOrWhiteSpace($txtSiteDesign.Text)) {
            $inputs.SiteDesignTitle = $SiteDesignTitleDefault
        } else {
            $inputs.SiteDesignTitle = $txtSiteDesign.Text
        }
        if ([string]::IsNullOrWhiteSpace($txtSiteScript.Text)) {
            $inputs.SiteScriptTitle = $SiteScriptTitleDefault
        } else {
            $inputs.SiteScriptTitle = $txtSiteScript.Text
        }
        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Close()
    })

    $btnCancel.Add_Click({
        $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $form.Close()
    })

    $null = $form.ShowDialog()
    if ($form.DialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
        return $inputs
    }
    else {
        return $null
    }
}

$userInputs = Show-ConfigForm
if ($null -eq $userInputs) {
    return
}

# Set execution policy if requested
if ($userInputs.SetRemoteSigned -and (Get-ExecutionPolicy -Scope Process) -ne "RemoteSigned") {
    try {
        Set-ExecutionPolicy RemoteSigned -Scope Process -Force
        [System.Windows.Forms.MessageBox]::Show("Policy was set to 'RemoteSigned'.", "Info", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("ERROR - ExecutionPolicy. Please check.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

# Check if module is imported
try {
    Import-Module $moduleName -ErrorAction Stop
}
catch {
    [System.Windows.Forms.MessageBox]::Show("ERROR: Could not import module '$moduleName'. Script will terminate.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    return
}

# Validierung der einzelnen Eigenschaften
if ([string]::IsNullOrWhiteSpace($userInputs.AdminUrl) -or 
    [string]::IsNullOrWhiteSpace($userInputs.SourceSiteUrl)) {
    [System.Windows.Forms.MessageBox]::Show("Admin-URL and Source Site URL must not be empty.", "Validation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    return
}

# Prüfen, ob entweder Ziel-Site oder Präfix angegeben wurde
if ([string]::IsNullOrWhiteSpace($userInputs.TargetSiteUrl) -and [string]::IsNullOrWhiteSpace($userInputs.Prefix)) {
    [System.Windows.Forms.MessageBox]::Show("Please enter a Target Site URL or a prefix for bulk updates.", "Validation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    return
}

try {
    [System.Windows.Forms.MessageBox]::Show("Starting MFA-based interactive sign in to the same tenant's Admin Center.", "MFA SignIn", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    Connect-PnPOnline -Url $($userInputs.AdminUrl) -Interactive -ErrorAction Stop
}
catch {
    [System.Windows.Forms.MessageBox]::Show("Error signing in to Admin Center: $($_.Exception.Message)", "SignIn Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    return
}

try {
    [System.Windows.Forms.MessageBox]::Show("Please enter your credentials for the source site.", "Credential Prompt", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    $Cred = Get-Credential -Message "Credentials for source site: $($userInputs.SourceSiteUrl)" -ErrorAction Stop
    if ($null -eq $Cred) {
        [System.Windows.Forms.MessageBox]::Show("No credentials provided. Script will terminate.", "Cancelled", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    Connect-PnPOnline -Url $($userInputs.SourceSiteUrl) -Credentials $Cred -ErrorAction Stop
}
catch {
    [System.Windows.Forms.MessageBox]::Show("Error signing in to the source site: $($_.Exception.Message)", "SignIn Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    return
}

# Site Script Export mit Null-Prüfung
try {
    $exportedSiteScript = Get-PnPSiteScriptFromWeb -IncludeBranding -IncludeTheme -ErrorAction Stop
    if ($null -eq $exportedSiteScript) {
        [System.Windows.Forms.MessageBox]::Show("Could not export any site script.", "Export Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    
    $jsonScript = $exportedSiteScript.JSON
    if ([string]::IsNullOrWhiteSpace($jsonScript)) {
        [System.Windows.Forms.MessageBox]::Show("The exported site script does not contain valid JSON.", "Empty Content", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
}
catch {
    [System.Windows.Forms.MessageBox]::Show("Error exporting the site script: $($_.Exception.Message)", "Export Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    return
}

# Verbindung zum Admin Center mit Fehlerbehandlung
try {
    Connect-PnPOnline -Url $($userInputs.AdminUrl) -Credentials $Cred -ErrorAction Stop
}
catch {
    [System.Windows.Forms.MessageBox]::Show("Error re-signing in to Admin Center: $($_.Exception.Message)", "SignIn Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    return
}

# Site Script und Design hinzufügen mit Null-Prüfung
try {
    $newSiteScript = Add-PnPSiteScript -Title $($userInputs.SiteScriptTitle) -Description "Exported from $($userInputs.SourceSiteUrl)" -Content $jsonScript -ErrorAction Stop
    if ($null -eq $newSiteScript -or $null -eq $newSiteScript.Id) {
        [System.Windows.Forms.MessageBox]::Show("Could not create site script or retrieve ID.", "Script Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    
    $newSiteDesign = Add-PnPSiteDesign -Title $($userInputs.SiteDesignTitle) -WebTemplate "64" -SiteScripts $newSiteScript.Id -Description "Design from $($userInputs.SourceSiteUrl)" -ErrorAction Stop
    if ($null -eq $newSiteDesign -or $null -eq $newSiteDesign.Id) {
        [System.Windows.Forms.MessageBox]::Show("Could not create site design or retrieve ID.", "Design Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
}
catch {
    [System.Windows.Forms.MessageBox]::Show("Error creating site script/design: $($_.Exception.Message)", "Creation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    return
}

# Design-Anwendung auf Ziel-Site oder via Präfix
if (-not [string]::IsNullOrWhiteSpace($userInputs.TargetSiteUrl)) {
    # Option 1: Direkt auf eine Ziel-Site anwenden
    try {
        Invoke-PnPSiteDesign -Identity $newSiteDesign.Id -WebUrl $userInputs.TargetSiteUrl -ErrorAction Stop
        [System.Windows.Forms.MessageBox]::Show("Design successfully applied to '$($userInputs.TargetSiteUrl)'.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error applying the design to $($userInputs.TargetSiteUrl): $($_.Exception.Message)", "Apply Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}
elseif (-not [string]::IsNullOrWhiteSpace($userInputs.Prefix)) {
    # Option 2: Auf alle Sites mit Präfix anwenden (bestehende Funktionalität)
    try {
        $targetSites = Get-PnPTenantSite -ErrorAction Stop | Where-Object { $_.Url -like "*$($userInputs.Prefix)*" }
        if ($null -eq $targetSites) {
            [System.Windows.Forms.MessageBox]::Show("No sites found or error retrieving sites.", "Info", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            return
        }
        
        if ($targetSites.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No sites found containing '$($userInputs.Prefix)'.", "Info", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            return
        }
        
        foreach ($site in $targetSites) {
            if ($null -ne $site -and -not [string]::IsNullOrWhiteSpace($site.Url)) {
                try {
                    Invoke-PnPSiteDesign -Identity $newSiteDesign.Id -WebUrl $site.Url -ErrorAction Stop
                }
                catch {
                    [System.Windows.Forms.MessageBox]::Show("Error applying the design to $($site.Url): $($_.Exception.Message)", "Apply Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                }
            }
        }
        
        [System.Windows.Forms.MessageBox]::Show("Done! All sites containing the prefix '$($userInputs.Prefix)' were updated.", "Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error retrieving or updating target sites: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
}