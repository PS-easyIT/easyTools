[CmdletBinding()]
param(
    [ValidateSet('OU','Department','Company','Group','GroupMembership','UserGroups')]
    [string]$FilterType,

    [ValidateSet('HTML','CSV','Excel','JSON','Both')]
    [string]$OutputType,

    [string]$OutputPath,

    [switch]$IncludeDisabled,
    [switch]$Silent,
    [switch]$GroupsOnly,
    [switch]$UsersOnly
)

# Dynamically determine script directory for default paths
$ScriptDirectory = $null
try {
    # Try to get script path via MyInvocation (more reliable in some PS 5.1 contexts)
    $InvocationInfo = (Get-Variable MyInvocation -Scope 1 -ErrorAction Stop).Value
    if ($InvocationInfo.MyCommand.Path) {
        $ScriptDirectory = Split-Path -Path $InvocationInfo.MyCommand.Path -Parent
    }
}
catch {
    # MyInvocation might not be available or suitable (e.g., in runspace, interactive debugging)
    # Fall back to PSScriptRoot if MyInvocation failed
    if (-not ([string]::IsNullOrEmpty($PSScriptRoot))) {
        $ScriptDirectory = $PSScriptRoot
    }
}

# Resolve OutputPath default if not provided by user
if (-not $PSBoundParameters.ContainsKey('OutputPath')) {
    if (-not ([string]::IsNullOrEmpty($ScriptDirectory))) {
        $OutputPath = Join-Path $ScriptDirectory 'ADUsersExport'
    } else {
        $OutputPath = Join-Path -Path (Get-Location).Path -ChildPath 'ADUsersExport' # Explicitly CWD
        Write-Warning "WARNUNG: Skript-Stammverzeichnis konnte nicht zuverlässig ermittelt werden. OutputPath ('$OutputPath') wird im aktuellen Arbeitsverzeichnis erstellt: '$((Get-Location).Path)'"
    }
}


#region ─── Helfer ───────────────────────────────────────────────────────────────
function Ensure-ADModule {
    if (-not (Get-Module -Name ActiveDirectory -ListAvailable)) { 
        throw 'ActiveDirectory Modul nicht installiert.' 
    }
    Import-Module -Name ActiveDirectory -ErrorAction Stop
    Write-Debug 'ActiveDirectory Modul geladen.'
}

function Choose-FilterType { 
    param($Current) 
    if ($Silent -or $Current){return $Current}
    Write-Host "Filtertyp wählen:`n1. OU`n2. Department`n3. Company" -ForegroundColor Cyan
    switch (Read-Host 'Ihre Wahl'){ 1{'OU'} 2{'Department'} 3{'Company'} default{Choose-FilterType} }
}

function Choose-OutputType { 
    param(
        [Parameter()]
        [string]$Current
    )
    
    if ($Silent -or $Current -eq 'CSV') { 
        return 'CSV' # Immer CSV zurückgeben, wenn $Silent oder bereits CSV
    }
    
    # Interaktiver Modus ist nicht mehr sinnvoll, da nur CSV möglich ist.
    # Write-Host "Exportformat wählen:`n1. CSV" -ForegroundColor Cyan
    # switch (Read-Host 'Ihre Wahl') { 
    #     1 { 'CSV' }
    #     default { Choose-OutputType }
    # }
    return 'CSV' # Standardmäßig CSV zurückgeben
}

function Select-Values {
    param($Type)
    switch ($Type){
        'OU'        { Get-ADOrganizationalUnit -Filter * | Sort DistinguishedName | Select -Expand DistinguishedName | Out-GridView -Title 'OU auswählen' -PassThru }
        'Department'{ Get-ADUser -Filter * -Properties Department | Where Department | Select -Expand Department -Unique | Sort | Out-GridView -Title 'Department auswählen' -PassThru }
        'Company'   { Get-ADUser -Filter * -Properties Company    | Where Company    | Select -Expand Company    -Unique | Sort | Out-GridView -Title 'Company auswählen' -PassThru }
    }
}

function Build-LdapFilter {
    param(
        $Type,
        $Values,
        [System.Nullable[bool]]$FilterByEnabledState = $false,
        [System.Nullable[bool]]$UsersShouldBeEnabled = $true,
        [string]$ObjectClass = 'user'
    )
    
    $orClauses = switch($Type){
        'OU'        { $Values | ForEach-Object {"(distinguishedName=*$_*)"} }
        'GroupType' { 
            $Values | ForEach-Object {
                switch ($_) {
                    'Security' { "(groupType:1.2.840.113556.1.4.803:=2147483648)" }
                    'Distribution' { "(!(groupType:1.2.840.113556.1.4.803:=2147483648))" }
                    'Global' { "(groupType:1.2.840.113556.1.4.803:=2)" }
                    'Universal' { "(groupType:1.2.840.113556.1.4.803:=8)" }
                    'DomainLocal' { "(groupType:1.2.840.113556.1.4.803:=4)" }
                    default { "($Type=$($_))" }
                }
            }
        }
        Default     { $Values | ForEach-Object {"($Type=$($_))"} }
    }
    
    if ($orClauses -and $orClauses.Count -gt 0) {
        $mainFilterClause = "(" + ($orClauses -join '|') + ")"
    } else {
        $mainFilterClause = ""
    }
    
    $finalFilterClauses = @("(objectClass=$ObjectClass)")
    if ($mainFilterClause) {
        $finalFilterClauses += $mainFilterClause
    }

    if ($FilterByEnabledState -eq $true) {
        if ($UsersShouldBeEnabled -eq $true) {
            $finalFilterClauses += "(!(userAccountControl:1.2.840.113556.1.4.803:=2))"
        } else {
            $finalFilterClauses += "(userAccountControl:1.2.840.113556.1.4.803:=2)"
        }
    }
    
    return "(&" + ($finalFilterClauses -join '') + ")"
}

function Get-SelectableADAttributes {
    $defaultAttributes = @(
        'SamAccountName',
        'Name',
        'Enabled',
        'mail',
        'Department',
        'Company',
        'DistinguishedName',
        'DisplayName',
        'GivenName',
        'Surname',
        'Title',
        'Office',
        'TelephoneNumber',
        'MobilePhone',
        'Manager',
        'LastLogonDate',
        'Created',
        'Modified',
        'Description',
        'EmployeeID',
        'EmployeeNumber',
        'EmployeeType',
        'Division',
        'OfficePhone',
        'HomePhone',
        'Pager',
        'Fax',
        'StreetAddress',
        'City',
        'State',
        'PostalCode',
        'Country',
        'Notes'
    )
    return $defaultAttributes
}

function Select-ExportAttributes {
    param($Current)
    if ($Silent -or $Current) { return $Current }
    
    $availableAttributes = Get-SelectableADAttributes
    $selectedAttributes = $availableAttributes | 
        ForEach-Object {
            [PSCustomObject]@{
                Attribute = $_
                Selected = $true # Default to selected
            }
        } | 
        Out-GridView -Title 'AD-Attribute für Export auswählen' -PassThru
    
    if ($selectedAttributes) {
        return $selectedAttributes | Where-Object Selected | Select-Object -ExpandProperty Attribute
    }
    return $null
}

function Export-Data {
    param(
        $Collection,
        [ValidateSet('CSV')] # Nur CSV erlauben
        [string]$Type = 'CSV', # Standard auf CSV setzen
        [string]$Path,
        [string[]]$Attributes,
        [string]$SortBy,
        [string]$ExcludePattern,
        [string]$RequiredAttribute,
        [string]$ReportTitle = "AD Audit Report"
    )
    try {
        Write-Debug "Starte Export ($Type) nach $Path"
        
        $Path = [System.IO.Path]::GetFullPath($Path)
        
        if (-not (Test-Path $Path)) {
            Write-Debug "Erstelle Export-Verzeichnis: $Path"
            New-Item $Path -ItemType Directory -Force | Out-Null
        }
        
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        # Dateiendung immer .csv
        $file = Join-Path $Path "$($ReportTitle.Replace(' ','_'))_$timestamp.csv"
        
        $exportData = $Collection | Select-Object $Attributes
        
        if ($RequiredAttribute) {
            Write-Debug "Filtere nach Pflichtattribut: $RequiredAttribute"
            $exportData = $exportData | Where-Object { $_.$RequiredAttribute }
        }
        
        if ($SortBy) {
            Write-Debug "Sortiere nach: $SortBy"
            $exportData = $exportData | Sort-Object $SortBy
        }
        
        if ($ExcludePattern) {
            Write-Debug "Filtere Benutzer mit Pattern: $ExcludePattern"
            $exportData = $exportData | Where-Object { $_.Name -notlike "*$ExcludePattern*" }
        }
        
        # Switch-Block ist jetzt nur noch für CSV relevant
        switch ($Type) {
            'CSV' { 
                Write-Debug "Exportiere CSV mit $($exportData.Count) Datensätzen"
                $exportData | Export-Csv -Path $file -NoType -Encoding UTF8 -ErrorAction Stop
            }
            # HTML, JSON, Excel cases entfernt
        }
        
        return $file
    } # Ende des try-Blocks
    catch {
        Write-Error "Fehler beim Export ($Type): $_"
        throw
    }
} # Ende der Funktion Export-Data
#endregion Helfer

#region ─── Moderne Windows 11 XAML GUI ─────────────────────────────────────────
$xamlContent = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="EasyADAudit - Active Directory Audit Tool" Height="840" Width="1900"
    WindowStartupLocation="CenterScreen" ResizeMode="CanResize"
    Background="#F3F3F3" FontFamily="Segoe UI" FontSize="12">
    
    <Window.Resources>
    <!-- Windows 11 Design System Colors -->
    <SolidColorBrush x:Key="AccentBrush" Color="#0078D4"/>
    <SolidColorBrush x:Key="AccentHoverBrush" Color="#106EBE"/>
    <SolidColorBrush x:Key="AccentPressedBrush" Color="#005A9E"/>
    <SolidColorBrush x:Key="CardBackgroundBrush" Color="#FFFFFF"/>
    <SolidColorBrush x:Key="SurfaceBrush" Color="#F9F9F9"/>
    <SolidColorBrush x:Key="BorderBrush" Color="#E5E5E5"/>
    <SolidColorBrush x:Key="TextPrimaryBrush" Color="#323130"/>
    <SolidColorBrush x:Key="TextSecondaryBrush" Color="#605E5C"/>
    
    <!-- Modern Button Style -->
    <Style x:Key="ModernButtonStyle" TargetType="Button">
        <Setter Property="Background" Value="{StaticResource AccentBrush}"/>
        <Setter Property="Foreground" Value="White"/>
        <Setter Property="BorderThickness" Value="0"/>
        <Setter Property="Padding" Value="16,8"/>
        <Setter Property="FontWeight" Value="SemiBold"/>
        <Setter Property="Cursor" Value="Hand"/>
        <Setter Property="Template">
        <Setter.Value>
            <ControlTemplate TargetType="Button">
            <Border Background="{TemplateBinding Background}" 
                CornerRadius="4" 
                BorderThickness="{TemplateBinding BorderThickness}"
                BorderBrush="{TemplateBinding BorderBrush}">
                <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                <Setter Property="Background" Value="{StaticResource AccentHoverBrush}"/>
                </Trigger>
                <Trigger Property="IsPressed" Value="True">
                <Setter Property="Background" Value="{StaticResource AccentPressedBrush}"/>
                </Trigger>
            </ControlTemplate.Triggers>
            </ControlTemplate>
        </Setter.Value>
        </Setter>
    </Style>
    
    <!-- Secondary Button Style -->
    <Style x:Key="SecondaryButtonStyle" TargetType="Button" BasedOn="{StaticResource ModernButtonStyle}">
        <Setter Property="Background" Value="{StaticResource SurfaceBrush}"/>
        <Setter Property="Foreground" Value="{StaticResource TextPrimaryBrush}"/>
        <Setter Property="BorderThickness" Value="1"/>
        <Setter Property="BorderBrush" Value="{StaticResource BorderBrush}"/>
    </Style>
    
    <!-- Card Style -->
    <Style x:Key="CardStyle" TargetType="Border">
        <Setter Property="Background" Value="{StaticResource CardBackgroundBrush}"/>
        <Setter Property="BorderBrush" Value="{StaticResource BorderBrush}"/>
        <Setter Property="BorderThickness" Value="1"/>
        <Setter Property="CornerRadius" Value="8"/>
        <Setter Property="Padding" Value="16"/>
        <Setter Property="Margin" Value="8"/>
        <Setter Property="Effect">
        <Setter.Value>
            <DropShadowEffect Color="#000000" Opacity="0.1" ShadowDepth="2" BlurRadius="8"/>
        </Setter.Value>
        </Setter>
    </Style>
    
    <!-- Modern ComboBox Style -->
    <Style x:Key="ModernComboBoxStyle" TargetType="ComboBox">
        <Setter Property="Height" Value="32"/>
        <Setter Property="Padding" Value="12,6"/>
        <Setter Property="BorderBrush" Value="{StaticResource BorderBrush}"/>
        <Setter Property="Background" Value="White"/>
    </Style>
    
    <!-- Modern TextBox Style -->
    <Style x:Key="ModernTextBoxStyle" TargetType="TextBox">
        <Setter Property="Height" Value="32"/>
        <Setter Property="Padding" Value="12,6"/>
        <Setter Property="BorderBrush" Value="{StaticResource BorderBrush}"/>
        <Setter Property="Background" Value="White"/>
        <Setter Property="BorderThickness" Value="1"/>
    </Style>
    </Window.Resources>

    <Grid>
    <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="*"/>
        <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    
    <!-- Header -->
    <Border Grid.Row="0" Background="{StaticResource AccentBrush}" Padding="20,12">
        <StackPanel>
        <TextBlock Text="EasyADAudit" FontSize="24" FontWeight="Bold" Foreground="White"/>
        <TextBlock Text="Active Directory Audit and Reporting Tool" FontSize="14" Foreground="White" Opacity="0.9"/>
        </StackPanel>
    </Border>
    
    <!-- Main Content -->
    <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" Padding="16">
        <StackPanel>
        
        <!-- Audit Type Selection Card -->
        <Border Style="{StaticResource CardStyle}">
            <StackPanel>
            <TextBlock Text="Select Audit Type" FontSize="16" FontWeight="SemiBold" Margin="0,0,0,12" Foreground="{StaticResource TextPrimaryBrush}"/>
            <Grid>
                <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                
                <RadioButton x:Name="RadioUsers" Content="User Audit" IsChecked="True" Margin="8" FontWeight="SemiBold"/>
                <RadioButton x:Name="RadioGroups" Content="Group Audit" Grid.Column="1" Margin="8" FontWeight="SemiBold"/>
                <RadioButton x:Name="RadioMemberships" Content="Membership Audit" Grid.Column="2" Margin="8" FontWeight="SemiBold"/>
            </Grid>
            </StackPanel>
        </Border>
        
        <!-- Configuration Sections in a Row -->
        <Grid>
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

            <!-- Filter Configuration Card -->
            <Border Grid.Column="0" Style="{StaticResource CardStyle}" VerticalAlignment="Stretch">
                <StackPanel>
                <TextBlock Text="Filter Configuration" FontSize="16" FontWeight="SemiBold" Margin="0,0,0,12" Foreground="{StaticResource TextPrimaryBrush}"/>
                
                <Grid>
                    <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <TextBlock Text="Filter Type:" VerticalAlignment="Center" Margin="0,0,12,0" FontWeight="SemiBold"/>
                    <ComboBox x:Name="ComboBoxFilterType" Grid.Column="1" Style="{StaticResource ModernComboBoxStyle}" Margin="0,4"/>
                    <Button x:Name="ButtonPopulateFilterValues" Content="Load Values" Grid.Column="2" Style="{StaticResource SecondaryButtonStyle}" Margin="8,4,0,4"/>
                    
                    <TextBlock Text="Available Values:" Grid.Row="1" VerticalAlignment="Top" Margin="0,12,12,0" FontWeight="SemiBold"/>
                    <Border Grid.Row="1" Grid.Column="1" Grid.ColumnSpan="2" BorderBrush="{StaticResource BorderBrush}" BorderThickness="1" CornerRadius="4" Margin="0,8,0,0" Height="120">
                    <ListBox x:Name="ListBoxFilterValues" Background="White" BorderThickness="0">
                        <ListBox.ItemTemplate>
                        <DataTemplate>
                            <CheckBox Content="{Binding DisplayName}" IsChecked="{Binding IsSelected}" Margin="2"/>
                        </DataTemplate>
                        </ListBox.ItemTemplate>
                    </ListBox>
                    </Border>
                    
                    <TextBlock x:Name="TextBlockSelectedFilterValues" Text="Selected Values: (None)" Grid.Row="2" Grid.ColumnSpan="3" Margin="0,8,0,0" FontStyle="Italic" Foreground="{StaticResource TextSecondaryBrush}"/>
                </Grid>
                
                <WrapPanel Orientation="Horizontal" Margin="0,12,0,0">
                    <CheckBox x:Name="CheckBoxIncludeDisabled" Content="Include Disabled Objects" Margin="0,0,16,4"/>
                    <CheckBox x:Name="CheckBoxIncludeNested" Content="Include Nested Groups" Margin="0,0,16,4"/>
                    <CheckBox x:Name="CheckBoxShowEmptyGroups" Content="Show Empty Groups" Margin="0,0,0,4"/>
                </WrapPanel>
                </StackPanel>
            </Border>
            
            <!-- Export Configuration Card -->
            <Border Grid.Column="1" Style="{StaticResource CardStyle}" VerticalAlignment="Stretch">
                <StackPanel>
                <TextBlock Text="Export Configuration" FontSize="16" FontWeight="SemiBold" Margin="0,0,0,12" Foreground="{StaticResource TextPrimaryBrush}"/>
                
                <Grid>
                    <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <StackPanel Grid.ColumnSpan="2" Margin="0,0,0,12">
                    <TextBlock Text="Export Format:" FontWeight="SemiBold" Margin="0,0,0,4"/>
                    <ComboBox x:Name="ComboBoxOutputType" Style="{StaticResource ModernComboBoxStyle}"/>
                    </StackPanel>
                    
                    <StackPanel Grid.Row="1" Grid.ColumnSpan="2" Margin="0,0,0,12">
                    <TextBlock Text="Export Path:" FontWeight="SemiBold" Margin="0,0,0,4"/>
                    <Grid>
                        <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <TextBox x:Name="TextBoxOutputPath" Style="{StaticResource ModernTextBoxStyle}"/>
                        <Button x:Name="ButtonBrowseOutputPath" Content="Browse" Grid.Column="1" Style="{StaticResource SecondaryButtonStyle}" Margin="8,0,0,0" Width="Auto"/>
                    </Grid>
                    </StackPanel>
                    
                    <StackPanel Grid.Row="2" Margin="0,0,8,0">
                    <TextBlock Text="Select Attributes:" FontWeight="SemiBold" Margin="0,0,0,4"/>
                    <Button x:Name="ButtonSelectExportAttributes" Content="Select Attributes..." Style="{StaticResource SecondaryButtonStyle}" HorizontalAlignment="Stretch"/>
                    <TextBlock x:Name="TextBlockSelectedAttributesCount" Text="Selected Attributes: (Please select)" Margin="0,4,0,0" FontSize="11" Foreground="{StaticResource TextSecondaryBrush}"/>
                    </StackPanel>
                    
                    <StackPanel Grid.Row="2" Grid.Column="1" Margin="8,0,0,0">
                    <TextBlock Text="Advanced Options:" FontWeight="SemiBold" Margin="0,0,0,4"/>
                    <TextBox x:Name="TextBoxRequiredAttribute" Style="{StaticResource ModernTextBoxStyle}" Text="" Margin="0,2"/>
                    <TextBlock Text="Required Attribute" FontSize="10" Foreground="{StaticResource TextSecondaryBrush}" Margin="2,1,0,4"/>
                    <TextBox x:Name="TextBoxSortBy" Style="{StaticResource ModernTextBoxStyle}" Text="Name" Margin="0,2"/>
                    <TextBlock Text="Sort By" FontSize="10" Foreground="{StaticResource TextSecondaryBrush}" Margin="2,1,0,4"/>
                    <TextBox x:Name="TextBoxExcludePattern" Style="{StaticResource ModernTextBoxStyle}" Text="" Margin="0,2"/>
                    <TextBlock Text="Exclude Pattern" FontSize="10" Foreground="{StaticResource TextSecondaryBrush}" Margin="2,1,0,0"/>
                    </StackPanel>
                </Grid>
                </StackPanel>
            </Border>
            
            <!-- Analysis & Preview Card -->
            <Border Grid.Column="2" Style="{StaticResource CardStyle}" VerticalAlignment="Stretch">
                <StackPanel>
                <TextBlock Text="Analysis &amp; Preview" FontSize="16" FontWeight="SemiBold" Margin="0,0,0,12" Foreground="{StaticResource TextPrimaryBrush}"/>
                
                <Grid>
                    <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    
                    <Border Background="#E3F2FD" CornerRadius="4" Padding="12" Margin="0,0,4,0">
                    <StackPanel>
                        <TextBlock Text="Object Count" FontWeight="SemiBold" Foreground="#1976D2"/>
                        <TextBlock x:Name="TextBlockUserCountPreview" Text="Users: -" FontSize="14" Margin="0,4,0,0"/>
                        <TextBlock x:Name="TextBlockGroupCountPreview" Text="Groups: -" FontSize="14"/>
                    </StackPanel>
                    </Border>
                    
                    <Border Background="#F3E5F5" CornerRadius="4" Padding="12" Margin="4,0" Grid.Column="1">
                    <StackPanel>
                        <TextBlock Text="Status" FontWeight="SemiBold" Foreground="#7B1FA2"/>
                        <TextBlock x:Name="TextBlockActiveCount" Text="Active: -" FontSize="14" Margin="0,4,0,0"/>
                        <TextBlock x:Name="TextBlockInactiveCount" Text="Inactive: -" FontSize="14"/>
                    </StackPanel>
                    </Border>
                    
                    <Border Background="#E8F5E8" CornerRadius="4" Padding="12" Margin="4,0,0,0" Grid.Column="2">
                    <StackPanel>
                        <TextBlock Text="Memberships" FontWeight="SemiBold" Foreground="#388E3C"/>
                        <TextBlock x:Name="TextBlockMembershipCount" Text="Assignments: -" FontSize="14" Margin="0,4,0,0"/>
                        <TextBlock x:Name="TextBlockOrphanedCount" Text="Orphaned: -" FontSize="14"/>
                    </StackPanel>
                    </Border>
                </Grid>
                
                <Button x:Name="ButtonUpdatePreview" Content="Update Preview" Style="{StaticResource SecondaryButtonStyle}" Margin="0,12,0,0" HorizontalAlignment="Center"/>
                </StackPanel>
            </Border>
        </Grid>
        
        <!-- Quick Reports Card -->
        <Border Style="{StaticResource CardStyle}">
            <StackPanel>
            <TextBlock Text="Quick Reports" FontSize="16" FontWeight="SemiBold" Margin="0,0,0,12" Foreground="{StaticResource TextPrimaryBrush}"/>
            
            <UniformGrid Columns="3" Margin="0,0,0,12">
                <Button x:Name="ButtonQuickReportInactiveUsers" Content="Inactive Users" Style="{StaticResource SecondaryButtonStyle}" Margin="0,0,4,4"/>
                <Button x:Name="ButtonQuickReportEmptyGroups" Content="Empty Groups" Style="{StaticResource SecondaryButtonStyle}" Margin="4,0,4,4"/>
                <Button x:Name="ButtonQuickReportAdmins" Content="Admin Groups" Style="{StaticResource SecondaryButtonStyle}" Margin="4,0,0,4"/>
                <Button x:Name="ButtonQuickReportNoLogin" Content="Never Logged In" Style="{StaticResource SecondaryButtonStyle}" Margin="0,4,4,0"/>
                <Button x:Name="ButtonQuickReportPasswordExpiry" Content="Password Expiry" Style="{StaticResource SecondaryButtonStyle}" Margin="4,4,4,0"/>
                <Button x:Name="ButtonQuickReportLargeGroups" Content="Large Groups" Style="{StaticResource SecondaryButtonStyle}" Margin="4,4,0,0"/>
            </UniformGrid>
            </StackPanel>
        </Border>
        </StackPanel>
    </ScrollViewer>
    
    <!-- Footer with Actions -->
    <Border Grid.Row="2" Background="{StaticResource SurfaceBrush}" BorderBrush="{StaticResource BorderBrush}" BorderThickness="0,1,0,0" Padding="16,12">
        <Grid>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        
        <StackPanel>
            <TextBlock x:Name="TextBlockStatus" Text="Ready. Please select audit type and filters." FontWeight="SemiBold" Foreground="{StaticResource TextPrimaryBrush}"/>
            <ProgressBar x:Name="ProgressBarStatus" Height="4" Margin="0,4,0,0" Background="Transparent" Foreground="{StaticResource AccentBrush}"/>
        </StackPanel>
        
        <Button x:Name="ButtonStartExport" Content="Start Export" Grid.Column="1" Style="{StaticResource ModernButtonStyle}" Margin="0,0,8,0" Padding="24,12"/>
        <Button x:Name="ButtonClose" Content="Close" Grid.Column="2" Style="{StaticResource SecondaryButtonStyle}" Padding="16,12"/>
        </Grid>
    </Border>
    </Grid>
</Window>
'@

# GUI-spezifische Funktionen für erweiterte AD-Audit-Funktionalität
function Get-GroupAuditData {
    param(
        [string[]]$FilterValues,
        [string]$FilterType,
        [bool]$IncludeDisabled = $false,
        [bool]$IncludeNested = $false,
        [bool]$ShowEmptyGroups = $false
    )
    
    try {
        Write-Debug "Starte Gruppen-Audit..."
        $allGroups = @()
        
        switch ($FilterType) {
            'OU' {
                foreach ($ou in $FilterValues) {
                    $groups = Get-ADGroup -SearchBase $ou -Filter * -Properties *
                    $allGroups += $groups
                }
            }
            'GroupType' {
                foreach ($type in $FilterValues) {
                    $filter = switch ($type) {
                        'Security' { "GroupCategory -eq 'Security'" }
                        'Distribution' { "GroupCategory -eq 'Distribution'" }
                        'Global' { "GroupScope -eq 'Global'" }
                        'Universal' { "GroupScope -eq 'Universal'" }
                        'DomainLocal' { "GroupScope -eq 'DomainLocal'" }
                        default { "*" }
                    }
                    $groups = Get-ADGroup -Filter $filter -Properties *
                    $allGroups += $groups
                }
            }
            default {
                $allGroups = Get-ADGroup -Filter * -Properties *
            }
        }
        
        # Erweiterte Gruppeninformationen hinzufügen
        $enrichedGroups = foreach ($group in $allGroups) {
            $members = Get-ADGroupMember -Identity $group.DistinguishedName -ErrorAction SilentlyContinue
            $memberCount = ($members | Measure-Object).Count
            
            if ($ShowEmptyGroups -or $memberCount -gt 0) {
                [PSCustomObject]@{
                    Name = $group.Name
                    DisplayName = $group.DisplayName
                    Description = $group.Description
                    GroupCategory = $group.GroupCategory
                    GroupScope = $group.GroupScope
                    MemberCount = $memberCount
                    Members = ($members | Select-Object -ExpandProperty Name) -join '; '
                    Created = $group.Created
                    Modified = $group.Modified
                    DistinguishedName = $group.DistinguishedName
                    SamAccountName = $group.SamAccountName
                    IsEmpty = $memberCount -eq 0
                    ObjectGUID = $group.ObjectGUID
                }
            }
        }
        
        return $enrichedGroups | Sort-Object Name
    }
    catch {
        Write-Error "Fehler beim Gruppen-Audit: $_"
        throw
    }
}

function Get-MembershipAuditData {
    param(
        [string[]]$FilterValues,
        [string]$FilterType,
        [bool]$IncludeNested = $false
    )
    
    try {
        Write-Debug "Starte Mitgliedschafts-Audit..."
        $membershipData = @()
        
        $users = Get-ADUser -Filter * -Properties MemberOf
        
        foreach ($user in $users) {
            $userGroups = @()
            
            foreach ($groupDN in $user.MemberOf) {
                try {
                    $group = Get-ADGroup -Identity $groupDN -Properties Name
                    $userGroups += $group.Name
                    
                    if ($IncludeNested) {
                        $nestedGroups = Get-ADGroup -Identity $groupDN -Properties MemberOf
                        foreach ($nestedGroupDN in $nestedGroups.MemberOf) {
                            $nestedGroup = Get-ADGroup -Identity $nestedGroupDN -Properties Name -ErrorAction SilentlyContinue
                            if ($nestedGroup) {
                                $userGroups += "$($group.Name) -> $($nestedGroup.Name)"
                            }
                        }
                    }
                }
                catch {
                    Write-Warning "Gruppe nicht gefunden: $groupDN"
                }
            }
            
            $membershipData += [PSCustomObject]@{
                UserName = $user.Name
                SamAccountName = $user.SamAccountName
                Enabled = $user.Enabled
                GroupCount = $userGroups.Count
                Groups = $userGroups -join '; '
                HasAdminGroups = ($userGroups | Where-Object { $_ -match 'admin|domain|enterprise' }).Count -gt 0
                DistinguishedName = $user.DistinguishedName
            }
        }
        
        return $membershipData | Sort-Object UserName
    }
    catch {
        Write-Error "Fehler beim Mitgliedschafts-Audit: $_"
        throw
    }
}

# Globale Variablen für das Hauptfenster und XAML-Elemente
$Global:MainWindow = $null
$Global:XamlControls = @{}
$Global:SelectedFilterValuesFromListBox = @()
$Global:SelectedExportAttributes = $null
$Global:UserCountPreviewJob = $null

# Async Funktion zur Aktualisierung der Benutzeranzahl-Vorschau
function Update-UserCountPreviewAsync {
    try {
        if ($Global:UserCountPreviewJob -and $Global:UserCountPreviewJob.State -eq 'Running') {
            Write-Debug "Bestehender UserCountPreviewJob wird gestoppt."
            Stop-Job -Job $Global:UserCountPreviewJob | Out-Null
            Remove-Job -Job $Global:UserCountPreviewJob -Force | Out-Null
            Get-EventSubscriber -SourceIdentifier "UserCountJobChanged" -ErrorAction SilentlyContinue | Unregister-Event -Force | Out-Null
        }

        $Global:MainWindow.Dispatcher.Invoke([action]{
            $Global:XamlControls.TextBlockUserCountPreview.Text = "Zähle Benutzer..."
            $Global:XamlControls.ProgressBarStatus.IsIndeterminate = $true
            $Global:XamlControls.TextBlockStatus.Text = "Aktualisiere Benutzeranzahl-Vorschau..."
        })

        $filterType = $Global:XamlControls.ComboBoxFilterType.SelectedItem
        $selectedValues = $Global:SelectedFilterValuesFromListBox
        $includeDisabledUsersInCount = $Global:XamlControls.CheckBoxIncludeDisabled.IsChecked -eq $true

        if (-not $filterType -or -not $selectedValues) {
            $Global:MainWindow.Dispatcher.Invoke([action]{
                $Global:XamlControls.TextBlockUserCountPreview.Text = "(Filter unvollständig)"
                $Global:XamlControls.ProgressBarStatus.IsIndeterminate = $false
                $Global:XamlControls.TextBlockStatus.Text = "Filterkriterien für Vorschau unvollständig."
            })
            return
        }

        Ensure-ADModule

        $applyEnabledFilterForCount = (-not $includeDisabledUsersInCount)
        $ldapFilterForCount = Build-LdapFilter -Type $filterType -Values $selectedValues -FilterByEnabledState $applyEnabledFilterForCount -UsersShouldBeEnabled $true

        Write-Debug "Update-UserCountPreviewAsync: LDAP Filter für Zählung = $ldapFilterForCount"

        $scriptBlock = {
            param($LdapFilterParam)
            try {
                Import-Module ActiveDirectory -ErrorAction Stop
                $count = (Get-ADUser -LDAPFilter $LdapFilterParam -Properties 'PrimaryGroupID' -ErrorAction Stop).Count
                return $count
            }
            catch {
                Write-Warning "Fehler im Hintergrundjob (UserCountPreview): $($_.Exception.Message)"
                return -1
            }
        }
        
        $Global:UserCountPreviewJob = Start-Job -ScriptBlock $scriptBlock -ArgumentList $ldapFilterForCount
        
        Register-ObjectEvent -InputObject $Global:UserCountPreviewJob -EventName StateChanged -SourceIdentifier "UserCountJobChanged" -Action {
            param($Sender, $EventArgs)
            try {
                $job = $Sender
                if ($job.State -in ('Completed', 'Failed', 'Stopped')) {
                    $countResult = Receive-Job -Job $job -Keep
                    Remove-Job -Job $job -Force | Out-Null
                    Unregister-Event -SourceIdentifier "UserCountJobChanged" | Out-Null

                    $Global:MainWindow.Dispatcher.Invoke([action]{
                        if ($job.State -eq 'Failed' -or $countResult -is [array] -or $countResult -eq -1) {
                            $Global:XamlControls.TextBlockUserCountPreview.Text = "Fehler bei Zählung"
                            $errorMsg = "Fehler bei der Benutzerzählung-Vorschau."
                            if ($job.ChildJobs[0].Error) {
                                $errorMsg += " Details: " + ($job.ChildJobs[0].Error | Out-String)
                            }
                            $Global:XamlControls.TextBlockStatus.Text = $errorMsg
                        } elseif ($countResult -eq $null) {
                             $Global:XamlControls.TextBlockUserCountPreview.Text = "0 Benutzer (Vorschau)"
                        } else {
                            $Global:XamlControls.TextBlockUserCountPreview.Text = "$($countResult) Benutzer (Vorschau)"
                            $Global:XamlControls.TextBlockStatus.Text = "Benutzeranzahl-Vorschau aktualisiert."
                        }
                        $Global:XamlControls.ProgressBarStatus.IsIndeterminate = $false
                    })
                    $Global:UserCountPreviewJob = $null
                }
            }
            catch {
                Write-Error "Fehler im UserCountJob StateChanged Event Handler: $($_.Exception.Message)"
                 $Global:MainWindow.Dispatcher.Invoke([action]{
                    $Global:XamlControls.TextBlockUserCountPreview.Text = "Fehler (Event Zählung)"
                    $Global:XamlControls.ProgressBarStatus.IsIndeterminate = $false
                })
            }
        } | Out-Null
    }
    catch {
        $errorMessage = "Fehler in Update-UserCountPreviewAsync Hauptfunktion: $($_.Exception.Message)"
        Write-Error $errorMessage
        $Global:MainWindow.Dispatcher.Invoke([action]{
            $Global:XamlControls.TextBlockUserCountPreview.Text = "Fehler bei Zählung"
            $Global:XamlControls.TextBlockStatus.Text = $errorMessage
            $Global:XamlControls.ProgressBarStatus.IsIndeterminate = $false
        })
    }
}

function Get-RawFilterValues {
    param($Type)
    Write-Debug "Get-RawFilterValues: Typ = $Type"
    try {
        switch ($Type) {
            'OU'         { Get-ADOrganizationalUnit -Filter * | Sort-Object DistinguishedName | Select-Object -ExpandProperty DistinguishedName }
            'Department' { Get-ADUser -Filter * -Properties Department | Where-Object {$_.Department} | Select-Object -ExpandProperty Department -Unique | Sort-Object }
            'Company'    { Get-ADUser -Filter * -Properties Company    | Where-Object {$_.Company}    | Select-Object -ExpandProperty Company    -Unique | Sort-Object }
            'GroupType'  { @('Security', 'Distribution', 'Global', 'Universal', 'DomainLocal') }
            default      { Write-Warning "Unbekannter Filtertyp in Get-RawFilterValues: $Type"; return @() }
        }
    }
    catch {
        Write-Error "Fehler in Get-RawFilterValues beim Abrufen von '$Type': $_"
        [System.Windows.MessageBox]::Show("Fehler beim Abrufen der Filterwerte für '$Type': $($_.Exception.Message)", "Datenabruffehler", "OK", "Error")
        return @()
    }
}

function Populate-FilterValuesListBox {
    param ([System.Windows.Controls.ListBox]$ListBoxElement)
    try {
        $selectedFilterType = $Global:XamlControls.ComboBoxFilterType.SelectedItem
        if (-not $selectedFilterType) {
            $ListBoxElement.ItemsSource = @()
            $Global:XamlControls.TextBlockStatus.Text = "Bitte zuerst einen Filtertyp auswählen."
            return
        }

        $Global:XamlControls.TextBlockStatus.Text = "Lade Werte für Filtertyp '$selectedFilterType'..."
        $Global:XamlControls.ProgressBarStatus.IsIndeterminate = $true

        $rawValues = Get-RawFilterValues -Type $selectedFilterType
        
        $listBoxItems = foreach ($value in $rawValues) {
            [PSCustomObject]@{ DisplayName = $value; IsSelected = $false }
        }
        
        $ListBoxElement.ItemsSource = $listBoxItems
        $Global:XamlControls.TextBlockSelectedFilterValues.Text = "Ausgewählte Werte: (Keine)"
        $Global:SelectedFilterValuesFromListBox = @()
        Update-UserCountPreviewAsync

        if ($listBoxItems.Count -eq 0) {
            $Global:XamlControls.TextBlockStatus.Text = "Keine Werte für Filtertyp '$selectedFilterType' gefunden."
        } else {
            $Global:XamlControls.TextBlockStatus.Text = "Werte für '$selectedFilterType' geladen. Bitte auswählen."
        }
    }
    catch {
        Write-Error "Fehler in Populate-FilterValuesListBox: $_"
        $Global:XamlControls.TextBlockStatus.Text = "Fehler beim Laden der Filterwerte: $($_.Exception.Message)"
        [System.Windows.MessageBox]::Show("Fehler beim Laden der Filterwerte-Liste: $($_.Exception.Message)", "Listenfehler", "OK", "Error")
    }
    finally {
        $Global:XamlControls.ProgressBarStatus.IsIndeterminate = $false
    }
}

function Get-SelectedItemsFromFilterValuesListBox {
    param ([System.Windows.Controls.ListBox]$ListBoxElement)
    $selectedItems = @()
    if ($ListBoxElement.ItemsSource) {
        foreach ($item in $ListBoxElement.ItemsSource) {
            if ($item.IsSelected) {
                $selectedItems += $item.DisplayName
            }
        }
    }
    $Global:SelectedFilterValuesFromListBox = $selectedItems
    $Global:XamlControls.TextBlockSelectedFilterValues.Text = "Ausgewählte Werte: $(if($selectedItems.Count -gt 0) {$selectedItems -join ', '} else {'(Keine)'})"
    return $selectedItems
}

function Handle-ListBoxFilterValueClick {
    param($sender, $e)

    try {
        $clickedElement = $e.OriginalSource
        if ($clickedElement -is [System.Windows.Controls.CheckBox]) {
            Get-SelectedItemsFromFilterValuesListBox -ListBoxElement $Global:XamlControls.ListBoxFilterValues
            Update-UserCountPreviewAsync
        }
    }
    catch {
        $errMsg = "Fehler in Handle-ListBoxFilterValueClick: $($_.Exception.Message)"
        Write-Error $errMsg
        $Global:XamlControls.TextBlockStatus.Text = "Fehler bei ListBox-Klick: $errMsg"
    }
}

function Select-ExportAttributesGui {
    try {
        $availableAttributes = Get-SelectableADAttributes | Sort-Object
        $selectedAttributes = $availableAttributes | 
            ForEach-Object {
                [PSCustomObject]@{
                    Attribute = $_
                    Selected = ($Global:SelectedExportAttributes -contains $_)
                }
            } | 
            Out-GridView -Title 'AD-Attribute für Export auswählen' -PassThru
        
        if ($selectedAttributes) {
            $Global:SelectedExportAttributes = $selectedAttributes | Where-Object Selected | Select-Object -ExpandProperty Attribute
            $Global:XamlControls.TextBlockSelectedAttributesCount.Text = "Ausgewählte Attribute: $($Global:SelectedExportAttributes.Count)"
            $Global:XamlControls.TextBlockStatus.Text = "$($Global:SelectedExportAttributes.Count) Attribute für den Export ausgewählt."
        } else {
            $Global:XamlControls.TextBlockSelectedAttributesCount.Text = "Ausgewählte Attribute: (Keine ausgewählt)"
            $Global:XamlControls.TextBlockStatus.Text = "Keine Attribute für den Export ausgewählt."
        }
    }
    catch {
        $errMsg = "Fehler in Select-ExportAttributesGui: $($_.Exception.Message)"
        Write-Error $errMsg
        $Global:XamlControls.TextBlockStatus.Text = "Fehler bei Attributauswahl: $errMsg"
        try {[System.Windows.MessageBox]::Show($errMsg, "Attributauswahl Fehler", "OK", "Error")} catch {}
    }
}

function Get-QuickReport {
    param(
        [ValidateSet('InactiveUsers','EmptyGroups','AdminGroups','NoLogin','PasswordExpiry','LargeGroups')]
        [string]$ReportType
    )
    
    try {
        switch ($ReportType) {
            'InactiveUsers' {
                $cutoffDate = (Get-Date).AddDays(-90)
                Get-ADUser -Filter * -Properties LastLogonDate, Enabled | 
                    Where-Object { $_.Enabled -eq $false -or $_.LastLogonDate -lt $cutoffDate } |
                    Select-Object Name, SamAccountName, Enabled, LastLogonDate, DistinguishedName
            }
            'EmptyGroups' {
                $allGroups = Get-ADGroup -Filter * -Properties Members
                foreach ($group in $allGroups) {
                    $memberCount = (Get-ADGroupMember -Identity $group.DistinguishedName -ErrorAction SilentlyContinue | Measure-Object).Count
                    if ($memberCount -eq 0) {
                        [PSCustomObject]@{
                            Name = $group.Name
                            Description = $group.Description
                            GroupCategory = $group.GroupCategory
                            Created = $group.Created
                            DistinguishedName = $group.DistinguishedName
                        }
                    }
                }
            }
            'AdminGroups' {
                Get-ADGroup -Filter "Name -like '*admin*' -or Name -like '*domain*' -or Name -like '*enterprise*'" -Properties Members, Description |
                    ForEach-Object {
                        $memberCount = (Get-ADGroupMember -Identity $_.DistinguishedName -ErrorAction SilentlyContinue | Measure-Object).Count
                        [PSCustomObject]@{
                            Name = $_.Name
                            Description = $_.Description
                            MemberCount = $memberCount
                            GroupCategory = $_.GroupCategory
                            GroupScope = $_.GroupScope
                            DistinguishedName = $_.DistinguishedName
                        }
                    }
            }
            'NoLogin' {
                Get-ADUser -Filter * -Properties LastLogonDate | 
                    Where-Object { $_.LastLogonDate -eq $null -and $_.Enabled -eq $true } |
                    Select-Object Name, SamAccountName, Created, DistinguishedName
            }
            'PasswordExpiry' {
                $futureDate = (Get-Date).AddDays(30)
                Get-ADUser -Filter * -Properties PasswordLastSet, PasswordNeverExpires | 
                    Where-Object { $_.PasswordNeverExpires -eq $false } |
                    ForEach-Object {
                        $maxPasswordAge = (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge
                        if ($maxPasswordAge -and $_.PasswordLastSet) {
                            $expiryDate = $_.PasswordLastSet.AddDays($maxPasswordAge.Days)
                            if ($expiryDate -le $futureDate) {
                                [PSCustomObject]@{
                                    Name = $_.Name
                                    SamAccountName = $_.SamAccountName
                                    PasswordLastSet = $_.PasswordLastSet
                                    ExpiryDate = $expiryDate
                                    DaysToExpiry = ($expiryDate - (Get-Date)).Days
                                    DistinguishedName = $_.DistinguishedName
                                }
                            }
                        }
                    }
            }
            'LargeGroups' {
                Get-ADGroup -Filter * -Properties Members |
                    ForEach-Object {
                        $memberCount = (Get-ADGroupMember -Identity $_.DistinguishedName -ErrorAction SilentlyContinue | Measure-Object).Count
                        if ($memberCount -gt 50) {
                            [PSCustomObject]@{
                                Name = $_.Name
                                Description = $_.Description
                                MemberCount = $memberCount
                                GroupCategory = $_.GroupCategory
                                DistinguishedName = $_.DistinguishedName
                            }
                        }
                    } | Sort-Object MemberCount -Descending
            }
        }
    }
    catch {
        Write-Error "Fehler beim Schnellbericht '$ReportType': $_"
        throw
    }
}

function Update-FilterOptionsForAuditType {
    if ($Global:XamlControls.RadioUsers.IsChecked) {
        $Global:XamlControls.ComboBoxFilterType.ItemsSource = @("OU", "Department", "Company")
    }
    elseif ($Global:XamlControls.RadioGroups.IsChecked) {
        $Global:XamlControls.ComboBoxFilterType.ItemsSource = @("OU", "GroupType")
    }
    elseif ($Global:XamlControls.RadioMemberships.IsChecked) {
        $Global:XamlControls.ComboBoxFilterType.ItemsSource = @("OU", "User", "Group")
    }
    $Global:XamlControls.ComboBoxFilterType.SelectedIndex = 0
}

function Start-QuickReport {
    param([string]$ReportType)
    
    try {
        $Global:XamlControls.TextBlockStatus.Text = "Erstelle Schnellbericht: $ReportType"
        $Global:XamlControls.ProgressBarStatus.IsIndeterminate = $true
        
        Ensure-ADModule
        $reportData = Get-QuickReport -ReportType $ReportType
        
        if ($reportData) {
            $reportTitle = switch ($ReportType) {
                'InactiveUsers' { "Inaktive Benutzer" }
                'EmptyGroups' { "Leere Gruppen" }
                'AdminGroups' { "Administrator Gruppen" }
                'NoLogin' { "Niemals angemeldete Benutzer" }
                'PasswordExpiry' { "Passwort läuft ab" }
                'LargeGroups' { "Große Gruppen" }
            }
            
            # $outputFormat wird immer CSV sein, da ComboBox nur CSV anbietet
            $outputFormat = 'CSV' # $Global:XamlControls.ComboBoxOutputType.SelectedItem
            $outputPath = $Global:XamlControls.TextBoxOutputPath.Text
            
            $exportedFile = Export-Data -Collection $reportData -Type $outputFormat -Path $outputPath -Attributes ($reportData[0].PSObject.Properties.Name) -ReportTitle $reportTitle
            
            $Global:XamlControls.TextBlockStatus.Text = "Schnellbericht erstellt: $exportedFile"
            [System.Windows.MessageBox]::Show("Schnellbericht '$reportTitle' wurde erfolgreich erstellt:`n$exportedFile", "Bericht erstellt", "OK", "Information")
        } else {
            $Global:XamlControls.TextBlockStatus.Text = "Keine Daten für Schnellbericht gefunden"
            [System.Windows.MessageBox]::Show("Für den Bericht '$ReportType' wurden keine Daten gefunden.", "Keine Ergebnisse", "OK", "Information")
        }
    }
    catch {
        $errorMsg = "Fehler beim Erstellen des Schnellberichts: $($_.Exception.Message)"
        $Global:XamlControls.TextBlockStatus.Text = $errorMsg
        [System.Windows.MessageBox]::Show($errorMsg, "Berichtsfehler", "OK", "Error")
    }
    finally {
        $Global:XamlControls.ProgressBarStatus.IsIndeterminate = $false
    }
}

function Start-ExportProcess {
    param(
        [string]$GuiFilterType,
        [string]$GuiOutputType, # Wird CSV sein
        [string]$GuiOutputPath,
        [bool]$GuiIncludeDisabled,
        [string[]]$GuiSelectedValues,
        [string[]]$GuiSelectedAttributes,
        [string]$GuiRequiredAttribute,
        [string]$GuiSortBy,
        [string]$GuiExcludePattern
    )
    
    try {
        $Global:XamlControls.TextBlockStatus.Text = "Bestimme Audit-Typ..."
        
        Ensure-ADModule
        
        $auditData = $null
        $reportTitle = ""
        
        if ($Global:XamlControls.RadioUsers.IsChecked) {
            $reportTitle = "Benutzer Audit"
            $Global:XamlControls.TextBlockStatus.Text = "Führe Benutzer-Audit durch..."
            
            # Verwende bestehende Benutzer-Audit-Logik
            $ldap = Build-LdapFilter $GuiFilterType $GuiSelectedValues -FilterByEnabledState (-not $GuiIncludeDisabled) -UsersShouldBeEnabled $true -ObjectClass 'user'
            $auditData = Get-ADUser -LDAPFilter $ldap -Properties $GuiSelectedAttributes
        }
        elseif ($Global:XamlControls.RadioGroups.IsChecked) {
            $reportTitle = "Gruppen Audit" 
            $Global:XamlControls.TextBlockStatus.Text = "Führe Gruppen-Audit durch..."
            
            $includeNested = $Global:XamlControls.CheckBoxIncludeNested.IsChecked -eq $true
            $showEmpty = $Global:XamlControls.CheckBoxShowEmptyGroups.IsChecked -eq $true
            $auditData = Get-GroupAuditData -FilterValues $GuiSelectedValues -FilterType $GuiFilterType -IncludeDisabled $GuiIncludeDisabled -IncludeNested $includeNested -ShowEmptyGroups $showEmpty
        }
        elseif ($Global:XamlControls.RadioMemberships.IsChecked) {
            $reportTitle = "Mitgliedschaften Audit"
            $Global:XamlControls.TextBlockStatus.Text = "Führe Mitgliedschafts-Audit durch..."
            
            $includeNested = $Global:XamlControls.CheckBoxIncludeNested.IsChecked -eq $true
            $auditData = Get-MembershipAuditData -FilterValues $GuiSelectedValues -FilterType $GuiFilterType -IncludeNested $includeNested
        }
        
        if ($auditData -and $auditData.Count -gt 0) {
            $Global:XamlControls.TextBlockStatus.Text = "Exportiere Audit-Daten..."
            
            # $formats ist jetzt immer nur CSV
            $formats = @('CSV') # if ($GuiOutputType -eq 'Both') { @('HTML', 'CSV') } else { @($GuiOutputType) }
            $exportedFiles = @()
            
            foreach ($format in $formats) { # Schleife läuft nur einmal für CSV
                $exportedFile = Export-Data -Collection $auditData -Type $format -Path $GuiOutputPath -Attributes $GuiSelectedAttributes -SortBy $GuiSortBy -ExcludePattern $GuiExcludePattern -RequiredAttribute $GuiRequiredAttribute -ReportTitle $reportTitle
                $exportedFiles += $exportedFile
            }
            
            $Global:XamlControls.TextBlockStatus.Text = "Audit erfolgreich abgeschlossen."
            $message = "$reportTitle wurde erfolgreich erstellt:`n" + ($exportedFiles -join "`n")
            [System.Windows.MessageBox]::Show($message, "Audit abgeschlossen", "OK", "Information")
        } else {
            $Global:XamlControls.TextBlockStatus.Text = "Keine Daten für Audit gefunden."
            [System.Windows.MessageBox]::Show("Für die gewählten Kriterien wurden keine Daten gefunden.", "Keine Ergebnisse", "OK", "Information")
        }
    }
    catch {
        $errorMessage = "Fehler im Audit-Prozess: $($_.Exception.Message)"
        $Global:XamlControls.TextBlockStatus.Text = $errorMessage
        [System.Windows.MessageBox]::Show($errorMessage, "Audit-Fehler", "OK", "Error")
    }
    finally {
        $Global:XamlControls.ProgressBarStatus.IsIndeterminate = $false
    }
}

function Start-ExportProcessGui {
    try {
        $Global:XamlControls.TextBlockStatus.Text = "Sammle Exportparameter aus GUI..."
        $Global:XamlControls.ProgressBarStatus.IsIndeterminate = $true

        $guiFilterType = $Global:XamlControls.ComboBoxFilterType.SelectedItem
        $guiOutputType = $Global:XamlControls.ComboBoxOutputType.SelectedItem
        $guiOutputPath = $Global:XamlControls.TextBoxOutputPath.Text
        $guiIncludeDisabled = $Global:XamlControls.CheckBoxIncludeDisabled.IsChecked -eq $true
        
        Get-SelectedItemsFromFilterValuesListBox -ListBoxElement $Global:XamlControls.ListBoxFilterValues
        $guiSelectedValues = $Global:SelectedFilterValuesFromListBox
        
        $guiSelectedAttributes = $Global:SelectedExportAttributes
        
        $guiRequiredAttribute = $Global:XamlControls.TextBoxRequiredAttribute.Text
        $guiSortBy = $Global:XamlControls.TextBoxSortBy.Text
        $guiExcludePattern = $Global:XamlControls.TextBoxExcludePattern.Text

        # Validierungen
        if (-not $guiFilterType) {
            [System.Windows.MessageBox]::Show("Bitte wählen Sie einen Filtertyp aus.", "Validierungsfehler", "OK", "Warning")
            $Global:XamlControls.TextBlockStatus.Text = "Validierungsfehler: Filtertyp fehlt."
            $Global:XamlControls.ProgressBarStatus.IsIndeterminate = $false
            return
        }
        if (-not $guiSelectedValues -or $guiSelectedValues.Count -eq 0) {
            [System.Windows.MessageBox]::Show("Bitte wählen Sie mindestens einen Filterwert aus (über 'Werte laden').", "Validierungsfehler", "OK", "Warning")
            $Global:XamlControls.TextBlockStatus.Text = "Validierungsfehler: Filterwerte fehlen."
            $Global:XamlControls.ProgressBarStatus.IsIndeterminate = $false
            return
        }
        if (-not $guiSelectedAttributes -or $guiSelectedAttributes.Count -eq 0) {
            [System.Windows.MessageBox]::Show("Bitte wählen Sie mindestens ein Attribut für den Export aus (über 'Attribute auswählen...').", "Validierungsfehler", "OK", "Warning")
            $Global:XamlControls.TextBlockStatus.Text = "Validierungsfehler: Exportattribute fehlen."
            $Global:XamlControls.ProgressBarStatus.IsIndeterminate = $false
            return
        }
        if (-not $guiOutputPath -or -not (Test-Path (Split-Path $guiOutputPath -Parent))) {
             [System.Windows.MessageBox]::Show("Der übergeordnete Ordner des Exportpfads '$guiOutputPath' existiert nicht oder ist ungültig. Bitte wählen Sie einen gültigen Pfad.", "Validierungsfehler", "OK", "Warning")
            $Global:XamlControls.TextBlockStatus.Text = "Validierungsfehler: Exportpfad ungültig."
            $Global:XamlControls.ProgressBarStatus.IsIndeterminate = $false
            return
        }

        $Global:XamlControls.TextBlockStatus.Text = "Starte Exportprozess..."
        Start-ExportProcess -GuiFilterType $guiFilterType `
                            -GuiOutputType $guiOutputType `
                            -GuiOutputPath $guiOutputPath `
                            -GuiIncludeDisabled $guiIncludeDisabled `
                            -GuiSelectedValues $guiSelectedValues `
                            -GuiSelectedAttributes $guiSelectedAttributes `
                            -GuiRequiredAttribute $guiRequiredAttribute `
                            -GuiSortBy $guiSortBy `
                            -GuiExcludePattern $guiExcludePattern
        
    }
    catch {
        $errMsg = "Schwerwiegender Fehler in Start-ExportProcessGui: $($_.Exception.Message)"
        Write-Error $errMsg
        $Global:XamlControls.TextBlockStatus.Text = "Fehler: $errMsg"
        try {[System.Windows.MessageBox]::Show($errMsg, "Exportfehler", "OK", "Error")} catch {}
    }
    finally {
        $Global:XamlControls.ProgressBarStatus.IsIndeterminate = $false
    }
}

function Initialize-Gui {
    try {
        # XAML direkt laden (keine externe Datei nötig)
        [xml]$xaml = $xamlContent
        $reader = (New-Object System.Xml.XmlNodeReader $xaml)
        $Global:MainWindow = [Windows.Markup.XamlReader]::Load($reader)

        # Korrigierte Steuerelemente-Sammlung
        # NamespaceManager für XPath-Abfragen mit Prefixen (wie x:Name) erstellen
        $xamlNs = New-Object System.Xml.XmlNamespaceManager($xaml.NameTable)
        $xamlNs.AddNamespace("x", "http://schemas.microsoft.com/winfx/2006/xaml")

        # Alle XML-Knoten auswählen, die ein x:Name-Attribut haben
        $xaml.SelectNodes("//*[@x:Name]", $xamlNs) | ForEach-Object {
            $node = $_
            # Den Wert des x:Name-Attributs abrufen
            $controlName = $node.GetAttribute("Name", "http://schemas.microsoft.com/winfx/2006/xaml")
            
            if (-not [string]::IsNullOrEmpty($controlName)) {
                # Das tatsächliche WPF-Steuerelement über FindName suchen
                $foundControl = $Global:MainWindow.FindName($controlName)
                if ($null -ne $foundControl) {
                    $Global:XamlControls[$controlName] = $foundControl
                } else {
                    # Warnung ausgeben, falls ein im XAML definiertes x:Name nicht per FindName gefunden wird.
                    # Dies kann passieren, wenn das Element z.B. in einem DataTemplate ist und nicht im Haupt-Namensbereich des Fensters.
                    Write-Warning "Initialize-Gui: Steuerelement '$controlName' (definiert mit x:Name im XAML) wurde von FindName() nicht im Hauptfenster gefunden."
                }
            }
        }

        # Explizite Prüfung, ob die kritischen ComboBoxen gefunden wurden
        if ($null -eq $Global:XamlControls.ComboBoxFilterType) { # Kurzform für $Global:XamlControls["ComboBoxFilterType"]
            throw "GUI-Steuerelement ComboBoxFilterType konnte nicht gefunden werden. Überprüfen Sie das XAML und den Namen des Steuerelements."
        }
        if ($null -eq $Global:XamlControls.ComboBoxOutputType) {
            throw "GUI-Steuerelement ComboBoxOutputType konnte nicht gefunden werden. Überprüfen Sie das XAML und den Namen des Steuerelements."
        }
        if ($null -eq $Global:XamlControls.RadioUsers) {
            throw "GUI-Steuerelement RadioUsers konnte nicht gefunden werden."
        }
        # Fügen Sie hier bei Bedarf weitere Prüfungen für wichtige Steuerelemente hinzu

        # ComboBoxen initialisieren
        # ComboBoxOutputType (unabhängig vom Audit-Typ)
        $Global:XamlControls.ComboBoxOutputType.ItemsSource = @("CSV") # Nur CSV anbieten
        $Global:XamlControls.ComboBoxOutputType.SelectedIndex = 0
        
        # ComboBoxFilterType basierend auf dem Standard-Audit-Typ (Benutzer) initialisieren.
        # Da RadioUsers.IsChecked="True" im XAML ist, setzen wir die ItemsSource und SelectedIndex hier direkt für den Startzustand.
        # Update-FilterOptionsForAuditType wird dann von den Event-Handlern für spätere Änderungen aufgerufen.
        if ($Global:XamlControls.RadioUsers.IsChecked) {
            $Global:XamlControls.ComboBoxFilterType.ItemsSource = @("OU", "Department", "Company")
        }
        elseif ($Global:XamlControls.RadioGroups.IsChecked) { # Fallback, falls XAML-Standard geändert wird
            $Global:XamlControls.ComboBoxFilterType.ItemsSource = @("OU", "GroupType")
        }
        elseif ($Global:XamlControls.RadioMemberships.IsChecked) { # Fallback
            $Global:XamlControls.ComboBoxFilterType.ItemsSource = @("OU", "User", "Group")
        }
        else { 
            # Standardfall, falls keiner der RadioButtons explizit im XAML als IsChecked=True markiert ist,
            # oder um einen definierten Zustand sicherzustellen.
            $Global:XamlControls.ComboBoxFilterType.ItemsSource = @("OU", "Department", "Company") 
        }
        $Global:XamlControls.ComboBoxFilterType.SelectedIndex = 0
        
        $Global:XamlControls.TextBoxOutputPath.Text = $OutputPath

        # Event-Handler zuweisen (erweitert)
        # Diese Handler rufen Update-FilterOptionsForAuditType auf, um ComboBoxFilterType zu aktualisieren,
        # wenn der Benutzer den Audit-Typ ÄNDERT.
        $Global:XamlControls.RadioUsers.add_Checked({ Update-FilterOptionsForAuditType })
        $Global:XamlControls.RadioGroups.add_Checked({ Update-FilterOptionsForAuditType })
        $Global:XamlControls.RadioMemberships.add_Checked({ Update-FilterOptionsForAuditType })
        
        # Schnellbericht-Buttons
        $Global:XamlControls.ButtonQuickReportInactiveUsers.add_Click({ Start-QuickReport -ReportType 'InactiveUsers' })
        $Global:XamlControls.ButtonQuickReportEmptyGroups.add_Click({ Start-QuickReport -ReportType 'EmptyGroups' })
        $Global:XamlControls.ButtonQuickReportAdmins.add_Click({ Start-QuickReport -ReportType 'AdminGroups' })
        $Global:XamlControls.ButtonQuickReportNoLogin.add_Click({ Start-QuickReport -ReportType 'NoLogin' })
        $Global:XamlControls.ButtonQuickReportPasswordExpiry.add_Click({ Start-QuickReport -ReportType 'PasswordExpiry' })
        $Global:XamlControls.ButtonQuickReportLargeGroups.add_Click({ Start-QuickReport -ReportType 'LargeGroups' })

        # Event-Handler zuweisen
        try {
            if ($Global:XamlControls.ButtonBrowseOutputPath) {
                $Global:XamlControls.ButtonBrowseOutputPath.add_Click({
                    try {
                        $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
                        $folderBrowser.Description = "Wählen Sie einen Ordner für den Export"
                        $folderBrowser.ShowNewFolderButton = $true
                        if ($Global:XamlControls.TextBoxOutputPath.Text -and (Test-Path $Global:XamlControls.TextBoxOutputPath.Text)) {
                            $folderBrowser.SelectedPath = $Global:XamlControls.TextBoxOutputPath.Text
                        }
                        if ($folderBrowser.ShowDialog((New-Object System.Windows.Forms.NativeWindow)) -eq [System.Windows.Forms.DialogResult]::OK) {
                            $Global:XamlControls.TextBoxOutputPath.Text = $folderBrowser.SelectedPath
                            $Global:XamlControls.TextBlockStatus.Text = "Exportpfad ausgewählt: $($folderBrowser.SelectedPath)"
                        }
                    }
                    catch {
                        Write-Error "Fehler beim Anzeigen des Ordnerauswahldialogs: $_"
                        $Global:XamlControls.TextBlockStatus.Text = "Fehler (Ordnerdialog): $($_.Exception.Message)"
                        try { [System.Windows.MessageBox]::Show("Fehler beim Ordnerauswahldialog: $($_.Exception.Message)", "Dialogfehler", "OK", "Error") } catch {}
                    }
                })
            }

            if ($Global:XamlControls.ComboBoxFilterType) {
                $Global:XamlControls.ComboBoxFilterType.add_SelectionChanged({
                    try {
                        $Global:XamlControls.ListBoxFilterValues.ItemsSource = @() 
                        $Global:XamlControls.TextBlockSelectedFilterValues.Text = "Ausgewählte Werte: (Keine)"
                        $Global:SelectedFilterValuesFromListBox = @()
                        Update-UserCountPreviewAsync 
                        $Global:XamlControls.TextBlockStatus.Text = "Filtertyp geändert. Bitte auf 'Werte anzeigen/aktualisieren' klicken."
                    } 
                    catch {
                        Write-Error "Fehler im ComboBoxFilterType.SelectionChanged: $_"
                        $Global:XamlControls.TextBlockStatus.Text = "Fehler (Filtertyp Auswahl): $($_.Exception.Message)"
                    }
                })
            }

            if ($Global:XamlControls.ButtonPopulateFilterValues) {
                $Global:XamlControls.ButtonPopulateFilterValues.add_Click({
                    try {
                        Populate-FilterValuesListBox -ListBoxElement $Global:XamlControls.ListBoxFilterValues
                    } 
                    catch {
                        Write-Error "Fehler im ButtonPopulateFilterValues.Click: $_"
                        $Global:XamlControls.TextBlockStatus.Text = "Fehler (Werte laden): $($_.Exception.Message)"
                    }
                })
            }

            if ($Global:XamlControls.ListBoxFilterValues) {
                $Global:XamlControls.ListBoxFilterValues.add_MouseLeftButtonUp({param($s,$e) 
                    try {
                        Handle-ListBoxFilterValueClick -sender $s -e $e 
                    }
                    catch {
                        Write-Error "Fehler im ListBoxFilterValues.MouseLeftButtonUp: $_"
                        $Global:XamlControls.TextBlockStatus.Text = "Fehler (Filterwert Auswahl): $($_.Exception.Message)"
                    }
                })
            }

            if ($Global:XamlControls.CheckBoxIncludeDisabled) {
                $Global:XamlControls.CheckBoxIncludeDisabled.add_Click({
                    try {
                        Get-SelectedItemsFromFilterValuesListBox -ListBoxElement $Global:XamlControls.ListBoxFilterValues
                        Update-UserCountPreviewAsync 
                    }
                    catch {
                        Write-Error "Fehler im CheckBoxIncludeDisabled.Click: $_"
                        $Global:XamlControls.TextBlockStatus.Text = "Fehler (Checkbox Deaktivierte): $($_.Exception.Message)"
                    }
                })
            }

            if ($Global:XamlControls.ButtonSelectExportAttributes) {
                $Global:XamlControls.ButtonSelectExportAttributes.add_Click({
                    try {
                        Select-ExportAttributesGui
                    }
                    catch {
                        Write-Error "Fehler im ButtonSelectExportAttributes.Click: $_"
                        $Global:XamlControls.TextBlockStatus.Text = "Fehler (Attributauswahl): $($_.Exception.Message)"
                    }
                })
            }

            if ($Global:XamlControls.ButtonStartExport) {
                $Global:XamlControls.ButtonStartExport.add_Click({
                    try {
                        Start-ExportProcessGui
                    }
                    catch {
                        Write-Error "Fehler im ButtonStartExport.Click: $_"
                        $Global:XamlControls.TextBlockStatus.Text = "Fehler (Exportstart): $($_.Exception.Message)"
                    }
                })
            }
            
            if ($Global:XamlControls.ButtonClose) {
                $Global:XamlControls.ButtonClose.add_Click({
                    try {
                        $Global:MainWindow.Close()
                    }
                    catch {
                        Write-Error "Fehler im ButtonClose.Click: $_"
                    }
                })
            }

        }
        catch {
            $errMsg = "Fehler beim Zuweisen der GUI Event-Handler: $($_.Exception.Message)"
            Write-Error $errMsg
            try { [System.Windows.MessageBox]::Show($errMsg, "GUI Event-Handler Fehler", "OK", "Error") } catch {}
        }

        Write-Host "EasyADAudit GUI wird gestartet..." -ForegroundColor Green
        $null = $Global:MainWindow.ShowDialog()
    }
    catch {
        Write-Error "GUI-Initialisierungsfehler: $_"
        # Optional: Zeige den Fehler auch in einem MessageBox an, falls die Konsole nicht sichtbar ist.
        # try { [System.Windows.MessageBox]::Show("Schwerwiegender GUI-Initialisierungsfehler: $($_.Exception.Message)", "GUI Fehler", "OK", "Error") } catch {}
        exit 1
    }
}
#endregion GUI

#region ─── Hauptlogik ──────────────────────────────────────────────────────────
# Konsolen-Modus Hauptlogik
function Start-ConsoleMode {
    Write-Host "EasyADAudit wird im Konsolen-Modus ausgeführt..." -ForegroundColor Cyan
    
    try {
        Ensure-ADModule
        
        $filterType = Choose-FilterType -Current $FilterType
        $values = Select-Values -Type $filterType
        $outputType = Choose-OutputType -Current $OutputType # Wird jetzt immer 'CSV' sein
        $attributes = Select-ExportAttributes -Current $null
        
        if (-not $values) {
            Write-Warning "Keine Filterwerte ausgewählt."
            return
        }
        
        if (-not $attributes) {
            Write-Warning "Keine Exportattribute ausgewählt."
            return
        }
        
        Write-Host "Sammle AD-Daten..." -ForegroundColor Yellow
        
        $ldapFilter = Build-LdapFilter -Type $filterType -Values $values -FilterByEnabledState (-not $IncludeDisabled) -UsersShouldBeEnabled $true
        $users = Get-ADUser -LDAPFilter $ldapFilter -Properties $attributes
        
        if ($users) {
            Write-Host "Gefunden: $($users.Count) Benutzer" -ForegroundColor Green
            
            # $formats ist jetzt immer nur CSV
            $formats = @('CSV') # if ($outputType -eq 'Both') { @('HTML', 'CSV') } else { @($outputType) }
            $exportedFiles = @()
            
            foreach ($format in $formats) { # Schleife läuft nur einmal für CSV
                $exportedFile = Export-Data -Collection $users -Type $format -Path $OutputPath -Attributes $attributes -ReportTitle "AD Benutzer Export"
                $exportedFiles += $exportedFile
                Write-Host "Exportiert: $exportedFile" -ForegroundColor Green
            }
            
            Write-Host "`nExport erfolgreich abgeschlossen!" -ForegroundColor Green
            Write-Host "Dateien:" -ForegroundColor Cyan
            $exportedFiles | ForEach-Object { Write-Host "  $_" -ForegroundColor White }
        } else {
            Write-Warning "Keine Benutzer für die gewählten Kriterien gefunden."
        }
    }
    catch {
        Write-Error "Fehler im Konsolen-Modus: $_"
        exit 1
    }
}

# Haupteinstiegspunkt
if ($FilterType -or $OutputType -or $Silent) {
    # Konsolen-Modus
    Start-ConsoleMode
} else {
    # GUI-Modus
    Write-Host "Starte EasyADAudit mit moderner GUI..." -ForegroundColor Green
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms
    Initialize-Gui
}
#endregion Hauptlogik