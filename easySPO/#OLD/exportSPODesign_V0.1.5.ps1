# WPF statt Windows.Forms verwenden
Write-Host "[DEBUG] Skript gestartet - Lade benötigte Assemblies..." -ForegroundColor Cyan
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Write-Host "[DEBUG] Assemblies geladen" -ForegroundColor Cyan

# Get current execution policy for use in the form
Write-Host "[DEBUG] Prüfe aktuelle Execution Policy..." -ForegroundColor Cyan
$epCurrent = Get-ExecutionPolicy -Scope Process -ErrorAction SilentlyContinue
Write-Host "[DEBUG] Execution Policy: $epCurrent" -ForegroundColor Cyan

# Check if module is installed for use in the form
Write-Host "[DEBUG] Prüfe, ob PnP.PowerShell installiert ist..." -ForegroundColor Cyan
$moduleName = "PnP.PowerShell"
$moduleInstalled = Get-Module -ListAvailable -Name $moduleName
if ($moduleInstalled) {
    Write-Host "[DEBUG] Modul $moduleName ist installiert: Version $($moduleInstalled.Version)" -ForegroundColor Green
} else {
    Write-Host "[DEBUG] Modul $moduleName ist NICHT installiert" -ForegroundColor Yellow
}

# Die XAML-Datei für die WPF-GUI laden
function Load-XamlGui {
    Write-Host "[DEBUG] Lade XAML GUI..." -ForegroundColor Cyan
    $xamlContent = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="SharePoint Design Export"
    Width="1300"
    Height="925"
    Background="#F3F3F3"
    WindowStartupLocation="CenterScreen">
    <Window.Resources>
        <!--  Styles für Windows 11 / Metro Look  -->
        <Style x:Key="HeaderTextBlockStyle" TargetType="TextBlock">
            <Setter Property="Foreground" Value="#202020" />
            <Setter Property="FontSize" Value="24" />
            <Setter Property="FontWeight" Value="SemiBold" />
            <Setter Property="Margin" Value="10,0,0,0" />
            <Setter Property="VerticalAlignment" Value="Center" />
        </Style>

        <Style x:Key="StatusOkTextBlockStyle" TargetType="TextBlock">
            <Setter Property="Foreground" Value="#4CAF50" />
            <Setter Property="FontSize" Value="13" />
            <Setter Property="FontWeight" Value="SemiBold" />
        </Style>

        <Style x:Key="StatusErrorTextBlockStyle" TargetType="TextBlock">
            <Setter Property="Foreground" Value="#F44336" />
            <Setter Property="FontSize" Value="13" />
            <Setter Property="FontWeight" Value="SemiBold" />
        </Style>

        <Style x:Key="StatusInfoTextBlockStyle" TargetType="TextBlock">
            <Setter Property="Foreground" Value="#2196F3" />
            <Setter Property="FontSize" Value="13" />
            <Setter Property="FontWeight" Value="Normal" />
        </Style>

        <Style x:Key="LabelStyle" TargetType="TextBlock">
            <Setter Property="Foreground" Value="#505050" />
            <Setter Property="FontSize" Value="13" />
            <Setter Property="Margin" Value="0,5,0,2" />
        </Style>

        <Style x:Key="GroupHeaderStyle" TargetType="TextBlock">
            <Setter Property="Foreground" Value="#303030" />
            <Setter Property="FontSize" Value="14" />
            <Setter Property="FontWeight" Value="SemiBold" />
            <Setter Property="Margin" Value="0,0,0,10" />
        </Style>

        <Style x:Key="FooterTextBlockStyle" TargetType="TextBlock">
            <Setter Property="Foreground" Value="#505050" />
            <Setter Property="FontSize" Value="12" />
            <Setter Property="HorizontalAlignment" Value="Center" />
            <Setter Property="VerticalAlignment" Value="Center" />
        </Style>

        <Style x:Key="ModernGroupBoxStyle" TargetType="GroupBox">
            <Setter Property="BorderBrush" Value="#DDDDDD" />
            <Setter Property="BorderThickness" Value="1" />
            <Setter Property="Padding" Value="10" />
            <Setter Property="Margin" Value="0,5,0,10" />
            <Setter Property="Background" Value="#FFFFFF" />
        </Style>

        <Style x:Key="ModernTextBoxStyle" TargetType="TextBox">
            <Setter Property="Height" Value="30" />
            <Setter Property="Padding" Value="5,0" />
            <Setter Property="VerticalContentAlignment" Value="Center" />
            <Setter Property="Background" Value="White" />
            <Setter Property="BorderBrush" Value="#CCCCCC" />
            <Setter Property="BorderThickness" Value="1" />
        </Style>

        <Style
            x:Key="RequiredTextBoxStyle"
            BasedOn="{StaticResource ModernTextBoxStyle}"
            TargetType="TextBox">
            <Setter Property="BorderBrush" Value="#F44336" />
            <Style.Triggers>
                <Trigger Property="Text" Value="">
                    <Setter Property="ToolTip" Value="Dieses Feld ist erforderlich" />
                </Trigger>
                <Trigger Property="Text" Value="{x:Null}">
                    <Setter Property="ToolTip" Value="Dieses Feld ist erforderlich" />
                </Trigger>
            </Style.Triggers>
        </Style>

        <Style x:Key="ModernButtonStyle" TargetType="Button">
            <Setter Property="Height" Value="30" />
            <Setter Property="Padding" Value="15,0" />
            <Setter Property="Background" Value="#0078D7" />
            <Setter Property="Foreground" Value="White" />
            <Setter Property="BorderThickness" Value="0" />
            <Setter Property="Cursor" Value="Hand" />
            <Style.Triggers>
                <Trigger Property="IsEnabled" Value="False">
                    <Setter Property="Background" Value="#CCCCCC" />
                    <Setter Property="Foreground" Value="#888888" />
                </Trigger>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#106EBE" />
                </Trigger>
            </Style.Triggers>
        </Style>

        <Style x:Key="SecondaryButtonStyle" TargetType="Button">
            <Setter Property="Height" Value="30" />
            <Setter Property="Padding" Value="15,0" />
            <Setter Property="Background" Value="#EFEFEF" />
            <Setter Property="Foreground" Value="#505050" />
            <Setter Property="BorderThickness" Value="0" />
            <Setter Property="Cursor" Value="Hand" />
            <Style.Triggers>
                <Trigger Property="IsEnabled" Value="False">
                    <Setter Property="Background" Value="#CCCCCC" />
                    <Setter Property="Foreground" Value="#888888" />
                </Trigger>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#E0E0E0" />
                </Trigger>
            </Style.Triggers>
        </Style>

        <Style x:Key="ModernCheckBoxStyle" TargetType="CheckBox">
            <Setter Property="Margin" Value="0,5" />
            <Setter Property="VerticalContentAlignment" Value="Center" />
        </Style>

        <Style x:Key="ModernRadioButtonStyle" TargetType="RadioButton">
            <Setter Property="Margin" Value="0,5" />
            <Setter Property="VerticalContentAlignment" Value="Center" />
        </Style>

        <Style x:Key="ModernProgressBarStyle" TargetType="ProgressBar">
            <Setter Property="Height" Value="10" />
            <Setter Property="Foreground" Value="#4CAF50" />
            <Setter Property="Background" Value="#E0E0E0" />
        </Style>

        <!--  Neue Workflow Styles  -->
        <Style x:Key="WorkflowIconContainer" TargetType="Border">
            <Setter Property="Width" Value="40" />
            <Setter Property="Height" Value="40" />
            <Setter Property="Background" Value="#F0F0F0" />
            <Setter Property="BorderBrush" Value="#DDDDDD" />
            <Setter Property="BorderThickness" Value="1" />
            <Setter Property="CornerRadius" Value="20" />
            <Setter Property="Margin" Value="10,5,10,5" />
            <Setter Property="Padding" Value="8" />
            <Setter Property="HorizontalAlignment" Value="Center" />
            <Setter Property="VerticalAlignment" Value="Center" />
        </Style>

        <Style x:Key="WorkflowLabelStyle" TargetType="TextBlock">
            <Setter Property="Foreground" Value="#505050" />
            <Setter Property="FontSize" Value="12" />
            <Setter Property="FontWeight" Value="Medium" />
            <Setter Property="TextAlignment" Value="Center" />
            <Setter Property="Margin" Value="0,3,0,2" />
        </Style>

        <!--  Workflow-Schritt-Status Styles  -->
        <Style
            x:Key="WorkflowIconCompleted"
            BasedOn="{StaticResource WorkflowIconContainer}"
            TargetType="Border">
            <Setter Property="Background" Value="#E3F2FD" />
            <Setter Property="BorderBrush" Value="#2196F3" />
        </Style>

        <Style
            x:Key="WorkflowIconActive"
            BasedOn="{StaticResource WorkflowIconContainer}"
            TargetType="Border">
            <Setter Property="Background" Value="#E8F5E9" />
            <Setter Property="BorderBrush" Value="#4CAF50" />
        </Style>
    </Window.Resources>

    <Grid Margin="0,0,0,0">
        <Grid.RowDefinitions>
            <RowDefinition Height="75" />
            <!--  Header  -->
            <RowDefinition Height="*" />
            <!--  Content  -->
            <RowDefinition Height="60" />
            <!--  Footer  -->
        </Grid.RowDefinitions>

        <!--  Header  -->
        <Border Grid.Row="0" Background="#0078D7">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto" />
                    <ColumnDefinition Width="*" />
                </Grid.ColumnDefinitions>

                <!--  Logo / App Icon  -->
                <Border
                    Grid.Column="0"
                    Width="50"
                    Height="50"
                    Margin="15,0,10,0"
                    Background="White"
                    CornerRadius="5">
                    <Path
                        Margin="10"
                        Data="M12,2C6.47,2 2,6.47 2,12C2,17.53 6.47,22 12,22C17.53,22 22,17.53 22,12C22,6.47 17.53,2 12,2M15.1,7.07C15.24,7.07 15.38,7.12 15.5,7.23L16.77,8.5C17,8.72 17,9.07 16.77,9.28L15.77,10.28L13.72,8.23L14.72,7.23C14.82,7.12 14.96,7.07 15.1,7.07M13.13,8.81L15.19,10.87L9.13,16.93H7.07V14.87L13.13,8.81Z"
                        Fill="#0078D7"
                        Stretch="Uniform" />
                </Border>

                <!--  App Title  -->
                <TextBlock
                    Grid.Column="1"
                    Foreground="White"
                    Style="{StaticResource HeaderTextBlockStyle}"
                    Text="SharePoint Design Export Tool" />
            </Grid>
        </Border>

        <!--  Content Area - 3 Columns  -->
        <Grid Grid.Row="1" Margin="20">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto" />
                <!--  Workflow Overview  -->
                <RowDefinition Height="*" />
                <!--  Main Content  -->
            </Grid.RowDefinitions>

            <!--  Workflow Overview  -->
            <GroupBox
                Grid.Row="0"
                Margin="0,0,0,30"
                Header="Ablauf">
                <StackPanel HorizontalAlignment="Center" Orientation="Horizontal">
                    <!--  Schritt 1: Module installieren  -->
                    <StackPanel>
                        <Border x:Name="workflowIcon1" Style="{StaticResource WorkflowIconContainer}">
                            <Viewbox>
                                <Canvas Width="16" Height="16">
                                    <!--  Down Arrow für Download  -->
                                    <Path Data="M2,0 L14,0 L8,10 Z" Fill="Gray" />
                                </Canvas>
                            </Viewbox>
                        </Border>
                        <TextBlock Style="{StaticResource WorkflowLabelStyle}" Text="Modul" />
                        <!--  Erklärung für Schritt 1  -->
                        <TextBlock
                            FontSize="9"
                            Foreground="Gray"
                            Text="1. SPO_PnP Modul installieren   "
                            TextAlignment="Center" />
                    </StackPanel>
                    <!--  Schritt 2: MFA Login  -->
                    <StackPanel>
                        <Border x:Name="workflowIcon2" Style="{StaticResource WorkflowIconContainer}">
                            <Viewbox>
                                <Canvas Width="16" Height="16">
                                    <!--  Shield Icon  -->
                                    <Path Data="M8,0 L16,4 L12,16 L4,16 L0,4 Z" Fill="Gray" />
                                </Canvas>
                            </Viewbox>
                        </Border>
                        <TextBlock Style="{StaticResource WorkflowLabelStyle}" Text="MFA Login" />
                        <!--  Erklärung für Schritt 2  -->
                        <TextBlock
                            FontSize="9"
                            Foreground="Gray"
                            Text="  2. Login per MFA    "
                            TextAlignment="Center" />
                    </StackPanel>
                    <!--  Schritt 3: Source Verbindung  -->
                    <StackPanel>
                        <Border x:Name="workflowIcon3" Style="{StaticResource WorkflowIconContainer}">
                            <Viewbox>
                                <Canvas Width="16" Height="16">
                                    <!--  Kettenglied  -->
                                    <Path
                                        Data="M4,4 A2,2 0 1,1 4,12 A2,2 0 1,1 4,4 M12,4 A2,2 0 1,1 12,12 A2,2 0 1,1 12,4"
                                        Fill="Transparent"
                                        Stroke="Gray"
                                        StrokeThickness="2" />
                                    <Line
                                        Stroke="Gray"
                                        StrokeThickness="2"
                                        X1="4"
                                        X2="12"
                                        Y1="8"
                                        Y2="8" />
                                </Canvas>
                            </Viewbox>
                        </Border>
                        <TextBlock Style="{StaticResource WorkflowLabelStyle}" Text="Source" />
                        <!--  Erklärung für Schritt 3  -->
                        <TextBlock
                            FontSize="9"
                            Foreground="Gray"
                            Text="   3. Quell-Site abrufen   "
                            TextAlignment="Center" />
                    </StackPanel>
                    <!--  Schritt 4: Site Script exportieren  -->
                    <StackPanel>
                        <Border x:Name="workflowIcon4" Style="{StaticResource WorkflowIconContainer}">
                            <Viewbox>
                                <Canvas Width="16" Height="16">
                                    <!--  Pfeil nach oben  -->
                                    <Path
                                        Data="M8,16 L8,4 M8,4 L4,8 M8,4 L12,8"
                                        Stroke="Gray"
                                        StrokeEndLineCap="Round"
                                        StrokeLineJoin="Round"
                                        StrokeStartLineCap="Round"
                                        StrokeThickness="2" />
                                </Canvas>
                            </Viewbox>
                        </Border>
                        <TextBlock Style="{StaticResource WorkflowLabelStyle}" Text="Export" />
                        <!--  Erklärung für Schritt 4  -->
                        <TextBlock
                            FontSize="9"
                            Foreground="Gray"
                            Text="   4. Exportieren   "
                            TextAlignment="Center" />
                    </StackPanel>
                    <!--  Schritt 5: Site Design erstellen  -->
                    <StackPanel>
                        <Border x:Name="workflowIcon5" Style="{StaticResource WorkflowIconContainer}">
                            <Viewbox>
                                <Canvas Width="16" Height="16">
                                    <!--  Palette Icon (vereinfacht)  -->
                                    <Ellipse
                                        Width="16"
                                        Height="16"
                                        Fill="Transparent"
                                        Stroke="Gray"
                                        StrokeThickness="2" />
                                    <Ellipse
                                        Canvas.Left="2"
                                        Canvas.Top="2"
                                        Width="4"
                                        Height="4"
                                        Fill="Gray" />
                                </Canvas>
                            </Viewbox>
                        </Border>
                        <TextBlock Style="{StaticResource WorkflowLabelStyle}" Text="Design" />
                        <!--  Erklärung für Schritt 5  -->
                        <TextBlock
                            FontSize="9"
                            Foreground="Gray"
                            Text="   5. Site Design erstellen   "
                            TextAlignment="Center" />
                    </StackPanel>
                    <!--  Schritt 6: Design anwenden  -->
                    <StackPanel>
                        <Border x:Name="workflowIcon6" Style="{StaticResource WorkflowIconContainer}">
                            <Viewbox>
                                <Canvas Width="16" Height="16">
                                    <!--  Häkchen  -->
                                    <Path
                                        Data="M2,8 L6,12 L14,2"
                                        Fill="Transparent"
                                        Stroke="Gray"
                                        StrokeEndLineCap="Round"
                                        StrokeLineJoin="Round"
                                        StrokeStartLineCap="Round"
                                        StrokeThickness="2" />
                                </Canvas>
                            </Viewbox>
                        </Border>
                        <TextBlock Style="{StaticResource WorkflowLabelStyle}" Text="Anwenden" />
                        <!--  Erklärung für Schritt 6  -->
                        <TextBlock
                            FontSize="9"
                            Foreground="Gray"
                            Text="    6. Design anwenden   "
                            TextAlignment="Center" />
                    </StackPanel>
                </StackPanel>
            </GroupBox>

            <!--  Main Content Area - 3 Columns  -->
            <Grid Grid.Row="1">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="1*" />
                    <!--  Left Column: Form Fields  -->
                    <ColumnDefinition Width="1*" />
                    <!--  Middle Column: Configuration  -->
                    <ColumnDefinition Width="1*" />
                    <!--  Right Column: Status & Buttons  -->
                </Grid.ColumnDefinitions>

                <!--  Left Column: Form Fields  -->
                <StackPanel Grid.Column="0" Margin="0,0,45,0">
                    <TextBlock Style="{StaticResource GroupHeaderStyle}" Text="Policy -&gt; Module -&gt; Verbindung -&gt; " />

                    <!--  Execution Policy Group  -->
                    <GroupBox
                        Height="100"
                        Header="Execution Policy"
                        Style="{StaticResource ModernGroupBoxStyle}">
                        <StackPanel>
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="Auto" />
                                    <ColumnDefinition Width="*" />
                                </Grid.ColumnDefinitions>
                                <TextBlock
                                    Grid.Column="0"
                                    VerticalAlignment="Center"
                                    Style="{StaticResource LabelStyle}"
                                    Text="Aktuelle Policy:" />
                                <TextBlock
                                    x:Name="txtCurrentPolicy"
                                    Grid.Column="1"
                                    Margin="10,0,0,0"
                                    VerticalAlignment="Center"
                                    Text="{Binding Policy}" />
                            </Grid>
                            <CheckBox
                                x:Name="chkRemoteSigned"
                                Margin="0,10,0,0"
                                Content="RemoteSigned setzen"
                                Style="{StaticResource ModernCheckBoxStyle}" Width="163" />
                        </StackPanel>
                    </GroupBox>

                    <!--  Module Installation  -->
                    <GroupBox Header="PnP.PowerShell Modul" Style="{StaticResource ModernGroupBoxStyle}">
                        <StackPanel>
                            <TextBlock
                                x:Name="txtModuleStatus"
                                Style="{StaticResource StatusInfoTextBlockStyle}"
                                Text="STATUS ÜBERPRÜFEN..." />
                            <Button
                                x:Name="btnInstallModule"
                                Width="150"
                                Margin="0,10,0,0"
                                HorizontalAlignment="Right"
                                Content="Modul installieren"
                                Style="{StaticResource ModernButtonStyle}" />
                        </StackPanel>
                    </GroupBox>

                    <!--  Login with MFA  -->
                    <GroupBox
                        Height="101"
                        Header="Login mit MFA"
                        Style="{StaticResource ModernGroupBoxStyle}">
                        <StackPanel>
                            <TextBlock
                                x:Name="txtMFAStatus"
                                Style="{StaticResource StatusErrorTextBlockStyle}"
                                Text="Nicht angemeldet" />
                            <Button
                                x:Name="btnMFALogin"
                                Width="150"
                                Margin="0,10,0,0"
                                HorizontalAlignment="Right"
                                Content="Bei MS365 anmelden"
                                Style="{StaticResource ModernButtonStyle}" />
                        </StackPanel>
                    </GroupBox>

                    <!--  Site URLs  -->
                    <TextBlock
                        Margin="0,53,0,0"
                        Style="{StaticResource LabelStyle}"
                        Text="Admin Center URL (z.B. https://IhrTenant-admin.sharepoint.com)" />
                    <TextBox x:Name="txtAdminUrl" Style="{StaticResource RequiredTextBoxStyle}" />
                </StackPanel>

                <!--  Middle Column: Configuration  -->
                <StackPanel Grid.ColumnSpan="2" Margin="380,0,31,0">
                    <TextBlock Style="{StaticResource GroupHeaderStyle}" Text="Export -&gt; Einzel oder Mehrere Sites" />
                    <TextBlock Style="{StaticResource LabelStyle}" Text="Quell-Website (z.B. https://IhrTenant.sharepoint.com/sites/Quelle)" />
                    <TextBox x:Name="txtSourceSite" Style="{StaticResource RequiredTextBoxStyle}" />

                    <!--  Info Panel  -->
                    <Border
                        Height="55"
                        Margin="0,20,0,20"
                        Padding="10"
                        Background="#E3F2FD"
                        BorderBrush="#2196F3"
                        BorderThickness="1">
                        <TextBlock
                            Foreground="#0D47A1"
                            Text="HINWEIS: Dieses Tool exportiert nur Site-Design-Elemente (Navigation, Themes, Layouts), keine Inhalte, Listen oder Dokumente."
                            TextWrapping="Wrap" />
                    </Border>

                    <!--  Target Options  -->
                    <GroupBox
                        Height="175"
                        Header="Export/Transfer Optionen"
                        Style="{StaticResource ModernGroupBoxStyle}">
                        <StackPanel>
                            <RadioButton
                                x:Name="radioSingleSite"
                                Content="Export zu einer Ziel-Website"
                                IsChecked="True"
                                Style="{StaticResource ModernRadioButtonStyle}" />
                            <Grid Margin="20,5,0,0">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="Auto" />
                                    <ColumnDefinition Width="*" />
                                </Grid.ColumnDefinitions>
                                <TextBlock
                                    Grid.Column="0"
                                    VerticalAlignment="Center"
                                    Style="{StaticResource LabelStyle}"
                                    Text="Ziel-URL:" />
                                <TextBox
                                    x:Name="txtTargetSite"
                                    Grid.Column="1"
                                    Margin="10,0,0,0"
                                    Style="{StaticResource ModernTextBoxStyle}" />
                            </Grid>

                            <RadioButton
                                x:Name="radioMultiSite"
                                Margin="0,15,0,0"
                                Content="Export zu mehreren Websites mit Praefix"
                                Style="{StaticResource ModernRadioButtonStyle}" />
                            <Grid Margin="20,5,0,0">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="Auto" />
                                    <ColumnDefinition Width="*" />
                                </Grid.ColumnDefinitions>
                                <TextBlock
                                    Grid.Column="0"
                                    VerticalAlignment="Center"
                                    Style="{StaticResource LabelStyle}"
                                    Text="Präfix:" />
                                <TextBox
                                    x:Name="txtPrefix"
                                    Grid.Column="1"
                                    Margin="10,0,0,0"
                                    IsEnabled="False"
                                    Style="{StaticResource ModernTextBoxStyle}" />
                            </Grid>
                        </StackPanel>
                    </GroupBox>

                    <!--  Advanced Options  -->
                    <CheckBox
                        x:Name="chkAdvancedOptions"
                        Margin="0,10,0,0"
                        Content="Erweiterte Export-Optionen aktivieren"
                        Style="{StaticResource ModernCheckBoxStyle}" />

                    <GroupBox
                        x:Name="grpAdvancedOptions"
                        Height="70"
                        Header="Erweiterte Export-Optionen"
                        IsEnabled="False"
                        Style="{StaticResource ModernGroupBoxStyle}">
                        <WrapPanel>
                            <CheckBox
                                x:Name="chkIncludeBranding"
                                Margin="5"
                                Content="Branding"
                                IsChecked="True"
                                Style="{StaticResource ModernCheckBoxStyle}" />
                            <CheckBox
                                x:Name="chkIncludeTheme"
                                Margin="5"
                                Content="Theme"
                                IsChecked="True"
                                Style="{StaticResource ModernCheckBoxStyle}" />
                            <CheckBox
                                x:Name="chkIncludeNavigation"
                                Margin="5"
                                Content="Navigation"
                                IsChecked="True"
                                Style="{StaticResource ModernCheckBoxStyle}" />
                            <CheckBox
                                x:Name="chkIncludePages"
                                Margin="5"
                                Content="Seiten-Layouts"
                                IsChecked="True"
                                Style="{StaticResource ModernCheckBoxStyle}" />
                        </WrapPanel>
                    </GroupBox>

                    <!--  Site Design/Script Titles  -->
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*" />
                            <ColumnDefinition Width="*" />
                        </Grid.ColumnDefinitions>
                        <StackPanel Grid.Column="0" Grid.ColumnSpan="2">
                            <TextBlock
                                Margin="0,10,0,0"
                                Style="{StaticResource LabelStyle}"
                                Text="Titel für Site Design" />
                            <TextBox
                                x:Name="txtSiteDesign"
                                Style="{StaticResource ModernTextBoxStyle}"
                                ToolTip="Standard: ExportedDesign" />
                        </StackPanel>
                    </Grid>
                    <TextBlock Style="{StaticResource LabelStyle}" Text="Titel für Site Script" />
                    <TextBox
                        x:Name="txtSiteScript"
                        Style="{StaticResource ModernTextBoxStyle}"
                        ToolTip="Standard: ExportedScript" />
                </StackPanel>

                <!--  Right Column: Status & Buttons  -->
                <StackPanel
                    Grid.Column="1"
                    Grid.ColumnSpan="2"
                    Margin="394,0,0,0">
                    <TextBlock Style="{StaticResource GroupHeaderStyle}" Text="-&gt; Status der Ausführung" />

                    <!--  Status Panel  -->
                    <GroupBox
                        Height="270"
                        Header="Status Meldungen"
                        Style="{StaticResource ModernGroupBoxStyle}">
                        <StackPanel>
                            <ScrollViewer Height="150" VerticalScrollBarVisibility="Auto">
                                <TextBlock
                                    x:Name="txtStatus"
                                    Style="{StaticResource StatusInfoTextBlockStyle}"
                                    Text="Bereit"
                                    TextWrapping="Wrap" Height="140" />
                            </ScrollViewer>

                            <!-- Device Code Information Panel -->
                            <TextBlock 
                                x:Name="txtDeviceCodeInfo" 
                                Margin="0,10,0,5"
                                Text="Bitte besuchen Sie microsoft.com/devicelogin und geben Sie den folgenden Code ein:" 
                                Visibility="Collapsed"
                                Style="{StaticResource StatusInfoTextBlockStyle}" />

                            <Grid x:Name="deviceCodeGrid" Visibility="Collapsed">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*" />
                                    <ColumnDefinition Width="Auto" />
                                </Grid.ColumnDefinitions>

                                <Border 
                                    Grid.Column="0"
                                    Background="#E3F2FD" 
                                    BorderBrush="#2196F3" 
                                    BorderThickness="1" 
                                    Padding="10,5">
                                    <TextBlock 
                                        x:Name="txtDeviceCode" 
                                        FontSize="16"
                                        FontWeight="Bold"
                                        HorizontalAlignment="Center" />
                                </Border>

                                <Button 
                                    x:Name="btnCopyCode" 
                                    Grid.Column="1"
                                    Content="Kopieren" 
                                    Margin="5,0,0,0"
                                    Style="{StaticResource SecondaryButtonStyle}" />
                            </Grid>
                        </StackPanel>
                    </GroupBox>

                    <!--  Progress Bar  -->
                    <TextBlock
                        Margin="0,10,0,0"
                        Style="{StaticResource LabelStyle}"
                        Text="Fortschritt" />
                    <ProgressBar
                        x:Name="progressOperation"
                        Margin="0,0,0,20"
                        Style="{StaticResource ModernProgressBarStyle}"
                        Value="0" />

                    <!--  Operations Buttons  -->
                    <StackPanel Margin="0,10,0,0" HorizontalAlignment="Center">
                        <Button
                            x:Name="btnValidateInputs"
                            Width="150"
                            Margin="0,5"
                            HorizontalAlignment="Right"
                            Content="Eingaben pruefen"
                            Style="{StaticResource ModernButtonStyle}" />
                        <Button
                            x:Name="btnExportDesign"
                            Width="150"
                            Margin="0,5"
                            HorizontalAlignment="Right"
                            Content="Design exportieren"
                            Style="{StaticResource ModernButtonStyle}" />
                        <Button
                            x:Name="btnCreateScriptDesign"
                            Width="150"
                            Margin="0,5"
                            HorizontalAlignment="Right"
                            Content="Script/Design erstellen"
                            IsEnabled="False"
                            Style="{StaticResource ModernButtonStyle}" />
                        <Button
                            x:Name="btnApplyDesign"
                            Width="150"
                            Margin="0,5"
                            HorizontalAlignment="Right"
                            Content="Design anwenden"
                            IsEnabled="False"
                            Style="{StaticResource ModernButtonStyle}" />
                        <Button
                            x:Name="btnCancel"
                            Width="150"
                            Margin="0,15"
                            HorizontalAlignment="Right"
                            Content="Abbrechen"
                            Style="{StaticResource SecondaryButtonStyle}" />
                    </StackPanel>
                </StackPanel>
            </Grid>
        </Grid>

        <!--  Footer  -->
        <Border
            Grid.Row="2"
            Background="#F0F0F0"
            BorderBrush="#E0E0E0"
            BorderThickness="0,1,0,0">
            <Grid>
                <TextBlock Style="{StaticResource FooterTextBlockStyle}" Text="SharePoint Design Export Tool v0.1.5" />
                <TextBlock
                    Margin="0,0,20,0"
                    HorizontalAlignment="Right"
                    Style="{StaticResource FooterTextBlockStyle}"
                    Text="© 2025 psscripts.de">
                    <TextBlock.ToolTip>https://psscripts.de</TextBlock.ToolTip>
                </TextBlock>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

    # Die XAML in ein XML-Objekt konvertieren
    Write-Host "[DEBUG] Konvertiere XAML in XML-Objekt..." -ForegroundColor Cyan
    $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader] $xamlContent)
    $window = [Windows.Markup.XamlReader]::Load($reader)
    Write-Host "[DEBUG] XAML wurde erfolgreich in Window-Objekt konvertiert" -ForegroundColor Green
    
    # GUI-Elemente als Hashtable speichern
    $gui = @{}
    
    # Alle benannten Elemente im GUI finden und in Hashtable speichern
    Write-Host "[DEBUG] Suche benannte Elemente im GUI..." -ForegroundColor Cyan
    $window.FindName('nothing') # Dummy-Aufruf, um FindName zu aktivieren
    foreach ($elementName in ($xamlContent -split "`n" | Select-String 'x:Name="([^"]*)"' -AllMatches | 
                             ForEach-Object { $_.Matches } | ForEach-Object { $_.Groups[1].Value })) {
        try {
            $element = $window.FindName($elementName)
            if ($null -ne $element) {
                $gui[$elementName] = $element
                Write-Host "[DEBUG] Element gefunden: $elementName" -ForegroundColor Cyan
            }
        } catch {
            Write-Host "[DEBUG] Fehler beim Finden des Elements: $elementName - $($_.Exception.Message)" -ForegroundColor Red
            Write-Warning "Element not found: $elementName"
        }
    }
    
    Write-Host "[DEBUG] GUI-Elemente geladen: $($gui.Count) Elemente gefunden" -ForegroundColor Green
    
    # Fenster und GUI-Elemente zurückgeben
    return @{
        Window = $window
        GUI = $gui
    }
}

function Update-WorkflowIconStatus {
    param (
        [Parameter(Mandatory=$true)]
        [int]$StepNumber,
        
        [Parameter(Mandatory=$true)]
        [ValidateSet("Default", "Active", "Completed")]
        [string]$Status,
        
        [Parameter(Mandatory=$false)]
        [hashtable]$GuiElements,
        
        [Parameter(Mandatory=$false)]
        [System.Windows.Window]$Window
    )
    
    Write-Host "[DEBUG] Update-WorkflowIconStatus: Schritt $StepNumber auf Status '$Status' setzen" -ForegroundColor Cyan
    
    if ($null -eq $GuiElements) {
        Write-Host "[DEBUG] GUI-Elemente fehlen - Status kann nicht aktualisiert werden" -ForegroundColor Red
        return
    }
    
    $iconName = "workflowIcon$StepNumber"
    $icon = $GuiElements[$iconName]
    
    if ($null -ne $icon) {
        Write-Host "[DEBUG] Icon '$iconName' gefunden, ändere Style zu $Status" -ForegroundColor Cyan
        switch ($Status) {
            "Default" {
                $icon.Style = $Window.FindResource("WorkflowIconContainer")
                $pathElement = FindPathInBorder $icon
                if ($pathElement) {
                    if ($pathElement.Fill -is [System.Windows.Media.SolidColorBrush]) {
                        $pathElement.Fill = [System.Windows.Media.Brushes]::Gray
                    }
                    if ($pathElement.Stroke -is [System.Windows.Media.SolidColorBrush]) {
                        $pathElement.Stroke = [System.Windows.Media.Brushes]::Gray
                    }
                }
            }
            "Active" {
                $icon.Style = $Window.FindResource("WorkflowIconActive")
                $pathElement = FindPathInBorder $icon
                if ($pathElement) {
                    if ($pathElement.Fill -is [System.Windows.Media.SolidColorBrush]) {
                        $pathElement.Fill = [System.Windows.Media.Brushes]::Green
                    }
                    if ($pathElement.Stroke -is [System.Windows.Media.SolidColorBrush]) {
                        $pathElement.Stroke = [System.Windows.Media.Brushes]::Green
                    }
                }
            }
            "Completed" {
                $icon.Style = $Window.FindResource("WorkflowIconCompleted")
                $pathElement = FindPathInBorder $icon
                if ($pathElement) {
                    if ($pathElement.Fill -is [System.Windows.Media.SolidColorBrush]) {
                        $pathElement.Fill = [System.Windows.Media.Brushes]::Blue
                    }
                    if ($pathElement.Stroke -is [System.Windows.Media.SolidColorBrush]) {
                        $pathElement.Stroke = [System.Windows.Media.Brushes]::Blue
                    }
                }
            }
        }
        Write-Host "[DEBUG] Icon-Status wurde auf $Status gesetzt" -ForegroundColor Green
    } else {
        Write-Host "[DEBUG] Icon '$iconName' wurde nicht gefunden" -ForegroundColor Red
    }
}

function FindPathInBorder {
    param (
        [Parameter(Mandatory=$true)]
        [System.Windows.Controls.Border]$Border
    )
    
    Write-Host "[DEBUG] Suche Path-Element in Border" -ForegroundColor Cyan
    
    # Get the child of the border (should be a Viewbox)
    $viewbox = $Border.Child
    if ($viewbox -is [System.Windows.Controls.Viewbox]) {
        # Get the child of the Viewbox (should be a Canvas)
        $canvas = $viewbox.Child
        if ($canvas -is [System.Windows.Controls.Canvas]) {
            # Look for Path or other drawable elements in the Canvas children
            foreach ($child in $canvas.Children) {
                if ($child -is [System.Windows.Shapes.Path]) {
                    Write-Host "[DEBUG] Path-Element in Border gefunden" -ForegroundColor Green
                    return $child
                }
                elseif ($child -is [System.Windows.Shapes.Line]) {
                    Write-Host "[DEBUG] Line-Element in Border gefunden" -ForegroundColor Green
                    return $child
                }
                elseif ($child -is [System.Windows.Shapes.Ellipse]) {
                    Write-Host "[DEBUG] Ellipse-Element in Border gefunden" -ForegroundColor Green
                    return $child
                }
            }
        }
    }
    
    Write-Host "[DEBUG] Kein zeichenbares Element in Border gefunden" -ForegroundColor Yellow
    return $null
}

# Event handler für "Copy Code" Button
function Copy-DeviceCodeToClipboard {
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$GuiElements
    )
    
    Write-Host "[DEBUG] Kopiere Device Code in die Zwischenablage" -ForegroundColor Cyan
    
    if ($null -eq $GuiElements -or $null -eq $GuiElements.txtDeviceCode) {
        Write-Host "[DEBUG] GUI-Elemente oder txtDeviceCode fehlen" -ForegroundColor Red
        return
    }
    
    $code = $GuiElements.txtDeviceCode.Text
    
    if (-not [string]::IsNullOrEmpty($code) -and $code -ne "Wird generiert...") {
        # Code in die Zwischenablage kopieren
        Write-Host "[DEBUG] Code gefunden: $code - kopiere in Zwischenablage" -ForegroundColor Cyan
        $code | Set-Clipboard
        
        # Bestätigung anzeigen
        $GuiElements.txtStatus.Text = "Device Code wurde in die Zwischenablage kopiert. Bitte bei Microsoft.com/devicelogin eingeben."
        Write-Host "[DEBUG] Code wurde in die Zwischenablage kopiert" -ForegroundColor Green
    } else {
        Write-Host "[DEBUG] Kein Code zum Kopieren vorhanden oder Code wird noch generiert" -ForegroundColor Yellow
        $GuiElements.txtStatus.Text = "Device Code wird noch generiert oder ist nicht verfügbar."
    }
}

# Main function to create and show the GUI
function Show-SPODesignExportTool {
    Write-Host "[DEBUG] Starte SPO Design Export Tool..." -ForegroundColor Cyan
    
    # Load the GUI
    Write-Host "[DEBUG] Lade GUI..." -ForegroundColor Cyan
    $guiData = Load-XamlGui
    $window = $guiData.Window
    $gui = $guiData.GUI
    Write-Host "[DEBUG] GUI geladen, $($gui.Count) Elemente verfügbar" -ForegroundColor Green
    
    # Set initial values
    Write-Host "[DEBUG] Setze initiale Werte..." -ForegroundColor Cyan
    $gui.txtCurrentPolicy.Text = if ($epCurrent) { $epCurrent.ToString() } else { "Unbekannt" }
    $gui.txtModuleStatus.Text = if ($moduleInstalled) { "PnP.PowerShell Modul ist installiert" } else { "PnP.PowerShell Modul ist NICHT installiert" }
    if ($moduleInstalled) {
        $gui.txtModuleStatus.Style = $window.FindResource("StatusOkTextBlockStyle")
        $gui.btnInstallModule.IsEnabled = $false
        Write-Host "[DEBUG] Modul ist installiert, Installation-Button deaktiviert" -ForegroundColor Green
    } else {
        $gui.txtModuleStatus.Style = $window.FindResource("StatusErrorTextBlockStyle")
        $gui.btnInstallModule.IsEnabled = $true
        Write-Host "[DEBUG] Modul ist NICHT installiert, Installation-Button aktiviert" -ForegroundColor Yellow
    }
    
    # Event Handler for advanced options checkbox
    Write-Host "[DEBUG] Registriere Event-Handler..." -ForegroundColor Cyan
    $gui.chkAdvancedOptions.Add_Checked({
        Write-Host "[DEBUG] Erweiterte Optionen aktiviert" -ForegroundColor Cyan
        $gui.grpAdvancedOptions.IsEnabled = $true
    })
    
    $gui.chkAdvancedOptions.Add_Unchecked({
        Write-Host "[DEBUG] Erweiterte Optionen deaktiviert" -ForegroundColor Cyan
        $gui.grpAdvancedOptions.IsEnabled = $false
    })
    
    # Event handler for radio buttons
    $gui.radioSingleSite.Add_Checked({
        Write-Host "[DEBUG] Einzelne Site-Option ausgewählt" -ForegroundColor Cyan
        $gui.txtTargetSite.IsEnabled = $true
        $gui.txtPrefix.IsEnabled = $false
    })
    
    $gui.radioMultiSite.Add_Checked({
        Write-Host "[DEBUG] Multi-Site-Option ausgewählt" -ForegroundColor Cyan
        $gui.txtTargetSite.IsEnabled = $false
        $gui.txtPrefix.IsEnabled = $true
    })
    
    # Event handler for Install Module button
    $gui.btnInstallModule.Add_Click({
        Write-Host "[DEBUG] Installiere Modul-Button geklickt" -ForegroundColor Cyan
        
        # Workflow-Status aktualisieren
        Update-WorkflowIconStatus -StepNumber 1 -Status "Active" -GuiElements $gui -Window $window
        
        $gui.txtStatus.Text = "Installiere PnP.PowerShell Modul..."
        $gui.progressOperation.Value = 10
        
        try {
            Write-Host "[DEBUG] Starte Modul-Installation..." -ForegroundColor Cyan
            Install-Module -Name $moduleName -Scope CurrentUser -Force
            Write-Host "[DEBUG] Modul erfolgreich installiert" -ForegroundColor Green
            
            $gui.txtModuleStatus.Text = "PnP.PowerShell Modul wurde erfolgreich installiert"
            $gui.txtModuleStatus.Style = $window.FindResource("StatusOkTextBlockStyle")
            $gui.btnInstallModule.IsEnabled = $false
            $gui.txtStatus.Text = "PnP.PowerShell Modul wurde erfolgreich installiert. Bitte führen Sie nun den MFA-Login aus."
            $gui.progressOperation.Value = 20
            
            # Nach erfolgreicher Installation
            Update-WorkflowIconStatus -StepNumber 1 -Status "Completed" -GuiElements $gui -Window $window
        }
        catch {
            Write-Host "[DEBUG] Fehler bei der Modul-Installation: $($_.Exception.Message)" -ForegroundColor Red
            $gui.txtModuleStatus.Text = "Fehler bei der Installation: $_"
            $gui.txtStatus.Text = "Fehler bei der Installation des PnP.PowerShell Moduls: $_"
            $gui.txtModuleStatus.Style = $window.FindResource("StatusErrorTextBlockStyle")
            $gui.progressOperation.Value = 0
        }
    })
    
    # Event handler for "Copy Code" button
    if ($null -ne $gui.btnCopyCode) {
        $gui.btnCopyCode.Add_Click({
            Write-Host "[DEBUG] Copy Code Button geklickt" -ForegroundColor Cyan
            Copy-DeviceCodeToClipboard -GuiElements $gui
        })
    } else {
        Write-Host "[DEBUG] Element 'btnCopyCode' wurde nicht gefunden" -ForegroundColor Red
        Write-Warning "Element 'btnCopyCode' wurde nicht gefunden. Der Kopieren-Button wird nicht funktionieren."
    }
    
    # Event handler for MFA Login button - umgestellt auf Device Code Auth
    $gui.btnMFALogin.Add_Click({
        Write-Host "[DEBUG] MFA Login Button geklickt" -ForegroundColor Cyan
        
        # First check if admin URL is provided
        if ([string]::IsNullOrEmpty($gui.txtAdminUrl.Text)) {
            Write-Host "[DEBUG] Admin URL fehlt" -ForegroundColor Red
            $gui.txtStatus.Text = "Bitte geben Sie eine Admin Center URL ein (z.B. https://IhrTenant-admin.sharepoint.com)"
            $gui.txtStatus.Style = $window.FindResource("StatusErrorTextBlockStyle")
            return
        }
        
        # Update workflow status and UI
        Update-WorkflowIconStatus -StepNumber 2 -Status "Active" -GuiElements $gui -Window $window
        $gui.txtStatus.Text = "Starte Device Code Login-Vorgang..."
        $gui.progressOperation.Value = 25
        $gui.btnMFALogin.IsEnabled = $false
        Write-Host "[DEBUG] Starte Device Code Login für URL: $($gui.txtAdminUrl.Text)" -ForegroundColor Cyan
        
        try {
            # Import the module if it's installed
            if (Get-Module -ListAvailable -Name "PnP.PowerShell") {
                Write-Host "[DEBUG] Importiere PnP.PowerShell Modul" -ForegroundColor Cyan
                Import-Module -Name "PnP.PowerShell" -ErrorAction Stop
                
                # Vorbereitung der UI für Device Code-Anzeige
                $gui.txtDeviceCodeInfo.Visibility = "Visible"
                $gui.deviceCodeGrid.Visibility = "Visible"
                $gui.txtDeviceCode.Text = "Wird generiert..."
                $gui.txtStatus.Text = "Warte auf Device Code von Microsoft..."
                
                # Verbindung mit SharePoint Online herstellen
                Write-Host "[DEBUG] Verbinde zu SharePoint Online mit Device Code" -ForegroundColor Cyan
                
                try {
                    # Wir verwenden nur den DeviceLogin Parameter
                    $connection = Connect-PnPOnline -Url $gui.txtAdminUrl.Text -DeviceLogin -ReturnConnection -ErrorAction Stop
                    
                    # Nach erfolgreicher Verbindung
                    Write-Host "[DEBUG] Verbindung hergestellt. Prüfe auf erfolgreiche Authentifizierung..." -ForegroundColor Green
                    
                    if ($connection) {
                        # Verbindung erfolgreich
                        $gui.txtDeviceCode.Text = "[Authentifizierung erfolgreich]"
                        $gui.txtMFAStatus.Text = "Angemeldet"
                        $gui.txtMFAStatus.Style = $window.FindResource("StatusOkTextBlockStyle")
                        $gui.txtStatus.Text = "Device Code Login erfolgreich. Sie können nun den Design-Export starten."
                        $gui.progressOperation.Value = 40
                        
                        # Nach erfolgreicher Anmeldung
                        Update-WorkflowIconStatus -StepNumber 2 -Status "Completed" -GuiElements $gui -Window $window
                        
                        # Device Code UI ausblenden nach erfolgreichem Login
                        $gui.deviceCodeGrid.Visibility = "Collapsed"
                        $gui.txtDeviceCodeInfo.Visibility = "Collapsed"
                        
                        # Button wieder aktivieren
                        $gui.btnMFALogin.IsEnabled = $true
                    } else {
                        throw "Verbindung wurde nicht hergestellt"
                    }
                }
                catch {
                    Write-Host "[DEBUG] Fehler beim Verbinden mit SharePoint Online: $($_.Exception.Message)" -ForegroundColor Red
                    $gui.txtMFAStatus.Text = "Fehler bei der Anmeldung"
                    $gui.txtMFAStatus.Style = $window.FindResource("StatusErrorTextBlockStyle")
                    $gui.txtStatus.Text = "Fehler beim Device Code Login: $($_.Exception.Message)"
                    $gui.progressOperation.Value = 20
                    
                    # Device Code UI wieder ausblenden
                    $gui.deviceCodeGrid.Visibility = "Collapsed"
                    $gui.txtDeviceCodeInfo.Visibility = "Collapsed"
                    
                    # Button wieder aktivieren
                    $gui.btnMFALogin.IsEnabled = $true
                }
            }
            else {
                Write-Host "[DEBUG] PnP.PowerShell Modul ist nicht installiert" -ForegroundColor Red
                throw "PnP.PowerShell Modul ist nicht installiert. Bitte installieren Sie es zuerst."
            }
        }
        catch {
            Write-Host "[DEBUG] Fehler beim Device Code Login: $($_.Exception.Message)" -ForegroundColor Red
            $gui.txtMFAStatus.Text = "Fehler bei der Anmeldung"
            $gui.txtMFAStatus.Style = $window.FindResource("StatusErrorTextBlockStyle")
            $gui.txtStatus.Text = "Fehler beim Device Code Login: $($_.Exception.Message)"
            $gui.progressOperation.Value = 20
            
            # Device Code UI ausblenden bei Fehler
            $gui.deviceCodeGrid.Visibility = "Collapsed"
            $gui.txtDeviceCodeInfo.Visibility = "Collapsed"
            
            # Button wieder aktivieren
            $gui.btnMFALogin.IsEnabled = $true
        }
    })
    
    # Event handler for Export Design button
    $gui.btnExportDesign.Add_Click({
        Write-Host "[DEBUG] Export Design Button geklickt" -ForegroundColor Cyan
        
        # Validate inputs
        if ([string]::IsNullOrEmpty($gui.txtAdminUrl.Text) -or [string]::IsNullOrEmpty($gui.txtSourceSite.Text)) {
            Write-Host "[DEBUG] Pflichtfelder nicht ausgefüllt" -ForegroundColor Red
            $gui.txtStatus.Text = "Bitte füllen Sie alle erforderlichen Felder aus (Admin URL, Quell-Website)"
            $gui.txtStatus.Style = $window.FindResource("StatusErrorTextBlockStyle")
            return
        }
        
        # Workflow-Status für Schritte 3 und 4 aktualisieren
        Update-WorkflowIconStatus -StepNumber 3 -Status "Active" -GuiElements $gui -Window $window
        
        $gui.txtStatus.Text = "Verbinde zur Quell-Website..."
        $gui.progressOperation.Value = 45
        
        try {
            Write-Host "[DEBUG] Versuche Verbindung zur Quell-Website: $($gui.txtSourceSite.Text)" -ForegroundColor Cyan
            # Hier würden Sie die Verbindung zur Quell-Website herstellen
            # Connect-PnPOnline -Url $gui.txtSourceSite.Text
            
            # Simulierte Verbindung für dieses Beispiel
            Write-Host "[DEBUG] Simuliere Verbindung (1 Sekunde)" -ForegroundColor Cyan
            Start-Sleep -Seconds 1
            
            # Nach erfolgreicher Verbindung zur Quell-Website
            Write-Host "[DEBUG] Verbindung zur Quell-Website erfolgreich" -ForegroundColor Green
            Update-WorkflowIconStatus -StepNumber 3 -Status "Completed" -GuiElements $gui -Window $window
            Update-WorkflowIconStatus -StepNumber 4 -Status "Active" -GuiElements $gui -Window $window
            
            $gui.txtStatus.Text = "Exportiere Site Design..."
            $gui.progressOperation.Value = 60
            
            Write-Host "[DEBUG] Starte Export des Site Designs" -ForegroundColor Cyan
            # Hier würden Sie den Export durchführen
            # $siteDesign = Get-PnPSiteDesign ...
            
            # Simulierter Export für dieses Beispiel
            Write-Host "[DEBUG] Simuliere Export (2 Sekunden)" -ForegroundColor Cyan
            Start-Sleep -Seconds 2
            
            Write-Host "[DEBUG] Export erfolgreich abgeschlossen" -ForegroundColor Green
            $gui.txtStatus.Text = "Site Design wurde erfolgreich exportiert. Sie können nun das Script/Design erstellen."
            $gui.progressOperation.Value = 75
            $gui.btnCreateScriptDesign.IsEnabled = $true
            
            # Nach erfolgreichem Export
            Update-WorkflowIconStatus -StepNumber 4 -Status "Completed" -GuiElements $gui -Window $window
        }
        catch {
            Write-Host "[DEBUG] Fehler beim Exportieren des Site Designs: $($_.Exception.Message)" -ForegroundColor Red
            $gui.txtStatus.Text = "Fehler beim Exportieren des Site Designs: $_"
            $gui.txtStatus.Style = $window.FindResource("StatusErrorTextBlockStyle")
            $gui.progressOperation.Value = 40
        }
    })
    
    # Event handler for Create Script/Design button
    $gui.btnCreateScriptDesign.Add_Click({
        Write-Host "[DEBUG] Create Script/Design Button geklickt" -ForegroundColor Cyan
        
        # Workflow-Status für Schritt 5 aktualisieren
        Update-WorkflowIconStatus -StepNumber 5 -Status "Active" -GuiElements $gui -Window $window
        
        $gui.txtStatus.Text = "Erstelle Site Script und Site Design..."
        $gui.progressOperation.Value = 80
        
        try {
            Write-Host "[DEBUG] Starte Erstellung von Site Script und Site Design" -ForegroundColor Cyan
            # Hier würden Sie das Site Script und Design erstellen
            # Add-PnPSiteScript ...
            # Add-PnPSiteDesign ...
            
            # Simulierte Erstellung für dieses Beispiel
            Write-Host "[DEBUG] Simuliere Erstellung (2 Sekunden)" -ForegroundColor Cyan
            Start-Sleep -Seconds 2
            
            Write-Host "[DEBUG] Site Script und Site Design erfolgreich erstellt" -ForegroundColor Green
            $gui.txtStatus.Text = "Site Script und Site Design wurden erfolgreich erstellt. Sie können nun das Design anwenden."
            $gui.progressOperation.Value = 90
            $gui.btnApplyDesign.IsEnabled = $true
            
            # Nach erfolgreicher Erstellung
            Update-WorkflowIconStatus -StepNumber 5 -Status "Completed" -GuiElements $gui -Window $window
        }
        catch {
            Write-Host "[DEBUG] Fehler beim Erstellen des Site Scripts/Designs: $($_.Exception.Message)" -ForegroundColor Red
            $gui.txtStatus.Text = "Fehler beim Erstellen des Site Scripts/Designs: $_"
            $gui.txtStatus.Style = $window.FindResource("StatusErrorTextBlockStyle")
            $gui.progressOperation.Value = 75
        }
    })
    
    # Event handler for Apply Design button
    $gui.btnApplyDesign.Add_Click({
        Write-Host "[DEBUG] Apply Design Button geklickt" -ForegroundColor Cyan
        
        # Workflow-Status für Schritt 6 aktualisieren
        Update-WorkflowIconStatus -StepNumber 6 -Status "Active" -GuiElements $gui -Window $window
        
        $gui.txtStatus.Text = "Wende Site Design an..."
        $gui.progressOperation.Value = 95
        
        try {
            Write-Host "[DEBUG] Starte Anwendung des Site Designs" -ForegroundColor Cyan
            # Hier würden Sie das Site Design anwenden
            # Invoke-PnPSiteDesign ...
            
            # Simulierte Anwendung für dieses Beispiel
            Write-Host "[DEBUG] Simuliere Anwendung (2 Sekunden)" -ForegroundColor Cyan
            Start-Sleep -Seconds 2
            
            Write-Host "[DEBUG] Site Design erfolgreich angewendet" -ForegroundColor Green
            $gui.txtStatus.Text = "Site Design wurde erfolgreich angewendet. Vorgang abgeschlossen."
            $gui.progressOperation.Value = 100
            
            # Nach erfolgreicher Anwendung
            Update-WorkflowIconStatus -StepNumber 6 -Status "Completed" -GuiElements $gui -Window $window
        }
        catch {
            Write-Host "[DEBUG] Fehler beim Anwenden des Site Designs: $($_.Exception.Message)" -ForegroundColor Red
            $gui.txtStatus.Text = "Fehler beim Anwenden des Site Designs: $_"
            $gui.txtStatus.Style = $window.FindResource("StatusErrorTextBlockStyle")
            $gui.progressOperation.Value = 90
        }
    })
    
    # Event handler for Cancel button
    $gui.btnCancel.Add_Click({
        Write-Host "[DEBUG] Cancel Button geklickt - Beende Anwendung" -ForegroundColor Cyan
        $window.Close()
    })
    
    # Event handler for Validate Inputs button
    $gui.btnValidateInputs.Add_Click({
        Write-Host "[DEBUG] Validate Inputs Button geklickt" -ForegroundColor Cyan
        
        if ([string]::IsNullOrEmpty($gui.txtAdminUrl.Text) -or [string]::IsNullOrEmpty($gui.txtSourceSite.Text)) {
            Write-Host "[DEBUG] Validierung fehlgeschlagen - Pflichtfelder fehlen" -ForegroundColor Red
            $gui.txtStatus.Text = "Validation fehlgeschlagen: Bitte füllen Sie alle erforderlichen Felder aus (Admin URL, Quell-Website)"
            $gui.txtStatus.Style = $window.FindResource("StatusErrorTextBlockStyle")
        } else {
            Write-Host "[DEBUG] Validierung erfolgreich" -ForegroundColor Green
            $gui.txtStatus.Text = "Validation erfolgreich. Alle erforderlichen Felder wurden ausgefüllt."
            $gui.txtStatus.Style = $window.FindResource("StatusOkTextBlockStyle")
        }
    })
    
    # Show the window
    Write-Host "[DEBUG] Zeige Anwendungsfenster" -ForegroundColor Cyan
    $window.ShowDialog() | Out-Null
    Write-Host "[DEBUG] Anwendung wurde beendet" -ForegroundColor Cyan
}

# Start the application
Write-Host "[DEBUG] Starte SPO Design Export Tool - Hauptanwendung" -ForegroundColor Cyan
Show-SPODesignExportTool
Write-Host "[DEBUG] Hauptanwendung beendet" -ForegroundColor Cyan
