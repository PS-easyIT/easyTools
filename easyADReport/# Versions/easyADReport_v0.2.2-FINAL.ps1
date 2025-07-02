[xml]$Global:XAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        Title="easyADReport v0.2.2" Height="950" Width="1650"
        WindowStartupLocation="CenterScreen" ResizeMode="CanResizeWithGrip"
        Background="#F9F9F9" FontFamily="Segoe UI">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="60"/>
            <!-- Header -->
            <RowDefinition Height="*"/>
            <!-- Content Area -->
            <RowDefinition Height="50"/>
            <!-- Footer -->
        </Grid.RowDefinitions>

        <!-- Header (Grid.Row="0") -->
        <Border Grid.Row="0" Background="#0078D7" BorderThickness="0">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel Grid.Column="0" Orientation="Horizontal" Margin="20,0">
                    <TextBlock Text="easyADReport" Foreground="White" FontSize="22" FontWeight="SemiBold" VerticalAlignment="Center"/>
                    <TextBlock Text="v0.2.2" Foreground="#E1E1E1" FontSize="14" VerticalAlignment="Center" Margin="10,0,0,0"/>
                </StackPanel>
                <Border Grid.Column="1" Background="#0063B1" Width="300" Height="50" CornerRadius="4" Margin="0,5,20,5">
                    <Grid x:Name="ResultCountGrid" Margin="5">
                        <Border Background="#0063B1" CornerRadius="2" Margin="2">
                            <Grid>
                                <TextBlock Text="SUMMARY" Foreground="White" FontSize="14" HorizontalAlignment="Center" VerticalAlignment="Top" Margin="0,5,0,0" Width="266"/>
                                <TextBlock x:Name="TotalResultCountText" Text="0" Foreground="White" FontSize="20" FontWeight="Bold" HorizontalAlignment="Center" VerticalAlignment="Top" Margin="0,10,0,0" Width="33" TextAlignment="Right"/>
                            </Grid>
                        </Border>
                    </Grid>
                </Border>
            </Grid>
        </Border>

        <!-- Content Area (Grid.Row="1") -->
        <Grid Grid.Row="1">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="350"/>
                <!-- Quick Reports -->
                <ColumnDefinition Width="*"/>
                <!-- Rest (Filter/Attributes + Results) -->
            </Grid.ColumnDefinitions>

            <!-- Quick Reports Panel (linke Spalte) -->
            <ScrollViewer Grid.Column="0" VerticalScrollBarVisibility="Auto" Margin="20,15,10,15">
                <Border Background="White" CornerRadius="4" BorderBrush="#DDDDDD" BorderThickness="1">
                    <GroupBox Header="Quick Reports" BorderThickness="0" Margin="5,0,0,0">
                        <StackPanel>
                            <GroupBox Header="Users" BorderThickness="0" Margin="0,5,0,0">
                                <StackPanel>
                                    <Button x:Name="ButtonQuickAllUsers" Content="All Users" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <Button x:Name="ButtonQuickDisabledUsers" Content="Disabled Users" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <Button x:Name="ButtonQuickLockedUsers" Content="Locked Users" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <Button x:Name="ButtonQuickNeverExpire" Content="Password Never Expires" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <Button x:Name="ButtonQuickInactiveUsers" Content="Inactive Users (90+ days)" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <Button x:Name="ButtonQuickAdminUsers" Content="Administrators" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <Button x:Name="ButtonQuickRecentlyCreatedUsers" Content="Recently Created Users (30d)" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <Button x:Name="ButtonQuickPasswordExpiringSoon" Content="Password Expiring Soon (7d)" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <Button x:Name="ButtonQuickExpiredPasswords" Content="Expired Passwords" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <Button x:Name="ButtonQuickNeverLoggedOn" Content="Never Logged On Accounts" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <Button x:Name="ButtonQuickRecentlyDeletedUsers" Content="Recently Deleted Users" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <Button x:Name="ButtonQuickRecentlyModifiedUsers" Content="Recently Modified Users" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <Button x:Name="ButtonQuickInactiveUsersXDays" Content="Inactive Users (letzte X Tage)" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <Button x:Name="ButtonQuickUsersWithoutManager" Content="Users without Manager" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <Button x:Name="ButtonQuickUsersMissingRequiredAttributes" Content="Users missing Required Attributes" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <Button x:Name="ButtonQuickUsersDuplicateLogonNames" Content="Users with Duplicate Logon Names" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <Button x:Name="ButtonQuickOrphanedSIDsUsers" Content="Orphaned SIDs (Users)" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                </StackPanel>
                            </GroupBox>
                            <GroupBox Header="Groups" BorderThickness="0" Margin="0,5,0,0">
                                <StackPanel>
                                    <Button x:Name="ButtonQuickGroups" Content="All Groups" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <Button x:Name="ButtonQuickSecurityGroups" Content="Security Groups" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <Button x:Name="ButtonQuickDistributionGroups" Content="Distribution Lists" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <Button x:Name="ButtonQuickRecentlyCreatedGroups" Content="Recently Created Groups" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <Button x:Name="ButtonQuickRecentlyDeletedGroups" Content="Recently Deleted Groups" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <Button x:Name="ButtonQuickEmptyGroups" Content="Empty Groups" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <Button x:Name="ButtonQuickVeryLargeGroupsXMembers" Content="Very Large Groups (&gt; X Mitglieder)" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <Button x:Name="ButtonQuickDeeplyNestedGroupsXDepth" Content="Deeply Nested Groups (Tiefe &gt; X)" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <Button x:Name="ButtonQuickGroupsWithoutOwner" Content="Groups Without Owner" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <Button x:Name="ButtonQuickGroupsWithoutDescription" Content="Groups Without Description" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <Button x:Name="ButtonQuickDistListsWithoutMailbox" Content="Distribution Lists Without Mailbox" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <Button x:Name="ButtonQuickGroupNestingInfo" Content="Group Nesting Information" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                </StackPanel>
                            </GroupBox>
                            <GroupBox Header="Computers" BorderThickness="0" Margin="0,5,0,0">
                                <StackPanel>
                                    <Button x:Name="ButtonQuickComputers" Content="All Computers" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <Button x:Name="ButtonQuickInactiveComputers" Content="Inactive Computers (90+ days)" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <Button x:Name="ButtonQuickRecentlyCreatedComputers30d" Content="Recently Created Computers (&lt; 30 Tage)" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <Button x:Name="ButtonQuickInactiveComputersXDays" Content="Inactive Computers (&gt; X Tage)" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <Button x:Name="ButtonQuickComputersNotLoggedOnXDays" Content="Computers Not Logged On Within X Tage" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <Button x:Name="ButtonQuickDisabledComputerAccounts" Content="Disabled Computer Accounts" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <Button x:Name="ButtonQuickOutdatedOSComputers" Content="Computers with Outdated Operating System" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <Button x:Name="ButtonQuickDecommissionPendingComputers" Content="Decommission-Pending Computers" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <Button x:Name="ButtonQuickComputersWithoutDescription" Content="Computers Without Description" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                </StackPanel>
                            </GroupBox>
                            <GroupBox Header="Gruppenrichtlinien (GPOs)" BorderThickness="0" Margin="0,5,0,0">
                                <StackPanel>
                                    <Button x:Name="ButtonQuickAllGPOs" Content="All GPOs" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <Button x:Name="ButtonQuickRecentlyCreatedGPOs" Content="Recently Created GPOs" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <Button x:Name="ButtonQuickRecentlyModifiedGPOs" Content="Recently Modified GPOs" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <Button x:Name="ButtonQuickDeletedGPOs" Content="Deleted GPOs" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <Button x:Name="ButtonQuickUnlinkedGPOs" Content="Unlinked GPOs" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <Button x:Name="ButtonQuickGPOsMissingPermissions" Content="GPOs with Missing Permissions" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <Button x:Name="ButtonQuickGPOsNotApplied" Content="GPOs Not Applied" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <Button x:Name="ButtonQuickGPOsWithWMIFilters" Content="GPOs with WMI Filters" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <Button x:Name="ButtonQuickGPOsLinkedToEmptyOUs" Content="GPOs Linked to Empty OUs" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <Button x:Name="ButtonQuickGPOsLoopbackProcessing" Content="GPOs with Loopback Processing" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                </StackPanel>
                            </GroupBox>
                            <GroupBox Header="Password Compliance" BorderThickness="0" Margin="0,5,0,0">
                                <StackPanel>
                                    <!-- Die meisten Reports dieser Kategorie (Passwords Expiring Soon, Expired Passwords, Users with Never-Expire Flag) sind bereits unter "Users" vorhanden. -->
                                    <TextBlock Text="Reports zu Passwörtern (Ablauf, Richtlinien) befinden sich aktuell unter 'Users'." Margin="5,5,5,5" TextWrapping="Wrap" Foreground="#808080"/>
                                </StackPanel>
                            </GroupBox>
                            <GroupBox Header="Login Anomalies" BorderThickness="0" Margin="0,5,0,0">
                                <StackPanel>
                                    <Button x:Name="ButtonQuickMultipleFailedLogons" Content="Multiple Failed Logon Attempts" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <Button x:Name="ButtonQuickLogonsOutsideBusinessHours" Content="Logons Outside Business Hours" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                </StackPanel>
                            </GroupBox>
                            <GroupBox Header="Compliance &amp; Auditing" BorderThickness="0" Margin="0,5,0,0">
                                <StackPanel>
                                    <!-- 'Accounts with PasswordNeverExpires' ist bereits unter 'Users' als 'Password Never Expires' vorhanden. -->
                                    <Button x:Name="ButtonQuickAccountsViolatingPasswordPolicy" Content="Accounts Violating Password Policy" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <Button x:Name="ButtonQuickTombstoneObjectsNearExpiry" Content="Tombstone Objects Near Retention Expiry" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <Button x:Name="ButtonQuickDeletedObjectsInTombstone" Content="Deleted Objects in Tombstone Store" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <Button x:Name="ButtonQuickOrphanedSIDsCompliance" Content="Orphaned SIDs (Compliance)" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                </StackPanel>
                            </GroupBox>
                            <GroupBox Header="Security Audit" BorderThickness="0" Margin="0,5,0,0">
                                <StackPanel>
                                    <Button x:Name="ButtonQuickWeakPasswordPolicy" Content="Weak Password Policies" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <Button x:Name="ButtonQuickRiskyGroupMemberships" Content="Risky Memberships" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <Button x:Name="ButtonQuickPrivilegedAccounts" Content="Privileged Accounts" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                </StackPanel>
                            </GroupBox>
                            <GroupBox Header="AD-Health" FontSize="11" BorderThickness="0" Margin="0,5,0,0">
                                <StackPanel>
                                    <Button x:Name="ButtonQuickFSMORoles" Content="FSMO Role Holders" Margin="0,2" FontSize="10" Height="15" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <Button x:Name="ButtonQuickDCStatus" Content="Domain Controller Status" Margin="0,2" FontSize="10" Height="15" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <Button x:Name="ButtonQuickReplicationStatus" Content="Replication Status" Margin="0,2" FontSize="10" Height="15" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                </StackPanel>
                            </GroupBox>
                            <GroupBox Header="OU - Topology" FontSize="11" BorderThickness="0" Margin="0,5,0,0">
                                <StackPanel>
                                    <Button x:Name="ButtonQuickOUHierarchy" Content="OU Hierarchy" Margin="0,2" FontSize="10" Height="15" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <Button x:Name="ButtonQuickSitesSubnets" Content="Sites &amp; Subnets" Margin="0,2" FontSize="10" Height="15" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <Button x:Name="ButtonQuickAllOUs" Content="All OUs" Margin="0,2" FontSize="10" Height="15" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <Button x:Name="ButtonQuickEmptyOUs" Content="Empty OUs" Margin="0,2" FontSize="10" Height="15" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <Button x:Name="ButtonQuickRecentlyCreatedOUs" Content="Recently Created OUs" Margin="0,2" FontSize="10" Height="15" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <Button x:Name="ButtonQuickRecentlyDeletedOUs" Content="Recently Deleted OUs" Margin="0,2" FontSize="10" Height="15" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <Button x:Name="ButtonQuickUnusedOUs" Content="Unused OUs" Margin="0,2" FontSize="10" Height="15" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <Button x:Name="ButtonQuickOrphanedOUs" Content="Orphaned OUs" Margin="0,2" FontSize="10" Height="15" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <Button x:Name="ButtonQuickOUHierarchyDepthReport" Content="OU Hierarchy Depth Report" Margin="0,2" FontSize="10" Height="15" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <!-- Weitere Buttons für DNS Zones hier später einfügen -->
                                </StackPanel>
                            </GroupBox>
                            <GroupBox Header="Permissions &amp; ACLs" BorderThickness="0" Margin="0,5,0,0">
                                <StackPanel>
                                    <Button x:Name="ButtonQuickExplicitNonInheritedPermissions" Content="Objects with Explicit Non-Inherited Permissions" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <Button x:Name="ButtonQuickBrokenInheritanceObjects" Content="Objects with Broken Inheritance" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <Button x:Name="ButtonQuickHighPrivilegeACEs" Content="High-Privilege ACEs" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <Button x:Name="ButtonQuickExcessiveDelegatedRights" Content="Excessive Delegated Rights" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                    <Button x:Name="ButtonQuickObjectsAccessibleByEveryone" Content="Objects Accessible by Everyone or Authenticated Users" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                </StackPanel>
                            </GroupBox>
                        </StackPanel>
                    </GroupBox>
                </Border>
            </ScrollViewer>

            <!-- Rechte Spalte: Filter/Attribute oben, Ergebnisse unten -->
            <Grid Grid.Column="1" Margin="0,15,20,15">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <!-- Für Filter/Attribute -->
                    <RowDefinition Height="*"/>
                    <!-- Für Ergebnisse -->
                </Grid.RowDefinitions>

                <!-- Filter and Attribute Selection (oben rechts) -->
                <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,15">
                    <Border Background="White" CornerRadius="4" BorderBrush="#DDDDDD" BorderThickness="1" Margin="0,0,15,0" Width="700">
                        <GroupBox Header="Filter" Margin="5" BorderThickness="0">
                            <StackPanel Margin="0,0,-2,0">
                                <StackPanel Orientation="Horizontal" Margin="0,0,0,10">
                                    <Label Content="Objekttyp:" VerticalAlignment="Center" Width="80"/>
                                    <RadioButton x:Name="RadioButtonUser" Content="User" IsChecked="True" Margin="5,0" VerticalAlignment="Center" />
                                    <RadioButton x:Name="RadioButtonGroup" Content="Group" Margin="15,0" VerticalAlignment="Center" />
                                    <RadioButton x:Name="RadioButtonComputer" Content="Computer" Margin="5,0,10,0" VerticalAlignment="Center" />
                                    <RadioButton x:Name="RadioButtonGroupMemberships" Content="Memberships" Margin="5,0" VerticalAlignment="Center" />
                                </StackPanel>
                                <StackPanel Orientation="Horizontal" Width="678">
                                    <Label Content="Filter Attribute:" VerticalAlignment="Center" Width="80"/>
                                    <ComboBox x:Name="ComboBoxFilterAttribute" Width="212" Margin="5,0" VerticalAlignment="Center" BorderThickness="1" BorderBrush="#CCCCCC"/>
                                    <Label Content="Filter Value:" VerticalAlignment="Center"/>
                                    <TextBox x:Name="TextBoxFilterValue" Width="300" Margin="5,0" VerticalAlignment="Center" BorderThickness="1" BorderBrush="#CCCCCC"/>
                                </StackPanel>
                            </StackPanel>
                        </GroupBox>
                    </Border>
                    <Border Background="White" CornerRadius="4" BorderBrush="#DDDDDD" BorderThickness="1" Margin="0,0,15,0" Width="346">
                        <GroupBox Header="Attributes to Export" Margin="5" BorderThickness="0">
                            <ListBox x:Name="ListBoxSelectAttributes" Height="80" SelectionMode="Multiple" BorderThickness="0" Margin="10,0,-2,0">
                                <ListBoxItem Content="DisplayName"/>
                                <ListBoxItem Content="SamAccountName"/>
                                <ListBoxItem Content="GivenName"/>
                                <ListBoxItem Content="Surname"/>
                                <ListBoxItem Content="mail"/>
                                <ListBoxItem Content="Department"/>
                                <ListBoxItem Content="Title"/>
                                <ListBoxItem Content="Enabled"/>
                            </ListBox>
                        </GroupBox>
                    </Border>
                    <StackPanel Orientation="Vertical" VerticalAlignment="Center" Height="90">
                        <Button x:Name="ButtonQueryAD" Content="S E A R C H" Width="166" Height="36" Margin="0,5,0,5" 
                    Background="#0078D7" Foreground="White" BorderThickness="0" FontWeight="Bold" FontSize="16">
                            <Button.Resources>
                                <Style TargetType="Border">
                                    <Setter Property="CornerRadius" Value="4"/>
                                </Style>
                            </Button.Resources>
                        </Button>
                        <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
                            <Button x:Name="ButtonExportCSV" Content="Export CSV" Width="80" Height="28" Background="#FFE8FFE8" BorderBrush="#CCCCCC" Margin="0,10,5,0" HorizontalAlignment="Left">
                                <Button.Resources>
                                    <Style TargetType="{x:Type Border}">
                                        <Setter Property="CornerRadius" Value="3"/>
                                    </Style>
                                </Button.Resources>
                            </Button>
                            <Button x:Name="ButtonExportHTML" Content="Export HTML" Width="80" Height="28" Background="#FFE8FFE8" BorderBrush="#CCCCCC" Margin="0,10,0,0" HorizontalAlignment="Right">
                                <Button.Resources>
                                    <Style TargetType="{x:Type Border}">
                                        <Setter Property="CornerRadius" Value="3"/>
                                    </Style>
                                </Button.Resources>
                            </Button>
                        </StackPanel>
                    </StackPanel>
                </StackPanel>

                <!-- Results (unten rechts) -->
                <Border Grid.Row="1" Background="White" CornerRadius="4" BorderBrush="#DDDDDD" BorderThickness="1">
                    <GroupBox BorderThickness="0" Margin="5">
                        <DataGrid x:Name="DataGridResults" AutoGenerateColumns="True" IsReadOnly="True" BorderThickness="0"
                                  Background="Transparent" GridLinesVisibility="All" RowBackground="White" AlternatingRowBackground="#F5F5F5"/>
                    </GroupBox>
                </Border>
            </Grid>
        </Grid>

        <!-- Footer (Grid.Row="2") -->
        <Border Grid.Row="2" Background="#F0F0F0" BorderThickness="0,1,0,0" BorderBrush="#DDDDDD">
            <Grid Margin="20,0">
                <TextBlock x:Name="TextBlockStatus" Text="Ready" VerticalAlignment="Center" Margin="0,0,460,0"/>
                <StackPanel Orientation="Horizontal" HorizontalAlignment="Right"/>
                <Button x:Name="ButtonClose" Content="EXIT" 
                            Background="#FFFF4242" Foreground="White" BorderThickness="0" FontWeight="Bold" FontSize="16" Margin="1474,5,0,8">
                    <Button.Resources>
                        <Style TargetType="{x:Type Border}">
                            <Setter Property="CornerRadius" Value="4"/>
                        </Style>
                    </Button.Resources>
                </Button>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

# Setze die Ausgabekodierung auf UTF-8, um Probleme mit Umlauten zu vermeiden
$Script:OutputEncoding = [Console]::InputEncoding = [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)

# Assembly für WPF laden
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms # Für SaveFileDialog

# DPI-Awareness für bessere Darstellung
Add-Type -TypeDefinition '
using System;
using System.Runtime.InteropServices;
public class DPIAware
{
    [DllImport("user32.dll")]
    public static extern bool SetProcessDPIAware();
}
' -ErrorAction SilentlyContinue

try {
    [void] [DPIAware]::SetProcessDPIAware()
} catch {
    Write-Warning "Could not set DPI awareness: $($_.Exception.Message)"
}

# --- Log-Funktion für konsistente Fehlerausgabe ---
Function Write-ADReportLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [ValidateSet('Info', 'Warning', 'Error')]
        [string]$Type = 'Info',
        
        [Parameter(Mandatory=$false)]
        [switch]$GUI,
        
        [Parameter(Mandatory=$false)]
        [switch]$Terminal
    )
    
    # Standardmäßig sowohl GUI als auch Terminal, wenn nicht explizit angegeben
    if (-not $GUI -and -not $Terminal) {
        $GUI = $true
        $Terminal = $true
    }
    
    # Ausgabe in der GUI
    if ($GUI -and $Global:TextBlockStatus) {
        $Global:TextBlockStatus.Text = $Message
    }
    
    # Ausgabe im Terminal
    if ($Terminal) {
        switch ($Type) {
            'Info'    { Write-Host $Message }
            'Warning' { Write-Warning $Message }
            'Error'   { Write-Error $Message }
        }
    }
}

# --- Funktion zum Abrufen von AD-Daten ---
Function Get-ADReportData {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false)]
        [string]$FilterAttribute,
        [Parameter(Mandatory=$false)]
        [string]$FilterValue,
        [Parameter(Mandatory=$true)]
        [System.Collections.IList]$SelectedAttributes,
        [Parameter(Mandatory=$false)]
        [string]$CustomFilter,
        [Parameter(Mandatory=$false)]
        [string]$ObjectType = "User"
    )

    # Überprüfen, ob das AD-Modul verfügbar ist
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        Write-ADReportLog -Message "Error: Active Directory module not found. Please install RSAT or AD PowerShell module." -Type Error
        return $null
    }

    # AD-Modul laden, falls nicht bereits geladen
    try {
        if (-not (Get-Module -Name ActiveDirectory)) {
            Import-Module ActiveDirectory -ErrorAction Stop
        }
    } catch {
        Write-ADReportLog -Message "Error: Could not load Active Directory module: $($_.Exception.Message)" -Type Error
        return $null
    }

    try {
        # Konvertiere SelectedAttributes zu String-Array und filtere leere/null Werte
        $PropertiesToLoad = @(
            $SelectedAttributes | ForEach-Object {
                if ($null -ne $_) {
                    # Wenn es sich um ListBoxItem handelt, Content verwenden
                    if ($_ -is [System.Windows.Controls.ListBoxItem]) {
                        $_.Content.ToString()
                    } else {
                        $_.ToString()
                    }
                }
            } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        )

        # Entferne doppelte Einträge aus der Properties-Liste
        $PropertiesToLoad = $PropertiesToLoad | Select-Object -Unique

        # Sicherstellen, dass immer einige Basiseigenschaften geladen werden
        if ('DistinguishedName' -notin $PropertiesToLoad) { # Wichtig für viele Operationen
            $PropertiesToLoad += 'DistinguishedName'
        }

        if ('ObjectClass' -notin $PropertiesToLoad) { # Wichtig für Visualisierung und Typbestimmung
            $PropertiesToLoad += 'ObjectClass'
        }

        # Sicherheitsfilter: Stelle sicher, dass 'memberOf' und 'Member' nur bei GroupMemberships-Abfragen geladen werden bzw. wenn explizit ausgewählt
        if ($ObjectType -ne "GroupMemberships") {
            $PropertiesToLoadBeforeFilter = @($PropertiesToLoad)
            $PropertiesToLoad = $PropertiesToLoad | Where-Object { $_ -notin @('memberOf', 'Member') }
            if (($PropertiesToLoadBeforeFilter.Count -ne $PropertiesToLoad.Count) -or ($PropertiesToLoadBeforeFilter -join ',') -ne ($PropertiesToLoad -join ',')) {
                 Write-ADReportLog -Message "Filtered 'memberOf' or 'Member' from PropertiesToLoad for $ObjectType query. Was: $($PropertiesToLoadBeforeFilter -join ', '), IsNow: $($PropertiesToLoad -join ', ')" -Type Info
            }
        }
        if ($PropertiesToLoad.Count -eq 0) {
            # Default properties if nothing was selected
            $PropertiesToLoad = @("DisplayName", "SamAccountName", "ObjectClass") 
        }

        $Filter = "*"
        if ($CustomFilter) {
            $Filter = $CustomFilter
        } elseif ($FilterAttribute -and $FilterValue) {
            $Filter = "$($FilterAttribute) -like '$($FilterValue)*'"
        } elseif ($FilterValue -and (-not $FilterAttribute)) {
            $Filter = "DisplayName -like '$($FilterValue)*'"
        }
        
        $Users = $null # Initialize Users

        # Special handling for LockedOut accounts
        if ($CustomFilter -and $CustomFilter.Trim() -eq "LockedOut -eq `$true") {
            Write-Host "Führe Search-ADAccount -LockedOut -UsersOnly aus, um gesperrte Benutzer zu finden."
            # Ensure AD module is loaded; this is checked at the beginning of the function.
            $LockedOutAccounts = Search-ADAccount -LockedOut -UsersOnly -ErrorAction Stop 
            
            if ($LockedOutAccounts) {
                Write-Host "$($LockedOutAccounts.Count) gesperrte(s) Konto/Konten gefunden. Rufe Details für ausgewählte Attribute ab..."
                
                $Users = foreach ($Account in $LockedOutAccounts) {
                    try {
                        Get-ADUser -Identity $Account.SamAccountName -Properties $PropertiesToLoad -ErrorAction SilentlyContinue | Select-Object -ExcludeProperty 'PropertyNames','AddedProperties','RemovedProperties','ModifiedProperties','PropertiesCount'
                    } catch {
                        Write-Warning "Konnte Details für Benutzer $($Account.SamAccountName) nicht abrufen: $($_.Exception.Message)"
                        $null
                    }
                }
                # Filter out any null results from failed Get-ADUser calls
                $Users = $Users | Where-Object {$_ -ne $null}

            } else {
                Write-Host "Keine gesperrten Benutzerkonten gefunden via Search-ADAccount."
            }
        } else {
            # Standard AD User Abfrage für andere Filter
            Write-Host "Executing Get-ADUser with filter '$Filter' and properties '$($PropertiesToLoad -join ', ')'"
            # Sicherstellen, dass Get-ADUser immer ein Array zurückgibt, auch bei nur einem Ergebnis
            $Users = @(Get-ADUser -Filter $Filter -Properties $PropertiesToLoad -ErrorAction Stop | Select-Object -ExcludeProperty 'PropertyNames','AddedProperties','RemovedProperties','ModifiedProperties','PropertiesCount')
            Write-ADReportLog -Message "Users found: $($Users.Count) - Type: $($Users.GetType().FullName)" -Type Info -Terminal
        }

        if ($Users) {
            Write-Host "DEBUG: Final Select. PropertiesToLoad: $($PropertiesToLoad -join ', ')"
            # Stelle sicher, dass nur explizit angeforderte und notwendige Standardeigenschaften zurückgegeben werden
            $Users = $Users | Select-Object -Property $PropertiesToLoad
            return $Users
        } else {
            Write-ADReportLog -Message "No users found for the specified filter." -Type Warning
            return $null
        }
    } catch {
        $ErrorMessage = "Error in AD query: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return $null
    }
}

# --- Funktion zum Abrufen der Gruppenmitgliedschaften eines Benutzers ---
Function Get-UserGroupMemberships {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]$SamAccountName
    )

    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        Write-ADReportLog -Message "Fehler: Active Directory Modul nicht gefunden." -Type Error
        return $null
    }

    try {
        $User = Get-ADUser -Identity $SamAccountName -Properties SamAccountName, Name -ErrorAction Stop | Select-Object -ExcludeProperty 'PropertyNames','AddedProperties','RemovedProperties','ModifiedProperties','PropertiesCount' # Hinzugefügt: Name für UserDisplayName
        if (-not $User) {
            Write-ADReportLog -Message "Benutzer $SamAccountName nicht gefunden." -Type Warning
            return $null
        }
        
        $Groups = Get-ADPrincipalGroupMembership -Identity $User -ErrorAction Stop | 
                  Get-ADGroup -Properties Name, SamAccountName, Description, GroupCategory, GroupScope -ErrorAction SilentlyContinue # Hinzugefügt: SamAccountName für GroupSamAccountName

        if ($Groups) {
            $GroupMemberships = $Groups | ForEach-Object {
                [PSCustomObject]@{
                    UserDisplayName = $User.Name
                    UserSamAccountName = $User.SamAccountName
                    GroupName = $_.Name
                    GroupSamAccountName = $_.SamAccountName
                    GroupDescription = $_.Description
                    GroupCategory = $_.GroupCategory
                    GroupScope = $_.GroupScope
                }
            }
            return $GroupMemberships
        } else {
            Write-ADReportLog -Message "Keine Gruppenmitgliedschaften für Benutzer $SamAccountName gefunden." -Type Info
            return [System.Collections.ArrayList]::new() # Leeres Array zurückgeben, um Fehler zu vermeiden
        }
    } catch {
        $ErrorMessage = "Fehler beim Abrufen der Gruppenmitgliedschaften fuer $($SamAccountName): $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return $null
    }
}

# --- Funktion zum Abrufen von AD-Gruppendaten ---
Function Get-ADGroupReportData {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [System.Collections.IList]$SelectedAttributes,
        [Parameter(Mandatory=$false)]
        [string]$CustomFilter
    )

    # Überprüfen, ob das AD-Modul verfügbar ist
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        Write-ADReportLog -Message "Fehler: Active Directory Modul nicht gefunden." -Type Error
        return $null
    }

    try {
        # Konvertiere SelectedAttributes zu String-Array und filtere leere/null Werte
        $PropertiesToLoad = @(
            $SelectedAttributes | ForEach-Object {
                if ($null -ne $_) {
                    # Wenn es sich um ListBoxItem handelt, Content verwenden
                    if ($_ -is [System.Windows.Controls.ListBoxItem]) {
                        $_.Content.ToString()
                    } else {
                        $_.ToString()
                    }
                }
            } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        )

        if ('ObjectClass' -notin $PropertiesToLoad) { # Wichtig für Visualisierung und Typbestimmung
            $PropertiesToLoad += 'ObjectClass'
        }

        if ($PropertiesToLoad.Count -eq 0) {
            $PropertiesToLoad = @("Name", "SamAccountName", "GroupCategory", "GroupScope", "ObjectClass") 
        }

        $FilterToUse = "*"
        if ($CustomFilter) {
            $FilterToUse = $CustomFilter
        }
        
        Write-Host "Führe Get-ADGroup mit Filter '$FilterToUse' und Eigenschaften '$($PropertiesToLoad -join ', ')' aus"
        $Groups = Get-ADGroup -Filter $FilterToUse -Properties $PropertiesToLoad -ErrorAction Stop

        if ($Groups) {
            # Erstelle Array mit den bereinigten Attributnamen für Select-Object
            $SelectAttributes = @(
                $SelectedAttributes | ForEach-Object {
                    if ($null -ne $_) {
                        if ($_ -is [System.Windows.Controls.ListBoxItem]) {
                            $_.Content.ToString()
                        } else {
                            $_.ToString()
                        }
                    }
                } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            )
            
            # Verwende Select-Object, um ein Array von Objekten zu erstellen
            # Dies stellt sicher, dass wir eine IEnumerable-Sammlung zurückgeben
            $Output = $Groups | Select-Object $SelectAttributes -ExcludeProperty 'PropertyNames','AddedProperties','RemovedProperties','ModifiedProperties','PropertiesCount'
            return $Output
        } else {
            Write-ADReportLog -Message "No groups found for the specified filter." -Type Warning
            return $null
        }
    } catch {
        $ErrorMessage = "Error in AD group query: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return $null
    }
}

# --- Funktion zum Abrufen von AD-Computerdaten ---
Function Get-ADComputerReportData {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [System.Collections.IList]$SelectedAttributes,
        [Parameter(Mandatory=$false)]
        [string]$CustomFilter
    )

    # Überprüfen, ob das AD-Modul verfügbar ist
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        Write-ADReportLog -Message "Fehler: Active Directory Modul nicht gefunden." -Type Error
        return $null
    }

    try {
        # Konvertiere SelectedAttributes zu String-Array und filtere leere/null Werte
        $PropertiesToLoad = @(
            $SelectedAttributes | ForEach-Object {
                if ($null -ne $_) {
                    # Wenn es sich um ListBoxItem handelt, Content verwenden
                    if ($_ -is [System.Windows.Controls.ListBoxItem]) {
                        $_.Content.ToString()
                    } else {
                        $_.ToString()
                    }
                }
            } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        )

        if ('ObjectClass' -notin $PropertiesToLoad) { # Wichtig für Visualisierung und Typbestimmung
            $PropertiesToLoad += 'ObjectClass'
        }

        if ($PropertiesToLoad.Count -eq 0) {
            $PropertiesToLoad = @("Name", "DNSHostName", "OperatingSystem", "Enabled", "ObjectClass") 
        }

        $FilterToUse = "*"
        if ($CustomFilter) {
            $FilterToUse = $CustomFilter
        }
        
        Write-Host "Führe Get-ADComputer mit Filter '$FilterToUse' und Eigenschaften '$($PropertiesToLoad -join ', ')' aus"
        $Computers = Get-ADComputer -Filter $FilterToUse -Properties $PropertiesToLoad -ErrorAction Stop

        if ($Computers) {
            # Erstelle Array mit den bereinigten Attributnamen für Select-Object
            $SelectAttributes = @(
                $SelectedAttributes | ForEach-Object {
                    if ($null -ne $_) {
                        if ($_ -is [System.Windows.Controls.ListBoxItem]) {
                            $_.Content.ToString()
                        } else {
                            $_.ToString()
                        }
                    }
                } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            )
            
            # Verwende Select-Object, um ein Array von Objekten zu erstellen
            $Output = $Computers | Select-Object $SelectAttributes -ExcludeProperty 'PropertyNames','AddedProperties','RemovedProperties','ModifiedProperties','PropertiesCount'
            return $Output
        } else {
            Write-ADReportLog -Message "No computers found for the specified filter." -Type Warning
            return $null
        }
    } catch {
        $ErrorMessage = "Error in AD computer query: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return $null
    }
}

# --- Funktion zum Abrufen von Gruppenmitgliedschaftsberichten ---
Function Get-ADGroupMembershipsReportData {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]$FilterAttribute,

        [Parameter(Mandatory=$true)]
        [string]$FilterValue
    )

    Write-ADReportLog -Message "Fetching group membership data for filter: $FilterAttribute = $FilterValue" -Type Info -Terminal
    $ReportOutput = @()

    try {
        $TargetObjects = Get-ADObject -Filter "$FilterAttribute -like '*$FilterValue*'" -Properties DisplayName, SamAccountName, MemberOf, Member, ObjectClass -ErrorAction SilentlyContinue

        if (-not $TargetObjects) {
            Write-ADReportLog -Message "No objects found for the specified filter '$FilterAttribute -like *$FilterValue*'." -Type Warning
            return $ReportOutput
        }

        foreach ($TargetObject in $TargetObjects) {
            # Korrekte Objektklasse abrufen
            $objectClassSimple = if ($TargetObject.ObjectClass -is [array]) {
                $TargetObject.ObjectClass[-1]
            } else {
                $TargetObject.ObjectClass
            }
            
            $objDisplayName = if ($TargetObject.DisplayName) { $TargetObject.DisplayName } else { $TargetObject.Name }
            $objSamAccountName = if ($TargetObject.SamAccountName) { $TargetObject.SamAccountName } else { "N/A" }

            Write-ADReportLog -Message "Processing object: $objDisplayName (Type: $objectClassSimple)" -Type Info

            if ($objectClassSimple -eq 'user' -or $objectClassSimple -eq 'computer') {
                if ($TargetObject.MemberOf) {
                    foreach ($groupDN in $TargetObject.MemberOf) {
                        try {
                            $groupObject = Get-ADGroup -Identity $groupDN -Properties DisplayName, SamAccountName -ErrorAction Stop
                            $groupDisplayName = if ($groupObject.DisplayName) { $groupObject.DisplayName } else { $groupObject.Name }
                            $ReportOutput += [PSCustomObject]@{ 
                                ObjectName = $objDisplayName
                                ObjectSAM  = $objSamAccountName
                                ObjectType = $objectClassSimple
                                Relationship = "Ist Mitglied von"
                                RelatedObject = $groupDisplayName
                                RelatedObjectSAM = $groupObject.SamAccountName
                                RelatedObjectType = "Group"
                            }
                        } catch {
                            Write-ADReportLog -Message "Error resolving group DN '$groupDN' for object '$objDisplayName': $($_.Exception.Message)" -Type Warning
                            $ReportOutput += [PSCustomObject]@{ 
                                ObjectName = $objDisplayName
                                ObjectSAM  = $objSamAccountName
                                ObjectType = $objectClassSimple
                                Relationship = "Ist Mitglied von (Fehler)"
                                RelatedObject = $groupDN 
                                RelatedObjectSAM = "N/A"
                                RelatedObjectType = "Group (Fehler)"
                            }
                        }
                    }
                } else {
                    $ReportOutput += [PSCustomObject]@{ 
                        ObjectName = $objDisplayName
                        ObjectSAM  = $objSamAccountName
                        ObjectType = $objectClassSimple
                        Relationship = "Ist Mitglied von"
                        RelatedObject = "(Keine Gruppenmitgliedschaften)"
                        RelatedObjectSAM = "N/A"
                        RelatedObjectType = "N/A"
                    }
                }
            } elseif ($objectClassSimple -eq 'group') {
                if ($TargetObject.Member) {
                    foreach ($memberDN in $TargetObject.Member) {
                        try {
                            $memberObject = Get-ADObject -Identity $memberDN -Properties DisplayName, SamAccountName, ObjectClass -ErrorAction Stop
                            $memberObjectClassSimple = if ($memberObject.ObjectClass -is [array]) {
                                $memberObject.ObjectClass[-1]
                            } else {
                                $memberObject.ObjectClass
                            }
                            $memberName = if ($memberObject.DisplayName) { $memberObject.DisplayName } else { $memberObject.Name }
                            $memberSam = if ($memberObject.SamAccountName) { $memberObject.SamAccountName } else { "N/A" }

                            $ReportOutput += [PSCustomObject]@{ 
                                ObjectName = $objDisplayName
                                ObjectSAM  = $objSamAccountName
                                ObjectType = $objectClassSimple
                                Relationship = "Hat Mitglied"
                                RelatedObject = $memberName
                                RelatedObjectSAM = $memberSam
                                RelatedObjectType = $memberObjectClassSimple
                            }
                        } catch {
                            Write-ADReportLog -Message "Error resolving member DN '$memberDN' for group '$objDisplayName': $($_.Exception.Message)" -Type Warning
                            $ReportOutput += [PSCustomObject]@{ 
                                ObjectName = $objDisplayName
                                ObjectSAM  = $objSamAccountName
                                ObjectType = $objectClassSimple
                                Relationship = "Hat Mitglied (Fehler)"
                                RelatedObject = $memberDN 
                                RelatedObjectSAM = "N/A"
                                RelatedObjectType = "Unbekannt (Fehler)"
                            }
                        }
                    }
                } else {
                    $ReportOutput += [PSCustomObject]@{ 
                        ObjectName = $objDisplayName
                        ObjectSAM  = $objSamAccountName
                        ObjectType = $objectClassSimple
                        Relationship = "Hat Mitglied"
                        RelatedObject = "(Keine Mitglieder)"
                        RelatedObjectSAM = "N/A"
                        RelatedObjectType = "N/A"
                    }
                }
            } else {
                Write-ADReportLog -Message "Object $objDisplayName is of unhandled type '$objectClassSimple' for membership report." -Type Info
            }
        }
    } catch {
        $ErrorMessage = "Error in Get-ADGroupMembershipsReportData: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
    }
    return $ReportOutput
}

# --- Security Audit Functions ---
Function Get-WeakPasswordPolicyUsers {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing users with weak password policies..." -Type Info -Terminal
        
        # Properties relevant for comprehensive password policy analysis
        $Properties = @(
            "DisplayName", "SamAccountName", "Enabled", "PasswordNeverExpires", 
            "PasswordNotRequired", "PasswordLastSet", "LastLogonDate", "AdminCount",
            "CannotChangePassword", "SmartcardLogonRequired", "TrustedForDelegation",
            "DoesNotRequirePreAuth", "UseDESKeyOnly", "AccountExpirationDate",
            "LastBadPasswordAttempt", "BadLogonCount", "LogonCount", "PrimaryGroup",
            "MemberOf", "ServicePrincipalNames", "UserAccountControl", "LockedOut",
            "TrustedToAuthForDelegation", "AllowReversiblePasswordEncryption",
            "whenCreated", "Description", "UserPrincipalName", "DistinguishedName"
        )
        
        # Retrieve Domain Password Policy for comparisons
        $DomainPasswordPolicy = Get-ADDefaultDomainPasswordPolicy
        $MinPasswordAge = $DomainPasswordPolicy.MinPasswordAge.Days
        $MaxPasswordAge = $DomainPasswordPolicy.MaxPasswordAge.Days
        $MinPasswordLength = $DomainPasswordPolicy.MinPasswordLength
        
        Write-ADReportLog -Message "Domain Password Policy - Min Age: $MinPasswordAge days, Max Age: $MaxPasswordAge days, Min Length: $MinPasswordLength chars" -Type Info -Terminal
        
        # Load all users with relevant properties
        $AllUsers = Get-ADUser -Filter * -Properties $Properties | Select-Object -ExcludeProperty 'PropertyNames','AddedProperties','RemovedProperties','ModifiedProperties','PropertiesCount'
        Write-ADReportLog -Message "$($AllUsers.Count) users loaded for analysis..." -Type Info -Terminal
        
        # Define high-risk service account patterns
        $ServiceAccountPatterns = @("svc", "service", "app", "sql", "iis", "web", "backup", "sync", "admin", "sa")
        $TestAccountPatterns = @("test", "temp", "demo", "guest", "anonymous", "trial")
        
        # Enhanced analysis for weak password policies
        $WeakPasswordUsers = foreach ($user in $AllUsers) {
            $issues = @()
            $riskLevel = 0
            $recommendations = @()
            $securityFlags = @()
            
            # 1. Password never expires
            if ($user.PasswordNeverExpires -eq $true) {
                $issues += "Password never expires"
                $riskLevel += 3
                $recommendations += "Enable password expiration"
                $securityFlags += "NO_EXPIRY"
            }
            
            # 2. Password not required
            if ($user.PasswordNotRequired -eq $true) {
                $issues += "Password not required"
                $riskLevel += 5  # Critical risk
                $recommendations += "Enforce password requirement"
                $securityFlags += "NO_PASSWORD_REQ"
            }
            
            # 3. No password set or extremely old password
            if ($user.PasswordLastSet -eq $null) {
                $issues += "No password set"
                $riskLevel += 5
                $recommendations += "Set password immediately"
                $securityFlags += "NO_PASSWORD"
            } elseif ($user.PasswordLastSet -lt (Get-Date).AddDays(-($MaxPasswordAge * 3))) {
                $issues += "Password extremely outdated (>$($MaxPasswordAge * 3) days)"
                $riskLevel += 4
                $recommendations += "Force password reset immediately"
                $securityFlags += "ANCIENT_PASSWORD"
            } elseif ($user.PasswordLastSet -lt (Get-Date).AddDays(-($MaxPasswordAge * 2))) {
                $issues += "Password very old (>$($MaxPasswordAge * 2) days)"
                $riskLevel += 3
                $recommendations += "Schedule password reset"
                $securityFlags += "OLD_PASSWORD"
            } elseif ($user.PasswordLastSet -lt (Get-Date).AddDays(-365)) {
                $issues += "Password older than 1 year"
                $riskLevel += 2
                $recommendations += "Password reset recommended"
                $securityFlags += "STALE_PASSWORD"
            }
            
            # 4. User cannot change password
            if ($user.CannotChangePassword -eq $true) {
                $issues += "Cannot change password"
                $riskLevel += 2
                $recommendations += "Allow password changes (except for service accounts)"
                $securityFlags += "NO_CHANGE_ALLOWED"
            }
            
            # 5. Kerberos Pre-Authentication disabled (ASREPRoast attack possible)
            if ($user.DoesNotRequirePreAuth -eq $true) {
                $issues += "Kerberos Pre-Auth disabled (ASREPRoast vulnerability)"
                $riskLevel += 4
                $recommendations += "Enable Kerberos Pre-Authentication"
                $securityFlags += "ASREP_ROASTABLE"
            }
            
            # 6. Weak encryption (DES)
            if ($user.UseDESKeyOnly -eq $true) {
                $issues += "Uses weak DES encryption"
                $riskLevel += 3
                $recommendations += "Disable DES encryption"
                $securityFlags += "WEAK_ENCRYPTION"
            }
            
            # 7. Reversible password encryption
            if ($user.AllowReversiblePasswordEncryption -eq $true) {
                $issues += "Reversible password encryption enabled"
                $riskLevel += 4
                $recommendations += "Disable reversible password encryption"
                $securityFlags += "REVERSIBLE_ENCRYPTION"
            }
            
            # 8. Smartcard authentication not used for privileged accounts
            if ($user.AdminCount -eq 1 -and $user.SmartcardLogonRequired -eq $false) {
                $issues += "Privileged account without smartcard requirement"
                $riskLevel += 3
                $recommendations += "Enable smartcard authentication for admin accounts"
                $securityFlags += "ADMIN_NO_SMARTCARD"
            }
            
            # 9. Delegation for normal user accounts
            if (($user.TrustedForDelegation -eq $true -or $user.TrustedToAuthForDelegation -eq $true) -and $user.AdminCount -ne 1) {
                $issues += "Delegation enabled for standard user"
                $riskLevel += 3
                $recommendations += "Restrict delegation to service accounts only"
                $securityFlags += "UNEXPECTED_DELEGATION"
            }
            
            # 10. Excessive failed logon attempts
            if ($user.BadLogonCount -gt 10) {
                $issues += "Excessive failed logon attempts ($($user.BadLogonCount))"
                $riskLevel += 2
                $recommendations += "Investigate account for potential compromise"
                $securityFlags += "HIGH_FAILED_LOGONS"
            } elseif ($user.BadLogonCount -gt 5) {
                $issues += "Multiple failed logon attempts ($($user.BadLogonCount))"
                $riskLevel += 1
                $recommendations += "Monitor account activity"
                $securityFlags += "FAILED_LOGONS"
            }
            
            # 11. Service account without SPN (potentially misconfigured)
            $isServiceAccount = $false
            foreach ($pattern in $ServiceAccountPatterns) {
                if ($user.SamAccountName -like "*$pattern*" -or $user.DisplayName -like "*$pattern*") {
                    $isServiceAccount = $true
                    break
                }
            }
            
            if ($isServiceAccount) {
                if (-not $user.ServicePrincipalNames) {
                    $issues += "Service account without SPN"
                    $riskLevel += 1
                    $recommendations += "Configure SPN for service account"
                    $securityFlags += "SERVICE_NO_SPN"
                }
                
                # Service accounts should not be interactive
                if ($user.LastLogonDate -and $user.LastLogonDate -gt (Get-Date).AddDays(-30)) {
                    $issues += "Interactive logons detected for service account"
                    $riskLevel += 2
                    $recommendations += "Review service account usage"
                    $securityFlags += "SERVICE_INTERACTIVE"
                }
            }
            
            # 12. Never logged in accounts with high privileges
            if ($user.LogonCount -eq 0 -and $user.AdminCount -eq 1) {
                $issues += "Admin account never used"
                $riskLevel += 3
                $recommendations += "Disable unused admin account"
                $securityFlags += "UNUSED_ADMIN"
            } elseif ($user.LogonCount -eq 0 -and $user.whenCreated -lt (Get-Date).AddDays(-30)) {
                $issues += "Account never used (created >30 days ago)"
                $riskLevel += 1
                $recommendations += "Consider disabling unused account"
                $securityFlags += "NEVER_USED"
            }
            
            # 13. Inactive privileged accounts
            if ($user.AdminCount -eq 1 -and $user.LastLogonDate -and $user.LastLogonDate -lt (Get-Date).AddDays(-60)) {
                $issues += "Privileged account inactive for >60 days"
                $riskLevel += 3
                $recommendations += "Review inactive admin account"
                $securityFlags += "INACTIVE_ADMIN"
            } elseif ($user.LastLogonDate -and $user.LastLogonDate -lt (Get-Date).AddDays(-180)) {
                $issues += "Account inactive for >180 days"
                $riskLevel += 2
                $recommendations += "Consider disabling inactive account"
                $securityFlags += "LONG_INACTIVE"
            }
            
            # 14. Test/temp accounts without expiration
            $isTestAccount = $false
            foreach ($pattern in $TestAccountPatterns) {
                if ($user.SamAccountName -like "*$pattern*" -or $user.DisplayName -like "*$pattern*") {
                    $isTestAccount = $true
                    break
                }
            }
            
            if ($isTestAccount) {
                if ($user.AccountExpirationDate -eq $null) {
                    $issues += "Test/temp account without expiration date"
                    $riskLevel += 2
                    $recommendations += "Set expiration date for temporary accounts"
                    $securityFlags += "TEST_NO_EXPIRY"
                }
                
                if ($user.whenCreated -lt (Get-Date).AddDays(-90)) {
                    $issues += "Old test account (>90 days)"
                    $riskLevel += 1
                    $recommendations += "Review necessity of old test account"
                    $securityFlags += "OLD_TEST_ACCOUNT"
                }
            }
            
            # 15. Locked accounts with admin privileges
            if ($user.LockedOut -eq $true -and $user.AdminCount -eq 1) {
                $issues += "Locked privileged account"
                $riskLevel += 2
                $recommendations += "Investigate locked admin account"
                $securityFlags += "LOCKED_ADMIN"
            }
            
            # 16. Accounts with suspicious creation patterns
            if ($user.whenCreated -gt (Get-Date).AddDays(-7) -and $user.AdminCount -eq 1) {
                $issues += "Recently created admin account"
                $riskLevel += 2
                $recommendations += "Verify legitimacy of new admin account"
                $securityFlags += "NEW_ADMIN"
            }
            
            # 17. Accounts with generic or weak naming
            $weakNames = @("admin", "administrator", "user", "guest", "test", "temp", "service", "default")
            foreach ($weakName in $weakNames) {
                if ($user.SamAccountName -eq $weakName -or $user.SamAccountName -like "$weakName*") {
                    $issues += "Generic/predictable account name"
                    $riskLevel += 1
                    $recommendations += "Use non-predictable account names"
                    $securityFlags += "WEAK_NAMING"
                    break
                }
            }
            
            # Only return users with identified vulnerabilities
            if ($issues.Count -gt 0) {
                # Categorize risk level
                $riskCategory = switch ([int]$riskLevel) {
                    {$_ -ge 10} { [string]"Critical" }
                    {$_ -ge 7} { [string]"High" }
                    {$_ -ge 4} { [string]"Medium" }
                    {$_ -ge 2} { [string]"Low" }
                    default    { [string]"Minimal" }
                }
                

                # Add context information
                $contextInfo = @()
                if ($user.AdminCount -eq 1) { $contextInfo += "Privileged Account" }
                if ($user.Enabled -eq $false) { $contextInfo += "Disabled" }
                if ($user.LockedOut -eq $true) { $contextInfo += "Locked" }
                if ($user.ServicePrincipalNames) { $contextInfo += "Service Account" }
                if ($isServiceAccount) { $contextInfo += "Service Pattern" }
                if ($isTestAccount) { $contextInfo += "Test/Temp Account" }
                if ($user.SamAccountName -match "^(admin|administrator|root|sa|service|svc)") { $contextInfo += "System Account" }
                
                # Calculate compliance status
                $complianceIssues = 0
                if ($user.PasswordNeverExpires) { $complianceIssues++ }
                if ($user.PasswordNotRequired) { $complianceIssues++ }
                if ($user.DoesNotRequirePreAuth) { $complianceIssues++ }
                if ($user.UseDESKeyOnly) { $complianceIssues++ }
                if ($user.AllowReversiblePasswordEncryption) { $complianceIssues++ }
                
                $complianceStatus = if ($complianceIssues -eq 0) { "Compliant" } 
                                   elseif ($complianceIssues -le 2) { "Partially Compliant" } 
                                   else { "Non-Compliant" }
                
                # Calculate password age in days
                $passwordAge = if ($user.PasswordLastSet) { 
                    [math]::Round((New-TimeSpan -Start $user.PasswordLastSet -End (Get-Date)).TotalDays) 
                } else { "Never Set" }
                
                # Calculate account age
                $accountAge = [math]::Round((New-TimeSpan -Start $user.whenCreated -End (Get-Date)).TotalDays)
                
                # Determine urgency level
                $urgencyLevel = if ($user.AdminCount -eq 1 -and $riskLevel -ge 7) { "Immediate Action Required" }
                               elseif ($riskLevel -ge 10) { "Critical" }
                               elseif ($riskLevel -ge 7) { "Urgent" }
                               elseif ($riskLevel -ge 4) { "High Priority" }
                               elseif ($riskLevel -ge 2) { "Medium Priority" }
                               else { "Low Priority" }
                
                # Enhanced output object with comprehensive analysis
                [PSCustomObject]@{
                    DisplayName = $user.DisplayName
                    SamAccountName = $user.SamAccountName
                    UserPrincipalName = $user.UserPrincipalName
                    Enabled = $user.Enabled
                    LockedOut = $user.LockedOut
                    Context = if ($contextInfo) { $contextInfo -join ", " } else { "Standard User" }
                    PasswordLastSet = $user.PasswordLastSet
                    PasswordAge = $passwordAge
                    AccountCreated = $user.whenCreated
                    AccountAge = $accountAge
                    LastLogonDate = $user.LastLogonDate
                    LogonCount = $user.LogonCount
                    BadLogonCount = $user.BadLogonCount
                    LastBadPasswordAttempt = $user.LastBadPasswordAttempt
                    Vulnerabilities = $issues -join "; "
                    SecurityFlags = $securityFlags -join ", "
                    RiskLevel = $riskCategory
                    RiskScore = $riskLevel
                    ComplianceStatus = $complianceStatus
                    ComplianceIssues = $complianceIssues
                    UrgencyLevel = $urgencyLevel
                    Recommendations = $recommendations -join "; "
                    Description = $user.Description
                    LastAssessment = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    TotalIssuesFound = $issues.Count
                    RequiresImmediateAction = ($urgencyLevel -eq "Immediate Action Required" -or $urgencyLevel -eq "Critical")
                }
            }
        }
        
        # Enhanced statistics for logging
        $totalIssues = $WeakPasswordUsers.Count
        $criticalIssues = ($WeakPasswordUsers | Where-Object { $_.RiskLevel -eq "Critical" }).Count
        $highIssues = ($WeakPasswordUsers | Where-Object { $_.RiskLevel -eq "High" }).Count
        $mediumIssues = ($WeakPasswordUsers | Where-Object { $_.RiskLevel -eq "Medium" }).Count
        $adminIssues = ($WeakPasswordUsers | Where-Object { $_.Context -like "*Privileged Account*" }).Count
        $nonCompliant = ($WeakPasswordUsers | Where-Object { $_.ComplianceStatus -eq "Non-Compliant" }).Count
        $immediateAction = ($WeakPasswordUsers | Where-Object { $_.RequiresImmediateAction -eq $true }).Count
        $serviceAccountIssues = ($WeakPasswordUsers | Where-Object { $_.Context -like "*Service*" }).Count
        $testAccountIssues = ($WeakPasswordUsers | Where-Object { $_.Context -like "*Test*" }).Count
        
        Write-ADReportLog -Message "Password Policy Analysis completed:" -Type Info -Terminal
        Write-ADReportLog -Message "  Total: $totalIssues users with vulnerabilities" -Type Info -Terminal
        Write-ADReportLog -Message "  Risk Distribution - Critical: $criticalIssues, High: $highIssues, Medium: $mediumIssues" -Type Info -Terminal
        Write-ADReportLog -Message "  Privileged accounts affected: $adminIssues" -Type Info -Terminal
        Write-ADReportLog -Message "  Service accounts with issues: $serviceAccountIssues" -Type Info -Terminal
        Write-ADReportLog -Message "  Test/temp accounts with issues: $testAccountIssues" -Type Info -Terminal
        Write-ADReportLog -Message "  Compliance violations: $nonCompliant" -Type Info -Terminal
        Write-ADReportLog -Message "  Requiring immediate action: $immediateAction" -Type Info -Terminal
        
        return $WeakPasswordUsers
        
    } catch {
        $ErrorMessage = "Error analyzing password policies: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return $null
    }
}

Function Get-RiskyGroupMemberships {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analysiere riskante Gruppenmitgliedschaften..." -Type Info -Terminal
        
        # Definiere hochprivilegierte/riskante Gruppen
        $RiskyGroups = @(
            "Domain Admins", "Domänen-Admins",
            "Enterprise Admins", "Organisations-Admins", 
            "Schema Admins", "Schema-Admins",
            "Administrators", "Administratoren",
            "Account Operators", "Konten-Operatoren",
            "Server Operators", "Server-Operatoren",
            "Backup Operators", "Sicherungs-Operatoren",
            "Print Operators", "Druck-Operatoren",
            "Replicator", "Replikations-Operator",
            "Remote Desktop Users", "Remotedesktopbenutzer",
            "Power Users", "Hauptbenutzer"
        )
        
        $RiskyUsers = @()
        
        # Für jede riskante Gruppe die Mitglieder ermitteln
        foreach ($groupName in $RiskyGroups) {
            try {
                $group = Get-ADGroup -Filter "Name -eq '$groupName'" -ErrorAction SilentlyContinue
                if ($group) {
                    $members = Get-ADGroupMember -Identity $group.SamAccountName -ErrorAction SilentlyContinue | 
                              Get-ADObject -Properties DisplayName, SamAccountName, ObjectClass -ErrorAction SilentlyContinue | 
                              Where-Object { $_.objectClass -eq "user" }
                    
                    foreach ($member in $members) {
                        try {
                            $userDetails = Get-ADUser -Identity $member.SamAccountName -Properties DisplayName, Enabled, LastLogonDate, PasswordLastSet -ErrorAction SilentlyContinue | Select-Object -ExcludeProperty 'PropertyNames','AddedProperties','RemovedProperties','ModifiedProperties','PropertiesCount'
                            
                            if ($userDetails) {
                                # Erstelle Objekt mit Risikobewertung
                                $riskUser = [PSCustomObject]@{
                                    DisplayName = $userDetails.DisplayName
                                    SamAccountName = $userDetails.SamAccountName
                                    Enabled = $userDetails.Enabled
                                    LastLogonDate = $userDetails.LastLogonDate
                                    PasswordLastSet = $userDetails.PasswordLastSet
                                    RisikoGruppe = $group.Name
                                    Risikostufe = switch ($group.Name) {
                                        {$_ -match "Domain Admins|Domänen-Admins|Enterprise Admins|Organisations-Admins|Schema Admins|Schema-Admins"} { "Kritisch" }
                                        {$_ -match "Administrators|Administratoren"} { "Hoch" }
                                        {$_ -match "Account Operators|Server Operators|Backup Operators"} { "Mittel" }
                                        default { "Niedrig" }
                                    }
                                    Empfehlung = if ($userDetails.Enabled -eq $false) { 
                                        "Deactivate Account delete from group" 
                                    } elseif ($userDetails.LastLogonDate -and $userDetails.LastLogonDate -lt (Get-Date).AddDays(-90)) { 
                                        "Check inactive account" 
                                    } else { 
                                        "Regularly review permissions" 
                                    }
                                }
                                $RiskyUsers += $riskUser
                            }
                        } catch {
                            Write-Warning "Konnte Details für Benutzer $($member.SamAccountName) nicht abrufen: $($_.Exception.Message)"
                        }
                    }
                }
            } catch {
                Write-Warning "Konnte Gruppe '$groupName' nicht finden oder darauf zugreifen: $($_.Exception.Message)"
            }
        }
        
        # Entferne Duplikate (Benutzer können in mehreren riskanten Gruppen sein)
        $UniqueRiskyUsers = $RiskyUsers | Sort-Object SamAccountName -Unique
        
        Write-ADReportLog -Message "$($UniqueRiskyUsers.Count) Benutzer mit riskanten Gruppenmitgliedschaften gefunden." -Type Info -Terminal
        return $UniqueRiskyUsers
        
    } catch {
        $ErrorMessage = "Fehler beim Analysieren der Gruppenmitgliedschaften: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return $null
    }
}

Function Get-PrivilegedAccounts {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analysiere Konten mit erhöhten Rechten..." -Type Info -Terminal
        
        # Eigenschaften für privilegierte Konten
        $Properties = @(
            "DisplayName", "SamAccountName", "Enabled", "AdminCount", 
            "LastLogonDate", "PasswordLastSet", "PasswordNeverExpires",
            "ServicePrincipalNames", "TrustedForDelegation", "TrustedToAuthForDelegation"
        )
        
        # Alle Benutzer mit AdminCount = 1 (historisch privilegiert)
        $AdminCountUsers = Get-ADUser -Filter "AdminCount -eq 1" -Properties $Properties | Select-Object -ExcludeProperty 'PropertyNames','AddedProperties','RemovedProperties','ModifiedProperties','PropertiesCount'
        
        # Service-Konten (Konten mit SPNs)
        $ServiceAccounts = Get-ADUser -Filter "ServicePrincipalNames -like '*'" -Properties $Properties | Select-Object -ExcludeProperty 'PropertyNames','AddedProperties','RemovedProperties','ModifiedProperties','PropertiesCount'
        
        # Konten mit Delegierungsrechten
        $DelegationAccounts = Get-ADUser -Filter "TrustedForDelegation -eq `$true -or TrustedToAuthForDelegation -eq `$true" -Properties $Properties | Select-Object -ExcludeProperty 'PropertyNames','AddedProperties','RemovedProperties','ModifiedProperties','PropertiesCount'
        
        # Alle privilegierten Konten zusammenführen
        $AllPrivilegedAccounts = @()
        $AllPrivilegedAccounts += $AdminCountUsers
        $AllPrivilegedAccounts += $ServiceAccounts
        $AllPrivilegedAccounts += $DelegationAccounts
        
        # Duplikate entfernen und analysieren
        $UniquePrivilegedAccounts = $AllPrivilegedAccounts | Sort-Object SamAccountName -Unique | ForEach-Object {
            $account = $_
            
            # Risikofaktoren analysieren
            $riskFactors = @()
            $riskLevel = 0
            
            if ($account.AdminCount -eq 1) {
                $riskFactors += "AdminCount set"
                $riskLevel += 2
            }
            
            if ($account.ServicePrincipalNames) {
                $riskFactors += "Service-Account (SPN)"
                $riskLevel += 1
            }
            
            if ($account.TrustedForDelegation) {
                $riskFactors += "Delegation activated"
                $riskLevel += 2
            }
            
            if ($account.TrustedToAuthForDelegation) {
                $riskFactors += "Constrained Delegation"
                $riskLevel += 1
            }
            
            if ($account.PasswordNeverExpires) {
                $riskFactors += "Password never expires"
                $riskLevel += 1
            }
            
            if ($account.Enabled -eq $false) {
                $riskFactors += "Account disabled"
                $riskLevel += 3  # Deactivated privileged accounts are a high risk
            }

            if ($account.LastLogonDate -and $account.LastLogonDate -lt (Get-Date).AddDays(-90)) {
                $riskFactors += "Inactive (>90 days)"
                $riskLevel += 2
            }
            
            # Privilegien-Level bestimmen
            $privilegeLevel = "Standard"
            if ($account.AdminCount -eq 1 -and ($account.TrustedForDelegation -or $account.TrustedToAuthForDelegation)) {
                $privilegeLevel = "Critical"
            } elseif ($account.AdminCount -eq 1) {
                $privilegeLevel = "High"
            } elseif ($account.ServicePrincipalNames -or $account.TrustedForDelegation) {
                $privilegeLevel = "Medium"
            }
            
            # Risikostufe bestimmen
            $overallRisk = switch ($riskLevel) {
                {$_ -ge 5} { "Critical" }
                {$_ -ge 3} { "High" }
                {$_ -ge 2} { "Medium" }
                default { "Low" }
            }
            
            # Empfehlungen generieren
            $recommendations = @()
            if ($account.Enabled -eq $false -and $account.AdminCount -eq 1) {
                $recommendations += "Reset AdminCount for deactivated account"
            }
            if ($account.LastLogonDate -and $account.LastLogonDate -lt (Get-Date).AddDays(-90)) {
                $recommendations += "Check account usage"
            }
            if ($account.PasswordNeverExpires -and $account.AdminCount -eq 1) {
                $recommendations += "Enable password expiration"
            }
            if ($account.TrustedForDelegation) {
                $recommendations += "Review delegation rights"
            }
            
            [PSCustomObject]@{
                AdminCount = $account.AdminCount
                ServiceAccount = [bool]$account.ServicePrincipalNames
                DisplayName = $account.DisplayName
                SamAccountName = $account.SamAccountName
                Enabled = $account.Enabled
                LastLogonDate = $account.LastLogonDate
                PasswordLastSet = $account.PasswordLastSet
                PrivilegeLevel = $privilegeLevel
                RiskFactors = $riskFactors -join "; "
                Recommendations = if ($recommendations) { $recommendations -join "; " } else { "Regular review" }
                Delegation = $account.TrustedForDelegation -or $account.TrustedToAuthForDelegation
            }
        }
        
        Write-ADReportLog -Message "$($UniquePrivilegedAccounts.Count) Konten mit erhöhten Rechten gefunden." -Type Info -Terminal
        return $UniquePrivilegedAccounts
        
    } catch {
        $ErrorMessage = "Fehler beim Analysieren der privilegierten Konten: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return $null
    }
}

# --- AD-Health Funktionen ---
Function Get-FSMORoles {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Retrieving FSMO role holders..." -Type Info -Terminal
        
        $Forest = Get-ADForest
        $Domain = Get-ADDomain
        
        $FSMORoles = @()
        
        # Forest-weite FSMO-Rollen
        $FSMORoles += [PSCustomObject]@{
            Role = "Schema Master"
            Type = "Forest-wide"
            Server = $Forest.SchemaMaster
            Domain = $Forest.Name
            Status = if (Test-Connection -ComputerName $Forest.SchemaMaster -Count 1 -Quiet) { "Online" } else { "Offline" }
            Description = "Manages the Active Directory schema"
        }

        $FSMORoles += [PSCustomObject]@{
            Role = "Domain Naming Master"
            Type = "Forest-wide"
            Server = $Forest.DomainNamingMaster
            Domain = $Forest.Name
            Status = if (Test-Connection -ComputerName $Forest.DomainNamingMaster -Count 1 -Quiet) { "Online" } else { "Offline" }
            Description = "Manages adding and removing domains"
        }
        
        # Domänen-spezifische FSMO-Rollen
        $FSMORoles += [PSCustomObject]@{
            Role = "PDC Emulator"
            Type = "Domain-specific"
            Server = $Domain.PDCEmulator
            Domain = $Domain.Name
            Status = if (Test-Connection -ComputerName $Domain.PDCEmulator -Count 1 -Quiet) { "Online" } else { "Offline" }
            Description = "Time synchronization and password changes"
        }

        $FSMORoles += [PSCustomObject]@{
            Role = "RID Master"
            Type = "Domain-specific"
            Server = $Domain.RIDMaster
            Domain = $Domain.Name
            Status = if (Test-Connection -ComputerName $Domain.RIDMaster -Count 1 -Quiet) { "Online" } else { "Offline" }
            Description = "Distributes RID pools to domain controllers"
        }

        $FSMORoles += [PSCustomObject]@{
            Role = "Infrastructure Master"
            Type = "Domain-specific"
            Server = $Domain.InfrastructureMaster
            Domain = $Domain.Name
            Status = if (Test-Connection -ComputerName $Domain.InfrastructureMaster -Count 1 -Quiet) { "Online" } else { "Offline" }
            Description = "Manages cross-domain references"
        }
        
        Write-ADReportLog -Message "$($FSMORoles.Count) FSMO roles found." -Type Info -Terminal
        return $FSMORoles
        
    } catch {
        $ErrorMessage = "Error retrieving FSMO roles: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return $null
    }
}

Function Get-DomainControllerStatus {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Collect AD Data ..." -Type Info -Terminal
        
        # AD-Health Sammlung - ähnlich wie dxdiag
        $ADHealthReport = @()
        
        # 1. Forest-Informationen
        try {
            $Forest = Get-ADForest
            $ForestInfo = [PSCustomObject]@{
            Category = "Forest-Information"
            Parameter = "Forest Name"
            Value = $Forest.Name
            Status = "OK"
            Details = "Forest Functional Level: $($Forest.ForestMode)"
            }
            $ADHealthReport += $ForestInfo
            
            # Schema Version korrekt auslesen
            $schemaVersion = "Unknown"
            if ($Forest.SchemaVersion) {
            # Prüfe ob es eine Collection ist und hole den ersten Wert
            if ($Forest.SchemaVersion -is [Microsoft.ActiveDirectory.Management.ADPropertyValueCollection]) {
                $schemaVersion = $Forest.SchemaVersion[0]
            } else {
                $schemaVersion = $Forest.SchemaVersion.ToString()
            }
            }
            
            $ADHealthReport += [PSCustomObject]@{
            Category = "Forest-Information"
            Parameter = "Schema Version"
            Value = $schemaVersion
            Status = "OK"
            Details = "Current schema version"
            }
        } catch {
            $ADHealthReport += [PSCustomObject]@{
            Category = "Forest-Information"
            Parameter = "Forest Access"
            Value = "Error"
            Status = "Critical"
            Details = $_.Exception.Message
            }
        }
        
        # 2. Domain-Informationen
        try {
            $Domain = Get-ADDomain
            $ADHealthReport += [PSCustomObject]@{
                Category = "Domain-Information"
                Parameter = "Domain Name"
                Value = $Domain.NetBIOSName
                Status = "OK"
                Details = "FQDN: $($Domain.DNSRoot), Level: $($Domain.DomainMode)"
            }
            
            # PDC Emulator Test
            $pdcTest = Test-Connection -ComputerName $Domain.PDCEmulator -Count 1 -Quiet -ErrorAction SilentlyContinue
            $ADHealthReport += [PSCustomObject]@{
                Category = "Domain-Information"
                Parameter = "PDC Emulator"
                Value = $Domain.PDCEmulator
                Status = if ($pdcTest) { "OK" } else { "Warning" }
                Details = if ($pdcTest) { "Reachable" } else { "Not reachable" }
            }
        } catch {
            $ADHealthReport += [PSCustomObject]@{
                Category = "Domain-Information"
                Parameter = "Domain Access"
                Value = "Error"
                Status = "Critical"
                Details = $_.Exception.Message
            }
        }
        
        # 3. Domain Controller Diagnose
        try {
            $DomainControllers = Get-ADDomainController -Filter * -ErrorAction Stop
            
            # Konvertiere zu Array und zähle dann
            $DCArray = @($DomainControllers)
            $DCCount = $DCArray.Count
            $ADHealthReport += [PSCustomObject]@{
                Category = "Domain Controller"
                Parameter = "Numbers of DCs"
                Value = $DCCount
                Status = if ($DCCount -ge 2) { "OK" } else { "Warning" }
                Details = "Recommended: Min. 2 DCs for redundancy"
            }

            foreach ($DC in $DomainControllers) {
                # Ping-Test
                $pingResult = Test-Connection -ComputerName $DC.HostName -Count 1 -Quiet -ErrorAction SilentlyContinue
                $ADHealthReport += [PSCustomObject]@{
                    Category = "Domain Controller"
                    Parameter = "DC Reachability"
                    Value = $DC.Name
                    Status = if ($pingResult) { "OK" } else { "Critical" }
                    Details = if ($pingResult) { "Ping successful" } else { "Ping failed" }
                }

                # LDAP-Port Test
                $ldapTest = Test-NetConnection -ComputerName $DC.HostName -Port 389 -InformationLevel Quiet -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
                $ADHealthReport += [PSCustomObject]@{
                    Category = "Domain Controller"
                    Parameter = "LDAP Service"
                    Value = "$($DC.Name):389"
                    Status = if ($ldapTest) { "OK" } else { "Critical" }
                    Details = if ($ldapTest) { "LDAP Port reachable" } else { "LDAP Port not reachable" }
                }
                
                # Global Catalog Test
                if ($DC.IsGlobalCatalog) {
                    $gcTest = Test-NetConnection -ComputerName $DC.HostName -Port 3268 -InformationLevel Quiet -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
                    $ADHealthReport += [PSCustomObject]@{
                        Category = "Domain Controller"
                        Parameter = "Global Catalog"
                        Value = "$($DC.Name):3268"
                        Status = if ($gcTest) { "OK" } else { "Warning" }
                        Details = if ($gcTest) { "GC Port reachable" } else { "GC Port not reachable    " }
                    }
                }
            }
        } catch {
            $ADHealthReport += [PSCustomObject]@{
                Category = "Domain Controller"
                Parameter = "DC Enumeration"
                Value = "Error"
                Status = "Critical"
                Details = $_.Exception.Message
            }
        }
        
        # 4. FSMO-Rollen Diagnose
        try {
            $ADHealthReport += [PSCustomObject]@{
                Category = "FSMO-Rollen"
                Parameter = "Schema Master"
                Value = $Forest.SchemaMaster
                Status = if (Test-Connection -ComputerName $Forest.SchemaMaster -Count 1 -Quiet -ErrorAction SilentlyContinue) { "OK" } else { "Critical" }
                Details = "Forest-wide role"
            }

            $ADHealthReport += [PSCustomObject]@{
                Category = "FSMO-Rollen"
                Parameter = "Domain Naming Master"
                Value = $Forest.DomainNamingMaster
                Status = if (Test-Connection -ComputerName $Forest.DomainNamingMaster -Count 1 -Quiet -ErrorAction SilentlyContinue) { "OK" } else { "Critical" }
                Details = "Forest-wide role"
            }
            
            $ADHealthReport += [PSCustomObject]@{
                Category = "FSMO-Rollen"
                Parameter = "PDC Emulator"
                Value = $Domain.PDCEmulator
                Status = if (Test-Connection -ComputerName $Domain.PDCEmulator -Count 1 -Quiet -ErrorAction SilentlyContinue) { "OK" } else { "Critical" }
                Details = "Time synchronization, password changes"
            }

            $ADHealthReport += [PSCustomObject]@{
                Category = "FSMO-Rollen"
                Parameter = "RID Master"
                Value = $Domain.RIDMaster
                Status = if (Test-Connection -ComputerName $Domain.RIDMaster -Count 1 -Quiet -ErrorAction SilentlyContinue) { "OK" } else { "Critical" }
                Details = "RID pool distribution"
            }
            
            $ADHealthReport += [PSCustomObject]@{
                Category = "FSMO-Rollen"
                Parameter = "Infrastructure Master"
                Value = $Domain.InfrastructureMaster
                Status = if (Test-Connection -ComputerName $Domain.InfrastructureMaster -Count 1 -Quiet -ErrorAction SilentlyContinue) { "OK" } else { "Critical" }
                Details = "Cross-domain references"
            }
        } catch {
            $ADHealthReport += [PSCustomObject]@{
                Category = "FSMO-Rollen"
                Parameter = "FSMO Access"
                Value = "Error"
                Status = "Critical"
                Details = $_.Exception.Message
            }
        }
        
        # 5. DNS-Diagnose
        try {
            $dnsResult = Resolve-DnsName -Name $Domain.DNSRoot -Type A -ErrorAction SilentlyContinue
            $ADHealthReport += [PSCustomObject]@{
                Category = "DNS-Diagnose"
                Parameter = "Domain DNS Resolution"
                Value = $Domain.DNSRoot
                Status = if ($dnsResult) { "OK" } else { "Warning" }
                Details = if ($dnsResult) { "DNS resolution successful" } else { "DNS resolution failed" }
            }
            
            # SRV-Records testen
            $srvTest = Resolve-DnsName -Name "_ldap._tcp.$($Domain.DNSRoot)" -Type SRV -ErrorAction SilentlyContinue
            $ADHealthReport += [PSCustomObject]@{
                Category = "DNS-Diagnose"
                Parameter = "LDAP SRV Records"
                Value = "_ldap._tcp.$($Domain.DNSRoot)"
                Status = if ($srvTest) { "OK" } else { "Warning" }
                Details = if ($srvTest) { "$($srvTest.Count) SRV Records found" } else { "No SRV Records found" }
            }
        } catch {
            $ADHealthReport += [PSCustomObject]@{
                Category = "DNS-Diagnose"
                Parameter = "DNS Test"
                Value = "Error"
                Status = "Warning"
                Details = $_.Exception.Message
            }
        }
        
        # 6. Replikations-Schnelltest
        try {
            $replPartners = Get-ADReplicationPartnerMetadata -Target (Get-ADDomainController).HostName -Partition (Get-ADDomain).DistinguishedName -ErrorAction SilentlyContinue | Select-Object -First 3
            
            if ($replPartners) {
                $recentRepl = $replPartners | Where-Object { $_.LastReplicationSuccess -gt (Get-Date).AddHours(-24) }
                $ADHealthReport += [PSCustomObject]@{
                    Category = "Replication"
                    Parameter = "Last Replication"
                    Value = "$($recentRepl.Count)/$($replPartners.Count) Partner"
                    Status = if ($recentRepl.Count -eq $replPartners.Count) { "OK" } elseif ($recentRepl.Count -gt 0) { "Warning" } else { "Critical" }
                    Details = "Replication at last 24h"
                }
            } else {
                $ADHealthReport += [PSCustomObject]@{
                    Category = "Replication"
                    Parameter = "Replication Partners"
                    Value = "None found"
                    Status = "Warning"
                    Details = "No replication partners found"
                }
            }
        } catch {
            $ADHealthReport += [PSCustomObject]@{
                Category = "Replication"
                Parameter = "Replication Test"
                Value = "Error"
                Status = "Warning"
                Details = $_.Exception.Message
            }
        }
        
        # 7. System-Gesundheit
        try {
            $ADHealthReport += [PSCustomObject]@{
                Category = "System-Info"
                Parameter = "Active Server"
                Value = $env:COMPUTERNAME
                Status = "Info"
                Details = "Ausfuehrungskontext: $($env:USERDOMAIN)\$($env:USERNAME)"
            }

            $ADHealthReport += [PSCustomObject]@{
                Category = "System-Info"
                Parameter = "PowerShell Version"
                Value = $PSVersionTable.PSVersion.ToString()
                Status = "Info"
                Details = "AD-Modul available: $(if (Get-Module -ListAvailable -Name ActiveDirectory) { 'Ja' } else { 'Nein' })"
            }

            # Zeitzone und Zeit
            $ADHealthReport += [PSCustomObject]@{
                Category = "System-Info"
                Parameter = "Systemtime"
                Value = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                Status = "Info"
                Details = "Zeitzone: UTC$((Get-TimeZone).BaseUtcOffset.ToString('hh\:mm'))"
            }
        } catch {
            $ADHealthReport += [PSCustomObject]@{
                Category = "System-Info"
                Parameter = "System-Check"
                Value = "Teilweise Fehler"
                Status = "Info"
                Details = $_.Exception.Message
            }
        }
        
        # Zusammenfassung erstellen
        $kritisch = ($ADHealthReport | Where-Object { $_.Status -eq "Critical" }).Count
        $warnungen = ($ADHealthReport | Where-Object { $_.Status -eq "Warning" }).Count
        $ok = ($ADHealthReport | Where-Object { $_.Status -eq "OK" }).Count
        
        $summary = [PSCustomObject]@{
            Category = "=== SUMMARY ==="
            Parameter = "AD-Health Status"
            Value = if ($kritisch -eq 0 -and $warnungen -eq 0) { "Healthy" } elseif ($kritisch -eq 0) { "Minor Issues" } else { "Critical Issues" }
            Status = if ($kritisch -eq 0 -and $warnungen -eq 0) { "OK" } elseif ($kritisch -eq 0) { "Warning" } else { "Critical" }
            Details = "OK: $ok, Warnings: $warnungen, Critical: $kritisch"
        }
        
        # Zusammenfassung an den Anfang setzen
        $FinalReport = @($summary) + $ADHealthReport
        
        Write-ADReportLog -Message "AD-Health Diagnose abgeschlossen: $($FinalReport.Count) Checks durchgeführt." -Type Info -Terminal
        return $FinalReport
        
    } catch {
        $ErrorMessage = "Schwerwiegender Fehler bei AD-Health Diagnose: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return @([PSCustomObject]@{
            Category = "FEHLER"
            Parameter = "AD-Health Check"
            Value = "Fehlgeschlagen"
            Status = "Kritisch"
            Details = $ErrorMessage
        })
    }
}

Function Get-ReplicationStatus {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Checking AD replication status..." -Type Info -Terminal
        
        $ReplicationData = @()
        $DomainControllers = Get-ADDomainController -Filter *
        
        # 1. Collect replication failures for all DCs
        Write-ADReportLog -Message "Checking replication failures..." -Type Info -Terminal
        foreach ($DC in $DomainControllers) {
            try {
                $ReplicationFailures = Get-ADReplicationFailure -Target $DC.HostName -ErrorAction SilentlyContinue
                
                if ($ReplicationFailures) {
                    foreach ($failure in $ReplicationFailures) {
                        $ReplicationData += [PSCustomObject]@{
                            Category = "Replication Failures"
                            SourceDC = $DC.Name
                            TargetDC = $failure.Partner
                            Partition = $failure.PartitionDN
                            Status = "ERROR"
                            ErrorType = $failure.FailureType
                            ErrorCount = $failure.FailureCount
                            FirstFailure = $failure.FirstFailureTime
                            LastFailure = $failure.LastFailureTime
                            Details = $failure.LastError
                        }
                    }
                } else {
                    # No errors for this DC
                    $ReplicationData += [PSCustomObject]@{
                        Category = "Replication Status"
                        SourceDC = $DC.Name
                        TargetDC = "All Partners"
                        Partition = "All Partitions"
                        Status = "OK"
                        ErrorType = "None"
                        ErrorCount = 0
                        FirstFailure = $null
                        LastFailure = $null
                        Details = "No replication failures found"
                    }
                }
            } catch {
                $ReplicationData += [PSCustomObject]@{
                    Category = "System Error"
                    SourceDC = $DC.Name
                    TargetDC = "N/A"
                    Partition = "N/A"
                    Status = "CRITICAL"
                    ErrorType = "Query Error"
                    ErrorCount = 1
                    FirstFailure = Get-Date
                    LastFailure = Get-Date
                    Details = "Error retrieving replication data: $($_.Exception.Message)"
                }
            }
        }
        
        # 2. Collect replication partner metadata
        Write-ADReportLog -Message "Collecting replication partner metadata..." -Type Info -Terminal
        foreach ($DC in $DomainControllers) {
            try {
                $PartnerMetadata = Get-ADReplicationPartnerMetadata -Target $DC.HostName -ErrorAction SilentlyContinue
                
                if ($PartnerMetadata) {
                    foreach ($partner in $PartnerMetadata) {
                        $timeSinceLastSync = "Unknown"
                        $syncStatus = "Unknown"
                        
                        if ($partner.LastReplicationSuccess) {
                            $timeSince = (Get-Date) - $partner.LastReplicationSuccess
                            $timeSinceLastSync = "$([math]::Round($timeSince.TotalHours, 1)) hours"
                            
                            # Status based on time since last sync
                            if ($timeSince.TotalHours -lt 1) {
                                $syncStatus = "EXCELLENT"
                            } elseif ($timeSince.TotalHours -lt 6) {
                                $syncStatus = "GOOD"
                            } elseif ($timeSince.TotalHours -lt 24) {
                                $syncStatus = "WARNING"
                            } else {
                                $syncStatus = "CRITICAL"
                            }
                        }
                        
                        $ReplicationData += [PSCustomObject]@{
                            Category = "Partner Metadata"
                            SourceDC = $DC.Name
                            TargetDC = $partner.Partner -replace ".*CN=NTDS Settings,CN=([^,]+),.*", '$1'
                            Partition = $partner.Partition
                            Status = $syncStatus
                            ErrorType = if ($partner.ConsecutiveReplicationFailures -gt 0) { "Consecutive Failures" } else { "None" }
                            ErrorCount = $partner.ConsecutiveReplicationFailures
                            FirstFailure = $partner.LastReplicationAttempt
                            LastFailure = $partner.LastReplicationSuccess
                            Details = "Last sync: $timeSinceLastSync ago | USN: $($partner.LastReplicationSuccess)"
                        }
                    }
                }
            } catch {
                Write-Warning "Error retrieving partner metadata for $($DC.Name): $($_.Exception.Message)"
            }
        }
        
        # 3. Additional replication health checks
        Write-ADReportLog -Message "Performing extended replication diagnostics..." -Type Info -Terminal
        
        # DC-to-DC connectivity tests
        foreach ($SourceDC in $DomainControllers) {
            foreach ($TargetDC in $DomainControllers) {
                if ($SourceDC.Name -ne $TargetDC.Name) {
                    try {
                        # Test LDAP connectivity
                        $ldapTest = Test-NetConnection -ComputerName $TargetDC.HostName -Port 389 -InformationLevel Quiet -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
                        
                        # Test RPC connectivity (for replication)
                        $rpcTest = Test-NetConnection -ComputerName $TargetDC.HostName -Port 135 -InformationLevel Quiet -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
                        
                        $connectionStatus = "OK"
                        $connectionDetails = "LDAP: OK, RPC: OK"
                        
                        if (-not $ldapTest -and -not $rpcTest) {
                            $connectionStatus = "CRITICAL"
                            $connectionDetails = "LDAP: FAILED, RPC: FAILED"
                        } elseif (-not $ldapTest) {
                            $connectionStatus = "WARNING"
                            $connectionDetails = "LDAP: FAILED, RPC: OK"
                        } elseif (-not $rpcTest) {
                            $connectionStatus = "WARNING"
                            $connectionDetails = "LDAP: OK, RPC: FAILED"
                        }
                        
                        $ReplicationData += [PSCustomObject]@{
                            Category = "DC Connectivity"
                            SourceDC = $SourceDC.Name
                            TargetDC = $TargetDC.Name
                            Partition = "Network Connectivity"
                            Status = $connectionStatus
                            ErrorType = if ($connectionStatus -ne "OK") { "Network Issue" } else { "None" }
                            ErrorCount = if ($connectionStatus -eq "CRITICAL") { 2 } elseif ($connectionStatus -eq "WARNING") { 1 } else { 0 }
                            FirstFailure = if ($connectionStatus -ne "OK") { Get-Date } else { $null }
                            LastFailure = if ($connectionStatus -ne "OK") { Get-Date } else { $null }
                            Details = $connectionDetails
                        }
                    } catch {
                        Write-Warning "Error testing connectivity between $($SourceDC.Name) and $($TargetDC.Name)"
                    }
                }
            }
        }
        
        # 4. Create replication summary
        $totalErrors = ($ReplicationData | Where-Object { $_.Status -eq "CRITICAL" -or $_.Status -eq "ERROR" }).Count
        $warnings = ($ReplicationData | Where-Object { $_.Status -eq "WARNING" }).Count
        $okCount = ($ReplicationData | Where-Object { $_.Status -eq "OK" -or $_.Status -eq "GOOD" -or $_.Status -eq "EXCELLENT" }).Count
        
        # Add summary to the beginning
        $summary = [PSCustomObject]@{
            Category = "SUMMARY"
            SourceDC = "All DCs"
            TargetDC = "All Partners"
            Partition = "Overall Status"
            Status = if ($totalErrors -eq 0 -and $warnings -eq 0) { "HEALTHY" } elseif ($totalErrors -eq 0) { "MINOR ISSUES" } else { "CRITICAL ISSUES" }
            ErrorType = "Overview"
            ErrorCount = $totalErrors
            FirstFailure = $null
            LastFailure = $null
            Details = "OK/Good: $okCount | Warnings: $warnings | Critical/Error: $totalErrors | DCs: $($DomainControllers.Count)"
        }
        
        # Final report with summary
        $FinalReplicationData = @($summary) + $ReplicationData
        
        Write-ADReportLog -Message "Replication analysis completed: $($FinalReplicationData.Count) entries found." -Type Info -Terminal
        return $FinalReplicationData
        
    } catch {
        $ErrorMessage = "Critical error during replication analysis: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return @([PSCustomObject]@{
            Category = "SYSTEM ERROR"
            SourceDC = "Unknown"
            TargetDC = "Unknown"
            Partition = "N/A"
            Status = "CRITICAL"
            ErrorType = "System Error"
            ErrorCount = 1
            FirstFailure = Get-Date
            LastFailure = Get-Date
            Details = $ErrorMessage
        })
    }
}

# --- Funktion zum Abrufen von OU-Hierarchie-Berichten ---
Function Get-ADOUHierarchyReport {
    [CmdletBinding()]
    param()

    Write-ADReportLog -Message "Generating OU Hierarchy Report..." -Type Info -Terminal
    try {
        $AllOUs = Get-ADOrganizationalUnit -Filter * -Properties DistinguishedName, Name, CanonicalName, whenCreated, whenChanged, Description, ProtectedFromAccidentalDeletion -ErrorAction SilentlyContinue
        if (-not $AllOUs) {
            Write-ADReportLog -Message "No Organizational Units found in the domain." -Type Warning -Terminal
            return $null
        }

        $OUReport = foreach ($OU in $AllOUs) {
            $ParentPath = "Domain Root"
            try {
                # Versuche, den Parent DN zu bekommen und daraus den Namen abzuleiten
                $parentDN = $OU.DistinguishedName.Substring($OU.DistinguishedName.IndexOf(',') + 1)
                if ($parentDN -and $parentDN -ne $OU.DistinguishedName) {
                    $parentOUObj = Get-ADOrganizationalUnit -Identity $parentDN -Properties Name -ErrorAction SilentlyContinue
                    if ($parentOUObj) {
                        $ParentPath = $parentOUObj.Name
                    } else {
                        # Fallback, wenn es keine OU ist, sondern z.B. eine Domain Komponente
                        if ($parentDN -match '^DC='){
                            $ParentPath = $parentDN
                        } else {
                           $ParentPath = $parentDN # Zeige den DN des Parents, wenn Name nicht ermittelbar
                        }
                    }
                }
            } catch {
                # Fehler beim Ermitteln des Parent, Standardwert bleibt
                Write-ADReportLog -Message "Could not determine parent for $($OU.Name) (DN: $($OU.DistinguishedName)). Error: $($_.Exception.Message)" -Type Warning -GUI
            }
            
            $DirectChildrenCount = 0 # Initialisieren für den Fall eines Fehlers
            try {
                 $DirectChildrenCount = (Get-ADObject -Filter * -SearchBase $OU.DistinguishedName -SearchScope OneLevel -ErrorAction SilentlyContinue).Count
            } catch {
                Write-ADReportLog -Message "Could not determine direct children count for $($OU.Name). Error: $($_.Exception.Message)" -Type Warning -GUI
            }

            [PSCustomObject]@{
                Name                        = $OU.Name
                DistinguishedName           = $OU.DistinguishedName
                ParentContainer             = $ParentPath
                Description                 = $OU.Description
                ProtectedFromDeletion       = $OU.ProtectedFromAccidentalDeletion
                Created                     = $OU.whenCreated
                Modified                    = $OU.whenChanged
                CanonicalName               = $OU.CanonicalName
                DirectChildrenCount         = $DirectChildrenCount
            }
        } # Ende foreach

        Write-ADReportLog -Message "Successfully generated OU Hierarchy Report for $($OUReport.Count) OUs." -Type Info -GUI
        return $OUReport | Sort-Object DistinguishedName
    } catch {
        $ErrorMessage = "Error generating OU Hierarchy Report: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error -GUI
        return $null # Oder ein Fehlerobjekt für das DataGrid
    }
}


# --- Funktion zum Abrufen deaktivierter Benutzer (Quick Report) ---
Function Get-ADDisabledUsersReport {
    [CmdletBinding()]
    param()

    Write-ADReportLog -Message "Lade 'Deaktivierte Benutzer' Report..." -Type Info -GUI
    try {
        $Users = Get-ADUser -Filter {Enabled -eq $false} -Properties DisplayName, SamAccountName, GivenName, Surname, Enabled, LastLogonDate, WhenCreated, DistinguishedName -ErrorAction Stop
        if ($Users) {
            Write-ADReportLog -Message "$($Users.Count) deaktivierte Benutzer gefunden." -Type Info -GUI
            return $Users | Select-Object DisplayName, SamAccountName, GivenName, Surname, Enabled, LastLogonDate, WhenCreated, DistinguishedName
        } else {
            Write-ADReportLog -Message "Keine deaktivierten Benutzer gefunden." -Type Info -GUI
            return $null
        }
    } catch {
        Write-ADReportLog -Message "Fehler beim Abrufen deaktivierter Benutzer: $($_.Exception.Message)" -Type Error -GUI
        return $null
    }
}

# --- Funktion zum Abrufen von inaktiven Benutzern (Quick Report) ---
Function Get-ADReportInactiveUsers {
    [CmdletBinding()]
    param(
        [int]$InactiveDays = 90
    )

    Write-ADReportLog -Message "Lade 'Inaktive Benutzer ($($InactiveDays)+ Tage)' Report..." -Type Info -GUI
    try {
        $Properties = @(
            "DisplayName", "SamAccountName", "GivenName", "Surname", "Enabled", 
            "LastLogonDate", "PasswordLastSet", "DistinguishedName", "Created", 
            "Modified", "Description", "mail", "ObjectClass"
        )
        $timespan = New-TimeSpan -Days $InactiveDays
        $inactiveDate = (Get-Date).Subtract($timespan)
        
        # Benutzer, deren LastLogonDate älter als $inactiveDate ist oder nie angemeldet waren (LastLogonDate ist $null)
        $InactiveUsers = Get-ADUser -Filter { (Enabled -eq $true) -and ( (LastLogonDate -lt $inactiveDate) -or (LastLogonDate -eq $null) ) } -Properties $Properties -ErrorAction SilentlyContinue
        
        if ($InactiveUsers) {
            Write-ADReportLog -Message "$($InactiveUsers.Count) inaktive Benutzer (seit $($InactiveDays)+ Tagen) gefunden." -Type Info -GUI
            return $InactiveUsers | Select-Object $Properties
        } else {
            Write-ADReportLog -Message "Keine inaktiven Benutzer (seit $($InactiveDays)+ Tagen) gefunden." -Type Warning -GUI
            return @()
        }
    }
    catch {
        $ErrorMessage = "Fehler beim Laden des 'Inaktive Benutzer' Reports: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error -GUI
        return $null
    }
}

# --- Funktion zum Abrufen von Administrator-Benutzern (Quick Report) ---
Function Get-ADReportAdminUsers {
    [CmdletBinding()]
    param()

    Write-ADReportLog -Message "Lade 'Administratoren' Report..." -Type Info -GUI
    try {
        $Properties = @(
            "DisplayName", "SamAccountName", "GivenName", "Surname", "Enabled", 
            "LastLogonDate", "PasswordLastSet", "AdminCount", "DistinguishedName", 
            "Created", "Modified", "Description", "mail", "ObjectClass"
        )
        # Benutzer mit AdminCount > 0
        $AdminUsers = Get-ADUser -Filter 'AdminCount -gt 0 -and Enabled -eq $true' -Properties $Properties -ErrorAction SilentlyContinue
        
        if ($AdminUsers) {
            Write-ADReportLog -Message "$($AdminUsers.Count) Administrator-Konten gefunden." -Type Info -GUI
            return $AdminUsers | Select-Object $Properties
        } else {
            Write-ADReportLog -Message "Keine Administrator-Konten gefunden." -Type Warning -GUI
            return @()
        }
    }
    catch {
        $ErrorMessage = "Fehler beim Laden des 'Administratoren' Reports: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error -GUI
        return $null
    }
}

# --- Funktion zum Abrufen von kürzlich geänderten Benutzern (Quick Report) ---
Function Get-ADReportRecentlyModifiedUsers {
    [CmdletBinding()]
    param(
        [int]$Days = 7 # Standardmäßig die letzten 7 Tage
    )

    Write-ADReportLog -Message "Lade 'Kürzlich geänderte Benutzer (letzte $($Days) Tage)' Report..." -Type Info -GUI
    try {
        $Properties = @(
            "DisplayName", "SamAccountName", "GivenName", "Surname", "Enabled", 
            "LastLogonDate", "PasswordLastSet", "DistinguishedName", "Created", 
            "Modified", "Description", "mail", "ObjectClass", "whenChanged"
        )
        $targetDate = (Get-Date).AddDays(-$Days)
        
        $RecentlyModifiedUsers = Get-ADUser -Filter { (Modified -ge $targetDate) -and (Enabled -eq $true) } -Properties $Properties -ErrorAction SilentlyContinue
        
        if ($RecentlyModifiedUsers) {
            Write-ADReportLog -Message "$($RecentlyModifiedUsers.Count) kürzlich geänderte Benutzer (seit $($Days) Tagen) gefunden." -Type Info -GUI
            return $RecentlyModifiedUsers | Select-Object $Properties
        } else {
            Write-ADReportLog -Message "Keine kürzlich geänderten Benutzer (seit $($Days) Tagen) gefunden." -Type Warning -GUI
            return @()
        }
    }
    catch {
        $ErrorMessage = "Fehler beim Laden des 'Kürzlich geänderte Benutzer' Reports: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error -GUI
        return $null
    }
}

# --- Funktion zum Abrufen von Benutzern mit fehlerhaften Kennworteingaben (Quick Report) ---
Function Get-ADReportUsersWBadPwdCount {
    [CmdletBinding()]
    param()

    Write-ADReportLog -Message "Lade 'Benutzer mit fehlerhaften Kennworteingaben' Report..." -Type Info -GUI
    try {
        $Properties = @(
            "DisplayName", "SamAccountName", "GivenName", "Surname", "Enabled", "LockedOut",
            "LastLogonDate", "PasswordLastSet", "DistinguishedName", "Created", 
            "Modified", "Description", "mail", "ObjectClass", "badPwdCount", "badPasswordTime"
        )
        
        $UsersWBadPwd = Get-ADUser -Filter 'badPwdCount -gt 0 -and Enabled -eq $true' -Properties $Properties -ErrorAction SilentlyContinue
        
        if ($UsersWBadPwd) {
            Write-ADReportLog -Message "$($UsersWBadPwd.Count) Benutzer mit fehlerhaften Kennworteingaben gefunden." -Type Info -GUI
            return $UsersWBadPwd | Select-Object $Properties
        } else {
            Write-ADReportLog -Message "Keine Benutzer mit fehlerhaften Kennworteingaben gefunden." -Type Warning -GUI
            return @()
        }
    }
    catch {
        $ErrorMessage = "Fehler beim Laden des 'Benutzer mit fehlerhaften Kennworteingaben' Reports: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error -GUI
        return $null
    }
}

# --- Funktion zum Abrufen aller Gruppen (Quick Report) ---
Function Get-ADReportAllGroups {
    [CmdletBinding()]
    param()

    Write-ADReportLog -Message "Lade 'Alle Gruppen' Report..." -Type Info -GUI
    try {
        $Properties = @(
            "Name", "SamAccountName", "GroupCategory", "GroupScope", "ManagedBy", 
            "Description", "DistinguishedName", "Created", "Modified", "whenChanged", "Members" 
            # Members wird für die Spalte MemberCount benötigt, aber nicht direkt angezeigt
        )
        $Groups = Get-ADGroup -Filter * -Properties $Properties -ErrorAction SilentlyContinue
        
        if ($Groups) {
            $GroupsWithMemberCount = $Groups | Select-Object *, @{Name="MemberCount";Expression={$_.Members.Count}}
            Write-ADReportLog -Message "$($Groups.Count) Gruppen gefunden." -Type Info -GUI
            return $GroupsWithMemberCount
        } else {
            Write-ADReportLog -Message "Keine Gruppen gefunden." -Type Warning -GUI
            return @()
        }
    }
    catch {
        $ErrorMessage = "Fehler beim Laden des 'Alle Gruppen' Reports: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error -GUI
        return $null
    }
}

# --- Funktion zum Abrufen von leeren Gruppen (Quick Report) ---
Function Get-ADReportEmptyGroups {
    [CmdletBinding()]
    param()

    Write-ADReportLog -Message "Lade 'Leere Gruppen' Report..." -Type Info -GUI
    try {
        $Properties = @(
            "Name", "SamAccountName", "GroupCategory", "GroupScope", "ManagedBy", 
            "Description", "DistinguishedName", "Created", "Modified", "whenChanged", "Members"
        )
        # Gruppen, bei denen das 'Members'-Attribut nicht existiert oder leer ist
        $EmptyGroups = Get-ADGroup -Filter 'Members -notlike "*"' -Properties $Properties -ErrorAction SilentlyContinue
        
        if ($EmptyGroups) {
            $EmptyGroupsWithMemberCount = $EmptyGroups | Select-Object *, @{Name="MemberCount";Expression={0}} # Explizit 0 für leere Gruppen
            Write-ADReportLog -Message "$($EmptyGroups.Count) leere Gruppen gefunden." -Type Info -GUI
            return $EmptyGroupsWithMemberCount
        } else {
            Write-ADReportLog -Message "Keine leeren Gruppen gefunden." -Type Warning -GUI
            return @()
        }
    }
    catch {
        $ErrorMessage = "Fehler beim Laden des 'Leere Gruppen' Reports: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error -GUI
        return $null
    }
}

# --- Funktion zum Abrufen von Gruppen ohne Manager (Quick Report) ---
Function Get-ADReportGroupsWNoManager {
    [CmdletBinding()]
    param()

    Write-ADReportLog -Message "Lade 'Gruppen ohne Manager' Report..." -Type Info -GUI
    try {
        $Properties = @(
            "Name", "SamAccountName", "GroupCategory", "GroupScope", "ManagedBy", 
            "Description", "DistinguishedName", "Created", "Modified", "whenChanged", "Members"
        )
        # Gruppen, bei denen das 'ManagedBy'-Attribut nicht existiert oder leer ist
        $GroupsWNoManager = Get-ADGroup -Filter 'ManagedBy -notlike "*"' -Properties $Properties -ErrorAction SilentlyContinue
        
        if ($GroupsWNoManager) {
            $GroupsWithMemberCount = $GroupsWNoManager | Select-Object *, @{Name="MemberCount";Expression={$_.Members.Count}}
            Write-ADReportLog -Message "$($GroupsWNoManager.Count) Gruppen ohne Manager gefunden." -Type Info -GUI
            return $GroupsWithMemberCount
        } else {
            Write-ADReportLog -Message "Keine Gruppen ohne Manager gefunden." -Type Warning -GUI
            return @()
        }
    }
    catch {
        $ErrorMessage = "Fehler beim Laden des 'Gruppen ohne Manager' Reports: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error -GUI
        return $null
    }
}

# --- Funktion zum Abrufen von kürzlich erstellten Gruppen (Quick Report) ---
Function Get-ADReportRecentlyCreatedGroups {
    [CmdletBinding()]
    param(
        [int]$Days = 7 # Standardmäßig die letzten 7 Tage
    )

    Write-ADReportLog -Message "Lade 'Kürzlich erstellte Gruppen (letzte $($Days) Tage)' Report..." -Type Info -GUI
    try {
        $Properties = @(
            "Name", "SamAccountName", "GroupCategory", "GroupScope", "ManagedBy", 
            "Description", "DistinguishedName", "Created", "Modified", "whenChanged", "Members"
        )
        $targetDate = (Get-Date).AddDays(-$Days)
        
        $RecentlyCreatedGroups = Get-ADGroup -Filter { Created -ge $targetDate } -Properties $Properties -ErrorAction SilentlyContinue
        
        if ($RecentlyCreatedGroups) {
            $GroupsWithMemberCount = $RecentlyCreatedGroups | Select-Object *, @{Name="MemberCount";Expression={$_.Members.Count}}
            Write-ADReportLog -Message "$($RecentlyCreatedGroups.Count) kürzlich erstellte Gruppen (seit $($Days) Tagen) gefunden." -Type Info -GUI
            return $GroupsWithMemberCount
        } else {
            Write-ADReportLog -Message "Keine kürzlich erstellten Gruppen (seit $($Days) Tagen) gefunden." -Type Warning -GUI
            return @()
        }
    }
    catch {
        $ErrorMessage = "Fehler beim Laden des 'Kürzlich erstellte Gruppen' Reports: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error -GUI
        return $null
    }
}

# --- Funktion zum Abrufen von kürzlich geänderten Gruppen (Quick Report) ---
Function Get-ADReportRecentlyModifiedGroups {
    [CmdletBinding()]
    param(
        [int]$Days = 7 # Standardmäßig die letzten 7 Tage
    )

    Write-ADReportLog -Message "Lade 'Kürzlich geänderte Gruppen (letzte $($Days) Tage)' Report..." -Type Info -GUI
    try {
        $Properties = @(
            "Name", "SamAccountName", "GroupCategory", "GroupScope", "ManagedBy", 
            "Description", "DistinguishedName", "Created", "Modified", "whenChanged", "Members"
        )
        $targetDate = (Get-Date).AddDays(-$Days)
        
        $RecentlyModifiedGroups = Get-ADGroup -Filter { Modified -ge $targetDate } -Properties $Properties -ErrorAction SilentlyContinue
        
        if ($RecentlyModifiedGroups) {
            $GroupsWithMemberCount = $RecentlyModifiedGroups | Select-Object *, @{Name="MemberCount";Expression={$_.Members.Count}}
            Write-ADReportLog -Message "$($RecentlyModifiedGroups.Count) kürzlich geänderte Gruppen (seit $($Days) Tagen) gefunden." -Type Info -GUI
            return $GroupsWithMemberCount
        } else {
            Write-ADReportLog -Message "Keine kürzlich geänderten Gruppen (seit $($Days) Tagen) gefunden." -Type Warning -GUI
            return @()
        }
    }
    catch {
        $ErrorMessage = "Fehler beim Laden des 'Kürzlich geänderte Gruppen' Reports: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error -GUI
        return $null
    }
}

# --- Funktion zum Abrufen von Verteilergruppen (Quick Report) ---
Function Get-ADReportDistributionGroups {
    [CmdletBinding()]
    param()

    Write-ADReportLog -Message "Lade 'Verteilergruppen' Report..." -Type Info -GUI
    try {
        $Properties = @(
            "Name", "SamAccountName", "GroupCategory", "GroupScope", "ManagedBy", 
            "Description", "DistinguishedName", "Created", "Modified", "whenChanged", "Members"
        )
        $DistributionGroups = Get-ADGroup -Filter 'GroupCategory -eq "Distribution"' -Properties $Properties -ErrorAction SilentlyContinue
        
        if ($DistributionGroups) {
            $GroupsWithMemberCount = $DistributionGroups | Select-Object *, @{Name="MemberCount";Expression={$_.Members.Count}}
            Write-ADReportLog -Message "$($DistributionGroups.Count) Verteilergruppen gefunden." -Type Info -GUI
            return $GroupsWithMemberCount
        } else {
            Write-ADReportLog -Message "Keine Verteilergruppen gefunden." -Type Warning -GUI
            return @()
        }
    }
    catch {
        $ErrorMessage = "Fehler beim Laden des 'Verteilergruppen' Reports: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error -GUI
        return $null
    }
}

# --- Funktion zum Abrufen von Sicherheitsgruppen (Quick Report) ---
Function Get-ADReportSecurityGroups {
    [CmdletBinding()]
    param()

    Write-ADReportLog -Message "Lade 'Sicherheitsgruppen' Report..." -Type Info -GUI
    try {
        $Properties = @(
            "Name", "SamAccountName", "GroupCategory", "GroupScope", "ManagedBy", 
            "Description", "DistinguishedName", "Created", "Modified", "whenChanged", "Members"
        )
        $SecurityGroups = Get-ADGroup -Filter 'GroupCategory -eq "Security"' -Properties $Properties -ErrorAction SilentlyContinue
        
        if ($SecurityGroups) {
            $GroupsWithMemberCount = $SecurityGroups | Select-Object *, @{Name="MemberCount";Expression={$_.Members.Count}}
            Write-ADReportLog -Message "$($SecurityGroups.Count) Sicherheitsgruppen gefunden." -Type Info -GUI
            return $GroupsWithMemberCount
        } else {
            Write-ADReportLog -Message "Keine Sicherheitsgruppen gefunden." -Type Warning -GUI
            return @()
        }
    }
    catch {
        $ErrorMessage = "Fehler beim Laden des 'Sicherheitsgruppen' Reports: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error -GUI
        return $null
    }
}

# --- Funktion zum Abrufen aller Computer (Quick Report) ---
Function Get-ADReportAllComputers {
    [CmdletBinding()]
    param()

    Write-ADReportLog -Message "Lade 'Alle Computer' Report..." -Type Info -GUI
    try {
        $Properties = @(
            "Name", "DNSHostName", "OperatingSystem", "OperatingSystemServicePack", "IPv4Address",
            "Enabled", "LastLogonDate", "DistinguishedName", "Created", "Modified", "Description", "ManagedBy", "whenChanged"
        )
        $Computers = Get-ADComputer -Filter * -Properties $Properties -ErrorAction SilentlyContinue
        
        if ($Computers) {
            Write-ADReportLog -Message "$($Computers.Count) Computer gefunden." -Type Info -GUI
            return $Computers | Select-Object $Properties
        } else {
            Write-ADReportLog -Message "Keine Computer gefunden." -Type Warning -GUI
            return @()
        }
    }
    catch {
        $ErrorMessage = "Fehler beim Laden des 'Alle Computer' Reports: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error -GUI
        return $null
    }
}

# --- Funktion zum Abrufen von deaktivierten Computern (Quick Report) ---
Function Get-ADReportDisabledComputers {
    [CmdletBinding()]
    param()

    Write-ADReportLog -Message "Lade 'Deaktivierte Computer' Report..." -Type Info -GUI
    try {
        $Properties = @(
            "Name", "DNSHostName", "OperatingSystem", "OperatingSystemServicePack", "IPv4Address",
            "Enabled", "LastLogonDate", "DistinguishedName", "Created", "Modified", "Description", "ManagedBy", "whenChanged"
        )
        $DisabledComputers = Get-ADComputer -Filter 'Enabled -eq $false' -Properties $Properties -ErrorAction SilentlyContinue
        
        if ($DisabledComputers) {
            Write-ADReportLog -Message "$($DisabledComputers.Count) deaktivierte Computer gefunden." -Type Info -GUI
            return $DisabledComputers | Select-Object $Properties
        } else {
            Write-ADReportLog -Message "Keine deaktivierten Computer gefunden." -Type Warning -GUI
            return @()
        }
    }
    catch {
        $ErrorMessage = "Fehler beim Laden des 'Deaktivierte Computer' Reports: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error -GUI
        return $null
    }
}

# --- Funktion zum Abrufen von inaktiven Computern (Quick Report) ---
Function Get-ADReportInactiveComputers {
    [CmdletBinding()]
    param(
        [int]$InactiveDays = 90
    )

    Write-ADReportLog -Message "Lade 'Inaktive Computer ($($InactiveDays)+ Tage)' Report..." -Type Info -GUI
    try {
        $Properties = @(
            "Name", "DNSHostName", "OperatingSystem", "OperatingSystemServicePack", "IPv4Address",
            "Enabled", "LastLogonDate", "DistinguishedName", "Created", "Modified", "Description", "ManagedBy", "whenChanged"
        )
        $timespan = New-TimeSpan -Days $InactiveDays
        $inactiveDate = (Get-Date).Subtract($timespan)
        
        $InactiveComputers = Get-ADComputer -Filter { (Enabled -eq $true) -and ( (LastLogonDate -lt $inactiveDate) -or (LastLogonDate -eq $null) ) } -Properties $Properties -ErrorAction SilentlyContinue
        
        if ($InactiveComputers) {
            Write-ADReportLog -Message "$($InactiveComputers.Count) inaktive Computer (seit $($InactiveDays)+ Tagen) gefunden." -Type Info -GUI
            return $InactiveComputers | Select-Object $Properties
        } else {
            Write-ADReportLog -Message "Keine inaktiven Computer (seit $($InactiveDays)+ Tagen) gefunden." -Type Warning -GUI
            return @()
        }
    }
    catch {
        $ErrorMessage = "Fehler beim Laden des 'Inaktive Computer' Reports: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error -GUI
        return $null
    }
}

# --- Funktion zum Abrufen von kürzlich erstellten Computern (Quick Report) ---
Function Get-ADReportRecentlyCreatedComputers {
    [CmdletBinding()]
    param(
        [int]$Days = 7 # Standardmäßig die letzten 7 Tage
    )

    Write-ADReportLog -Message "Lade 'Kürzlich erstellte Computer (letzte $($Days) Tage)' Report..." -Type Info -GUI
    try {
        $Properties = @(
            "Name", "DNSHostName", "OperatingSystem", "OperatingSystemServicePack", "IPv4Address",
            "Enabled", "LastLogonDate", "DistinguishedName", "Created", "Modified", "Description", "ManagedBy", "whenChanged"
        )
        $targetDate = (Get-Date).AddDays(-$Days)
        
        $RecentlyCreatedComputers = Get-ADComputer -Filter { Created -ge $targetDate } -Properties $Properties -ErrorAction SilentlyContinue
        
        if ($RecentlyCreatedComputers) {
            Write-ADReportLog -Message "$($RecentlyCreatedComputers.Count) kürzlich erstellte Computer (seit $($Days) Tagen) gefunden." -Type Info -GUI
            return $RecentlyCreatedComputers | Select-Object $Properties
        } else {
            Write-ADReportLog -Message "Keine kürzlich erstellten Computer (seit $($Days) Tagen) gefunden." -Type Warning -GUI
            return @()
        }
    }
    catch {
        $ErrorMessage = "Fehler beim Laden des 'Kürzlich erstellte Computer' Reports: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error -GUI
        return $null
    }
}

# --- Funktion zum Abrufen von kürzlich geänderten Computern (Quick Report) ---
Function Get-ADReportRecentlyModifiedComputers {
    [CmdletBinding()]
    param(
        [int]$Days = 7 # Standardmäßig die letzten 7 Tage
    )

    Write-ADReportLog -Message "Lade 'Kürzlich geänderte Computer (letzte $($Days) Tage)' Report..." -Type Info -GUI
    try {
        $Properties = @(
            "Name", "DNSHostName", "OperatingSystem", "OperatingSystemServicePack", "IPv4Address",
            "Enabled", "LastLogonDate", "DistinguishedName", "Created", "Modified", "Description", "ManagedBy", "whenChanged"
        )
        $targetDate = (Get-Date).AddDays(-$Days)
        
        $RecentlyModifiedComputers = Get-ADComputer -Filter { Modified -ge $targetDate } -Properties $Properties -ErrorAction SilentlyContinue
        
        if ($RecentlyModifiedComputers) {
            Write-ADReportLog -Message "$($RecentlyModifiedComputers.Count) kürzlich geänderte Computer (seit $($Days) Tagen) gefunden." -Type Info -GUI
            return $RecentlyModifiedComputers | Select-Object $Properties
        } else {
            Write-ADReportLog -Message "Keine kürzlich geänderten Computer (seit $($Days) Tagen) gefunden." -Type Warning -GUI
            return @()
        }
    }
    catch {
        $ErrorMessage = "Fehler beim Laden des 'Kürzlich geänderte Computer' Reports: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error -GUI
        return $null
    }
}

# --- Funktion zum Abrufen aller GPOs (Quick Report) ---
Function Get-ADReportAllGPOs {
    [CmdletBinding()]
    param()

    Write-ADReportLog -Message "Lade 'Alle GPOs' Report..." -Type Info -GUI
    try {
        if (-not (Get-Module -ListAvailable -Name GroupPolicy)) {
            Write-ADReportLog -Message "Das GroupPolicy-Modul ist nicht verfügbar. GPO-Berichte können nicht erstellt werden." -Type Error -GUI
            return $null
        }
        Import-Module GroupPolicy -ErrorAction SilentlyContinue # Sicherstellen, dass es geladen ist

        $Properties = @(
            "DisplayName", "Id", "Owner", "GpoStatus", 
            "CreationTime", "ModificationTime", "UserVersion", "ComputerVersion", "WmiFilter"
        )
        # Get-GPO -All benötigt keine expliziten Properties, da das Objekt bereits die meisten Infos enthält.
        # Wir selektieren die gewünschten Properties am Ende für Konsistenz.
        $AllGPOs = Get-GPO -All -ErrorAction SilentlyContinue
        
        if ($AllGPOs) {
            Write-ADReportLog -Message "$($AllGPOs.Count) GPOs gefunden." -Type Info -GUI
            return $AllGPOs | Select-Object $Properties 
        } else {
            Write-ADReportLog -Message "Keine GPOs gefunden." -Type Warning -GUI
            return @()
        }
    }
    catch {
        $ErrorMessage = "Fehler beim Laden des 'Alle GPOs' Reports: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error -GUI
        return $null
    }
}

# --- Funktion zum Abrufen von kürzlich geänderten GPOs (Quick Report) ---
Function Get-ADReportRecentlyModifiedGPOs {
    [CmdletBinding()]
    param(
        [int]$Days = 7 # Standardmäßig die letzten 7 Tage
    )

    Write-ADReportLog -Message "Lade 'Kürzlich geänderte GPOs (letzte $($Days) Tage)' Report..." -Type Info -GUI
    try {
        if (-not (Get-Module -ListAvailable -Name GroupPolicy)) {
            Write-ADReportLog -Message "Das GroupPolicy-Modul ist nicht verfügbar. GPO-Berichte können nicht erstellt werden." -Type Error -GUI
            return $null
        }
        Import-Module GroupPolicy -ErrorAction SilentlyContinue

        $Properties = @(
            "DisplayName", "Id", "Owner", "GpoStatus", 
            "CreationTime", "ModificationTime", "UserVersion", "ComputerVersion", "WmiFilter"
        )
        $targetDate = (Get-Date).AddDays(-$Days)
        
        $RecentlyModifiedGPOs = Get-GPO -All -ErrorAction SilentlyContinue | Where-Object { $_.ModificationTime -ge $targetDate }
        
        if ($RecentlyModifiedGPOs) {
            Write-ADReportLog -Message "$($RecentlyModifiedGPOs.Count) kürzlich geänderte GPOs (seit $($Days) Tagen) gefunden." -Type Info -GUI
            return $RecentlyModifiedGPOs | Select-Object $Properties
        } else {
            Write-ADReportLog -Message "Keine kürzlich geänderten GPOs (seit $($Days) Tagen) gefunden." -Type Warning -GUI
            return @()
        }
    }
    catch {
        $ErrorMessage = "Fehler beim Laden des 'Kürzlich geänderte GPOs' Reports: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error -GUI
        return $null
    }
}

# --- Funktion zum Abrufen von deaktivierten GPOs (Quick Report) ---
Function Get-ADReportDisabledGPOs {
    [CmdletBinding()]
    param()

    Write-ADReportLog -Message "Lade 'Deaktivierte GPOs' Report..." -Type Info -GUI
    try {
        if (-not (Get-Module -ListAvailable -Name GroupPolicy)) {
            Write-ADReportLog -Message "Das GroupPolicy-Modul ist nicht verfügbar. GPO-Berichte können nicht erstellt werden." -Type Error -GUI
            return $null
        }
        Import-Module GroupPolicy -ErrorAction SilentlyContinue

        $Properties = @(
            "DisplayName", "Id", "Owner", "GpoStatus", 
            "CreationTime", "ModificationTime", "UserVersion", "ComputerVersion", "WmiFilter"
        )
        
        $DisabledGPOs = Get-GPO -All -ErrorAction SilentlyContinue | Where-Object { $_.GpoStatus -eq 'AllSettingsDisabled' }
        
        if ($DisabledGPOs) {
            Write-ADReportLog -Message "$($DisabledGPOs.Count) deaktivierte GPOs gefunden." -Type Info -GUI
            return $DisabledGPOs | Select-Object $Properties
        } else {
            Write-ADReportLog -Message "Keine deaktivierten GPOs gefunden." -Type Warning -GUI
            return @()
        }
    }
    catch {
        $ErrorMessage = "Fehler beim Laden des 'Deaktivierte GPOs' Reports: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error -GUI
        return $null
    }
}

# --- Funktion zum Abrufen von GPOs ohne Benutzer-/Computereinstellungen (Quick Report) ---
Function Get-ADReportGPOsWithNoSettings {
    [CmdletBinding()]
    param()

    Write-ADReportLog -Message "Lade 'GPOs ohne Benutzer-/Computereinstellungen' Report..." -Type Info -GUI
    try {
        if (-not (Get-Module -ListAvailable -Name GroupPolicy)) {
            Write-ADReportLog -Message "Das GroupPolicy-Modul ist nicht verfügbar. GPO-Berichte können nicht erstellt werden." -Type Error -GUI
            return $null
        }
        Import-Module GroupPolicy -ErrorAction SilentlyContinue

        $Properties = @(
            "DisplayName", "Id", "Owner", "GpoStatus", 
            "CreationTime", "ModificationTime", "UserVersion", "ComputerVersion", "WmiFilter"
        )
        
        # Sucht nach GPOs, die nicht 'AllSettingsDisabled' sind, aber UserVersion UND ComputerVersion = 0 haben.
        # Dies deutet darauf hin, dass keine Einstellungen konfiguriert sind, obwohl die GPO selbst nicht komplett deaktiviert ist.
        $GPOsNoSettings = Get-GPO -All -ErrorAction SilentlyContinue | Where-Object { 
            ($_.GpoStatus -ne 'AllSettingsDisabled') -and 
            ($_.UserVersion -eq 0) -and 
            ($_.ComputerVersion -eq 0) 
        }
        
        if ($GPOsNoSettings) {
            Write-ADReportLog -Message "$($GPOsNoSettings.Count) GPOs ohne Benutzer-/Computereinstellungen gefunden." -Type Info -GUI
            return $GPOsNoSettings | Select-Object $Properties
        } else {
            Write-ADReportLog -Message "Keine GPOs ohne Benutzer-/Computereinstellungen gefunden." -Type Warning -GUI
            return @()
        }
    }
    catch {
        $ErrorMessage = "Fehler beim Laden des 'GPOs ohne Benutzer-/Computereinstellungen' Reports: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error -GUI
        return $null
    }
}

# --- Funktion zum Abrufen von GPOs mit WMI-Filter (Quick Report) ---
Function Get-ADReportGPOsWithWmiFilter {
    [CmdletBinding()]
    param()

    Write-ADReportLog -Message "Lade 'GPOs mit WMI-Filter' Report..." -Type Info -GUI
    try {
        if (-not (Get-Module -ListAvailable -Name GroupPolicy)) {
            Write-ADReportLog -Message "Das GroupPolicy-Modul ist nicht verfügbar. GPO-Berichte können nicht erstellt werden." -Type Error -GUI
            return $null
        }
        Import-Module GroupPolicy -ErrorAction SilentlyContinue

        $Properties = @(
            "DisplayName", "Id", "Owner", "GpoStatus", 
            "CreationTime", "ModificationTime", "UserVersion", "ComputerVersion", "WmiFilter"
        )
        
        $GPOsWithWmiFilter = Get-GPO -All -ErrorAction SilentlyContinue | Where-Object { $null -ne $_.WmiFilter }
        
        if ($GPOsWithWmiFilter) {
            Write-ADReportLog -Message "$($GPOsWithWmiFilter.Count) GPOs mit WMI-Filter gefunden." -Type Info -GUI
            return $GPOsWithWmiFilter | Select-Object $Properties
        } else {
            Write-ADReportLog -Message "Keine GPOs mit WMI-Filter gefunden." -Type Warning -GUI
            return @()
        }
    }
    catch {
        $ErrorMessage = "Fehler beim Laden des 'GPOs mit WMI-Filter' Reports: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error -GUI
        return $null
    }
}

# --- Funktion zum Abrufen von nicht verknüpften GPOs (Quick Report) ---
Function Get-ADReportGPOsNotLinked {
    [CmdletBinding()]
    param()

    Write-ADReportLog -Message "Lade 'Nicht verknüpfte GPOs' Report..." -Type Info -GUI
    try {
        if (-not (Get-Module -ListAvailable -Name GroupPolicy)) {
            Write-ADReportLog -Message "Das GroupPolicy-Modul ist nicht verfügbar. GPO-Berichte können nicht erstellt werden." -Type Error -GUI
            return $null
        }
        Import-Module GroupPolicy -ErrorAction SilentlyContinue

        $Properties = @(
            "DisplayName", "Id", "Owner", "GpoStatus", 
            "CreationTime", "ModificationTime", "UserVersion", "ComputerVersion", "WmiFilter"
        )
        
        $AllGPOs = Get-GPO -All -ErrorAction SilentlyContinue
        $NotLinkedGPOs = @()

        foreach ($GPO in $AllGPOs) {
            # Erstelle temporären Pfad für den Report
            $tempReportPath = Join-Path $env:TEMP "$($GPO.Id.Guid).xml"
            try {
                Get-GPOReport -Guid $GPO.Id -ReportType Xml -Path $tempReportPath -ErrorAction Stop
                $GPOReportXml = [xml](Get-Content $tempReportPath -ErrorAction Stop)
                
                # Überprüfe, ob der Knoten für Verknüpfungen existiert und Kinder hat
                $linksNode = $GPOReportXml.DocumentElement.LinksTo
                if ($null -eq $linksNode -or -not $linksNode.HasChildNodes) {
                    $NotLinkedGPOs += $GPO
                }
            }
            catch {
                Write-ADReportLog -Message "Fehler beim Verarbeiten des Reports für GPO '$($GPO.DisplayName)' (ID: $($GPO.Id)): $($_.Exception.Message)" -Type Warning
            }
            finally {
                if (Test-Path $tempReportPath) {
                    Remove-Item $tempReportPath -ErrorAction SilentlyContinue
                }
            }
        }
        
        if ($NotLinkedGPOs) {
            Write-ADReportLog -Message "$($NotLinkedGPOs.Count) nicht verknüpfte GPOs gefunden." -Type Info -GUI
            return $NotLinkedGPOs | Select-Object $Properties
        } else {
            Write-ADReportLog -Message "Keine nicht verknüpften GPOs gefunden." -Type Warning -GUI
            return @()
        }
    }
    catch {
        $ErrorMessage = "Fehler beim Laden des 'Nicht verknüpfte GPOs' Reports: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error -GUI
        return $null
    }
}

# --- Funktion zum Abrufen aller DNS-Zonen (Quick Report) ---
Function Get-ADReportDNSZones {
    [CmdletBinding()]
    param()

    Write-ADReportLog -Message "Lade 'DNS-Zonen' Report..." -Type Info -GUI
    try {
        if (-not (Get-Module -ListAvailable -Name DnsServer)) {
            Write-ADReportLog -Message "Das DnsServer-Modul ist nicht verfügbar. DNS-Berichte können nicht erstellt werden." -Type Error -GUI
            return $null
        }
        Import-Module DnsServer -ErrorAction SilentlyContinue

        $Properties = @(
            "ZoneName", "ZoneType", "IsReverseLookupZone", "IsDsIntegrated", 
            "DynamicUpdate", "ReplicationScope", "DirectoryPartitionName" 
        )
        # Get-DnsServerZone benötigt keinen -ComputerName, wenn der lokale Server abgefragt wird.
        $DnsZones = Get-DnsServerZone -ErrorAction SilentlyContinue
        
        if ($DnsZones) {
            Write-ADReportLog -Message "$($DnsZones.Count) DNS-Zonen gefunden." -Type Info -GUI
            return $DnsZones | Select-Object $Properties
        } else {
            Write-ADReportLog -Message "Keine DNS-Zonen gefunden." -Type Warning -GUI
            return @()
        }
    }
    catch {
        $ErrorMessage = "Fehler beim Laden des 'DNS-Zonen' Reports: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error -GUI
        return $null
    }
}

# --- Funktion zum Abrufen aller DNS-Records (Quick Report) ---
Function Get-ADReportAllDNSRecords {
    [CmdletBinding()]
    param()

    Write-ADReportLog -Message "Lade 'Alle DNS-Records' Report..." -Type Info -GUI
    try {
        if (-not (Get-Module -ListAvailable -Name DnsServer)) {
            Write-ADReportLog -Message "Das DnsServer-Modul ist nicht verfügbar. DNS-Berichte können nicht erstellt werden." -Type Error -GUI
            return $null
        }
        Import-Module DnsServer -ErrorAction SilentlyContinue

        $Properties = @(
            "HostName", "RecordType", "RecordClass", "TimeToLive", "Timestamp", "RecordData"
        )
        
        $AllDnsZones = Get-DnsServerZone -ErrorAction SilentlyContinue
        $AllDnsRecords = @()

        if ($AllDnsZones) {
            foreach ($Zone in $AllDnsZones) {
                $RecordsInZone = Get-DnsServerResourceRecord -ZoneName $Zone.ZoneName -ErrorAction SilentlyContinue
                if ($RecordsInZone) {
                    $AllDnsRecords += $RecordsInZone
                }
            }
        }
        
        if ($AllDnsRecords) {
            Write-ADReportLog -Message "$($AllDnsRecords.Count) DNS-Records in allen Zonen gefunden." -Type Info -GUI
            # RecordData kann komplex sein, daher ist es besser, das ganze Objekt zurückzugeben oder spezifisch zu formatieren
            return $AllDnsRecords | Select-Object $Properties, @{Name="ZoneName"; Expression={$_.ZoneName}}
        } else {
            Write-ADReportLog -Message "Keine DNS-Records gefunden." -Type Warning -GUI
            return @()
        }
    }
    catch {
        $ErrorMessage = "Fehler beim Laden des 'Alle DNS-Records' Reports: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error -GUI
        return $null
    }
}

Function Get-ADReportAllUsers {
    [CmdletBinding()]
    param()

    Write-ADReportLog -Message "Lade 'Alle Benutzer' Report..." -Type Info -GUI
    try {
        $Properties = @(
            "DisplayName", "SamAccountName", "GivenName", "Surname", "Enabled", 
            "LastLogonDate", "PasswordLastSet", "PasswordNeverExpires", "LockedOut",
            "DistinguishedName", "Created", "Modified", "Description", "mail", "ObjectClass"
        )
        $Users = Get-ADUser -Filter * -Properties $Properties -ErrorAction SilentlyContinue
        
        if ($Users) {
            Write-ADReportLog -Message "$($Users.Count) Benutzer gefunden." -Type Info -GUI
            return $Users | Select-Object $Properties
        } else {
            Write-ADReportLog -Message "Keine Benutzer gefunden." -Type Warning -GUI
            return @()
        }
    }
    catch {
        $ErrorMessage = "Fehler beim Laden des 'Alle Benutzer' Reports: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error -GUI
        return $null
    }
}

# --- Funktion zum Abrufen deaktivierter Benutzer (Quick Report) ---
Function Get-ADReportDisabledUsers {
    [CmdletBinding()]
    param()

    Write-ADReportLog -Message "Lade 'Deaktivierte Benutzer' Report..." -Type Info -GUI
    try {
        $Properties = @(
            "DisplayName", "SamAccountName", "GivenName", "Surname", "Enabled", 
            "LastLogonDate", "PasswordLastSet", "DistinguishedName", "Created", 
            "Modified", "Description", "mail", "ObjectClass"
        )
        # Benutzer, die nicht aktiviert sind
        $DisabledUsers = Get-ADUser -Filter 'Enabled -eq $false' -Properties $Properties -ErrorAction SilentlyContinue
        
        if ($DisabledUsers) {
            Write-ADReportLog -Message "$($DisabledUsers.Count) deaktivierte Benutzer gefunden." -Type Info -GUI
            return $DisabledUsers | Select-Object $Properties
        } else {
            Write-ADReportLog -Message "Keine deaktivierten Benutzer gefunden." -Type Warning -GUI
            return @()
        }
    }
    catch {
        $ErrorMessage = "Fehler beim Laden des 'Deaktivierte Benutzer' Reports: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error -GUI
        return $null
    }
}


Function Get-ADPasswordNeverExpiresReport {
    [CmdletBinding()]
    param()

    Write-ADReportLog -Message "Lade 'Password Never Expires' Report..." -Type Info -GUI
    try {
        # Das Attribut für 'PasswordNeverExpires' ist Teil von UserAccountControl
        # UserAccountControl -band 0x10000 (ADS_UF_DONT_EXPIRE_PASSWD)
        # Get-ADUser gibt 'PasswordNeverExpires' als berechnete Eigenschaft zurück.
        $Users = Get-ADUser -Filter {PasswordNeverExpires -eq $true} -Properties DisplayName, SamAccountName, GivenName, Surname, Enabled, LastLogonDate, WhenCreated, DistinguishedName, PasswordNeverExpires -ErrorAction Stop
        if ($Users) {
            Write-ADReportLog -Message "$($Users.Count) Benutzer mit 'Password Never Expires' gefunden." -Type Info -GUI
            return $Users | Select-Object DisplayName, SamAccountName, GivenName, Surname, Enabled, LastLogonDate, WhenCreated, DistinguishedName, PasswordNeverExpires
        } else {
            Write-ADReportLog -Message "Keine Benutzer mit 'Password Never Expires' gefunden." -Type Info -GUI
            return $null
        }
    } catch {
        Write-ADReportLog -Message "Fehler beim Abrufen von Benutzern mit 'Password Never Expires': $($_.Exception.Message)" -Type Error -GUI
        return $null
    }
}

# --- Funktion zum Abrufen kürzlich erstellter Benutzer (Quick Report) ---
Function Get-ADRecentlyCreatedUsersReport {
    [CmdletBinding()]
    param(
        [int]$DaysAgo = 30 # Standardmäßig die letzten 30 Tage
    )

    Write-ADReportLog -Message "Lade 'Kürzlich erstellte Benutzer (letzte $($DaysAgo) Tage)' Report..." -Type Info -GUI
    try {
        $cutoffDate = (Get-Date).AddDays(-$DaysAgo)
        $Users = Get-ADUser -Filter {WhenCreated -ge $cutoffDate} -Properties DisplayName, SamAccountName, GivenName, Surname, Enabled, WhenCreated, DistinguishedName, LockedOut, LastLogonDate -ErrorAction Stop # LockedOut, LastLogonDate für Vollständigkeit hinzugefügt

        if ($Users) {
            $UserList = $Users | Select-Object DisplayName, SamAccountName, GivenName, Surname, Enabled, WhenCreated, DistinguishedName, LockedOut, LastLogonDate
            Write-ADReportLog -Message "$($UserList.Count) kürzlich erstellte Benutzer (letzte $($DaysAgo) Tage) gefunden." -Type Info -GUI
            return $UserList
        } else {
            Write-ADReportLog -Message "Keine kürzlich erstellten Benutzer (letzte $($DaysAgo) Tage) gefunden." -Type Info -GUI
            return $null
        }
    } catch {
        Write-ADReportLog -Message "Fehler beim Abrufen kürzlich erstellter Benutzer: $($_.Exception.Message)" -Type Error -GUI
        return $null
    }
}

# --- Funktion zum Abrufen inaktiver Computer (Quick Report) ---
Function Get-ADInactiveComputersReport {
    [CmdletBinding()]
    param(
        [int]$InactiveDays = 90
    )

    Write-ADReportLog -Message "Lade 'Inaktive Computer (seit $($InactiveDays) Tagen)' Report..." -Type Info -GUI
    try {
        $timespan = New-TimeSpan -Days $InactiveDays
        # Get-ADComputer -Inactive -TimeSpan verwendet LastLogonTimestamp, was nicht immer repliziert wird.
        # Eine Alternative ist das Filtern nach PasswordLastSet oder LastLogonDate, wenn verfügbar und zuverlässig.
        # Für eine einfachere Implementierung hier zunächst der Versuch mit -Inactive.
        # Beachte: LastLogonTimestamp ist genauer, aber nicht auf allen DCs aktuell.
        # LastLogonDate ist eine berechnete, benutzerfreundlichere Version von LastLogon.
        $Computers = Get-ADComputer -Filter * -Properties DisplayName, SamAccountName, OperatingSystem, Enabled, LastLogonDate, WhenCreated, DistinguishedName, PasswordLastSet -ErrorAction Stop | Where-Object {
            ($_.LastLogonDate -ne $null -and $_.LastLogonDate -lt (Get-Date).Subtract($timespan)) -or 
            ($_.LastLogonDate -eq $null -and $_.PasswordLastSet -ne $null -and $_.PasswordLastSet -lt (Get-Date).Subtract($timespan)) -or
            ($_.LastLogonDate -eq $null -and $_.PasswordLastSet -eq $null -and $_.WhenCreated -lt (Get-Date).Subtract($timespan)) # Fallback, wenn beide Datumswerte fehlen
        }

        if ($Computers) {
            $ComputerList = $Computers | Select-Object DisplayName, SamAccountName, OperatingSystem, Enabled, LastLogonDate, PasswordLastSet, WhenCreated, DistinguishedName
            Write-ADReportLog -Message "$($ComputerList.Count) inaktive Computer (seit $($InactiveDays) Tagen) gefunden." -Type Info -GUI
            return $ComputerList
        } else {
            Write-ADReportLog -Message "Keine inaktiven Computer (seit $($InactiveDays) Tagen) gefunden." -Type Info -GUI
            return $null
        }
    } catch {
        Write-ADReportLog -Message "Fehler beim Abrufen inaktiver Computer: $($_.Exception.Message)" -Type Error -GUI
        return $null
    }
}

# --- Funktion zum Abrufen gesperrter Benutzer (Quick Report) ---
Function Get-ADLockedOutUsersReport {
    [CmdletBinding()]
    param()

    Write-ADReportLog -Message "Lade 'Gesperrte Benutzer' Report..." -Type Info -GUI
    try {
        $Users = Get-ADUser -Filter {LockedOut -eq $true} -Properties DisplayName, SamAccountName, GivenName, Surname, Enabled, LastLogonDate, WhenCreated, DistinguishedName, LockedOut -ErrorAction Stop
        if ($Users) {
            Write-ADReportLog -Message "$($Users.Count) gesperrte Benutzer gefunden." -Type Info -GUI
            return $Users | Select-Object DisplayName, SamAccountName, GivenName, Surname, Enabled, LastLogonDate, WhenCreated, DistinguishedName, LockedOut
        } else {
            Write-ADReportLog -Message "Keine gesperrten Benutzer gefunden." -Type Info -GUI
            return $null
        }
    } catch {
        Write-ADReportLog -Message "Fehler beim Abrufen gesperrter Benutzer: $($_.Exception.Message)" -Type Error -GUI
        return $null
    }
}

# --- Funktion zum Abrufen aller Benutzer (Quick Report) ---
Function Get-ADAllUsersReport {
    [CmdletBinding()]
    param()

    Write-ADReportLog -Message "Lade 'Alle Benutzer' Report..." -Type Info -GUI
    try {
        $Users = Get-ADUser -Filter * -Properties DisplayName, SamAccountName, GivenName, Surname, Enabled, LastLogonDate, WhenCreated, DistinguishedName -ErrorAction Stop
        if ($Users) {
            Write-ADReportLog -Message "$($Users.Count) Benutzer gefunden." -Type Info -GUI
            return $Users | Select-Object DisplayName, SamAccountName, GivenName, Surname, Enabled, LastLogonDate, WhenCreated, DistinguishedName
        } else {
            Write-ADReportLog -Message "Keine Benutzer gefunden." -Type Info -GUI
            return $null
        }
    } catch {
        Write-ADReportLog -Message "Fehler beim Abrufen aller Benutzer: $($_.Exception.Message)" -Type Error -GUI
        return $null
    }
}

# --- Funktion zum Abrufen von AD Sites und Subnets ---
Function Get-ADSitesAndSubnetsReport {
    [CmdletBinding()]
    param()

    Write-ADReportLog -Message "Gathering AD Sites and Subnets information..." -Type Info -GUI
    Initialize-ResultCounters

    try {
        $Report = @()

        # Sites abrufen
        $Sites = Get-ADReplicationSite -Filter * -Properties Description, Options, InterSiteTopologyGenerator -ErrorAction SilentlyContinue
        if ($Sites) {
            foreach ($Site in $Sites) {
                # Server für diesen Site separat abrufen
                $ServersCount = 0
                try {
                    $Servers = Get-ADDomainController -Filter { Site -eq $Site.Name } -ErrorAction SilentlyContinue
                    $ServersCount = if ($Servers) { @($Servers).Count } else { 0 }
                } catch {
                    Write-ADReportLog -Message "Could not get servers for site $($Site.Name): $($_.Exception.Message)" -Type Warning -GUI
                }

                # Site Links separat abrufen
                $SiteLinksText = ""
                try {
                    $SiteLinks = Get-ADReplicationSiteLink -Filter "(Sites -eq '$($Site.DistinguishedName)')" -ErrorAction SilentlyContinue
                    $SiteLinksText = if ($SiteLinks) { ($SiteLinks | ForEach-Object { $_.Name }) -join ", " }
                } catch {
                    Write-ADReportLog -Message "Could not get site links for site $($Site.Name): $($_.Exception.Message)" -Type Warning -GUI
                }

                $Report += [PSCustomObject]@{
                    Type = "Site"
                    Name = $Site.Name
                    DistinguishedName = $Site.DistinguishedName
                    Description = $Site.Description
                    ServersInSiteCount = $ServersCount
                    InterSiteTopologyGenerator = $Site.InterSiteTopologyGenerator
                    Options = $Site.Options
                    SiteLinks = $SiteLinksText
                    Location = $null # Für Konsistenz
                    AssociatedSite = $null # Für Konsistenz mit Subnets
                }
            }
            Write-ADReportLog -Message "Successfully retrieved $($Sites.Count) AD Replication Sites." -Type Info -GUI
        } else {
            Write-ADReportLog -Message "No AD Replication Sites found." -Type Warning -GUI
        }

        # Subnets abrufen
        $Subnets = Get-ADReplicationSubnet -Filter * -Properties Description, Location, Site -ErrorAction SilentlyContinue
        if ($Subnets) {
            foreach ($Subnet in $Subnets) {
                $associatedSiteValue = $null
                try {
                    # Versuche, den Namen der Site abzurufen
                    $siteObject = Get-ADReplicationSite -Identity $Subnet.Site -ErrorAction SilentlyContinue
                    if ($siteObject) {
                        $associatedSiteValue = $siteObject.Name
                    } else {
                        # Wenn Get-ADReplicationSite $null zurückgibt (Site nicht gefunden, aber kein Fehler), verwende den Rohwert
                        $associatedSiteValue = $Subnet.Site 
                    }
                } catch {
                    # Im Fehlerfall (z.B. ungültige Identity), verwende den Rohwert
                    $associatedSiteValue = $Subnet.Site
                }

                $Report += [PSCustomObject]@{
                    Type = "Subnet"
                    Name = $Subnet.Name
                    DistinguishedName = $Subnet.DistinguishedName
                    Description = $Subnet.Description
                    ServersInSiteCount = $null 
                    InterSiteTopologyGenerator = $null
                    Options = $null
                    SiteLinks = $null # Für Konsistenz
                    Location = $Subnet.Location
                    AssociatedSite = $associatedSiteValue
                }
            }
            Write-ADReportLog -Message "Successfully retrieved $($Subnets.Count) AD Replication Subnets." -Type Info -GUI
        } else {
            Write-ADReportLog -Message "No AD Replication Subnets found." -Type Warning -GUI
        }

        if ($Report.Count -gt 0) {
            Write-ADReportLog -Message "Successfully generated Sites and Subnets Report for $($Report.Count) entries." -Type Info -GUI
            return $Report | Sort-Object Type, Name
        } else {
            Write-ADReportLog -Message "No data found for Sites and Subnets Report." -Type Info -GUI
            return @(
                [PSCustomObject]@{
                    Type = "Information"
                    Name = "No Data"
                    DistinguishedName = "N/A"
                    Description = "No Sites or Subnets found in Active Directory"
                    ServersInSiteCount = 0
                    InterSiteTopologyGenerator = "N/A"
                    Options = "N/A"
                    SiteLinks = "N/A"
                    Location = "N/A"
                    AssociatedSite = "N/A"
                }
            )
        }
    }
    catch {
        $ErrorMessage = "Error generating Sites and Subnets Report: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error -GUI
            $errorObject = [PSCustomObject]@{
                Type = "Error"
                Name = "System Error"
                DistinguishedName = "N/A"
                Description = $ErrorMessage
                ServersInSiteCount = 0
                InterSiteTopologyGenerator = "N/A"
                Options = "N/A"
                SiteLinks = "N/A"
                Location = "N/A"
                AssociatedSite = "N/A"
            }
            return @($errorObject)
    }
}

# Funktion zum Aktualisieren der Ergebniszähler im Header
    Function Initialize-ResultCounters {
        # Gesamtergebniszähler auf 0 setzen
        if ($null -ne $Global:TotalResultCountText) {
            $Global:TotalResultCountText.Text = "0"
        }
        
        # Sicherstellen, dass alle Zähler zurückgesetzt werden
        # if ($null -ne $Global:UserCountText) {
        #     $Global:UserCountText.Text = "0"
        # }
        
        # if ($null -ne $Global:ComputerCountText) {
        #     $Global:ComputerCountText.Text = "0"
        # }
        
        # if ($null -ne $Global:GroupCountText) {
        #     $Global:GroupCountText.Text = "0"
        # }
        
        # Status zurücksetzen
        if ($null -ne $Global:TextBlockStatus) {
            $Global:TextBlockStatus.Text = "Ready for query..."
        }
        
        # DataGrid leeren
        if ($null -ne $Global:DataGridResults) {
            $Global:DataGridResults.ItemsSource = $null
        }
    }

    # Funktion zum Aktualisieren der Ergebniszähler im Header
    Function Update-ResultCounters {
        param (
            [Parameter(Mandatory=$false)]
            [System.Collections.IList]$Results = @()
        )

        try {
            if ($null -eq $Results -or $Results.Count -eq 0) {
                if ($null -ne $Global:TotalResultCountText) { 
                    $Global:TotalResultCountText.Text = "0" 
                }
                return
            }
        } catch {
            Write-Warning "Error in Update-ResultCounters: $($_.Exception.Message)"
            if ($null -ne $Global:TotalResultCountText) { 
                $Global:TotalResultCountText.Text = "Error" 
            }
            return
        }

        $totalCount = $Results.Count
        $userCount = 0
        $computerCount = 0
        $groupCount = 0

        if ($totalCount -gt 0) {
            if ($null -ne $Results) {
                foreach ($item in $Results) {
                    if ($null -ne $item -and $null -ne $item.PSObject) {
                        if ($item.PSObject.Properties["ObjectClass"] -and $item.ObjectClass -is [array] -and $item.ObjectClass -contains "user") {
                            $userCount++
                        } elseif ($item.PSObject.Properties["ObjectClass"] -and $item.ObjectClass -is [string] -and $item.ObjectClass -eq "user") {
                            $userCount++
                        } elseif ($item.PSObject.Properties["ObjectClass"] -and $item.ObjectClass -is [array] -and $item.ObjectClass -contains "computer") {
                            $computerCount++
                        } elseif ($item.PSObject.Properties["ObjectClass"] -and $item.ObjectClass -is [string] -and $item.ObjectClass -eq "computer") {
                            $computerCount++
                        } elseif ($item.PSObject.Properties["ObjectClass"] -and $item.ObjectClass -is [array] -and $item.ObjectClass -contains "group") {
                            $groupCount++
                        } elseif ($item.PSObject.Properties["ObjectClass"] -and $item.ObjectClass -is [string] -and $item.ObjectClass -eq "group") {
                            $groupCount++
                        } elseif ($item.PSObject.Properties["ObjectCategory"] -and $item.ObjectCategory -like "*CN=Person,CN=Schema,CN=Configuration,*") {
                            $userCount++
                        } elseif ($item.PSObject.Properties["ObjectCategory"] -and $item.ObjectCategory -like "*CN=Computer,CN=Schema,CN=Configuration,*") {
                            $computerCount++
                        } elseif ($item.PSObject.Properties["ObjectCategory"] -and $item.ObjectCategory -like "*CN=Group,CN=Schema,CN=Configuration,*") {
                            $groupCount++
                        }
                    }
                }
            }
        }

        try {
            if ($null -ne $Global:TotalResultCountText) { 
                $Global:TotalResultCountText.Text = $totalCount.ToString() 
            }
        } catch {
            Write-Warning "Error updating result counters display: $($_.Exception.Message)"
        }
    }

    # Funktion zur universellen Aktualisierung der Ergebnisanzeige und DataGrid
    Function Update-ADReportResults {
    param (
        [Parameter(Mandatory=$false)]
        $Results = $null,
        
        [Parameter(Mandatory=$false)]
        [string]$StatusMessage
    )

    if ($null -ne $Global:DataGridResults) {
        $Global:DataGridResults.ItemsSource = @($Results) # Sicherstellen, dass es immer eine Sammlung ist
    }

    if (-not [string]::IsNullOrWhiteSpace($StatusMessage) -and $null -ne $Global:TextBlockStatus) {
        $Global:TextBlockStatus.Text = $StatusMessage
    }
    
    # Sicherstellen, dass wir keine leere Auflistung übergeben, sondern ein Array
    $ResultsCollection = if ($null -eq $Results) { @() } else { @($Results) }
    Update-ResultCounters -Results $ResultsCollection
    Update-ResultVisualization -Results $ResultsCollection
} 

# --- Globale Variablen für UI Elemente --- 
Function Initialize-ADReportForm {
    param($XAMLContent)
    # Überprüfen, ob das Window-Objekt bereits existiert und zurücksetzen
    if ($Global:Window) {
        Remove-Variable -Name Window -Scope Global -ErrorAction SilentlyContinue
    }
    
    try {
        $reader = New-Object System.Xml.XmlNodeReader $XAMLContent
        $Global:Window = [Windows.Markup.XamlReader]::Load( $reader )
        
        if ($null -eq $Global:Window) {
            Write-Error "Failed to load XAML window"
            return
        }
    } catch {
        Write-Error "Error loading XAML: $($_.Exception.Message)"
        return
    }

    # --- UI Elemente zu globalen Variablen zuweisen --- 
    # Objekttyp Radio Buttons
    $Global:RadioButtonUser = $Window.FindName("RadioButtonUser")
    $Global:RadioButtonGroup = $Window.FindName("RadioButtonGroup")
    $Global:RadioButtonComputer = $Window.FindName("RadioButtonComputer")
    $Global:RadioButtonGroupMemberships = $Window.FindName("RadioButtonGroupMemberships")

    # Filter und Attribute
    $Global:ComboBoxFilterAttribute = $Window.FindName("ComboBoxFilterAttribute")
    $Global:TextBoxFilterValue = $Window.FindName("TextBoxFilterValue")
    $Global:ListBoxSelectAttributes = $Window.FindName("ListBoxSelectAttributes")
    $Global:ButtonQueryAD = $Window.FindName("ButtonQueryAD")

    # Quick Reports - Benutzer
    $Global:ButtonQuickAllUsers = $Window.FindName("ButtonQuickAllUsers")
    $Global:ButtonQuickDisabledUsers = $Window.FindName("ButtonQuickDisabledUsers")
    $Global:ButtonQuickLockedUsers = $Window.FindName("ButtonQuickLockedUsers")
    $Global:ButtonQuickNeverExpire = $Window.FindName("ButtonQuickNeverExpire")
    $Global:ButtonQuickInactiveUsers = $Window.FindName("ButtonQuickInactiveUsers")
    $Global:ButtonQuickAdminUsers = $Window.FindName("ButtonQuickAdminUsers")
    $Global:ButtonQuickRecentlyCreatedUsers = $Window.FindName("ButtonQuickRecentlyCreatedUsers")
    $Global:ButtonQuickPasswordExpiringSoon = $Window.FindName("ButtonQuickPasswordExpiringSoon")
    $Global:ButtonQuickExpiredPasswords = $Window.FindName("ButtonQuickExpiredPasswords")
    $Global:ButtonQuickNeverLoggedOn = $Window.FindName("ButtonQuickNeverLoggedOn")
    
    # Quick Reports - Gruppen
    $Global:ButtonQuickGroups = $Window.FindName("ButtonQuickGroups")
    $Global:ButtonQuickSecurityGroups = $Window.FindName("ButtonQuickSecurityGroups")
    $Global:ButtonQuickDistributionGroups = $Window.FindName("ButtonQuickDistributionGroups")
    
    # Quick Reports - Computer
    $Global:ButtonQuickComputers = $Window.FindName("ButtonQuickComputers")
    $Global:ButtonQuickInactiveComputers = $Window.FindName("ButtonQuickInactiveComputers")
    
    # Security Audit
    $Global:ButtonQuickWeakPasswordPolicy = $Window.FindName("ButtonQuickWeakPasswordPolicy")
    $Global:ButtonQuickRiskyGroupMemberships = $Window.FindName("ButtonQuickRiskyGroupMemberships")
    $Global:ButtonQuickPrivilegedAccounts = $Window.FindName("ButtonQuickPrivilegedAccounts")
    
    # AD-Health
    $Global:ButtonQuickFSMORoles = $Window.FindName("ButtonQuickFSMORoles")
    $Global:ButtonQuickDCStatus = $Window.FindName("ButtonQuickDCStatus")
    $Global:ButtonQuickReplicationStatus = $Window.FindName("ButtonQuickReplicationStatus")
    
    # User Reports
    # Die folgenden Zeilen wurden entfernt, da sie Duplikate von vorherigen Zuweisungen waren:
    # $Global:ButtonQuickAllUsers = $Window.FindName("ButtonQuickAllUsers")
    # $Global:ButtonQuickDisabledUsers = $Window.FindName("ButtonQuickDisabledUsers")
    # $Global:ButtonQuickRecentlyCreatedUsers = $Window.FindName("ButtonQuickRecentlyCreatedUsers")
    # $Global:ButtonQuickInactiveComputers = $Window.FindName("ButtonQuickInactiveComputers")


    # OU & Topology Reports
    $Global:ButtonQuickOUHierarchy = $Window.FindName("ButtonQuickOUHierarchy")
    $Global:ButtonQuickSitesSubnets = $Window.FindName("ButtonQuickSitesSubnets")

    # Ergebnisanzahl-Anzeige
    $Global:ResultCountGrid = $Window.FindName("ResultCountGrid")
    $Global:TotalResultCountText = $Window.FindName("TotalResultCountText")
    # $Global:UserCountText = $Window.FindName("UserCountText")
    # $Global:ComputerCountText = $Window.FindName("ComputerCountText")
    # $Global:GroupCountText = $Window.FindName("GroupCountText") # Auskommentiert, da GroupCountText UI Element nicht existiert
    
    # Ergebnisse
    $Global:DataGridResults = $Window.FindName("DataGridResults")

    # Status und Export
    $Global:TextBlockStatus = $Window.FindName("TextBlockStatus")
    $Global:ButtonExportCSV = $Window.FindName("ButtonExportCSV")
    $Global:ButtonExportHTML = $Window.FindName("ButtonExportHTML")
    $Global:ButtonClose = $Window.FindName("ButtonClose")

    # Weitere fehlende Button-Referenzen
    $Global:ButtonQuickRecentlyDeletedUsers = $Window.FindName("ButtonQuickRecentlyDeletedUsers")
    $Global:ButtonQuickRecentlyModifiedUsers = $Window.FindName("ButtonQuickRecentlyModifiedUsers")
    $Global:ButtonQuickInactiveUsersXDays = $Window.FindName("ButtonQuickInactiveUsersXDays")
    $Global:ButtonQuickUsersWithoutManager = $Window.FindName("ButtonQuickUsersWithoutManager")
    $Global:ButtonQuickUsersMissingRequiredAttributes = $Window.FindName("ButtonQuickUsersMissingRequiredAttributes")
    $Global:ButtonQuickUsersDuplicateLogonNames = $Window.FindName("ButtonQuickUsersDuplicateLogonNames")
    $Global:ButtonQuickOrphanedSIDsUsers = $Window.FindName("ButtonQuickOrphanedSIDsUsers")
    $Global:ButtonQuickRecentlyCreatedGroups = $Window.FindName("ButtonQuickRecentlyCreatedGroups")
    $Global:ButtonQuickRecentlyDeletedGroups = $Window.FindName("ButtonQuickRecentlyDeletedGroups")
    $Global:ButtonQuickEmptyGroups = $Window.FindName("ButtonQuickEmptyGroups")
    $Global:ButtonQuickVeryLargeGroupsXMembers = $Window.FindName("ButtonQuickVeryLargeGroupsXMembers")
    $Global:ButtonQuickDeeplyNestedGroupsXDepth = $Window.FindName("ButtonQuickDeeplyNestedGroupsXDepth")
    $Global:ButtonQuickGroupsWithoutOwner = $Window.FindName("ButtonQuickGroupsWithoutOwner")
    $Global:ButtonQuickGroupsWithoutDescription = $Window.FindName("ButtonQuickGroupsWithoutDescription")
    $Global:ButtonQuickDistListsWithoutMailbox = $Window.FindName("ButtonQuickDistListsWithoutMailbox")
    $Global:ButtonQuickGroupNestingInfo = $Window.FindName("ButtonQuickGroupNestingInfo")

    # Event Handler für Quick Reports (OU & Topology)
    if ($null -ne $Global:ButtonQuickRecentlyCreatedUsers) {
        $Global:ButtonQuickRecentlyCreatedUsers.Add_Click({
            Write-ADReportLog -Message "'Kürzlich erstellte Benutzer' Report wird geladen..." -Type Info
            Initialize-ResultCounters
            $ReportData = Get-ADRecentlyCreatedUsersReport # Standardmäßig 30 Tage
            Update-ADReportResults -Results $ReportData -StatusMessage "'Kürzlich erstellte Benutzer (30 Tage)' Report geladen."
        })
    } else {
        Write-ADReportLog -Message "GUI Element 'ButtonQuickRecentlyCreatedUsers' nicht gefunden. Event-Handler nicht zugewiesen." -Type Warning
    }

    # Event Handler für ButtonQuickAllUsers
    if ($null -ne $Global:ButtonQuickAllUsers) {
        $Global:ButtonQuickAllUsers.add_Click({
            Write-ADReportLog -Message "Button 'Alle Benutzer' geklickt." -Type Info
            Initialize-ResultCounters # Zähler und UI zurücksetzen
            $ReportData = Get-ADReportAllUsers
            Update-ADReportResults -Results $ReportData 
        })
    } else {
        Write-ADReportLog -Message "GUI Element 'ButtonQuickAllUsers' nicht gefunden. Event-Handler nicht zugewiesen." -Type Warning
    }

    # Event Handler für ButtonQuickDisabledUsers
    if ($null -ne $Global:ButtonQuickDisabledUsers) {
        $Global:ButtonQuickDisabledUsers.add_Click({
            Write-ADReportLog -Message "Button 'Deaktivierte Benutzer' geklickt." -Type Info
            Initialize-ResultCounters # Zähler und UI zurücksetzen
            $ReportData = Get-ADReportDisabledUsers
            Update-ADReportResults -Results $ReportData 
        })
    } else {
        Write-ADReportLog -Message "GUI Element 'ButtonQuickDisabledUsers' nicht gefunden. Event-Handler nicht zugewiesen." -Type Warning
    }

    # Event Handler für ButtonQuickLockedUsers
    if ($null -ne $Global:ButtonQuickLockedUsers) {
        $Global:ButtonQuickLockedUsers.add_Click({
            Write-ADReportLog -Message "Button 'Gesperrte Benutzer' geklickt." -Type Info
            Initialize-ResultCounters # Zähler und UI zurücksetzen
            $ReportData = Get-ADLockedOutUsersReport # Bestehende Funktion verwenden
            Update-ADReportResults -Results $ReportData 
        })
    } else {
        Write-ADReportLog -Message "GUI Element 'ButtonQuickLockedUsers' nicht gefunden. Event-Handler nicht zugewiesen." -Type Warning
    }

    # Event Handler für ButtonQuickNeverExpire
    if ($null -ne $Global:ButtonQuickNeverExpire) {
        $Global:ButtonQuickNeverExpire.add_Click({
            Write-ADReportLog -Message "Button 'Password Never Expires' geklickt." -Type Info
            Initialize-ResultCounters # Zähler und UI zurücksetzen
            $ReportData = Get-ADPasswordNeverExpiresReport # Bestehende Funktion verwenden
            Update-ADReportResults -Results $ReportData 
        })
    } else {
        Write-ADReportLog -Message "GUI Element 'ButtonQuickNeverExpire' nicht gefunden. Event-Handler nicht zugewiesen." -Type Warning
    }

    # Event Handler für ButtonQuickInactiveUsers
    if ($null -ne $Global:ButtonQuickInactiveUsers) {
        $Global:ButtonQuickInactiveUsers.add_Click({
            Write-ADReportLog -Message "Button 'Inaktive Benutzer (90+ Tage)' geklickt." -Type Info
            Initialize-ResultCounters # Zähler und UI zurücksetzen
            $ReportData = Get-ADReportInactiveUsers # Standardmäßig 90 Tage
            Update-ADReportResults -Results $ReportData 
        })
    } else {
        Write-ADReportLog -Message "GUI Element 'ButtonQuickInactiveUsers' nicht gefunden. Event-Handler nicht zugewiesen." -Type Warning
    }

    # Event Handler für ButtonQuickAdminUsers
    if ($null -ne $Global:ButtonQuickAdminUsers) {
        $Global:ButtonQuickAdminUsers.add_Click({
            Write-ADReportLog -Message "Button 'Administratoren' geklickt." -Type Info
            Initialize-ResultCounters # Zähler und UI zurücksetzen
            $ReportData = Get-ADReportAdminUsers
            Update-ADReportResults -Results $ReportData 
        })
    } else {
        Write-ADReportLog -Message "GUI Element 'ButtonQuickAdminUsers' nicht gefunden. Event-Handler nicht zugewiesen." -Type Warning
    }

    # Event Handler für ButtonQuickRecentlyModifiedUsers
    if ($null -ne $Global:ButtonQuickRecentlyModifiedUsers) {
        $Global:ButtonQuickRecentlyModifiedUsers.add_Click({
            Write-ADReportLog -Message "Button 'Kürzlich geänderte Benutzer' geklickt." -Type Info
            Initialize-ResultCounters # Zähler und UI zurücksetzen
            $ReportData = Get-ADReportRecentlyModifiedUsers # Standardmäßig letzte 7 Tage
            Update-ADReportResults -Results $ReportData 
        })
    } else {
        Write-ADReportLog -Message "GUI Element 'ButtonQuickRecentlyModifiedUsers' nicht gefunden. Event-Handler nicht zugewiesen." -Type Warning
    }

    # Event Handler für ButtonQuickUsersWBadPwdCount
    if ($null -ne $Global:ButtonQuickUsersWBadPwdCount) {
        $Global:ButtonQuickUsersWBadPwdCount.add_Click({
            Write-ADReportLog -Message "Button 'Benutzer mit fehlerhaften Kennworteingaben' geklickt." -Type Info
            Initialize-ResultCounters # Zähler und UI zurücksetzen
            $ReportData = Get-ADReportUsersWBadPwdCount
            Update-ADReportResults -Results $ReportData 
        })
    } else {
        Write-ADReportLog -Message "GUI Element 'ButtonQuickUsersWBadPwdCount' nicht gefunden. Event-Handler nicht zugewiesen." -Type Warning
    }

    # Event Handler für ButtonQuickAllGroups
    if ($null -ne $Global:ButtonQuickGroups) {
        $Global:ButtonQuickGroups.add_Click({
            Write-ADReportLog -Message "Button 'Alle Gruppen' geklickt." -Type Info
            Initialize-ResultCounters # Zähler und UI zurücksetzen
            $ReportData = Get-ADReportAllGroups
            Update-ADReportResults -Results $ReportData 
        })
    } else {
        Write-ADReportLog -Message "GUI Element 'ButtonQuickGroups' (for All Groups) nicht gefunden. Event-Handler nicht zugewiesen." -Type Warning
    }

    # Event Handler für ButtonQuickEmptyGroups
    if ($null -ne $Global:ButtonQuickEmptyGroups) {
        $Global:ButtonQuickEmptyGroups.add_Click({
            Write-ADReportLog -Message "Button 'Leere Gruppen' geklickt." -Type Info
            Initialize-ResultCounters # Zähler und UI zurücksetzen
            $ReportData = Get-ADReportEmptyGroups
            Update-ADReportResults -Results $ReportData -StatusMessage "Leere Gruppen Report geladen."
        })
    } else {
        Write-ADReportLog -Message "GUI Element 'ButtonQuickEmptyGroups' nicht gefunden. Event-Handler nicht zugewiesen." -Type Warning
    }

    # Event Handler für ButtonQuickGroupsWNoManager
    if ($null -ne $Global:ButtonQuickGroupsWNoManager) {
        $Global:ButtonQuickGroupsWNoManager.add_Click({
            Write-ADReportLog -Message "Button 'Gruppen ohne Manager' geklickt." -Type Info
            Initialize-ResultCounters # Zähler und UI zurücksetzen
            $ReportData = Get-ADReportGroupsWNoManager
            Update-ADReportResults -Results $ReportData 
        })
    } else {
        Write-ADReportLog -Message "GUI Element 'ButtonQuickGroupsWNoManager' nicht gefunden. Event-Handler nicht zugewiesen." -Type Warning
    }

    # Event Handler for a DNS Button (Button does not exist in XAML)
    # $Global:ButtonQuickDNSRecords.add_Click({
    #     Write-ADReportLog -Message "Button 'Alle DNS-Records' geklickt." -Type Info -Terminal
    #     Initialize-ResultCounters # Zähler und UI zurücksetzen
    #     $ReportData = Get-ADReportAllDNSRecords
    #     Update-ADReportResults -Results $ReportData -StatusMessage "Alle DNS-Records"
    # })

    $Global:ButtonQuickSitesSubnets.Add_Click({
        Write-ADReportLog -Message "Loading AD Sites & Subnets Report..." -Type Info -Terminal
        Initialize-ResultCounters
        $ReportData = Get-ADSitesAndSubnetsReport
        Update-ADReportResults -Results $ReportData -StatusMessage "Sites & Subnets Report loaded."
    })

    # --- UI Elemente füllen ---
    # Standardattribute für die Auswahl-ListBox füllen
    if ($null -ne $Global:ListBoxSelectAttributes) {
        $DefaultAttributes = @("DisplayName", "SamAccountName", "GivenName", "Surname", "mail", "Department", "Title", "Enabled", "LastLogonDate", "whenCreated")
        try {
            $Global:ListBoxSelectAttributes.Items.Clear()
            $DefaultAttributes | ForEach-Object {
                $listBoxItem = New-Object System.Windows.Controls.ListBoxItem
                $listBoxItem.Content = $_ 
                $Global:ListBoxSelectAttributes.Items.Add($listBoxItem)
                if ($_ -in @("DisplayName", "SamAccountName", "Enabled")) { # Standardmäßig ausgewählte Attribute
                    $Global:ListBoxSelectAttributes.SelectedItems.Add($listBoxItem)
                }
            }
        } catch {
            Write-Warning "Error initializing ListBoxSelectAttributes: $($_.Exception.Message)"
        }
    } else {
        Write-Warning "ListBoxSelectAttributes not found in XAML"
    }

    # Standardattribute für die Filter-ComboBox füllen
    if ($null -ne $Global:ComboBoxFilterAttribute) {
        try {
            $FilterableAttributes = @("DisplayName", "SamAccountName", "GivenName", "Surname", "mail", "Department", "Title")
            $Global:ComboBoxFilterAttribute.Items.Clear()
            $FilterableAttributes | ForEach-Object { $Global:ComboBoxFilterAttribute.Items.Add($_) }
            if ($Global:ComboBoxFilterAttribute.Items.Count -gt 0) { 
                $Global:ComboBoxFilterAttribute.SelectedIndex = 0 
            }
        } catch {
            Write-Warning "Error initializing ComboBoxFilterAttribute: $($_.Exception.Message)"
        }
    } else {
        Write-Warning "ComboBoxFilterAttribute not found in XAML"
    }

    # Sicherstellen, dass alle Zähler zurückgesetzt werden
    if ($null -ne $Global:UserCountText) {
        $Global:UserCountText.Text = "0"
    }
    
    if ($null -ne $Global:ComputerCountText) {
        $Global:ComputerCountText.Text = "0"
        $Global:DataGridResults.ItemsSource = $null
    }
}

# Funktion zur Aktualisierung der Ergebnisanzeige
Function Update-ResultVisualization {
    param (
        [Parameter(Mandatory=$true)]
        [array]$Results
    )
    
    try {
        # Analysieren der Suchergebnisse basierend auf dem Objekttyp
        # Standardwerte für die Visualisierungen
        $userData = @{"Total" = 0}
        $computerData = @{"Total" = 0}
        $groupData = @{"Total" = 0}
        
        # Farben für die Diagramme
        $userColors = @{
            "Enabled" = "#4CAF50";
            "Disabled" = "#F44336";
            "Locked" = "#FFC107";
            "Admin" = "#2196F3";
            "Total" = "#607D8B";
            "Default" = "#9E9E9E"
        }
        
        $computerColors = @{
            "Enabled" = "#4CAF50";
            "Disabled" = "#F44336";
            "Inactive" = "#FFC107";
            "Total" = "#607D8B";
            "Default" = "#9E9E9E"
        }
        
        $groupColors = @{
            "Security" = "#FF9800";
            "Distribution" = "#9C27B0";
            "Total" = "#607D8B";
            "Default" = "#9E9E9E"
        }
        
        # Objekttyp bestimmen basierend auf Eigenschaften des ersten Elements, falls vorhanden
        if ($Results.Count -gt 0) {
            $firstObject = $Results[0]
            
            # Überprüfen auf Benutzer, Computer oder Gruppe basierend auf ObjectClass Eigenschaft
            if ($firstObject.PSObject.Properties.Name -contains "ObjectClass") {
                $objectClass = $firstObject.ObjectClass
                
                if ($objectClass -eq "user") {
                    # Benutzerdetails analysieren
                    $enabledUsers = $Results | Where-Object {$_.Enabled -eq $true}
                    $disabledUsers = $Results | Where-Object {$_.Enabled -eq $false}
                    
                    $userData = @{
                        "Total" = $Results.Count
                    }
                    
                    if ($enabledUsers.Count -gt 0) {
                        $userData["Enabled"] = $enabledUsers.Count
                    }
                    
                    if ($disabledUsers.Count -gt 0) {
                        $userData["Disabled"] = $disabledUsers.Count
                    }
                    
                    # Wenn die Eigenschaften LockedOut oder AdminCount vorhanden sind, diese auch berücksichtigen
                    if ($firstObject.PSObject.Properties.Name -contains "LockedOut") {
                        $lockedUsers = $Results | Where-Object {$_.LockedOut -eq $true}
                        if ($lockedUsers.Count -gt 0) {
                            $userData["Locked"] = $lockedUsers.Count
                        }
                    }
                    
                    if ($firstObject.PSObject.Properties.Name -contains "AdminCount") {
                        $adminUsers = $Results | Where-Object {$_.AdminCount -eq 1}
                        if ($adminUsers.Count -gt 0) {
                            $userData["Admin"] = $adminUsers.Count
                        }
                    }
                }
                elseif ($objectClass -eq "computer") {
                    # Computerdetails analysieren
                    $enabledComputers = $Results | Where-Object {$_.Enabled -eq $true}
                    $disabledComputers = $Results | Where-Object {$_.Enabled -eq $false}
                    
                    $computerData = @{
                        "Total" = $Results.Count
                    }
                    
                    if ($enabledComputers.Count -gt 0) {
                        $computerData["Enabled"] = $enabledComputers.Count
                    }
                    
                    if ($disabledComputers.Count -gt 0) {
                        $computerData["Disabled"] = $disabledComputers.Count
                    }
                    
                    # Wenn InactiveDays vorhanden ist, auch inaktive Computer anzeigen
                    if ($firstObject.PSObject.Properties.Name -contains "InactiveDays") {
                        $inactiveComputers = $Results | Where-Object {$_.InactiveDays -gt 30}
                        if ($inactiveComputers.Count -gt 0) {
                            $computerData["Inactive"] = $inactiveComputers.Count
                        }
                    }
                }
                elseif ($objectClass -eq "group") {
                    # Gruppendetails analysieren
                    if ($firstObject.PSObject.Properties.Name -contains "GroupCategory") {
                        $securityGroups = $Results | Where-Object {$_.GroupCategory -eq "Security"}
                        $distributionGroups = $Results | Where-Object {$_.GroupCategory -eq "Distribution"}
                        
                        $groupData = @{
                            "Total" = $Results.Count
                        }
                        
                        if ($securityGroups.Count -gt 0) {
                            $groupData["Security"] = $securityGroups.Count
                        }
                        
                        if ($distributionGroups.Count -gt 0) {
                            $groupData["Distribution"] = $distributionGroups.Count
                        }
                    }
                    else {
                        # Wenn GroupCategory nicht verfügbar ist, nur Gesamtzahl anzeigen
                        $groupData = @{
                            "Total" = $Results.Count
                        }
                    }
                }
                else {
                    # Für andere Objekttypen nur Gesamtzahl anzeigen
                    $userData = @{"Total" = $Results.Count}
                }
            }
            else {
                # Wenn keine ObjectClass-Eigenschaft vorhanden ist, zeige nur die Gesamtzahl an
                $userData = @{"Total" = $Results.Count}
            }
        }
        
        # Visualisierungen erstellen mit Null-Check
        # Canvas-Elemente sind nicht im aktuellen XAML definiert - Platzhalter für zukünftige Implementierung
        Write-ADReportLog -Message "Result visualization data prepared (visualization UI not yet implemented)." -Type Info -Terminal
        
        Write-ADReportLog -Message "Result visualization updated successfully." -Type Info -Terminal
    }
    catch {
        Write-ADReportLog -Message "Error updating result visualization: $($_.Exception.Message)" -Type Error
    }
}



    # --- Event Handler zuweisen --- 
    # Event Handler für RadioButtons zum Umschalten zwischen Objekttypen
    # Event Handler für ListBoxSelectAttributes zur Steuerung der Attributauswahl
    # Define the SelectionChanged event handler script block
    $script:ListBoxSelectionChangedHandler = {
        param($sourceControl, $e)

        # Get all items from the ListBox.
        $allItems = $Global:ListBoxSelectAttributes.Items
        # Get currently selected items. These are expected to be ListBoxItem objects.
        $selectedListBoxItems = $Global:ListBoxSelectAttributes.SelectedItems

        $memberOfSelected = $false
        if ($null -ne $selectedListBoxItems) {
            foreach ($selectedItem_obj in $selectedListBoxItems) {
                if ($selectedItem_obj.Content -eq "MemberOf") {
                    $memberOfSelected = $true
                    break
                }
            }
        }

        if ($memberOfSelected) {
            # "MemberOf" is selected. Disable all other items.
            foreach ($item_iter in $allItems) {

                if ($item_iter.Content -ne "MemberOf") {
                    $item_iter.IsEnabled = $false
                } else {
                    $item_iter.IsEnabled = $true # Ensure "MemberOf" itself remains enabled
                }
            }
        } else {
            # "MemberOf" is NOT selected.
            $otherAttributeSelected = $false
            if ($null -ne $selectedListBoxItems -and $selectedListBoxItems.Count -gt 0) {
                $otherAttributeSelected = $true
            }

            if ($otherAttributeSelected) {
                # Other attributes are selected. Disable "MemberOf". Enable all others.
                foreach ($item_iter in $allItems) {

                    if ($item_iter.Content -eq "MemberOf") {
                        $item_iter.IsEnabled = $false
                        if ($item_iter.IsSelected) {
                            $item_iter.IsSelected = $false
                        }
                    } else {
                        $item_iter.IsEnabled = $true
                    }
                }
            } else {
                # No attributes are selected (or only "MemberOf" was selected and now it's not).
                # Enable all attributes.
                foreach ($item_iter in $allItems) {
                    if ($item_iter -isnot [System.Windows.Controls.ListBoxItem]) {
                        Write-Warning "SelectionChanged: Item '$($item_iter)' (no selection) is not a ListBoxItem. Skipping."
                        continue
                    }
                    $item_iter.IsEnabled = $true
                }
            }
        }
    }
    # Add the event handler
    $Global:ListBoxSelectAttributes.add_SelectionChanged($script:ListBoxSelectionChangedHandler)

    $RadioButtonUser.add_Checked({
        $Global:ListBoxSelectAttributes.IsEnabled = $true
        # Funktion wird ausgeführt, wenn RadioButtonUser ausgewählt wird
        Write-ADReportLog -Message "Object type changed to User" -Type Info -Terminal
        
        # ComboBoxFilterAttribute leeren und mit benutzerspezifischen Attributen füllen
        $Global:ComboBoxFilterAttribute.Items.Clear()
        $UserFilterAttributes = @("DisplayName", "SamAccountName", "GivenName", "Surname", "mail", 
                                  "Department", "Title", "EmployeeID", "EmployeeNumber", "UserPrincipalName")
        $UserFilterAttributes | ForEach-Object { $Global:ComboBoxFilterAttribute.Items.Add($_) }
        if ($Global:ComboBoxFilterAttribute.Items.Count -gt 0) { 
            $Global:ComboBoxFilterAttribute.SelectedIndex = 0 
        }
        
        # ListBoxSelectAttributes leeren und mit benutzerspezifischen Attributen füllen
        $Global:ListBoxSelectAttributes.Items.Clear()
        $UserAttributes = @("DisplayName", "SamAccountName", "GivenName", "Surname", "mail", 
                            "Department", "Title", "Enabled", "LastLogonDate", "whenCreated", 
                            "PasswordLastSet", "PasswordNeverExpires", "LockedOut", "Description",
                            "Office", "OfficePhone", "MobilePhone", "Company")
        $UserAttributes | ForEach-Object {
            $newItem = New-Object System.Windows.Controls.ListBoxItem
            $newItem.Content = $_ 
            $Global:ListBoxSelectAttributes.Items.Add($newItem)
        }
        
        # Status aktualisieren
        $Global:TextBlockStatus.Text = "Bereit für Benutzerabfrage"
    })

    $RadioButtonGroup.add_Checked({
        $Global:ListBoxSelectAttributes.IsEnabled = $true
        # Funktion wird ausgeführt, wenn RadioButtonGroup ausgewählt wird
        Write-ADReportLog -Message "Object type changed to Group" -Type Info -Terminal
        
        # ComboBoxFilterAttribute leeren und mit gruppenspezifischen Attributen füllen
        $Global:ComboBoxFilterAttribute.Items.Clear()
        $GroupFilterAttributes = @("Name", "SamAccountName", "Description", "GroupCategory", "GroupScope")
        $GroupFilterAttributes | ForEach-Object { $Global:ComboBoxFilterAttribute.Items.Add($_) }
        if ($Global:ComboBoxFilterAttribute.Items.Count -gt 0) { 
            $Global:ComboBoxFilterAttribute.SelectedIndex = 0 
        }
        
        # ListBoxSelectAttributes leeren und mit gruppenspezifischen Attributen füllen
        $Global:ListBoxSelectAttributes.Items.Clear()
        $GroupAttributes = @("Name", "SamAccountName", "Description", "GroupCategory", "GroupScope", 
                            "whenCreated", "whenChanged", "ManagedBy", "mail", "info", "MemberOf")
        $GroupAttributes | ForEach-Object {
            $newItem = New-Object System.Windows.Controls.ListBoxItem
            $newItem.Content = $_ 
            $Global:ListBoxSelectAttributes.Items.Add($newItem)
        }
        
        # Status aktualisieren
        $Global:TextBlockStatus.Text = "Bereit für Gruppenabfrage"
    })

    $RadioButtonComputer.add_Checked({
        $Global:ListBoxSelectAttributes.IsEnabled = $true
        # Funktion wird ausgeführt, wenn RadioButtonComputer ausgewählt wird
        Write-ADReportLog -Message "Object type changed to Computer" -Type Info -Terminal
        
        # ComboBoxFilterAttribute leeren und mit computerspezifischen Attributen füllen
        $Global:ComboBoxFilterAttribute.Items.Clear()
        $ComputerFilterAttributes = @("Name", "DNSHostName", "OperatingSystem", "Description")
        $ComputerFilterAttributes | ForEach-Object { $Global:ComboBoxFilterAttribute.Items.Add($_) }
        if ($Global:ComboBoxFilterAttribute.Items.Count -gt 0) { 
            $Global:ComboBoxFilterAttribute.SelectedIndex = 0 
        }
        
        # ListBoxSelectAttributes leeren und mit computerspezifischen Attributen füllen
        $Global:ListBoxSelectAttributes.Items.Clear()
        $ComputerAttributes = @("Name", "DNSHostName", "OperatingSystem", "OperatingSystemVersion", 
                              "Enabled", "LastLogonDate", "whenCreated", "IPv4Address", "Description",
                              "Location", "ManagedBy", "PasswordLastSet")
        $ComputerAttributes | ForEach-Object {
            $newItem = New-Object System.Windows.Controls.ListBoxItem
            $newItem.Content = $_ 
            $Global:ListBoxSelectAttributes.Items.Add($newItem)
        }
        
        # Status aktualisieren
        $Global:TextBlockStatus.Text = "Bereit für Computerabfrage"
    })

    $RadioButtonGroupMemberships.add_Checked({
        Write-ADReportLog -Message "Object type changed to GroupMemberships" -Type Info -Terminal
        $Global:ListBoxSelectAttributes.IsEnabled = $false
        # Optional: ComboBoxFilterAttribute anpassen oder leeren, falls erforderlich
        # $Global:ComboBoxFilterAttribute.Items.Clear()
        # $Global:ComboBoxFilterAttribute.Items.Add("SamAccountName") # Beispiel
        # if ($Global:ComboBoxFilterAttribute.Items.Count -gt 0) { $Global:ComboBoxFilterAttribute.SelectedIndex = 0 }
        Write-ADReportLog -Message "Attribute selection disabled for GroupMemberships query." -Type Info
    })

    # Event Handler für ButtonQueryAD anpassen, um Objekttyp zu berücksichtigen
    $ButtonQueryAD.add_Click({
        Write-ADReportLog -Message "Executing query..." -Type Info
        try {
            $SelectedFilterAttribute = $Global:ComboBoxFilterAttribute.SelectedItem.ToString()
            $FilterValue = $Global:TextBoxFilterValue.Text
            $SelectedAttributes = $Global:ListBoxSelectAttributes.SelectedItems
            Write-Host "DEBUG: SelectedItems in ButtonClickHandler: $($Global:ListBoxSelectAttributes.SelectedItems | ForEach-Object {$_.Content -join '; '})"
            $isUserSearch = $Global:RadioButtonUser.IsChecked

            if ($SelectedAttributes.Count -eq 0) {
                [System.Windows.MessageBox]::Show("Please select at least one attribute for export.", "Warnung", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
                return
            }

            # Bestimme den aktuell ausgewählten Objekttyp
            $ObjectType = if ($Global:RadioButtonUser.IsChecked) { "User" } 
                        elseif ($Global:RadioButtonGroup.IsChecked) { "Group" } 
                        elseif ($Global:RadioButtonGroupMemberships.IsChecked) { "GroupMemberships" } 
                        else { "Computer" }
            
            # AD-Abfrage basierend auf Objekttyp durchführen
            $ReportData = $null
            switch ($ObjectType) {
                "User" {
                    $ReportData = Get-ADReportData -FilterAttribute $SelectedFilterAttribute -FilterValue $FilterValue -SelectedAttributes $SelectedAttributes
                }
                "Group" {
                    # Für Gruppen wird der Filter erst nach dem Abrufen aller Gruppen angewendet
                    $ReportData = Get-ADGroupReportData -CustomFilter "*" -SelectedAttributes $SelectedAttributes
                    if ($FilterValue -and $SelectedFilterAttribute) {
                        $ReportData = $ReportData | Where-Object { $_.$SelectedFilterAttribute -like "*$FilterValue*" }
                    }
                }
                "Computer" {
                    # Für Computer wird der Filter erst nach dem Abrufen aller Computer angewendet
                    $ReportData = Get-ADComputerReportData -CustomFilter "*" -SelectedAttributes $SelectedAttributes
                    if ($FilterValue -and $SelectedFilterAttribute) {
                        $ReportData = $ReportData | Where-Object { $_.$SelectedFilterAttribute -like "*$FilterValue*" }
                    }
                }
                "GroupMemberships" {
                    if (-not ([string]::IsNullOrWhiteSpace($FilterValue)) -and -not ([string]::IsNullOrWhiteSpace($SelectedFilterAttribute))) {
                        $ReportData = Get-ADGroupMembershipsReportData -FilterAttribute $SelectedFilterAttribute -FilterValue $FilterValue
                    } else {
                        Write-ADReportLog -Message "Filter attribute or value is empty for GroupMemberships query. Please specify a filter." -Type Warning
                        [System.Windows.Forms.MessageBox]::Show("Bitte geben Sie einen Filter (Attribut und Wert) für die Mitgliedschaftsabfrage an.", "Hinweis", "OK", "Information")
                        $ReportData = @()
                    }
                }
            }
            
            if ($ReportData) {
                try { # Inner try for processing $ReportData
                    if ($Global:DataGridResults) {
                        $Global:DataGridResults.ItemsSource = $null # Clear items first
                        if ($Global:DataGridResults.Columns) { $Global:DataGridResults.Columns.Clear() } # Clear columns
                    }
                    # Debug-Informationen
                    Write-ADReportLog -Message "ReportData Typ: $($ReportData.GetType().FullName)" -Type Info -Terminal
                    
                    # Wir brauchen sicherzustellen, dass wir immer eine Liste/Sammlung haben, auch bei einzelnen Objekten
                    # Verwende @() um es als Array zu erzwingen
                    $ReportCollection = @($ReportData)
                    
                    # Direkte Zuweisung an DataGrid
                    $Global:DataGridResults.ItemsSource = $ReportCollection
                    
                    # Zähle die Anzahl der Ergebnisse
                    $Count = $ReportCollection.Count
                    Write-ADReportLog -Message "Query completed. $Count result(s) found for $ObjectType." -Type Info
                    
                    # Ergebniszähler im Header aktualisieren
                    Update-ResultCounters -Results $ReportCollection
                    
                    if ($isUserSearch -and $ReportCollection.Count -eq 1 -and $ReportCollection[0].PSObject.Properties['SamAccountName']) {
                        $userSamAccountName = $ReportCollection[0].SamAccountName
                        
                        # Check if the "Mitgliedschaften" RadioButton is checked
                        if ($Global:RadioButtonGroupMemberships.IsChecked -eq $true) {
                            Write-ADReportLog -Message "Rufe Gruppenmitgliedschaften für Benutzer $($userSamAccountName) ab (RadioButton 'Mitgliedschaften' ist aktiv)..." -Type Info
                            $GroupMemberships = Get-UserGroupMemberships -SamAccountName $userSamAccountName
                            
                            if ($GroupMemberships -and $GroupMemberships.Count -gt 0) {
                                Write-ADReportLog -Message "$($GroupMemberships.Count) Gruppenmitgliedschaften gefunden." -Type Info
                                $DisplayData = $GroupMemberships | Select-Object @{
                                    Name = 'Benutzer';
                                    Expression = {$_.UserDisplayName}
                                }, @{
                                    Name = 'Benutzer (SAM)';
                                    Expression = {$_.UserSamAccountName}
                                }, @{
                                    Name = 'Gruppe';
                                    Expression = {$_.GroupName}
                                }, @{
                                    Name = 'Gruppe (SAM)';
                                    Expression = {$_.GroupSamAccountName}
                                }, @{
                                    Name = 'Gruppentyp';
                                    Expression = {"$($_.GroupCategory) / $($_.GroupScope)"}
                                }, @{
                                    Name = 'Beschreibung';
                                }
                            } else {
                                # RadioButton "Mitgliedschaften" is NOT checked. Display user data as usual.
                                Write-ADReportLog -Message "RadioButton 'Mitgliedschaften' ist nicht aktiv. Zeige Benutzerdaten für $userSamAccountName." -Type Info
                                $Global:DataGridResults.ItemsSource = $ReportCollection
                                $Global:TextBlockStatus.Text = "Benutzer $userSamAccountName gefunden. Mitgliedschaften nicht abgefragt."
                                Update-ResultCounters -Results $ReportCollection
                                Update-ResultVisualization -Results $ReportCollection
                            }
                        } else {
                            # RadioButton "Mitgliedschaften" is NOT checked. Display user data as usual.
                            Write-ADReportLog -Message "RadioButton 'Mitgliedschaften' ist nicht aktiv. Zeige Benutzerdaten für $userSamAccountName." -Type Info
                            $Global:DataGridResults.ItemsSource = $ReportCollection
                            $Global:TextBlockStatus.Text = "Benutzer $userSamAccountName gefunden. Mitgliedschaften nicht abgefragt."
                            Update-ResultCounters -Results $ReportCollection
                            Update-ResultVisualization -Results $ReportCollection
                        }
                    } else {
                        # This is the original 'else' for cases other than single user search.
                        # It should remain as is, displaying the $ReportCollection.
                        $Global:DataGridResults.ItemsSource = $ReportCollection
                        Update-ResultCounters -Results $ReportCollection
                        Update-ResultVisualization -Results $ReportCollection # Ensure visualization is updated here too.
                    }
                } catch {
                    $ErrorMessage = "Error in query process: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    if ($Global:DataGridResults) {
                        $Global:DataGridResults.ItemsSource = $null
                        if ($Global:DataGridResults.Columns) { $Global:DataGridResults.Columns.Clear() }
                    }
                    Update-ResultCounters -Results @() # Leeres Array für die Zähler
                } 
            } else {
                Write-ADReportLog -Message "No data returned from query for $ObjectType." -Type Info
                if ($Global:DataGridResults) {
                    $Global:DataGridResults.ItemsSource = $null
                    if ($Global:DataGridResults.Columns) { $Global:DataGridResults.Columns.Clear() }
                }
                Update-ResultCounters -Results @()
                if ($Global:TextBlockStatus) { $Global:TextBlockStatus.Text = "No results found for $ObjectType." }
            }
        } catch { # Outer catch for the whole ButtonQueryAD.add_Click
            $OuterCatchErrorMessage = "An unexpected error occurred during query execution: $($_.Exception.Message)"
            Write-ADReportLog -Message $OuterCatchErrorMessage -Type Error
            try {
                if ($Global:DataGridResults) {
                    $Global:DataGridResults.ItemsSource = $null
                    if ($Global:DataGridResults.Columns) { $Global:DataGridResults.Columns.Clear() }
                }
                Update-ResultCounters -Results @()
                if ($Global:TextBlockStatus) { $Global:TextBlockStatus.Text = "Error: Query failed." } 
            } catch {
                Write-ADReportLog -Message "CRITICAL: Error within the main query button's outer catch block: $($_.Exception.Message)" -Type Error
            }
        }
    }) # End of .add_Click({
            $ButtonQuickAllUsers.add_Click({
        Write-ADReportLog -Message "Loading all users..." -Type Info
        try {
            if ($script:ListBoxSelectionChangedHandler) {
                $Global:ListBoxSelectAttributes.remove_SelectionChanged($script:ListBoxSelectionChangedHandler)
            }
            $Global:RadioButtonUser.IsChecked = $true # Sicherstellen, dass Benutzermodus aktiv ist
            $Global:ComboBoxFilterAttribute.SelectedIndex = -1
            $Global:TextBoxFilterValue.Text = ""
            
            $QuickReportAttributes = @("DisplayName", "SamAccountName", "GivenName", "Surname", "mail", "Enabled", "LastLogonDate", "whenCreated", "LockedOut")
            $Global:ListBoxSelectAttributes.SelectedItems.Clear()
            foreach ($attr in $QuickReportAttributes) {
                $itemsFound = $Global:ListBoxSelectAttributes.Items | Where-Object { $_.Content -eq $attr }
                if ($itemsFound) {
                    foreach ($singleItemInFoundList in $itemsFound) {
                        if ($singleItemInFoundList -is [System.Windows.Controls.ListBoxItem]) {
                            $singleItemInFoundList.IsSelected = $true
                        }
                    }
                }
            }

            $ReportData = Get-ADReportData -CustomFilter "*" -SelectedAttributes $QuickReportAttributes
            if ($null -eq $ReportData) { $ReportData = @() }
            
            $Global:DataGridResults.ItemsSource = $null
            $Global:DataGridResults.Columns.Clear()
            $Global:DataGridResults.ItemsSource = $ReportData
            Update-ResultCounters -Results $ReportData
            $Global:TextBlockStatus.Text = "All users loaded. $($ReportData.Count) record(s) found."
            Write-ADReportLog -Message "All users loaded. $($ReportData.Count) result(s) found." -Type Info
        } catch {
            $ErrorMessage = "Error loading all users: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
            $Global:DataGridResults.Columns.Clear()
            Update-ResultCounters -Results @()
            $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
        }
    })

    $ButtonQuickDisabledUsers.add_Click({
        Write-ADReportLog -Message "Loading disabled users..." -Type Info
        try {
            if ($script:ListBoxSelectionChangedHandler) {
                $Global:ListBoxSelectAttributes.remove_SelectionChanged($script:ListBoxSelectionChangedHandler)
            }
            $Global:RadioButtonUser.IsChecked = $true
            $Global:ComboBoxFilterAttribute.SelectedIndex = -1
            $Global:TextBoxFilterValue.Text = ""

            $QuickReportAttributes = @("DisplayName", "SamAccountName", "Enabled", "LastLogonDate")
            $Global:ListBoxSelectAttributes.SelectedItems.Clear()
            foreach ($attr in $QuickReportAttributes) {
                $itemsFound = 1 # DEBUG
                if ($itemsFound) {
                    foreach ($singleItemInFoundList in $itemsFound) {
                        if ($singleItemInFoundList -is [System.Windows.Controls.ListBoxItem]) {
                            $singleItemInFoundList.IsSelected = $true
                        }
                    }
                }
            }
            
            $ReportData = Get-ADReportData -CustomFilter "Enabled -eq `$false" -SelectedAttributes $QuickReportAttributes
            if ($null -eq $ReportData) { $ReportData = @() }
            $Global:DataGridResults.ItemsSource = $null
            $Global:DataGridResults.Columns.Clear()
            $Global:DataGridResults.ItemsSource = $ReportData
            Update-ResultCounters -Results $ReportData
            $Global:TextBlockStatus.Text = "Disabled users loaded. $($ReportData.Count) record(s) found."
            Write-ADReportLog -Message "Disabled users loaded. $($ReportData.Count) result(s) found." -Type Info
        } catch {
            $ErrorMessage = "Error loading disabled users: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
            $Global:DataGridResults.Columns.Clear()
            Update-ResultCounters -Results @()
            $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
        }
    })

    $ButtonQuickLockedUsers.add_Click({
        Write-ADReportLog -Message "Loading locked out users..." -Type Info
        try {
            if ($script:ListBoxSelectionChangedHandler) {
                $Global:ListBoxSelectAttributes.remove_SelectionChanged($script:ListBoxSelectionChangedHandler)
            }
            $Global:RadioButtonUser.IsChecked = $true
            $Global:ComboBoxFilterAttribute.SelectedIndex = -1
            $Global:TextBoxFilterValue.Text = ""

            $QuickReportAttributes = @("DisplayName", "SamAccountName", "LockedOut", "LastLogonDate", "BadLogonCount")
            $Global:ListBoxSelectAttributes.SelectedItems.Clear()
            foreach ($attr in $QuickReportAttributes) {
                $itemsFound = $Global:ListBoxSelectAttributes.Items | Where-Object { $_.Content -eq $attr }
                if ($itemsFound) {
                    foreach ($singleItemInFoundList in $itemsFound) {
                        if ($singleItemInFoundList -is [System.Windows.Controls.ListBoxItem]) {
                            $singleItemInFoundList.IsSelected = $true
                        }
                    }
                }
            }

            $ReportData = Get-ADReportData -CustomFilter "LockedOut -eq `$true" -SelectedAttributes $QuickReportAttributes
            if ($null -eq $ReportData) { $ReportData = @() }
            $Global:DataGridResults.ItemsSource = $null
            $Global:DataGridResults.Columns.Clear()
            $Global:DataGridResults.ItemsSource = $ReportData
            Update-ResultCounters -Results $ReportData
            $Global:TextBlockStatus.Text = "Locked out users loaded. $($ReportData.Count) record(s) found."
            Write-ADReportLog -Message "Locked out users loaded. $($ReportData.Count) result(s) found." -Type Info
        } catch {
            $ErrorMessage = "Error loading locked out users: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
            $Global:DataGridResults.Columns.Clear()
            Update-ResultCounters -Results @()
            $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
        }
    })

    $ButtonQuickGroups.add_Click({
        Write-ADReportLog -Message "Loading all groups..." -Type Info
        try {
            if ($script:ListBoxSelectionChangedHandler) {
                $Global:ListBoxSelectAttributes.remove_SelectionChanged($script:ListBoxSelectionChangedHandler)
            }
            $Global:RadioButtonGroup.IsChecked = $true
            $Global:ComboBoxFilterAttribute.SelectedIndex = -1
            $Global:TextBoxFilterValue.Text = ""

            $QuickReportAttributes = @("Name", "SamAccountName", "GroupCategory", "GroupScope")
            $Global:ListBoxSelectAttributes.SelectedItems.Clear()
            foreach ($attr in $QuickReportAttributes) {
                $itemsFound = $Global:ListBoxSelectAttributes.Items | Where-Object { $_.Content -eq $attr }
                if ($itemsFound) {
                    foreach ($singleItemInFoundList in $itemsFound) {
                        if ($singleItemInFoundList -is [System.Windows.Controls.ListBoxItem]) {
                            $singleItemInFoundList.IsSelected = $true
                        }
                    }
                }
            }

            $ReportData = Get-ADGroupReportData -CustomFilter "*" -SelectedAttributes $QuickReportAttributes
            if ($null -eq $ReportData) { $ReportData = @() }
            $Global:DataGridResults.ItemsSource = $null
            $Global:DataGridResults.Columns.Clear()
            $Global:DataGridResults.ItemsSource = $ReportData
            Update-ResultCounters -Results $ReportData
            $Global:TextBlockStatus.Text = "All groups loaded. $($ReportData.Count) record(s) found."
            Write-ADReportLog -Message "All groups loaded. $($ReportData.Count) result(s) found." -Type Info
        } catch {
            $ErrorMessage = "Error loading all groups: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
            $Global:DataGridResults.Columns.Clear()
            Update-ResultCounters -Results @()
            $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
        }
    })

    $ButtonQuickSecurityGroups.add_Click({
        Write-ADReportLog -Message "Loading security groups..." -Type Info
        try {
            $QuickReportAttributes = @("Name", "SamAccountName", "GroupCategory", "GroupScope")
            $ReportData = Get-ADGroupReportData -CustomFilter "GroupCategory -eq 'Security'" -SelectedAttributes $QuickReportAttributes
            if ($ReportData) {
                $Global:DataGridResults.ItemsSource = $null
                $Global:DataGridResults.Columns.Clear()
                $Global:DataGridResults.ItemsSource = $ReportData
                Write-ADReportLog -Message "Security groups loaded. $($ReportData.Count) result(s) found." -Type Info
                Update-ResultCounters -Results $ReportData
            } else {
                $Global:DataGridResults.ItemsSource = $null
                $Global:DataGridResults.Columns.Clear()
            }
        } catch {
            $ErrorMessage = "Error loading security groups: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
            $Global:DataGridResults.Columns.Clear()
        }
    })

    $ButtonQuickDistributionGroups.add_Click({
        Write-ADReportLog -Message "Loading distribution lists..." -Type Info
        try {
            $QuickReportAttributes = @("Name", "SamAccountName", "GroupCategory", "GroupScope")
            $ReportData = Get-ADGroupReportData -CustomFilter "GroupCategory -eq 'Distribution'" -SelectedAttributes $QuickReportAttributes
            if ($ReportData) {
                $Global:DataGridResults.ItemsSource = $null
                $Global:DataGridResults.Columns.Clear()
                $Global:DataGridResults.ItemsSource = $ReportData
                Write-ADReportLog -Message "Distribution lists loaded. $($ReportData.Count) result(s) found." -Type Info
                Update-ResultCounters -Results $ReportData
            } else {
                $Global:DataGridResults.ItemsSource = $null
                $Global:DataGridResults.Columns.Clear()
            }
        } catch {
            $ErrorMessage = "Error loading distribution lists: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
            $Global:DataGridResults.Columns.Clear()
        }
    })

    # Neue Funktionen für Benutzer
    $ButtonQuickNeverExpire.add_Click({
        Write-ADReportLog -Message "Loading users with non-expiring passwords..." -Type Info
        try {
            $QuickReportAttributes = @("DisplayName", "SamAccountName", "Enabled", "PasswordNeverExpires", "LastLogonDate")
            $ReportData = Get-ADReportData -CustomFilter "PasswordNeverExpires -eq `$true" -SelectedAttributes $QuickReportAttributes
            if ($ReportData) {
                $Global:DataGridResults.ItemsSource = $null
                $Global:DataGridResults.Columns.Clear()
                $Global:DataGridResults.ItemsSource = $ReportData
                Write-ADReportLog -Message "Users with non-expiring passwords loaded. $($ReportData.Count) result(s) found." -Type Info
                Update-ResultCounters -Results $ReportData
            } else {
                $Global:DataGridResults.ItemsSource = $null
                $Global:DataGridResults.Columns.Clear()
            }
        } catch {
            $ErrorMessage = "Error loading users with non-expiring passwords: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
            $Global:DataGridResults.Columns.Clear()
        }
    })

    $ButtonQuickInactiveUsers.add_Click({
        Write-ADReportLog -Message "Loading inactive users (>90 days)..." -Type Info
        try {
            $QuickReportAttributes = @("DisplayName", "SamAccountName", "Enabled", "LastLogonDate")
            $InactiveCutoffDate = (Get-Date).AddDays(-90)
            $ReportData = Get-ADReportData -CustomFilter "(LastLogonTimeStamp -lt '$($InactiveCutoffDate.ToFileTime())') -or (LastLogonTimeStamp -notlike '*')" -SelectedAttributes $QuickReportAttributes
            if ($ReportData) {
                # Nachbearbeitung: Nur wirklich inaktive Benutzer behalten
                $FilteredData = $ReportData | Where-Object { 
                    !$_.LastLogonDate -or ($_.LastLogonDate -lt $InactiveCutoffDate)
                }
                $Global:DataGridResults.ItemsSource = $null
                $Global:DataGridResults.Columns.Clear()
                $Global:DataGridResults.ItemsSource = $FilteredData
                Write-ADReportLog -Message "Inactive users loaded. $($FilteredData.Count) result(s) found." -Type Info
                Update-ResultCounters -Results $FilteredData
            } else {
                $Global:DataGridResults.ItemsSource = $null
                $Global:DataGridResults.Columns.Clear()
            }
        } catch {
            $ErrorMessage = "Error loading inactive users: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
            $Global:DataGridResults.Columns.Clear()
        }
    })

    $ButtonQuickAdminUsers.add_Click({
        Write-ADReportLog -Message "Loading administrators..." -Type Info
        try {
            # Verbesserte Methode zum Finden von Admin-Benutzern
            # Zuerst alle Benutzer laden
            $AllUsers = Get-ADUser -Filter * -Properties DisplayName, SamAccountName, Enabled, LastLogonDate | Select-Object -ExcludeProperty 'PropertyNames','AddedProperties','RemovedProperties','ModifiedProperties','PropertiesCount'
            
            # Bekannte Admin-Gruppenbezeichnungen (deutsch und englisch)
            $AdminGroups = @(
                "Domain Admins", "Domänen-Admins",
                "Enterprise Admins", "Organisations-Admins",
                "Administrators", "Administratoren", 
                "Schema Admins", "Schema-Admins",
                "AAD DC Administrators", "Azure AD-DC-Administratoren",
                "Server Operators", "Server-Operatoren",
                "Account Operators", "Konten-Operatoren",
                "Backup Operators", "Sicherungs-Operatoren"
            )
            
            $AdminUsers = @()
            # Prüfe Benutzer auf Admin-Rechte - erst SIDs der Admin-Gruppen ermitteln
            $AdminGroupSIDs = @()
            foreach ($groupName in $AdminGroups) {
                try {
                    $group = Get-ADGroup -Filter "Name -eq '$groupName'" -ErrorAction SilentlyContinue
                    if ($group) {
                        $AdminGroupSIDs += $group.SID.Value
                    }
                } catch {
                    # Ignoriere Fehler bei nicht existierenden Gruppen
                }
            }
            
            # Nur fortfahren, wenn wir Admin-Gruppen gefunden haben
            if ($AdminGroupSIDs.Count -gt 0) {
                # Für jeden Benutzer die Gruppenmitgliedschaften prüfen
                foreach ($user in $AllUsers) {
                    $memberOf = Get-ADPrincipalGroupMembership -Identity $user.SamAccountName -ErrorAction SilentlyContinue
                    foreach ($group in $memberOf) {
                        if ($AdminGroupSIDs -contains $group.SID.Value) {
                            # Benutzer ist Mitglied einer Admin-Gruppe
                            $adminUser = $user | Select-Object DisplayName, SamAccountName, Enabled, LastLogonDate, 
                                @{Name='AdminGroups'; Expression={
                                    ($memberOf | Where-Object { $AdminGroupSIDs -contains $_.SID.Value } | Select-Object -ExpandProperty Name) -join ", "
                                }}
                            $AdminUsers += $adminUser
                            break # Genug, wenn wir eine Admin-Gruppe gefunden haben
                        }
                    }
                }
            }
            
            if ($AdminUsers.Count -gt 0) {
                $Global:DataGridResults.ItemsSource = $null
                $Global:DataGridResults.Columns.Clear()
                $Global:DataGridResults.ItemsSource = $AdminUsers
                Write-ADReportLog -Message "Administrators loaded. $($AdminUsers.Count) result(s) found." -Type Info
                Update-ResultCounters -Results $AdminUsers
            } else {
                $Global:DataGridResults.ItemsSource = $null
                $Global:DataGridResults.Columns.Clear()
                Write-ADReportLog -Message "No administrator accounts found." -Type Warning
            }
        } catch {
            $ErrorMessage = "Error loading administrators: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
            $Global:DataGridResults.Columns.Clear()
        }
    })

    # Neue Funktionen für Computer
    $ButtonQuickComputers.add_Click({
        Write-ADReportLog -Message "Loading all computers..." -Type Info
        try {
            $QuickReportAttributes = @("Name", "DNSHostName", "OperatingSystem", "Enabled", "LastLogonDate")
            $ReportData = Get-ADComputer -Filter * -Properties $QuickReportAttributes | Select-Object $QuickReportAttributes
            if ($ReportData) {
                $Global:DataGridResults.ItemsSource = $null
                $Global:DataGridResults.Columns.Clear()
                $Global:DataGridResults.ItemsSource = $ReportData
                Write-ADReportLog -Message "All computers loaded. $($ReportData.Count) result(s) found." -Type Info
                Update-ResultCounters -Results $ReportData
            } else {
                $Global:DataGridResults.ItemsSource = $null
                $Global:DataGridResults.Columns.Clear()
                Write-ADReportLog -Message "No computers found." -Type Warning
            }
        } catch {
            $ErrorMessage = "Error loading all computers: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
            $Global:DataGridResults.Columns.Clear()
        }
    })

    $ButtonQuickInactiveComputers.add_Click({
        Write-ADReportLog -Message "Loading inactive computers (>90 days)..." -Type Info
        try {
            $QuickReportAttributes = @("Name", "DNSHostName", "OperatingSystem", "Enabled", "LastLogonDate", "PasswordLastSet")
            
            # Direktes Anfordern der spezifischen Computer-Eigenschaften, die wir brauchen
            $AllComputers = Get-ADComputer -Filter * -Properties Name, DNSHostName, OperatingSystem, Enabled, LastLogonDate, LastLogonTimeStamp, PasswordLastSet
            
            # Filtern nach inaktiven Computern basierend auf LastLogonDate und LastLogonTimeStamp
            $InactiveComputers = $AllComputers | Where-Object {
                ($_.LastLogonDate -lt ((Get-Date).AddDays(-90)) -or $_.LastLogonDate -eq $null) -and
                ($_.LastLogonTimeStamp -lt ((Get-Date).AddDays(-90)).ToFileTime() -or $_.LastLogonTimeStamp -eq $null)
            } | Select-Object Name, DNSHostName, OperatingSystem, Enabled, LastLogonDate, PasswordLastSet,
              @{Name='InactiveDays'; Expression={
                if ($_.LastLogonDate) {
                    [math]::Round((New-TimeSpan -Start $_.LastLogonDate -End (Get-Date)).TotalDays)
                } else { "N/A" }
              }}
            
            if ($InactiveComputers -and $InactiveComputers.Count -gt 0) {
                $Global:DataGridResults.ItemsSource = $null
                $Global:DataGridResults.Columns.Clear()
                $Global:DataGridResults.ItemsSource = $InactiveComputers
                Write-ADReportLog -Message "Inactive computers loaded. $($InactiveComputers.Count) result(s) found." -Type Info
                Update-ResultCounters -Results $InactiveComputers
            } else {
                $Global:DataGridResults.ItemsSource = $null
                $Global:DataGridResults.Columns.Clear()
                Write-ADReportLog -Message "No inactive computers found." -Type Warning
            }
        } catch {
            $ErrorMessage = "Error loading inactive computers: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
            $Global:DataGridResults.Columns.Clear()
        }
    })

    # --- Security Audit Event Handler ---
    $ButtonQuickWeakPasswordPolicy.add_Click({
        Write-ADReportLog -Message "Loading users with weak password policies..." -Type Info
        try {
            $WeakPasswordUsers = Get-WeakPasswordPolicyUsers
            if ($WeakPasswordUsers -and $WeakPasswordUsers.Count -gt 0) {
                $Global:DataGridResults.ItemsSource = $null
                $Global:DataGridResults.Columns.Clear()
                $Global:DataGridResults.ItemsSource = $WeakPasswordUsers
                Write-ADReportLog -Message "Users with weak password policies loaded. $($WeakPasswordUsers.Count) result(s) found." -Type Info
                Update-ResultCounters -Results $WeakPasswordUsers
            } else {
                $Global:DataGridResults.ItemsSource = $null
                $Global:DataGridResults.Columns.Clear()
                Write-ADReportLog -Message "No users with weak password policies found." -Type Warning
            }
        } catch {
            $ErrorMessage = "Error loading users with weak password policies: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
            $Global:DataGridResults.Columns.Clear()
        }
    })

    $ButtonQuickRiskyGroupMemberships.add_Click({
        Write-ADReportLog -Message "Loading users with risky group memberships..." -Type Info
        try {
            $RiskyUsers = Get-RiskyGroupMemberships
            if ($RiskyUsers -and $RiskyUsers.Count -gt 0) {
                $Global:DataGridResults.ItemsSource = $RiskyUsers
                Write-ADReportLog -Message "Users with risky group memberships loaded. $($RiskyUsers.Count) result(s) found." -Type Info
                Update-ResultCounters -Results $RiskyUsers
            } else {
                $Global:DataGridResults.ItemsSource = $null
                Write-ADReportLog -Message "No users with risky group memberships found." -Type Warning
            }
        } catch {
            $ErrorMessage = "Error loading users with risky group memberships: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
        }
    })

    $ButtonQuickPrivilegedAccounts.add_Click({
        Write-ADReportLog -Message "Loading privileged accounts..." -Type Info
        try {
            $PrivilegedAccounts = Get-PrivilegedAccounts
            if ($PrivilegedAccounts -and $PrivilegedAccounts.Count -gt 0) {
                $Global:DataGridResults.ItemsSource = $PrivilegedAccounts
                Write-ADReportLog -Message "Privileged accounts loaded. $($PrivilegedAccounts.Count) result(s) found." -Type Info
                Update-ResultCounters -Results $PrivilegedAccounts
            } else {
                $Global:DataGridResults.ItemsSource = $null
                Write-ADReportLog -Message "No privileged accounts found." -Type Warning
            }
        } catch {
            $ErrorMessage = "Error loading privileged accounts: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
        }
    })

    # --- AD-Health Event Handler ---
    $ButtonQuickFSMORoles.add_Click({
        Write-ADReportLog -Message "Loading FSMO role holders..." -Type Info
        try {
            $FSMORoles = Get-FSMORoles
            if ($FSMORoles -and $FSMORoles.Count -gt 0) {
                $Global:DataGridResults.ItemsSource = $FSMORoles
                Write-ADReportLog -Message "FSMO role holders loaded. $($FSMORoles.Count) result(s) found." -Type Info
                Update-ResultCounters -Results $FSMORoles
            } else {
                $Global:DataGridResults.ItemsSource = $null
                Write-ADReportLog -Message "No FSMO role information found." -Type Warning
            }
        } catch {
            $ErrorMessage = "Error loading FSMO role holders: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
        }
    })

    $ButtonQuickDCStatus.add_Click({
        Write-ADReportLog -Message "Loading domain controller status..." -Type Info
        try {
            $DomainControllers = Get-DomainControllerStatus
            if ($DomainControllers -and $DomainControllers.Count -gt 0) {
                $Global:DataGridResults.ItemsSource = $DomainControllers
                Write-ADReportLog -Message "Domain controller status loaded. $($DomainControllers.Count) result(s) found." -Type Info
                Update-ResultCounters -Results $DomainControllers
            } else {
                $Global:DataGridResults.ItemsSource = $null
                Write-ADReportLog -Message "No domain controllers found." -Type Warning
            }
        } catch {
            $ErrorMessage = "Error loading domain controller status: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
        }
    })

    $ButtonQuickReplicationStatus.add_Click({
        Write-ADReportLog -Message "Loading AD replication status..." -Type Info
        try {
            $ReplicationStatus = Get-ReplicationStatus
            if ($ReplicationStatus -and $ReplicationStatus.Count -gt 0) {
                $Global:DataGridResults.ItemsSource = $ReplicationStatus
                Write-ADReportLog -Message "AD replication status loaded. $($ReplicationStatus.Count) result(s) found." -Type Info
                Update-ResultCounters -Results $ReplicationStatus
            } else {
                $Global:DataGridResults.ItemsSource = $null
                Write-ADReportLog -Message "No replication status information found." -Type Warning
            }
        } catch {
            $ErrorMessage = "Error loading AD replication status: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
        }
    })

    $ButtonExportCSV.add_Click({
        Write-ADReportLog -Message "Preparing CSV export..." -Type Info
        if ($null -eq $Global:DataGridResults.ItemsSource -or $Global:DataGridResults.Items.Count -eq 0) {
            Write-ADReportLog -Message "No data available for export." -Type Warning
            [System.Windows.Forms.MessageBox]::Show("Es sind keine Daten zum Exportieren vorhanden.", "Hinweis", "OK", "Information")
            return
        }

        $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $SaveFileDialog.Filter = "CSV (Comma delimited) (*.csv)|*.csv"
        $SaveFileDialog.Title = "Save CSV file as"
        $SaveFileDialog.InitialDirectory = [Environment]::GetFolderPath("MyDocuments")

        if ($SaveFileDialog.ShowDialog() -eq "OK") {
            $FilePath = $SaveFileDialog.FileName
            try {
                $Global:DataGridResults.ItemsSource | Export-Csv -Path $FilePath -NoTypeInformation -Encoding UTF8 -Delimiter ';'
                $Global:TextBlockStatus.Text = "Daten erfolgreich nach $FilePath exportiert."
                [System.Windows.Forms.MessageBox]::Show("Data successfully exported!", "CSV Export", "OK", "Information")
            } catch {
                $ErrorMessage = "Error in CSV export: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                [System.Windows.Forms.MessageBox]::Show("Fehler beim Exportieren der Daten: $($_.Exception.Message)", "Export Fehler", "OK", "Error")
            }
        } else {
            Write-ADReportLog -Message "CSV export canceled by user." -Type Info
        }
    })

    $ButtonExportHTML.add_Click({
        Write-ADReportLog -Message "Preparing HTML export..." -Type Info
        if ($null -eq $Global:DataGridResults.ItemsSource -or $Global:DataGridResults.Items.Count -eq 0) {
            $Global:TextBlockStatus.Text = "No data available for export."
            [System.Windows.Forms.MessageBox]::Show("Es sind keine Daten zum Exportieren vorhanden.", "Hinweis", "OK", "Information")
            return
        }

        $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $SaveFileDialog.Filter = "HTML File (*.html;*.htm)|*.html;*.htm"
        $SaveFileDialog.Title = "Save HTML file as"
        $SaveFileDialog.InitialDirectory = [Environment]::GetFolderPath("MyDocuments")

        if ($SaveFileDialog.ShowDialog() -eq "OK") {
            $FilePath = $SaveFileDialog.FileName
            try {
                # Create a more attractive HTML header
                $HtmlHead = @"
<meta http-equiv='Content-Type' content='text/html; charset=utf-8' />
<title>Active Directory Report</title>
<style>
  body { font-family: Segoe UI, Arial, sans-serif; margin: 20px; }
  table { border-collapse: collapse; width: 90%; margin: 20px auto; box-shadow: 0 0 10px rgba(0,0,0,0.1); }
  th, td { border: 1px solid #ddd; padding: 10px; text-align: left; }
  th { background-color: #0078D4; color: white; }
  tr:nth-child(even) { background-color: #f2f2f2; }
  h1 { text-align: center; color: #333; }
</style>
"@
                $DateTimeNow = Get-Date -Format "dd.MM.yyyy HH:mm:ss"
                $ReportTitle = "Active Directory Report - Created on $DateTimeNow"

                $Global:DataGridResults.ItemsSource | ConvertTo-Html -Head $HtmlHead -Body "<h1>$ReportTitle</h1>" | Out-File -FilePath $FilePath -Encoding UTF8
                $Global:TextBlockStatus.Text = "Daten erfolgreich nach $FilePath exportiert."
                [System.Windows.Forms.MessageBox]::Show("Data successfully exported!", "HTML Export", "OK", "Information")
            } catch {
                $ErrorMessage = "Error in HTML export: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                [System.Windows.Forms.MessageBox]::Show("Fehler beim Exportieren der Daten: $($_.Exception.Message)", "Export Fehler", "OK", "Error")
            }
        } else {
            Write-ADReportLog -Message "HTML export canceled by user." -Type Info
        }
    })

    # Close Button Event Handler
    if ($null -ne $Global:ButtonClose) {
        $Global:ButtonClose.add_Click({
            if ($null -ne $Global:Window) {
                $Global:Window.Close()
            }
        })
    } else {
        Write-Warning "ButtonClose not found in XAML"
    }

    # Fenster anzeigen
    $null = $Window.ShowDialog()

# Platzhalter-Funktion für initiale Visualisierung beim Start
Function Initialize-ResultVisualization {
    [CmdletBinding()]
    param()
    
    # Erstellt eine leere Visualisierung, bis Suchergebnisse vorliegen
    $emptyData = @{"Total" = 0}
    
    # Visualisierungen sind nicht im aktuellen XAML implementiert - Platzhalter für zukünftige Implementierung
    Write-ADReportLog -Message "Result visualization initialized (UI components not yet implemented)." -Type Info -Terminal
}

# --- Hauptlogik ---
Function Start-ADReportGUI {
    [CmdletBinding()]
    param()
    
    Write-Host "easyADReport v0.2.2 wird gestartet..." -ForegroundColor Green
    
    # Überprüfen, ob das AD-Modul verfügbar ist
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        Write-Warning "Active Directory PowerShell module not found. Please install RSAT or AD PowerShell module."
        Write-Host "The application will start but AD functionality will be limited." -ForegroundColor Yellow
    }
    
    # Bereinige alle alten globalen UI-Variablen vor dem Start
    $UiVariables = @("Window", "ComboBoxFilterAttribute", "TextBoxFilterValue", "ListBoxSelectAttributes",
                  "ButtonQueryAD", "ButtonQuickAllUsers", "ButtonQuickDisabledUsers", "ButtonQuickLockedUsers",
                  "ButtonQuickGroups", "ButtonQuickSecurityGroups", "ButtonQuickDistributionGroups",
                  "ButtonQuickNeverExpire", "ButtonQuickInactiveUsers", "ButtonQuickAdminUsers",
                  "ButtonQuickComputers", "ButtonQuickInactiveComputers", 
                  "ButtonQuickWeakPasswordPolicy", "ButtonQuickRiskyGroupMemberships", "ButtonQuickPrivilegedAccounts",
                  "DataGridResults", "TextBlockStatus", "ButtonExportCSV", "ButtonExportHTML",
                  "ResultCountGrid", "TotalResultCountText", "ButtonClose") 
    
    foreach ($var in $UiVariables) {
        Remove-Variable -Name $var -Scope Global -ErrorAction SilentlyContinue
    }

    try {
        # Ruft Initialize-ADReportForm auf, welche die UI lädt, Elemente zuweist und füllt.
        Initialize-ADReportForm -XAMLContent $Global:XAML
    } catch {
        Write-Error "Error starting GUI: $($_.Exception.Message)"
        Write-Host "Press any key to continue..." -ForegroundColor Red
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
}

# --- Skriptstart ---
Start-ADReportGUI