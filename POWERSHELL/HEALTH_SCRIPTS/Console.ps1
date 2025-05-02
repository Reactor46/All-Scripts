#===========================================================================
# Build the GUI Runspace
#===========================================================================

Set-Location C:\LazyWinAdmin
$Jobs = [system.collections.arraylist]::Synchronized((New-Object System.Collections.ArrayList))
$UiHash = [hashtable]::Synchronized(@{})
$RunspaceHash = [hashtable]::Synchronized(@{})
$RunspaceHash.Host = $Host
$RunspaceHash.Runspace = [runspacefactory]::CreateRunspace()
$RunspaceHash.Runspace.ApartmentState = "STA"
$RunspaceHash.Runspace.ThreadOptions = "ReuseThread"
$RunspaceHash.Runspace.Open()
$runspaceHash.psCmd = {Add-Type -AssemblyName PresentationFramework}.GetPowerShell()
$InitialSessionState = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
$InitialSessionState.ImportPSSnapIn('Microsoft.Exchange.Management.PowerShell.E2010',[ref]'') | Out-Null
$InitialSessionState.ImportPSModule('ActiveDirectory')
$RunspaceHash.Runspace.SessionStateProxy.SetVariable('InitialSessionState',$InitialSessionState)
$RunspaceHash.Runspace.SessionStateProxy.SetVariable('UiHash',$UiHash)
$RunspaceHash.Runspace.SessionStateProxy.SetVariable('RunspaceHash',$RunspaceHash)
$RunspaceHash.Runspace.SessionStateProxy.SetVariable('Jobs',$Jobs)
$RunspaceHash.Runspace.SessionStateProxy.SetVariable('ScriptsPath', ((get-itemproperty HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\Setup).MsiInstallPath + "scripts"))
$RunspaceHash.psCmd.Runspace = $RunspaceHash.Runspace 
$RunspaceHash.Handle = $RunspaceHash.psCmd.AddScript({

$RunspaceHash.Host.UI.RawUI.ForegroundColor="cyan"
$RunspaceHash.Host.UI.WriteLine("Loading Application GUI, Please Wait..")

Function Log-Error($block)
{
    $Activity = $Error[0].CategoryInfo.Activity
    $ExceptionMessage = $_.Exception.Message
    $Line = $Error[0].InvocationInfo.ScriptLineNumber
    $Char = $Error[0].InvocationInfo.OffsetInLine
    If($block)	 
    {
        $UiHash.Window.Dispatcher.invoke(
        [action]{$UiHash.$block.text = $ExceptionMessage}
        )
    }
    $(date -f "dd.MM.yyyy HH:mm:ss") + " " + "(At Line: $Line, Char: $Char)" + " " + "$Activity $ExceptionMessage" >> $Errors
}

#===========================================================================
# Build the GUI Layout
#===========================================================================

[xml]$xaml = @"
<Window
	xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
	xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Console" Height="500" Width="1000" ResizeMode="CanMinimize" WindowStartupLocation="CenterScreen">
    <Window.Resources>
        <Style TargetType="DatePickerTextBox">
            <Setter Property="VerticalAlignment" Value="Center" />
            <Setter Property="HorizontalAlignment" Value="Center" />
        </Style>
        <Style TargetType="Label">
            <Setter Property="Padding" Value="4" />
        </Style>
        <Style TargetType="{x:Type Window}">
            <Setter Property="FontSize" Value="11" />
            <Setter Property="FontFamily" Value="Segoe UI"/>
        </Style>
    </Window.Resources>
    <Grid Margin="0,-1,0,1" Background="#FFE4E4E4">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="125" MaxWidth="125"/>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="880*"/>
        </Grid.ColumnDefinitions>
        <Grid.RowDefinitions>
            <RowDefinition Height="356*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="117" MinHeight="57" MaxHeight="272"/>
        </Grid.RowDefinitions>
        <ListView  Name="QResult_lsv" HorizontalAlignment="Stretch" Height="235" VerticalAlignment="Center" Margin="10,201,10,37" BorderThickness="1" Background="#FFECECEC" Grid.Column="2" Grid.RowSpan="3">
            <ListView.View>
                <GridView>
                    <GridViewColumn Header="Identity" DisplayMemberBinding ="{Binding Identity}" Width="Auto"/>
                    <GridViewColumn Header="Delivery Type" DisplayMemberBinding ="{Binding DeliveryType}" Width="Auto"/>
                    <GridViewColumn Header="Status" DisplayMemberBinding ="{Binding Status}" Width="Auto"/>
                    <GridViewColumn Header="Message Count" DisplayMemberBinding ="{Binding MessageCount}" Width="Auto"/>
                    <GridViewColumn Header="Next Hop Domain" DisplayMemberBinding ="{Binding NextHopDomain}" Width="Auto"/>
                </GridView>
            </ListView.View>
        </ListView>
        <Label Name="MsgCount_lbl" Content="Message count higher than:" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="274,172,0,0" Grid.Column="2" Height="23" Width="145"/>
        <TextBox Name="CheckPrereq_box" HorizontalAlignment="Stretch" TextWrapping="Wrap" VerticalAlignment="Stretch" Margin="10,201,10,0" AcceptsReturn="True" ScrollViewer.HorizontalScrollBarVisibility="Auto" ScrollViewer.VerticalScrollBarVisibility="Auto" ScrollViewer.CanContentScroll="True" IsReadOnly="True" BorderThickness="1" Background="#FFECECEC" Grid.Column="2" Grid.Row="0"/>
        <TextBox Name="VrfFAccess_box" HorizontalAlignment="Stretch" Height="235" TextWrapping="Wrap" VerticalAlignment="Top" Margin="10,201,10,0" AcceptsReturn="True" ScrollViewer.HorizontalScrollBarVisibility="Auto" ScrollViewer.VerticalScrollBarVisibility="Auto" ScrollViewer.CanContentScroll="True" IsReadOnly="True" BorderThickness="1" Background="#FFECECEC" Grid.Column="2" Grid.RowSpan="3"/>
        <TextBox Name="Body_box" HorizontalAlignment="Stretch" Height="235" TextWrapping="Wrap" VerticalAlignment="Top" Margin="10,201,10,0" AcceptsReturn="True" ScrollViewer.HorizontalScrollBarVisibility="Auto" ScrollViewer.VerticalScrollBarVisibility="Auto" ScrollViewer.CanContentScroll="True" BorderThickness="1" Background="#FFECECEC" Grid.Column="2" Grid.RowSpan="3"/>
        <DataGrid Name="TrackResult_gvw" HorizontalAlignment="Stretch" Height="235" VerticalAlignment="Center" Margin="10,201,10,37" BorderThickness="1" Background="#FFECECEC" Grid.Column="2" Grid.RowSpan="3" GridLinesVisibility="None" AutoGenerateColumns="False" ColumnWidth="Auto">
            <DataGrid.Columns>            
                    <DataGridTextColumn Header="Time Stamp" Binding ="{Binding TimeStamp}"/>
                    <DataGridTextColumn Header="Event ID" Binding ="{Binding EventID}"/>
                    <DataGridTextColumn Header="Source" Binding ="{Binding Source}"/>
                    <DataGridTextColumn Header="Sender" Binding ="{Binding Sender}"/>
                    <DataGridTextColumn Header="Recipients" Binding ="{Binding Recipients}"/>
                    <DataGridTextColumn Header="Recipient Count" Binding ="{Binding RecipientCount}"/>
                    <DataGridTextColumn Header="Client Host Name" Binding ="{Binding ClientHostName}"/>
                    <DataGridTextColumn Header="Server Host Name" Binding ="{Binding ServerHostName}"/>
                    <DataGridTextColumn Header="Connector ID" Binding ="{Binding ConnectorID}"/>
                    <DataGridTextColumn Header="Message Subject" Binding ="{Binding MessageSubject}"/>
                    <DataGridTextColumn Header="Message ID" Binding ="{Binding MessageID}"/>
                    <DataGridTextColumn Header="Size(MB)" Binding ="{Binding Size(MB)}"/>
            </DataGrid.Columns>
        </DataGrid>
        <ListView Name="BkpResult_lsv" HorizontalAlignment="Stretch" Height="235" VerticalAlignment="Center" Margin="10,201,10,37" BorderThickness="1" Background="#FFECECEC" Grid.Column="2" Grid.RowSpan="3">
            <ListView.View>
                <GridView>
                    <GridViewColumn Header="Name" DisplayMemberBinding ="{Binding Name}" Width="Auto"/>
                    <GridViewColumn Header="Server" DisplayMemberBinding ="{Binding Server}" Width="Auto"/>
                    <GridViewColumn Header="Last Incremental Backup" DisplayMemberBinding ="{Binding LastIncrementalBackup}" Width="Auto"/>
                    <GridViewColumn Header="Last Full Backup" DisplayMemberBinding ="{Binding LastFullBackup}" Width="Auto"/>
                    <GridViewColumn Header="Backup In Progress" DisplayMemberBinding ="{Binding BackupInProgress}" Width="Auto"/>
                </GridView>
            </ListView.View>
        </ListView>
        <ListView  Name="MDBCopyStatus_lsv" HorizontalAlignment="Stretch" Margin="10,0,10,35" VerticalAlignment="Stretch" BorderThickness="1" Grid.Column="2" Grid.Row="2" Background="#FFECECEC">
            <ListView.View>
                <GridView>
                    <GridViewColumn Header="Name" DisplayMemberBinding ="{Binding Name}" Width="Auto"/>
                    <GridViewColumn Header="Status" DisplayMemberBinding ="{Binding Status}" Width="Auto"/>
                    <GridViewColumn Header="Copy Queue Length" DisplayMemberBinding ="{Binding CopyQueueLength}" Width="Auto"/>
                    <GridViewColumn Header="Replay Queue Length" DisplayMemberBinding ="{Binding ReplayQueueLength}" Width="Auto"/>
                    <GridViewColumn Header="Last Inspected Log Time" DisplayMemberBinding ="{Binding ReplayQueueLength}" Width="Auto"/>
                    <GridViewColumn Header="Content Index State" DisplayMemberBinding ="{Binding ContentIndexState}" Width="Auto"/>
                </GridView>
            </ListView.View>
        </ListView>
        <Button Name="GetBackup_btn" Content="Get" HorizontalAlignment="Right" Height="21" Margin="0,174,10,0" VerticalAlignment="Top" Width="63" Grid.Column="2"/>
        <Label Name="CheckPoint_lbl" Content="Checkpoint:" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,152,0,0" Grid.Column="2" Height="23" Width="70"/>
        <DatePicker Name="CheckpointDate_pkr" HorizontalAlignment="Left" Margin="78,152,0,0" VerticalAlignment="Top" SelectedDateFormat="Short" FirstDayOfWeek="Monday" DisplayDateEnd="$((Get-Date).ToString('yyyy-MM-dd'))" DisplayDateStart="$((Get-Date).AddDays(-30).ToString('yyyy-MM-dd'))" SelectedDate="$((Get-Date).ToString('yyyy-MM-dd'))" Width="100" BorderThickness="1" Padding="3,0,0,0" Height="20" Background="#FFECECEC" Grid.Column="2"/>
        <ComboBox Name="CheckpointTime_cmb" HorizontalAlignment="Left" VerticalAlignment="Top" Width="55" Margin="180,152,0,0" Background="#FFECECEC" Grid.Column="2" Height="20" IsEditable="True" SelectedIndex="0"/>
        <Button Name="CheckPrereq_btn" Content="Check Prerequisites" HorizontalAlignment="Right" Height="21" Margin="0,174,10,0" VerticalAlignment="Top" Width="110" Grid.Column="2"/>
        <Button Name="Verify_btn" Content="Verify" HorizontalAlignment="Right" Height="21" Margin="0,174,10,0" VerticalAlignment="Top" Width="63" Grid.Column="2"/>
        <Button Name="Close_btn" Content="Close" HorizontalAlignment="Right" Height="21" Margin="0,0,10,10" VerticalAlignment="Bottom" Width="63" Grid.Column="2" Grid.Row="2"/>
        <Button Name="Proceed_btn" Content="Proceed" HorizontalAlignment="Right" Height="21" Margin="0,0,73,10" VerticalAlignment="Bottom" Width="63" Grid.Column="2" Grid.Row="2"/>
        <Button Name="ExportToCSV_btn" Content="Export to CSV" HorizontalAlignment="Right" Height="21" Margin="0,0,73,10" VerticalAlignment="Bottom" Width="100" Grid.Column="2" Grid.Row="2"/>
        <Button Name="Send_btn" Content="Send" HorizontalAlignment="Right" Height="21" Margin="0,0,73,10" VerticalAlignment="Bottom" Width="63" Grid.Column="2" Grid.Row="2"/>
        <Button Name="Track_btn" Content="Track" HorizontalAlignment="Right" Height="21" Margin="0,174,10,0" VerticalAlignment="Top" Width="63" Grid.Column="2"/>
        <Button Name="Clean_btn" Content="Clean" HorizontalAlignment="Right" Height="21" Margin="0,174,10,0" VerticalAlignment="Top" Width="63" Grid.Column="2"/>
        <Button Name="GetQ_btn" Content="Get" HorizontalAlignment="Right" Height="21" Margin="0,174,10,0" VerticalAlignment="Top" Width="63" Grid.Column="2"/>
        <TextBlock Name="BkpProcess_bck" Text="" HorizontalAlignment="Stretch" VerticalAlignment="Bottom" Margin="10,0,78,6" Height="25" TextTrimming="CharacterEllipsis" Padding="0,5,0,0" Grid.Column="2" Grid.Row="2"/>
        <TextBlock Name="MsgQProcess_bck" Text="" HorizontalAlignment="Stretch" VerticalAlignment="Bottom" Margin="10,0,78,6" Height="25" TextTrimming="CharacterEllipsis" Padding="0,5,0,0" Grid.Column="2" Grid.Row="2"/>
        <TextBlock Name="SwitchSrvProcess_bck" Text="" HorizontalAlignment="Stretch" VerticalAlignment="Bottom" Margin="10,0,141,6" Height="25" TextTrimming="CharacterEllipsis" Padding="0,5,0,0" Grid.Column="2" Grid.Row="2"/>
        <TextBlock Name="CSVProcess_bck" Text="" HorizontalAlignment="Stretch" VerticalAlignment="Bottom" Margin="10,0,178,6" Height="25" TextTrimming="CharacterEllipsis" Padding="0,5,0,0" Grid.Column="2" Grid.Row="2"/>
        <TextBlock Name="FAccessProcess_bck" Text="" HorizontalAlignment="Stretch" VerticalAlignment="Bottom" Margin="10,0,78,6" Height="25" TextTrimming="CharacterEllipsis" Padding="0,5,0,0" Grid.Column="2" Grid.Row="2"/>
        <TextBlock Name="SendMailProcess_bck" Text="" HorizontalAlignment="Stretch" VerticalAlignment="Bottom" Margin="10,0,141,6" Height="25" TextTrimming="CharacterEllipsis" Padding="0,5,0,0" Grid.Column="2" Grid.Row="2"/>
        <TextBox Name="Sender_box" HorizontalAlignment="Left" Height="20" VerticalAlignment="Top" Width="189" Margin="78,108,0,0" BorderThickness="1" Background="#FFECECEC" Grid.Column="2"/>
        <TextBox Name="Rcpt_box" HorizontalAlignment="Left" Height="20" VerticalAlignment="Top" Width="189" Margin="78,130,0,0" BorderThickness="1" Background="#FFECECEC" Grid.Column="2"/>
        <Label Name="Sender_lbl" Content="Sender:" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,108,0,0" Grid.Column="2" Height="23" Width="48"/>
        <Label Name="Rcpt_lbl" Content="Recipient:" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,130,0,0" Grid.Column="2" Height="23" Width="58"/>
        <Label Name="Start_lbl" Content="Start:" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,152,0,0" Grid.Column="2" Height="23" Width="37"/>
        <DatePicker Name="StartDate_pkr" HorizontalAlignment="Left" Margin="78,152,0,0" VerticalAlignment="Top" SelectedDateFormat="Short" FirstDayOfWeek="Monday" DisplayDateEnd="$((Get-Date).ToString('yyyy-MM-dd'))" DisplayDateStart="$((Get-Date).AddDays(-30).ToString('yyyy-MM-dd'))" SelectedDate="$((Get-Date).ToString('yyyy-MM-dd'))" Width="100" BorderThickness="1" Padding="3,0,0,0" Height="20" Background="#FFECECEC" Grid.Column="2"/>
        <ComboBox  Name="StartTime_cmb" HorizontalAlignment="Left" VerticalAlignment="Top" Width="55" Margin="180,152,0,0" Background="#FFECECEC" Grid.Column="2" Height="20" IsEditable="True" SelectedIndex="0"/>
        <Label Name="End_lbl" Content="End:" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="274,152,0,0" Grid.Column="2" Height="23" Width="32"/>
        <DatePicker Name="EndDate_pkr" HorizontalAlignment="Left" Margin="380,152,0,0" VerticalAlignment="Top" SelectedDateFormat="Short" FirstDayOfWeek="Monday" DisplayDateEnd="$((Get-Date).ToString('yyyy-MM-dd'))" DisplayDateStart="$((Get-Date).AddDays(-30).ToString('yyyy-MM-dd'))" SelectedDate="$((Get-Date).ToString('yyyy-MM-dd'))" Width="100" BorderThickness="1" Padding="3,0,0,0" Height="20" Background="#FFECECEC" Grid.Column="2"/>
        <ComboBox Name="EndTime_cmb" HorizontalAlignment="Left" VerticalAlignment="Top" Width="55" Margin="482,152,0,0" Background="#FFECECEC" Grid.Column="2" Height="20" IsEditable="True" SelectedIndex="23"/>
        <Label Name="MsgSubject_lbl" Content="Message Subject:" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="274,108,0,0" Grid.Column="2" Height="23" Width="95"/>
        <TextBox Name="MsgSubject_box" HorizontalAlignment="Left" Height="20" VerticalAlignment="Top" Width="189" Margin="380,108,0,0" BorderThickness="1" Background="#FFECECEC" Grid.Column="2"/>
        <Label Name="MsgID_lbl" Content="Message ID:" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="274,130,0,0" Grid.Column="2" Height="23" Width="71"/>
        <TextBox Name="MsgID_box" HorizontalAlignment="Left" Height="20" VerticalAlignment="Top" Width="189" Margin="380,130,0,0" BorderThickness="1" Background="#FFECECEC" Grid.Column="2"/>
        <Label Name="EvtID_lbl" Content="Event ID:" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="274,172,0,0" Grid.Column="2" Height="23" Width="56"/>
        <ComboBox Name="EvtID_cmb" HorizontalAlignment="Left" VerticalAlignment="Top" Width="90" Margin="380,174,0,0" Background="#FFECECEC" Grid.Column="2" Height="21"/>
        <TextBox Name="Server_box" HorizontalAlignment="Left" Height="20" VerticalAlignment="Top" Width="189" Margin="78,174,0,0" BorderThickness="1" Background="#FFECECEC" Grid.Column="2"/>
        <Label Name="Server_lbl" Content="Server:" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,174,0,0" Width="45" Grid.Column="2" Height="23"/>
        <TextBox Name="To_box" HorizontalAlignment="Left" Height="20" VerticalAlignment="Top" Width="189" Margin="78,108,0,0" BorderThickness="1" Background="#FFECECEC" Grid.Column="2"/>
        <Label Name="To_lbl" Content="To:" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,108,0,0" Grid.Column="2" Height="23" Width="26"/>
        <Label Name="Cc_lbl" Content="Cc:" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,130,0,0" Grid.Column="2" Height="23" Width="26"/>
        <TextBox Name="Cc_box" HorizontalAlignment="Left" Height="20" VerticalAlignment="Top" Width="189" Margin="78,130,0,0" BorderThickness="1" Background="#FFECECEC" Grid.Column="2"/>
        <Label Name="Subject_lbl" Content="Subject:" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,152,0,0" Grid.Column="2" Height="23" Width="50"/>
        <TextBox Name="Subject_box" HorizontalAlignment="Left" Height="20" VerticalAlignment="Top" Width="240" Margin="78,152,0,0" BorderThickness="1" Background="#FFECECEC" Grid.Column="2"/>
        <Label Name="Who_lbl" Content="Who:" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,130,0,0" Grid.Column="2" Height="23" Width="36"/>
        <Label Name="Where_lbl" Content="Where:" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,152,0,0" Grid.Column="2" Height="23" Width="45"/>
        <Button Name="Grant_btn" Content="Grant" HorizontalAlignment="Left" Height="21" Margin="10,174,0,0" VerticalAlignment="Top" Width="63" Grid.Column="2"/>
        <Button Name="Remove_btn" Content="Remove" HorizontalAlignment="Left" Height="21" Margin="73,174,0,0" VerticalAlignment="Top" Width="63" Grid.Column="2"/>
        <TextBox Name="Who_box" HorizontalAlignment="Left" Height="20" VerticalAlignment="Top" Width="189" Margin="78,130,0,0" BorderThickness="1" Background="#FFECECEC" Grid.Column="2"/>
        <TextBox Name="Where_box" HorizontalAlignment="Left" Height="20" VerticalAlignment="Top" Width="189" Margin="78,152,0,0" BorderThickness="1" Background="#FFECECEC" Grid.Column="2"/>
        <RadioButton Name="SwitchSrvRadio_btn" Content="Switchover Server" HorizontalAlignment="Left" Margin="10,130,0,0" VerticalAlignment="Top" Grid.Column="2" Height="13" Width="102"/>
        <RadioButton Name="RedistributeDBsRadio_btn" Content="Redistribute Active Databases" HorizontalAlignment="Left" Margin="10,152,0,0" VerticalAlignment="Top" Grid.Column="2" Height="13" Width="159"/>
        <DockPanel Name="Slider_dck" Margin="428,173,0,0" HorizontalAlignment="Left" VerticalAlignment="Top" Width="192" Grid.Column="2" Height="22">
            <TextBox Name="MsgCount_box" Text="{Binding Value, ElementName=slValue, UpdateSourceTrigger=PropertyChanged}" DockPanel.Dock="Right" TextAlignment="Center" Width="28" VerticalContentAlignment="Center" BorderThickness="1" Background="#FFECECEC" Height="22"/>
            <Slider Name="slValue" Height="22" IsSnapToTickEnabled="True" Width="158" Maximum="300" TickFrequency="30" Value="30"/>
        </DockPanel>
        <Label Name="MenuPosition_lbl" Content="Message Tracking" HorizontalAlignment="Left" Height="30" Margin="10,15,0,0" VerticalAlignment="Top" Width="240" FontSize="14" Grid.Column="2"/>
        <TextBlock Name="Tip_bck" HorizontalAlignment="Left" Height="56" Margin="10,50,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="720" FontSize="10" Grid.Column="2"/>
        <TreeView HorizontalAlignment="Left" Margin="1,25,0,0" VerticalAlignment="Top" BorderThickness="0" Background="{x:Null}" ScrollViewer.HorizontalScrollBarVisibility="Disabled" Grid.Column="0" Grid.RowSpan="3">
            <TreeViewItem Header="Exchange" IsExpanded="True">
                <TreeViewItem Name="MessageTracking" Header="Message Tracking" HorizontalAlignment="Left" Margin="-16,0,0,0" Width="117" IsSelected="True"/>
                <TreeViewItem Name="MessageQueues" Header="Message Queues" HorizontalAlignment="Left" Margin="-16,0,0,0" Width="117"/>
                <TreeViewItem Name="Backups" Header="Backups" HorizontalAlignment="Left" Margin="-16,0,0,0" Width="117"/>
                <TreeViewItem Name="SwitchoverServer" Header="Switchover Server" HorizontalAlignment="Left" Margin="-16,0,0,0" Width="117"/>
                <TreeViewItem Name="FullAccess" Header="Full Access" HorizontalAlignment="Left" Margin="-16,0,0,0" Width="117"/>
                <TreeViewItem Name="SendEmail" Header="Send E-Mail" HorizontalAlignment="Left" Margin="-16,0,0,0" Width="117"/>
            </TreeViewItem>
            <TreeViewItem Header="Tools" IsExpanded="False">
                <TreeViewItem Name="ActiveDirectory" Header="Active Directory" HorizontalAlignment="Left" Margin="-16,0,0,0" Width="117"/>
                <TreeViewItem Name="ADSIEdit" Header="ADSIEdit" HorizontalAlignment="Left" Margin="-16,0,0,0" Width="117"/>
                <TreeViewItem Name="OWA" Header="OWA" HorizontalAlignment="Left" Margin="-16,0,0,0" Width="117"/>
                <TreeViewItem Name="TaskScheduler" Header="Task Scheduler" HorizontalAlignment="Left" Margin="-16,0,0,0" Width="117"/>
            </TreeViewItem>
        </TreeView>
        <GridSplitter HorizontalAlignment="Center" Margin="0,10" VerticalAlignment="Stretch" Width="2" Grid.Column="1" Grid.RowSpan="3" Background="#FFD5D5D5"/>
        <GridSplitter Name='SrvSwitch_gsr' HorizontalAlignment="Stretch" Height="2" Margin="10,1,10,0" VerticalAlignment="Center" Grid.Column="2" Grid.Row="1" Background="#FFD5D5D5"/>
    </Grid>
</Window>
"@

Try
{
    #===========================================================================
    # Read XAML
    #===========================================================================

    $reader = (New-Object System.Xml.XmlNodeReader $xaml)
    $UiHash.Window = [Windows.Markup.XamlReader]::Load($reader)


    #===========================================================================
    # Store XAML Objects In PowerShell
    #===========================================================================

    $xaml.SelectNodes("//*[@Name]")| %{$UiHash.($_.name) = $UiHash.Window.FindName($_.Name)}


    #===========================================================================
    # Define the Check-Prerequisites function
    #===========================================================================

    Function Check-Prerequisites($Servers)
    {
        Foreach($Server in $servers)
        {
	        #===========================================================================
            # Test-ServiceHealth
            #===========================================================================

            $UiHash.Window.Dispatcher.invoke(
            [action]{$UiHash.SwitchSrvProcess_bck.Text = 'Processing: Test-ServiceHealth in progress..'}
            )

            If($ServiceHealth = Test-ServiceHealth -ser $Server | ?{$_.RequiredServicesRunning -ne $True})
            {
                $ServiceHealth.ServicesNotRunning | %{[String]$ServicesNotRunning += "$_`n"}
                $UiHash.Window.Dispatcher.invoke(
                [action]{$UiHash.CheckPrereq_box.text += "Required services are not running on $server`:`n$ServicesNotRunning"}
                )
                Throw "Required services are not running on $Server"
                $ServiceStatus++
            }
            Else
            {
                $UiHash.Window.Dispatcher.invoke(
                [action]{$UiHash.CheckPrereq_box.text += "Required services are running on $server ...OK`n"}
                )
            }

            #===========================================================================
            # Test-ReplicationHealth
            #===========================================================================

            $UiHash.Window.Dispatcher.invoke(
            [action]{$UiHash.SwitchSrvProcess_bck.Text = 'Processing: Test-ReplicationHealth in progress..'}
            )

            If($ReplicationHealth = Test-ReplicationHealth -ser $Server | ?{$_.Error -ne "$null"})
            {
                $ReplicationHealth.Error | %{[String]$ReplicationHealthError += "$_`n"}
                $UiHash.Window.Dispatcher.invoke(
                [action]{$UiHash.CheckPrereq_box.text += "Replication health check failed on $server`:`n$ReplicationHealthError"}
                )
                Throw "Test-Replication health check failed on $Server"
                $ReplicationStatus++
            }
            Else
            {
                $UiHash.Window.Dispatcher.invoke(
                [action]{$UiHash.CheckPrereq_box.text += "Replication health check was successful on $server ...OK`n"}
                )
            }

            #===========================================================================
            # Mailbox Database Copy Status Check
            #===========================================================================

            $UiHash.Window.Dispatcher.invoke(
            [action]{$UiHash.SwitchSrvProcess_bck.Text = 'Processing: Checking mailbox database copy status..'}
            )

            If([Array]$MailboxDatabaseCopyStatus = mailboxdatabase -ser $Server | ?{$_.recovery -ne $True -and $_.ReplicationType -eq 'Remote'} | mailboxdatabasecopystatus | ?{$_.mailboxserver -eq $($Server.split('.')[0]) -and ($_.CopyQueueLength -ge '5' -or $_.ReplayQueueLength -ge '5' -or $_.ContentIndexState -ne 'Healthy')} | sort name)
            {
                $UiHash.Window.Dispatcher.invoke(
                [action]{$UiHash.MDBCopyStatus_lsv.ItemsSource = $MailboxDatabaseCopyStatus;$UiHash.CheckPrereq_box.text += "Mailbox database copy status check failed on $server`:"}
                )
                Throw "Mailbox database copy status check failed on $Server"
                $MailboxDatabaseCopyStatus++
            }
            Else
            {
                $UiHash.Window.Dispatcher.invoke(
                [action]{$UiHash.CheckPrereq_box.text += "Mailbox database copy status check was successful on $server ...OK`n"}
                )
            }
            $UiHash.Window.Dispatcher.invoke(
            [action]{$UiHash.CheckPrereq_box.text += "------------------------------------------------------------------------`n"}
            )
        }
    }


    #===========================================================================
    # Build the Exchange Runspacepool
    #===========================================================================

    $Function = Get-Content Function:\Log-Error
    $SessionStateFunction = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList 'Log-Error', $Function
    $InitialSessionState.Commands.Add($SessionStateFunction)
    $Function = Get-Content Function:\Check-Prerequisites
    $SessionStateFunction = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList 'Check-Prerequisites', $Function
    $InitialSessionState.Commands.Add($SessionStateFunction)
    $Errors = New-object System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList 'Errors', '.\Errors.txt', $null
    $InitialSessionState.Variables.Add($Errors)
    $RunspaceHash.RunspacePool = [runspacefactory]::CreateRunspacePool(1, 1, $InitialSessionState, $Host)
    $RunspaceHash.RunspacePool.Open()


    #===========================================================================
    # Actually make the objects work
    #===========================================================================
 							  
    Function Get-ExchangeServers
    {
        Return $((($UiHash.Server_box.text) -replace " ", "").split(','))
    }   

    Function Visibility($a1, $b1, $c1, $d1, $e1, $f1, $g1)
    {
        $a = $b = $c = $d = $e = $f = $g = $null
	
	    Foreach($a in $FullAccessGroup){($a).visibility = "$a1"}
	    Foreach($b in $BackupGroup){($b).visibility = "$b1"}
        Foreach($c in $MessageQueueGroup){($c).visibility = "$c1"}
	    Foreach($d in $SendEMailGroup){($d).visibility = $d1}
	    Foreach($e in $MessageTrackingGroup){($e).visibility = "$e1"}
	    Foreach($f in $ServerGroup){($f).visibility = "$f1"}
	    Foreach($g in $SwitchoverServerGroup){($g).visibility = "$g1"}
    }


    #===========================================================================
    # Assigning GUI Objects into Visibility Groups
    #===========================================================================

    $FullAccessGroup = @()
    $FullAccessGroup = $UiHash.Who_lbl, $UiHash.Where_lbl, $UiHash.Grant_btn, $UiHash.Remove_btn, $UiHash.Who_box, $UiHash.Where_box, $UiHash.FAccessProcess_bck, $UiHash.VrfFAccess_box, $UiHash.Verify_btn

    $BackupGroup = @()
    $BackupGroup = $UiHash.CheckPoint_lbl, $UiHash.CheckpointDate_pkr, $UiHash.CheckpointTime_cmb, $UiHash.BkpResult_lsv, $UiHash.GetBackup_btn, $UiHash.BkpProcess_bck

    $MessageQueueGroup = @()
    $MessageQueueGroup = $UiHash.QResult_lsv, $UiHash.GetQ_btn, $UiHash.MsgQProcess_bck, $UiHash.Slider_dck, $UiHash.MsgCount_lbl, $UiHash.MsgCount_box

    $SendEMailGroup = @()
    $SendEMailGroup = $UiHash.To_lbl, $UiHash.Cc_lbl, $UiHash.To_box, $UiHash.Cc_box, $UiHash.Subject_lbl, $UiHash.Subject_box, $UiHash.Body_box, $UiHash.Send_btn, $UiHash.SendMailProcess_bck, $UiHash.Clean_btn

    $MessageTrackingGroup = @()
    $MessageTrackingGroup = $UiHash.EndTime_cmb, $UiHash.EndDate_pkr, $UiHash.End_lbl, $UiHash.EvtID_cmb, $UiHash.EvtID_lbl, $UiHash.MsgID_box, $UiHash.MsgID_lbl, $UiHash.MsgSubject_box, $UiHash.MsgSubject_lbl, $UiHash.Rcpt_box, $UiHash.Rcpt_lbl, $UiHash.Sender_box, $UiHash.Sender_lbl, $UiHash.StartTime_cmb, $UiHash.StartDate_pkr, $UiHash.Start_lbl, $UiHash.Track_btn, $UiHash.TrackResult_gvw, $UiHash.ExportToCSV_btn, $UiHash.CSVProcess_bck

    $ServerGroup = @()
    $ServerGroup = $UiHash.Server_lbl, $UiHash.Server_box

    $SwitchoverServerGroup = @()
    $SwitchoverServerGroup = $UiHash.SwitchSrvRadio_btn, $UiHash.RedistributeDBsRadio_btn, $UiHash.CheckPrereq_btn, $UiHash.CheckPrereq_box, $UiHash.SwitchSrvProcess_bck, $UiHash.Proceed_btn, $UiHash.MDBCopyStatus_lsv, $UiHash.SrvSwitch_gsr

    Visibility "Hidden" "Hidden" "Hidden" "Hidden" "Visible" "Visible" "Hidden"

    $UiHash.MenuPosition_lbl.Visibility = 'Visible'

    $UiHash.Tip_bck.Text = 'Tip: Save your findings by click on "Export to CSV" button.'

    '','BADMAIL','DELIVER','DSN','EXPAND','FAIL','POISONMESSAGE','RECEIVE','REDIRECT','RESOLVE','SEND','SUBMIT','TRANSFER' | %{$UiHash.EvtID_cmb.Items.Add($_)}


    For($o = 0; $o -lt 24; $o++){$UiHash.StartTime_cmb.Items.Add("$(date -h $o -f HH)`:00"); $UiHash.EndTime_cmb.Items.Add("$(date -h $o -f HH)`:00"); $UiHash.CheckpointTime_cmb.Items.Add("$(date -h $o -f HH)`:00")}

    $UiHash.Server_box.text = 'Server01'


    $UiHash.MessageTracking.Add_MouseLeftButtonUp(
    {
	    Visibility "Hidden" "Hidden" "Hidden" "Hidden" "Visible" "Visible" "Hidden"

        $UiHash.MenuPosition_lbl.Content = 'Message Tracking'
        $UiHash.Tip_bck.Text = 'Tip: Save your findings by click on "Export to CSV" button.'    
    })

    $UiHash.FullAccess.Add_MouseLeftButtonUp(
    {
	    Visibility "Visible" "Hidden" "Hidden" "Hidden" "Hidden" "Hidden" "Hidden"

        $UiHash.MenuPosition_lbl.Content = 'Full Access'
        $UiHash.Tip_bck.Text = 'Tip: Fill the "Who:" field in with an email address and click on the "Verify" button, this will give you back list of users, where used smtp address has the full access. Fill "Where:" field and click on the "Verify" button to get list of users having access to used email address mailbox. These result will be returned based on automapping feature, if full access was granted earlier with this feature bypassed particular item will not be returned.'
    })
	
    $UiHash.MessageQueues.Add_MouseLeftButtonUp(
    {
	    Visibility "Hidden" "Hidden" "Visible" "Hidden" "Hidden" "Visible" "Hidden"

        $UiHash.MenuPosition_lbl.Content = 'Message Queues'
        $UiHash.Tip_bck.Text = 'Tip: Set the minimum message count per queue by slider.'
    })

    $UiHash.SwitchoverServer.Add_MouseLeftButtonUp(
    {
	    Visibility "Hidden" "Hidden" "Hidden" "Hidden" "Hidden" "Visible" "Visible"

        $UiHash.MenuPosition_lbl.Content = 'Switchover Server'
        $UiHash.Tip_bck.Text = 'Tip: For server switchover, fill the source and the target server names in (delimited by comma), click "Check Prerequisites" button, if all requirements are evaluated successfully, you can continue by click on "Proceed" button. For database redistribution, fill an Exchange server name (or its DAG name) in and click on "Check Prerequisites" button, if all requirements are evaluated successfully, you can continue by click on "Proceed" button.'
    })

    $UiHash.Backups.Add_MouseLeftButtonUp(
    {
	    Visibility "Hidden" "Visible" "Hidden" "Hidden" "Hidden" "Visible" "Hidden"

        $UiHash.MenuPosition_lbl.Content = 'Backups'
        $UiHash.Tip_bck.Text = 'Tip: Display databases where last backup date/time is less than the checkpoint.'
    })

    $UiHash.SendEMail.Add_MouseLeftButtonUp(
    {
	    Visibility "Hidden" "Hidden" "Hidden" "Visible" "Hidden" "Visible" "Hidden"

        $UiHash.MenuPosition_lbl.Content = 'Send E-Mail'
        $UiHash.Tip_bck.Text = 'Tip: This feature allows you to send easily one or many email messages, you can use "Clean" button to remove previously filled parameters.'
    })

    $UiHash.Send_btn.Add_Click({
    If($UiHash.To_box.text -and $UiHash.Subject_box.text)
    {
        $ScriptBlock = {
	    param(
            $UiHash,
            $parameters,
            $ErrorActionPreference = 'stop'
        )

        Try
	    {
            $UiHash.Window.Dispatcher.invoke(
            [action]{$UiHash.SendMailProcess_bck.Text = 'Processing: Sending the e-mail..'}
            )
            Send-MailMessage -From $((mailbox $(whoami)).primarysmtpaddress) @parameters
            $UiHash.Window.Dispatcher.invoke(
            [action]{$UiHash.SendMailProcess_bck.Text = 'Processing: E-mail was sent.'}
            )
	    }
	    Catch
	    {
    	    Log-Error -block 'SendMailProcess_bck'
	    }
        }
        $parameters = @{
            To          = $($UiHash.To_box.text)
            Subject     = $($UiHash.Subject_box.text)
	    } 
	    If($UiHash.Cc_box.text){$parameters.Add("Cc",$($UiHash.Cc_box.text))}
        If($UiHash.Body_box.text){$parameters.Add("Body",$($UiHash.Body_box.text))}
        If($UiHash.Server_box.text){$parameters.Add("SmtpServer",$($UiHash.Server_box.text))}

        $Powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($UiHash).AddArgument($parameters)
        $Powershell.RunspacePool = $RunspaceHash.RunspacePool
        $Jobs.Add($Powershell.BeginInvoke())
    }
    })
		
    $UiHash.GetBackup_btn.Add_Click({
    $ScriptBlock = {
    Param(
        $UiHash,
        $Servers,
        $CheckPoint,
        $ErrorActionPreference = 'stop'
    )

    Try
    {
        $UiHash.Window.Dispatcher.invoke(
        [action]{$UiHash.BkpResult_lsv.ItemsSource = $null;$UiHash.BkpResult_lsv.Items.Clear();$UiHash.BkpProcess_bck.text = 'Processing: Searching for obsolete backups..'}
        )
        $Backup = @()

        If($CheckPoint)
        {
            Foreach($Server in ($servers | get-exchangeserver).name)
            {
                $Backup += get-mailboxdatabase -server $Server -status | ?{$_.server -like $server.split(".")[0] -and ($_.lastfullbackup -lt $CheckPoint -and $_.lastincrementalbackup -lt $CheckPoint) -and $_.name -notlike "*test*" -and $_.name -notlike "*restore*"} | Select name, server, lastincrementalbackup, lastfullbackup, backupinprogress
            }
        }
        Else
        {
            Foreach($Server in ($servers | get-exchangeserver).name)
            {
                $Backup += get-mailboxdatabase -server $Server -status | ?{$_.server -like $server.split(".")[0] -and $_.name -notlike "*test*" -and $_.name -notlike "*restore*"} | Select name, server, lastincrementalbackup, lastfullbackup, backupinprogress
            }
        }                
        $UiHash.Window.Dispatcher.invoke(
        [action]{$UiHash.BkpProcess_bck.text = [String]::Empty;$UiHash.BkpResult_lsv.ItemsSource = ($Backup | sort name)}
        )
    }
    Catch
	{
        Log-Error -block 'BkpProcess_bck'
	}
    }
    $Powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($UiHash).AddArgument($(Get-ExchangeServers)).AddArgument("$($UiHash.CheckpointDate_pkr.Text) $($UiHash.CheckpointTime_cmb.Text)")
    $Powershell.RunspacePool = $RunspaceHash.RunspacePool
    $Jobs.Add($Powershell.BeginInvoke())
    })

    $UiHash.GetQ_btn.Add_Click({
    $ScriptBlock = {
        param(    
            $UiHash,
            $Servers,
            $MsgCount,
            $ErrorActionPreference = 'stop'
        )
        
    Try
    {
        $UiHash.Window.Dispatcher.invoke(
        [action]{$UiHash.QResult_lsv.ItemsSource = $null;$UiHash.QResult_lsv.Items.Clear();$UiHash.MsgQProcess_bck.Text = 'Processing: Searching message queues..'}
        )       
        $Queue = @()	
        Foreach($Server in ($servers | Get-ExchangeServer).name)
	    {
	        $Queue += Get-Queue -server $Server | ?{$_.Messagecount -gt "$MsgCount" -and $_.DeliveryType -ne "ShadowRedundancy"} | Select Identity, DeliveryType, Status, MessageCount, NextHopDomain
	    }
	    $UiHash.Window.Dispatcher.invoke(
        [action]{$UiHash.MsgQProcess_bck.text = [String]::Empty;$UiHash.QResult_lsv.ItemsSource = $Queue}
        )
    }
	Catch
	{
        Log-Error -block 'MsgQProcess_bck'
	}
    }
    $Powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($UiHash).AddArgument($(Get-ExchangeServers)).AddArgument($($UiHash.MsgCount_box.text))
    $Powershell.RunspacePool = $RunspaceHash.RunspacePool
    $Jobs.Add($Powershell.BeginInvoke())
    })
 
    $UiHash.Track_btn.Add_Click({
    $ScriptBlock = {
    param(
        $UiHash,
        $Servers,
        $parameters,
        $ErrorActionPreference = 'stop'
    )

    Try
    {
        $UiHash.Window.Dispatcher.invoke(
        [action]{$UiHash.TrackResult_gvw.ItemsSource = $null;$UiHash.TrackResult_gvw.Items.Clear();$UiHash.CSVProcess_bck.text = 'Processing: Message tracking in progress..'}
        )
        $UiHash.Remove('Result')     
        $UiHash.Result = @(($servers | Get-ExchangeServer).name | Get-MessageTrackingLog @parameters -resultsize unlimited | select timestamp, eventid, source, sender, @{n='recipients';e={$_.recipients}}, clienthostname, serverhostname, connectorID, messagesubject, recipientcount, messageid, @{n='Size(MB)';e={[math]::round($_.totalbytes/1MB)}} | sort timestamp)
        $UiHash.Window.Dispatcher.invoke(
        [action]{$UiHash.CSVProcess_bck.text = [String]::Empty;$UiHash.TrackResult_gvw.ItemsSource = $UiHash.Result}
        )
    }
    Catch
    {
        Log-Error -block 'CSVProcess_bck'
    }
    }

    $parameters = @{}
    If($UiHash.StartDate_pkr.Text -and $UiHash.StartTime_cmb.Text){$Parameters.Add('Start',"$($UiHash.StartDate_pkr.Text) $($UiHash.StartTime_cmb.Text)")}
    If($UiHash.EndDate_pkr.Text -and $UiHash.EndTime_cmb.Text){$Parameters.Add('End',"$($UiHash.EndDate_pkr.Text) $($UiHash.EndTime_cmb.Text)")}
    If($UiHash.Sender_box.Text){$Parameters.Add('Sender',$($UiHash.Sender_box.Text))}
    If($UiHash.Rcpt_box.Text){$Parameters.Add('Recipient',$($UiHash.Rcpt_box.Text))}
    If($UiHash.MsgSubject_box.Text){$Parameters.Add('MessageSubject',$($UiHash.MsgSubject_box.Text))}
    If($UiHash.MsgID_box.Text){$Parameters.Add('MessageID',$($UiHash.MsgID_box.Text))}
    If($UiHash.EvtID_cmb.Text){$Parameters.Add('EventID',$($UiHash.EvtID_cmb.Text))}
    
    $Powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($UiHash).AddArgument($(Get-ExchangeServers)).AddArgument($parameters)
    $Powershell.RunspacePool = $RunspaceHash.RunspacePool
    $Jobs.Add($Powershell.BeginInvoke())
    })

    $UiHash.ExportToCSV_btn.Add_Click({
    If($UiHash.Result)
    {
	    Try
	    {
	        $UiHash.Result | Export-Csv -Path .\MessageTracking.csv -NoTypeInfo
	        $UiHash.CSVProcess_bck.Text = 'Processing: Export-CSV - Completed!'
	    }
	    Catch
	    {
	        Log-Error -block 'CSVProcess_bck'
	    }
    }
    })

    $UiHash.Grant_btn.Add_Click({
    If($UiHash.Who_box.text -and $UiHash.Where_box.text)
    {
        $ScriptBlock = {
        param(
            $UiHash,
            $parameters,
            $ErrorActionPreference = 'stop'
        )
        
        Try
        {
            $UiHash.Window.Dispatcher.invoke(
            [action]{$UiHash.VrfFAccess_box.Clear();$UiHash.FAccessProcess_bck.Text = 'Processing: Attempting to add full access to mailbox..'}
            )
            Set-ADServerSettings -ViewEntireForest $true -WarningAction SilentlyContinue 
    	    Add-MailboxPermission @parameters -AccessRights Fullaccess -InheritanceType all | out-null
            $UiHash.Window.Dispatcher.invoke(
            [action]{$UiHash.FAccessProcess_bck.Text = 'Processing: Full Access Granted.'}
            )            
        }
        Catch
        {
            Log-Error -block 'FAccessProcess_bck'
        }
    }
    $parameters = @{ 
        Identity    = $($UiHash.Where_box.text)
        User        = $($UiHash.Who_box.text)
    }
    $Powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($UiHash).AddArgument($parameters)
    $Powershell.RunspacePool = $RunspaceHash.RunspacePool
    $Jobs.Add($Powershell.BeginInvoke())
    }
    })

    $UiHash.Remove_btn.Add_Click({
    If($UiHash.Who_box.text -and $UiHash.Where_box.text)
    {
        $ScriptBlock = {
        param(
            $UiHash,
            $Identity,
            $User,
            $ErrorActionPreference = 'stop'
        )
        
        Try
        {
            $UiHash.Window.Dispatcher.invoke(
            [action]{$UiHash.VrfFAccess_box.Clear();$UiHash.FAccessProcess_bck.Text = 'Processing: Attempting to remove full access from mailbox..'}
            )

            Set-ADServerSettings -ViewEntireForest $true -WarningAction SilentlyContinue
            $Identity | %{Remove-MailboxPermission -Identity $_ -user $User -AccessRights Fullaccess -InheritanceType all -confirm:$False}

            $UiHash.Window.Dispatcher.invoke(
            [action]{$UiHash.FAccessProcess_bck.Text = 'Processing: Full Access Removed.'}
            )            
        }
        Catch
        {
            Log-Error -block 'FAccessProcess_bck'
        }
    }
    $Identity = $($UiHash.Where_box.text).replace(" ","").split(",")
    $User = $($UiHash.Who_box.text)

    $Powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($UiHash).AddArgument($Identity).AddArgument($User)
    $Powershell.RunspacePool = $RunspaceHash.RunspacePool
    $Jobs.Add($Powershell.BeginInvoke())
    }
    })

    $UiHash.Verify_btn.Add_Click({
    If($UiHash.Where_box.text -and !($UiHash.Who_box.text))
    {
        $ScriptBlock = {
        param(
            $UiHash,
            $Identity,
            $ErrorActionPreference = 'stop'
        )
        
        Try
        {
            $UiHash.Window.Dispatcher.invoke(
            [action]{$UiHash.VrfFAccess_box.Clear();$UiHash.FAccessProcess_bck.Text = 'Processing: Searching for users having access to mailbox..'}
            )
            $list = @();[String]::Empty
            Set-ADServerSettings -ViewEntireForest $true -WarningAction SilentlyContinue
            $SAMAccountName = (recipient $Identity).SAMAccountName
            (get-aduser -filter {samaccountname -eq $SAMAccountName} -properties * -Server "$((ADDomainController).hostname):3268").msExchDelegateListLink | %{$list += (recipient $_).primarysmtpaddress}
            
            $UiHash.Window.Dispatcher.invoke(
            [action]{$UiHash.FAccessProcess_bck.text = [String]::Empty;$UiHash.VrfFAccess_box.Clear();$UiHash.VrfFAccess_box.text = ($list -join ", ")}
            )
        }
        Catch
	    {
	        Log-Error -block 'FAccessProcess_bck'
	    }
        }
        $Identity = $($UiHash.Where_box.text)

        $Powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($UiHash).AddArgument($Identity)
        $Powershell.RunspacePool = $RunspaceHash.RunspacePool
        $Jobs.Add($Powershell.BeginInvoke())
    }

    If($UiHash.Who_box.text -and !($UiHash.Where_box.text))
    {
        $ScriptBlock = {
        param(
            $UiHash,
            $User,
            $ErrorActionPreference = 'stop'
        )
        
        Try
        {
            $UiHash.Window.Dispatcher.invoke(
            [action]{$UiHash.VrfFAccess_box.Clear();$UiHash.FAccessProcess_bck.Text = 'Processing: Searching for mailboxes where user has access..'}
            )
            $list = @();[String]::Empty
            Set-ADServerSettings -ViewEntireForest $true -WarningAction SilentlyContinue
            $SAMAccountName = (recipient $User).SAMAccountName
            (get-aduser -filter {samaccountname -eq $SAMAccountName} -properties * -Server "$((ADDomainController).hostname):3268").msExchDelegateListBL | %{$list += (recipient $_).primarysmtpaddress}

            $UiHash.Window.Dispatcher.invoke(
            [action]{$UiHash.FAccessProcess_bck.text = [String]::Empty;$UiHash.VrfFAccess_box.Clear();$UiHash.VrfFAccess_box.text = ($list -join ", ")}
            )
        }
        Catch
	    {
	        Log-Error -block 'FAccessProcess_bck'
	    }
    }
    $User = $UiHash.Who_box.text
        
    $Powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($UiHash).AddArgument($User)
    $Powershell.RunspacePool = $RunspaceHash.RunspacePool
    $Jobs.Add($Powershell.BeginInvoke())
    }
    })

    $UiHash.CheckPrereq_btn.Add_Click({
    If($UiHash.SwitchSrvRadio_btn.IsChecked)
    {
        $ScriptBlock = {
        param(
            $UiHash,
            $Servers,
            $ErrorActionPreference = 'stop'
        )
    
        Try
        {
            $UiHash.Window.Dispatcher.invoke(
            [action]{$UiHash.CheckPrereq_box.Clear();$UiHash.MDBCopyStatus_lsv.ItemsSource = $null;$UiHash.MDBCopyStatus_lsv.Items.Clear();$UiHash.SwitchSrvProcess_bck.Text = 'Processing: Prerequisite check in progress..'}
            )
        
            If($Servers.Count -eq 2)
            {
                If(!(DatabaseAvailabilityGroup | ?{$_.servers -match "$($Servers[0])" -and $_.servers -match "$($Servers[1])"}))
                {
                    Throw "Servers are not members of the same DAG."
                    break
                }
                ElseIf("$($Servers[0])" -eq "$($Servers[1])")
                {
                    Throw "Server names are not unique."
                    break
                }
                Else
                {
                    $DAG = DatabaseAvailabilityGroup | ?{$_.servers -match "$($Servers[0])" -and $_.servers -match "$($Servers[1])"}
                    $Servers = $Servers | %{[System.Net.Dns]::GetHostByName("$_").hostname.ToLower()}
                    $UiHash.Window.Dispatcher.invoke(
                    [action]{$UiHash.CheckPrereq_box.text += "Database Availability Group: $($DAG.name)`n------------------------------------------------------------------------`n"}
                    )
                }
            }
            Else
            {
                Throw "Not expected server count. Required number of servers is 2."
                break
            }

            Check-Prerequisites -Servers $Servers

            If(!($ServiceStatus -and $ReplicationStatus -and $MailboxDatabaseCopyStatus))
            {
                $Databases = Get-MailboxDatabase -Server $Servers[1] -status | ?{$_.ReplicationType -eq "Remote"} | sort name
                $FailoverStatus = 0
	            Foreach($Database in $Databases)
                {
                    If($Servers[1].split('.')[0] -notmatch $Database.server)
	                {
                        $FailoverStatus++
                    }
                }
                $UiHash.PrerequisiteResult = "OK"
                $UiHash.Window.Dispatcher.invoke(
                [action]{$UiHash.SwitchSrvProcess_bck.Text = [String]::Empty;$UiHash.CheckPrereq_box.text += "Test successful.`n------------------------------------------------------------------------`nCount of databases to be moved from: $($Servers[0].split('.')[0]) to: $($Servers[1].split('.')[0]): $FailoverStatus`n------------------------------------------------------------------------`nIf you want to continue, click `"Proceed`" button.`n"}
                )
                $UiHash.Servers = $Servers
            }
            Else
            {
                $UiHash.PrerequisiteResult = "Not OK"
            }
        }
        Catch
        {
            Log-Error -block 'SwitchSrvProcess_bck'
        }
    }
    $Powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($UiHash).AddArgument($(Get-ExchangeServers))
    $Powershell.RunspacePool = $RunspaceHash.RunspacePool
    $Jobs.Add($Powershell.BeginInvoke())
    }

    If($UiHash.RedistributeDBsRadio_btn.IsChecked)
    {
        $ScriptBlock = {
        param(
            $UiHash,
            $Servers,
            $ErrorActionPreference = 'stop'
        )
    
        Try
        {
            $UiHash.Window.Dispatcher.invoke(
            [action]{$UiHash.CheckPrereq_box.Clear();$UiHash.MDBCopyStatus_lsv.ItemsSource = $null;$UiHash.MDBCopyStatus_lsv.Items.Clear();$UiHash.SwitchSrvProcess_bck.Text = 'Processing: Prerequisite check in progress..'}
            )

            If($Servers.Count -eq 1)
            {
                If($DAG = databaseavailabilitygroup | ?{$_.name -match $Servers})
                {
                    $Servers = $DAG.Servers | sort name | %{[System.Net.Dns]::GetHostByName("$_").hostname.ToLower()}
                    $UiHash.Window.Dispatcher.invoke(
                    [action]{$UiHash.CheckPrereq_box.text += "Database Availability Group: $($DAG.name)`n------------------------------------------------------------------------`n"}
                    )
                }
                ElseIf($DAG = databaseavailabilitygroup | ?{$_.servers -match $Servers})
                {
                    $Servers = $DAG.Servers | sort name | %{[System.Net.Dns]::GetHostByName("$_").hostname.ToLower()}
                    $UiHash.Window.Dispatcher.invoke(
                    [action]{$UiHash.CheckPrereq_box.text += "Database Availability Group: $($DAG.name)`n------------------------------------------------------------------------`n"}
                    )
                }
            }
            Else
            {
                Throw "Not Expected Server Count. Required Number of Servers is 1."
                break
            }

            Check-Prerequisites -Servers $Servers

            If(!($ServiceStatus -and $ReplicationStatus -and $MailboxDatabaseCopyStatus))
            {
                $Databases = Get-MailboxDatabase -Server $Servers[0] -status | ?{$_.ReplicationType -eq "Remote"} | sort name
                $FailoverStatus = 0
	            Foreach($Database in $Databases)
                {
                    If($Database.ActivationPreference[0] -notmatch $Database.Server)
	                {
                        $FailoverStatus++
                    }
                }
                $UiHash.PrerequisiteResult = "OK"                
                $UiHash.Window.Dispatcher.invoke(
                [action]{$UiHash.SwitchSrvProcess_bck.Text = [String]::Empty;$UiHash.CheckPrereq_box.text += "Test succeessful.`n------------------------------------------------------------------------`nCount of databases to be moved according to activation preference: $FailoverStatus`n------------------------------------------------------------------------`nIf you want to continue, click `"Proceed`" button.`n"}
                )
                $UiHash.DAG = $DAG.name
            }
            Else
            {
                $UiHash.PrerequisiteResult = "Not OK"
            }
        }
        Catch
        {
            Log-Error -block 'SwitchSrvProcess_bck'
        }
        }
        $Powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($UiHash).AddArgument($(Get-ExchangeServers))
        $Powershell.RunspacePool = $RunspaceHash.RunspacePool
        $Jobs.Add($Powershell.BeginInvoke())
    }
    })

    $UiHash.Proceed_btn.Add_Click({
    If($UiHash.SwitchSrvRadio_btn.IsChecked -and $UiHash.PrerequisiteResult -eq "OK")
    {
        Try
        {
            Start powershell -ArgumentList "-NoExit", "-sta", "-noprofile" , "-command asnp Microsoft.Exchange.Management.PowerShell.E2010; Move-ActiveMailboxDatabase -Server $($UiHash.Servers[0]) -ActivateOnServer $($UiHash.Servers[1]) -MountDialOverride lossless"
            $UiHash.Remove('PrerequisiteResult')
            $UiHash.Remove('Servers')
        }
        Catch
        {
            Log-Error -block 'SwitchSrvProcess_bck'
        }
    }

    If($UiHash.RedistributeDBsRadio_btn.IsChecked -and $UiHash.PrerequisiteResult -eq "OK")
    {
        Try
        {
            start powershell -ArgumentList "-NoExit", "-sta", "-noprofile", "-command & '$ScriptsPath\RedistributeActiveDatabases.ps1' -DagName $($UiHash.DAG) -BalanceDbsByActivationPreference"
            $UiHash.Remove('PrerequisiteResult')
            $UiHash.Remove('DAG')
        }
        Catch
        {
            Log-Error -block 'SwitchSrvProcess_bck'
        }    
    }
    })

    $UiHash.ActiveDirectory.Add_MouseDoubleClick({
    dsa.msc
    })

    $UiHash.ADSIEdit.Add_MouseDoubleClick({
    ADSIEdit.msc
    })

    $UiHash.OWA.Add_MouseDoubleClick({
    $ScriptBlock = {
    param(
        $Servers
    )
    
    If($VirtualDirectory = Get-OwaVirtualDirectory -ser $Servers)
    {
        start $(($VirtualDirectory.InternalURL).absoluteUri)
    }
    }
    $Powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument(($(Get-ExchangeServers) -split ",")[0])
    $Powershell.RunspacePool = $RunspaceHash.RunspacePool
    $Jobs.Add($Powershell.BeginInvoke())
    })

    $UiHash.TaskScheduler.Add_MouseDoubleClick({
    Taskschd.msc
    })

    $UiHash.Clean_btn.Add_Click({
    $UiHash.To_box.text = $UiHash.Cc_box.text = $UiHash.Subject_box.text = $UiHash.Body_box.text = $UiHash.SendMailProcess_bck.text = [String]::Empty
    })

    $UiHash.Close_btn.Add_Click({
    $UiHash.Window.Close()
    $RunspaceHash.RunspacePool.Close()
    $RunspaceHash.RunspacePool.Dispose()
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()
    Stop-Process -Id $PID -Force
    })
    
    $WindowCode = '[DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);' 
    $AsyncWindow = Add-Type -MemberDefinition $WindowCode -name Win32ShowWindowAsync -namespace Win32Functions -PassThru 
    $null = $AsyncWindow::ShowWindowAsync((Get-Process -PID $PID).MainWindowHandle, 0)
    
    $UiHash.Window.ShowDialog()
}
Catch
{
    $RunspaceHash.Errors = $Error[0]
    $($Error[0]) >> .\RunspaceErrors.txt
    $RunspaceHash.Host.UI.WriteErrorLine("$($RunspaceHash.Errors.FullyQualifiedErrorId)")
}
}).BeginInvoke()