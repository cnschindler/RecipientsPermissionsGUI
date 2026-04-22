#
# Exchange Recipients Permissions GUI by Christian Schindler, NTx BOCG, christian.schinder@ntx.at
#
# Latest Version at "https://github.com/cnschindler/RecipientPermissionsGUI"
#
# Provided as is. No liabilty.
#

[cmdletbinding(SupportsShouldProcess = $true)]
Param()

# Variable definition for Logging
#
# Modulepath for file based modules
[System.IO.DirectoryInfo]$ModulePath = Join-Path -Path $PSScriptRoot -ChildPath "Modules"

# Logging variables to control the initial logging behavior, logfile name and path
# Use the script name as the base for the logfile name, add a timestamp to it.
[System.IO.FileInfo]$ScriptName = $MyInvocation.MyCommand.Name
[System.IO.FileInfo]$LogfileName = ($ScriptName.BaseName + "_{0:yyyyMMdd-HHmmss}.log" -f [DateTime]::Now)

# Combine the script directory and logfile name to create the full path of the logfile
[System.IO.FileInfo]$script:LogFileFullPath = Join-Path -Path $PSScriptRoot -ChildPath $LogfileName

# Set start and stop messages for the logfile
[string]$Script:LogFileStart = "{0:dd.MM.yyyy H:mm:ss} : {1}" -f [DateTime]::Now, "Logging started"
[string]$Script:LogFileStop = "Logging stopped"

# Set logging variables to control the initial logging behavior
$Script:LoggingEnabled = $true
$Script:FileLoggingEnabled = $true

# Common Log Messages
[string]$MSGExchangeMgmgtToolsNotInstalled = "Exchange Management Tools not found. Please install the Exchange Management Tools and try again."
[string]$MSGStopScriptExecution = "Stopping script execution."

function Write-LogFile
{
    # Logging function, used for progress and error logging
    # Uses the globally (script scoped) configured variables 'LogFileFullPath' to identify the logfile, 'LoggingEnabled' to enable/disable logging
    # and 'FileLoggingEnabled' to enable/disable file based logging
    #
    [CmdLetBinding()]

    param
    (
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [System.Management.Automation.ErrorRecord]$ErrorInfo = $null
    )

    # Prefix the string to write with the current Date and Time, add error message if present...
    if ($ErrorInfo)
    {
        $logLine = "{0:dd.MM.yyyy H:mm:ss} : ERROR : {1} The error is: {2}" -f [DateTime]::Now, $Message, $ErrorInfo.Exception.Message
    }

    Else
    {
        $logLine = "{0:dd.MM.yyyy H:mm:ss} : INFO : {1}" -f [DateTime]::Now, $Message
    }

    # If logging is enabled...
    if ($Script:LoggingEnabled)
    {
        # If file based logging is enabled, write to the logfile
        if ($Script:FileLoggingEnabled)
        {
            # Create the Script:LogfileFullPath and folder structure if it doesn't exist
            if (-not (Test-Path $script:LogFileFullPath -PathType Leaf))
            {
                New-Item -ItemType File -Path $script:LogFileFullPath -Force -Confirm:$false -WhatIf:$false | Out-Null
                Add-Content -Value $Script:LogFileStart -Path $script:LogFileFullPath -Encoding UTF8 -WhatIf:$false -Confirm:$false
            }

            # Write to Script:LogfileFullPath
            Add-Content -Value $logLine -Path $script:LogFileFullPath -Encoding UTF8 -WhatIf:$false -Confirm:$false
            Write-Verbose $logLine
        }

        # If file based logging is not enabled, Output the log line to the console
        else
        {
            # If an errorinfo was given, format the output in red
            If ($ErrorInfo)
            {
                Write-Host -ForegroundColor Red -Object $logLine
            }

            Else
            {
                Write-Host -Object $logLine
            }

        }
    }
}
Function Connect-Exchange
{
    # Check if a connection to an exchange server exists. If no connection exists, load EMS and connect
    if (-NOT (Get-PSSession | Where-Object ConfigurationName -EQ "Microsoft.Exchange"))
    {
        # Define EMS script path
        $EMSModuleFile = (Get-ItemProperty HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup -ErrorAction SilentlyContinue).MsiInstallPath + "bin\RemoteExchange.ps1"

        # If the EMS Module wasn't found
        if ([System.String]::IsNullOrEmpty($EMSModuleFile) -or -Not (Test-Path $EMSModuleFile))
        {
            # Write Error and exit the script
            Write-LogFile -Message $MSGExchangeMgmgtToolsNotInstalled
            Write-LogFile -Message $MSGStopScriptExecution
            Exit
        }

        else
        {
            # Load Exchange Management Shell
            try
            {
                # Dot source the EMS Script
                . $($EMSModuleFile) -ErrorAction Stop | Out-Null
                Write-LogFile -Message "Successfully loaded Exchange Management Shell Module."
            }

            catch
            {
                Write-LogFile -Message "Unable to load Exchange Management Shell Module." -ErrorInfo $_
                Write-LogFile -Message $MSGStopScriptExecution
                Exit
            }

            # Connect to Exchange Server
            try
            {
                Connect-ExchangeServer -auto -ClientApplication:ManagementShell -ErrorAction Stop | Out-Null
                Write-LogFile -Message "Successfully connected to Exchange Server."
            }

            catch
            {
                Write-LogFile -Message "Unable to connect to Exchange Server." -ErrorInfo $_
                Write-LogFile -Message $MSGStopScriptExecution
                Exit
            }
        }
    }
}
function LoadFileBasedModules
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [System.IO.DirectoryInfo]
        $Path
    )

    $ModuleFiles = Get-ChildItem -Path $Path -Recurse -Filter "*.psm1"

    foreach ($File in $ModuleFiles)
    {
        Import-Module -Name $file.Fullname -Force -DisableNameChecking
    }
}
function Get-ObjectPickerSelection
{
    # The ID of a Universal Group is either 8 or -2147483640
    $UGGroupID = "-2147483640"

    $Attributes = "SamAccountName","Mail","distinguishedName","GroupType"
    $AllowedObjectTypes = "Users","Groups"

    $ReturnObject = Show-ActiveDirectoryObjectPicker -AttributesToFetch $Attributes -AllowedObjectTypes $AllowedObjectTypes
    if (-Not [System.String]::IsNullOrEmpty($ReturnObject.Name))
    {
        if ($ReturnObject.SchemaClassName -eq "Group" -and -Not ($ReturnObject.FetchedAttributes[3] -match $UGGroupID))
        {
            $Message = "The selected group is not a Universal Group. Please select a Universal Group and try again."
            Write-LogFile -Message $Message
            $Textbox_Messages.Text = $Message
            Return $null
        }
        Return $ReturnObject
    }

    else
    {
        $Message = "No object was selected in the Object Picker Dialog."
        Write-LogFile -Message $Message
        $Textbox_Messages.Text = $Message
        Return $null
    }

}
Function Manage-SendAsPermissions
{
    [cmdletbinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Recipient,
        [Parameter(Mandatory = $true)]
        [string]$Assignee,
        [switch]$RemovePermission
    )

    $Recipient = Get-Recipient -Identity $Recipient -ErrorAction Stop
    $Assignee = Get-Recipient -Identity $Assignee -ErrorAction Stop

    if ($RemovePermission)
    {
        try
        {
            Remove-ADPermission -Identity $Recipient -User $Assignee -ExtendedRights "Send As" -ErrorAction Stop -Confirm:$false
            Write-LogFile -Message "Successfully removed Send As permission for $Assignee on $Recipient."
            $Textbox_Messages.Text = "Successfully removed Send As permission for $Assignee on $Recipient."
        }

        catch
        {
            Write-LogFile -Message "Unable to remove Send As permission for $Assignee on $Recipient." -ErrorInfo $_
            $Textbox_Messages.Text = "Unable to remove Send As permission for $Assignee on $Recipient. The error is: $_"
        }
    }

    else
    {
        try
        {
            Add-ADPermission -Identity $Recipient -User $Assignee -ExtendedRights "Send As" -ErrorAction Stop
            Write-LogFile -Message "Successfully added Send As permission for $Assignee on $Recipient."
            $Textbox_Messages.Text = "Successfully added Send As permission for $Assignee on $Recipient."
        }

        catch
        {
            Write-LogFile -Message "Unable to add Send As permission for $Assignee on $Recipient." -ErrorInfo $_
            $Textbox_Messages.Text = "Unable to add Send As permission for $Assignee on $Recipient. The error is: $_"
        }
    }

}

#Region XAMLForm
Add-Type -AssemblyName PresentationFramework, System.Drawing, System.Windows.Forms, WindowsFormsIntegration

[xml]$XAMLForm = @'
<Window x:Name="Windows_Laps_Viewer" x:Class="EchangeRecipientPermissionsGUI.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:EchangeRecipientPermissionsGUI"
        mc:Ignorable="d"
        Title="NTx Exchange Recipient Permissions GUI" Width="530" MinWidth="550" Height="380" MinHeight="380" ResizeMode="NoResize" ScrollViewer.VerticalScrollBarVisibility="Disabled" SizeToContent="Height">
    <Grid VerticalAlignment="Top" HorizontalAlignment="Left" Width="530" Height="360">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>
        <Grid.RowDefinitions>
            <RowDefinition Height="60"/>
            <RowDefinition Height="60"/>
            <RowDefinition Height="60"/>
            <RowDefinition Height="140"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>
        <Label x:Name="Label_Recipient" Content="Recipient" HorizontalAlignment="Left" Margin="20,10,0,0" VerticalAlignment="Top" Width="72" FontWeight="Bold"/>
        <TextBox x:Name="Textbox_Recipient" HorizontalAlignment="Left" Height="25" Margin="20,0,0,2" TextWrapping="Wrap" VerticalAlignment="Bottom" Width="370" TabIndex="1" FontSize="12"/>
        <Button x:Name="Button_SelectRecipient" HorizontalAlignment="Right" Height="30" Margin="0,0,20,0" VerticalAlignment="Bottom" Width="100" Grid.Column="1">
            <AccessText Text="Select Recipient..." TextWrapping="Wrap" TextAlignment="Center"/>
        </Button>
        <Label x:Name="Label_Assignee" Content="Assignee" HorizontalAlignment="Left" Height="27" Margin="20,10,0,0" VerticalAlignment="Top" Width="65" FontWeight="Bold" Grid.Row="1"/>
        <TextBox x:Name="Textbox_Assignee" HorizontalAlignment="Left" Height="25" Margin="20,0,0,2" VerticalAlignment="Bottom" Width="370" FontSize="12" Grid.Row="1"/>
        <Button x:Name="Button_SelectAssignee" HorizontalAlignment="Right" Height="30" Margin="0,0,20,0" VerticalAlignment="Bottom" Width="100" Grid.Column="1" Grid.Row="1">
            <AccessText Text="Select Assignee..." TextWrapping="Wrap" TextAlignment="Center"/>
        </Button>
        <Button x:Name="Button_AddSendAs" HorizontalAlignment="Left" Height="30" Width="150" Grid.Column="1" Grid.Row="2" Margin="20,0,7,0">
            <AccessText Text="Add SendAs Permission" TextWrapping="Wrap" TextAlignment="Center"/>
        </Button>
        <Button x:Name="Button_RemoveSendAs" HorizontalAlignment="Left" Height="30" Width="160" Grid.Column="1" Grid.Row="2" Margin="190,0,0,0">
            <AccessText Text="Remove SendAs Permission" TextWrapping="Wrap" TextAlignment="Center"/>
        </Button>
        <Label x:Name="Label_Messages" Content="Output Messages" HorizontalAlignment="Left" Height="27" Margin="20,0,0,0" VerticalAlignment="Top" Width="137" FontWeight="Bold" Grid.ColumnSpan="2" Grid.Row="3"/>
        <TextBox x:Name="Textbox_Messages" HorizontalAlignment="Left" Height="110" Margin="20,20,0,0" Width="490" FontSize="12" Grid.ColumnSpan="2" Grid.Row="3" TextWrapping="Wrap"/>
        <TextBlock x:Name="Textblock_Info" HorizontalAlignment="Left" Margin="20,0,0,0" TextWrapping="Wrap" Text="Exchange Recipients Permissions GUI by Christian Schindler, NTx BOCG, christian.schindler@ntx.at" VerticalAlignment="Top" Width="482" Grid.ColumnSpan="2" Grid.Row="4"/>
    </Grid>
</Window>
'@ -replace 'mc:Ignorable="d"', '' -replace "x:Name", 'Name' -replace '^<Win.*', '<Window' -replace 'x:Class="\S+"', ''

# Read XAMLForm
$form = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $XAMLForm))
$XAMLForm.SelectNodes("//*[@Name]") | Where-Object { Set-Variable -Name ($_.Name) -Value $Form.FindName($_.Name) }

# Icon definition
#
# Icon in Base64 format
[string]$IconB64 = @"
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
"@

# Convert icon to bitmap
$bitmap = New-Object System.Windows.Media.Imaging.BitmapImage
$bitmap.BeginInit()
$bitmap.StreamSource = [System.IO.MemoryStream][System.Convert]::FromBase64String($IconB64)
$bitmap.EndInit()
$bitmap.Freeze()

# Add icon to form
$Form.Icon = $bitmap
#endregion XAMLForm

# Disable several Buttons by default
$Button_SelectAssignee.IsEnabled = $false
$Button_AddSendAs.IsEnabled = $false
$Button_RemoveSendAs.IsEnabled = $false

# Set Focus on the Computername Textbox
$Textbox_Recipient.Focus() | Out-Null

# Handler for changed text in the Textbox
$Textbox_Recipient.Add_TextChanged(
    {
        $Textbox_Messages.Clear()
        $Button_SelectAssignee.IsEnabled = $true
    }
)

# Handler for changed text in the Textbox
$Textbox_Assignee.Add_TextChanged(
    {
        $Textbox_Messages.Clear()
        $Button_AddSendAs.IsEnabled = $true
        $Button_RemoveSendAs.IsEnabled = $true
    }
)

# Handler for Select Recipient Button click
$Button_SelectRecipient.Add_Click(
    {
        $Recipient = Get-ObjectPickerSelection
        if ([System.String]::IsNullOrEmpty($Recipient.Name))
        {
            $Textbox_Messages.Text = $Textbox_Messages.Text + "`nNo Recipient was selected!"
            Return
        }

        Else
        {
            $Textbox_Recipient.Text = $Recipient.fetchedAttributes[1]
            $Script:SelectedRecipient = $Recipient.fetchedAttributes[2]
        }
    }
)

# Handler for Select Assignee Button click
$Button_SelectAssignee.Add_Click(
    {
        $Assignee = Get-ObjectPickerSelection
        if ([System.String]::IsNullOrEmpty($Assignee.Name))
        {
            $Textbox_Messages.Text = $Textbox_Messages.Text + "`nNo Assignee was selected!"
            Return
        }

        else
        {
            $Textbox_Assignee.Text = $Assignee.fetchedAttributes[1]
            $Script:SelectedAssignee = $Assignee.fetchedAttributes[1]
        }
    }
)

# Handler for Add Send As Button click
$Button_AddSendAs.Add_Click(
    {
        $Textbox_Messages.Clear()

        if (-not $Script:SelectedRecipient -or -not $Script:SelectedAssignee)
        {
            $Textbox_Messages.Text = "No Recipient or Assignee was specified!"
        }

        else
        {
            Manage-SendAsPermissions -Recipient $Script:SelectedRecipient -Assignee $Script:SelectedAssignee
        }
    }
)

# Handler for Remove Send As Button click
$Button_RemoveSendAs.Add_Click(
    {
        $Textbox_Messages.Clear()

        if (-not $Script:SelectedRecipient -or -not $Script:SelectedAssignee)
        {
            $Textbox_Messages.Text = "No Recipient or Assignee was specified!"
        }

        else
        {
            Manage-SendAsPermissions -Recipient $Script:SelectedRecipient -Assignee $Script:SelectedAssignee -RemovePermission
        }
    }
)

# Load file based modules from the Modules folder
LoadFileBasedModules -Path $ModulePath
# Connect to Exchange Server, load EMS if no connection exists
Connect-Exchange | Out-Null

# Load Form
$Form.ShowDialog() | Out-Null