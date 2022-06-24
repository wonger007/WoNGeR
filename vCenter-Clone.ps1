<#
    Modified from @guyrleech 2018's - Clone one or more VMware ESXi VMs from a 'template' VM
    https://github.com/guyrleech/VMware
#>

<#
.SYNOPSIS

Clone one ore more VMware vCenter VMs from a Template

.DESCRIPTION

Clone a Template multiple times via GUI within an vCenter environment.

.PARAMETER esxihost

The name or IP address of the vCenter host to use.

.PARAMETER templateName

The exact name or regular expression matching the template VM to use. Must only match one VM

.PARAMETER dataStore

The datastore to create the copied hard disks in

.PARAMETER vmName

The name of the VM to create. If creating more than one, it must contain %d which will be replaced by the number of the clone

.PARAMETER snapshot

The name of the snapshot to use when creating a linked clone. Specifying this automatically enables the linked clone disk feature

.PARAMETER count

The number of clones to create

.PARAMETER startFrom

The number to start the cloning naming from.

.PARAMETER notes

Notes to assign to the created VM(s)

.PARAMETER powerOn

Power on the VM(s) once creation is complete

.PARAMETER disconnect

Disconnect from ESXi before exit

.PARAMETER maxVmdkDescriptorSize

If the vmdk file exceeds this size then the script will not attempt to edit it because it is probably a binary file and not a text descriptor

.EXAMPLE

& '.\ESXi cloner.ps1'

Run the user interface which will require various fields to be completed and when "OK" is clicked, create VMs as per these fields

.NOTES
Thanks to guyrleech @ https://github.com/guyrleech

#>

[CmdletBinding()]

Param
(
    [string]$esxihost ,
    [string]$templateName ,
    [string]$dataStore ,
    [string]$vmName ,
    [string]$snapshot ,
    [switch]$noGui ,
    [int]$count = 1 ,
    [int]$startFrom = 1 ,
    [string]$notes ,
    [switch]$powerOn ,
    [switch]$disconnect ,
    [string]$username ,
    [string]$password ,
    ## it is advised not to use the following parameters
    [int]$maxVmdkDescriptorSize = 10KB
)

## Adding so we can make it app modal as well as system
Add-Type @'
using System;
using System.Runtime.InteropServices;

namespace PInvoke.Win32
{
    public static class Windows
    {
        [DllImport("user32.dll")]
        public static extern int MessageBox(int hWnd, String text, String caption, uint type);
    }
}
'@

#region Functions

Function Connect-Hypervisor( $GUIobject , $servers , [bool]$pregui, [ref]$vServer , [string]$username , [string]$password )
{
    [hashtable]$connectParams = @{}

    if( ! [string]::IsNullOrEmpty( $username ) )
    {
        $connectParams.Add( 'User' , $username )
    }
    elseif( ! [string]::IsNullOrEmpty( $env:esxiusername ) )
    {
        $connectParams.Add( 'User' , $env:esxiusername )
    }
    if( ! [string]::IsNullOrEmpty( $password ) )
    {
        $connectParams.Add( 'Password' , $password )
    }
    elseif( ! [string]::IsNullOrEmpty( $env:esxipassword ) )
    {
        $connectParams.Add( 'Password' , $env:esxipassword )
    }

    $vServer.Value = Connect-VIServer -Server $servers -ErrorAction Continue @connectParams

    if( ! $vServer.Value )
    {
        $null = Display-MessageBox -window $GUIobject -text "Failed to connect to $($servers -join ' , ')" -caption 'Unable to Connect' -buttons OK -icon Error
    }
    elseif( ! $pregui )
    {   
        $_.Handled = $true
        $WPFbtnConnect.Content = 'Connected'
        $WPFcomboTemplate.Items.Clear()
        $WPFcomboTemplate.IsEnabled = $true          
        Get-Template | Select -ExpandProperty Name | ForEach-Object { $WPFcomboTemplate.items.add( $_ ) }
        if( $WPFcomboTemplate.items.Count -eq 1 )
        {
            $WPFcomboTemplate.SelectedIndex = 0
        }
        elseif( ! [string]::IsNullOrEmpty( $template ) )
        {
            $WPFcomboTemplate.SelectedValue = $template
        }
        $WPFcomboDatastore.Items.Clear()
        $WPFcomboDatastore.IsEnabled = $true          
        Get-Datastore | Select -ExpandProperty Name | ForEach-Object { $WPFcomboDatastore.items.add( $_ ) }
        if( $WPFcomboDatastore.items.Count -eq 1 )
        {
            $WPFcomboDatastore.SelectedIndex = 0
        }
        elseif( ! [string]::IsNullOrEmpty( $dataStore ) )
        {
            $WPFcomboDatastore.SelectedValue = $dataStore
        }
        $WPFcomboFolder.Items.Clear()
        $WPFcomboFolder.IsEnabled = $true          
        Get-Folder | Where 'Type' -EQ 'VM' | Where 'Name' -NE 'vm' | Select -ExpandProperty Name | ForEach-Object { $null = $WPFcomboFolder.items.add( $_ ) }
        if( $WPFcomboFolder.items.Count -eq 1 )
        {
            $WPFcomboFolder.SelectedIndex = 0
        }
        elseif( ! [string]::IsNullOrEmpty( $folder ) )
        {
            $WPFcomboFolder.SelectedValue = $folder
        }
        $WPFcomboResource.Items.Clear()
        $WPFcomboResource.IsEnabled = $true
        Get-VMHost | Select -ExpandProperty Name | ForEach-Object { $null = $WPFcomboResource.items.add( $_ ) }
        Get-Cluster | Select -ExpandProperty Name | ForEach-Object { $null = $WPFcomboResource.items.add( $_ ) }
        Get-ResourcePool | Where 'Name' -NE 'Resources' | Select -ExpandProperty Name | ForEach-Object { $null = $WPFcomboResource.items.add( $_ ) }
        Get-VApp | Select -ExpandProperty Name | ForEach-Object { $null = $WPFcomboResource.items.add( $_ ) }
        if( $WPFcomboResource.items.Count -eq 1 )
        {
            $WPFcomboResource.SelectedIndex = 0
        }
        elseif( ! [string]::IsNullOrEmpty( $resource ) )
        {
            $WPFcomboResource.SelectedValue = $resource
        }
    }
}


Function Validate-Fields( $guiobject )
{
    $_.Handled = $true
    
    if( [string]::IsNullOrEmpty( $WPFtxtVMName.Text ) )
    {
        $null = Display-MessageBox -window $guiobject -text 'No VM Name Specified' -caption 'Unable to Clone' -buttons OK -icon Error
        return $false
    }
    if( ! $WPFcomboTemplate.SelectedItem )
    {
        $null = Display-MessageBox -window $guiobject -text 'No Template VM Selected' -caption 'Unable to Clone' -buttons OK -icon Error
        return $false
    }
    if( ! $WPFcomboDatastore.SelectedItem )
    {
        $null = Display-MessageBox -window $guiobject -text 'No Datastore Selected' -caption 'Unable to Clone' -buttons OK -icon Error
        return $false
    }
    if( ! $WPFcomboFolder.SelectedItem )
    {
        $null = Display-MessageBox -window $guiobject -text 'No Folder Selected' -caption 'Unable to Clone' -buttons OK -icon Error
        return $false
    }
    if( ! $WPFcomboResource.SelectedItem )
    {
        $null = Display-MessageBox -window $guiobject -text 'No Resource Selected' -caption 'Unable to Clone' -buttons OK -icon Error
        return $false
    }
    $result = $null
    [int]$clonesStartFrom = -1
    if( ! [int]::TryParse( $WPFtxtCloneStart.Text , [ref]$clonesStartFrom ) -or $clonesStartFrom -lt 0 )
    {
        $null = Display-MessageBox -window $guiobject -text 'Specified clone numbering start value is invalid' -caption 'Unable to Clone' -buttons OK -icon Error
        return $false
    }
    if( ! [int]::TryParse( $WPFtxtNumberClones.Text , [ref]$result ) -or ! $result )
    {
        $null = Display-MessageBox -window $guiobject -text 'Specified number of clones is invalid' -caption 'Unable to Clone' -buttons OK -icon Error
        return $false
    }
    if( $result -gt 1 -and $WPFtxtVMName.Text -notmatch '%d' )
    {
        $null = Display-MessageBox -window $guiobject -text 'Must specify %d replacement pattern in the VM name when creating more than one clone (ie. CloneVM-%d)' -caption 'Unable to Clone' -buttons OK -icon Error
        return $false
    }
    
    if( $result -eq 1 -and $WPFtxtVMName.Text -match '%' )
    {
        $null = Display-MessageBox -window $guiobject -text 'Illegal character ''%'' in VM name - did you mean to make more than one clone ?' -caption 'Unable to Clone' -buttons OK -icon Error
        return $false
    }
    if( $WPFtxtVMName.Text -match '[\$\*\\/]' )
    {
        $null = Display-MessageBox -window $guiobject -text 'Illegal character(s) in VM name' -caption 'Unable to Clone' -buttons OK -icon Error
        return $false
    }
    if( ( $result -eq 1 -or $WPFtxtVMName.Text -notmatch '%d' ) -and ( Get-VM -Name $WPFtxtVMName.Text -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq $WPFtxtVMName.Text } ) ) ## must check exact name match
    {
        $null = Display-MessageBox -window $guiobject -text "VM `"$($WPFtxtVMName.Text)`" already exists" -caption 'Unable to Clone' -buttons OK -icon Error
        return $false
    }
    if( $result -gt 1 )
    {
        [string[]]$existing = $null
        $clonesStartFrom..($result + $clonesStartFrom - 1) | ForEach-Object `
        {
            [string]$thisVMName = $WPFtxtVMName.Text -replace '%d' , $_
            $existing += @( Get-VM -Name $thisVMName -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq $thisVMName } | Select -ExpandProperty Name )
        }
        if( $existing -and $existing.Count )
        {
            $null = Display-MessageBox -window $guiobject -text "VMs `"$($existing -join '","')`"  already exist" -caption 'Unable to Clone' -buttons OK -icon Error
            return $false
        }
    }
    return $true
}

Function Display-MessageBox( $window , $text , $caption , [System.Windows.MessageBoxButton]$buttons , [System.Windows.MessageBoxImage]$icon )
{
    if( $window -and $window.PSObject.Properties[ 'handle' ] -and $window.Handle )
    {
        [int]$modified = switch( $buttons )
            {
                'OK' { [System.Windows.MessageBoxButton]::OK }
                'OKCancel' { [System.Windows.MessageBoxButton]::OKCancel }
                'YesNo' { [System.Windows.MessageBoxButton]::YesNo }
                'YesNoCancel' { [System.Windows.MessageBoxButton]::YesNo }
            }
        [int]$choice = [PInvoke.Win32.Windows]::MessageBox( $Window.handle , $text , $caption , ( ( $icon -as [int] ) -bor $modified ) )  ## makes it app modal so UI blocks
        switch( $choice )
        {
            ([MessageBoxReturns]::IDYES -as [int]) { 'Yes' }
            ([MessageBoxReturns]::IDNO -as [int]) { 'No' }
            ([MessageBoxReturns]::IDOK -as [int]) { 'Ok' } 
            ([MessageBoxReturns]::IDABORT -as [int]) { 'Abort' } 
            ([MessageBoxReturns]::IDCANCEL -as [int]) { 'Cancel' } 
            ([MessageBoxReturns]::IDCONTINUE -as [int]) { 'Continue' } 
            ([MessageBoxReturns]::IDIGNORE -as [int]) { 'Ignore' } 
            ([MessageBoxReturns]::IDRETRY -as [int]) { 'Retry' } 
            ([MessageBoxReturns]::IDTRYAGAIN -as [int]) { 'TryAgain' } 
        }       
    }
    else
    {
        [Windows.MessageBox]::Show( $text , $caption , $buttons , $icon )
    }
}


#endregion Functions

Remove-Module -Name Hyper-V -ErrorAction SilentlyContinue ## lest it clashes as there is some overlap in cmdlet names
[string]$oldVerbosity = $VerbosePreference
$VerbosePreference = 'SilentlyContinue'
Import-Module -Name VMware.PowerCLI -ErrorAction Stop
$VerbosePreference = $oldVerbosity


$vServer = $null


#region XAML&Modules

[string]$mainwindowXAML = @'
<Window x:Class="MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:WpfApplication1"
        mc:Ignorable="d"
        Title="ESXi VM Cloner" Height="520.399" Width="442.283" FocusManager.FocusedElement="{Binding ElementName=txtTargetComputer}">
    <Grid Margin="10,10,46,13" >
        <Grid VerticalAlignment="Top" Height="345" Margin="0,20,0,0">
            <Grid.RowDefinitions>
                <RowDefinition></RowDefinition>
                <RowDefinition></RowDefinition>
                <RowDefinition></RowDefinition>
                <RowDefinition></RowDefinition>
                <RowDefinition></RowDefinition>
                <RowDefinition></RowDefinition>
                <RowDefinition></RowDefinition>
                <RowDefinition></RowDefinition>
                <RowDefinition></RowDefinition>
                <RowDefinition></RowDefinition>
                <RowDefinition></RowDefinition>
                <RowDefinition></RowDefinition>
                <RowDefinition></RowDefinition>
                <RowDefinition></RowDefinition>
                <RowDefinition></RowDefinition>
                <RowDefinition></RowDefinition>
                <RowDefinition></RowDefinition>
                <RowDefinition></RowDefinition>
            </Grid.RowDefinitions>
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="200*"></ColumnDefinition>
                <ColumnDefinition Width="250*"></ColumnDefinition>
                <ColumnDefinition Width="90*"></ColumnDefinition>
                <ColumnDefinition Width="150*"></ColumnDefinition>
            </Grid.ColumnDefinitions>

            <TextBlock Grid.Row="0" Grid.Column="0" Text="vCenter"></TextBlock>
            <TextBlock Grid.Row="1" Grid.Column="0" Text="New VM Name"></TextBlock>
            <TextBlock Grid.Row="2" Grid.Column="0" Text="Template"></TextBlock>
            <TextBlock Grid.Row="3" Grid.Column="0" Text="VM Folder"></TextBlock>
            <TextBlock Grid.Row="4" Grid.Column="0" Text="Resource"></TextBlock>
            <TextBlock Grid.Row="8" Grid.Column="0" Text="Datastore"></TextBlock>
            <TextBlock Grid.Row="9" Grid.Column="0" Text="Notes"></TextBlock>
            <TextBlock Grid.Row="10" Grid.Column="0" Text="Number of Clones"></TextBlock>
            <TextBlock Grid.Row="11" Grid.Column="0" Text="Clones Start From"></TextBlock>
            <TextBlock Grid.Row="13" Grid.Column="0" Text="Options:"></TextBlock>
            <TextBlock Grid.Row="14" Grid.Column="0" Text=""></TextBlock>

            <TextBox x:Name="txtESXiHost" Grid.Row="0" Grid.Column="1"></TextBox>
            <Button x:Name="btnConnect" Grid.Row="0"  Grid.Column="3" Content="_Connect"></Button>
            <TextBox x:Name="txtVMName" Grid.Row="1" Grid.Column="1" Text=""></TextBox>
            <ComboBox x:Name="comboTemplate" Grid.Row="2" Grid.Column="1"/>
            <ComboBox x:Name="comboFolder" Grid.Row="3" Grid.Column="1"></ComboBox>
            <ComboBox x:Name="comboResource" Grid.Row="4" Grid.Column="1"></ComboBox>
            <ComboBox x:Name="comboDatastore" Grid.Row="8" Grid.Column="1" Text="5"></ComboBox>
            <TextBox x:Name="txtNotes" Grid.Row="9" Grid.Column="1"></TextBox>
            <TextBox x:Name="txtNumberClones" Grid.Row="10" Grid.Column="1" Text="1"></TextBox>
            <TextBox x:Name="txtCloneStart" Grid.Row="11" Grid.Column="1" Text="1"></TextBox>

            <CheckBox x:Name="chkStart" Content="_Start after Creation" Grid.Row="15" Grid.Column="1"/>
            <CheckBox x:Name="chkDisconnect" Content="_Disconnect on Exit" Grid.Row="17" Grid.Column="1"/>
        </Grid>
        <Button x:Name="btnOk" Content="OK" HorizontalAlignment="Left" Height="31" Margin="10,412,0,0" VerticalAlignment="Top" Width="90" IsDefault="True"/>
        <Button x:Name="btnCancel" Content="Cancel" HorizontalAlignment="Left" Height="31" Margin="116,412,0,0" VerticalAlignment="Top"  Width="90" IsDefault="False" IsCancel="True"/>
    </Grid>
</Window>
'@

Function Load-GUI( $inputXml )
{
    $form = $NULL
    $inputXML = $inputXML -replace 'mc:Ignorable="d"' , '' -replace 'x:N' ,'N'  -replace '^<Win.*' , '<Window'
 
    [xml]$XAML = $inputXML
 
    $reader = New-Object Xml.XmlNodeReader $xaml

    try
    {
        $Form = [Windows.Markup.XamlReader]::Load( $reader )
    }
    catch
    {
        Write-Host "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .NET is installed.`n$_"
        return $null
    }
 
    $xaml.SelectNodes('//*[@Name]') | ForEach-Object `
    {
        Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name) -Scope Global
    }

    return $form
}

#endregion XAML&Modules

[void][Reflection.Assembly]::LoadWithPartialName('Presentationframework')

$mainForm = Load-GUI $mainwindowXAML

if( ! $mainForm )
{
    return
}

if( $DebugPreference -eq 'Inquire' )
{
    Get-Variable -Name WPF*
}
## set up call backs

$WPFbtnConnect.Add_Click({
    $_.Handled = $true
    if( [string]::IsNullOrEmpty( $wpftxtESXiHost.Text ) )
    {
        Display-MessageBox -window $mainForm -text 'No ESXi Host Specified' -caption 'Unable to Connect' -buttons OK -icon Error
    }
    else
    {
        Connect-Hypervisor -GUIobject $mainForm -servers $wpftxtESXiHost.Text -pregui $false -vServer ([ref]$vServer) -username $username -password $password
    }
})
$WPFbtnOk.add_Click({
    $_.Handled = $true
    $validated = Validate-Fields -guiobject $mainForm
    if( $validated -eq $true )
    {
        $mainForm.DialogResult = $true
        $mainForm.Close()
    }
})

$WPFtxtCloneStart.Text = $startFrom
$WPFtxtNumberClones.Text = $count
$WPFchkStart.IsChecked = $powerOn
$WPFtxtESXiHost.Text = $esxihost
$WPFtxtVMName.Text = $vmName
$WPFcomboTemplate.IsEnabled = $vServer -ne $null
$WPFcomboDatastore.IsEnabled = $vServer -ne $null
$WPFcomboResource.IsEnabled = $vServer -ne $null
$WPFcomboFolder.IsEnabled = $vServer -ne $null

$mainForm.add_Loaded({
    if( $_.Source.WindowState -eq 'Minimized' )
    {
        $_.Source.WindowState = 'Normal'
    }
    $_.Handled = $true
})

$result = $mainForm.ShowDialog()
    
$disconnect = $wpfchkDisconnect.IsChecked
if( ! $result )
{
    if( $disconnect )
    {
        $vServer | Disconnect-VIServer -Confirm:$false
    }

    return
}

$chosenTemplate = Get-Template -Name $WPFcomboTemplate.SelectedItem
$chosenDatastore = Get-Datastore -Name $WPFcomboDatastore.SelectedItem
$chosenFolder = Get-Folder -Name $WPFcomboFolder.SelectedItem
$chosenName = $WPFtxtVMName.Text

if( (Get-VMHost -Name $WPFcomboResource.SelectedItem -ErrorAction Ignore))
{
    $chosenResource = Get-VMHost -Name $WPFcomboResource.SelectedItem
}
elseif( (Get-Cluster -Name $WPFcomboResource.SelectedItem -ErrorAction Ignore))
{
    $chosenResource = Get-Cluster -Name $WPFcomboResource.SelectedItem
}
elseif( (Get-ResourcePool -Name $WPFcomboResource.SelectedItem -ErrorAction Ignore))
{
    $chosenResource = Get-ResourcePool -Name $WPFcomboResource.SelectedItem
}
else
{
    $chosenResource = Get-VApp -Name $WPFcomboResource.SelectedItem
}

$count = $WPFtxtNumberClones.Text
$startFrom = $WPFtxtCloneStart.Text
$notes = $WPFtxtNotes.Text
$powerOn = $WPFchkStart.IsChecked

if( $count -gt 1 -and $chosenName -notmatch '%d' )
{
    Throw 'When creating multiple clones, the name must contain %d which will be replaced by the number of the clone'
}

[datetime]$startTime = [datetime]::Now

For( [int]$vmNumber = $startFrom ; $vmNumber -lt $startFrom + $count ; $vmNumber++ )
{
    [string]$thisChosenName = if( $count -gt 1 ) { $chosenName -replace '%d',$vmNumber } else { $chosenName }

    Write-Verbose "Creating VM # $vmNumber : $thisChosenName"

    $cloneVM = New-VM -Name $thisChosenName -Template $chosenTemplate -ResourcePool $chosenResource -Datastore $chosenDatastore -Location $chosenFolder -Notes $notes

    if( ! $cloneVM )
    {
        Throw "Failed to clone new VM `"$thisChosenName`""
    }

    if( $powerOn )
    {
        $poweredOn = Start-VM -VM $cloneVM
        if( ! $poweredOn -or $poweredOn.PowerState -ne 'PoweredOn' )
        {
            Write-Warning "Error powering on `"$($cloneVM.Name)`""
        }
    }
}

Write-Verbose "$count VMs cloned in $([math]::Round((New-TimeSpan -Start $startTime -End ([datetime]::Now)).TotalSeconds,2)) seconds"

if( $disconnect )
{
    $vServer|Disconnect-VIServer -Confirm:$false
}

$null = Read-Host 'Hit <enter> to exit'
