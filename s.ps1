# Set-ExecutionPolicy Bypass -Scope Process

$deploymentfolder = ''

$deploymentfolder_NL = '\\network.local\dfs\Engineering\Software\Autodesk\Deployment'
$deploymentfolder_TR = 'TBT'

$foldername_full = "2024"
$foldername_vault_office = "2024_Vault_Office"
$foldername_autocad_lt= "2024_Autocad_LT"
$foldername_autocad_lt_vault_office = "2024_Autocad_LT+Vault_Office"

$firstlogonurl = "https://github.com/sbrandsen/mavitec_autodesk_installer/raw/main/MavitecFirstLogon.zip"
$workingfolderxml = 'https://raw.githubusercontent.com/sbrandsen/mavitec_autodesk_installer/main/WorkingFolders.xml' 

$inputXML = @"
<Window x:Class="playground_2.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:playground_2"
        mc:Ignorable="d"
        Title="Installer (last tested 2022)" Height="450" Width="400">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="*" />
            <RowDefinition Height="20" />
            <RowDefinition Height="auto" />
        </Grid.RowDefinitions>
        <Grid Grid.Row="2" Margin="5,0">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="auto" />
                <ColumnDefinition Width="*" />
            </Grid.ColumnDefinitions>
            <Grid.RowDefinitions>
                <RowDefinition Height="auto" />
                <RowDefinition Height="10" />
            </Grid.RowDefinitions>
            <Label Content="Local working folder path:" Grid.Row="0" VerticalAlignment="Stretch" />
            <TextBox HorizontalAlignment="Stretch" Grid.Row="0" Grid.Column="1" x:Name="workingfolderpath" VerticalAlignment="Stretch" />
        </Grid>
        <Grid Grid.Row="0">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*" />
                <ColumnDefinition Width="*" />
            </Grid.ColumnDefinitions>
            <Grid.RowDefinitions>
                <RowDefinition Height="*" />
                <RowDefinition Height="*" />
                <RowDefinition Height="*" />
                <RowDefinition Height="*" />
                <RowDefinition Height="*" />
                <RowDefinition Height="*" />
                <RowDefinition Height="*" />
                <RowDefinition Height="*" />
            </Grid.RowDefinitions>
            <Label Content="1 - Uninstall: Old Versions" FontWeight="Bold" Grid.ColumnSpan="2" HorizontalContentAlignment="Center" VerticalContentAlignment="Center" Grid.Row="0" />
            <Button Content="Launch Uninstall Control Panel" Margin="5" x:Name="CB_Uninstall_Tool" Grid.Row="1" Grid.ColumnSpan="2" />
            <Label Content="2 - Install 2024 release: Choose one" FontWeight="Bold" Grid.ColumnSpan="2" HorizontalContentAlignment="Center" VerticalContentAlignment="Center" Grid.Row="2" />
            <Label Grid.Row="3" VerticalAlignment="Top" Content="Installing from" Grid.ColumnSpan="2" HorizontalAlignment="Center" />
            <RadioButton Grid.Row="3" HorizontalAlignment="Center" Content="Netherlands"  VerticalAlignment="Bottom" GroupName="Location" IsChecked="True" x:Name="RB_Netherlands" Margin="0,0,0,8" />
            <RadioButton Grid.Row="3" Grid.Column="1" HorizontalAlignment="Center" Content="Turkey"  VerticalAlignment="Bottom" GroupName="Location" x:Name="RB_Turkey" Margin="0,0,0,8" />
            <Button Content="Full Suite" Margin="5" x:Name="CB_Deployment_Full" Grid.Row="4" />
            <Button Content="AutoCAD LT + Vault Office" x:Name="CB_Deployment_AutoCAD_Vault" Grid.Row="4" Grid.Column="1" Margin="5" />
            <Button Content="AutoCAD LT only" Margin="5" Grid.Row="5" x:Name="CB_Deployment_AutoCAD_LT" />
            <Button Content="Vault Office only" Margin="5" Grid.Column="1" Grid.Row="5" x:Name="CB_Deployment_Office" />
            <Label Content="3 - Configure: For full suite only" FontWeight="Bold" Grid.Row="6" Grid.ColumnSpan="2" HorizontalContentAlignment="Center" VerticalContentAlignment="Center" />
            <Button Content="Setup full Suite (Old Vault)"  Margin="5" Grid.Row="7" Grid.ColumnSpan="1" x:Name="CB_SetupOldVault" />
            <Button Content="Setup full Suite (New Vault)"  Margin="5" Grid.Column="1" Grid.Row="7" Grid.ColumnSpan="1" x:Name="CB_SetupNewVault" />
        </Grid>
    </Grid>
</Window>


"@ 
 
$inputXML = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N' -replace '^<Win.*', '<Window'
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = $inputXML
#Read XAML
 
$reader=(New-Object System.Xml.XmlNodeReader $xaml)
try{
    $Form=[Windows.Markup.XamlReader]::Load( $reader )
}
catch{
    Write-Warning "Unable to parse XML, with error: $($Error[0])`n Ensure that there are NO SelectionChanged or TextChanged properties in your textboxes (PowerShell cannot process them)"
    throw
}
 
#===========================================================================
# Load XAML Objects In PowerShell
#===========================================================================
  
$xaml.SelectNodes("//*[@Name]") | %{
    try {
        Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name) -ErrorAction Stop
    }
    catch {
        throw
    }
}
 
Function Get-FormVariables{
if ($global:ReadmeDisplay -ne $true){Write-host "If you need to reference this display again, run Get-FormVariables" -ForegroundColor Yellow;$global:ReadmeDisplay=$true}
get-variable WPF*
}
 
#===========================================================================
# Use this space to add code to the various form elements in your GUI
#===========================================================================
Function GetXMLLocation([string]$servername) {
    $appDataRoaming = $env:AppData
    $workingFoldersPath = Join-Path -Path $appDataRoaming -ChildPath "Autodesk\VaultCommon\Servers\Services_Security_6_29_2011\$servername\Vaults\Vault\Objects\WorkingFolders.xml"
    return $workingFoldersPath
}


Function SetWorkingPath([string]$newpath, [string] $servername){
    $workingFoldersPath = GetXMLLocation -servername $servername
    if(-Not (Test-Path -Path $workingFoldersPath)){
        $folder = $workingFoldersPath.Substring(0, $workingFoldersPath.lastIndexOf('\'))
        [System.IO.Directory]::CreateDirectory($folder)
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]'Tls,Tls11,Tls12'
    	Invoke-WebRequest -Uri $workingfolderxml -OutFile $workingFoldersPath
    }

    # Load the XML file
    $xml = [xml](Get-Content $workingFoldersPath)

    # Modify the physical path
    $xml.WorkingFolders.Folder.PhysicalPath = $newPath

    # Save the modified XML to a new file
    $xml.Save($workingFoldersPath)
}

function GetVaultVersion {
    $executable = ""
    try {
    $executable = (Get-ItemProperty -Path "HKCU:\SOFTWARE\Autodesk\Inventor\Current Version" -ErrorAction Stop).Executable
    }
    catch {
        Write-Error "Inventor Version could not be detected, is it installed? Exiting."
        Exit
    }
    $match = Select-String -InputObject $executable -Pattern "Autodesk\\Inventor (.*?)\\bin" -AllMatches | Select-Object -Expand Matches
    $version = $match.Groups[1].Value
    return $version
}

Function GetWorkingPath(){
    $workingFoldersPath = GetXMLLocation

    if(-Not (Test-Path -Path $workingFoldersPath)){
        return ""
    }

    # Load the XML file
    $xml = [xml](Get-Content $workingFoldersPath)

    # Modify the physical path
    return $xml.WorkingFolders.Folder.PhysicalPath
}

Function Configure([string] $servername){

    $detectedversion = GetVaultVersion
    $vaultloc = "C:\Program Files\Autodesk\Vault Client $detectedversion\Explorer\Connectivity.VaultPro.exe"
            
    if(-Not (Test-Path -Path $vaultloc)){
        [System.Windows.MessageBox]::Show("Could not find Vault on the C:\ drive, is it installed?")
        return
    }
   
    $workingfolder = $WPFworkingfolderpath.Text
    $workingfolder = $workingfolder -replace '"', ""
    if($workingfolder -eq ""){
        [System.Windows.MessageBox]::Show('Invalid Working folder path')
        return
    }

    if(-Not (Test-Path -Path $workingfolder)){
        [System.Windows.MessageBox]::Show('Invalid Working folder path')
        return
    }

    SetWorkingPath($workingfolder, $servername)

    $ProcessActive = Get-Process "Connectivity.VaultPro" -ErrorAction SilentlyContinue
	if($ProcessActive -eq $null){  
        Start-Process -FilePath $vaultloc 
	}
	
    $parent = [System.IO.Path]::GetTempPath()
    $output = [IO.Path]::Combine($parent, "MavitecFirstLogon.zip")
    $checkPath = [IO.Path]::Combine($parent, "MavitecFirstRun$detectedversion")

    $extractionPath = "C:\ProgramData\Autodesk\Vault $detectedversion\Extensions"

    Invoke-WebRequest $firstlogonurl -OutFile $output

    Expand-Archive -Path $output -DestinationPath $extractionPath -Force
    if (Test-Path $output) {
        Remove-Item $output
    }
    exit
}

Function LaunchUninstallTool(){


    $results = Check-InstalledAutodeskPrograms

    if ($results) {
        $dialogResult = [System.Windows.Forms.MessageBox]::Show("This will help to remove Autodesk Programs, so that everyone uses the same version`n`nPress YES: Automatically remove ALL Inventor, AutoCAD, Vault and Libraries. USE WITH CAUTION, can take a long while!`n`nPress NO: Manually remove (specific) versions, safer but slower.", "Remove Autodesk Versions", "YesNoCancel")

        switch ($dialogResult) {
            "Yes" {
                $dialogResult = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to uninstall ALL Autodesk products?", "Confirmation", "YesNo")
                switch ($dialogResult) {
                    "No" {
                        return
                    }
                }

                $count = Auto-UninstallAutodesk
                if($count -le 40){
                    [System.Windows.Forms.MessageBox]::Show($count.ToString() + " programs succesfully uninstalled!")
                } else {                   
                    [System.Windows.Forms.MessageBox]::Show("Could not automatically remove all programs, do it manually.")
                    Manual-UninstallAutodesk
                }
            }
            "No" {
                Manual-UninstallAutodesk
            }
        }

    } else {
        [System.Windows.Forms.MessageBox]::Show("No Autodesk products found. Continue to the next step")
    }

}

Function Auto-UninstallAutodesk {
    $programs = Check-InstalledAutodeskPrograms
    $count = 0
    While($programs.Count -ne 0 -and $count -le 40){
        foreach ($program in $programs) {
            $path = $program.UninstallString
            $newString = $originalString -replace '^msiexec\s', 'msiexec.exe '

            if ($path -match "AdOdis") {
                $path = $path -replace "-i", "-q -i"
            }
            if ($path -match "MsiExec") {
                $path += " /qn"
            }

            $filepath = ""
            $arguments = ""
            $firstSpaceAfterDotIndex = 0

            $dotIndex = $path.IndexOf('.')
            if ($dotIndex -gt 0) {            
                $firstSpaceAfterDotIndex = $path.IndexOf(' ', $dotIndex)
                if ($firstSpaceAfterDotIndex -lt 0) {
                    $firstSpaceAfterDotIndex = $path.Length
                }
            } else {
                $firstSpaceAfterDotIndex = $path.IndexOf(' ')
            }


            if ($firstSpaceAfterDotIndex -gt 0) {
                $arguments = $path.Substring($firstSpaceAfterDotIndex).Trim()
            }

            $filePath = $path.Substring(0, $firstSpaceAfterDotIndex)

            Write-Host "Uninstalling"$program.DisplayName
            Start-Process -FilePath $filePath -ArgumentList $arguments -Wait
            Write-Host "Uninstalled"$program.DisplayName
            Write-Host ""
                     
            $count++
            $programs = Check-InstalledAutodeskPrograms
        }
    }

    return $count
}

function Manual-UninstallAutodesk {
    control appwiz.cpl
    Start-Sleep -Milliseconds 500
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
    [System.Windows.Forms.SendKeys]::SendWait("Autodesk*20")
    $results2 = ($results | Select-Object -ExpandProperty DisplayName) -join "`n"
    Start-Sleep -Seconds 2
    [System.Windows.Forms.MessageBox]::Show("Some will not appear, this is normal. This means the program is part of another. Just uninstall anything with Autodesk and an old year name.`n`n" + $results2, "Autodesk programs to Uninstall")    
}

Function Check-InstalledAutodeskPrograms {

    $infos = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | 
            Where-object {$_.DisplayName -ne $null -and $_.SystemComponent -ne "1"} |
            select DisplayName, Publisher, DisplayVersion, UninstallString

    $infos += Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* |
            Where-object {$_.DisplayName -ne $null -and $_.SystemComponent -ne "1"} |
            select DisplayName, Publisher, DisplayVersion, UninstallString

    $infos += Get-ItemProperty HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* |
            Where-object {$_.DisplayName -ne $null -and $_.SystemComponent -ne "1"} |
            select DisplayName, Publisher, DisplayVersion, UninstallString

    return $infos | Where-Object { $_.Publisher -like "Autodesk*" -and $_.DisplayName -match '\d{4}' -and  $_.DisplayName -notlike "*Update*" }

}

function CheckInstalledPrograms {
    $touninstall = Check-InstalledAutodeskPrograms

    if($touninstall.Count -gt 0){
    $touninstall = ($touninstall | Select-Object -ExpandProperty DisplayName) -join "`n"
        $result = [System.Windows.Forms.MessageBox]::Show("There are still programs to uninstall:  `n`n$touninstall`n`nContinue anyways?", "No clean starting point found", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning) 
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            return $true
        } else {
            return $false
        }

    }

    return $true
}

function InstallAutoCADLT(){
    InstallDeployment -ProgramName "AutoCAD LT" -ProgramFolder $foldername_autocad_lt
}

function InstallAutoCADVault(){ 
    InstallDeployment -ProgramName "AutoCAD LT + Vault Office" -ProgramFolder $foldername_autocad_lt_vault_office
}

function InstallVaultOffice(){
    InstallDeployment -ProgramName "Vault Office" -ProgramFolder $foldername_autocad_lt
}

Function InstallFullSuite(){
    InstallDeployment -ProgramName "Full Suite" -ProgramFolder $foldername_full
}

function InstallDeployment {
    param (
        [string]$ProgramName,
        [string]$ProgramFolder,
        [string]$InstallerVersion = "1.39.0.174"
    )

    $answer = CheckInstalledPrograms

    if (-Not $answer) {
        return
    }

    if($WPFRB_Netherlands.IsChecked){
        $deploymentfolder = $deploymentfolder_NL
    }

    if($WPFRB_Turkey.IsChecked){
        $deploymentfolder = $deploymentfolder_TR
    }


    # Show notification before starting installation
    Show-TrayNotification -Title "$ProgramName Installer" -Description "$ProgramName installation started, there will be a confirmation when finished."

    # Start installation process and wait for it to finish
    $installArgs = @(
        '-i', 'deploy', '--offline_mode', '-q', '-o',
        "$deploymentfolder\$ProgramFolder\image\Collection.xml",
        '--installer_version', $InstallerVersion
    )

    $installPath = "$deploymentfolder\$ProgramFolder\image\Installer.exe"
    Start-Process -FilePath $installPath -ArgumentList $installArgs -Wait

    # Show notification after installation is finished
    [System.Windows.MessageBox]::Show("Installation of $ProgramName finished.", "Installation done")
}

function Show-TrayNotification {
    param (
        [string]$Title,
        [string]$Description
    )
    
    $notification = New-Object System.Windows.Forms.NotifyIcon
    $notification.Visible = $true
    
    $notification.ShowBalloonTip(5000, $Title, $Description, [System.Windows.Forms.ToolTipIcon]::Info)
    $notification.Dispose()
}

Add-Type -AssemblyName System.Windows.Forms
$WPFCB_Uninstall_Tool.Add_Click({LaunchUninstallTool})

$WPFCB_Deployment_Full.Add_Click({InstallFullSuite})  #todo
$WPFCB_Deployment_AutoCAD_Vault.Add_Click({InstallAutoCADVault})  #done
$WPFCB_Deployment_AutoCAD_LT.Add_Click({InstallAutoCADLT})  #done
$WPFCB_Deployment_Office.Add_Click({InstallVaultOffice})  #done

$WPFCB_SetupOldVault.Add_Click({Configure -servername "mavitec-vault"})  #done
$WPFCB_SetupNewVault.Add_Click({Configure -servername "mavitec-vaultprod"})  #done

$Form.ShowDialog() | out-null
