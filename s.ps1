# Set-ExecutionPolicy Bypass -Scope Process

$installerurl = 'https://github.com/sbrandsen/mavitec_autodesk_installer/raw/main/2022.exe'
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
        Title="Installer (last tested 2022)" Height="350" Width="400">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="*" />
            <RowDefinition Height="20"/>
            <RowDefinition Height="auto" />
        </Grid.RowDefinitions>
        <Grid Grid.Row="2" Margin="5,0">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="auto" />
                <ColumnDefinition Width="*" />
            </Grid.ColumnDefinitions>
            <Grid.RowDefinitions>
                <RowDefinition Height="auto" />
                <RowDefinition Height="10"/>
            </Grid.RowDefinitions>
            <Label Content="Local working folder path:" Grid.Row="0" VerticalAlignment="Stretch" />
            <TextBox HorizontalAlignment="Stretch" Grid.Row="0" Grid.Column="1" x:Name="workingfolderpath" VerticalAlignment="Stretch" />

        </Grid>
        <Grid Grid.Row="0">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Grid.RowDefinitions>
                <RowDefinition Height="*"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="*"/>
            </Grid.RowDefinitions>
            <Label Content="1 - Uninstall: Old Versions" FontWeight="Bold" Grid.ColumnSpan="2" HorizontalContentAlignment="Center" VerticalContentAlignment="Center" Grid.Row="0"/>
            <Button Content="Launch Uninstall Control Panel" Margin="5" x:Name="CB_Uninstall_Tool" Grid.Row="1" Grid.ColumnSpan="2"/>
            <Label Content="2 - Install 2024 release: Choose one" FontWeight="Bold" Grid.ColumnSpan="2" HorizontalContentAlignment="Center" VerticalContentAlignment="Center" Grid.Row="2"/>
            <Button Content="Full Suite" Margin="5" x:Name="CB_Deployment_Full" Grid.Row="3"/>
            <Button Content="AutoCAD LT + Vault Office" x:Name="CB_Deployment_AutoCAD_Vault" Grid.Row="3" Grid.Column="1" Margin="5"/>
            <Button Content="AutoCAD LT only" Margin="5" Grid.Row="4" x:Name="CB_Deployment_AutoCAD_LT"/>
            <Button Content="Vault Office only" Margin="5" Grid.Column="1" Grid.Row="4" x:Name="CB_Deployment_Office"/>
            <Label Content="3 - Configure: For full suite only" FontWeight="Bold" Grid.Row="5" Grid.ColumnSpan="2" HorizontalContentAlignment="Center" VerticalContentAlignment="Center" />
            <Button Content="Setup full Suite (Old Vault)"  Margin="5" Grid.Row="6" Grid.ColumnSpan="1" x:Name="CB_SetupOldVault"/>
            <Button Content="Setup full Suite (New Vault)"  Margin="5" Grid.Column="1" Grid.Row="6" Grid.ColumnSpan="1" x:Name="CB_SetupNewVault"/>
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
  
$xaml.SelectNodes("//*[@Name]") | %{"trying item $($_.Name)";
    try {Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name) -ErrorAction Stop}
    catch{throw}
    }
 
Function Get-FormVariables{
if ($global:ReadmeDisplay -ne $true){Write-host "If you need to reference this display again, run Get-FormVariables" -ForegroundColor Yellow;$global:ReadmeDisplay=$true}
get-variable WPF*
}
 
#===========================================================================
# Use this space to add code to the various form elements in your GUI
#===========================================================================
Function GetXMLLocation(){
    $appDataRoaming = $env:AppData
    $workingFoldersPath = Join-Path -Path $appDataRoaming -ChildPath "Autodesk\VaultCommon\Servers\Services_Security_6_29_2011\mavitec-vaultprod\Vaults\Vault\Objects\WorkingFolders.xml"
    return $workingFoldersPath
}

Function SetWorkingPath([string]$newpath){
    $workingFoldersPath = GetXMLLocation
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

Function Configure(){

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

    SetWorkingPath($workingfolder)

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
    iwr
    if (Test-Path $output) {
        Remove-Item $output
    }

    if(-Not (Test-Path $checkPath)) {     
        $form.Close()
        exit
    }


    $ipj = Join-Path -Path $workingfolder -ChildPath "MavitecVault.ipj"
    if (-Not (Test-Path $ipj)) {
    	Set-Clipboard -Value 'mavitec-prodvault'
        [System.Windows.MessageBox]::Show("MavitecVault.ipj not found in local working folder. GET it from the Vault`n`nServer: mavitec-vaultprod`nWas copied to your clipboard.")
        return
    
    }

    $vbs = Join-Path -Path $workingfolder -ChildPath "Application Settings\Inventor\Applications\Setup Mavitec Addin\Setup Addin.vbs"
    if (-Not (Test-Path $vbs)) {
        [System.Windows.MessageBox]::Show("Setup Addin folder not found in local working folder. GET Setup Mavitec Addin folder from the Vault")
        return
    }

    Start-Process -FilePath $vbs
    $form.Close()
    exit

}

Function InstallProducts(){
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]'Tls,Tls11,Tls12'
    $form.Cursor = [System.Windows.Input.Cursors]::Wait
    Invoke-WebRequest -Uri $installerurl -OutFile $env:temp\setup.exe
    Start-Process -FilePath $env:temp\setup.exe
    $form.Cursor = [System.Windows.Input.Cursors]::Arrow
}

Function LaunchUninstallTool(){


    $results = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | 
    Select-Object DisplayName, DisplayVersion, Publisher, InstallDate | 
    Where-Object { $_.DisplayName -like "*Autodesk*" -and $_.DisplayName -match "(\d{4})" -and ([int]$matches[1] -ge 2000 -and [int]$matches[1] -le 2023) } |
    Select-Object -ExpandProperty DisplayName

    control appwiz.cpl

    if ($results) {
        $results2 =$results -join "`n"
        Start-Sleep -Seconds 2
        [System.Windows.Forms.MessageBox]::Show("Some will not appear, this is normal. This means the program is part of another. Just uninstall anything with Autodesk and an old year name.`n`n" + $results2, "Autodesk programs to Uninstall")
    }

}

function CheckInstalledPrograms {
    Add-Type -AssemblyName System.Windows.Forms
    
    $touninstall = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | 
    Select-Object DisplayName, DisplayVersion, Publisher, InstallDate | 
    Where-Object { $_.DisplayName -like "*Autodesk*" -and $_.DisplayName -match "(\d{4})" -and ([int]$matches[1] -ge 2000 -and [int]$matches[1] -le 2099) } |
    Select-Object -ExpandProperty DisplayName

    if($touninstall){
    $touninstall = $touninstall -join "`n"
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
    InstallDeployment -ProgramName "AutoCAD LT" -ProgramFolder "2024_Autocad_LT"
}

function InstallAutoCADVault(){ 
    InstallDeployment -ProgramName "AutoCAD LT + Vault Office" -ProgramFolder "2024_Autocad_LT+Vault_Office"
}

function InstallVaultOffice(){
    InstallDeployment -ProgramName "Vault Office" -ProgramFolder "2024_Autocad_LT"
}

Function InstallFullSuite(){
    InstallDeployment -ProgramName "Full Suite" -ProgramFolder "2024"
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

    # Show notification before starting installation
    Show-TrayNotification -Title "$ProgramName Installer" -Description "$ProgramName installation started, there will be a confirmation when finished."

    # Start installation process and wait for it to finish
    $installArgs = @(
        '-i', 'deploy', '--offline_mode', '-q', '-o',
        "\\network.local\dfs\Engineering\Software\Autodesk\Deployment\$ProgramFolder\image\Collection.xml",
        '--installer_version', $InstallerVersion
    )

    $installPath = "\\network.local\dfs\Engineering\Software\Autodesk\Deployment\$ProgramFolder\image\Installer.exe"
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
$WPFCB_Deployment_AutoCAD_Vault.Add_Click({InstallAutoCADVault})  #todo
$WPFCB_Deployment_AutoCAD_LT.Add_Click({InstallAutoCADLT})  #done
$WPFCB_Deployment_Office.Add_Click({InstallVaultOffice})  #todo

$WPFCB_SetupOldVault.Add_Click({InstallAutoCADLT})  #todo
$WPFCB_SetupNewVault.Add_Click({InstallAutoCADLT})  #todo

$Form.ShowDialog() | out-null
