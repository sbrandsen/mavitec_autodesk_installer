# Set-ExecutionPolicy Bypass -Scope Process

$installerurl = 'https://github.com/sbrandsen/mavitec_autodesk_installer/raw/main/2022.exe'
$inputXML = @"
<Window x:Class="playground_2.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:playground_2"
        mc:Ignorable="d"
        Title="Installer (last tested 2022)" Height="250" Width="400">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="*" />
            <RowDefinition Height="100px" />
        </Grid.RowDefinitions>
        <Grid>
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="auto" />
                <ColumnDefinition Width="*" />
            </Grid.ColumnDefinitions>
            <Grid.RowDefinitions>
                <RowDefinition Height="auto" />
            </Grid.RowDefinitions>
            <Label Content="Local working folder path:" Grid.Row="0" />
            <TextBox HorizontalAlignment="Stretch" Grid.Row="0" Grid.Column="1" x:Name="workingfolderpath" />
        </Grid>
        <StackPanel Grid.Row="1" Orientation="Horizontal" HorizontalAlignment="Center">
            <Button Content="Install 2022 Suite" Padding="15" Margin="5,0" x:Name="CB_Install"/>
            <Button Content="Configure" Padding="40,10,40,10" Margin="5,0" x:Name="CB_Configure"/>
        </StackPanel>
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
    	Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/sbrandsen/mavitec_autodesk_installer/main/WorkingFolders.xml' -OutFile $workingFoldersPath
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
        [System.Windows.MessageBox]::Show("Could not find Vault $detectedversion on the C:\ drive, is it installed?")
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

    $url = "https://github.com/sbrandsen/mavitec_autodesk_installer/raw/main/MavitecFirstLogon.zip"
    $parent = [System.IO.Path]::GetTempPath()
    $output = [IO.Path]::Combine($parent, "MavitecFirstLogon.zip")
    $checkPath = [IO.Path]::Combine($parent, "MavitecFirstRun")

    $extractionPath = "C:\ProgramData\Autodesk\Vault 2022\Extensions"

    Invoke-WebRequest $url -OutFile $output

    Expand-Archive -Path $output -DestinationPath $extractionPath -Force

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

$WPFCB_Install.Add_Click({InstallProducts}) 
$WPFCB_Configure.Add_Click({Configure})    

$Form.ShowDialog() | out-null
