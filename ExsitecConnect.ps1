$vbs_server = '192.168.50.45'
$vbs_port = 2001
$Server = $env:COMPUTERNAME
$DeploymentPath = 'https://raw.githubusercontent.com/orjan-berg/ExsitecConnect/a9c2f208f6bd41d7c4ce1c8614596f982fb16359/DeploymentConfigTemplate.xml'
$appPoolName = 'ExsitecConnect'
$SitePath = 'c:\inetpub'
$User = 'IIS AppPool\ExsitecConnect'

# Define the log file path
$logFilePath = "$PSScriptRoot\logfile.log"

# Function to log messages to a file
function Write-LogMessage {
    param (
        [string]$Message
    )
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logEntry = "$timestamp - $Message"
    Add-Content -Path $logFilePath -Value $logEntry
}

# windows features and roles

$IIS = Get-WindowsFeature -ComputerName $Server -Name Web-Server
if ($IIS.Installed) {
    Write-Output 'IIS er installert!'
} else {
    Write-Output 'IIS er ikke installert, installerer nå...'
    Install-WindowsFeature -Name Web-Server -IncludeManagementTools -ComputerName $Server -Restart
}

Write-Output 'Installerer nødvendige roles and features...'
Install-WindowsFeature -ConfigurationFilePath $DeploymentPath -ComputerName $Server -Restart -Verbose
Write-Output 'Ferdig!'

# Import modules
Import-Module IISAdministration
Import-Module WebAdministration


# site mappe
if (Test-Path $SitePath\$appPoolName) {
    'Site mappe er opprettet fra før'
} else {
    'Site mappe er ikke opprettet'
    'Lager ny Site mappe'
    New-Item -ItemType Directory -Path $SitePath -Name $appPoolName
}

# Application pool

if (Test-Path IIS:\AppPools\$appPoolName) {
    'AppPool er opprettet fra før'
} else {
    'AppPool er ikke opprettet'
    'Lager ny AppPool'
    New-WebAppPool -Name $appPoolName
    Set-ItemProperty -Path IIS:\AppPools\$appPoolName managedRunTimeVersion 'V4.0'
    Set-ItemProperty -Path IIS:\AppPools\$appPoolName enable32BitAppOnWin64 'True'
}

# new website

Start-IISCommitDelay
$TestSite = New-IISSite -Name $appPoolName -BindingInformation '*:8082:' -PhysicalPath $SitePath\$appPoolName -Passthru
$TestSite.Applications['/'].ApplicationPoolName = $appPoolName
Stop-IISCommitDelay

# Application pool user

$Acl = Get-Acl $SitePath\$appPoolName
$Ar = New-Object System.Security.AccessControl.FileSystemAccessRule($User, 'Modify', 'ContainerInherit,ObjectInherit', 'None', 'Allow')
$Acl.SetAccessRule($Ar)
Set-Acl $SitePath\$appPoolName $Acl
Write-Output('Application pool bruker er opprettet')

Write-Output('Jobb utført!')

$connection = Test-NetConnection -ComputerName $vbs_server -Port $vbs_port -Verbose
if ( $connection.TcpTestSucceeded) { 
    Write-Host 'Connection to VBS is OK!'
} else {
    Write-Host 'Connection to VBS is NOT ok!'
}