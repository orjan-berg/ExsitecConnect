$vbs_server = '192.168.50.45'
$vbs_port = 2001
$Server = $env:COMPUTERNAME
$FeaturesToInstallPath = 'https://raw.githubusercontent.com/orjan-berg/ExsitecConnect/refs/heads/main/FeaturesToInstall.csv'
$appPoolName = 'ExsitecConnect'
$SitePath = 'c:\inetpub'
$User = 'IIS AppPool\ExsitecConnect'

# Define the log file path
$logFilePath = '.\logfile.log'

# Function to log messages to a file
function Write-LogMessage {
    param (
        [string]$Message
    )
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logEntry = "$timestamp - $Message"
    Add-Content -Path $logFilePath -Value $logEntry
}


Write-LogMessage 'Laster ned FeaturesToInstall.csv'
$FeaturesToInstallFile = Invoke-WebRequest $FeaturesToInstallPath
$FeaturesToInstallFile.Content | Out-File '.\FeaturesToInstall.csv'
Write-LogMessage 'Nedlasting fullført'



# windows features and roles

$IIS = Get-WindowsFeature -ComputerName $Server -Name Web-Server
if ($IIS.Installed) {
    Write-LogMessage 'IIS er installert!'
} else {
    Write-LogMessage 'IIS er ikke installert, installerer nå...'
    Install-WindowsFeature -Name Web-Server -IncludeManagementTools -ComputerName $Server -Restart
}

$RequiredFeatures = Import-Csv -Path '.\FeaturesToInstall.csv'

$WindowsFeatures = Get-WindowsFeature -Name Web-*, Net-*

# Funksjon for å sjekke og installere manglende funksjoner
foreach ($feature in $RequiredFeatures) {
    $isInstalled = $WindowsFeatures | Where-Object { $_.Name -eq $feature -and $_.Installed -eq $true }

    if (-not $isInstalled) {
        Write-Host "Funksjonen $feature er ikke installert. Forsøker å installere..."
        Install-WindowsFeature -Name $feature -IncludeManagementTools -ErrorAction Stop
        Write-Host "Funksjonen $feature ble installert."
    } else {
        Write-Host "Funksjonen $feature er allerede installert."
    }
}

# Import modules
Import-Module IISAdministration
Import-Module WebAdministration


# site mappe
if (Test-Path $SitePath\$appPoolName) {
    Write-LogMessage 'Site mappe er opprettet fra før'
} else {
    Write-LogMessage 'Site mappe er ikke opprettet'
    Write-LogMessage 'Lager ny Site mappe'
    New-Item -ItemType Directory -Path $SitePath -Name $appPoolName
}

# Application pool

if (Test-Path IIS:\AppPools\$appPoolName) {
    Write-LogMessage 'AppPool er opprettet fra før'
} else {
    Write-LogMessage 'AppPool er ikke opprettet'
    Write-LogMessage 'Lager ny AppPool'
    New-WebAppPool -Name $appPoolName
    Set-ItemProperty -Path IIS:\AppPools\$appPoolName managedRunTimeVersion 'V4.0'
    Set-ItemProperty -Path IIS:\AppPools\$appPoolName enable32BitAppOnWin64 'True'
}

# new website
Write-LogMessage 'Start-IISCommitDelay'
Start-IISCommitDelay
$TestSite = New-IISSite -Name $appPoolName -BindingInformation '*:8082:' -PhysicalPath $SitePath\$appPoolName -Passthru
$TestSite.Applications['/'].ApplicationPoolName = $appPoolName
Stop-IISCommitDelay
Write-LogMessage 'Stop-IISCommitDelay'

# Application pool user

$Acl = Get-Acl $SitePath\$appPoolName
$Ar = New-Object System.Security.AccessControl.FileSystemAccessRule($User, 'Modify', 'ContainerInherit,ObjectInherit', 'None', 'Allow')
$Acl.SetAccessRule($Ar)
Set-Acl $SitePath\$appPoolName $Acl
Write-LogMessage('Application pool bruker er opprettet')

Write-LogMessage('Jobb utført!')

$connection = Test-NetConnection -ComputerName $vbs_server -Port $vbs_port -Verbose
if ( $connection.TcpTestSucceeded) { 
    Write-LogMessage 'Connection to VBS is OK!'
} else {
    Write-LogMessage 'Connection to VBS is NOT ok!'
}

