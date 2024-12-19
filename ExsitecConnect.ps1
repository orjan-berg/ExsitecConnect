$vbs_server = "192.168.50.190"
$vbs_port = 2001
$Server = $env:COMPUTERNAME
$DeploymentPath = "\\192.168.50.83\Shared\Jobb\ps_scripts\DeploymentConfigTemplate.xml"
$appPoolName = 'ExsitecConnect'
$SitePath = "c:\inetpub"
$User ='IIS AppPool\ExsitecConnect'

# windows features and roles

$IIS = get-windowsfeature -ComputerName $Server -Name Web-Server
if ($IIS.Installed)
{
    Write-Output "IIS er installert!"
}
else
{
    Write-Output "IIS er ikke installert, installerer nå..."
    Install-WindowsFeature -name Web-Server -IncludeManagementTools -ComputerName $Server -Restart
}

Write-Output "Installerer nødvendige roles and features..."
Install-WindowsFeature -ConfigurationFilePath $DeploymentPath -ComputerName $Server -Restart -Verbose
Write-Output "Ferdig!"

# Import modules
Import-Module IISAdministration
Import-Module WebAdministration


# site mappe
if(Test-Path $SitePath\$appPoolName)
{
"Site mappe er opprettet fra før"
}
else
{
"Site mappe er ikke opprettet"
"Lager ny Site mappe"
New-item -ItemType Directory -Path $SitePath -Name $appPoolName
}

# Application pool

if(Test-Path IIS:\AppPools\$appPoolName)
{
"AppPool er opprettet fra før"
}
else
{
"AppPool er ikke opprettet"
"Lager ny AppPool"
New-WebAppPool -Name $appPoolName
Set-ItemProperty -Path IIS:\AppPools\$appPoolName managedRunTimeVersion "V4.0"
Set-ItemProperty -Path IIS:\AppPools\$appPoolName enable32BitAppOnWin64 "True"
}

# new website

Start-IISCommitDelay
$TestSite = New-IISSite -Name $appPoolName -BindingInformation "*:8082:" -PhysicalPath $SitePath\$appPoolName -Passthru
$TestSite.Applications["/"].ApplicationPoolName = $appPoolName
Stop-IISCommitDelay

# Application pool user

$Acl = Get-Acl $SitePath\$appPoolName
$Ar = New-Object System.Security.AccessControl.FileSystemAccessRule($User,"Modify", "ContainerInherit,ObjectInherit","None", "Allow")
$Acl.SetAccessRule($Ar)
Set-Acl $SitePath\$appPoolName $Acl
Write-Output("Application pool bruker er opprettet")

Write-Output("Jobb utført!")

$connection = Test-NetConnection -ComputerName $vbs_server -port $vbs_port -Verbose
if ( $connection.TcpTestSucceeded) 
{ 
    write-host "Connection to VBS is OK!"
}
else
{
    write-host "Connection to VBS is NOT ok!"
}