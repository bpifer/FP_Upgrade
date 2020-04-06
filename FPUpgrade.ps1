###################################
## CA Formula Pro Upgrade Script ##
###################################

# Check to see if script should run
if (($env:computername) -like "WSCA*-F*"){
exit
}
Else {

$passwd = "Bu1ldItN0w"
$UpdateFolder = "C:\Temp\Upgrade"

if (Test-Path $UpdateFolder) {
    Write-Output "$UpdateFolder exists. Skipping."
}
Else {
    Write-Output "The folder '$UpdateFolder' doesn't exist. This folder will be used for storing logs created after the script runs. Creating now."
    Start-Sleep 1
    New-Item -Path "$UpdateFolder" -ItemType Directory
    Write-Output "The folder $UpdateFolder was successfully created."
}

Function AccountInfo {
    $passwd = "Bu1ldItN0w"

    #Get computer and user account for build
    $machine = ($env:computername).Substring(0,9)+"FP"
    $account = ($env:computername).Substring(1,8)+"FP"

    $number = 1

    $computercheck = Get-PPGADObject -ObjectClass computer -Property name -value "$machine$number"

    while ($computercheck) {
    $number++
    $computercheck = Get-PPGADObject -ObjectClass computer -Property name -value "$machine$number"
    }

    "$machine$number does not exist. The PC will be renamed with this name"
    $pcName = "$machine$number"

    $usercheck = Get-PPGADObject -Filter "(&(objectClass=user)(samaccountname=$account$number))" | Select-Object -ExpandProperty samaccountname

    if($usercheck){
    "$account$number exists"
    $user = "$account$number"
    }
    ELSE{
    "$account$number does not exist. Please create the account"
    $nullUser = "$account$number does not exist. Please create the account"
    Sleep 5
    }  

    # Auto Login
    if($nullUser -ne "$account$number does not exist. Please create the account"){
        Write-host "Setting auto login"
        $autologon = "C:\PPG\FP_Update\Autologon.exe"
        $domain = "PPGNA"
        $password = "F0rward2020"
        Start-Process $autologon -ArgumentList "/accepteula",$user,$domain,$password
        }

     #Rename PC
     Write-host "Renaming PC"
    (Get-WmiObject win32_computersystem).Rename( $pcName,$passwd,'ppgna\s003968')

    #Remove Legal Notice
     Write-host "removing legal notice"
    Start-Process "C:\PPG\FP_Update\Remove_Notice.exe"
    

    #Spectro Check
$spectroFolder = "C:\ProgramData\Datacolor\DCIDriver\DCIDriver_Terminal_$env:computername"

if(Test-Path $spectroFolder) {
    Write-Output "DCIDriverDCIDriver_Terminal_$env:computername `nBacking up folder to C:\PPG\FP_Update\Spectro_Backup"
    
    $source = $spectroFolder
    $dest = "C:\PPG\FP_Update\Spectro_Backup"
    New-Item -Path $dest -ItemType Directory
    robocopy   $source $dest /COPYALL /E /SEC /R:0 /W:0 /NFL

    Write-Output "Creating folder for new PC name: $pcName"
    $source = "C:\PPG\FP_Update\Spectro_Backup"
    $dest = "C:\ProgramData\Datacolor\DCIDriver\DCIDriver_Terminal_$pcName"
    New-Item -Path $dest -ItemType Directory
    robocopy   $source $dest /COPYALL /E /SEC /R:0 /W:0 /NFL

}
Else {
    Write-Output "The folder '$spectroFolder' doesn't exist."
    }

    if ($nullUser -ne "$account$number does not exist. Please create the account"){
        # Set ACL permissions
        $DCFolder = "C:\ProgramData\Datacolor"
        If (Test-Path $DCFolder) {
        Write-Host "Setting Folder Permissions"
        $acl = Get-Acl $DCFolder
        $AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule($user,"FullControl","ContainerInherit,ObjectInherit","None","Allow")
        $acl.SetAccessRule($AccessRule)
        $acl | Set-Acl $DCFolder}
        Else {
        Write-Host "DataColor folder doesnt exist"
        }
    }

}

Start-Transcript -OutputDirectory "$UpdateFolder"

AccountInfo

#Get Environment 
$currentUser = $env:UserName
Write-Host "`nRunning this script as $currentUser `nComputer Name: $pcName `nUser Name: $user`n"
Write-Host $nullUser

#Check if FP is installed
$app = Get-WmiObject -Class Win32_Product | Where-Object { $_.Name -match "PPG FormulaPro"}
if($app.name -eq "PPG FormulaPro"){
    Write-host "FormaulaPro is installed. Backing up FP."
    cmd.exe /c "C:\PPG\FP_Update\Backup Tool\backup_Win10.bat"

    Write-host "Uninstalling FormulaPro"
    $app.Uninstall()

    Write-Host "Installing 1.5"
        cmd.exe /c '"c:\PPG\FP_Update\Formula Pro\USCA-Stores-FP-setup-150-105.exe" /v"STATION=17" /s /w /debuglog"c:\Temp\USCA-Stores-FP-setup-150-105.exe.log" /v"/qn'

$restore = "C:\EFB_DATA\Backups\Rebuild"
    if(Test-Path $restore) {
        Write-Output "$restore exists."
    }
    Else {
        Write-Output "The folder '$restore' doesn't exist."
        New-Item -Path "$restore" -ItemType Directory
    }
    Write-Host "Restore FP"
    Copy-Item "C:\EFB_DATA\Backups\CYA\CustomerData_Full.bak" -Destination "C:\EFB_DATA\Backups\Rebuild\"
    cmd.exe /c "C:\PPG\FP_Update\Restore Tool\restore.bat"
}

Else {
Write-host "FormulaPro is not installed. Installing."
        cmd.exe /c '"c:\PPG\FP_Update\Formula Pro\USCA-Stores-FP-setup-150-105.exe" /v"STATION=17" /s /w /debuglog"c:\Temp\USCA-Stores-FP-setup-150-105.exe.log" /v"/qn'
 }

# Install Misura driver
$Prima = Get-WmiObject -Class Win32_Product | where {$_.Name -like "*PPG USA Prima M16*"} 
if($Prima.name -eq "PPG USA Prima M16"){
Write-Host "Installing Misura driver"
Start-Process -FilePath "C:\PPG\FP_Update\Misura\Dromont-MisuraRetail-1.41.0.exe" -Verb runas -ArgumentList /silent }
Else {
Write-host "The PC does not have the Prima software installed" }

# Remove shortcuts
Write-Host "Removing Shortcuts"

$D_SET = "C:\Users\Public\Desktop\D_SET.EXE.lnk"
if (Test-Path $D_SET) {Remove-Item $D_SET}

$D_DSP = "C:\Users\Public\Desktop\D_DSP.EXE.lnk"
if (Test-Path $D_DSP) {Remove-Item $D_DSP}

$Paint = "C:\Users\Public\Desktop\Paint.lnk"
if (Test-Path $Paint) {Remove-Item $Paint}

# Disable DC Service
Write-Host "Checking for DC Service"
$serviceName = 'MSSQL$DCISQL2014'

If (Get-Service $serviceName -ErrorAction SilentlyContinue) {

    If ((Get-Service $serviceName).Status -eq 'Running') {

        Stop-Service $serviceName
        Write-Host "Stopping $serviceName"
        Set-Service -name $serviceName -StartupType Disabled

    } Else {
        Write-Host "$serviceName found, but it is not running."
        }

} Else {
    Write-Host "$serviceName not found"
}

#Finish
Write-host "Restarting computer"
Sleep 2
Restart-Computer -Force

Stop-Transcript
}