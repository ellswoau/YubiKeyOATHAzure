<# PowerShell tool for sequentially deploying multi YubiKeys as OATH tokens. This tool
generates unique secret keys for each key, installs them to YubiKey, and appends a
.csv that is required by Azure to upload OATH token information. 

Author: Austin Ellsworth
Last update: January 4, 2022
#>

#initialize running variable used for menu loop and selection variable to 0
$running = $true
$selection = 0
$global:scriptDir = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent

#Verify that Yubikey manager and yubico authenticator are installed on this machine.
if(Test-Path -path "C:\Program Files\Yubico\YubiKey Manager\ykman.exe") {
    $ykmanStatus = "Installed"
}
else {
    $ykmanStatus = "Not Installed!"
}
if(Test-Path -path "C:\Program Files\Yubico\Yubico Authenticator\yubioath-desktop.exe") {
    $yubicoAuthStatus = "Installed"
}
else {
    $yubicoAuthStatus = "Not Installed!"
}
if(Test-Path -path "C:\Program Files\PowerShell\7\pwsh.exe") {
    $ps7Status = "Installed"
}
else {
    $ps7Status = "Not Installed!"
}

function Show-Menu-Main
{
    param (
        [string]$Title = 'YubiKey OATH Tool'
    )
    Clear-Host
    Write-Host "===================== $Title ======================"
    Write-Output "Status: `nYkMan: $ykmanStatus      YubiCo Auth: $yubicoAuthStatus      PowerShell 7: $ps7Status `n"
    Write-Host "1: Enter '1' to deploy a YubiKey."
    Write-Host "2: Enter '2' to clear a YubiKey."
    Write-Host "3: Enter '3' to install YkMan, YubiCo Authenticator, or PowerShell 7."
    Write-Host "4: Enter '4' to open Azure MFA OATH Tokens web page."
    Write-Host "Q: Enter 'Q' to quit."
}

function Show-Menu-Installers
{
    param (
        [string]$Title = 'YubiKey Installers'
    )
    Clear-Host
    Write-Host "===================== $Title ====================="
    Write-Host "1: Enter '1' to install YkMan."
    Write-Host "2: Enter '2' to install Yubico Authenticator."
    Write-Host "3: Enter '3' to install PowerShell 7."
    Write-Host "M: Enter 'M' to return to main menu."
}

function Show-Serial-Accounts {
    #Get serial number, show it, and store it
    $global:newKeySerial = cmd.exe /c "C:\Program Files\Yubico\YubiKey Manager\ykman.exe" list --serials
    if($newKeySerial -eq $null) {
        Write-Host "No Yubi-Key Found, please connect a Yubi-Key and try again"
        $running = $false
        Start-Sleep -seconds 5
        break
    }
    else {
        Write-Host "`nConnected YubiKey serial number: $newKeySerial"
    }

    #Show slot status in yubikey, ask continue
    $newkeyPreexistingAccounts = cmd.exe /c "C:\Program Files\Yubico\YubiKey Manager\ykman.exe" oath accounts list
    if($newKeyPreexistingAccounts -eq $null) {
        Write-Host "No preexisting accounts installed on YubiKey"
    }
    else{
        Write-Host "Pre-existing oath account(s) on key, can hold two. Found account(s):`n$newKeyPreexistingAccounts `n"
    }
    Read-Host -Prompt "Press enter to continue" 
}

while($running -eq $true) {
    Show-Menu-Main 
    $selection = Read-Host "Please make a selection"
    if($selection -eq 'Q') {
        $running = $false
    }

#Check that yubikey is connected, else break with error

#Break with error if ykman or yubico auth not installed

Switch ($selection) {

    1{
        Show-Serial-Accounts
        #prompt for UPN
        $newKeyUPN = Read-Host "`nPlease enter the user principle name"

        #Generate secret key, store as variable
        $RNG = [Security.Cryptography.RNGCryptoServiceProvider]::Create()
[Byte[]]$x=1
        for($r=''; $r.length -lt 64){$RNG.GetBytes($x); if([char]$x[0] -clike '[2-7A-Z]'){$r+=[char]$x[0]}}
        $newKeySecretKey = $r
        Write-Host "Secret key generated: $newKeySecretKey"

        #install to key
        Write-Host "`Installing to Key"
        cmd.exe /c "C:\Program Files\Yubico\YubiKey Manager\ykman.exe" oath accounts add -i Microsoft $newkeyUPN $newKeySecretKey

        #check if spreadsheet exists in current folder with date, else create one, append new key
        
        #if (Test-Path "$scriptDir\YubiKey-Batch $(get-date -f MM-dd-yyyy).csv" -eq $false) {
            #create spreadsheet
            [pscustomobject] @{
            upn = $newKeyUPN;
            "serial number" = $newKeySerial;
            "secret key" = $newKeySecretKey;
            "time interval" = [int]30;
            manufacturer = "YubiKey";
            model = "OTP1"
            } | Export-Csv -UseQuotes Never -Path "$scriptDir\YubiKey-Batch $(get-date -f MM-dd-yyyy).csv" -NoTypeInformation -Append
            Read-Host "`nKey installed & csv updated, press enter to continue" 
        #}
    }

    2{
        Show-Serial-Accounts
        cmd.exe /c "C:\Program Files\Yubico\YubiKey Manager\ykman.exe" oath reset
    }

    3{
        Show-Menu-Installers
        Write-Host "Installers run from network drive, please ensure \\file-path\ is mounted before proceeding!"
        $selectionInstaller = Read-Host "Please make a selection"
        if($selectionInstaller -eq 'Q') {
            break
        }
        Switch ($selectionInstaller) {
            1{
                cmd.exe /c "Z:\Yubico\yubikey-manager-qt-1.2.4-win64.exe"
                Read-Host "Installer exited, press enter to continue"
            }
            2{
                cmd.exe /c "Z:\Yubico\yubioath-desktop-5.1.0-win64.msi"
                Read-Host "Installer exited, press enter to continue"
            }
            3{
                cmd.exe /c "Z:\PowerShell\PowerShell-7.2.1-win-x64.msi"
                Read-Host "Installer exited, press enter to continue"
            }            
            "M" {break}
        }
    }

    4{
        If(Test-Path "$scriptDir\YubiKey-Batch $(get-date -f MM-dd-yyyy).csv") {
            Write-Host `n"YubiKey sheet found at $scriptDir\YubiKey-Batch $(get-date -f MM-dd-yyyy).csv"
            Read-Host "Press enter to follow link to Azure upload page"
            cmd.exe /c "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" https://portal.azure.com/#blade/Microsoft_AAD_IAM/MultifactorAuthenticationMenuBlade/HardwareTokens
        }
        else {
            Write-Host "YubiKey spreadsheet not found, press enter to return to main page"
        }
        
    }

    Q{ 
    break 
    }
}
}