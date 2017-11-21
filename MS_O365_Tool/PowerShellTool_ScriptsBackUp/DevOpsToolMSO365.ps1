############################################################################################################
#   ArcSigt DevOps Team Costa Rica Tool to Testing MS Office 365 througth PowerShell                       #
#   Made by Andres Duran Campos DevOps CR andres.duran-campos@hpe.com                                      #
#   This Tool is for internal use only to testing MS O365                                                  #
#                                                                                                          #
#                                                                                                          #
#                                                                                                          #
############################################################################################################
<#  
.SYNOPSIS
   	Downloads, installs, and has connection options to O365 and online tenant information via PowerShell.

.DESCRIPTION  
    Installs the required modules to access O365 tenant information via PowerShell.
                            
    
    Rights Required		: Local admin on workshop for installing applications
                        : Set-ExecutionPolicy to 'Unrestricted' for the .ps1 file to execute the installs
                        : Requires PowerShell (or ISE) to 'Run as Administrator' to install the applications or modules

    

.FUNCTIONALITY
   This script displays options that simplify the process of installing the pre-requisites needed for 
   logging onto individual O365 components: Exchange online, Active Derectory, OneDrive and SharePoint. 
#>

param(
	[parameter(ValueFromPipeline=$false, ValueFromPipelineByPropertyName=$true, Mandatory=$false)] 
	[string] $TargetFolder = "C:\Install" ,
	[parameter(ValueFromPipeline=$false, ValueFromPipelineByPropertyName=$false, Mandatory=$false)] 
	[bool] $HasInternetAccess = ([Activator]::CreateInstance([Type]::GetTypeFromCLSID([Guid]'{DCB00C01-570F-4A9B-8D69-199FDBA5723B}')).IsConnectedToInternet)
)

#region Detect PS version and 64-bit OS

# Check for 64-bit OS
If($env:PROCESSOR_ARCHITECTURE -match '86') {
        Write-Host "`nThis script only installs 64-bit modules. This machine is not a 64-bit Operating System.`n" -ForegroundColor Red

} # End Check for 64-bit OS

# Check for PowerShell version compatibility
If (($PSVersionTable.PSVersion).Major -le 2.0) {
       Write-Host "`nThis script requires a version of PowerShell 3 or higher, which this is not.`n" -ForegroundColor Red
       Write-Host "PS 3.0: https://www.microsoft.com/en-us/download/details.aspx?id=34595" -ForegroundColor Yellow
       Write-Host "PS 4.0: https://www.microsoft.com/en-us/download/details.aspx?id=40855" -ForegroundColor Yellow
       Write-Host "PS 5.0: https://www.microsoft.com/en-us/download/details.aspx?id=48729" -ForegroundColor Yellow
       Write-Host "Please review the System Requirements to decide which version to install onto your computer.`n" -ForegroundColor Cyan
       Exit
} # End Check for PowerShell version compatibility

#endregion End Detect PS version and 64-bit OS

Clear-Host

#region Menu display using here string
[string] $menu = @'

	*******************************************************************
	    DevOps Team ArcSight Logon/Test/Install O365 Services Tool
	*******************************************************************
	
	Please select an option from the list below:


     1) Log onto MS Azure Active Directory
     2) Send Email with Exchange Online 
     3) Log onto Exchange Online
     4) Log onto OneDrive Online and check The Site
     5) Log onto SharePoint/OneDrive Online and Upload/Download Files
	 

     10) Launch Windows Update
    	
     11) Install - .NET 4.5.2
     12) Install - MS Online Service Sign-In Assistance for IT Professionals RTW
     13) Install - Windows Azure Active Directory Module
     14) Install - SharePoint Online Module (Reboot required)
     15) Install - Skype for Business Online Module (Reboot required)
     
     20) Install - O365_Logon Module 1.0

     30) Enable PS Remoting on this local computer (Fix WinRM issue)

     31) Launch PowerShell 3.0 download website
     32) Launch PowerShell 4.0 download website
     33) Launch PowerShell 5.0 download website

     90) Launch SharePoint Online Management Shell download website
     91) Launch OneDrive PowerShell Gallery Module How to Install
	
     98) Restart this workstation
     99) Exit this script

Select an option.. [1-99]?
'@

#endregion Menu display

#region Installs

Function TestTargetPath { # Test for target path for install temporary directory.
            If ((Test-Path $targetfolder) -eq $true) {
			    Write-Host "Folder: $targetfolder exists." -ForegroundColor Green
		    } 
            Else{
			    Write-Host "Folder: $targetfolder does not exist, creating..." -NoNewline
			    New-Item $targetfolder -type Directory | Out-Null
			    Write-Host "created!" -ForegroundColor Green
            } 
}# End Test for target path for install temporary directory.

function Install-DotNET452{
	$val = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" -Name "Release"
	if ($val.Release -lt "379893") {
    		GetIt "http://download.microsoft.com/download/E/2/1/E21644B5-2DF2-47C2-91BD-63C560427900/NDP452-KB2901907-x86-x64-AllOS-ENU.exe"
	    	Set-Location $targetfolder
    		[string]$expression = ".\NDP452-KB2901907-x86-x64-AllOS-ENU.exe /quiet /norestart /l* $targetfolder\DotNET452.log"
	    	Write-Host "File: NDP452-KB2901907-x86-x64-AllOS-ENU.exe installing..." -NoNewLine
    		Invoke-Expression $expression
    		Start-Sleep -Seconds 20
    		Write-Host "`n.NET 4.5.2 should be installed by now." -Foregroundcolor Yellow
	} else {
    		Write-Host "`n.NET 4.5.2 already installed." -Foregroundcolor Green
    }
} # end Install .NET 4.5.2

#region Install Windows Azure Active Directory module

Function Check-SIA_Installed { # Check for Sign In Assistant before WaaD can install
        $CheckForSignInAssistant = Test-Path "HKLM:\SOFTWARE\Microsoft\MSOIdentityCRL"
        If ($CheckForSignInAssistant -eq $true) {
                $SignInAssistantVersion = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\MSOIdentityCRL"
                Write-Host "`Sign In Assistant version"$SignInAssistantVersion.MSOIDCRLVersion"is installed" -Foregroundcolor Green
                Install-WAADModule
        }
        Else {
                Write-Host "Windows Azure Active Directory Module stopping installation...`n" -Foregroundcolor Green 
                Write-Host "`nThe Sign In Assistant needs to be installed before the Windows Azure Active Directory module.`n" -Foregroundcolor Red   
        } 
} # End Check for Sign In Assistant before WaaD can install

Function Install-WAADModule {
            Check-Bits #Confirms if BitsTransfer is running on the local host
            $WAADUrl = "https://bposast.vo.msecnd.net/MSOPMW/Current/amd64/AdministrationConfig-FR.msi"
            Start-BitsTransfer -Source $WAADUrl -Description "Windows Azure Active Directory" -Destination $env:temp -DisplayName "Windows Azure Active Directory"
            Start-Process -FilePath msiexec.exe -ArgumentList "/i $env:temp\$(Split-Path $WAADUrl -Leaf) /quiet /passive"
            Start-Sleep -Seconds 5
            $LoopError = 1 # Variable to error out the loop
            Do {$CheckForWAAD = Test-Path "$env:windir\System32\WindowsPowerShell\v1.0\Modules\MSOnline"
                Write-Host "Windows Azure Active Directory Module being installed..." -Foregroundcolor Green
                Start-Sleep -Seconds 10
                $LoopError = $LoopError + 1
            }
            Until ($CheckForWAAD -eq $true -or $LoopError -eq 10)
                Start-Sleep -Seconds 5
                If ($CheckForWAAD -eq $true){
                        $WaaDModuleVersion = (get-item C:\Windows\System32\WindowsPowerShell\v1.0\Modules\MSOnline\Microsoft.Online.Administration.Automation.PSModule.dll).VersionInfo.FileVersion
                        Write-Host "`nWindows Azure Active Directory Module version $WaaDModuleVersion is now installed." -Foregroundcolor Green  
                }
                Else {
                        Write-Host "`nAn error may have occured. Windows Azure Active Directory online module could be installed or is still installing. Rerun this step to confirm." -ForegroundColor Red
                }
}

Function Install-WindowsAADModule {
        $CheckForWAAD = Test-Path "$env:windir\System32\WindowsPowerShell\v1.0\Modules\MSOnline"
        If ($CheckForWAAD -eq $false){
            Write-Host "`nWindows Azure Active Directory Module starting installation...`n" -Foregroundcolor Green
            Check-SIA_Installed
        }
        Else {
            $WaaDModuleVersion = (get-item C:\Windows\System32\WindowsPowerShell\v1.0\Modules\MSOnline\Microsoft.Online.Administration.Automation.PSModule.dll).VersionInfo.FileVersion
            If ($WaaDModuleVersion -ge "1.0.8070.2"){
                Write-Host "`nWindows Azure Active Directory Module version $WaaDModuleVersion already installed." -Foregroundcolor Green
            }
            Else {
                Write-Host "`nWindows Azure Active Directory Module version $WaaDModuleVersion already installed." -Foregroundcolor Green
                Write-Host "However, there is a newer version available for download." -ForegroundColor Yellow
                Write-Host "You will need to uninstall your current version and re-install a newer version." -ForegroundColor Yellow
            }            
        }
} #endregion Install Windows Azure Active Directory module

#region Install Sign in Assistant (SIA)
Function Install-SIA {
              Check-Bits #Confirms if BitsTransfer is running on the local host
              $MsolUrl = "http://download.microsoft.com/download/5/0/1/5017D39B-8E29-48C8-91A8-8D0E4968E6D4/en/msoidcli_64.msi"
              Start-BitsTransfer -Source $MsolUrl -Description "Microsoft Online services" -Destination $env:temp -DisplayName "Microsoft Online Services"
              Start-Process -FilePath msiexec.exe -ArgumentList "/i $env:temp\$(Split-Path $MsolUrl -Leaf) /quiet /passive"
              Start-Sleep -Seconds 10
              $LoopError = 1 # Variable to error out the loop
              Do {$CheckForSignInAssistant = Test-Path "HKLM:\SOFTWARE\Microsoft\MSOIdentityCRL"
                    Write-Host "Sign In Assistant being installed..." -Foregroundcolor Green
                    Start-Sleep -Seconds 10
                    $LoopError = $LoopError + 1
              }
              Until ($CheckForSignInAssistant -eq $true -or $LoopError -eq 10)
                    Start-Sleep -Seconds 10
                    If ($CheckForSignInAssistant -eq $true){
                            $SignInAssistantVersion = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\MSOIdentityCRL"
                            Write-Host "`nSign In Assistant version"$SignInAssistantVersion.MSOIDCRLVersion"is now installed." -Foregroundcolor Green  
                                       
                    }
                Else {
                        Write-Host "`nAn error may have occured. The Sign In Assistant could be installed or still installing. Rerun this step to confirm." -ForegroundColor Red
                }
                    
                    
}

Function Install-SignInAssistant {
    $CheckForSignInAssistant = Test-Path "HKLM:\SOFTWARE\Microsoft\MSOIdentityCRL"
        If ($CheckForSignInAssistant -eq $false) {
        Write-Host "`nSign In Assistant starting installation...`n" -Foregroundcolor Green
        Install-SIA
        }
            Else {
            $SignInAssistantVersion = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\MSOIdentityCRL"
                If ($SignInAssistantVersion.MSOIDCRLVersion -lt "7.250.4551.0") {
                    Write-Host "`nSign In Assistant starting installation...`n" -Foregroundcolor Green
                    Install-SIA
                }
                Else {
                    Write-Host "`nSign In Assistant version"$SignInAssistantVersion.MSOIDCRLVersion"is already installed" -Foregroundcolor Green
                }
        }
} #endregion Install Sign in Assistant

#region Install Skype for Business Module
Function Install-SfbOModule {
            GetIt "https://download.microsoft.com/download/2/0/5/2050B39B-4DA5-48E0-B768-583533B42C3B/SkypeOnlinePowershell.exe"
            Set-Location $targetfolder
            [string]$expression = ".\SkypeOnlinePowershell.exe /quiet /norestart /l* $targetfolder\SkypeOnlinePowerShell.log"
            Write-Host "Skype for Business online starting installation...`n" -NoNewLine -ForegroundColor Green
            Invoke-Expression $expression
    		Start-Sleep -Seconds 5
            $LoopError = 1 # Variable to error out the loop
            Do {$CheckForSfbO = Test-Path "$env:ProgramFiles\Common Files\Skype for business Online\Modules"
                Write-Host "Skype for Business online module being installed..." -Foregroundcolor Green
                Start-Sleep -Seconds 15
                $LoopError = $LoopError + 1
            }
            Until ($CheckForSfbO -eq $true -or $LoopError -eq 10)
    		    If ($CheckForSfbO -eq $true){
                    Start-Sleep -Seconds 10
                    If ($CheckForSfbO -eq $True) {
                        Write-Host "Skype for Business online module now installed." -Foregroundcolor Green
                    }
                    Else {                {
                        Write-Host "`nAn error may have occured. Skype for Business online module could be installed or is still installing. Rerun this step to confirm." -ForegroundColor Red
                }
                Write-Host "             Reboot eventually needed before this module will work.               " -BackgroundColor Red -ForegroundColor Black
            }
            }
}

Function Install-SfbO {
        $CheckForSfbO = Test-Path "$env:ProgramFiles\Common Files\Skype for business Online\Modules"
        If ($CheckForSfbO -eq $false){
            Install-SfboModule
        }
        Else {
            Write-Host "`nSkype for Business Online Module already installed`n" -Foregroundcolor Green
        }
 } #endregion Install Skype for Business Module

#region Install SharePoint Module
Function Install-SPOModule {
              Check-Bits #Confirms if BitsTransfer is running on the local host
              $MsolUrl = "http://blogs.technet.com/cfs-filesystemfile.ashx/__key/telligent-evolution-components-attachments/01-9846-00-00-03-65-75-65/sharepointonlinemanagementshell_5F00_4613_2D00_1211_5F00_x64_5F00_en_2D00_us.msi"
              Start-BitsTransfer -Source $MsolUrl -Description "SharePoint Online Module" -Destination $env:temp -DisplayName "SharePoint Online Module"
              Start-Process -FilePath msiexec.exe -ArgumentList "/i $env:temp\$(Split-Path $MsolUrl -Leaf) /quiet /passive"
              
              #Logic to confirm that the file downloaded to local client. If not, then launch to website for manual download.
                    $CheckForSPOFileDownload = Test-Path "$env:temp\sharepointonlinemanagementshell_5F00_4613_2D00_1211_5F00_x64_5F00_en_2D00_us.msi"
                    If ($CheckForSPOFileDownload -eq $false) { # Install calls download website for install if download file does not exist
                        Start-Process "https://www.microsoft.com/en-us/download/details.aspx?id=35588"
                    }
                    Else {
                        Start-Sleep -Seconds 5
                        $LoopError = 1 # Variable to error out the loop
                        Do {$CheckForSPO = Test-Path "$env:ProgramFiles\SharePoint Online Management Shell"
                            Write-Host "SharePoint Online module being installed..." -Foregroundcolor Green
                            Start-Sleep -Seconds 5
                            $LoopError = $LoopError + 1
                        }
                        Until ($CheckForSPO -eq $true -or $LoopError -eq 10)
                            Start-Sleep -Seconds 10
                            If ($CheckForSPO -eq $true){
                            Write-Host "`nSharePoint online module installation is now complete." -Foregroundcolor Green 
                            }
                            Else {
                                Write-Host "`nAn error may have occured. SharePoint online module could be installed. Rerun this step to confirm." -ForegroundColor Red
                            }
                    }
}

Function Install-SPO {
        $CheckForSPO = Test-Path "$env:ProgramFiles\SharePoint Online Management Shell"
        If ($CheckForSPO -eq $false){
             Install-SPOModule
             Write-Host "             Reboot eventually needed before this module will work.             " -BackgroundColor Red -ForegroundColor Black
        }
        Else {
             Write-Host "`nSharePoint Online Module already installed." -Foregroundcolor Green
        }

} #endregion Install SharePoint Module

#region Install O365_Logon Module

# O365_Logon Module download and extraction
Function Install-O365_LogonModule {  
        $url = "https://gallery.technet.microsoft.com/scriptcenter/O365Logon-Module-a1d9baf2/file/145181/4/O365_Logon.zip" # MS Script Center location
        $output = $env:TEMP
        Import-Module BitsTransfer  
        Start-BitsTransfer -Source $url -Destination $output

function Expand-ZIPFile($file, $destination) {
        $shell = new-object -com shell.application
        $zip = $shell.NameSpace($file)
    foreach($item in $zip.items())
    {
        $shell.Namespace($destination).copyhere($item)
    }
} 

Expand-ZIPFile –File “$output\O365_Logon.zip” –Destination “$env:windir\System32\WindowsPowerShell\v1.0\Modules\”
        Start-Sleep -Seconds 2
        $LoopError = 1 # Variable to error out the loop
        Do {$CheckForO365_LogonModule = Test-Path "$env:windir\System32\WindowsPowerShell\v1.0\Modules\O365_Logon"
            Write-Host "O365_Logon Module being installed..." -Foregroundcolor Green
            Start-Sleep -Seconds 2
            $LoopError = $LoopError + 1
        } 
        Until ($CheckForO365_LogonModule -eq $true -or $LoopError -eq 10)
            Start-Sleep -Seconds 2
            If ($CheckForO365_LogonModule -eq $true){
                Write-Host "`nO365_Logon Module now installed." -Foregroundcolor Green 
            }
            Else {
                Write-Host "`nAn error may have occured. O365_Logon module could be installed or is still installing. Rerun this step to confirm." -ForegroundColor Red
            }
            
} # End O365_Logon Module download and extraction

# Install O365_Logon module logic to check if already installed or needs to install 
Function Install-O365_Logon {
                If (((Test-Path "$env:windir\System32\WindowsPowerShell\v1.0\Modules\O365_Logon") -or (Test-Path "$env:USERPROFILE\Documents\WindowsPowerShell\Modules\O365_Logon") -or (Test-Path "$env:ProgramFiles\WindowsPowerShell\Modules\O365_Logon")) -eq $true){
                        Write-Host "`nO365_Logon Module already installed." -Foregroundcolor Green    
                }      
                Else {
                        Write-Host "O365_Logon Module starting installaion...`n" -Foregroundcolor Green
                        Install-O365_LogonModule  
                }  # End Install O365_Logon module logic to check if already installed or needs to install
} #endregion End Install O365_Logon Module


# Get-Bits function
Function Check-Bits{
    if ((Get-Module BitsTransfer) -eq $null){
			Write-Host "BitsTransfer: Installing..." -NoNewLine
			Import-Module BitsTransfer	
			Write-Host "Installed." -ForegroundColor Green
		}
} # End Get-Bits Function

# GetIt Module
function GetIt ([string]$sourcefile)	{
	if ($HasInternetAccess){
		Check-Bits # check if BitsTransfer is installed
		[string] $targetfile = $sourcefile.Substring($sourcefile.LastIndexOf("/") + 1) 
		TestTargetPath # Function to confirn or create the $Targetpath download for installable file
		if (Test-Path "$targetfolder\$targetfile"){
			Write-Host "File: $targetfile exists."
		}else{	
			Write-Host "File: $targetfile does not exist, downloading..." -NoNewLine
			Start-BitsTransfer -Source "$SourceFile" -Destination "$targetfolder\$targetfile"
			Write-Host "Downloaded." -ForegroundColor Green
		}
	}else{
		Write-Host "Internet Access not detected. Please resolve and try again." -foregroundcolor red
	}
} # End GetIt Module function

# Unzip function
function UnZipIt ([string]$source, [string]$target){
	if (Test-Path "$targetfolder\$target"){
		Write-Host "File: $target exists."
	}else{
		Write-Host "File: $target doesn't exist, unzipping..." -NoNewLine
		$sh = new-object -com shell.application
		$zipfolder = $sh.namespace("$targetfolder\$source") 
		$item = $zipfolder.parsename("$target")      
		$targetfolder2 = $sh.namespace("$targetfolder")       
		Set-Location $targetfolder
		$targetfolder2.copyhere($item)
		Write-Host "`b`b`b`b`b`b`b`b`b`b`b`bunzipped!   " -ForegroundColor Green
		Remove-Item $source
	}
} # End UnZipIt Function

# New-FileDownload function
function New-FileDownload {
	param (
	[parameter(ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, Mandatory=$true, HelpMessage="No source file specified")] 
	[string]$SourceFile,
    [parameter(ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, Mandatory=$false, HelpMessage="No destination folder specified")] 
    [string]$DestFolder,
    [parameter(ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, Mandatory=$false, HelpMessage="No destination file specified")] 
    [string]$DestFile
	)
	$error.clear()
	if (!($DestFolder)){$DestFolder = $TargetFolder}
	Get-ModuleStatus -name BitsTransfer
	if (!($DestFile)){[string] $DestFile = $SourceFile.Substring($SourceFile.LastIndexOf("/") + 1)}
	if (Test-Path $DestFolder){
		Write-Host "Folder: `"$DestFolder`" exists."
	} else{
		Write-Host "Folder: `"$DestFolder`" does not exist, creating..." -NoNewline
		New-Item $DestFolder -type Directory
		Write-Host "Done! " -ForegroundColor Green
	}
	if (Test-Path "$DestFolder\$DestFile"){
		Write-Host "File: $DestFile exists."
	}else{
		if ($HasInternetAccess){
			Write-Host "File: $DestFile does not exist, downloading..." -NoNewLine
			Start-BitsTransfer -Source "$SourceFile" -Destination "$DestFolder\$DestFile"
			Write-Host "Done! " -ForegroundColor Green
		}else{
			Write-Host "Internet access not detected. Please resolve and try again." -ForegroundColor red
		}
	}
} # End New-FileDownload funtion

# Function Get-ModuleStatus
function Get-ModuleStatus { 
	param	(
		[parameter(ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, Mandatory=$true, HelpMessage="No module name specified!")] 
		[string]$name
	)
	if(!(Get-Module -name "$name")) { 
		if(Get-Module -ListAvailable | ? {$_.name -eq "$name"}) { 
			Import-Module -Name "$name" 
			# module was imported
			return $true
		} else {
			# module was not available
			return $false
		}
	}else {
		# module was already imported
		# Write-Host "$name module already imported"
		return $true
	}
} # End Function Get-ModuleStatus function

#endregion Installs

#region Common used scriptblocks

# Enter Credentials to log onto tenant scriptblock
        $Global:UserCredential = {
            $Global:Credential = Get-Credential -Message "Your logon is your e-mail address for O365."
        } # End Enter Credentials to log onto tenant scriptblock

# Start logic to confirm if logged on user has access to MS online module. Then present users' information as confirmation. 
        $Global:MSolUserScriptBlock = {
            $Global:MSolUser =  Get-MsolUser -UserPrincipalName ($Credential.UserName) #User variable used if logging on user has not mailbox. This confrims that MS Online module is connected
                If ($MSolUser -eq $null){
                        Write-Host "`nYou are not logged into Azure Active Dirctory.`n" -ForegroundColor Red
                }
                Else {
                        Write-Host "`nHello $($MSolUser.DisplayName), you are now logged onto Azure Active Directory.`n" -ForegroundColor Green
                }
 } # End User variable used if logging on user has access to MSOnline. This confrims that MS Online module is connected
 
#endregion Common used scriptblocks

#region All Connect/disconnect functions

function connect-MSAzureAD {
############################################################################################################
#   ArcSigt DevOps Team Costa Rica Tool to Testing MS Office 365 througth PowerShell                       #
#   Made by Andres Duran Campos DevOps CR andres.duran-campos@hpe.com                                      #
#   This Tool is for internal use only to testing MS O365                                                  #
#                                                                                                          #
#                                                                                                          #
#                                                                                                          #
############################################################################################################


###########################################################
# Set Global Variables and Utilities                      #
#                                                         #
###########################################################

$str_path = "C:\Users\duranand\Documents\MS_AZURE_Files\MS_O365_Tool\Log_Files\o365testing.txt"
$str_licenses = "arstconnazureo365:DEVELOPERPACK"
$str_new_user = "devopscrtesting@arstconnazureo365.onmicrosoft.com"
$str_display_name = "DevOpsCR"
$str_usage_location = "US"
$str_role = "User Account Administrator"

# Create a log into path selected
Out-File -FilePath $str_path

##########################################################
# Connect to MS O365 Account, show active user           #
# and licenses                                           #
##########################################################
'START PROCESS OF TESTING MS O365' | Add-Content -Path $str_path

#Connection to Ms O365
Connect-MsolService


# Show list of active user
'SHOW LIST OF USER ACTIVE' | Add-Content -Path $str_path 
$get_user = Get-MsolUser 
out-file -filePath $str_path -Append -InputObject $get_user

# Show the active License 
'SHOW ACTIVE LICENSES' | Add-Content -Path $str_path 
$get_licenses = Get-MsolAccountSku 
out-file -filePath $str_path -Append -InputObject $get_licenses

# Show the active License 
'SHOW ACTIVE LICENSES' | Add-Content -Path $str_path 
$get_licenses_services = Get-MsolAccountSku | Select -ExpandProperty ServiceStatus
out-file -filePath $str_path -Append -InputObject $get_licenses_services

# Create new user
'SHOW NEW USER WAS CREATED INTO THE LIST OF USERS' | Add-Content -Path $str_path 
$create_User = New-MsolUser -UserPrincipalName $str_new_user -DisplayName $str_display_name -UsageLocation $str_usage_location
out-file -filePath $str_path -Append -InputObject $create_user
$get_user = Get-MsolUser 
out-file -filePath $str_path -Append -InputObject $get_user

# Assing License to user 
'SHOW ASSING LICENSE AT NEW USER' | Add-Content -Path $str_path 
Set-MsolUserLicense -UserPrincipalName $str_new_user -AddLicenses $str_licenses
$get_user = Get-MsolUser 
out-file -filePath $str_path -Append -InputObject $get_user

#Show details of service and licenses of new user
'SHOW DETAILS OF SERVICE AND LICENSES OF NEW USER' | Add-Content -Path $str_path  
$get_deteils = Get-MsolUser -UserPrincipalName $str_new_user | Select-Object -ExpandProperty Licenses | Select-Object -ExpandProperty ServiceStatus
out-file -filePath $str_path -Append -InputObject $get_deteils

# Display the list of available roles 
'SHOW List of Role Available' | Add-Content -Path $str_path
$get_role = Get-MsolRole | Sort Name | Select Name,Description
out-file -filePath $str_path -Append -InputObject $get_role


# Assing role at user
'THE ROLE WAS ASSIGNED AT THE NEW USER SUCCESSFULLY' | Add-Content -Path $str_path
$add_role = Add-MsolRoleMember -RoleMemberEmailAddress (Get-MsolUser | Where DisplayName -eq $str_display_name).UserPrincipalName -RoleName $str_role
out-file -filePath $str_path -Append -InputObject $add_role

#Configure user account properties 
'THE FOLOWING IS THE CONFIGURATION OF USER PROPERTIES' | Add-Content -Path $str_path
$upn=(Get-MsolUser | where {$_.DisplayName -eq $str_display_name}).UserPrincipalName
Set-MsolUser -UserPrincipalName $upn -UsageLocation "FR"
$get_location = Get-MsolUser | Select-Object DisplayName, Department, UsageLocation
out-file -filePath $str_path -Append -InputObject $get_location

#Remove License to user
'SHOW REMOVE LICENSE AT NEW USER' | Add-Content -Path $str_path 
Set-MsolUserLicense -UserPrincipalName $str_new_user -RemoveLicenses $str_licenses
$get_user = Get-MsolUser 
out-file -filePath $str_path -Append -InputObject $get_user

#Remove User to active list
'SHOW THE USER WAS REMOVED IS NO LOGER AT LIST USERS' | Add-Content -Path $str_path 
Remove-MsolUser -UserPrincipalName $str_new_user 
$get_user = Get-MsolUser 
out-file -filePath $str_path -Append -InputObject $get_user
}

function Connect-Graph{
############################################################################################################
#   ArcSigt DevOps Team Costa Rica Tool to Testing MS Office 365 througth PowerShell                       #
#   Made by Andres Duran Campos DevOps CR andres.duran-campos@hpe.com                                      #
#   This Tool is for internal use only to testing MS O365                                                  #
#                                                                                                          #
#                                                                                                          #
#                                                                                                          #
############################################################################################################

Import-Module -name 'PSMSGraph'

#In the credential prompt, provide your application's Client ID as the username and Client Secret as the password
$ClientCredential = Get-Credential
$GraphAppParams = @{
    Name = 'DevOpsMSGraphAppPS'
    ClientCredential = $ClientCredential
    RedirectUri = 'https://localhost/'
    Tenant = 'arstconnazureo365.onmicrosoft.com'

}
$GraphApp = New-GraphApplication @GraphAppParams
# This will prompt you to log in with your O365/Azure credentials. 
# This is required at least once to authorize the application to act on behalf of your account
# The username and password is not passed back to or stored by PowerShell.
$AuthCode = $GraphApp | Get-GraphOauthAuthorizationCode 
# see the following help for what resource to use. 
# get-help Get-GraphOauthAccessToken -Parameter Resource
$GraphAccessToken = $AuthCode | Get-GraphOauthAccessToken -Resource 'https://graph.windows.net'
$GraphAccessToken | Export-GraphOAuthAccessToken -Path 'C:\Users\duranand\Documents\MS_AZURE_Files\MS_O365_Tool\AccessToken.XML'


$GraphAccessToken = Import-GraphOAuthAccessToken -Path 'C:\Users\duranand\Documents\MS_AZURE_Files\MS_O365_Tool\AccessToken.XML'


Write-Output $GraphAccessToken
}

function Connect-EXO {
############################################################################################################
#   ArcSigt DevOps Team Costa Rica Tool to Testing MS Office 365 througth PowerShell                       #
#   Made by Andres Duran Campos DevOps CR andres.duran-campos@hpe.com                                      #
#   This Tool is for internal use only to testing MS O365                                                  #
#                                                                                                          #
#                                                                                                          #
#                                                                                                          #
############################################################################################################


###########################################################
# Set Global Variables and Utilities                      #
#                                                         #
###########################################################

$str_path = "C:\Users\duranand\Documents\MS_AZURE_Files\MS_O365_Tool\Log_Files\psmsgraphtoolo365.txt"

# Create a log into path selected
Out-File -FilePath $str_path


##########################################################
# Connect to MS Graph  Web App                           #
#                                                        #
##########################################################
'START PROCESS OF TESTING MS O365' | Add-Content -Path $str_path




# The resource URI
$resource = "https://graph.microsoft.com"
# Your Client ID and Client Secret obainted when registering your WebApp
$clientid = "44d6a59c-0b68-434f-813f-d825da6669f3"
$clientSecret = "VUYmoDfbior2SwTkejbQJAg"

$redirectUri = "https://localhost:8001/"

# UrlEncode the ClientID and ClientSecret and URL's for special characters 
$clientIDEncoded = [System.Web.HttpUtility]::UrlEncode($clientid)
$clientSecretEncoded = [System.Web.HttpUtility]::UrlEncode($clientSecret)
$redirectUriEncoded = [System.Web.HttpUtility]::UrlEncode($redirectUri)
$resourceEncoded = [System.Web.HttpUtility]::UrlEncode($resource)
$scopeEncoded = [System.Web.HttpUtility]::UrlEncode("https://graph.microsoft.com/v1.0/")

# Function to popup Auth Dialog Windows Form
Function Get-AuthCode {
    Add-Type -AssemblyName System.Windows.Forms

    $form = New-Object -TypeName System.Windows.Forms.Form -Property @{Width=440;Height=640}
    $web  = New-Object -TypeName System.Windows.Forms.WebBrowser -Property @{Width=420;Height=600;Url=($url -f ($Scope -join "%20")) }

    $DocComp  = {
        $Global:uri = $web.Url.AbsoluteUri        
        if ($Global:uri -match "error=[^&]*|code=[^&]*") {$form.Close() }
    }
    $web.ScriptErrorsSuppressed = $true
    $web.Add_DocumentCompleted($DocComp)
    $form.Controls.Add($web)
    $form.Add_Shown({$form.Activate()})
    $form.ShowDialog() | Out-Null

    $queryOutput = [System.Web.HttpUtility]::ParseQueryString($web.Url.Query)
    $output = @{}
    foreach($key in $queryOutput.Keys){
        $output["$key"] = $queryOutput[$key]
    }

    $output
}



##########################################################
# Get Authentication Code                                #
#                                                        #
##########################################################

# Get AuthCode
$url = "https://login.microsoftonline.com/common/oauth2/authorize?response_type=code&redirect_uri=$redirectUriEncoded&client_id=$clientID&resource=$resourceEncoded&prompt=admin_consent&scope=$scopeEncoded"
Get-AuthCode
# Extract Access token from the returned URI
$regex = '(?<=code=)(.*)(?=&)'
$authCode  = ($uri | Select-string -pattern $regex).Matches[0].Value

Write-output "Received an authCode, $authCode"

##########################################################
# Get Access Token                                       #
#                                                        #
##########################################################

#get Access Token
$body = "grant_type=authorization_code&redirect_uri=$redirectUri&client_id=$clientId&client_secret=$clientSecretEncoded&code=$authCode&resource=$resource"
$Authorization = Invoke-RestMethod https://login.microsoftonline.com/common/oauth2/token `
    -Method Post -ContentType "application/x-www-form-urlencoded" `
    -Body $body `
    -ErrorAction STOP

##########################################################
# Run Querys on Ms Graph using Access Token              #
#                                                        #
##########################################################

'SHOW AUTHORIZATION TOKEN' | Add-Content -Path $str_path
Write-output $Authorization.access_token
$accesstoken = $Authorization.access_token

out-file -filePath $str_path -Append -InputObject $Authorization

'SHOW MY PROFILE' | Add-Content -Path $str_path
$me = Invoke-RestMethod -Headers @{Authorization = "Bearer $accesstoken"} `
                        -Uri  https://graph.microsoft.com/v1.0/me `
                        -Method Get
out-file -filePath $str_path -Append -InputObject $me


'SHOW MY MESSAGES' | Add-Content -Path $str_path
$messages = $obj = Invoke-RestMethod -Headers @{Authorization = "Bearer $accesstoken"} `
                        -Uri  https://graph.microsoft.com/v1.0/me/messages `
                        -Method Get
$messages | ConvertTo-JSON
out-file -filePath $str_path -Append -InputObject $messages|ConvertTo-Json 
            

'SHOW MY EVENTS' | Add-Content -Path $str_path
$calendars = Invoke-RestMethod -Headers @{Authorization = "Bearer $accesstoken"} `
                        -Uri  https://graph.microsoft.com/v1.0/me/calendars `
                        -Method Get
$calendars|ConvertTo-Json
out-file -filePath $str_path -Append -InputObject $calendars|ConvertTo-Json


 
'SHOW MY FILES' | Add-Content -Path $str_path
$file = Invoke-RestMethod -Headers @{Authorization = "Bearer $accesstoken"} `
                        -Uri  https://graph.microsoft.com/v1.0/me/drive/root/children `
                        -Method Get
$file|ConvertTo-Json
out-file -filePath $str_path -Append -InputObject $file|ConvertTo-Json



'SHOW MY CONTACTS' | Add-Content -Path $str_path
$contacts = Invoke-RestMethod -Headers @{Authorization = "Bearer $accesstoken"} `
                        -Uri  https://graph.microsoft.com/v1.0/me/contacts `
                        -Method Get
$contacts|ConvertTo-Json
out-file -filePath $str_path -Append -InputObject $contacts|ConvertTo-Json

}

function Connect-OneDrive {
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
$AppID="990bfab9-6eb4-4cf3-b925-2abcf226dd04"       # Application GUID
 
$AppKey="QPFCiIw16XG8trhLoPmNbhQUAdrahRA/ifgDiKyGWP0="  # Application Secret Key
 
$RedirectURI="https://sepago.de/1Drive4Business"     # Single-Sign-On URL
 
$ResourceID="https://arstconnazureo365-my.sharepoint.com/" # Resource ID
 
 
 
$form = New-Object Windows.Forms.Form
 
$form.text = "Authenticate to OneDrive for Business"
 
$form.size = New-Object Drawing.size @(700,600)
 
$form.Width = 675
 
$form.Height = 880
 
$web = New-object System.Windows.Forms.WebBrowser
 
$web.IsWebBrowserContextMenuEnabled = $true
 
$web.Width = 600
 
$web.Height = 760
 
$web.Location = "25, 25"
 
$web.navigate("https://login.microsoftonline.com/common/oauth2/authorize?response_type=code&client_id="+$AppID+"&redirect_uri="+$RedirectURI)
 
$DocComplete  = {
 
    $Global:uri = $web.Url.AbsoluteUri
 
    if ($web.Url.AbsoluteUri -match "code=|error") {$form.Close() }
 
}
 
$web.Add_DocumentCompleted($DocComplete)
 
$form.Controls.Add($web)
 
$form.showdialog() | out-null
 
# Build object from last URI (which should contains the token)
 
$ReturnURI=($web.Url).ToString().Replace("#","&")
 
$Code = New-Object PSObject
 
ForEach ($element in $ReturnURI.Split("?")[1].Split("&")) 
 
{
 
    $Code | add-member Noteproperty $element.split("=")[0] $element.split("=")[1]
 
}
 
 
 
if($Code.Code) 
 
{
 
    #Code exist
 
    write-host("Access code:")
 
    write-host($Code.code)
 
 
 
    #Get authentication token
 
    $body="client_id="+$AppID+"&redirect_uri="+$RedirectURI+"&client_secret="+$AppKey+"&code="+$Code.code+"&grant_type=authorization_code&resource="+$ResourceID
 
    $Response = Invoke-WebRequest -Method Post -ContentType "application/x-www-form-urlencoded" "https://login.microsoftonline.com/common/oauth2/token" -body $body
 
    $Authentication = $Response.Content|ConvertFrom-Json
 
 
 
    if($Authentication.access_token)
 
    {
 
        #Authentication token exist
 
        write-host("`nAuthentication:")
 
        $Authentication
 
 
 
        #Samples:
 
        write-host("Drive info:")
 
        $Response=Invoke-WebRequest -Method GET -Uri ($ResourceID+"_api/v2.0"+"/drive") -Header @{ Authorization = "BEARER "+$Authentication.access_token} -ErrorAction Stop
 
        $responseObject = ConvertFrom-Json $Response.Content
 
        $responseObject
 
 
 
        write-host("Files and folders in root drive:")
 
        $Response=Invoke-WebRequest -Method GET -Uri ($ResourceID+"_api/v2.0"+"/drive/root/children:") -Header @{ Authorization = "BEARER "+$Authentication.access_token} -ErrorAction Stop
 
        $responseObject = ($Response.Content|ConvertFrom-Json).value
 
        $responseObject
 
 
 
 
 
 
 
    } else
 
    {    
 
        #No Code
 
        write-host("No authentication token recieved to log in")
 
    }
 
} else
 
 
 
{
 
    #No Code
 
    write-host("No code received to log in")
 
}

}


function Connect-SherePointPnP{
#######Script to test performance of SharePoint Online in the RDC environment#######

#######Load SDK#######
$path1 = "$pwd\Support_Files\Microsoft.SharePoint.Client.Runtime.dll"
$path2 = "$pwd\Support_Files\Microsoft.SharePoint.Client.dll"
$path3 = "$pwd\Support_Files\Microsoft.Online.SharePoint.Client.Tenant.dll"
$path4 = "$pwd\Support_Files\Microsoft.SharePoint.Client.Taxonomy.dll"

Add-Type -Path $path1
Add-Type -Path $path2
Add-Type -Path $path3
Add-Type -Path $path4

#######Import PNP Mpdule#######
Import-Module $pwd\Support_Files\SharePointPnPPowerShellOnline

#######Log Function#######
$fileFormat = Get-Date -format "dd-MMM-yy_HHmmss"

#######Logging function#######
$now=Get-Date -format "dd-MMM-yy,HH:mm:ss"
$logtime = Get-Date -Format HH:mm:ss
Function LogWrite
{
   Param ([string]$logstring)
   Add-content $Logfile -value $logstring
}

#######HTTP Download function#######
function downloadFile($url, $targetFile)
{ 
    "Downloading $url" 
    $uri = New-Object "System.Uri" "$url" 
    $request = [System.Net.HttpWebRequest]::Create($uri) 
    $request.set_Timeout(15000) #15 second timeout 
    $response = $request.GetResponse() 
    $totalLength = [System.Math]::Floor($response.get_ContentLength()/1024) 
    $responseStream = $response.GetResponseStream() 
    $targetStream = New-Object -TypeName System.IO.FileStream -ArgumentList $targetFile, Create 
    $buffer = new-object byte[] 10KB 
    $count = $responseStream.Read($buffer,0,$buffer.length) 
    $downloadedBytes = $count 
    while ($count -gt 0) 
    { 
        [System.Console]::CursorLeft = 0 
        [System.Console]::Write("Downloaded {0}K of {1}K", [System.Math]::Floor($downloadedBytes/1024), $totalLength) 
        $targetStream.Write($buffer, 0, $count) 
        $count = $responseStream.Read($buffer,0,$buffer.length) 
        $downloadedBytes = $downloadedBytes + $count 
    } 
    "`nFinished Download" 
    $targetStream.Flush()
    $targetStream.Close() 
    $targetStream.Dispose() 
    $responseStream.Dispose() 
}
 

#######Credentials and details#######
Write-Host "Please ensure you have owner permissions on the destination site" -ForegroundColor Green
$SiteURL = Read-Host -Prompt "Enter Site URL without final /"
connect-pnponline -url $SiteURL -useweblogin
$TestSiteURL = "https://arstconnazureo365.sharepoint.com"
$Logfile = $pwd.path + "/Log_Files/Performance_" + $fileFormat + ".csv"

write-progress -id 1 -activity "DevOps Team SharePoint Online Performance Testing" -status "Performing web request test" -percentComplete 5
LogWrite "Time,Action"
$logtime = Get-Date -Format HH:mm:ss
LogWrite "$logtime,$SiteURL"
LogWrite "$logtime,Testing SharePoint contact speed"
foreach($i in 1..4){

 $timeTaken = Measure-Command -Expression {
 $site = Invoke-WebRequest -Uri $TestSiteURL -UseBasicParsing
 }
 $milliseconds = $timeTaken.TotalMilliseconds
 $milliseconds = [Math]::Round($milliseconds, 0)
 $totalmills = $totalmills + $milliseconds
 $logtime = Get-Date -Format HH:mm:ss
 LogWrite "$logtime,Contact $i $milliseconds ms"
}


$Average = $totalmills / 4
Write-Host "Average contact speed out of 4 attemts: $Average ms" -ForegroundColor Green
$logtime = Get-Date -Format HH:mm:ss
LogWrite "$logtime,Average contact speed of $Average ms"


write-progress -id 1 -activity "DevOps Team SharePoint Online Performance Testing" -status "Authenticating with SharePoint Online" -percentComplete 10

$logtime = Get-Date -Format HH:mm:ss
LogWrite "$logtime,Authenticating"


#######Download Test Files#######
write-progress -id 1 -activity "DevOps Team SharePoint Online Performance Testing" -status "Downloading sample files from web" -percentComplete 15

$logtime = Get-Date -Format HH:mm:ss
LogWrite "$logtime,Downloading Sample Files"
$logtime = Get-Date -Format HH:mm:ss
LogWrite "$logtime,Downloading 100kb files"
foreach($i in 1..4){
$TestFileName = $pwd.path + "/Sample_Files/Word_Document_100_" + $i + ".doc"
downloadFile "http://www.sample-videos.com/doc/Sample-doc-file-100kb.doc" $TestFileName
}
$logtime = Get-Date -Format HH:mm:ss
LogWrite "$logtime,Completed 100kb files" 
$logtime = Get-Date -Format HH:mm:ss
LogWrite "$logtime,Downloading 5000kb files"
foreach($i in 1..4){
$TestFileName = $pwd.path + "/Sample_Files/Word_Document_5000_" + $i + ".doc"
downloadFile "http://www.sample-videos.com/doc/Sample-doc-file-5000kb.doc" $TestFileName
}
$logtime = Get-Date -Format HH:mm:ss
LogWrite "$logtime,Completed 5000kb files"
$logtime = Get-Date -Format HH:mm:ss
LogWrite "$logtime,Downloading 2000kb files"
foreach($i in 1..4){
$TestFileName = $pwd.path + "/Sample_Files/Word_Document_2000_" + $i + ".doc"
downloadFile "http://www.sample-videos.com/doc/Sample-doc-file-2000kb.doc" $TestFileName
}
$logtime = Get-Date -Format HH:mm:ss
LogWrite "$logtime,Completed 2000kb files"
$logtime = Get-Date -Format HH:mm:ss
LogWrite "$logtime,Downloading 200MB files"
$TestFileName = $pwd.path + "/Sample_Files/Large_Zip_" + $i + ".zip"
downloadFile "http://ipv4.download.thinkbroadband.com/200MB.zip" $TestFileName
$logtime = Get-Date -Format HH:mm:ss
LogWrite "$logtime,Completed 200MB files"

#######Create Document Library#######
Clear-Host
write-progress -id 1 -activity "DevOps Team SharePoint Online Performance Testing" -status "Creating test SharePoint document library" -percentComplete 20

$Context = Get-PnPContext
$ListInfo = New-Object Microsoft.SharePoint.Client.ListCreationInformation
$ListInfo.Title = "Performance_Library"
$ListInfo.TemplateType = 101 #Document Library
$List = $Context.Web.Lists.Add($ListInfo)
$List.Description = "Library to test SharePoint performance"
$List.Update()
$Context.ExecuteQuery()
$logtime = Get-Date -Format HH:mm:ss
LogWrite "$logtime,Document library created"
write-progress -id 1 -activity "DevOps Team SharePoint Online Performance Testing" -status "Test SharePoint document library created" -percentComplete 25


#######Upload SharePoint Files#######

write-progress -id 1 -activity "DevOps Team SharePoint Online Performance Testing" -status "Uploading sample files to SharePoint library" -percentComplete 35

$DocLibName = "Performance_Library"
$TestFileFolder = $pwd.path + "/Sample_Files"
$Context = Get-PnPContext
$List = $Context.Web.Lists.GetByTitle($DocLibName)
$Context.Load($List)
$Context.ExecuteQuery()

$logtime = Get-Date -Format HH:mm:ss
LogWrite "$logtime,Start of upload"

Foreach ($File in (dir $TestFileFolder -File))
{
$FileStream = New-Object IO.FileStream($File.FullName,[System.IO.FileMode]::Open)
$FileCreationInfo = New-Object Microsoft.SharePoint.Client.FileCreationInformation
$FileCreationInfo.Overwrite = $true
$FileCreationInfo.ContentStream = $FileStream
$FileCreationInfo.URL = $File
$Upload = $List.RootFolder.Files.Add($FileCreationInfo)
$Context.Load($Upload)
$Context.Load($List.ContentTypes)
$Context.ExecuteQuery()
$item = $Upload.ListItemAllFields
$logtime = Get-Date -Format HH:mm:ss
LogWrite "$logtime,$File.BaseName uploaded"
$item["Title"] = $File.BaseName
$item.Update()
$Context.ExecuteQuery()
}

$logtime = Get-Date -Format HH:mm:ss
LogWrite "$logtime,End of upload"
write-progress -id 1 -activity "DevOps Team SharePoint Online Performance Testing" -status "Performing web request test" -percentComplete 55
LogWrite "$logtime,Testing SharePoint contact speed to new library"
foreach($i in 1..4){

 $timeTaken = Measure-Command -Expression {
 $site = Invoke-WebRequest -Uri $SiteURL -UseBasicParsing
 }
 $milliseconds = $timeTaken.TotalMilliseconds
 $milliseconds = [Math]::Round($milliseconds, 0)
 $totalmills = $totalmills + $milliseconds
 $logtime = Get-Date -Format HH:mm:ss
 LogWrite "$logtime,Contact $i $milliseconds ms"
}

$Average = $totalmills / 4
Write-Host "Average contact speed out of 4 attemts: $Average ms"
$logtime = Get-Date -Format HH:mm:ss
LogWrite "$logtime,Average contact speed of $Average ms"

write-progress -id 1 -activity "DevOps Team SharePoint Online Performance Testing" -status "Deleting local sample files" -percentComplete 65
#######Delete local files#######
Get-ChildItem -Path "$pwd/Sample_Files" -Include *.* -File -Recurse | foreach { $_.Delete()}
$logtime = Get-Date -Format HH:mm:ss
LogWrite "$logtime,local files deleted"

#######Download SharePoint Files#######
write-progress -id 1 -activity "DevOps Team SharePoint Online Performance Testing" -status "Downloading sample files from SharePoint Document Library" -percentComplete 70
$logtime = Get-Date -Format HH:mm:ss
LogWrite "$logtime,Starting Download from SPO"

$Context = Get-PnPContext
$destinationLocalFolder = $pwd.path + "/Sample_Files/"
$DocLibName = "Performance_Library"

$DocLib = $Context.Web.Lists.GetByTitle($DocLibName)
$Context.Load($DocLib)
$Context.ExecuteQuery()

$ListRoot = $DocLib.RootFolder
$FilesInFolder = $ListRoot.Files
$Context.Load($FilesInFolder)
$Context.ExecuteQuery()

foreach ($file in $FilesInFolder)
{
$logtime = Get-Date -Format HH:mm:ss
LogWrite "$logtime,$file.Name has been downloaded"
$FileRef = $file.ServerRelativeUrl
$FileInfo = [Microsoft.SharePoint.Client.File]::OpenBinaryDirect($Context,$FileRef)
$FileName = $destinationLocalFolder + $file.Name
$filestream = [System.IO.File]::Create($FileName)
$FileInfo.Stream.CopyTo($filestream)
LogWrite "$logtime,$Filename has been downloaded"
$filestream.Close()
}

$logtime = Get-Date -Format HH:mm:ss
LogWrite "$logtime,Download completed"
write-progress -id 1 -activity "DevOps Team SharePoint Online Performance Testing" -status "Download Complete" -percentComplete 85
#######Delete local files#######
write-progress -id 1 -activity "DevOps Team SharePoint Online Performance Testing" -status "Deleting local sample files" -percentComplete 90
Get-ChildItem -Path "$pwd/Sample_Files" -Include *.* -File -Recurse | foreach { $_.Delete()}
$logtime = Get-Date -Format HH:mm:ss
LogWrite "$logtime,local files deleted"


#######Delete Document Library#######
write-progress -id 1 -activity "DevOps Team SharePoint Online Performance Testing" -status "Deleting test SharePoint document library" -percentComplete 95
$Context = Get-PnPContext
$list = $Context.Web.Lists.GetByTitle($DocLibName)
$Context.Load($list)
$list.DeleteObject()
$Context.ExecuteQuery()
$logtime = Get-Date -Format HH:mm:ss
LogWrite "$logtime,Document library deleted"
write-progress -id 1 -activity "DevOps Team SharePoint Online Performance Testing" -status "Deleted test SharePoint document library" -percentComplete 98

disconnect-pnponline
write-progress -id 1 -activity "DevOps Team SharePoint Online Performance Testing" -status "Test complete, Thanks for your cooperation" -percentComplete 100

#Add-Type -Assembly "Microsoft.Office.Interop.Outlook"
#$ol = New-Object -com Outlook.Application 
#$message = $ol.CreateItem(0)
#$message.Recipients.Add("andres@arstconnazureo365.onmicrosoft.com")
#$message.Recipients.Add("andres@arstconnazureo365.onmicrosoft.com")  
#$message.Subject = "SharePoint Online Performance Testing $SiteURL"  
#$message.Body = "See attached file for the performance testing of $SiteURL"
#$message.Attachments.Add($Logfile)
#$message.Send()



Read-Host -Prompt "Press Enter to exit"
}


function Run-PythonFile {
Start-Process -FilePath http://localhost:5000/
$python = python manage.py runserver

}





#endregion All Connect/disconnect functions

#region Menu action

Do { 	
	if ($opt -ne "None") {Write-Host "Last command: "$opt -foregroundcolor Yellow}	
	$opt = Read-Host $menu

	switch ($opt)    {
    			
	  	1 { # Log onto MS Azure AD
            connect-MSAzureAD
            
        }

        2 { # Send Email with Exchange Online
            Run-PythonFile
            
        }

        3 { # Log onto Exchange Online
            Connect-EXO
            
        }

        4 { # Log onto OneDrive Online and check The Site
            Connect-OneDrive
            
        }
		
		5 { # Log onto SharePoint/OneDrive Online and Upload/Download Files
            Connect-SherePointPnP
		}

	  	10 { # Windows Update
			Invoke-Expression "$env:windir\system32\wuapp.exe startmenu"
		}
		
		11 { # Install - .NET 4.5.2
			 Install-DotNET452
		}

        12 { # Install MS Online Service Sign-In Assistance for IT Professionals RTW
            Install-SignInAssistant
        }

        13 { # Install - Windows Azure Active Directory Module for Windows PowerShell (64-bit)
            Install-WindowsAADModule
        }

        14 { # Install - SharePoint Online Module
            Install-SPO
        }

        15 { # Install - Skype for Business Online Module
            Install-SfbO
        }
         
        20 { # Install - O365_Logon Module 1.0
            Install-O365_Logon
        }

        30 { # Enable PS Remoting. This fixes WinRM error
            Enable-PSRemoting -Force -SkipNetworkProfileCheck
        }

        31 { # Launches PS 3.0 install website
            Start-Process "https://www.microsoft.com/en-us/download/details.aspx?id=34595"
        }

        32 { # Launches PS 4.0 install website
            Start-Process "https://www.microsoft.com/en-us/download/details.aspx?id=40855"
        }

        33 { # Launches PS 5.0 install website
            Start-Process "https://www.microsoft.com/en-us/download/details.aspx?id=48729"
        }

        90 { # Launch SharePoint Online Management Shell install website
            Start-Process "https://www.microsoft.com/en-us/download/details.aspx?id=35588" 
        }
       
        91 { # Launch OneDrive PowerShell Gallery Module How to Install
            Start-Process "https://www.powershellgallery.com/packages/OneDrive/1.0.3"
        }

		98 { # Exit and restart
			Restart-Computer -computername localhost -force
		}

		99 { # Exit
			if (($WasInstalled -eq $false) -and (Get-Module BitsTransfer)){
				Write-Host "BitsTransfer: Removing..." -NoNewLine
				Remove-Module BitsTransfer
				Write-Host "Removed." -ForegroundColor Green
			}
			Write-Host "Exiting..."
		}
		
        default {Write-Host "You haven't selected any of the available options."}
	}
} while ($opt -ne 99)

#endregion Menu action