############################################################################################################
#   ArcSigt DevOps Team Costa Rica Tool to Testing MS Office 365 througth PowerShell                       #
#   Made by Andres Duran Campos DevOps CR andres.duran-campos@microfocus.com                                      #
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
	    DevOps Team ArcSight Logon/Test O365 Services Tool
	*******************************************************************
	
	Please select an option from the list below:


     1) Log onto MS Azure Active Directory
     2) Send Email with Exchange Online 
     3) Log onto Exchange Online
     4) Log onto OneDrive Online and check The Site
     5) Log onto SharePoint/OneDrive Online and Upload/Download Files
	 
     10) Exit this tool

Select an option.. [1-10]?
'@

#endregion Menu display



#region All Connect/disconnect functions

function connect-MSAzureAD {

$path = "$pwd/Log_Files/MSAzureADLog"

#######Log Function#######
$fileFormat = Get-Date -format "dd-MMM-yy_HHmmss"

$Logfile = $path + $fileFormat + ".csv"
#######Logging function#######
$now=Get-Date -format "dd-MMM-yy,HH:mm:ss"
$logtime = Get-Date -Format HH:mm:ss
Function LogWrite
{
   Param ([string]$logstring)
   Add-content $Logfile -value $logstring
}

$logtime = Get-Date -Format HH:mm:ss
LogWrite "$logtime, Start Process on Active Directory"
LogWrite "Time,Action"

$str_licenses = "arstconnazureo365:DEVELOPERPACK"
$str_new_user = "devopscrtesting@arstconnazureo365.onmicrosoft.com"
$str_display_name = "DevOpsCR"
$str_usage_location = "US"
$str_role = "User Account Administrator"


#Connection to Ms O365
Connect-MsolService
$logtime = Get-Date -Format HH:mm:ss
LogWrite "$logtime,Testing Connection to Ms O365"

# Show list of active user
write-host -ForegroundColor Green ("SHOW LIST OF USER ACTIVE:")
$logtime = Get-Date -Format HH:mm:ss
LogWrite "$logtime,SHOW LIST OF USER ACTIVE"
$get_user = Get-MsolUser 
$get_user | ConvertTo-JSON
LogWrite $get_user

# Show the active License 
write-host -ForegroundColor Green ("SHOW ACTIVE LICENSES:")
$logtime = Get-Date -Format HH:mm:ss
LogWrite "$logtime,SHOW ACTIVE LICENSES" 
$get_licenses = Get-MsolAccountSku 
$get_licenses | ConvertTo-JSON
LogWrite $get_licenses

# Show the active License 
write-host -ForegroundColor Green ("SHOW ACTIVE LICENSES AND DETAILS:")
$logtime = Get-Date -Format HH:mm:ss
LogWrite "$logtime,SHOW ACTIVE LICENSES AND DETAILS"
$get_licenses_services = Get-MsolAccountSku | Select -ExpandProperty ServiceStatus
$get_licenses_services | ConvertTo-JSON
LogWrite $get_licenses_services

# Create new user
write-host -ForegroundColor Green ("SHOW NEW USER WAS CREATED INTO THE LIST OF USERS:")
$logtime = Get-Date -Format HH:mm:ss
LogWrite "$logtime,SHOW NEW USER WAS CREATED INTO THE LIST OF USERS" 
$create_User = New-MsolUser -UserPrincipalName $str_new_user -DisplayName $str_display_name -UsageLocation $str_usage_location
$get_user = Get-MsolUser 
$get_user | ConvertTo-JSON
LogWrite $get_user

# Assing License to user 
write-host -ForegroundColor Green ("SHOW ASSING LICENSE AT NEW USER:")
$logtime = Get-Date -Format HH:mm:ss
LogWrite "$logtime,SHOW ASSING LICENSE AT NEW USER"
Set-MsolUserLicense -UserPrincipalName $str_new_user -AddLicenses $str_licenses
$get_user = Get-MsolUser
$get_user | ConvertTo-JSON
LogWrite $get_user

#Show details of service and licenses of new user
write-host -ForegroundColor Green ("SHOW DETAILS OF SERVICE AND LICENSES OF NEW USER:")
$logtime = Get-Date -Format HH:mm:ss
LogWrite "$logtime,SHOW DETAILS OF SERVICE AND LICENSES OF NEW USER" 
$get_deteils = Get-MsolUser -UserPrincipalName $str_new_user | Select-Object -ExpandProperty Licenses | Select-Object -ExpandProperty ServiceStatus
$get_deteils | ConvertTo-JSON
LogWrite $get_deteils

# Display the list of available roles 
write-host -ForegroundColor Green ("SHOW List of Role Available:")
$logtime = Get-Date -Format HH:mm:ss
LogWrite "$logtime,SHOW List of Role Available"
$get_role = Get-MsolRole | Sort Name | Select Name,Description
$get_role | ConvertTo-JSON
LogWrite $get_role


# Assing role at user
write-host -ForegroundColor Green ("THE ROLE WAS ASSIGNED AT THE NEW USER SUCCESSFULLY:")
$logtime = Get-Date -Format HH:mm:ss
LogWrite "$logtime,THE ROLE WAS ASSIGNED AT THE NEW USER SUCCESSFULLY"
$add_role = Add-MsolRoleMember -RoleMemberEmailAddress (Get-MsolUser | Where DisplayName -eq $str_display_name).UserPrincipalName -RoleName $str_role
$add_role | ConvertTo-JSON
LogWrite $add_role

#Configure user account properties 
write-host -ForegroundColor Green ("THE FOLOWING IS THE CONFIGURATION OF USER PROPERTIES:")
$logtime = Get-Date -Format HH:mm:ss
LogWrite "$logtime,THE FOLOWING IS THE CONFIGURATION OF USER PROPERTIES"
$upn=(Get-MsolUser | where {$_.DisplayName -eq $str_display_name}).UserPrincipalName
Set-MsolUser -UserPrincipalName $upn -UsageLocation "FR"
$get_location = Get-MsolUser | Select-Object DisplayName, Department, UsageLocation
$get_location | ConvertTo-JSON
LogWrite $get_location

#Remove License to user
write-host -ForegroundColor Green ("SHOW REMOVE LICENSE AT NEW USER:")
$logtime = Get-Date -Format HH:mm:ss
LogWrite "$logtime,SHOW REMOVE LICENSE AT NEW USER"
Set-MsolUserLicense -UserPrincipalName $str_new_user -RemoveLicenses $str_licenses
$get_user = Get-MsolUser 
$get_user | ConvertTo-JSON
LogWrite $get_user

#Remove User to active list
write-host -ForegroundColor Green ("SHOW THE USER WAS REMOVED IS NO LOGER AT LIST USERS:")
$logtime = Get-Date -Format HH:mm:ss
LogWrite "$logtime,SHOW THE USER WAS REMOVED IS NO LOGER AT LIST USERS" 
Remove-MsolUser -UserPrincipalName $str_new_user 
$get_user = Get-MsolUser
$get_user | ConvertTo-JSON
LogWrite $get_user


$logtime = Get-Date -Format HH:mm:ss
LogWrite "$logtime, End Process on Active Directory"
}


function Connect-EXO {

$path = "$pwd/Log_Files/ExchangeLog"

#######Log Function#######
$fileFormat = Get-Date -format "dd-MMM-yy_HHmmss"

$Logfile = $path + $fileFormat + ".csv"
#######Logging function#######
$now=Get-Date -format "dd-MMM-yy,HH:mm:ss"
$logtime = Get-Date -Format HH:mm:ss
Function LogWrite
{
   Param ([string]$logstring)
   Add-content $Logfile -value $logstring
}

$logtime = Get-Date -Format HH:mm:ss
LogWrite "$logtime, Start Process on Exchange"
LogWrite "Time,Action"


$logtime = Get-Date -Format HH:mm:ss
LogWrite "$logtime,Testing Connect to MS Graph  Web App"



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


# Get AuthCode
$url = "https://login.microsoftonline.com/common/oauth2/authorize?response_type=code&redirect_uri=$redirectUriEncoded&client_id=$clientID&resource=$resourceEncoded&prompt=admin_consent&scope=$scopeEncoded"
Get-AuthCode
# Extract Access token from the returned URI
$regex = '(?<=code=)(.*)(?=&)'
$authCode  = ($uri | Select-string -pattern $regex).Matches[0].Value

write-host -ForegroundColor Green ("Received an authCode:")
Write-output "Received an authCode, $authCode"
$logtime = Get-Date -Format HH:mm:ss
LogWrite "$logtime, Get Authentication Code"


#get Access Token
$body = "grant_type=authorization_code&redirect_uri=$redirectUri&client_id=$clientId&client_secret=$clientSecretEncoded&code=$authCode&resource=$resource"
$Authorization = Invoke-RestMethod https://login.microsoftonline.com/common/oauth2/token `
    -Method Post -ContentType "application/x-www-form-urlencoded" `
    -Body $body `
    -ErrorAction STOP
$logtime = Get-Date -Format HH:mm:ss
LogWrite "$logtime, Get Access Token"


#Run Querys on Ms Graph using Access Token
$logtime = Get-Date -Format HH:mm:ss
LogWrite "$logtime, Run Querys on Ms Graph using Access Token"
write-host -ForegroundColor Green ("Authorization Token:")
Write-output $Authorization.access_token
$accesstoken = $Authorization.access_token

$logtime = Get-Date -Format HH:mm:ss
LogWrite "$logtime, Authorization Token"
LogWrite $accesstoken

write-host -ForegroundColor Green ("My Profile:")
$logtime = Get-Date -Format HH:mm:ss
LogWrite "$logtime, SHOW MY PROFILE"
$me = Invoke-RestMethod -Headers @{Authorization = "Bearer $accesstoken"} `
                        -Uri  https://graph.microsoft.com/v1.0/me `
                        -Method Get
LogWrite $me
$me


write-host -ForegroundColor Green ("My Messages:")
$logtime = Get-Date -Format HH:mm:ss
LogWrite "$logtime, SHOW MY MESSAGES"
$messages = $obj = Invoke-RestMethod -Headers @{Authorization = "Bearer $accesstoken"} `
                        -Uri  https://graph.microsoft.com/v1.0/me/messages `
                        -Method Get
$messages | ConvertTo-JSON

LogWrite $messages|ConvertTo-Json 
   


write-host -ForegroundColor Green ("My Events:")   
$logtime = Get-Date -Format HH:mm:ss
LogWrite "$logtime, SHOW MY EVENTS"
$calendars = Invoke-RestMethod -Headers @{Authorization = "Bearer $accesstoken"} `
                        -Uri  https://graph.microsoft.com/v1.0/me/calendars `
                        -Method Get
$calendars|ConvertTo-Json

LogWrite $calendars|ConvertTo-Json


write-host -ForegroundColor Green ("My Files:")
$logtime = Get-Date -Format HH:mm:ss
LogWrite "$logtime, SHOW MY FILES"
$file = Invoke-RestMethod -Headers @{Authorization = "Bearer $accesstoken"} `
                        -Uri  https://graph.microsoft.com/v1.0/me/drive/root/children `
                        -Method Get
$file|ConvertTo-Json

LogWrite $file|ConvertTo-Json


write-host -ForegroundColor Green ("My Contacts:")
$logtime = Get-Date -Format HH:mm:ss
LogWrite "$logtime, SHOW MY CONTACTS"
$contacts = Invoke-RestMethod -Headers @{Authorization = "Bearer $accesstoken"} `
                        -Uri  https://graph.microsoft.com/v1.0/me/contacts `
                        -Method Get
$contacts|ConvertTo-Json

LogWrite $contacts|ConvertTo-Json


$logtime = Get-Date -Format HH:mm:ss
LogWrite "$logtime, End Process on Exchange"
}

function Connect-OneDrive {
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")


$path = "$pwd/Log_Files/OneDriveLog"


#######Log Function#######
$fileFormat = Get-Date -format "dd-MMM-yy_HHmmss"

$Logfile = $path + $fileFormat + ".csv"


#######Logging function#######
$now=Get-Date -format "dd-MMM-yy,HH:mm:ss"
$logtime = Get-Date -Format HH:mm:ss
Function LogWrite
{
   Param ([string]$logstring)
   Add-content $Logfile -value $logstring
}


$logtime = Get-Date -Format HH:mm:ss
LogWrite "$logtime, Start Process on OneDrive"
LogWrite "Time,Action"

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
 
$logtime = Get-Date -Format HH:mm:ss
LogWrite "$logtime,Testing OneDrive Single-Sign-On"
 
if($Code.Code) 
 
{
 
    #Code exist
 
    write-host -ForegroundColor Green ("Access code:")
 
    write-host ($Code.code)
	
	$logtime = Get-Date -Format HH:mm:ss
	LogWrite "$logtime,Code Exist"
 
 
 
    #Get authentication token
 
    $body="client_id="+$AppID+"&redirect_uri="+$RedirectURI+"&client_secret="+$AppKey+"&code="+$Code.code+"&grant_type=authorization_code&resource="+$ResourceID
 
    $Response = Invoke-WebRequest -Method Post -ContentType "application/x-www-form-urlencoded" "https://login.microsoftonline.com/common/oauth2/token" -body $body
 
    $Authentication = $Response.Content|ConvertFrom-Json
	
	$logtime = Get-Date -Format HH:mm:ss
	LogWrite "$logtime, Get Authentication Token"
 
 
 
    if($Authentication.access_token)
 
    {
 
        #Authentication token exist
 
        write-host -ForegroundColor Green ("`nAuthentication:")
 
        $Authentication
		
		$logtime = Get-Date -Format HH:mm:ss
		LogWrite "$logtime, Authentication Token Exist"
		LogWrite "$Authentication"
 
 
 
        #Samples:
 
        write-host -ForegroundColor Green ("Drive info:")
 
        $Response=Invoke-WebRequest -Method GET -Uri ($ResourceID+"_api/v2.0"+"/drive") -Header @{ Authorization = "BEARER "+$Authentication.access_token} -ErrorAction Stop
 
        $responseObject = ConvertFrom-Json $Response.Content
 
        $responseObject
		
		$logtime = Get-Date -Format HH:mm:ss
		LogWrite "$logtime, Drive Information"
		LogWrite "$responseObject"
 
 
 
        write-host -ForegroundColor Green ("Files and folders in root drive:")
 
        $Response=Invoke-WebRequest -Method GET -Uri ($ResourceID+"_api/v2.0"+"/drive/root/children:") -Header @{ Authorization = "BEARER "+$Authentication.access_token} -ErrorAction Stop
 
        $responseObject = ($Response.Content|ConvertFrom-Json).value
 
        $responseObject
		
		$logtime = Get-Date -Format HH:mm:ss
		LogWrite "$logtime, Files and Folders in root drive"
		LogWrite "$responseObject"
 
 
 
 
 
 
 
    } else
 
    {    
 
        #No Code
 
        write-host("No authentication token recieved to log in")
		
		$logtime = Get-Date -Format HH:mm:ss
		LogWrite "$logtime, No authentication token recieved to log in"
 
    }
 
} else
 
 
 
{
 
    #No Code
 
    write-host("No code received to log in")
	
	$logtime = Get-Date -Format HH:mm:ss
	LogWrite "$logtime, No code received to log in"
 
}

$logtime = Get-Date -Format HH:mm:ss
LogWrite "$logtime, End Process"

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

		10 { # Exit
			if (($WasInstalled -eq $false) -and (Get-Module BitsTransfer)){
				Write-Host "BitsTransfer: Removing..." -NoNewLine
				Remove-Module BitsTransfer
				Write-Host "Removed." -ForegroundColor Green
			}
			Write-Host "Exiting..."
		}
		
        default {Write-Host "You haven't selected any of the available options."}
	}
} while ($opt -ne 10)

#endregion Menu action