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

$str_path = "C:\Users\duranand\Documents\MS_AZURE_Files\MS_O365_Tool\o365testing.txt"
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

