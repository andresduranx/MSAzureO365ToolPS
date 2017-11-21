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

$str_path = "C:\Users\duranand\Documents\MS_AZURE_Files\MS_O365_Tool\sharepointtesting.txt"

$orgName="arstconnazureo365"

# Create a log into path selected
Out-File -FilePath $str_path

##########################################################
# Connect to MS O365 Account                             #
#                                                        #
##########################################################
'START PROCESS OF TESTING MS O365' | Add-Content -Path $str_path


Connect-SPOService -Url https://$orgName-admin.sharepoint.com

'SHOW LIST OF SITES' | Add-Content -Path $str_path 
$get_sites = Get-SPOSite
out-file -filePath $str_path -Append -InputObject $get_sites

'SHOW LIST OF GROUP' | Add-Content -Path $str_path 
$get_group = Get-SPOSite | ForEach-Object {Get-SPOSiteGroup -Site $_.Url} |Format-Table
out-file -filePath $str_path -Append -InputObject $get_group

'SHOW LIST OF USERS' | Add-Content -Path $str_path 
$get_user = Get-SPOSite | ForEach-Object {Get-SPOUser -Site $_.Url}
out-file -filePath $str_path -Append -InputObject $get_user

'SHOW LIST OF ALL GROUP AND USER OF THEM' | Add-Content -Path $str_path
$siteURL = Get-SPOSite

foreach ($y in $x)
    {
        Write-Host $y.Url -ForegroundColor "Yellow"
        $z = Get-SPOSiteGroup -Site $y.Url
        foreach ($a in $z)
            {
                 $b = Get-SPOSiteGroup -Site $y.Url -Group $a.Title 
                 Write-Host $b.Title -ForegroundColor "Cyan"
                 $b | Select-Object -ExpandProperty Users
                 Write-Host
            }
    }
out-file -filePath $str_path -Append -InputObject $siteURL





Disconect-SPOService

