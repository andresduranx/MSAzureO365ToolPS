Import-Module -name 'PSMSGraph'
$GraphAccessToken =  Import-GraphOAuthAccessToken -Path 'C:\Users\duranand\Documents\MS_AZURE_Files\MS_O365_Tool\AccessToken.XML'
$GraphAccessToken | Update-GraphOAuthAccessToken -Force

$AADUsers = Get-AADUserAll -AccessToken $GraphAccessToken
$AADUsers | 
    Select-Object -Property * -ExcludeProperty _AccessToken | 
    Export-Csv -Path 'C:\Users\duranand\Documents\MS_AZURE_Files\MS_O365_Tool\AADUsers.csv' -NoTypeInformation

$GraphAccessToken  | Export-GraphOAuthAccessToken -Path 'C:\Users\duranand\Documents\MS_AZURE_Files\MS_O365_Tool\AccessToken.XML'


