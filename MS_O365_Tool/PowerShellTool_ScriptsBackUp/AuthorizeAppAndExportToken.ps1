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

