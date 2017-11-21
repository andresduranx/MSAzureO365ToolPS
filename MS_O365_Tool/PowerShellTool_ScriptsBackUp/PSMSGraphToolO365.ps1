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

$str_path = "C:\Users\duranand\Documents\MS_AZURE_Files\MS_O365_Tool\psmsgraphtoolo365.txt"

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



