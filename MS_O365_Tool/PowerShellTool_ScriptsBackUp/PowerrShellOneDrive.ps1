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
		
		
		write-host("Recent Files and folders in root drive:")
 
        $Response=Invoke-WebRequest -Method GET -Uri ($ResourceID+"_api/v2.0"+"/drive/recent") -Header @{ Authorization = "BEARER "+$Authentication.access_token} -ErrorAction Stop
 
        $responseObject = ($Response.Content|ConvertFrom-Json).value
 
        $responseObject
		
		
		write-host("Create Folder")
 
        $Response=POST /drive/root/children
		Content-Type: application/json

		{
			"name": "FolderA",
			"folder": { }
		}
		
		
 
 
 
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