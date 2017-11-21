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
$TestSiteURL = "https://arstconnazureo365.sharepoint.com/SitePages/DevHome.aspx"
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

