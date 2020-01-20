##########################################################################
# This code does the following:
# 1) Finds the reference csv file from the specified SharePoint Library and creates a local copy of it
# 2) Then it parses local csv file and populates specified SharePoint List with csv data
# 3) Then the local csv file is destroyed
##########################################################################

#Load SharePoint CSOM Assemblies <-- You may have to download them. check to see if path exists
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

# User Credentials 
$UserName = "example@example.onmicrosoft.com" # <-- Use your Office 365 login
$Password = Get-Content "C:\temp\ExportedPassword.txt" #<-- Create txt file with password.

# Function below downloads csv file from specified location on SharePoint
Function Download-FileFromLibrary()
{
    param
    (
        [Parameter(Mandatory=$true)] [string] $SiteURL,
        [Parameter(Mandatory=$true)] [string] $SourceFile,
        [Parameter(Mandatory=$true)] [string] $TargetFile
    )
 
    Try {
        #Setup Credentials to connect
        $Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($UserName,(ConvertTo-SecureString $Password -AsPlainText -Force))
 
        #Setup the context
        $Ctx = New-Object Microsoft.SharePoint.Client.ClientContext($SiteURL)
        $Ctx.Credentials = $Credentials
     
        #sharepoint online powershell download file from library
        $FileInfo = [Microsoft.SharePoint.Client.File]::OpenBinaryDirect($Ctx,$SourceFile)
        $WriteStream = [System.IO.File]::Open($TargetFile,[System.IO.FileMode]::Create)
        $FileInfo.Stream.CopyTo($WriteStream)
        $WriteStream.Close()
 
        Write-host -f Green "File '$SourceFile' Downloaded to '$TargetFile' Successfully!" $_.Exception.Message
  }
    Catch {
        write-host -f Red "Error Downloading File!" $_.Exception.Message
    }
}
 
#Set parameter values
$SiteURL="https://[yourTenant].sharepoint.com/sites/[yourSiteName]/" # <-- replace the tenant name
$SourceFile="/sites/[yourSiteName]/[LibraryName]/[fileName.csv]"  # <-- replace with your relative URL where the file is located
$TargetFile="C:\Temp\[fileName.csv]" # <-- this is the location on your machine. Change as needed.
Write-host "Local file C:\Temp\[fileName.csv] has been created"
 
#Call the function to download file
Download-FileFromLibrary -SiteURL $SiteURL -SourceFile $SourceFile -TargetFile $TargetFile

##Variables for Processing
$ListName ="PowershellList" # <--List where you want your data to populate
$ImportFile ="C:\Temp\[fileName.csv]" # <-- The local file where the csv data is coming from

#Setup Credentials to connect .... this was part of the original code
$Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($UserName,(ConvertTo-SecureString $Password -AsPlainText -Force))

#Set up the context 
$Context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteUrl)
$Context.Credentials = $credentials

#Get the List 
$List = $Context.web.Lists.GetByTitle($ListName)
Write-host ($ListName)
Write-host ($LibraryName)
Write-host ($ImportFile)
Write-host -f Green "Successfully connected"

# #Get the Data from CSV and Add to SharePoint List 
$data = Import-Csv $ImportFile
Foreach ($row in $data) {
     
    #add item to List   item===SharePoint Column name  row===CSV Column Name
    $ListItemInfo = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation
    $Item = $List.AddItem($ListItemInfo)
    $Item["Title"] = $row.Something
    $Item["FirstName"] = $row.FirstName  # <-- replace these with data that matches your csv / list columns
    $Item["LastName"] = $row.LastName
    $Item["Phone"] = $row.Phone
    $Item["Email"] = $row.Email
    $Item.Update()
    $Context.ExecuteQuery()
    
}
Write-host -f Green "CSV data Imported to SharePoint List Successfully!"
Remove-Item -Path C:\Temp\LRcsv.csv -Force
Write-host -f Yellow "Local file at C:\Temp\[fileName.csv] has been removed"
