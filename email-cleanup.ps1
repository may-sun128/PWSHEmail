### Connect to Outlook 

# add DDLs 
Add-Type -assembly "Microsoft.Office.Interop.Outlook"
add-type -assembly "System.Runtime.Interopservices"

# check if outlook is running
try
{
    $outlook = [Runtime.Interopservices.Marshal]::GetActiveObject('Outlook.Application')
    $outlookWasAlreadyRunning = $true
}

# if not, write message
catch
{
    # try
    # {
    #     $Outlook = New-Object -comobject Outlook.Application
    #     $outlookWasAlreadyRunning = $false
    # }
    # catch
    # {
    #     write-host "You must exit Outlook first."
    #     exit
    # }
    Write-Host "Could not find instance of outlook running."
    $outlookWasAlreadyRunning = $false 
}

# if outlook is running, create a namespace 
# TODO consolidate with the check above 
if ($outlookWasAlreadyRunning) 
{
    $namespace = $Outlook.GetNameSpace("MAPI")    
}
else
{
    Write-Host "Namespace not created."
}

Write-Host $outlookWasAlreadyRunning


### Get Outlook Objects


# get inbox object 
$inbox = $namespace.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderInbox)

# get emails (TODO figure out what data type $emails is)
$emails = $inbox.Items | Where-Object {$_.Subject -like $subjectComparisonExpression}

# Write-Host $inbox
# Write-Host $emails

# Loop through emails 
ForEach ($email in $emails)
{
    Write-Host $email
}

# wait to close application; for debugging  
Read-Host "Press enter to quit."