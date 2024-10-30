# Create an Outlook COM object
$Outlook = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNamespace("MAPI")

# Specify the folder path (e.g., "Inbox\SubfolderName" or just "Inbox")
$folderPath = "Inbox"

# Get the folder
$Folder = $Namespace.Folders.Item(1).Folders.Item($folderPath)

# Check if the folder exists
if (-not $Folder) {
    Write-Output "Folder not found: $folderPath"
    exit
}

# Get all items in the folder
$Items = $Folder.Items

# Loop through each item in the folder
foreach ($Item in $Items) {
    # Check if the item is a MailItem
    if ($Item -is [Microsoft.Office.Interop.Outlook.MailItem]) {
        # Get subject and body
        $Subject = $Item.Subject
        $Body = $Item.Body

        # Output subject and body
        Write-Output "Subject: $Subject"
        Write-Output "Body: $Body"
        Write-Output "----------------------------------"
    }
}

# Release COM objects
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Folder) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Namespace) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook) | Out-Null
