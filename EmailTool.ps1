# This tool allow you to download and rename
# email file attachement to a desired location 
# from your inbox folders in your outlook application.

#-----------TEST----------------- 
# use MAPI name space
#$outlook = new-object -com outlook.application; 
#$mapi = $outlook.GetNameSpace("MAPI");

# we need Inbox folder
#$olDefaultFolderInbox = 6
#$inbox = $mapi.GetDefaultFolder($olDefaultFolderInbox) 

# see folders under Inbox
#$inbox.Folders | SELECT FolderPath

#-----------END-TEST-----------------

# link to the folder 
$olFolderPath = "\\romainob98@gmail.com\Inbox\Test"
# set the desired file name
$attachmentFileName = 'test2.xlsx'
# set the location to temporary file
$filePath = "U:\Projects\EmailProj"
# use MAPI name space
$outlook = new-object -com outlook.application; 
$mapi = $outlook.GetNameSpace("MAPI");
# set the Inbox folder id
$olDefaultFolderInbox = 6
$inbox = $mapi.GetDefaultFolder($olDefaultFolderInbox) 
# access the target subfolder
$olTargetFolder = $inbox.Folders | Where-Object { $_.FolderPath -eq $olFolderPath }
# load emails
$emails = $olTargetFolder.Items
# process the emails
foreach ($email in $emails) {
    
    # format the timestamp
    $timestamp = $email.ReceivedTime.ToString("MM-dd-yyyy")
    # filter out the attachments
    #$email.Attachments | Where-Object {$_.FileName -eq $attachmentFileName} | foreach {
    $email.Attachments | foreach { 
        
        # insert the timestamp into the file name
        $fileName = $_.FileName
        #write-host $_.FileName.ex
        $pos = $fileName.IndexOf(".")
        $extension = $fileName.Substring($pos)
        $fileName = $fileName.Replace($fileName,$timestamp+$extension)
        # save the attachment
        $_.saveasfile((Join-Path $filePath $fileName)) 
    } 
} 
