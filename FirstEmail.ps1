
## Open Outlook connection and take folder
$olFolderInbox = 6
$outlook = new-object -com outlook.application;
$ns = $outlook.GetNameSpace("MAPI");

$inbox = $ns.GetDefaultFolder($olFolderInbox)

$PSTtoSelect = "C:\Users\en_ti\Documents\Outlook Files\Solutiontest.pst"
$PST = $ns.Stores | ? {$_.FilePath -eq $PSTtoSelect}
$PSTRoot = $PST.GetRootFolder()

#Open Database
$DatabaseName = "C:\Users\en_ti\Desktop\Solution Shop\RGF\Access\Database1.mdb"

#Open Access DB
$ConnectionString = "Provider = Microsoft.ACE.OLEDB.12.0;Data Source=$DatabaseName"
$Connection = New-Object System.Data.OleDb.OleDbConnection $ConnectionString

#SQL Query and create table
$Query = "SELECT * FROM Emails100"
$Command  = New-Object System.Data.OleDb.OleDbCommand $Query, $Connection
$Connection.Open()
$Adapter = New-Object System.Data.OleDb.OleDbDataAdapter $Command
$Dataset = New-Object System.Data.DataSet
[void] $Adapter.Fill($DataSet)
$Connection.Close()


$PSTRoot.Folders

$Emails = $Dataset.Tables[0]
    foreach ($email in $Emails) {
   
        $mail = $outlook.CreateItem(0)
        
        $mail.Subject = $email.Subject
        $mail.CC = $email.CCAddresses
        $mail.BCC = $email.BCCAddresses
        $mail.To = $email.sender
        $mail.HTMLBody=$email.Body
        $mail.to = $email.FromAddress
       # $attachment = $u.Path + $u.Filename
       # $attachment 
       # [void] $Mail.Attachments.Add($attachment) 
       # add-content $log1 "Message Sent: $toaddress1" 
      #  $mail.ReceivedTime = $u.EmailDate
    $mail.Save()
    ## add-content $log1 $mail.EntryID -  $u.MessageId
    write-host $mail.EntryID is $u.MessageId

   $targetfolder = $PSTRoot.Folders | where-object { $_.name -eq "TEST" }


   [void] $mail.Move($targetfolder)
   
}


    ## add-content $log1 $mail.EntryID -  $u.MessageId
