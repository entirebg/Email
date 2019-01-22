
#----------------- Open Outlook connection and take folder---------
        $olFolderInbox = 6
        $outlook = new-object -com outlook.application;
        $ns = $outlook.GetNameSpace("MAPI");


#---------------------- For creating different PST file ------------
        #$namespace.AddStore(“PST_Test.pst”)
        #$namespace.Session.Folders.GetLast().Name = (“PST_Test”)
        #$TargetFolder = $namespace.Session.Folders.GetLast()

        $inbox = $ns.GetDefaultFolder($olFolderInbox)

        $PSTtoSelect = "C:\Users\en_ti\Documents\Outlook Files\Solutiontest.pst"
        $PST = $ns.Stores | ? {$_.FilePath -eq $PSTtoSelect}
        $PSTRoot = $PST.GetRootFolder()

#-------------------------------Open Database---------------
        $DatabaseName = "C:\Users\en_ti\Desktop\Solution Shop\RGF\Access\Database1.mdb"
        $ConnectionString = "Provider = Microsoft.ACE.OLEDB.12.0;Data Source=$DatabaseName"
        $Connection = New-Object System.Data.OleDb.OleDbConnection $ConnectionString

#--------------------------SQL Query and create table----------
        $Query = "SELECT * FROM Emails100"
        $Command  = New-Object System.Data.OleDb.OleDbCommand $Query, $Connection
        $Connection.Open()
        $Adapter = New-Object System.Data.OleDb.OleDbDataAdapter $Command
        $Dataset = New-Object System.Data.DataSet
        [void] $Adapter.Fill($DataSet)
        $Connection.Close()

#---------------- Create new email item ----------------
        $Emails = $Dataset.Tables[0]

    foreach ($email in $Emails) {
        $mail = $outlook.CreateItem(0)
        $mail.Subject = $email.Subject
        $mail.CC = $email.CCAddresses
        $mail.BCC = $email.BCCAddresses
        $mail.To = $email.sender
        $mail.HTMLBody=$email.Body
        $mail.to = $email.FromAddress
     
     
#------------ Atachments -------------------------------------     
       # $attachment = $email.Path + $$email.Filename
       # $attachment 
       # [void] $Mail.Attachments.Add($attachment) 
       # add-content $log1 "Message Sent: $toaddress1" 
       # $mail.ReceivedTime = $email.EmailDate

    $mail.Save()

#------------- Check and create Folder ---------------- 
    If ($targetfolder = $PSTRoot.Folders | where-object { $_.name -eq $email.FolderName }) 
         {[void] $mail.Move($targetfolder)}
    else 
         {[void]$PSTRoot.Folders.Add($email.FolderName)
         $targetfolder = $PSTRoot.Folders | where-object { $_.name -eq $email.FolderName}
         [void] $mail.Move($targetfolder)}
   }
   
  