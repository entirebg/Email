function func1{
       
     foreach ($singleemail in $emailDB){
     # write-host res is $res num is $num[0] mytemp $MyTemp
      if ($res -eq $singleemail.Recid) {
       $MyTemp += $singleemail.FolderName+"\"+$MyTemp
       
       if ($singleemail.Parent.Length){$result =$singleemail.UserID+$MyTemp
        
           } else { 
           $res = $singleemail.Parent
             func1 ($res) }
               }
            }
            return($result)
          }


$DatabaseName = "C:\Users\en_ti\Desktop\Solution Shop\RGF\Access\Database1.mdb"
$Query = "SELECT * FROM Email_Folders where UserId = 'Admin'"
$ConnectionString = "Provider = Microsoft.ACE.OLEDB.12.0;Data Source=$DatabaseName"
$Connection = New-Object System.Data.OleDb.OleDbConnection $ConnectionString
$Command  = New-Object System.Data.OleDb.OleDbCommand $Query, $Connection
$Connection.Open()
$Adapter = New-Object System.Data.OleDb.OleDbDataAdapter $Command
$Dataset = New-Object System.Data.DataSet
[void] $Adapter.Fill($DataSet)
$Connection.Close()
$emailDB = $Dataset.Tables[0]
foreach ($singleemail in $emailDB) {
     func1 ($res = $singleemail.Recid) }
     
#if ($singleemail.Parent.Length){ Write-Host  $singleemail.Parent.Length not 0} 
#else { Write-Host  $singleemail.Parent.Length is 0}
   
 
   

