## Open Outlook connection and take folder
$olFolderInbox = 6
$outlook = new-object -com outlook.application;
$ns = $outlook.GetNameSpace("MAPI");
$inbox = $ns.GetDefaultFolder($olFolderInbox)
#Open Database
$DatabaseName = "C:\Users\en_ti\Desktop\Solution Shop\RGF\Access\Database1.mdb"

#Open Access DB
$ConnectionString = "Provider = Microsoft.ACE.OLEDB.12.0;Data Source=$DatabaseName"
$Connection = New-Object System.Data.OleDb.OleDbConnection $ConnectionString

#SQL Query and create table
$Query = "SELECT * FROM EMAIL_FOLDERS"
$Command  = New-Object System.Data.OleDb.OleDbCommand $Query, $Connection
$Connection.Open()
$Adapter = New-Object System.Data.OleDb.OleDbDataAdapter $Command
$Dataset = New-Object System.Data.DataSet
[void] $Adapter.Fill($DataSet)
$Connection.Close()

$raws = $Dataset.Tables[0]
    foreach ($raw in $raws) {
    
    if ($raw.UserID  -eq "BROWE"){ Write-Host $raw.FolderName }
    
    }