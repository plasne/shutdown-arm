Param(
  [String] $DBServer = "????",
  [String] $DBName = "????",
  [String] $EmailFrom = "????",
  [String] $EmailTo = "????"
)

function New-LogEntry {
  Param(
    [Parameter(Mandatory=$True)]  [DateTime]  $Timestamp,
    [Parameter(Mandatory=$True)]  [String]    $RuleId,
    [Parameter(Mandatory=$True)]  [String]    $SubscriptionId,
    [Parameter(Mandatory=$True)]  [String]    $Region,
    [Parameter(Mandatory=$True)]  [String]    $VM,
    [Parameter(Mandatory=$True)]  [String]    $Status
  )

  $obj = New-Object System.Object
  $obj | Add-Member -type NoteProperty -name Timestamp -value $Timestamp
  $obj | Add-Member -type NoteProperty -name RuleId -value $RuleId
  $obj | Add-Member -type NoteProperty -name SubscriptionId -value $SubscriptionId
  $obj | Add-Member -type NoteProperty -name Region -value $Region
  $obj | Add-Member -type NoteProperty -name VM -value $VM
  $obj | Add-Member -type NoteProperty -name Status -value $Status

  return $obj
}

# credentials for local testing (comment these out for production)
# $dbUsername = "????"
# $dbPassword = "????"
# $dbSecurePassword = ConvertTo-SecureString $dbPassword -AsPlainText -Force
# $dbSecurePassword.MakeReadOnly()
# $dbSqlCredential = New-Object System.Data.SqlClient.SqlCredential($dbUsername, $dbSecurePassword)

# credentials for production (comment these out for local testing)
$dbPsCredential = Get-AutomationPSCredential -Name "DataWarehouse"
$dbPsCredential.Password.MakeReadOnly()
$dbSqlCredential = New-Object System.Data.SqlClient.SqlCredential($dbPsCredential.UserName, $dbPsCredential.Password)

# open the database connection
$connection = New-Object System.Data.SqlClient.SqlConnection("Server=tcp:$DBServer,1433;Initial Catalog=$DBName;Persist Security Info=False;MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;", $dbSqlCredential)
$connection.Open()

# query all log entries in the past 24 hours
$entries = New-Object System.Collections.Generic.List[System.Object]
$command = $connection.CreateCommand()
$command.CommandText = "SELECT * FROM [dbo].[ShutdownLog] WHERE [Timestamp] > DateAdd(hour, -24, GetDate()) ORDER BY [Timestamp] DESC"
$reader = $command.ExecuteReader()
while ($reader.Read()) {
  $entry = New-LogEntry -Timestamp $reader["Timestamp"] -RuleId $reader["RuleId"] -SubscriptionId $reader["SubscriptionId"] -Region $reader["Region"] -VM $reader["VM"] -Status $reader["Status"]
  $entries.Add($entry)
}

# format into a table
$output = ($entries | ConvertTo-Html Timestamp, RuleId, SubscriptionId, Region, VM, Status) | Out-String

# credentials for local testing (comment these out in production)
# $sgUsername = "????"
# $sgPassword = "????"
# $sgSecurePassword = ConvertTo-SecureString $sgPassword -AsPlainText -Force
# $sgPsCredential = New-Object System.Management.Automation.PSCredential $sgUsername, $sgSecurePassword

# credentials for production (comment this out for local testing)
$sgPsCredential = Get-AutomationPSCredential -Name "SendGrid"

# email results
Send-MailMessage -smtpServer "smtp.sendgrid.net" -Credential $sgPsCredential -Usessl -Port 587 -From $EmailFrom -To $EmailTo -Subject "VMs shutdown in the past 24 hours" -BodyAsHtml $output

# output
$entries | Format-Table Timestamp, RuleId, SubscriptionId, Region, VM, Status
