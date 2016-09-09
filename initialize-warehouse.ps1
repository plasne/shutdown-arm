Param(
  [String] $DBServer = "????",
  [String] $DBName = "????",
  [String] $DBUsername = "????",
  [String] $DBPassword = "????"
)

# open a connection to the SQL database
$connection = New-Object System.Data.SqlClient.SqlConnection
$connection.ConnectionString = "Server=tcp:$DBServer,1433;Initial Catalog=$DBName;Persist Security Info=False;User ID=$DBUsername;Password=$DBPassword;MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;"
$connection.Open()

# create the shutdown log table
$command = $connection.CreateCommand()
$command.CommandText = "CREATE TABLE [dbo].[ShutdownLog] ([Timestamp] [datetime] NOT NULL, [RuleId] [nvarchar](255) NOT NULL, [SubscriptionId] [nvarchar](50) NOT NULL, [Region] [nvarchar](50) NOT NULL, [VM] [nvarchar](255) NOT NULL, [Status] [nvarchar](255) NOT NULL)"
$result = $command.ExecuteNonQuery()
if ($result -eq -1) {
  Write-Output "The ShutdownLog table was successfully created."
}
