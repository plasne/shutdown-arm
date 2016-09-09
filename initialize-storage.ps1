Param(
  [String] $SubscriptionId = "????",
  [String] $StorageAccountName = "????",
  [String] $StorageAccountKey = "????",
  [String] $RulesTableName = "automationRules",
  [String] $TimezoneTableName = "automationTimezones"
)

function Add-Rule() {
  [CmdletBinding()]
  Param(
     [Parameter(Mandatory=$True, ValueFromPipeline=$True)]  [System.Object] $Table,
     [Parameter(Mandatory=$True)]  [String] $PartitionKey,
     [Parameter(Mandatory=$True)]  [String] $RowId,
     [Parameter(Mandatory=$True)]  [String] $SubscriptionId,
     [Parameter(Mandatory=$False)] [String] $Region,
     [Parameter(Mandatory=$False)] [String] $VM,
     [Parameter(Mandatory=$False)] [String] $ShutdownWindowFrom,
     [Parameter(Mandatory=$False)] [String] $ShutdownWindowTo,
     [Parameter(Mandatory=$False)] [String] $GracePeriod,
     [Parameter(Mandatory=$True)]  [Bool]   $Shutdown,
     [Parameter(Mandatory=$True)]  [Bool]   $Configure
  )

  # create the entity
  $entity = New-Object -TypeName Microsoft.WindowsAzure.Storage.Table.DynamicTableEntity -ArgumentList PartitionKey, $RowId
  $entity.Properties.Add("SubscriptionId", $SubscriptionId)
  $entity.Properties.Add("Region", $Region)
  $entity.Properties.Add("VM", $VM)
  $entity.Properties.Add("ShutdownWindowFrom", $ShutdownWindowFrom)
  $entity.Properties.Add("ShutdownWindowTo", $ShutdownWindowTo)
  $entity.Properties.Add("GracePeriod", $GracePeriod)
  $entity.Properties.Add("Shutdown", $Shutdown)
  $entity.Properties.Add("Configure", $Configure)

  # add the entity
  $result = $Table.CloudTable.Execute([Microsoft.WindowsAzure.Storage.Table.TableOperation]::Insert($entity))

}

function Add-Timezone() {
  [CmdletBinding()]
  Param(
     [Parameter(Mandatory=$True, ValueFromPipeline=$True)]  [System.Object] $Table,
     [Parameter(Mandatory=$True)]  [String] $PartitionKey,
     [Parameter(Mandatory=$True)]  [String] $Region,
     [Parameter(Mandatory=$True)]  [String] $Location,
     [Parameter(Mandatory=$True)]  [String] $Timezone
  )

  # create the entity
  $entity = New-Object -TypeName Microsoft.WindowsAzure.Storage.Table.DynamicTableEntity -ArgumentList $PartitionKey, $Region
  $entity.Properties.Add("Location", $Location)
  $entity.Properties.Add("Timezone", $Timezone)

  # add the entity
  $result = $Table.CloudTable.Execute([Microsoft.WindowsAzure.Storage.Table.TableOperation]::Insert($entity))

}

# get the storage context
$ctx = New-AzureStorageContext $StorageAccountName -StorageAccountKey $StorageAccountKey

# create the rules table (with a sample row) if it doesn't exist
$rulesTable = Get-AzureStorageTable -Name $RulesTableName -Context $ctx
if (!$rulesTable) {
  $rulesTable = New-AzureStorageTable �Name $RulesTableName �Context $ctx
  $rulesTable | Add-Rule -PartitionKey "automation" -RowId 1 -SubscriptionId bb8b8c18-67da-4a87-be7a-680da44f18e0 -Shutdown $True -Configure $True
}

# create the timezone table with all rows
$timezoneTable = Get-AzureStorageTable -Name $TimezoneTableName -Context $ctx
if (!$timezoneTable) {
  $timezoneTable = New-AzureStorageTable -Name $TimezoneTableName -Context $ctx
  $timezoneTable | Add-Timezone -PartitionKey "timezone" -Region "eastasia" -Location "East Asia" -Timezone "China Standard Time"
  $timezoneTable | Add-Timezone -PartitionKey "timezone" -Region "southeastasia" -Location "Southeast Asia" -Timezone "Singapore Standard Time"
  $timezoneTable | Add-Timezone -PartitionKey "timezone" -Region "centralus" -Location "Central US" -Timezone "Central Standard Time"
  $timezoneTable | Add-Timezone -PartitionKey "timezone" -Region "eastus" -Location "East US" -Timezone "Eastern Standard Time"
  $timezoneTable | Add-Timezone -PartitionKey "timezone" -Region "eastus2" -Location "East US 2" -Timezone "Eastern Standard Time"
  $timezoneTable | Add-Timezone -PartitionKey "timezone" -Region "westus" -Location "West US" -Timezone "Pacific Standard Time"
  $timezoneTable | Add-Timezone -PartitionKey "timezone" -Region "westus2" -Location "West US 2" -Timezone "Pacific Standard Time"
  $timezoneTable | Add-Timezone -PartitionKey "timezone" -Region "westcentralus" -Location "West Central US" -Timezone "Mountain Standard Time"
  $timezoneTable | Add-Timezone -PartitionKey "timezone" -Region "northcentralus" -Location "North Central US" -Timezone "Central Standard Time"
  $timezoneTable | Add-Timezone -PartitionKey "timezone" -Region "southcentralus" -Location "South Central US" -Timezone "Central Standard Time"
  $timezoneTable | Add-Timezone -PartitionKey "timezone" -Region "northeurope" -Location "North Europe" -Timezone "GMT Standard Time"
  $timezoneTable | Add-Timezone -PartitionKey "timezone" -Region "westeurope" -Location "West Europe" -Timezone "Central European Standard Time"
  $timezoneTable | Add-Timezone -PartitionKey "timezone" -Region "japanwest" -Location "Japan West" -Timezone "Tokyo Standard Time"
  $timezoneTable | Add-Timezone -PartitionKey "timezone" -Region "japaneast" -Location "Japan East" -Timezone "Tokyo Standard Time"
  $timezoneTable | Add-Timezone -PartitionKey "timezone" -Region "brazilsouth" -Location "Brazil South" -Timezone "E. South America Standard Time"
  $timezoneTable | Add-Timezone -PartitionKey "timezone" -Region "australiaeast" -Location "Australia East" -Timezone "AUS Eastern Standard Time"
  $timezoneTable | Add-Timezone -PartitionKey "timezone" -Region "australiasoutheast" -Location "Australia Southeast" -Timezone "AUS Eastern Standard Time"
  $timezoneTable | Add-Timezone -PartitionKey "timezone" -Region "southindia" -Location "South India" -Timezone "India Standard Time"
  $timezoneTable | Add-Timezone -PartitionKey "timezone" -Region "centralindia" -Location "Central India" -Timezone "India Standard Time"
  $timezoneTable | Add-Timezone -PartitionKey "timezone" -Region "westindia" -Location "West India" -Timezone "India Standard Time"
  $timezoneTable | Add-Timezone -PartitionKey "timezone" -Region "canadacentral" -Location "Canada Central" -Timezone "Eastern Standard Time"
  $timezoneTable | Add-Timezone -PartitionKey "timezone" -Region "canadaeast" -Location "Canada East" -Timezone "Eastern Standard Time"
}