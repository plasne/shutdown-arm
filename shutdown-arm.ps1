  Param(
    [String] $TenantId = "????",
    [String] $SubscriptionId = "????",
    [String] $StorageAccountResourceGroup = "????",
    [String] $StorageAccountName = "????",
    [String] $RulesTableName = "automationRules",
    [String] $DBServer = "????",
    [String] $DBName = "????"
  )

  # global variables
  $operations = New-Object System.Collections.Generic.List[System.Object]
  $potential = New-Object System.Collections.Generic.List[System.Object]
  $shutdownQ = [System.Collections.Queue]::Synchronized( (New-Object System.Collections.Queue) )
  $profilePath = ".\profile.json"
  $connection = New-Object System.Data.SqlClient.SqlConnection

  function New-Operation {
    Param(
                                    [String]        $RuleId,
      [Parameter(Mandatory=$True)]  [String]        $SubscriptionId,
                                    [String]        $Region,
      [Parameter(Mandatory=$True)]  [System.Object] $VM,
                                    [Int]           $Priority,
                                    [String]        $Shutdown
    )

    $obj = New-Object System.Object
    $obj | Add-Member -type NoteProperty -name RuleId -value $RuleId
    $obj | Add-Member -type NoteProperty -name SubscriptionId -value $SubscriptionId
    $obj | Add-Member -type NoteProperty -name Region -value $Region
    $obj | Add-Member -type NoteProperty -name VM -value $VM
    $obj | Add-Member -type NoteProperty -name VMName -value $VM.Name
    $obj | Add-Member -type NoteProperty -name Priority -value $Priority
    $obj | Add-Member -type NoteProperty -name Shutdown -value $Shutdown

    return $obj
  }

  function Find-VMs {
    Param(
      [Parameter(Mandatory=$True)]  [System.Object]  $Rule
    )

    If ($Rule.VMName.Length -gt 0) {
      $found = $potential | Where-Object { $_.SubscriptionId -eq $Rule.SubscriptionId -and $_.Region -eq $Rule.Region -and $_.VMName -eq $Rule.VMName }
    } ElseIf ($Rule.Region -gt 0) {
      $found = $potential | Where-Object { $_.SubscriptionId -eq $Rule.SubscriptionId -and $_.Region -eq $Rule.Region }
    } Else {
      $found = $potential | Where-Object SubscriptionId -eq $Rule.SubscriptionId
    }

    return $found
  }

  function Get-Rules {
    Param(
      [Parameter(Mandatory=$True)] [String] $StorageAccountName,
      [Parameter(Mandatory=$True)] [String] $StorageAccountKey,
      [String] $RulesTableName = "automationRules"
    )

    function Is-Integer($Value) {
      return $Value -match "(?<![-.])\b[0-9]+\b(?!\.[0-9])"
    }

    # connect to Azure Storage Table
    $ctx = New-AzureStorageContext $StorageAccountName -StorageAccountKey $StorageAccountKey
    $table = Get-AzureStorageTable �Name $RulesTableName -Context $ctx

    # create a query
    $query = New-Object Microsoft.WindowsAzure.Storage.Table.TableQuery
    $list = New-Object System.Collections.Generic.List[String]
    $list.Add("SubscriptionId")
    $list.Add("Region")
    $list.Add("VM")
    $list.Add("ShutdownWindowFrom")
    $list.Add("ShutdownWindowTo")
    $list.Add("GracePeriod")
    $list.Add("Shutdown")
    $list.Add("Configure")

    # execute the query
    $query.FilterString = "PartitionKey eq 'automation'"
    $query.SelectColumns = $list
    $query.TakeCount = 1000
    $entities = $table.CloudTable.ExecuteQuery($query)

    # create consumable objects from the returns
    $rules = New-Object System.Collections.Generic.List[System.Object]
    ForEach ($entity in $entities) {
      if ($entity.Properties.Shutdown.StringValue -eq "ignore") {
        // ignore
      } else {
        $obj = New-Object System.Object
        $obj | Add-Member -type NoteProperty -name Id -value $entity.RowKey
        $obj | Add-Member -type NoteProperty -name SubscriptionId -value $entity.Properties.SubscriptionId.StringValue
        $obj | Add-Member -type NoteProperty -name Region -value $entity.Properties.Region.StringValue
        $obj | Add-Member -type NoteProperty -name VM -value $entity.Properties.VM.StringValue
        if (Is-Integer $entity.Properties.ShutdownWindowFrom.StringValue) {
          $val = [Int]$entity.Properties.ShutdownWindowFrom.StringValue
          $obj | Add-Member -type NoteProperty -name ShutdownWindowFrom -value $val
        } else {
          $obj | Add-Member -type NoteProperty -name ShutdownWindowFrom -value 19    # default: 7pm
        }
        if (Is-Integer $entity.Properties.ShutdownWindowTo.StringValue) {
          $val = $entity.Properties.ShutdownWindowTo.StringValue
          $obj | Add-Member -type NoteProperty -name ShutdownWindowTo -value $val
        } else {
          $obj | Add-Member -type NoteProperty -name ShutdownWindowTo -value 5      # default: 5am
        }
        if (Is-Integer $entity.Properties.GracePeriod.StringValue) {
          $val = $entity.Properties.GracePeriod.StringValue
          $obj | Add-Member -type NoteProperty -name GracePeriod -value $val
        } else {
          $obj | Add-Member -type NoteProperty -name GracePeriod -value 4     # default: 4 hours
        }
        $obj | Add-Member -type NoteProperty -name Shutdown -value $entity.Properties.Shutdown.StringValue
        $obj | Add-Member -type NoteProperty -name Configure -value $entity.Properties.Configure.StringValue
        $rules.Add($obj)
      }
    }

    return $rules
  }

  function Get-Timezones {
    Param(
      [Parameter(Mandatory=$True)] [String] $StorageAccountName,
      [Parameter(Mandatory=$True)] [String] $StorageAccountKey,
      [String] $TimezoneTableName = "automationTimezones"
    )

    # connect to Azure Storage Table
    $ctx = New-AzureStorageContext $StorageAccountName -StorageAccountKey $StorageAccountKey
    $table = Get-AzureStorageTable �Name $TimezoneTableName -Context $ctx

    # create a query
    $query = New-Object Microsoft.WindowsAzure.Storage.Table.TableQuery
    $list = New-Object System.Collections.Generic.List[String]
    $list.Add("Timezone")

    # execute the query
    $query.FilterString = "PartitionKey eq 'timezone'"
    $query.SelectColumns = $list
    $query.TakeCount = 1000
    $entities = $table.CloudTable.ExecuteQuery($query)

    # create consumable objects from the returns
    $timezones = New-Object System.Collections.Generic.List[System.Object]
    ForEach ($entity in $entities) {
      $obj = New-Object System.Object
      $obj | Add-Member -type NoteProperty -name Region -value $entity.RowKey
      $obj | Add-Member -type NoteProperty -name Location -value $entity.Properties.Location.StringValue
      $obj | Add-Member -type NoteProperty -name Timezone -value $entity.Properties.Timezone.StringValue
      $timezones.Add($obj)
    }

    return $timezones
  }

  function Get-Regions {
    Param(
      [Parameter(Mandatory=$True)]  [System.Object] $Rule
    )
    $regions = New-Object System.Collections.Generic.List[String]

    if ($Rule.Region.Length > 0) {

      # add the specified region
      $regions.Add($Rule.Region)

    } else {

      # query to find all regions in the subscription
      $vnets = Get-AzureRmVirtualNetwork
      ForEach ($vnet in $vnets) {
        if (!$regions.Contains($vnet.Location)) {
          $regions.Add($vnet.Location)
        }
      }

    }

    return $regions
  }

  function Get-Priority {
    Param(
      [Parameter(Mandatory=$True)]  [System.Object] $Rule
    )

    $priority = 4    # subscription
    if ($Rule.Region.Length -gt 0) { $priority += 8 }
    #if ($Rule.VNet.Length -gt 0) { $priority += 16 }   # not implemented
    if ($Rule.VM.Length -gt 0) { $priority += 32 }

    # +1 for grace period
    # +2 for future expansion

    return $priority
  }

  function Log-Operation {
    Param(
      [Parameter(Mandatory=$True)]  [System.Object] $Operation,
      [Parameter(Mandatory=$True)]  [String]        $Status
    )

    # open a connection to the SQL database
    if ($connection.State -eq "Closed") {

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

    }

    # log the result
    $command = $connection.CreateCommand()
    $command.CommandText = "INSERT INTO [dbo].[ShutdownLog] ([Timestamp], [RuleId], [SubscriptionId], [Region], [VM], [Status]) VALUES (GetDate(), @ruleId, @subscriptionId, @region, @vm, @status)"
    $command.Parameters.Add("@ruleId", [System.Data.SqlDbType]::NVarChar, 255).Value = $Operation.RuleId
    $command.Parameters.Add("@subscriptionId", [System.Data.SqlDbType]::NVarChar, 50).Value = $Operation.SubscriptionId
    $command.Parameters.Add("@region", [System.Data.SqlDbType]::NVarChar, 50).Value = $Operation.Region
    $command.Parameters.Add("@vm", [System.Data.SqlDbType]::NVarChar, 255).Value = $Operation.VMName
    $command.Parameters.Add("@status", [System.Data.SqlDbType]::NVarChar, 255).Value = $Status
    try {
      $success = $command.ExecuteNonQuery()
    } catch {
      Write-Error "Error writing to database: $_"
    }

  }

  function Start-Shutdown {
    if ($shutdownQ.Count -gt 0) {

      # define the action to take on the background thread
      $script = {
        Param(
          [String]        $TenantId,
          [String]        $ProfilePath,
          [System.Object] $Operation
        )

        # attempt shutdown logging the success or failure
        try {
          $read = Select-AzureRmProfile -Path $ProfilePath
          $sub = Get-AzureRmSubscription -TenantId $TenantId -SubscriptionId $Operation.SubscriptionId | Set-AzureRmContext
          $vm = ((Get-AzureRmVM -ResourceGroupName $Operation.VM.ResourceGroupName -Name $Operation.VM.Name -Status).Statuses | Where-Object Code -like "PowerState/*")
          if ($vm.Code -eq "PowerState/deallocated") {
            return "ignored, already shutdown"
          } else {
            $stop = $Operation.VM | Stop-AzureRmVM -Force
            return "successful shutdown"
          }
        } catch {
          return "failed shutdown: $_"
        }

      }

      # start the job
      $job = Start-Job -ScriptBlock $script -ArgumentList $TenantId, $profilePath, $shutdownQ.Dequeue()

      # log the status
      $status = $job | Wait-Job | Receive-Job
      Log-Operation -Operation $Operation -Status $status
      $Operation | Add-Member -type NoteProperty -name Status -value $status
      Write-Output $Operation

      # iterate
      Register-ObjectEvent -InputObject $job -EventName StateChanged -Action { Start-Shutdown; Unregister-Event $eventsubscriber.SourceIdentifier; Remove-Job $eventsubscriber.SourceIdentifier } | Out-Null

    }
  }

  # login
  $conn = Get-AutomationConnection -Name AzureRunAsConnection
  $acnt = Add-AzureRMAccount -ServicePrincipal -Tenant $TenantId -ApplicationId $conn.ApplicationID -CertificateThumbprint $conn.CertificateThumbprint
  $sub = Get-AzureRmSubscription -TenantId $TenantId -SubscriptionId $SubscriptionId | Set-AzureRmContext

  # get the rules & timezones
  $storageAccountKey = (Get-AzureRmStorageAccountKey -ResourceGroup $StorageAccountResourceGroup -Name $StorageAccountName)[0].Value
  $rules = Get-Rules -StorageAccountName $StorageAccountName -StorageAccountKey $storageAccountKey
  $timezones = Get-Timezones -StorageAccountName $StorageAccountName -StorageAccountKey $storageAccountKey

  # first find anything that is in the grace period
  $subscriptions = $rules | Group-Object $SubscriptionId
  ForEach ($subscription in $subscriptions) {

    # set to the proper subscription
    $subscriptionId = $subscription.Group[0].SubscriptionId;
    $sub = Get-AzureRmSubscription -TenantId $TenantId -SubscriptionId $subscriptionId | Set-AzureRmContext

    # get all VMs in the subscription
    $vms = Get-AzureRmVM
    ForEach ($vm in $vms) {
      $operation = New-Operation -SubscriptionId $subscriptionId -Region $vm.Location -VM $vm
      $potential.Add($operation)
    }

    # get VMs that have started up in the past 24 hours
    $utcNow = [System.TimeZoneInfo]::ConvertTimeToUtc((Get-Date))
    $startups = Get-AzureRmLog -StartTime (Get-Date).AddHours(-24) -Status Accepted | Where-Object { $_.Authorization.Action -eq "Microsoft.Compute/virtualMachines/start/action" }
    ForEach ($startup in $startups) {
      $scope = $startup.Authorization.Scope.Split("/")
      $group = $scope[4]
      $name = $scope[8]
      $diff = $utcNow.Subtract($startup.EventTimestamp)
      ForEach ($rule in $subscription.Group) {
        if ($diff.Hours -lt $rule.GracePeriod) {
          $vms = Find-VMs -Rule $rule
          if ($vms.Count -gt 0) {
            $priority = Get-Priority -Rule $rule
            ForEach ($vm in $vms) {
              if ($vm.VMName -eq $name) {
                $operation = New-Operation -RuleId $rule.Id -SubscriptionId $subscriptionId -Region $vm.Region -VM $vm.VM -Priority ($priority + 1) -Shutdown "grace"
                $operations.Add($operation)
              }
            }
          }
        }
      }
    }

  }

  # establish a set of operations by subscription/region and sometimes subscription/region/vm
  ForEach ($rule in $rules) {

    # set to the proper subscription
    $sub = Get-AzureRmSubscription -TenantId $TenantId -SubscriptionId $rule.SubscriptionId | Set-AzureRmContext

    # determine this rule's priority
    $priority = Get-Priority -Rule $rule

    # iterate through each region
    $regions = Get-Regions -Rule $rule
    ForEach ($region in $regions) {

      # calculate the current time for each region
      $timezoneId = ($timezones | Where-Object Region -eq $region).Timezone
      $timezone = [System.TimeZoneInfo]::GetSystemTimeZones() | Where-Object Id -eq $timezoneId
      $time = [System.TimeZoneInfo]::ConvertTime([DateTime]::Now, $timezone)

      # determine if the region should be shutdown
      if ($time.Hour -ge $rule.ShutdownWindowFrom -and $time.Hour -lt $rule.ShutdownWindowTo) {
        $vms = Find-VMs -Rule $rule
        if ($vms.Count -gt 0) {
          $priority = Get-Priority -Rule $rule
          ForEach ($vm in $vms) {
            $operation = New-Operation -RuleId $rule.Id -SubscriptionId $subscriptionId -Region $vm.Region -VM $vm.VM -Priority $priority -Shutdown $rule.Shutdown
            $operations.Add($operation)
          }
        }
      }

    }

  }

  # for each VM, get the highest priority operation
  $vmsToShutdown = New-Object System.Collections.Generic.List[System.Object]
  $grouped = $operations | Group-Object SubscriptionId, Region, VMName
  ForEach ($entry in $grouped) {
    $sorted = $entry.Group | Sort-Object Priority -Descending
    switch ($sorted[0].Shutdown) {
      "true" {
        $vmsToShutdown.Add($sorted[0])
      }
      "grace" {
        Log-Operation -Operation $sorted[0] -Status "ignored, grace period"
        Write-Output $sorted[0]
      }
      "false" {
        # ignore
      }
    }
  }

  # enqueue all VMs to shutdown
  if ($vmsToShutdown.Count -gt 0) {
    Save-AzureRmProfile -Path $profilePath -Force
    ForEach ($operation in $vmsToShutdown) {
      $shutdownQ.Enqueue($operation)
    }
    $maxConcurrentJobs = 4;
    for ( $i = 0; $i -lt $maxConcurrentJobs; $i++ ) {
      Start-Shutdown
    }
  }
