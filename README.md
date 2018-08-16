# Using PowerShell to create Multiple Azure NSG Rules from Excel file 


I recently came across a scenario, where my colleague was trying to create multiple Network Security group rules from input file for a customer.  While creating the NSGs from PowerShell, if you provide input values directly in cmdlet or using variables, it works absolutely fine. 


// This works fine

    $InputObject = @{"TestSQLVMName" = "#TestSQLVMName" ; "TestSQLVMRG" = "#TestSQLVMRG" ; "ProdSQLVMName" = "#ProdSQLVMName" ; "ProdSQLVMRG" = "#ProdSQLVMRG"; "Paths" = @{"1"="#sqlserver:\sql\sqlazureVM\default\availabilitygroups\ag1";"2"="#sqlserver:\sql\sqlazureVM\default\availabilitygroups\ag2"}}
        New-AzureRmAutomationVariable -Name "#RecoveryPlanName" -ResourceGroupName "#AutomationAccountResourceGroup" -AutomationAccountName "#AutomationAccountName" -Value $RPDetails -Encrypted $false 
        New-AzureRmAutomationVariable -Name "#RecoveryPlanName" -ResourceGroupName "#AutomationAccountResourceGroup" -AutomationAccountName "#AutomationAccountName" -Value $RPDetails -Encrypted $false 

 
  
  Add-AzureRmNetworkSecurityRuleConfig
  -NetworkSecurityGroup $nsg -Name Baseline-rule -Description "Test-Baseline"
  -Access Allow -Protocol Tcp -Direction Inbound -Priority $newRulePriority
  -SourceAddressPrefix
  100.82.10.150,100.82.10.160,100.157.10.150,100.221.10.160 -SourcePortRange *
  -DestinationAddressPrefix * -DestinationPortRange 20-23,4750,139,445 


  
 






  
 





