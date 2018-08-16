# Using PowerShell to create Multiple Azure NSG Rules from Excel file 


I recently came across a scenario, where my colleague was trying to create multiple Network Security group rules from input file for a customer.  While creating the NSGs from PowerShell, if you provide input values directly in cmdlet or using variables, it works absolutely fine. 


##### This works fine

    Add-AzureRmNetworkSecurityRuleConfig -NetworkSecurityGroup $nsg -Name Baseline-rule -Description "Test-Baseline" -Access Allow -Protocol Tcp -Direction Inbound 
    ` -Priority $newRulePriority -SourceAddressPrefix  100.82.10.150,100.82.10.160,100.157.10.150,100.221.10.160 -SourcePortRange * -DestinationAddressPrefix * -DestinationPortRange 20-23,4750,139,445


**But the goal** was to read rules from an excel file & apply them to an NSG, but while doing so I’m getting an exception in the Set-AzureRmNetworkSecurityGroup statement, while saving state of NSG as below:

    Add-AzureRmNetworkSecurityRuleConfig -NetworkSecurityGroup $nsg -Name $RuleName -Description $Description -Access $Action -Protocol $Protocol -Direction $Direction 
    `  -Priority $Priority -SourceAddressPrefix $SourceAddress.ToString() -SourcePortRange $SourcePort.ToString()
    -DestinationAddressPrefix $DestinationAddress.ToString() -DestinationPortRange $DestinationPort.ToString() 
    
    Set-AzureRmNetworkSecurityGroup -NetworkSecurityGroup $nsg


All the variables above are passed by type-casting in ‘String’ data type.

##### This returns with an error "Security Rule has invalid Port range value provided "

![Error](https://github.com/Daya-Patil/Azure-NSG-In-Bulk/blob/readme-edits/Error.png)

Below is the Input excel file rule screenshot:
  
![Inputfile](https://github.com/Daya-Patil/Azure-NSG-In-Bulk/blob/readme-edits/InputFile.png)

Even if we do not convert the values for the ports with toString() to a String, you still get the same error.

So, what is happening here?

If you drill in detail and look at the actual Json file posted to the server, you will see below difference.

 ##### The config output for working one:

    "DestinationPortRange": [
      "20-23",
      "4750",
      "139",
      "445"
    ],
    "SourceAddressPrefix": [
      "155.64.64.0/24",
      "155.64.106.0/23"
    ]
##### Below is the config output for non-working one:

    "DestinationPortRange": [
      "\"22-23\",\"443\""
       ],
    "SourceAddressPrefix": [
      "\"100.82.10.150\",\"100.82.10.160\""
    ]


So it’s clear that the input is not getting sent in the expected format to the server. 

Let’s look at it again:















  
 





