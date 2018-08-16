$excel=new-object -com excel.application
$wb =$excel.workbooks.open("C:\users\dapatil\Desktop\nsg1.xlsx")
$sh = $wb.Sheets.Item(1)
$priority = $sh.Cells.Item(2, 1).text
$RuleName = $sh.Cells.Item(2, 2).text
$SourceAddress = $sh.Cells.Item(2, 3).text
$SourcePort = $sh.Cells.Item(2, 4).text
$DestinationAddress = $sh.Cells.Item(2, 5).text
$DestinationPort    = $sh.Cells.Item(2, 6).text
$ports = New-Object System.Collections.Generic.List``1[System.String]
$ports.add('22-23')
$ports.add(443)
$Protocol    = $sh.Cells.Item(2, 7).text
$Action      = $sh.Cells.Item(2, 8).text
$Direction   = $sh.Cells.Item(2, 9).text
$Description = $sh.Cells.Item(2, 10).text
$SourceIP = New-Object System.Collections.Generic.List``1[System.String]
$SourceIP.add('100.82.10.150')
$SourceIP.add('100.82.10.160')

$NSG = New-AzureRmNetworkSecurityGroup -ResourceGroupName 'Exchlab' -Location 'East US 2' -Name 'Copy-of-Windows-NSG4'
Add-AzureRmNetworkSecurityRuleConfig -NetworkSecurityGroup $nsg -Name new6 -Description $Description -Access $Action -Protocol $Protocol -Direction $Direction -Priority $Priority -SourceAddressPrefix $SourceIP -SourcePortRange $SourcePort.ToString() -DestinationAddressPrefix $DestinationAddress.ToString() -DestinationPortRange $Ports

Set-AzureRmNetworkSecurityGroup -NetworkSecurityGroup $nsg
