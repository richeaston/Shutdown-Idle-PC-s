$path = "C:\\AllComputersInOU.txt"
Get-ADComputer -Filter * -SearchBase "OU=[pc'sOU],DC=domain,DC=org" -Properties *  | Select -Property Name | Sort-Object -Property name, ws | Out-File $path 
[System.Windows.MessageBox]::Show("PC list exported to $path")
