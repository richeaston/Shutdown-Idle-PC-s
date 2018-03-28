$path = "C:\\AllComputersInOU.txt"
Get-ADComputer -Filter * -SearchBase "OU=Blocks,OU=Whickham_Workstations,DC=whickhamschool,DC=org" -Properties *  | Select -Property Name | Sort-Object -Property name, ws | Out-File $path 
[System.Windows.MessageBox]::Show("PC list exported to $path")