Enable-PSRemoting -Force
Set-Item wsman:\localhost\client\trustedhosts *
Restart-Service WinRM
Test-WsMan 10.11.48.13
Invoke-Command -ComputerName 10.11.48.13 -ScriptBlock { Get-ChildItem C:\ } -credential a-lflojo
Enter-PSSession -ComputerName 10.11.48.13 -Credential a-lflojo