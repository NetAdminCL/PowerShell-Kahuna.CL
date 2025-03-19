#Script That queries all servers on the network
$Servers = Get-ADComputer -Filter * | Select-Object -ExpandProperty Name

# Query Active Directory for IT department administrators
$AdminEmails = Get-ADUser -Filter { Department -eq "IT" -and Title -like "*Admin*" } | Select-Object -ExpandProperty EmailAddress -ErrorAction SilentlyContinue
$AdminEmails = $AdminEmails -join ","

#Creates and saves the error log file
$Date = Get-Date -Format "yyyyMMdd"
$ErrorLog = "\\Server\Share\Server Logs\ErrorLog_$Date.csv"
if (-Not (Test-Path $ErrorLog)) {
    New-Item -Path $ErrorLog -ItemType File -Force
    Add-Content -Path $ErrorLog -Value "Time,Server,Error"
}

# Query the network for the email server and specify the email for notifications
$SMTPServer = (Get-ADComputer -Filter { Name -like "*mail*" } | Select-Object -ExpandProperty Name | Select-Object -First 1) + "@adatum.com"
$FromEmail = christian.lasdulce@adatum.com

#Loops through the servers and checks if they are up or down
#If the server is down, it logs the error and sends an email to the administrators

while ($true) {
    foreach ($Server in $Servers) {
        try {
            if (Test-Connection -ComputerName $Server -Count 1 -Quiet) {
                Write-Host "$Server is up" -ForegroundColor Green
            } else {
                Write-Host "$Server is down" -ForegroundColor Red
                $ErrorMessage = "Server is down"
                $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                Add-Content -Path $ErrorLog -Value "$TimeStamp,$Server,$ErrorMessage"
                Send-MailMessage -From $FromEmail -To $AdminEmails -Subject "Server Down Alert: $Server" -Body "$TimeStamp - $Server is down. Error Message: $ErrorMessage" -SmtpServer $SMTPServer
            }
        } catch {
            $ErrorMessage = $_.Exception.Message
            $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            Add-Content -Path $ErrorLog -Value "$TimeStamp,$Server,$ErrorMessage"
            Send-MailMessage -From $FromEmail -To $AdminEmails -Subject "Server Error Alert: $Server" -Body "$TimeStamp - $Server encountered an error: $ErrorMessage" -SmtpServer $SMTPServer
        }
    }
    Start-Sleep -Seconds 1800
}