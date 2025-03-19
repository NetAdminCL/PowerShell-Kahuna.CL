param (
    [string]$ErrorLogPath = "\\Server\Share\Server Logs"
)

#Script That queries all servers on the network
$Servers = Get-ADComputer -Filter { OperatingSystem -like "*Server*" -and Enabled -eq $true } -Property Name | Select-Object -ExpandProperty Name

# Query Active Directory for IT department administrators 
try {
    $AdminEmails = Get-ADUser -Filter { Department -eq "IT" -and Title -like "*Admin*" } | Select-Object -ExpandProperty EmailAddress
    $AdminEmails = $AdminEmails -join ","
} catch {
    $ErrorMessage = $_.Exception.Message
$ErrorLog = Join-Path -Path $ErrorLogPath -ChildPath "ErrorLog_$Date.csv"
    Add-Content -Path $ErrorLog -Value "$TimeStamp,Get-ADUser,$ErrorMessage"
    $AdminEmails = ""
}
$AdminEmails = $AdminEmails -join ","

#Creates and saves the error log file
$Date = Get-Date -Format "yyyyMMdd"
$ErrorLog = "\\Server\Share\Server Logs\ErrorLog_$Date.csv"
if (-Not (Test-Path $ErrorLog)) {
    New-Item -Path $ErrorLog -ItemType File -Force
$SMTPServers = Get-ADComputer -Filter { Name -like "*mail*" } | Select-Object -ExpandProperty Name
if ($SMTPServers.Count -gt 0) {
    $SMTPServer = $SMTPServers[0] + "@adatum.com"
} else {
    throw "No mail servers found."
}
}

# Query the network for the email server and specify the email for notifications
$SMTPServer = (Get-ADComputer -Filter { Name -like "*mail*" } | Select-Object -ExpandProperty Name | Select-Object -First 1) + "@adatum.com"

# Schedule the script to run every 30 minutes using Task Scheduler
$Action = New-ScheduledTaskAction -Execute 'PowerShell.exe' -Argument "-File `"$PSScriptRoot\PowerShellProject_18Mar2025.psm1`""
$Trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(30) -RepetitionInterval (New-TimeSpan -Minutes 30) -RepetitionDuration ([TimeSpan]::MaxValue)
$Principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
$Settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -DontStopOnIdleEnd -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1)
Register-ScheduledTask -TaskName "ServerCheck" -Action $Action -Trigger $Trigger -Principal $Principal -Settings $Settings

#Loops through the servers and checks if they are up or down
#If the server is down, it logs the error and sends an email to the administrators
foreach ($Server in $Servers) {
    try {
        if (Test-Connection -ComputerName $Server -Count 1 -Quiet) {
            Write-Host "$Server is up" -ForegroundColor Green
        } else {
            Write-Host "$Server is down" -ForegroundColor Red
            $ErrorMessage = "Server is down"
            $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            Add-Content -Path $ErrorLog -Value "$TimeStamp,$Server,$ErrorMessage"
            try {
                Send-MailMessage -From $FromEmail -To $AdminEmails -Subject "Server Down Alert: $Server" -Body "$TimeStamp - $Server is down. Error Message: $ErrorMessage" -SmtpServer $SMTPServer
            } catch {
                $EmailErrorMessage = $_.Exception.Message
                Add-Content -Path $ErrorLog -Value "$TimeStamp,$Server,EmailError,$EmailErrorMessage"
            }
        }
    } catch {
        $ErrorMessage = $_.Exception.Message
        $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Add-Content -Path $ErrorLog -Value "$TimeStamp,$Server,$ErrorMessage"
        try {
            Send-MailMessage -From $FromEmail -To $AdminEmails -Subject "Server Error Alert: $Server" -Body "$TimeStamp - $Server encountered an error: $ErrorMessage" -SmtpServer $SMTPServer
        } catch {
            $EmailErrorMessage = $_.Exception.Message
            Add-Content -Path $ErrorLog -Value "$TimeStamp,$Server,EmailError,$EmailErrorMessage"
        }
    }
}               Add-Content -Path $ErrorLog -Value "$TimeStamp,$Server,EmailError,$EmailErrorMessage"
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