param (
    [string]$ErrorLogPath = "C:\Users\Administrator.ADATUM\Desktop\Error Logs"
)

# Queries all servers on the network
$Servers = Get-ADComputer -Filter { OperatingSystem -like "*Server*" -and Enabled -eq $true } -Property Name | Select-Object -ExpandProperty Name

# Queries Active Directory for IT department administrators 
$Date = Get-Date -Format "yyyyMMdd"
try {
    $AdminEmails = Get-ADUser -Filter { Department -eq "IT" -and Title -like "*Admin*" } | Select-Object -ExpandProperty EmailAddress
    $AdminEmails = $AdminEmails -join ","
} catch {
    $ErrorMessage = $_.Exception.Message
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Log-ErrorToExcel -Date (Get-Date -Format "yyyy-MM-dd") -Time (Get-Date -Format "HH:mm:ss") -Error "Get-ADUser,$ErrorMessage"
    $AdminEmails = ""
}

# Creates and saves the error log file
$Date = Get-Date -Format "yyyyMMdd"
$ErrorLog = Join-Path -Path $ErrorLogPath -ChildPath "ErrorLog_$Date.xlsx"

# Check if the Excel file exists, if not create it with headers
if (-Not (Test-Path $ErrorLog)) {
    $Excel = New-Object -ComObject Excel.Application
    $Workbook = $Excel.Workbooks.Add()
    $Worksheet = $Workbook.Worksheets.Item(1)
    $Worksheet.Cells.Item(1, 1) = "Date"
    $Worksheet.Cells.Item(1, 2) = "Time"
    $Worksheet.Cells.Item(1, 3) = "Error"
    $Workbook.SaveAs($ErrorLog)
    $Workbook.Close($false)
    $Excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)
}

# Function to log errors to the Excel file
function Log-ErrorToExcel {
    param (
        [string]$Date,
        [string]$Time,
        [string]$Error
    )
    $Excel = New-Object -ComObject Excel.Application
    $Workbook = $Excel.Workbooks.Open($ErrorLog)
    $Worksheet = $Workbook.Worksheets.Item(1)
    $LastRow = $Worksheet.UsedRange.Rows.Count + 1
    $Worksheet.Cells.Item($LastRow, 1) = $Date
    $Worksheet.Cells.Item($LastRow, 2) = $Time
    $Worksheet.Cells.Item($LastRow, 3) = $Error
    $Workbook.Save()
    $Workbook.Close($false)
    $Excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)
}

# Query the network for the email server and specify the email for notifications
$SMTPServer = (Get-ADComputer -Filter { Name -like "*mail*" } | Select-Object -ExpandProperty Name | Select-Object -First 1) + "@adatum.com"
$Action = New-ScheduledTaskAction -Execute 'PowerShell.exe' -Argument "-File `"$PSScriptRoot\PowerShellProject_18Mar2025.psm1`""
$Trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(30) -RepetitionInterval (New-TimeSpan -Minutes 30) -RepetitionDuration ([TimeSpan]::MaxValue)
$Principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
$Settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -DontStopOnIdleEnd -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1)
Register-ScheduledTask -TaskName "ServerCheck" -Action $Action -Trigger $Trigger -Principal $Principal -Settings $Settings

# Loops through the servers and checks if they are up or down, then emails the IT department if a server is down
foreach ($Server in $Servers) {
    try {
        if (Test-Connection -ComputerName $Server -Count 1 -Quiet) {
            Write-Host "$Server is up" -ForegroundColor Green
        } else {
            Write-Host "$Server is down" -ForegroundColor Red
            $ErrorMessage = "Server is down"
            $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            Log-ErrorToExcel -Date (Get-Date -Format "yyyy-MM-dd") -Time (Get-Date -Format "HH:mm:ss") -Error "$Server,$ErrorMessage"
            try {
                Send-MailMessage -From $FromEmail -To $AdminEmails -Subject "Server Down Alert: $Server" -Body "$TimeStamp - $Server is down. Error Message: $ErrorMessage" -SmtpServer $SMTPServer
            } catch {
                $EmailErrorMessage = $_.Exception.Message
                Log-ErrorToExcel -Date (Get-Date -Format "yyyy-MM-dd") -Time (Get-Date -Format "HH:mm:ss") -Error "$Server,EmailError,$EmailErrorMessage"
            }
        }
    } catch {
        $ErrorMessage = $_.Exception.Message
        $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Log-ErrorToExcel -Date (Get-Date -Format "yyyy-MM-dd") -Time (Get-Date -Format "HH:mm:ss") -Error "$Server,$ErrorMessage"
        try {
            Send-MailMessage -From $FromEmail -To $AdminEmails -Subject "Server Error Alert: $Server" -Body "$TimeStamp - $Server encountered an error: $ErrorMessage" -SmtpServer $SMTPServer
        } catch {
            $EmailErrorMessage = $_.Exception.Message
            Log-ErrorToExcel -Date (Get-Date -Format "yyyy-MM-dd") -Time (Get-Date -Format "HH:mm:ss") -Error "$Server,EmailError,$EmailErrorMessage"
        }
    }
    Start-Sleep -Seconds 1800
}