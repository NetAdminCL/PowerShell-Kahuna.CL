#Parameter for the "Error Log" Excel documentation
param 
(
    [string]$ErrorLogPath = "C:\PowerShell Automatic Script\ServerCheck"
)

#Checks the Active Directory for any Server
$Servers = Get-ADComputer -Filter {OperatingSystem -like "*Server*" -and Enabled -eq $true} -Property Name | Select-Object -ExpandProperty Name

#Checks to see if the "Error Log" document for the day exists and if it doesn not, creates one
$Date = Get-Date -Format "yyyyMMdd"
$ErrorLog = Join-Path -Path $ErrorLogPath -ChildPath "ErrorLog_$Date.xlsx"

if (-Not (Test-Path $ErrorLog)) 
{
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
function Log-ErrorToExcel
{
    param 
    (
        [string]$Date,
        [string]$Time,
        [string]$Errors
    )

    $Excel = New-Object -ComObject Excel.Application
    $Workbook = $Excel.Workbooks.Open($ErrorLog)
    $Worksheet = $Workbook.Worksheets.Item(1)
    $LastRow = $Worksheet.UsedRange.Rows.Count + 1
    $Worksheet.Cells.Item($LastRow, 1) = $Date
    $Worksheet.Cells.Item($LastRow, 2) = $Time
    $Worksheet.Cells.Item($LastRow, 3) = $Errors
    $Workbook.Save()
    $Workbook.Close($false)
    $Excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)
}

#Creates a schedule task "ServerCheck" if it does not exist set to run every 5 minutes
$Task = Get-ScheduledTask -TaskName "ServerCheck" -ErrorAction SilentlyContinue

if ($Task.TaskName -eq "ServerCheck")
{
    write-output "Task already exists."
}
else
{
    $PSSCriptRoot = $env:USERPROFILE
    $ScriptPath = "C:\PowerShell Automatic Script\ServerCheck\ServerCheckv1_20Mar2025.ps1"
    $Action = New-ScheduledTaskAction -Execute 'PowerShell.exe' -Argument "-File $ScriptPath" -WorkingDirectory $env:USERPROFILE
    $trigger = New-ScheduledTaskTrigger -Once -At (Get-Date) -RepetitionInterval (New-TimeSpan -Minutes 5) 
    $trigger.Repetition.Duration = "PT5M"
    $trigger.ExecutionTimeLimit = 'PT5M'
    $Principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
    $Settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -DontStopOnIdleEnd -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1) -ExecutionTimeLimit 5 -Compatibility Win8
    Register-ScheduledTask -TaskName "ServerCheck" -Action $Action -Trigger $Trigger -Principal $Principal -Settings $Settings -ErrorAction SilentlyContinue

}

# Loops through the servers and checks if they are up or down, then emails the IT department if a server is down
foreach ($Server in $Servers) 
{
    try 
    {
        if (Test-Connection -ComputerName $Server -Count 1 -Quiet) 
        {
            Write-Host "$Server is up" -ForegroundColor Green
        } 
        else 
        {
            Write-Host "$Server is down" -ForegroundColor Red
            $ErrorMessage = "Server is down"
            $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            Log-ErrorToExcel -Date (Get-Date -Format "yyyy-MM-dd") -Time (Get-Date -Format "HH:mm:ss") -Error "$Server,$ErrorMessage"
        }
    } 
    catch 
    {
        $ErrorMessage = $_.Exception.Message
        $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Log-ErrorToExcel -Date (Get-Date -Format "yyyy-MM-dd") -Time (Get-Date -Format "HH:mm:ss") -Error "$Server,$ErrorMessage"
    }
}