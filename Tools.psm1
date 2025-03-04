#Add [] to the string to output multiple computers via the script

function Get-CorpCompSYSInfo
{
    [cmdletbinding()]

    Param([String[]]$ComputerName)

    ForEach($Computer in $ComputerName)
    {
        $CompSYS = get-ciminstance -classname win32_ComputerSystem -ComputerName $Computer
        $BIOS = get-ciminstance -classname win32_BIOS -computername $Computer
        $DiskInfo = get-ciminstance -classname win32_LogicalDisk -computername $Computer
        $Properties = [ordered]@{
            'ComputerName' = $Computer;
            'BioSerial'    = $BIOS.SerialNumber;
            'Manufacturer' = $CompSYS.Manufacturer;
            'Model'        = $CompSYS.Model;
            'DriveLetter'  = $DiskInfo.DeviceID
                       }

        $OutputObject = New-Object -TypeName psobject -Property $Properties
        Write-Output $OutputObject
    }
}

#Use to test the script(Delete Hashtag):
#get-CorpCompSYSInfo -ComputerName LON-DC1,LON-SVR1 | Format-List