#Check for users with "London" city and move them to "London" OU and "London Users" group

$OUName = "London" 
$DomainDN = "DC=Adatum,DC=com" 
$OUPath = "OU=$OUName,$DomainDN" 
$GroupName = “London Users”
$LondonUser = @()

ForEach ($LondonUser in (Get-ADUser -Filter {City -eq "London"})) 
{ 
    Move-ADObject -Identity $LondonUser -TargetPath $OUPath 
    Add-ADGroupMember -Identity $GroupName -Members $LondonUser 
}
