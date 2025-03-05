#Check for users with "London" city and move them to "London" OU and "London Users" group

$OUName = "London" 
$DomainDN = "DC=Adatum,DC=com" 
$OUPath = "OU=$OUName,$DomainDN" 
$GroupName = “London Users”

ForEach ($User in (Get-ADUser -Filter {City -eq "London"})) 
{ 
    Move-ADObject -Identity $User -TargetPath $OUPath 
    Add-ADGroupMember -Identity $GroupName -Members $User 
}
