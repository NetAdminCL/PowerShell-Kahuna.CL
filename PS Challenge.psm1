#Creates a global security group "London Users" if it doesn't exist, or says it already exists

$OUName = "London" 
$DomainDN = "DC=Adatum,DC=com" 
$OUPath = "OU=$OUName,$DomainDN" 
$GroupName = “London Users”

If (-not (Get-ADGroup -Filter {Name -eq $GroupName} -SearchBase $OUPath))
{ 
    New-ADGroup - Name $GroupName -GroupScope Global -GroupCategory Security -Path $OUPath
    Write-Output "Global Security Group '$GroupName' has been created in '$OUName'."
}
Else
{
    Write-Output "'$GroupName' already exists in 'OUPath'."
}