Import-Csv 'C:\Office365Users.csv' | ForEach-Object {
$upn = $_."UserPrincipalName"
$newupn = $_."EmailAddress"
Write-Host "Changing UPN value from: "$upn" to: " $newupn -ForegroundColor Yellow
Set-AzureADUser -ObjectId $upn  -UserPrincipalName $newupn
}