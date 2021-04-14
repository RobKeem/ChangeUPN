Import-Module ActiveDirectory
$DomainController = 'ho-pth-dc09.dhw.wa.gov.au'
$OldUPN = Read-Host "Enter the current UPN for the user"
$NewUPN = Read-Host "Enter the new UPN for the user"
$ExchangeServer = 'ho-azu-exh01.dhw.wa.gov.au'
$SAM = Get-ADUser -Filter {UserPrincipalName -eq $OldUPN}| Select-Object -ExpandProperty SamAccountName

#Checks for duplicate UPNs
If (Get-ADUser -Filter "UserPrincipalName -eq '$NewUPN'") {
    Write-Host "$NewUPN already exists"
    Break
}

Write-Host "Changing UPN for $OldUPN" -ForegroundColor Yellow
Set-ADUser -Identity $SAM -UserPrincipalName $NewUPN

#Connect to Hybrid Exchange
Try {
    $Creds = Get-Credential -Message "Enter DHW Admin Credentials"
    $HASession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$ExchangeServer/PowerShell/ -Authentication Kerberos -Credential $Creds
    Import-PSSession $HASession -DisableNameChecking -AllowClobber
} Catch {
    Write-Host "Unable to connect to Exchange Server: $ExchangeServer." -ForegroundColor Red
    Break
}

$Mailbox = Get-RemoteMailbox $OldUPN -DomainController $DomainController
$Mailbox | Set-RemoteMailbox -EmailAddressPolicyEnabled $False -DomainController $DomainController
$Mailbox.EmailAddresses | Where-Object {$_ -like "sip:*"} | ForEach-Object {
                $Mailbox | Set-RemoteMailbox -EmailAddresses @{remove=$_} -DomainController $DomainController
            }
$Mailbox | Set-RemoteMailbox -EmailAddresses @{add="sip:$NewUPN"} -DomainController $DomainController
$Mailbox | Set-RemoteMailbox -EmailAddressPolicyEnabled $True -DomainController $DomainController
    
#Queries Hybrid to ensure the UPN has been updated
If (Get-RemoteMailbox -Identity $NewUPN) {
    Write-Host "UPN Updated to: $NewUPN" -ForegroundColor Green
} Else {
    Write-Host "$SAM has failed to update to $NewUPN, check the account and try again." -ForegroundColor Red
}
