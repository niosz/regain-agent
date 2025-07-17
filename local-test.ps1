$server = Read-Host -Prompt "Server Address"
$username1 = Read-Host -Prompt "Username 1"
$password1 = Read-Host -Prompt "Password 1" -AsSecureString | ConvertFrom-SecureString
$username2 = Read-Host -Prompt "Username 2"
$password2 = Read-Host -Prompt "Password 2" -AsSecureString | ConvertFrom-SecureString
$isBasic = "y"
try {
    Write-Information "Verifica di connessione utenza 1 : $username1"
    Connect-ManagementServer -Server $server -Credential ([pscredential]::new($username1, ($password1 | ConvertTo-SecureString))) -BasicUser:($isBasic -eq 'y')
    Disconnect-ManagementServer
    Write-Information "Verifica di connessione utenza 2 : $username2"
    Connect-ManagementServer -Server $server -Credential ([pscredential]::new($username2, ($password2 | ConvertTo-SecureString))) -BasicUser:($isBasic -eq 'y')
    Disconnect-ManagementServer
    Write-Information "Verifiche eseguete con successo."
}
catch {
    Write-Error "Errore durante la verifica delle credenziali: $_"
    exit 1
}