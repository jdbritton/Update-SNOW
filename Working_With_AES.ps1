# This is just a bunch of snippets for working with AES encryption and using it to
# encrypt passwords/generating keys.

#GENERATE KEY
$aeskeypath = ".\aeskey.key"
$AESKey = New-Object Byte[] 32
[Security.Cryptography.RNGCryptoServiceProvider]::Create().GetBytes($AESKey)
Set-Content $aeskeypath $AESKey 

#GIVE IT A PASSWORD TO ENCRYPT: 
$pw = Read-Host "type in a password!"-AsSecureString
$pw
$key = Get-Content .\aeskey.key
$encryptpw = $pw | ConvertFrom-SecureString -Key $key
#copy content to file 
Set-Content .\cred.txt $encryptpw

# RECOVERING PLAINTEXT:
$key = (Get-Content ".\aeskey.key")
$password = Get-Content ".\cred.txt" | ConvertTo-SecureString -Key $key
$temp = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($password)
$PlainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($temp)
$PlainPassword

$password = Write-Output "=" | ConvertTo-SecureString -Key $Key
