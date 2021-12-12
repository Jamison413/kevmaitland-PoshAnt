Write-Host "Creating CertificateRequest(CSR) for $CertName `r "

Invoke-Command -ComputerName testbox -ScriptBlock {

$CertName = "newcert.contoso.com"
$CSRPath = "$env:TEMP\$($CertName)_.csr"
$INFPath = "$env:TEMP\$($CertName)_.inf"
$Signature = '$Windows NT$' 


$INF =
@"
[Version]
Signature= "$Signature" 

[NewRequest]
Subject = "CN=$CertName, OU=Contoso East Division, O=Contoso Inc, L=Boston, S=Massachusetts, C=US"
KeySpec = 1
KeyLength = 2048
Exportable = TRUE
MachineKeySet = TRUE
SMIME = False
PrivateKeyArchive = FALSE
UserProtected = FALSE
UseExistingKeySet = FALSE
ProviderName = "Microsoft RSA SChannel Cryptographic Provider"
ProviderType = 12
RequestType = PKCS10
KeyUsage = 0xa0

[EnhancedKeyUsageExtension]

OID=1.3.6.1.5.5.7.3.1 
"@

write-Host "Certificate Request is being generated `r "
$INF | out-file -filepath $INFPath -force
certreq -new $INFPath $CSRPath

}
write-output "Certificate Request has been generated"