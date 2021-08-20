param(
    [parameter(Mandatory=$true , Position = 0, HelpMessage = "Existing or a new Issuer (can be anything).", ValueFromPipeline=$false)]
    [string]$IssuerName,

    [parameter(Mandatory=$true , Position = 1, HelpMessage = "The certificate provider.", ValueFromPipeline=$false)]
    [string]$ProviderName,

    [parameter(Mandatory=$true , Position = 2, HelpMessage = "Name of an existing Key Vault that will store the new certificate.", ValueFromPipeline=$false)]
    [ValidatePattern({^[a-zA-Z][a-zA-Z0-9]+$})]
    [string]$KeyVaultName,

    [parameter(Mandatory=$true , Position = 3, HelpMessage = "The domain for the certificate.", ValueFromPipeline=$false)]
    [ValidatePattern({^[a-zA-Z][a-zA-Z0-9\.\-]+$})]
    [string]$SAN,

    [parameter(Mandatory=$true , Position = 4, HelpMessage = "Certificate name (can be anything ^[0-9a-zA-Z-]+$).", ValueFromPipeline=$false)]
    [ValidatePattern({^[0-9a-zA-Z\-]+$})]
    [string]$CertificateName,

    [parameter(Mandatory=$true , Position = 5, HelpMessage = "The subject of your cert, ideally your whitelisted domain. Must start with 'CN='. Example: CN=something.com", ValueFromPipeline=$false)]
    [ValidatePattern({^CN=.+})]
    [string]$CertificateSubject    
)

function downloadKeyVaultCertificate([string]$keyvaultName, [string]$certificateName, [string]$pfxFilePath)
{
    $secret      = Get-AzKeyVaultSecret -VaultName $keyvaultName -Name $certificateName;
    $secretByte  = [Convert]::FromBase64String($secret.SecretValueText);
    $certFlags   = [X509KeyStorageFlags]::Exportable + [X509KeyStorageFlags]::PersistKeySet;
    $x509Cert    = [X509Certificate2]::new($secretByte, $null, $certFlags);

    $pfxFileByte = $x509Cert.Export([X509ContentType]::Pfx);
    [System.IO.File]::WriteAllBytes($pfxFilePath, $pfxFileByte);

    Import-PfxCertificate -CertStoreLocation 'Cert:\CurrentUser\My' -FilePath $pfxFilePath -Exportable;

    $thumbprint = $x509Cert.Thumbprint;

    if (($thumbprint.GetType()).IsArray)
    {
        $thumbprint = $thumbprint[0];
    }
    if ($null -ne (get-member -InputObject $thumbprint -Name 'Thumbprint'))
    {
        $thumbprint = $thumbprint.Thumbprint;
    }
    return $thumbprint;

    return $thumbprint;
}

$ErrorActionPreference = 'Stop';

$context = Get-AzContext;
if ($null -eq $context)
{
    throw "Please log in to Azure and select a subscription, then try again.";
}

Write-Host "`n`nThis will create certificate $CertificateName under account $($context.Account.Id) and subscription $($context.Subscription.Name)" -ForegroundColor Green;
Write-Host "Is this what you want to do? <y/N>" -ForegroundColor Green;
$yn = Read-Host;
if ($yn -ne 'y')
{
    return;
}

Set-AzKeyVaultCertificateIssuer -VaultName $KeyVaultName -IssuerProvider $ProviderName -Name $IssuerName
$policy=New-AzKeyVaultCertificatePolicy -SubjectName $CertificateSubject -IssuerName $IssuerName -DnsNames $SAN
Add-AzKeyVaultCertificate -VaultName $KeyVaultName -Name $CertificateName -CertificatePolicy $policy

while ($(Get-AzKeyVaultCertificateOperation $KeyVaultName $CertificateName).Status.Equals("inProgress"))
{
    Start-Sleep -s 1
    Write-Host -NoNewLine "." -ForegroundColor Cyan
}

# Display error if any
if (!$(Get-AzKeyVaultCertificateOperation $KeyVaultName $CertificateName).Status.Equals("completed"))
{
    throw "Error Message: $($(Get-AzKeyVaultCertificateOperation $KeyVaultName $CertificateName).ErrorMessage)"
}

# if no error, cert has been created
$certificateFileName = ($CertificateName.Replace(' ','')) + '.pfx';
$certificateFilePath = Join-Path $env:TEMP $certificateFileName;
$thumbprint = downloadKeyVaultCertificate $KeyVaultName $CertificateName $certificateFilePath;
Write-Host "Done, certificate thumbprint is $thumbprint" -ForegroundColor Green;
return $thumbprint;
