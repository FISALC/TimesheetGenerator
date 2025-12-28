$cert = New-SelfSignedCertificate -Type Custom -Subject "CN=TimesheetGenerator" -KeyUsage DigitalSignature -FriendlyName "TimesheetGeneratorWrapper" -CertStoreLocation "Cert:\CurrentUser\My" -TextExtension @("2.5.29.37={text}1.3.6.1.5.5.7.3.3", "2.5.29.19={text}")
$password = ConvertTo-SecureString -String "123456" -Force -AsPlainText
Export-PfxCertificate -Cert $cert -FilePath "TimesheetGenerator.Maui\TimesheetGenerator_SignKey.pfx" -Password $password
Export-Certificate -Cert $cert -FilePath "TimesheetGenerator.Maui\TimesheetGenerator_SignKey.cer"
Write-Host "Certificate Created Successfully"
