// Self-signed certificate generator for OfficeAI dev server
// Generates .certs/localhost.pfx for HTTPS on port 3000

import { execSync } from 'child_process';
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'fs';
import { join, dirname } from 'path';
import { fileURLToPath } from 'url';

const __dirname = dirname(fileURLToPath(import.meta.url));
const certsDir = join(__dirname, '..', '.certs');
const pfxPath = join(certsDir, 'localhost.pfx');

mkdirSync(certsDir, { recursive: true });

console.log('Generating self-signed certificate for localhost...');

try {
  // Windows: use PowerShell with .NET
  const psCmd = `
    $cert = New-SelfSignedCertificate -DnsName 'localhost' -CertStoreLocation 'Cert:\\CurrentUser\\My' -KeyExportPolicy Exportable -NotAfter (Get-Date).AddYears(2) -KeyAlgorithm RSA -KeyLength 2048;
    $pfxBytes = $cert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Pfx, 'officeai123');
    [System.IO.File]::WriteAllBytes('${pfxPath.replace(/\\/g, '\\\\')}', $pfxBytes);
    Write-Output 'OK'
  `.trim();

  const result = execSync(`powershell -NoProfile -Command "${psCmd.replace(/"/g, '\\"')}"`, {
    encoding: 'utf8',
    timeout: 30000,
  });

  if (result.includes('OK')) {
    console.log(`Certificate generated: ${pfxPath}`);
  } else {
    throw new Error('Unexpected PowerShell output: ' + result);
  }
} catch (err) {
  // Fallback: try openssl
  console.log('PowerShell cert generation failed, trying openssl...');
  try {
    execSync(
      `openssl req -x509 -newkey rsa:2048 -keyout "${certsDir}/localhost.key" -out "${certsDir}/localhost.crt" -days 730 -nodes -subj "/CN=localhost" -addext "subjectAltName=DNS:localhost"`,
      { timeout: 15000 }
    );
    console.log('Generated PEM certs via openssl. Creating PFX...');
    execSync(
      `openssl pkcs12 -export -out "${pfxPath}" -inkey "${certsDir}/localhost.key" -in "${certsDir}/localhost.crt" -passout pass:officeai123`,
      { timeout: 15000 }
    );
    console.log(`Certificate generated: ${pfxPath}`);
  } catch (opensslErr) {
    console.error('Both PowerShell and openssl failed.');
    console.error('Please manually place a .certs/localhost.pfx file with passphrase "officeai123".');
    console.error('Or install openssl and try again.');
    process.exit(1);
  }
}

// Verify the PFX was created
if (!existsSync(pfxPath)) {
  console.error('ERROR: PFX file was not created.');
  process.exit(1);
}

console.log('Done. The dev server will use HTTPS on port 3000.');
