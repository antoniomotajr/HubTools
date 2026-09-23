param(
    [string]$MsixPath = "",
    [string]$CertificatePath = ""
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path

function Test-IsAdministrator {
    $Identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $Principal = New-Object Security.Principal.WindowsPrincipal($Identity)
    return $Principal.IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator
    )
}

function Quote-Argument([string]$Value) {
    return '"' + $Value.Replace('"', '\"') + '"'
}

if (-not $CertificatePath) {
    $CertificatePath = Join-Path $Root "TechToolHub.cer"

    if (-not (Test-Path $CertificatePath)) {
        $CertificatePath = Join-Path $Root "release\TechToolHub.cer"
    }
}

if (-not $MsixPath) {
    $SearchFolders = @(
        $Root,
        (Join-Path $Root "release")
    )

    $Latest = $null

    foreach ($Folder in $SearchFolders) {
        if (-not (Test-Path $Folder)) {
            continue
        }

        $Candidate = Get-ChildItem `
            $Folder `
            -Filter "TechToolHub_*_x64.msix" `
            -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTime -Descending |
            Select-Object -First 1

        if ($Candidate) {
            $Latest = $Candidate
            break
        }
    }

    if (-not $Latest) {
        throw "Nenhum TechToolHub_*_x64.msix foi encontrado."
    }

    $MsixPath = $Latest.FullName
}

$MsixPath = (Resolve-Path $MsixPath).Path
$CertificatePath = (Resolve-Path $CertificatePath).Path

if (-not (Test-IsAdministrator)) {
    Write-Host "Solicitando permissão de Administrador..." -ForegroundColor Yellow

    $Arguments = @(
        "-NoProfile",
        "-ExecutionPolicy", "Bypass",
        "-File", (Quote-Argument $PSCommandPath),
        "-MsixPath", (Quote-Argument $MsixPath),
        "-CertificatePath", (Quote-Argument $CertificatePath)
    ) -join " "

    $Process = Start-Process `
        -FilePath "powershell.exe" `
        -Verb RunAs `
        -ArgumentList $Arguments `
        -Wait `
        -PassThru

    exit $Process.ExitCode
}

Write-Host ""
Write-Host "==> Confiando no certificado MSIX" -ForegroundColor Cyan
Write-Host "Certificado: $CertificatePath" -ForegroundColor DarkGray

$Cert = New-Object `
    System.Security.Cryptography.X509Certificates.X509Certificate2(
        $CertificatePath
    )

$TrustedMachine = Get-ChildItem "Cert:\LocalMachine\TrustedPeople" |
    Where-Object { $_.Thumbprint -eq $Cert.Thumbprint } |
    Select-Object -First 1

if (-not $TrustedMachine) {
    Import-Certificate `
        -FilePath $CertificatePath `
        -CertStoreLocation "Cert:\LocalMachine\TrustedPeople" |
        Out-Null

    Write-Host "Certificado adicionado a LocalMachine\TrustedPeople." -ForegroundColor Green
}
else {
    Write-Host "Certificado já é confiável no Computador Local." -ForegroundColor Green
}

$Verified = Get-ChildItem "Cert:\LocalMachine\TrustedPeople" |
    Where-Object { $_.Thumbprint -eq $Cert.Thumbprint } |
    Select-Object -First 1

if (-not $Verified) {
    throw "O certificado não foi encontrado em LocalMachine\TrustedPeople após a importação."
}

Write-Host ""
Write-Host "==> Verificando assinatura" -ForegroundColor Cyan

$Signature = Get-AuthenticodeSignature -FilePath $MsixPath

Write-Host "Status assinatura: $($Signature.Status)" -ForegroundColor DarkGray
Write-Host "Thumbprint pacote: $($Signature.SignerCertificate.Thumbprint)" -ForegroundColor DarkGray
Write-Host "Thumbprint confiável: $($Cert.Thumbprint)" -ForegroundColor DarkGray

if (
    $Signature.SignerCertificate -and
    $Signature.SignerCertificate.Thumbprint -ne $Cert.Thumbprint
) {
    throw "O certificado fornecido não corresponde ao certificado usado para assinar o MSIX."
}

Write-Host ""
Write-Host "==> Instalando TECH TOOL HUB" -ForegroundColor Cyan
Write-Host "MSIX: $MsixPath" -ForegroundColor DarkGray

$Existing = Get-AppxPackage `
    -Name "TechToolHub.LocalWorkspace" `
    -ErrorAction SilentlyContinue

if ($Existing) {
    Write-Host "Versão atualmente instalada: $($Existing.Version)" -ForegroundColor DarkGray
}

Add-AppxPackage -Path $MsixPath

Write-Host ""
Write-Host "TECH TOOL HUB instalado com sucesso." -ForegroundColor Green
Write-Host "Abra pelo Menu Iniciar: TECH TOOL HUB" -ForegroundColor Green
