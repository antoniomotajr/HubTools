param(
    [string]$PackageVersion = "2.26.2.0",
    [string]$Publisher = "CN=TechToolHub",
    [string]$Python = "python",
    [switch]$Install,
    [switch]$ResumePackage
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$Packaging = Join-Path $Root "packaging"
$BuildRoot = Join-Path $Root "build-msix"
$PyWork = Join-Path $BuildRoot "pyinstaller-work"
$DistRoot = Join-Path $BuildRoot "dist"
$Layout = Join-Path $BuildRoot "layout"
$Release = Join-Path $Root "release"
$AppName = "TechToolHub"
$IdentityName = "TechToolHub.LocalWorkspace"

function Write-Step([string]$Message) {
    Write-Host ""
    Write-Host "==> $Message" -ForegroundColor Cyan
}

function Find-WindowsSdkTool([string]$ToolName) {
    $KitsRoot = "${env:ProgramFiles(x86)}\Windows Kits\10\bin"

    if (-not (Test-Path $KitsRoot)) {
        throw "Windows SDK não encontrado. Instale o Windows 10/11 SDK."
    }

    $Candidates = Get-ChildItem $KitsRoot -Directory -ErrorAction SilentlyContinue |
        Sort-Object Name -Descending

    foreach ($Candidate in $Candidates) {
        $ToolPath = Join-Path $Candidate.FullName "x64\$ToolName"
        if (Test-Path $ToolPath) {
            return $ToolPath
        }
    }

    throw "$ToolName não encontrado no Windows SDK."
}

function Ensure-Pillow {
    & $Python -c "from PIL import Image" *> $null
    if ($LASTEXITCODE -eq 0) {
        return
    }

    Write-Host "Pillow não encontrado. Instalando para gerar os ícones..." -ForegroundColor Yellow
    & $Python -m pip install pillow

    if ($LASTEXITCODE -ne 0) {
        throw "Falha ao instalar Pillow."
    }
}

function Ensure-PackagingFiles {
    Write-Step "Verificando estrutura de packaging"

    $Brand = Join-Path $Root "brand_logo.png"
    if (-not (Test-Path $Brand)) {
        throw "Arquivo obrigatório não encontrado: $Brand"
    }

    $Assets = Join-Path $Packaging "Assets"
    New-Item -ItemType Directory -Force -Path $Packaging, $Assets | Out-Null

    $ManifestPath = Join-Path $Packaging "AppxManifest.xml"

    if (-not (Test-Path $ManifestPath)) {
        Write-Host "AppxManifest.xml ausente. Criando automaticamente..." -ForegroundColor Yellow

        $ManifestTemplate = @'
<?xml version="1.0" encoding="utf-8"?>
<Package
  xmlns="http://schemas.microsoft.com/appx/manifest/foundation/windows10"
  xmlns:uap="http://schemas.microsoft.com/appx/manifest/uap/windows10"
  xmlns:rescap="http://schemas.microsoft.com/appx/manifest/foundation/windows10/restrictedcapabilities"
  IgnorableNamespaces="uap rescap">

  <Identity
    Name="TechToolHub.LocalWorkspace"
    Publisher="__PUBLISHER__"
    Version="__VERSION__"
    ProcessorArchitecture="x64" />

  <Properties>
    <DisplayName>TECH TOOL HUB</DisplayName>
    <PublisherDisplayName>TECH TOOL HUB</PublisherDisplayName>
    <Description>Local Workspace para ferramentas, links e projetos.</Description>
    <Logo>Assets\StoreLogo.png</Logo>
  </Properties>

  <Resources>
    <Resource Language="pt-BR" />
  </Resources>

  <Dependencies>
    <TargetDeviceFamily
      Name="Windows.Desktop"
      MinVersion="10.0.17763.0"
      MaxVersionTested="10.0.26100.0" />
  </Dependencies>

  <Applications>
    <Application
      Id="TechToolHub"
      Executable="TechToolHub\TechToolHub.exe"
      EntryPoint="Windows.FullTrustApplication">

      <uap:VisualElements
        DisplayName="TECH TOOL HUB"
        Description="LOCAL WORKSPACE"
        BackgroundColor="transparent"
        Square150x150Logo="Assets\Square150x150Logo.png"
        Square44x44Logo="Assets\Square44x44Logo.png">

        <uap:DefaultTile
          Square310x310Logo="Assets\Square310x310Logo.png"
          Wide310x150Logo="Assets\Wide310x150Logo.png" />
      </uap:VisualElements>
    </Application>
  </Applications>

  <Capabilities>
    <rescap:Capability Name="runFullTrust" />
  </Capabilities>

</Package>
'@

        [System.IO.File]::WriteAllText(
            $ManifestPath,
            $ManifestTemplate,
            (New-Object System.Text.UTF8Encoding($false))
        )
    }

    # Corrige automaticamente manifestos de versões anteriores.
    $ExistingManifest = Get-Content $ManifestPath -Raw

    if (
        $ExistingManifest -match 'Square310x310Logo="Assets\\Square310x310Logo.png"' -and
        $ExistingManifest -notmatch 'Wide310x150Logo='
    ) {
        Write-Host "Atualizando AppxManifest.xml: adicionando Wide310x150Logo..." -ForegroundColor Yellow

        $ExistingManifest = $ExistingManifest -replace `
            'Square310x310Logo="Assets\\Square310x310Logo.png"\s*/>', `
            "Square310x310Logo=`"Assets\Square310x310Logo.png`"`r`n          Wide310x150Logo=`"Assets\Wide310x150Logo.png`" />"

        [System.IO.File]::WriteAllText(
            $ManifestPath,
            $ExistingManifest,
            (New-Object System.Text.UTF8Encoding($false))
        )
    }

    $RequiredImages = @(
        (Join-Path $Packaging "TechToolHub.ico"),
        (Join-Path $Assets "Square44x44Logo.png"),
        (Join-Path $Assets "Square150x150Logo.png"),
        (Join-Path $Assets "Square310x310Logo.png"),
        (Join-Path $Assets "Wide310x150Logo.png"),
        (Join-Path $Assets "StoreLogo.png")
    )

    $MissingImages = @($RequiredImages | Where-Object { -not (Test-Path $_) })

    if ($MissingImages.Count -gt 0) {
        Write-Host "Ícones MSIX ausentes. Gerando a partir de brand_logo.png..." -ForegroundColor Yellow
        Ensure-Pillow

        $Generator = Join-Path $env:TEMP "tech_tool_hub_generate_msix_assets.py"

        $PythonCode = @'
from pathlib import Path
from PIL import Image, ImageOps
import sys

brand = Path(sys.argv[1])
packaging = Path(sys.argv[2])
assets = packaging / "Assets"
assets.mkdir(parents=True, exist_ok=True)

image = Image.open(brand).convert("RGBA")

def make_square(size, path, fill_ratio=0.82):
    canvas = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    target = max(1, int(size * fill_ratio))
    fitted = ImageOps.contain(image, (target, target), Image.Resampling.LANCZOS)
    x = (size - fitted.width) // 2
    y = (size - fitted.height) // 2
    canvas.alpha_composite(fitted, (x, y))
    canvas.save(path, "PNG")

def make_wide(width, height, path):
    canvas = Image.new("RGBA", (width, height), (0, 0, 0, 0))
    fitted = ImageOps.contain(
        image,
        (118, 118),
        Image.Resampling.LANCZOS,
    )
    x = (width - fitted.width) // 2
    y = (height - fitted.height) // 2
    canvas.alpha_composite(fitted, (x, y))
    canvas.save(path, "PNG")

make_square(44, assets / "Square44x44Logo.png", 0.78)
make_square(150, assets / "Square150x150Logo.png", 0.82)
make_square(310, assets / "Square310x310Logo.png", 0.82)
make_wide(310, 150, assets / "Wide310x150Logo.png")
make_square(50, assets / "StoreLogo.png", 0.82)

ico_canvas = Image.new("RGBA", (256, 256), (0, 0, 0, 0))
fitted = ImageOps.contain(image, (220, 220), Image.Resampling.LANCZOS)
ico_canvas.alpha_composite(
    fitted,
    ((256 - fitted.width) // 2, (256 - fitted.height) // 2)
)
ico_canvas.save(
    packaging / "TechToolHub.ico",
    format="ICO",
    sizes=[(16,16),(24,24),(32,32),(48,48),(64,64),(128,128),(256,256)]
)
'@

        [System.IO.File]::WriteAllText(
            $Generator,
            $PythonCode,
            (New-Object System.Text.UTF8Encoding($false))
        )

        try {
            & $Python $Generator $Brand $Packaging

            if ($LASTEXITCODE -ne 0) {
                throw "Falha ao gerar os ícones do MSIX."
            }
        }
        finally {
            Remove-Item $Generator -Force -ErrorAction SilentlyContinue
        }
    }

    Write-Host "Packaging pronto: $Packaging" -ForegroundColor Green
}

function Ensure-InstallScript {
    $Installer = Join-Path $Root "install-msix.ps1"

    if (Test-Path $Installer) {
        return
    }

    Write-Host "install-msix.ps1 ausente. Criando automaticamente..." -ForegroundColor Yellow

    $InstallerText = @'
param(
    [string]$MsixPath = "",
    [string]$CertificatePath = ""
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path

if (-not $CertificatePath) {
    $CertificatePath = Join-Path $Root "TechToolHub.cer"
}

if (-not $MsixPath) {
    $Latest = Get-ChildItem $Root -Filter "TechToolHub_*_x64.msix" |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1

    if (-not $Latest) {
        throw "Nenhum arquivo TechToolHub_*_x64.msix foi encontrado em $Root."
    }

    $MsixPath = $Latest.FullName
}

if (-not (Test-Path $CertificatePath)) {
    throw "Certificado não encontrado: $CertificatePath"
}

if (-not (Test-Path $MsixPath)) {
    throw "MSIX não encontrado: $MsixPath"
}

Write-Host "Confiando no certificado para o usuário atual..." -ForegroundColor Cyan

$Cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2(
    $CertificatePath
)

$Trusted = Get-ChildItem "Cert:\CurrentUser\TrustedPeople" |
    Where-Object { $_.Thumbprint -eq $Cert.Thumbprint } |
    Select-Object -First 1

if (-not $Trusted) {
    Import-Certificate `
        -FilePath $CertificatePath `
        -CertStoreLocation "Cert:\CurrentUser\TrustedPeople" | Out-Null
}

Write-Host "Instalando: $MsixPath" -ForegroundColor Cyan
Add-AppxPackage -Path $MsixPath

Write-Host ""
Write-Host "TECH TOOL HUB instalado com sucesso." -ForegroundColor Green
Write-Host "Abra pelo Menu Iniciar: TECH TOOL HUB" -ForegroundColor Green
'@

    [System.IO.File]::WriteAllText(
        $Installer,
        $InstallerText,
        (New-Object System.Text.UTF8Encoding($false))
    )
}

function Ensure-DevCertificate([string]$Subject) {
    Write-Step "Preparando certificado de desenvolvimento"

    $Cert = Get-ChildItem "Cert:\CurrentUser\My" |
        Where-Object {
            $_.Subject -eq $Subject -and
            $_.HasPrivateKey -and
            $_.NotAfter -gt (Get-Date).AddDays(30)
        } |
        Sort-Object NotAfter -Descending |
        Select-Object -First 1

    if (-not $Cert) {
        $Cert = New-SelfSignedCertificate `
            -Type Custom `
            -Subject $Subject `
            -FriendlyName "TECH TOOL HUB MSIX Development" `
            -KeyAlgorithm RSA `
            -KeyLength 2048 `
            -HashAlgorithm SHA256 `
            -KeyUsage DigitalSignature `
            -CertStoreLocation "Cert:\CurrentUser\My" `
            -NotAfter (Get-Date).AddYears(5) `
            -TextExtension @("2.5.29.37={text}1.3.6.1.5.5.7.3.3")
    }

    New-Item -ItemType Directory -Force -Path $Release | Out-Null

    $CerPath = Join-Path $Release "TechToolHub.cer"
    Export-Certificate -Cert $Cert -FilePath $CerPath -Force | Out-Null

    Write-Host "Certificado: $($Cert.Subject)" -ForegroundColor DarkGray
    Write-Host "Thumbprint:   $($Cert.Thumbprint)" -ForegroundColor DarkGray

    return $Cert
}

function Migrate-LegacyData {
    Write-Step "Migrando dados atuais para LOCALAPPDATA"

    $Target = Join-Path $env:LOCALAPPDATA "TechToolHub"
    New-Item -ItemType Directory -Force -Path $Target | Out-Null

    $Files = @(
        "apps_data.json",
        "links_data.json",
        "local_apps_data.json",
        "projects_data.json",
        "user_profile.json",
        ".apps_catalog_v1_done",
        ".official_icons_v1_done"
    )

    foreach ($Name in $Files) {
        $Source = Join-Path $Root $Name
        $Destination = Join-Path $Target $Name

        if ((Test-Path $Source) -and -not (Test-Path $Destination)) {
            Copy-Item $Source $Destination
            Write-Host "Migrado: $Name" -ForegroundColor DarkGray
        }
    }

    foreach ($FolderName in @("icons", "user_media")) {
        $SourceFolder = Join-Path $Root $FolderName
        $DestinationFolder = Join-Path $Target $FolderName

        if (Test-Path $SourceFolder) {
            New-Item -ItemType Directory -Force -Path $DestinationFolder | Out-Null
            Copy-Item `
                (Join-Path $SourceFolder "*") `
                $DestinationFolder `
                -Recurse `
                -Force `
                -ErrorAction SilentlyContinue
        }
    }

    Write-Host "Dados persistentes: $Target" -ForegroundColor Green
}

Write-Step "Validando arquivos principais"

foreach ($File in @(
    (Join-Path $Root "app.py"),
    (Join-Path $Root "brand_logo.png")
)) {
    if (-not (Test-Path $File)) {
        throw "Arquivo obrigatório não encontrado: $File"
    }
}

if ($PackageVersion -notmatch '^\d+\.\d+\.\d+\.\d+$') {
    throw "PackageVersion precisa ter 4 partes, por exemplo 2.22.1.0."
}

Ensure-PackagingFiles
Ensure-InstallScript

Write-Step "Validando arquivos de packaging"

$Required = @(
    (Join-Path $Packaging "AppxManifest.xml"),
    (Join-Path $Packaging "TechToolHub.ico"),
    (Join-Path $Packaging "Assets\Square44x44Logo.png"),
    (Join-Path $Packaging "Assets\Square150x150Logo.png"),
    (Join-Path $Packaging "Assets\Square310x310Logo.png"),
    (Join-Path $Packaging "Assets\StoreLogo.png")
)

foreach ($File in $Required) {
    if (-not (Test-Path $File)) {
        throw "Arquivo obrigatório não encontrado: $File"
    }
}

Migrate-LegacyData

Write-Step "Verificando PyInstaller e dependências de runtime"

$RequiredPyInstaller = "6.22.3"

$CurrentPyInstaller = & $Python -c "import PyInstaller; print(PyInstaller.__version__)" 2>$null

if (
    $LASTEXITCODE -ne 0 -or
    "$CurrentPyInstaller".Trim() -ne $RequiredPyInstaller
) {
    $OldVersion = if ($CurrentPyInstaller) {
        "$CurrentPyInstaller".Trim()
    }
    else {
        "não instalado"
    }

    Write-Host "Atualizando PyInstaller: $OldVersion -> $RequiredPyInstaller" -ForegroundColor Yellow

    & $Python -m pip install `
        --upgrade `
        "pyinstaller==$RequiredPyInstaller"

    if ($LASTEXITCODE -ne 0) {
        throw "Falha ao atualizar PyInstaller para $RequiredPyInstaller."
    }
}

Write-Host "PyInstaller: $RequiredPyInstaller" -ForegroundColor Green

# O primeiro teste pode falhar porque jaraco ainda não está instalado.
# Com ErrorActionPreference=Stop, stderr de python.exe viraria um erro fatal
# antes que o script pudesse verificar $LASTEXITCODE. Temporariamente
# usamos Continue e restauramos o comportamento original logo depois.
$PreviousErrorActionPreference = $ErrorActionPreference
$ErrorActionPreference = "Continue"

& $Python -c "import jaraco.text, jaraco.functools, jaraco.context" 2>$null
$JaracoProbeExitCode = $LASTEXITCODE

$ErrorActionPreference = $PreviousErrorActionPreference

if ($JaracoProbeExitCode -ne 0) {
    Write-Host "Instalando dependências jaraco para pkg_resources..." -ForegroundColor Yellow

    & $Python -m pip install `
        --upgrade `
        jaraco.text `
        jaraco.functools `
        jaraco.context

    if ($LASTEXITCODE -ne 0) {
        throw "Falha ao instalar dependências jaraco."
    }
}

# Verificação final, também protegida contra NativeCommandError do
# Windows PowerShell 5.1.
$PreviousErrorActionPreference = $ErrorActionPreference
$ErrorActionPreference = "Continue"

& $Python -c "import jaraco.text, jaraco.functools, jaraco.context; print('jaraco: OK')" 2>&1 |
    ForEach-Object { Write-Host $_ }

$JaracoVerifyExitCode = $LASTEXITCODE

$ErrorActionPreference = $PreviousErrorActionPreference

if ($JaracoVerifyExitCode -ne 0) {
    throw "As dependências jaraco ainda não podem ser importadas."
}

Write-Step "Localizando Windows SDK"

$MakeAppx = Find-WindowsSdkTool "makeappx.exe"
$SignTool = Find-WindowsSdkTool "signtool.exe"

Write-Host "MakeAppx: $MakeAppx" -ForegroundColor DarkGray
Write-Host "SignTool: $SignTool" -ForegroundColor DarkGray

$BuiltApp = Join-Path $DistRoot $AppName
$BuiltExe = Join-Path $BuiltApp "$AppName.exe"

if ($ResumePackage) {
    Write-Step "Retomando a partir do EXE já gerado"

    if (-not (Test-Path $BuiltExe)) {
        throw "Não foi possível retomar: $BuiltExe não existe. Execute novamente sem -ResumePackage."
    }

    # Mantém dist/TechToolHub gerado pelo PyInstaller e recria apenas o layout.
    if (Test-Path $Layout) {
        Remove-Item $Layout -Recurse -Force
    }

    New-Item `
        -ItemType Directory `
        -Force `
        -Path $Layout, $Release | Out-Null

    Write-Host "Reutilizando: $BuiltExe" -ForegroundColor Green
}
else {
    Write-Step "Limpando build anterior"

    if (Test-Path $BuildRoot) {
        Remove-Item $BuildRoot -Recurse -Force
    }

    New-Item `
        -ItemType Directory `
        -Force `
        -Path $BuildRoot, $PyWork, $DistRoot, $Layout, $Release | Out-Null

    Write-Step "Gerando TechToolHub.exe com PyInstaller - modo enxuto"

    $PyInstallerArgs = @(
    "-m", "PyInstaller",
    "--noconfirm",
    "--clean",
    "--onedir",
    "--noconsole",
    "--name", $AppName,
    "--icon", (Join-Path $Packaging "TechToolHub.ico"),
    "--add-data", "$(Join-Path $Root 'brand_logo.png');.",
    "--add-data", "$(Join-Path $Root 'README-MSIX.txt');.",

    # Uvicorn: apenas os módulos dinâmicos realmente necessários.
    "--hidden-import", "uvicorn.logging",
    "--hidden-import", "uvicorn.loops.auto",
    "--hidden-import", "uvicorn.protocols.http.auto",
    "--hidden-import", "uvicorn.protocols.websockets.auto",
    "--hidden-import", "uvicorn.lifespan.on",

    # Compatibilidade com pkg_resources/setuptools atuais.
    "--hidden-import", "jaraco.text",
    "--hidden-import", "jaraco.functools",
    "--hidden-import", "jaraco.context",
    "--collect-submodules", "jaraco",
    "--collect-data", "jaraco.text",

    # Exclui bibliotecas de desenvolvimento/ciência/UI que não pertencem
    # ao TECH TOOL HUB e estavam fazendo o PyInstaller analisar milhares
    # de arquivos do ambiente Python global.
    "--exclude-module", "IPython",
    "--exclude-module", "matplotlib",
    "--exclude-module", "numpy",
    "--exclude-module", "pandas",
    "--exclude-module", "scipy",
    "--exclude-module", "sklearn",
    "--exclude-module", "pytest",
    "--exclude-module", "sphinx",
    "--exclude-module", "docutils",
    "--exclude-module", "jedi",
    "--exclude-module", "parso",
    "--exclude-module", "nbformat",
    "--exclude-module", "notebook",
    "--exclude-module", "jupyter",
    "--exclude-module", "jupyter_core",
    "--exclude-module", "jupyter_client",
    "--exclude-module", "zmq",
    "--exclude-module", "pygame",
    "--exclude-module", "PyQt5",
    "--exclude-module", "PyQt6",
    "--exclude-module", "PySide2",
    "--exclude-module", "PySide6",
    "--exclude-module", "tkinter",
    "--exclude-module", "astroid",
    "--exclude-module", "pylint",
    "--exclude-module", "lxml",

    "--distpath", $DistRoot,
    "--workpath", $PyWork,
    "--specpath", $BuildRoot,
    (Join-Path $Root "app.py")
)

    & $Python @PyInstallerArgs

    if ($LASTEXITCODE -ne 0) {
        throw "PyInstaller falhou."
    }

    if (-not (Test-Path $BuiltExe)) {
        throw "TechToolHub.exe não foi gerado."
    }
}

Write-Step "Montando layout do MSIX"

$PackageAppDir = Join-Path $Layout $AppName
New-Item -ItemType Directory -Force -Path $PackageAppDir | Out-Null
Copy-Item (Join-Path $BuiltApp "*") $PackageAppDir -Recurse -Force

$AssetsDest = Join-Path $Layout "Assets"
New-Item -ItemType Directory -Force -Path $AssetsDest | Out-Null
Copy-Item (Join-Path $Packaging "Assets\*") $AssetsDest -Force

$ManifestTemplate = Get-Content `
    (Join-Path $Packaging "AppxManifest.xml") `
    -Raw

$Manifest = $ManifestTemplate.
    Replace("__VERSION__", $PackageVersion).
    Replace("__PUBLISHER__", $Publisher)

$ManifestPath = Join-Path $Layout "AppxManifest.xml"

[System.IO.File]::WriteAllText(
    $ManifestPath,
    $Manifest,
    (New-Object System.Text.UTF8Encoding($false))
)

Write-Step "Criando MSIX"

$MsixName = "TechToolHub_${PackageVersion}_x64.msix"
$MsixPath = Join-Path $Release $MsixName

if (Test-Path $MsixPath) {
    Remove-Item $MsixPath -Force
}

& $MakeAppx pack /o /d $Layout /p $MsixPath

if ($LASTEXITCODE -ne 0) {
    throw "MakeAppx falhou."
}

$Cert = Ensure-DevCertificate $Publisher

Write-Step "Assinando MSIX"

& $SignTool sign /fd SHA256 /sha1 $Cert.Thumbprint /s My $MsixPath

if ($LASTEXITCODE -ne 0) {
    throw "SignTool falhou."
}

Copy-Item `
    (Join-Path $Root "install-msix.ps1") `
    (Join-Path $Release "install-msix.ps1") `
    -Force

Write-Step "Build concluído"

Write-Host "MSIX:        $MsixPath" -ForegroundColor Green
Write-Host "Certificado: $(Join-Path $Release 'TechToolHub.cer')" -ForegroundColor Green
Write-Host "Dados:       $(Join-Path $env:LOCALAPPDATA 'TechToolHub')" -ForegroundColor Green

if ($Install) {
    Write-Step "Instalando TECH TOOL HUB"

    $Installer = Join-Path $Release "install-msix.ps1"

    if (-not (Test-Path $Installer)) {
        throw "Instalador não encontrado: $Installer"
    }

    Write-Host "A instalação solicitará permissão de Administrador." -ForegroundColor Yellow

    & powershell.exe `
        -NoProfile `
        -ExecutionPolicy Bypass `
        -File $Installer `
        -MsixPath $MsixPath `
        -CertificatePath (Join-Path $Release "TechToolHub.cer")

    if ($LASTEXITCODE -ne 0) {
        throw "A instalação do MSIX falhou com código $LASTEXITCODE."
    }
}
