$ErrorActionPreference = "Stop"

Write-Host "TECH TOOL HUB - Git Security Check" -ForegroundColor Cyan
Write-Host "==================================" -ForegroundColor Cyan

git status

Write-Host ""
Write-Host "Arquivos rastreados/candidatos sensíveis:" -ForegroundColor Yellow

$Patterns = @(
    '\.env($|\.)',
    '\.pfx$',
    '\.p12$',
    '\.pem$',
    '\.key$',
    '\.msix$',
    '\.exe$',
    '^release/',
    '^build/',
    '^dist/',
    '^build-msix/',
    '^build-dashboard/',
    'apps_data\.json$',
    'links_data\.json$',
    'projects_data\.json$',
    'local_apps_data\.json$',
    'user_profile\.json$',
    '^user_media/',
    '^icons/'
)

$Candidates = @(
    git ls-files --cached --others --exclude-standard
)

$Problems = @()

foreach ($File in $Candidates) {
    foreach ($Pattern in $Patterns) {
        if ($File -match $Pattern) {
            $Problems += $File
            break
        }
    }
}

$Problems = @($Problems | Sort-Object -Unique)

if ($Problems.Count -eq 0) {
    Write-Host "OK: nenhum arquivo proibido encontrado." -ForegroundColor Green
}
else {
    Write-Host "ATENÇÃO:" -ForegroundColor Red
    $Problems | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
}
