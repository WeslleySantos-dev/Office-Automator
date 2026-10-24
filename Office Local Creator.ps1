# Pasta onde o script está
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$sourceIconFolder = Join-Path $scriptPath "icons"

# Verifica se a pasta icons existe junto com o script
if (-not (Test-Path $sourceIconFolder)) {
    Write-Host "ERRO: Pasta 'icons' não encontrada junto com o script!" -ForegroundColor Red
    Write-Host "Coloque a pasta 'icons' na mesma pasta deste script." -ForegroundColor Yellow
    pause
    exit 1
}

# Pasta de destino dos ícones (dentro do perfil do Chrome)
$chromeUserData = Join-Path $env:LOCALAPPDATA "Google\Chrome\User Data"
$destIconFolder = Join-Path $chromeUserData "icons"

# Cria a pasta de destino se não existir
if (-not (Test-Path $destIconFolder)) {
    New-Item -ItemType Directory -Path $destIconFolder -Force | Out-Null
    Write-Host "Pasta criada: $destIconFolder" -ForegroundColor Cyan
}

# Lista das aplicações
$apps = @(
    @{
        Nome     = "Word Online"
        Url      = "https://word.cloud.microsoft"
        IconFile = "Microsoft Word.ico"
    },
    @{
        Nome     = "Excel Online"
        Url      = "https://excel.cloud.microsoft"
        IconFile = "Microsoft Excel.ico"
    },
    @{
        Nome     = "PowerPoint Online"
        Url      = "https://powerpoint.cloud.microsoft"
        IconFile = "Microsoft PowerPoint.ico"
    }
)

# Copia os ícones para a pasta do Chrome
Write-Host "`nCopiando ícones..." -ForegroundColor Cyan
foreach ($app in $apps) {
    $sourceIcon = Join-Path $sourceIconFolder $app.IconFile
    $destIcon = Join-Path $destIconFolder $app.IconFile

    if (Test-Path $sourceIcon) {
        Copy-Item -Path $sourceIcon -Destination $destIcon -Force
        Write-Host "  Copiado: $($app.IconFile)" -ForegroundColor Green
    }
    else {
        Write-Host "  Ícone não encontrado: $($app.IconFile)" -ForegroundColor Yellow
    }
}

# Localiza o Chrome
$chromePaths = @(
    "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
    "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
    "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe"
)

$chrome = $chromePaths | Where-Object { Test-Path $_ } | Select-Object -First 1

if (-not $chrome) {
    Write-Host "ERRO: Google Chrome não encontrado!" -ForegroundColor Red
    pause
    exit 1
}

Write-Host "`nChrome encontrado: $chrome`n" -ForegroundColor Cyan

# Função para criar atalho
function New-ChromeApp {
    param (
        [string]$Nome,
        [string]$Url,
        [string]$ChromePath,
        [string]$IconPath
    )

    $desktop = [Environment]::GetFolderPath("Desktop")
    $startMenu = Join-Path ([Environment]::GetFolderPath("StartMenu")) "Programs"

    # Lista dos locais onde o atalho será criado
    $locais = @(
        (Join-Path $desktop "$Nome.lnk"),
        (Join-Path $startMenu "$Nome.lnk")
    )

    foreach ($atalho in $locais) {
        $shell = New-Object -ComObject WScript.Shell
        $shortcut = $shell.CreateShortcut($atalho)

        $shortcut.TargetPath = $ChromePath
        $shortcut.Arguments = "--app=`"$Url`""
        $shortcut.WorkingDirectory = Split-Path $ChromePath
        $shortcut.Description = "Microsoft $Nome"
        $shortcut.WindowStyle = 1

        if (Test-Path $IconPath) {
            $shortcut.IconLocation = $IconPath
        }

        $shortcut.Save()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($shell) | Out-Null
    }

    Write-Host "Criado: $Nome (Área de Trabalho + Menu Iniciar)" -ForegroundColor Green
}
# Cria os atalhos usando os ícones da pasta do Chrome
foreach ($app in $apps) {
    $iconPath = Join-Path $destIconFolder $app.IconFile
    New-ChromeApp -Nome $app.Nome -Url $app.Url -ChromePath $chrome -IconPath $iconPath
}
Write-Host "`nPronto!" -ForegroundColor Yellow
Write-Host "Atalhos criados na Área de Trabalho."
Write-Host "Ícones salvos em: $destIconFolder" -ForegroundColor Gray