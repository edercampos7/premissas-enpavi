# Instala a atualizacao automatica do site "Encontro de Contas Enpavi".
#
# Rode UMA vez no computador onde a planilha e salva. No PowerShell, cole:
#
#     irm https://raw.githubusercontent.com/edercampos7/premissas-enpavi/main/scripts/instalar.ps1 | iex
#
# O que ele faz:
#   1. confere se o Git esta instalado (se nao estiver, tenta instalar pelo winget);
#   2. baixa o projeto para Documentos\premissas-enpavi;
#   3. confere o acesso a sua conta do GitHub (na primeira vez abre uma janela para entrar);
#   4. publica o site uma vez com a planilha atual;
#   5. programa o "vigia" para iniciar junto com o Windows: sempre que a planilha
#      "... ENCONTRO ENPAVI.xlsx" da pasta Custo Enpavi for salva, o site e atualizado.
# Pode rodar de novo sem problema (atualiza a instalacao).

function Instalar-PremissasEnpavi {
    $ErrorActionPreference = 'Continue'
    $RepoUrl = 'https://github.com/edercampos7/premissas-enpavi.git'
    $SiteUrl = 'https://edercampos7.github.io/premissas-enpavi/'
    $Destino = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'premissas-enpavi'
    $NomeTarefa = 'Atualizar site Premissas Enpavi'
    $PsExe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'

    function Passo([string]$t) { Write-Host ''; Write-Host ('>> ' + $t) -ForegroundColor Cyan }
    function Ok([string]$t) { Write-Host ('   OK  ' + $t) -ForegroundColor Green }
    function Atencao([string]$t) { Write-Host ('   !!  ' + $t) -ForegroundColor Yellow }

    Write-Host ''
    Write-Host '=== Site Encontro de Contas Enpavi - atualizacao automatica ===' -ForegroundColor White

    # 1. Git
    Passo 'Verificando o Git'
    $git = Get-Command git -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $git) {
        foreach ($c in @("$env:ProgramFiles\Git\cmd", "$env:LOCALAPPDATA\Programs\Git\cmd")) {
            if (Test-Path (Join-Path $c 'git.exe')) { $env:Path = $env:Path + ';' + $c }
        }
        $git = Get-Command git -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    }
    if (-not $git -and (Get-Command winget -ErrorAction SilentlyContinue)) {
        Atencao 'Git nao encontrado. Instalando pelo winget (o Windows pode pedir confirmacao)...'
        winget install --id Git.Git -e --source winget --accept-package-agreements --accept-source-agreements
        $env:Path = [Environment]::GetEnvironmentVariable('Path', 'Machine') + ';' + [Environment]::GetEnvironmentVariable('Path', 'User')
        foreach ($c in @("$env:ProgramFiles\Git\cmd", "$env:LOCALAPPDATA\Programs\Git\cmd")) {
            if (Test-Path (Join-Path $c 'git.exe')) { $env:Path = $env:Path + ';' + $c }
        }
        $git = Get-Command git -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    }
    if (-not $git) {
        Atencao 'Nao consegui instalar o Git sozinho. Vou abrir a pagina de download:'
        Atencao 'instale com as opcoes padrao, feche esta janela e rode o instalador de novo.'
        Start-Process 'https://git-scm.com/download/win'
        return
    }
    Ok (& git --version)

    # 2. projeto
    Passo ('Baixando o projeto para ' + $Destino)
    if (Test-Path (Join-Path $Destino '.git')) {
        & git -C $Destino checkout -- index.html 2>$null
        & git -C $Destino pull --ff-only
        if ($LASTEXITCODE -ne 0) { Atencao 'Nao consegui atualizar a pasta do projeto (veja a mensagem acima).'; return }
        Ok 'Projeto ja existia; atualizado'
    } else {
        & git clone $RepoUrl $Destino
        if ($LASTEXITCODE -ne 0) { Atencao 'Nao consegui baixar o projeto do GitHub (veja a mensagem acima).'; return }
        Ok 'Projeto baixado'
    }
    if (-not (& git -C $Destino config user.name)) { & git -C $Destino config user.name 'Eder Azevedo' }
    if (-not (& git -C $Destino config user.email)) { & git -C $Destino config user.email 'edercampos7@hotmail.com' }

    # 3. planilha
    Passo 'Procurando a planilha (... ENCONTRO ENPAVI.xlsx na pasta Custo Enpavi da Area de Trabalho)'
    $pastas = @((Join-Path ([Environment]::GetFolderPath('Desktop')) 'Custo Enpavi'), (Join-Path $env:USERPROFILE 'Desktop\Custo Enpavi'))
    if ($env:OneDrive) { $pastas += (Join-Path $env:OneDrive 'Desktop\Custo Enpavi') }
    $arq = $null
    foreach ($p in ($pastas | Select-Object -Unique)) {
        if (-not (Test-Path -LiteralPath $p)) { continue }
        $arq = Get-ChildItem -LiteralPath $p -Filter '*.xlsx' -File | Where-Object { $_.Name -like '*ENCONTRO ENPAVI.xlsx' } |
            Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if ($arq) { break }
    }
    if ($arq) { Ok $arq.FullName } else { Atencao 'Nao achei a planilha agora. O vigia vai continuar procurando nessa pasta.' }

    # 4. acesso ao GitHub
    Passo 'Conferindo o acesso ao GitHub (se abrir uma janela, entre com a conta edercampos7)'
    & git -C $Destino push --dry-run origin HEAD:main
    if ($LASTEXITCODE -ne 0) {
        Atencao 'O GitHub nao liberou o envio. Entre com a conta edercampos7 quando a janela abrir e rode o instalador de novo.'
        return
    }
    Ok 'Acesso ao GitHub confirmado'

    # 5. primeira publicacao
    $vigia = Join-Path $Destino 'scripts\vigiar_planilha.ps1'
    if ($arq) {
        Passo 'Publicando o site com a planilha atual'
        & $PsExe -NoProfile -ExecutionPolicy Bypass -File $vigia -UmaVez -Forcar -SemAviso
    }

    # 6. vigia ao entrar no Windows
    Passo 'Programando o vigia para iniciar junto com o Windows'
    Get-CimInstance Win32_Process -ErrorAction SilentlyContinue |
        Where-Object { $_.CommandLine -like '*vigiar_planilha.ps1*' -and $_.CommandLine -notlike '*-UmaVez*' } |
        ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
    $argumentos = '-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "' + $vigia + '"'
    $usuario = $env:USERDOMAIN + '\' + $env:USERNAME
    $ok = $false
    try {
        $acao = New-ScheduledTaskAction -Execute $PsExe -Argument $argumentos -WorkingDirectory $Destino
        $gatilho = New-ScheduledTaskTrigger -AtLogOn -User $usuario
        $config = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable `
            -ExecutionTimeLimit ([TimeSpan]::Zero) -MultipleInstances IgnoreNew -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1)
        $principal = New-ScheduledTaskPrincipal -UserId $usuario -LogonType Interactive -RunLevel Limited
        Register-ScheduledTask -TaskName $NomeTarefa -Action $acao -Trigger $gatilho -Settings $config -Principal $principal `
            -Description 'Atualiza o site Encontro de Contas Enpavi sempre que a planilha for salva.' -Force -ErrorAction Stop | Out-Null
        Start-ScheduledTask -TaskName $NomeTarefa -ErrorAction Stop
        $ok = $true
        Ok ("Tarefa agendada '" + $NomeTarefa + "' criada e iniciada")
    } catch {
        Atencao ('Nao consegui criar a tarefa agendada (' + $_.Exception.Message + '). Vou usar a pasta Inicializar do Windows.')
    }
    if (-not $ok) {
        try {
            $lnk = Join-Path ([Environment]::GetFolderPath('Startup')) 'Vigia Premissas Enpavi.lnk'
            $sh = New-Object -ComObject WScript.Shell
            $s = $sh.CreateShortcut($lnk)
            $s.TargetPath = $PsExe; $s.Arguments = $argumentos; $s.WorkingDirectory = $Destino; $s.WindowStyle = 7
            $s.Save()
            Start-Process -FilePath $PsExe -ArgumentList $argumentos -WindowStyle Hidden
            Ok 'Atalho criado na pasta Inicializar e vigia iniciado'
        } catch {
            Atencao ('Tambem nao consegui criar o atalho: ' + $_.Exception.Message)
            return
        }
    }

    Write-Host ''
    Write-Host '=== Pronto ===' -ForegroundColor Green
    Write-Host ('Site: ' + $SiteUrl)
    Write-Host 'A partir de agora, ao salvar a planilha no Excel, o site se atualiza sozinho em cerca de 2 minutos.'
    Write-Host ('Para atualizar na hora: ' + (Join-Path $Destino 'atualizar_agora.bat'))
    Write-Host ('Registro do que aconteceu: ' + (Join-Path $Destino 'scripts\log.txt'))
    Write-Host ''
}
Instalar-PremissasEnpavi
