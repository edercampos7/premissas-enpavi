<#
Vigia a planilha do encontro de contas Triade x Enpavi. Sempre que ela for salva
no Excel, regenera o index.html e publica no GitHub (commit + push). O GitHub
Pages atualiza o site cerca de 1 a 2 minutos depois.

Normalmente e iniciado sozinho ao entrar no Windows (ver instalar.ps1).
Uso manual:
    powershell -ExecutionPolicy Bypass -File scripts\vigiar_planilha.ps1                  (fica vigiando)
    powershell -ExecutionPolicy Bypass -File scripts\vigiar_planilha.ps1 -UmaVez -Forcar  (publica agora e sai)

Qual planilha: a mais recente cujo nome termina em "ENCONTRO ENPAVI.xlsx" na pasta
"Custo Enpavi" da Area de Trabalho (ou a informada em -Pasta / -Xlsx).
Registro do que aconteceu: scripts\log.txt
#>
param(
    [string]$Pasta,
    [string]$Xlsx,
    [int]$Intervalo = 20,
    [int]$Estavel = 30,
    [int]$ConferirGitHubMin = 15,
    [switch]$UmaVez,
    [switch]$Forcar,
    [switch]$SemAviso
)
$ErrorActionPreference = 'Stop'
$Repo = Split-Path -Parent $PSScriptRoot
$LogPath = Join-Path $PSScriptRoot 'log.txt'
$EstadoPath = Join-Path $PSScriptRoot 'estado.txt'
$Painel = Join-Path $PSScriptRoot 'atualizar_painel.ps1'
$utf8 = New-Object System.Text.UTF8Encoding($false)

if (-not $UmaVez) {
    # em segundo plano nao ha ninguem para digitar senha: falha e registra no log
    $env:GIT_TERMINAL_PROMPT = '0'
    $env:GCM_INTERACTIVE = 'Never'
}

function Log([string]$msg) {
    if ($UmaVez) { Write-Host $msg }
    try {
        if ((Test-Path $LogPath) -and ((Get-Item $LogPath).Length -gt 1MB)) { Move-Item -Force $LogPath ($LogPath + '.old') }
        [System.IO.File]::AppendAllText($LogPath, (Get-Date).ToString('yyyy-MM-dd HH:mm:ss') + '  ' + $msg + "`r`n", $utf8)
    } catch { }
}

# 04/11/2026: site desativado a pedido do Eder. O vigia sai sem fazer nada.
# Para tirar a tarefa agendada do Windows: botao direito em scripts\desinstalar.ps1 > Executar com o PowerShell.
Log 'Site desativado em 04/11/2026: o vigia nao faz mais nada.'
exit 0

function Aviso([string]$titulo, [string]$texto, [bool]$erro) {
    if ($SemAviso -or ($env:OS -ne 'Windows_NT')) { return }
    try {
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing
        $n = New-Object System.Windows.Forms.NotifyIcon
        $icones = 'System.Drawing.SystemIcons' -as [type]
        $n.Icon = $icones::Information
        $n.BalloonTipTitle = $titulo
        $n.BalloonTipText = $texto
        if ($erro) { $n.BalloonTipIcon = 'Error' } else { $n.BalloonTipIcon = 'Info' }
        $n.Visible = $true
        $n.ShowBalloonTip(8000)
        Start-Sleep -Seconds 8
        $n.Dispose()
    } catch { }
}

$GitExe = $null
$cmd = Get-Command git -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
if ($cmd) { $GitExe = $cmd.Source }
if (-not $GitExe) {
    foreach ($c in @("$env:ProgramFiles\Git\cmd\git.exe", "${env:ProgramFiles(x86)}\Git\cmd\git.exe", "$env:LOCALAPPDATA\Programs\Git\cmd\git.exe")) {
        if ($c -and (Test-Path $c)) { $GitExe = $c; break }
    }
}

function Invoke-Git {
    $old = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        $out = & $GitExe -C $Repo @args 2>&1 | ForEach-Object { "$_" }
        $code = $LASTEXITCODE
    } finally { $ErrorActionPreference = $old }
    [pscustomobject]@{ Code = $code; Out = (($out | Where-Object { $_ }) -join ' | ') }
}

function Get-Pastas {
    $l = @()
    if ($Pasta) { $l += $Pasta }
    try { $l += (Join-Path ([Environment]::GetFolderPath('Desktop')) 'Custo Enpavi') } catch { }
    if ($env:USERPROFILE) { $l += (Join-Path $env:USERPROFILE 'Desktop\Custo Enpavi') }
    if ($env:OneDrive) { $l += (Join-Path $env:OneDrive 'Desktop\Custo Enpavi') }
    return ($l | Select-Object -Unique)
}

function Find-Planilha {
    if ($Xlsx) {
        if (Test-Path -LiteralPath $Xlsx) { return (Get-Item -LiteralPath $Xlsx) }
        return $null
    }
    foreach ($p in (Get-Pastas)) {
        if (-not (Test-Path -LiteralPath $p)) { continue }
        $f = Get-ChildItem -LiteralPath $p -Filter '*.xlsx' -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -like '*ENCONTRO ENPAVI.xlsx' -and $_.Name -notlike '~$*' } |
            Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if ($f) { return $f }
    }
    return $null
}

function Get-Hash([string]$path) {
    # FileShare ReadWrite: funciona com a planilha aberta no Excel
    $fs = New-Object System.IO.FileStream($path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]'ReadWrite, Delete')
    try {
        $sha = [System.Security.Cryptography.SHA256]::Create()
        return ([System.BitConverter]::ToString($sha.ComputeHash($fs))).Replace('-', '')
    } finally { $fs.Dispose() }
}

function Read-Estado {
    $e = @{}
    if (Test-Path $EstadoPath) {
        foreach ($l in [System.IO.File]::ReadAllLines($EstadoPath)) {
            $i = $l.IndexOf('=')
            if ($i -gt 0) { $e[$l.Substring(0, $i)] = $l.Substring($i + 1) }
        }
    }
    return $e
}

function Save-Estado([string]$hash, [string]$resultado, [string]$arquivo) {
    $linhas = @(('hash=' + $hash), ('resultado=' + $resultado), ('arquivo=' + $arquivo), ('quando=' + (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')))
    try { [System.IO.File]::WriteAllLines($EstadoPath, [string[]]$linhas, $utf8) } catch { }
}

# Gera o index.html com a planilha e envia ao GitHub.
# Retorna: ok | sem-mudanca | erro-planilha | erro-git
$script:UltimoResumo = ''
function Publicar($arq) {
    $mtx = New-Object System.Threading.Mutex($false, 'Local\PremissasEnpaviPublicar')
    $tem = $false
    try { $tem = $mtx.WaitOne([TimeSpan]::FromMinutes(5)) } catch [System.Threading.AbandonedMutexException] { $tem = $true }
    if (-not $tem) { Log 'Outra atualizacao ja esta em andamento; tento de novo depois.'; $mtx.Dispose(); return 'erro-git' }
    try {
        # o index.html e sempre gerado de novo: descarta a versao local antes de baixar o que ha no GitHub
        Invoke-Git checkout '--' index.html | Out-Null
        $r = Invoke-Git pull --ff-only -q
        if ($r.Code -ne 0) {
            $r = Invoke-Git pull --rebase --autostash -q
            if ($r.Code -ne 0) {
                Invoke-Git rebase --abort | Out-Null
                Log ('Falha ao baixar do GitHub: ' + $r.Out)
                return 'erro-git'
            }
        }
        try {
            $saida = & $Painel -Xlsx $arq.FullName -ProjectDir $Repo *>&1 | ForEach-Object { "$_" }
            $script:UltimoResumo = ($saida | Where-Object { $_ -like 'Corte:*' } | Select-Object -Last 1)
            Log ('Gerado com ' + $arq.Name + ' -> ' + $script:UltimoResumo)
        } catch {
            Log ('Erro ao ler a planilha ' + $arq.Name + ': ' + $_.Exception.Message)
            Invoke-Git checkout '--' index.html | Out-Null
            $script:UltimoResumo = $_.Exception.Message
            return 'erro-planilha'
        }
        Invoke-Git add index.html | Out-Null
        $d = Invoke-Git diff --cached --quiet
        if ($d.Code -eq 0) { Log 'O site ja estava com esses valores; nada a enviar.'; return 'sem-mudanca' }
        $msg = 'Atualiza site (planilha salva em ' + $arq.LastWriteTime.ToString('dd/MM/yyyy HH:mm') + ')'
        $c = Invoke-Git commit -q -m $msg
        if ($c.Code -ne 0) { Log ('Falha no commit: ' + $c.Out); Invoke-Git reset -q | Out-Null; return 'erro-git' }
        $p = Invoke-Git push -q origin HEAD:main
        if ($p.Code -ne 0) {
            Log ('Falha ao enviar para o GitHub: ' + $p.Out)
            # desfaz o commit local para nao divergir; sera refeito na proxima tentativa
            Invoke-Git reset -q --soft HEAD~1 | Out-Null
            Invoke-Git reset -q | Out-Null
            return 'erro-git'
        }
        Log ('Publicado: ' + $msg)
        return 'ok'
    } finally {
        $mtx.ReleaseMutex()
        $mtx.Dispose()
    }
}

function Avisar-Resultado([string]$res) {
    if ($res -eq 'ok') { Aviso 'Site Enpavi atualizado' ('Publicado. O site muda em 1 a 2 minutos. ' + $script:UltimoResumo) $false }
    elseif ($res -eq 'erro-planilha') { Aviso 'Site Enpavi NAO atualizado' ('Problema na planilha: ' + $script:UltimoResumo) $true }
    elseif ($res -eq 'erro-git') { Aviso 'Site Enpavi NAO atualizado' 'Falha ao enviar ao GitHub. Nova tentativa em 5 minutos. Veja scripts\log.txt.' $true }
}

if (-not $GitExe) { Log 'Git nao encontrado. Instale o Git (https://git-scm.com/download/win) e rode o instalar.ps1.'; exit 1 }

# ---------------------------------------------------------------- publicar uma vez (atualizar_agora.bat)
if ($UmaVez) {
    $arq = Find-Planilha
    if (-not $arq) { Log ('Planilha "... ENCONTRO ENPAVI.xlsx" nao encontrada em: ' + ((Get-Pastas) -join '; ')); exit 1 }
    Log ('Planilha: ' + $arq.FullName)
    $h = Get-Hash $arq.FullName
    $est = Read-Estado
    if ((-not $Forcar) -and ($est['hash'] -eq $h)) { Log 'A planilha nao mudou desde a ultima publicacao.'; exit 0 }
    $res = Publicar $arq
    if ($res -ne 'erro-git') { Save-Estado $h $res $arq.FullName }
    Avisar-Resultado $res
    if ($res -eq 'ok') { Log 'Pronto. O site atualiza em 1 a 2 minutos.' }
    if ($res -like 'erro*') { exit 1 }
    exit 0
}

# ---------------------------------------------------------------- modo vigia (segundo plano)
$unico = New-Object System.Threading.Mutex($false, 'Local\PremissasEnpaviVigia')
$dono = $false
try { $dono = $unico.WaitOne(0) } catch [System.Threading.AbandonedMutexException] { $dono = $true }
if (-not $dono) { exit 0 }   # ja existe um vigia rodando

Log ('Vigia iniciado. Procurando "... ENCONTRO ENPAVI.xlsx" em: ' + ((Get-Pastas) -join '; '))
$ultimaAssinatura = $null
$mudouEm = Get-Date '2000-01-01'      # na partida confere logo se a planilha mudou com o PC desligado
$proxTentativa = [DateTime]::MinValue
$proxConferencia = (Get-Date).AddMinutes($ConferirGitHubMin)
$avisouSemPlanilha = $false
$avisouGit = $null

while ($true) {
    try {
        $arq = Find-Planilha
        if (-not $arq) {
            if (-not $avisouSemPlanilha) { Log 'Planilha nao encontrada; continuo procurando.'; $avisouSemPlanilha = $true }
        } else {
            $avisouSemPlanilha = $false
            $assin = $arq.FullName + '|' + $arq.LastWriteTimeUtc.Ticks + '|' + $arq.Length
            if ($assin -ne $ultimaAssinatura) {
                if ($ultimaAssinatura) { Log ('Planilha salva: ' + $arq.Name + ' (' + $arq.LastWriteTime.ToString('dd/MM/yyyy HH:mm:ss') + ')'); $mudouEm = Get-Date }
                $ultimaAssinatura = $assin
            }
            $agora = Get-Date
            if ($mudouEm -and (($agora - $mudouEm).TotalSeconds -ge $Estavel) -and ($agora -ge $proxTentativa)) {
                $h = Get-Hash $arq.FullName
                $est = Read-Estado
                if ($est['hash'] -ne $h) {
                    $res = Publicar $arq
                    if ($res -eq 'erro-git') {
                        $proxTentativa = (Get-Date).AddMinutes(5)
                        if ($avisouGit -ne $h) { Avisar-Resultado $res; $avisouGit = $h }
                    } else {
                        Save-Estado $h $res $arq.FullName
                        $mudouEm = $null
                        Avisar-Resultado $res
                    }
                } else { $mudouEm = $null }
            }
        }
        # de tempos em tempos: se o modelo do site mudou no GitHub, gera de novo com a planilha atual
        if ((Get-Date) -ge $proxConferencia) {
            $proxConferencia = (Get-Date).AddMinutes($ConferirGitHubMin)
            $f = Invoke-Git fetch -q origin
            if ($f.Code -eq 0 -and $arq) {
                $local = (Invoke-Git rev-parse HEAD).Out.Trim()
                $remoto = (Invoke-Git rev-parse origin/main).Out.Trim()
                if ($local -and $remoto -and ($local -ne $remoto)) {
                    Log 'Ha mudancas no GitHub; gerando o site de novo.'
                    $res = Publicar $arq
                    if ($res -eq 'erro-git') { $proxConferencia = (Get-Date).AddMinutes(5) }
                }
            }
        }
    } catch {
        Log ('Erro: ' + $_.Exception.Message)
    }
    Start-Sleep -Seconds $Intervalo
}
