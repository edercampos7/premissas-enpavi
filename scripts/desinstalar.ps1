# Desliga a atualizacao automatica do site "Encontro de Contas Enpavi".
# Uso: botao direito > "Executar com o PowerShell", ou
#      powershell -ExecutionPolicy Bypass -File scripts\desinstalar.ps1
# A pasta do projeto (Documentos\premissas-enpavi) e o site continuam como estao.

$NomeTarefa = 'Atualizar site Premissas Enpavi'
try {
    Unregister-ScheduledTask -TaskName $NomeTarefa -Confirm:$false -ErrorAction Stop
    Write-Host ("Tarefa agendada '" + $NomeTarefa + "' removida.") -ForegroundColor Green
} catch { Write-Host 'Nenhuma tarefa agendada encontrada.' }

$lnk = Join-Path ([Environment]::GetFolderPath('Startup')) 'Vigia Premissas Enpavi.lnk'
if (Test-Path $lnk) { Remove-Item -Force $lnk; Write-Host 'Atalho da pasta Inicializar removido.' -ForegroundColor Green }

$n = 0
Get-CimInstance Win32_Process -ErrorAction SilentlyContinue |
    Where-Object { $_.CommandLine -like '*vigiar_planilha.ps1*' } |
    ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue; $n++ }
Write-Host ('Vigias encerrados: ' + $n)
Write-Host 'Pronto. O site para de se atualizar sozinho (o atualizar_agora.bat continua funcionando).'
