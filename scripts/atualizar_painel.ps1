<#
Regenera a apresentação (index.html) a partir da planilha do encontro de contas
Tríade x Enpavi (abas PREMISSAS ENPAVI, ENCONTRO ENPAVI, RECEITAS ENPAVI-USINAS SP-ESO
e os caixas das obras). Uso:

    powershell -ExecutionPolicy Bypass -File scripts\atualizar_painel.ps1
    powershell -ExecutionPolicy Bypass -File scripts\atualizar_painel.ps1 -Xlsx "C:\caminho\planilha.xlsx"

Sem -Xlsx, usa a planilha padrão da pasta Custo Enpavi (ou o .xlsx mais recente da pasta do projeto).
A planilha precisa ter sido salva pelo Excel (o script lê os valores calculados que o Excel grava).
Funciona no Windows PowerShell 5.1 e no PowerShell 7.
#>
param(
    [string]$Xlsx,
    [string]$ProjectDir = (Split-Path -Parent $PSScriptRoot)
)
$ErrorActionPreference = 'Stop'
$utf8 = New-Object System.Text.UTF8Encoding($false)
$inv = [System.Globalization.CultureInfo]::InvariantCulture

if (-not $Xlsx) {
    # planilha mais recente cujo nome termina em "ENCONTRO ENPAVI.xlsx" na pasta Custo Enpavi da Área de Trabalho
    # (assim o mês no começo do nome pode mudar: 082026, 092026...)
    $pastas = @()
    try { $pastas += (Join-Path ([Environment]::GetFolderPath('Desktop')) 'Custo Enpavi') } catch { }
    if ($env:USERPROFILE) { $pastas += (Join-Path $env:USERPROFILE 'Desktop\Custo Enpavi') }
    $pastas += $ProjectDir
    foreach ($p in $pastas) {
        if (-not $p -or -not (Test-Path $p)) { continue }
        $f = Get-ChildItem -Path $p -Filter '*.xlsx' -File | Where-Object { $_.Name -like '*ENCONTRO ENPAVI.xlsx' -and $_.Name -notlike '~$*' } |
            Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if ($f) { $Xlsx = $f.FullName; break }
    }
    if (-not $Xlsx) { throw "Nenhuma planilha '... ENCONTRO ENPAVI.xlsx' encontrada. Informe -Xlsx." }
}
Write-Host "Planilha: $Xlsx"
$xlsxInfo = Get-Item $Xlsx

# ---------------------------------------------------------------- leitura do .xlsx (sem Excel)
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
# FileShare.ReadWrite: o Excel costuma manter o arquivo aberto enquanto é editado
$stream = New-Object System.IO.FileStream($Xlsx, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
$zip = New-Object System.IO.Compression.ZipArchive($stream, [System.IO.Compression.ZipArchiveMode]::Read)
function Read-Entry([string]$name) {
    $e = $zip.GetEntry($name)
    if (-not $e) { return $null }
    $sr = New-Object System.IO.StreamReader($e.Open(), $utf8)
    $t = $sr.ReadToEnd(); $sr.Dispose(); return $t
}
function Decode([string]$s) { if ($null -eq $s) { return $null }; return [System.Net.WebUtility]::HtmlDecode($s) }

# textos compartilhados
$shared = New-Object System.Collections.Generic.List[string]
$ssText = Read-Entry 'xl/sharedStrings.xml'
if ($ssText) {
    foreach ($m in [regex]::Matches($ssText, '<si>(.*?)</si>', 'Singleline')) {
        $parts = [regex]::Matches($m.Groups[1].Value, '<t(?:\s[^>]*)?>(.*?)</t>', 'Singleline') | ForEach-Object { $_.Groups[1].Value }
        $shared.Add((Decode ($parts -join '')))
    }
}

# nome da aba -> arquivo xml
$wbXml = Read-Entry 'xl/workbook.xml'
$relsXml = Read-Entry 'xl/_rels/workbook.xml.rels'
$rid2target = @{}
foreach ($m in [regex]::Matches($relsXml, '<Relationship [^>]*>')) {
    $id = [regex]::Match($m.Value, 'Id="([^"]+)"').Groups[1].Value
    $tg = [regex]::Match($m.Value, 'Target="([^"]+)"').Groups[1].Value
    $rid2target[$id] = $tg
}
$sheetFile = @{}
foreach ($m in [regex]::Matches($wbXml, '<sheet [^>]*>')) {
    $nm = Decode ([regex]::Match($m.Value, 'name="([^"]+)"').Groups[1].Value)
    $rid = [regex]::Match($m.Value, 'r:id="([^"]+)"').Groups[1].Value
    $tg = $rid2target[$rid]
    if ($tg.StartsWith('/')) { $tg = $tg.TrimStart('/') } else { $tg = 'xl/' + $tg }
    $sheetFile[$nm] = $tg
}
function Need-Sheet([string]$name) {
    if (-not $sheetFile.ContainsKey($name)) { throw "A planilha não tem a aba '$name'." }
    return $sheetFile[$name]
}

# lê as células de uma aba: coord -> @{ v = valor; f = fórmula; e = erro }
# $colRegex limita as colunas lidas (acelera as abas grandes dos caixas)
function Read-Sheet([string]$name, [string]$colRegex = '[A-Z]{1,3}') {
    $xml = Read-Entry (Need-Sheet $name)
    $cells = @{}
    $pat = '<c r="(' + $colRegex + ')(\d+)"([^>]*?)(?:/>|>(.*?)</c>)'
    foreach ($m in [regex]::Matches($xml, $pat, 'Singleline')) {
        $attrs = $m.Groups[3].Value; $inner = $m.Groups[4].Value
        if ($attrs.EndsWith('/')) { continue }
        $t = [regex]::Match($attrs, 't="([^"]+)"').Groups[1].Value
        $f = $null; $v = $null; $err = $false
        $fm = [regex]::Match($inner, '<f(?:\s[^>]*)?>(.*?)</f>', 'Singleline')
        if ($fm.Success) { $f = Decode $fm.Groups[1].Value }
        $vm = [regex]::Match($inner, '<v>(.*?)</v>', 'Singleline')
        if ($t -eq 's') { if ($vm.Success) { $v = $shared[[int]$vm.Groups[1].Value] } }
        elseif ($t -eq 'inlineStr') {
            $v = Decode ((([regex]::Matches($inner, '<t(?:\s[^>]*)?>(.*?)</t>', 'Singleline')) | ForEach-Object { $_.Groups[1].Value }) -join '')
        }
        elseif ($t -eq 'str') { if ($vm.Success) { $v = Decode $vm.Groups[1].Value } }
        elseif ($t -eq 'e') { $err = $true; if ($vm.Success) { $v = $vm.Groups[1].Value } }
        elseif ($t -eq 'b') { if ($vm.Success) { $v = [double]$vm.Groups[1].Value } }
        else {
            if ($vm.Success) {
                $d = 0.0
                if ([double]::TryParse($vm.Groups[1].Value, [System.Globalization.NumberStyles]::Float, $inv, [ref]$d)) { $v = $d } else { $v = $vm.Groups[1].Value }
            }
        }
        $cells[$m.Groups[1].Value + $m.Groups[2].Value] = @{ v = $v; f = $f; e = $err; hasF = $fm.Success -or ($inner -match '<f') }
    }
    return $cells
}
function CV($cells, [string]$ref) { if ($cells.ContainsKey($ref)) { return $cells[$ref].v }; return $null }
function CF($cells, [string]$ref) { if ($cells.ContainsKey($ref)) { return $cells[$ref].f }; return $null }
function Num($x) { if ($x -is [double]) { return $x }; return 0.0 }
function Norm([string]$s) {
    if (-not $s) { return '' }
    $d = $s.Normalize([System.Text.NormalizationForm]::FormD)
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $d.ToCharArray()) {
        if ([System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch) -ne [System.Globalization.UnicodeCategory]::NonSpacingMark) { [void]$sb.Append($ch) }
    }
    return ($sb.ToString().ToUpperInvariant() -replace '\s+', ' ').Trim()
}
$excelBase = New-Object DateTime(1899, 12, 30)
function ToDate($serial) { if ($serial -is [double] -and $serial -gt 0) { return $excelBase.AddDays([math]::Floor($serial)) }; return $null }
function Iso($serial) { $d = ToDate $serial; if ($d) { return $d.ToString('yyyy-MM-dd') }; return $null }
function C2N([string]$c) { $n = 0; foreach ($ch in $c.ToCharArray()) { $n = $n * 26 + ([int]$ch - 64) }; return $n }
function N2C([int]$n) { $s = ''; while ($n -gt 0) { $r = ($n - 1) % 26; $s = [char](65 + $r) + $s; $n = [math]::Floor(($n - 1) / 26) }; return $s }
function Round2($x) { return [math]::Round([double]$x, 2) }

# ---------------------------------------------------------------- PREMISSAS ENPAVI
$P = Read-Sheet 'PREMISSAS ENPAVI'
$rowsP = @{}
foreach ($k in $P.Keys) { $r = [int]($k -replace '[A-Z]', ''); $rowsP[$r] = $true }
$maxRowP = ($rowsP.Keys | Measure-Object -Maximum).Maximum

$params = [ordered]@{
    inicio = Iso (CV $P 'C3'); corte = Iso (CV $P 'C4')
    juros = Num (CV $P 'C5'); formaJuros = CV $P 'C6'; pctPadrao = Num (CV $P 'C7'); meses = Num (CV $P 'C8')
}
if (-not $params.corte) { throw "PREMISSAS ENPAVI!C4 (data de corte) está vazia ou sem valor calculado." }
$resumo = [ordered]@{}
foreach ($k in 4..11) {
    $c = $P["K$k"]
    if (-not $c -or -not ($c.v -is [double])) { throw "PREMISSAS ENPAVI!K$k sem valor calculado. Abra a planilha no Excel, salve e rode de novo." }
    $resumo["K$k"] = $c.v
}
$resumo['texto'] = CV $P 'I12'
$rotulos = [ordered]@{}
foreach ($k in 4..11) { $rotulos["K$k"] = CV $P "I$k" }

# cabeçalho do item 1 (linha com CÓDIGO na coluna A)
$hdr1 = $null
for ($r = 1; $r -le $maxRowP; $r++) { if ((Norm (CV $P "A$r")) -eq 'CODIGO') { $hdr1 = $r; break } }
if (-not $hdr1) { throw "Não achei o cabeçalho do item 1 (CÓDIGO) na PREMISSAS ENPAVI." }
function Map-Headers($cells, [int]$row) {
    $map = @{}
    foreach ($ci in 1..20) {
        $col = N2C $ci; $h = Norm (CV $cells "$col$row")
        if (-not $h) { continue }
        if ($h.StartsWith('CODIGO')) { $map['cod'] = $col }
        elseif ($h.StartsWith('DESCRICAO')) { $map['desc'] = $col }
        elseif ($h.StartsWith('GRUPO')) { $map['grupo'] = $col }
        elseif ($h.StartsWith('ABA DE ORIGEM')) { $map['origem'] = $col }
        elseif ($h.StartsWith('CONSIDERAR')) { $map['sn'] = $col }
        elseif ($h.StartsWith('% TRIADE')) { $map['E'] = $col }
        elseif ($h.StartsWith('% ENPAVI SOBRE')) { $map['F'] = $col }
        elseif ($h.StartsWith('% ENPAVI EFETIVO')) { $map['G'] = $col }
        elseif ($h.StartsWith('RESULTADO') -or $h.StartsWith('VALOR 100')) { $map['H'] = $col }
        elseif ($h.StartsWith('PARTE TRIADE')) { $map['I'] = $col }
        elseif ($h.StartsWith('PARTE ENPAVI')) { $map['J'] = $col }
        elseif ($h.StartsWith('OBSERVA')) { $map['obs'] = $col }
    }
    foreach ($need in 'cod', 'desc', 'E', 'F', 'G', 'H', 'I', 'J') { if (-not $map.ContainsKey($need)) { throw "Coluna '$need' não encontrada no cabeçalho da linha $row da PREMISSAS ENPAVI." } }
    return $map
}
$h1 = Map-Headers $P $hdr1

# ---------------------------------------------------------------- ENCONTRO ENPAVI
$E = Read-Sheet 'ENCONTRO ENPAVI'
$cutSerial = CV $E 'B3'; $startSerial = CV $E 'B5'
if (-not ($cutSerial -is [double])) { throw "ENCONTRO ENPAVI!B3 sem valor calculado." }
$meses = New-Object System.Collections.Generic.List[string]
$mesRows = New-Object System.Collections.Generic.List[int]
$r = 12
while ($E.ContainsKey("A$r") -and ((CV $E "A$r") -is [double])) { $meses.Add((Iso (CV $E "A$r"))); $mesRows.Add($r); $r++ }
if ($meses.Count -lt 12) { throw "Não achei os meses da ENCONTRO ENPAVI (coluna A a partir da linha 12)." }
$cutIdx = $meses.IndexOf((Iso $cutSerial))
if ($cutIdx -lt 0) { throw "O mês de corte $(Iso $cutSerial) não está na lista de meses da ENCONTRO ENPAVI." }

function Series([string]$col) {
    $arr = New-Object 'double[]' $mesRows.Count
    for ($i = 0; $i -lt $mesRows.Count; $i++) { $arr[$i] = Round2 (Num (CV $E ($col + $mesRows[$i]))) }
    return , $arr
}
# colunas de fluxo (100%) = códigos na linha 7 antes da coluna vazia que separa o resumo
$flowCols = [ordered]@{}
$ci = 2
while ($true) {
    $col = N2C $ci; $code = CV $E "${col}7"
    if (-not $code) { break }
    $flowCols[(Norm $code)] = @{ col = $col; codigo = $code; desc = CV $E "${col}8"; serie = (Series $col) }
    $ci++
}
# colunas do encontro (resumo mensal)
$encCols = [ordered]@{ AC = 'RESULTADO TRIADE'; AD = 'PARTE ENPAVI'; AE = 'CUSTO 029'; AF = 'OBRAS ANTIGAS'; AG = 'RETIRADAS'; AH = 'RETIRADAS'; AI = 'PROVIS'; AJ = 'SALDO ENPAVI SEM JUROS'; AK = 'TAXA'; AL = 'JUROS DO M'; AM = 'JUROS ACUMULADOS'; AN = 'SALDO ENPAVI COM JUROS' }
$enc = [ordered]@{}
foreach ($col in $encCols.Keys) {
    $h = Norm (CV $E "${col}7")
    if (-not $h.StartsWith($encCols[$col])) { throw "ENCONTRO ENPAVI!${col}7 deveria ser '$($encCols[$col])...' e está '$h'. A estrutura da aba mudou." }
    $enc[$col] = Series $col
}

# ---------------------------------------------------------------- linhas do item 1
$linhas = New-Object System.Collections.Generic.List[object]
$totalRow = $null
$adminObras = @{}
for ($r = $hdr1 + 1; $r -le $maxRowP; $r++) {
    $cod = CV $P "$($h1.cod)$r"
    if (-not $cod) { continue }
    if ((Norm $cod) -eq 'TOTAL') { $totalRow = $r; break }
    $hF = CF $P "$($h1.H)$r"
    $grupo = $null; if ($h1.grupo) { $grupo = CV $P "$($h1.grupo)$r" }
    $origem = $null; if ($h1.origem) { $origem = CV $P "$($h1.origem)$r" }
    $dataGate = $null
    if ($origem -is [double]) { $dataGate = Iso $origem; $origem = $null }
    $tipo = 'outro'; $col = $null; $adm = $null
    $ng = Norm $grupo; $nc = Norm $cod
    if ($flowCols.Contains($nc)) { $col = $flowCols[$nc].col }
    if ($hF -and $hF -match "SUMIFS\('CAIXA OBRA") {
        $tipo = 'adm'
        $partes = New-Object System.Collections.Generic.List[object]
        foreach ($m in [regex]::Matches($hF, "SUMIFS\('CAIXA OBRA ([^']+)'!\`$K:\`$K[^)]*\)(\*\`$E\`$(\d+))?")) {
            $ob = $m.Groups[1].Value; $er = $null
            if ($m.Groups[3].Success) { $er = [int]$m.Groups[3].Value }
            $partes.Add([ordered]@{ obra = $ob; linhaE = $er }); $adminObras[$ob] = $true
        }
        $adm = [ordered]@{ partes = $partes }
    }
    elseif ($hF -and $hF -match "RECEITAS") { $tipo = 'retirada' }
    elseif ($nc -eq '029') { $tipo = 'c029' }
    elseif ($ng -like 'OBRAS ANTIGAS*') { $tipo = 'antiga' }
    elseif ($ng -like 'USINA*' -or $nc -like 'CA?APAV*') { $tipo = 'usina' }
    elseif ($ng -like 'MATRIZ*') { $tipo = 'matriz' }
    elseif ($ng -like 'CONSORCIO*') { $tipo = 'consorcio' }
    elseif ($ng -like 'OBRA PROPRIA*') { $tipo = 'propria' }
    if (($tipo -ne 'adm') -and ($tipo -ne 'retirada') -and (-not $col)) { throw "A linha $r ($cod) da PREMISSAS não tem coluna correspondente na ENCONTRO ENPAVI (linha 7)." }
    $obs = $null; if ($h1.obs) { $obs = CV $P "$($h1.obs)$r" }
    $linhas.Add([ordered]@{
        linha = $r; codigo = $cod; descricao = (CV $P "$($h1.desc)$r"); grupo = $grupo; origem = $origem; data = $dataGate; tipo = $tipo; col = $col
        E = Num (CV $P "$($h1.E)$r"); F = Num (CV $P "$($h1.F)$r"); G = Num (CV $P "$($h1.G)$r")
        H = Round2 (Num (CV $P "$($h1.H)$r")); I = Round2 (Num (CV $P "$($h1.I)$r")); J = Round2 (Num (CV $P "$($h1.J)$r"))
        Ivazio = -not ((CV $P "$($h1.I)$r") -is [double])
        obs = $obs; adm = $adm
    })
}
if (-not $totalRow) { throw "Não achei a linha TOTAL do item 1 na PREMISSAS ENPAVI." }
$total1 = [ordered]@{ I = Round2 (Num (CV $P "$($h1.I)$totalRow")); J = Round2 (Num (CV $P "$($h1.J)$totalRow")) }

# item 2 (provisões)
$prov = New-Object System.Collections.Generic.List[object]
$provTotal = 0.0
$hdr2 = $null
for ($r = $totalRow + 1; $r -le $maxRowP; $r++) { if ((Norm (CV $P "A$r")) -eq 'CODIGO') { $hdr2 = $r; break } }
if ($hdr2) {
    $h2 = Map-Headers $P $hdr2
    $parteEnpaviCol = $h2.J
    for ($r = $hdr2 + 1; $r -le $maxRowP; $r++) {
        $cod = CV $P "A$r"; if (-not $cod) { continue }
        if ((Norm $cod) -eq 'TOTAL') { $provTotal = Round2 (Num (CV $P "$parteEnpaviCol$r")); break }
        $sn = $null; if ($h2.sn) { $sn = CV $P "$($h2.sn)$r" }
        $obs = $null; if ($h2.obs) { $obs = CV $P "$($h2.obs)$r" }
        $prov.Add([ordered]@{ codigo = $cod; descricao = (CV $P "$($h2.desc)$r"); considerar = $sn; valor100 = Round2 (Num (CV $P "$($h2.H)$r")); parteEnpavi = Round2 (Num (CV $P "$parteEnpaviCol$r")); obs = $obs })
    }
}

# ---------------------------------------------------------------- receita acumulada por obra (base da adm. central)
$receitaObras = [ordered]@{}
$mesEnd = @(); foreach ($iso in $meses) { $mesEnd += [DateTime]::ParseExact($iso, 'yyyy-MM-dd', $inv) }
$start = ToDate $startSerial
foreach ($ob in ($adminObras.Keys | Sort-Object)) {
    $cx = Read-Sheet ("CAIXA OBRA " + $ob) '[HKM]'
    $porMes = New-Object 'double[]' $meses.Count
    foreach ($k in $cx.Keys) {
        if ($k[0] -ne 'K') { continue }
        $rr = $k.Substring(1)
        if ([int]$rr -lt 8) { continue }
        $val = $cx[$k].v
        if (-not ($val -is [double]) -or $val -eq 0) { continue }
        $flag = $null; if ($cx.ContainsKey("M$rr")) { $flag = $cx["M$rr"].v }
        if ($flag -and ((Norm ([string]$flag)) -eq 'N')) { continue }
        $d = $null; if ($cx.ContainsKey("H$rr")) { $d = ToDate $cx["H$rr"].v }
        if (-not $d) { continue }
        if ($start -and $d -lt $start) { continue }
        for ($i = 0; $i -lt $mesEnd.Count; $i++) { if ($d -le $mesEnd[$i]) { $porMes[$i] += $val; break } }
    }
    $acc = New-Object 'double[]' $meses.Count; $s = 0.0
    for ($i = 0; $i -lt $meses.Count; $i++) { $s += $porMes[$i]; $acc[$i] = Round2 $s }
    $receitaObras[$ob] = $acc
}

# ---------------------------------------------------------------- RECEITAS ENPAVI-USINAS SP-ESO (retiradas)
$receitas = $null
if ($sheetFile.ContainsKey('RECEITAS ENPAVI-USINAS SP-ESO')) {
    $RC = Read-Sheet 'RECEITAS ENPAVI-USINAS SP-ESO' '[A-F]'
    $porAba = New-Object System.Collections.Generic.List[object]
    $porNome = New-Object System.Collections.Generic.List[object]
    $maxR = 0; foreach ($k in $RC.Keys) { $n = [int]($k -replace '[A-Z]', ''); if ($n -gt $maxR) { $maxR = $n } }
    $sec = 0; $totalAba = $null; $totalNome = $null
    for ($r = 1; $r -le [math]::Min($maxR, 120); $r++) {
        $a = Norm (CV $RC "A$r")
        if ($a -eq 'ABA DE ORIGEM') { $sec = 1; continue }
        if ($a -eq 'EMPRESA' -and (Norm (CV $RC "B$r")).StartsWith('NOME')) { $sec = 2; continue }
        if ($a -eq 'EMPRESA') { break }
        if ($sec -eq 1 -and $a) {
            if ($a -eq 'TOTAL') { $totalAba = [ordered]@{ enpavi = Round2 (Num (CV $RC "B$r")); usinas = Round2 (Num (CV $RC "C$r")); eso = Round2 (Num (CV $RC "D$r")); total = Round2 (Num (CV $RC "E$r")); qtd = Num (CV $RC "F$r") }; $sec = 0; continue }
            $porAba.Add([ordered]@{ aba = (CV $RC "A$r"); enpavi = Round2 (Num (CV $RC "B$r")); usinas = Round2 (Num (CV $RC "C$r")); eso = Round2 (Num (CV $RC "D$r")); total = Round2 (Num (CV $RC "E$r")); qtd = Num (CV $RC "F$r") })
        }
        if ($sec -eq 2 -and $a) {
            if ($a -eq 'TOTAL') { $totalNome = Round2 (Num (CV $RC "E$r")); $sec = 0; continue }
            $porNome.Add([ordered]@{ empresa = (CV $RC "A$r"); nome = (CV $RC "B$r"); receita = Round2 (Num (CV $RC "E$r")); qtd = Num (CV $RC "F$r") })
        }
    }
    $receitas = [ordered]@{ porAba = $porAba; porNome = $porNome; totalAba = $totalAba; totalNome = $totalNome }
}
$zip.Dispose(); $stream.Dispose()

# ---------------------------------------------------------------- JSON + HTML
$flowOut = [ordered]@{}
foreach ($k in $flowCols.Keys) { $fc = $flowCols[$k]; $flowOut[$fc.col] = [ordered]@{ codigo = $fc.codigo; desc = $fc.desc; serie = $fc.serie } }
$site = [ordered]@{
    meta = [ordered]@{
        generatedAt = (Get-Date).ToString('dd/MM/yyyy HH:mm')
        planilhaSalvaEm = $xlsxInfo.LastWriteTime.ToString('dd/MM/yyyy HH:mm')
        sourceFile = $xlsxInfo.Name
    }
    params = $params
    resumo = $resumo
    rotulos = $rotulos
    linhas = $linhas
    total1 = $total1
    provisoes = [ordered]@{ linhas = $prov; totalEnpavi = $provTotal }
    meses = $meses
    corteIdx = $cutIdx
    fluxos = $flowOut
    encontro = $enc
    receitaObras = $receitaObras
    receitas = $receitas
}
$json = $site | ConvertTo-Json -Depth 10 -Compress
$json = $json.Replace('</script', '<\/script')

$tpl = [System.IO.File]::ReadAllText((Join-Path $PSScriptRoot 'template.html'), $utf8)
$logo = [System.IO.File]::ReadAllText((Join-Path $PSScriptRoot 'logo_base64.txt'), [System.Text.Encoding]::ASCII).Trim()
$final = $tpl.Replace('__SITE_DATA_JSON__', $json).Replace('__LOGO_BASE64__', $logo)
$outPath = Join-Path $ProjectDir 'index.html'
[System.IO.File]::WriteAllText($outPath, $final, $utf8)

$fmt = { param($v) ([double]$v).ToString('N2', [System.Globalization.CultureInfo]::GetCultureInfo('pt-BR')) }
Write-Host ("Apresentação atualizada: " + $outPath)
Write-Host ("Corte: " + $params.corte + " | Saldo sem juros: R$ " + (& $fmt $resumo.K9) + " | Juros: R$ " + (& $fmt $resumo.K10) + " | Saldo com juros: R$ " + (& $fmt $resumo.K11))
