# gerar_dados.ps1 — Gera dados.js a partir da planilha Excel (compativel PS 5.1)
$ErrorActionPreference = "Stop"

$ArquivoExcel = Join-Path $PSScriptRoot "Fluxo de caixa TRB.xlsx"
$ArquivoJS    = Join-Path $PSScriptRoot "dados.js"

if (-not (Test-Path $ArquivoExcel)) {
    Write-Host "ERRO: Arquivo nao encontrado: $ArquivoExcel" -ForegroundColor Red
    Read-Host "Pressione Enter para sair"
    exit 1
}

# Normaliza string: minusculo, sem acento (usando NFD Unicode — independente de encoding)
function Norm([string]$s) {
    if ([string]::IsNullOrEmpty($s)) { return "" }
    $s = $s.Trim().ToLower()
    $decomp = $s.Normalize([System.Text.NormalizationForm]::FormD)
    $sb = [System.Text.StringBuilder]::new()
    foreach ($ch in $decomp.ToCharArray()) {
        $cat = [System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch)
        if ($cat -ne [System.Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$sb.Append($ch)
        }
    }
    return $sb.ToString()
}

# Escape para uso em string JSON
function Esc([string]$s) {
    $s = $s.Trim()
    $s = $s -replace '\\','\\'
    $s = $s -replace '"','\"'
    $s = $s -replace "`r",''
    $s = $s -replace "`n",' '
    return $s
}

# Converte qualquer valor para string sem lancar erro
function SafeStr($v) {
    if ($null -eq $v) { return "" }
    return [string]$v
}

$excel = $null
try {
    Write-Host "Iniciando Excel..." -ForegroundColor Cyan
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible        = $false
    $excel.DisplayAlerts  = $false
    $excel.ScreenUpdating = $false

    Write-Host "Abrindo planilha..." -ForegroundColor Cyan
    $wb = $excel.Workbooks.Open($ArquivoExcel, 0, $true)

    # ─────────────────────────────────────────────────────────────────────────
    # 1. Ler aba "Unidades 2" PRIMEIRO — monta os mapas de Empresa 2
    # ─────────────────────────────────────────────────────────────────────────
    $wsUnid = $null
    foreach ($s in $wb.Worksheets) {
        if ($s.Name -eq "Unidades 2") { $wsUnid = $s; break }
    }
    if ($null -eq $wsUnid) {
        foreach ($s in $wb.Worksheets) {
            if ($s.Name -match "Unidade") { $wsUnid = $s; break }
        }
    }

    $empresasTipoMap   = @{}   # Empresa2 → Tipo
    $empresasGrupoMap  = @{}   # Empresa2 → Grupo
    $empresasRegimeMap = @{}   # Empresa2 → Regime
    $cnEmp2Map         = @{}   # "EmpOrig|||CN" → Empresa2
    $empFallbackMap    = @{}   # EmpOrig → primeira Empresa2 encontrada
    $permutantesEmp2Set = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    if ($null -ne $wsUnid) {
        Write-Host "Lendo aba: $($wsUnid.Name)..." -ForegroundColor Cyan
        $usadoUnid = $wsUnid.UsedRange
        $nRowsUnid = $usadoUnid.Rows.Count
        $nColsUnid = $usadoUnid.Columns.Count
        $valsUnid  = $usadoUnid.Value2

        $hMapUnid = @{}
        for ($c = 1; $c -le $nColsUnid; $c++) {
            $raw = SafeStr $valsUnid[1,$c]
            $key = Norm $raw
            if ($key -ne "") { $hMapUnid[$key] = $c }
        }

        $uEmpCol    = if ($hMapUnid.ContainsKey("empresa"))           { $hMapUnid["empresa"]           } else { 1 }
        $uCNCol     = if ($hMapUnid.ContainsKey("centro de negocio")) { $hMapUnid["centro de negocio"] } else { $null }
        $uEmp2Col   = if ($hMapUnid.ContainsKey("empresa 2"))         { $hMapUnid["empresa 2"]         } else { $null }
        $uTipoCol   = if ($hMapUnid.ContainsKey("tipo"))              { $hMapUnid["tipo"]              } else { 2 }
        $uGrupoCol  = if ($hMapUnid.ContainsKey("grupo"))             { $hMapUnid["grupo"]             } else { $null }
        $uRegimeCol = if ($hMapUnid.ContainsKey("regime"))            { $hMapUnid["regime"]            } else { $null }

        for ($r = 2; $r -le $nRowsUnid; $r++) {
            $empOrig = (SafeStr $valsUnid[$r,$uEmpCol]).Trim()
            if ($empOrig -eq "") { continue }

            $cn     = if ($null -ne $uCNCol)     { (SafeStr $valsUnid[$r,$uCNCol]).Trim()    } else { "" }
            $emp2   = if ($null -ne $uEmp2Col)   { (SafeStr $valsUnid[$r,$uEmp2Col]).Trim()  } else { $empOrig }
            $tipo   = (SafeStr $valsUnid[$r,$uTipoCol]).Trim()
            $grupo  = if ($null -ne $uGrupoCol)  { (SafeStr $valsUnid[$r,$uGrupoCol]).Trim()  } else { "" }
            $regime = if ($null -ne $uRegimeCol) { (SafeStr $valsUnid[$r,$uRegimeCol]).Trim() } else { "" }

            if ($emp2 -eq "") { $emp2 = $empOrig }

            if ($cn -ne "") {
                $cnKey = "$empOrig|||$cn"
                if (-not $cnEmp2Map.ContainsKey($cnKey)) { $cnEmp2Map[$cnKey] = $emp2 }
            }
            if (-not $empFallbackMap.ContainsKey($empOrig)) { $empFallbackMap[$empOrig] = $emp2 }

            if ($grupo -eq "Permutante") {
                [void]$permutantesEmp2Set.Add($emp2)
            } elseif (-not $empresasTipoMap.ContainsKey($emp2)) {
                $empresasTipoMap[$emp2]   = $tipo
                $empresasGrupoMap[$emp2]  = $grupo
                $empresasRegimeMap[$emp2] = $regime
            }
        }
        Write-Host "Unidades 2: $($empresasTipoMap.Count) empreendimentos." -ForegroundColor Cyan
    } else {
        Write-Host "Aba 'Unidades 2' nao encontrada - ignorada." -ForegroundColor Yellow
    }

    # ─────────────────────────────────────────────────────────────────────────
    # 2. Ler aba "Fluxo Realizado"
    # ─────────────────────────────────────────────────────────────────────────
    $ws = $null
    foreach ($s in $wb.Worksheets) {
        if ($s.Name -match "Fluxo.{0,5}Realizado" -or $s.Name -match "Realizado") {
            $ws = $s; break
        }
    }

    if ($null -eq $ws) {
        $nomes = @()
        foreach ($s in $wb.Worksheets) { $nomes += $s.Name }
        Write-Host "ERRO: Aba 'Fluxo Realizado' nao encontrada." -ForegroundColor Red
        Write-Host "Abas disponiveis: $($nomes -join ', ')" -ForegroundColor Yellow
        $wb.Close($false)
        Read-Host "Pressione Enter para sair"
        exit 1
    }

    Write-Host "Aba encontrada: $($ws.Name)" -ForegroundColor Green

    $usado = $ws.UsedRange
    $nRows = $usado.Rows.Count
    $nCols = $usado.Columns.Count

    Write-Host "Lendo $($nRows - 1) linhas de dados..." -ForegroundColor Cyan

    $vals = $usado.Value2

    $hMap = @{}
    for ($c = 1; $c -le $nCols; $c++) {
        $raw = SafeStr $vals[1,$c]
        $key = Norm $raw
        if ($key -ne "") { $hMap[$key] = $c }
    }

    function Col([string]$name) {
        $n = Norm $name
        if ($hMap.ContainsKey($n)) { return $hMap[$n] }
        return $null
    }

    $iEmpresa = Col "empresa"
    $iValor   = Col "valor"
    $iCI      = Col "class interna"
    $iTipo    = Col "tipo"
    $iPR      = Col "p/r"
    $iData    = Col "data"
    $iCN      = Col "centro de negocio"
    $iMes     = if ($hMap.ContainsKey("mes")) { $hMap["mes"] } else { Col "mes" }

    $faltando = @()
    if ($null -eq $iEmpresa) { $faltando += "Empresa" }
    if ($null -eq $iMes)     { $faltando += "Mes" }
    if ($null -eq $iValor)   { $faltando += "Valor" }
    if ($null -eq $iCI)      { $faltando += "Class Interna" }
    if ($faltando.Count -gt 0) {
        Write-Host "AVISO: Colunas nao localizadas: $($faltando -join ', ')" -ForegroundColor Yellow
    }

    $linhas = [System.Collections.Generic.List[string]]::new()

    for ($r = 2; $r -le $nRows; $r++) {

        $emp = ""
        if ($null -ne $iEmpresa) { $emp = (SafeStr $vals[$r,$iEmpresa]).Trim() }
        if ($emp -eq "") { continue }

        $mes = 0
        if ($null -ne $iMes) {
            $mesRaw = $vals[$r,$iMes]
            if ($null -ne $mesRaw) { try { $mes = [int][double]$mesRaw } catch { $mes = 0 } }
        }
        if ($mes -lt 1 -or $mes -gt 12) { continue }

        $valor = 0.0
        if ($null -ne $iValor) {
            $valorRaw = $vals[$r,$iValor]
            if ($null -ne $valorRaw) { try { $valor = [double]$valorRaw } catch { $valor = 0.0 } }
        }

        $ci = ""
        if ($null -ne $iCI) {
            $ci = (SafeStr $vals[$r,$iCI]).Trim()
            $ci = $ci -replace '\s*\(Permutante\)\s*$', ''
            $ci = $ci -replace '\s*\(permutante\)\s*$', ''
        }
        if ($ci -eq "") { continue }

        $tipo = ""
        if ($null -ne $iTipo) { $tipo = (SafeStr $vals[$r,$iTipo]).Trim() }

        $pr = "R"
        if ($null -ne $iPR) {
            $prRaw = (SafeStr $vals[$r,$iPR]).Trim().ToUpper()
            if ($prRaw -eq "P") { $pr = "P" }
        }

        $dataStr = ""
        if ($null -ne $iData) {
            $dv = $vals[$r,$iData]
            if ($null -ne $dv) {
                try {
                    $dt = [DateTime]::FromOADate([double]$dv)
                    $dataStr = $dt.ToString("yyyy-MM-dd")
                } catch { $dataStr = "" }
            }
        }

        $cnJ = ""
        if ($null -ne $iCN) { $cnJ = (SafeStr $vals[$r,$iCN]).Trim() }

        # Resolver Empresa 2 via CN lookup
        $empFinal = $emp
        if ($cnEmp2Map.Count -gt 0) {
            $cnKey = "$emp|||$cnJ"
            if ($cnEmp2Map.ContainsKey($cnKey)) {
                $empFinal = $cnEmp2Map[$cnKey]
            } elseif ($empFallbackMap.ContainsKey($emp)) {
                $empFinal = $empFallbackMap[$emp]
            }
        }

        # Ignorar Permutantes
        if ($permutantesEmp2Set.Contains($empFinal)) { continue }

        $valorStr = $valor.ToString("F2", [System.Globalization.CultureInfo]::InvariantCulture)
        $linha = "{`"Empresa`":`"$(Esc $empFinal)`",`"Mes`":$mes,`"Valor`":$valorStr,`"Class Interna`":`"$(Esc $ci)`",`"Tipo`":`"$(Esc $tipo)`",`"PR`":`"$pr`",`"Data`":`"$dataStr`",`"CN`":`"$(Esc $cnJ)`"}"
        $linhas.Add($linha)
    }

    Write-Host "Fluxo Realizado: $($linhas.Count) lancamentos." -ForegroundColor Cyan

    # ─────────────────────────────────────────────────────────────────────────
    # 3. Ler aba Recebíveis
    # ─────────────────────────────────────────────────────────────────────────
    $wsRec = $null
    foreach ($s in $wb.Worksheets) {
        if ($s.Name -match "Receb") { $wsRec = $s; break }
    }

    if ($null -ne $wsRec) {
        Write-Host "Lendo aba: $($wsRec.Name)..." -ForegroundColor Cyan
        $usadoRec = $wsRec.UsedRange
        $nRowsRec = $usadoRec.Rows.Count
        $nColsRec = $usadoRec.Columns.Count
        $valsRec  = $usadoRec.Value2

        $hMapRec = @{}
        for ($c = 1; $c -le $nColsRec; $c++) {
            $raw = SafeStr $valsRec[1,$c]
            $key = Norm $raw
            if ($key -ne "") { $hMapRec[$key] = $c }
        }

        function ColRec([string]$name) {
            $n = Norm $name
            if ($hMapRec.ContainsKey($n)) { return $hMapRec[$n] }
            return $null
        }

        # Usar PRIMEIRA coluna "Empresa" (código curto) — há duas colunas com esse nome
        $rEmpresa = $null
        for ($c = 1; $c -le $nColsRec; $c++) {
            if ((Norm (SafeStr $valsRec[1,$c])) -eq "empresa") { $rEmpresa = $c; break }
        }
        $rMes     = ColRec "mes"
        $rAno     = ColRec "ano"
        $rValor   = ColRec "cr aberto"
        $rCI      = ColRec "class interna"
        $rTipo    = ColRec "tipo"
        $rPR      = ColRec "p/r"

        # Coluna "Mês" no formato YYYY-MM (texto) — detectada por valor, usada como fallback
        $rMesData = $null
        for ($c = 1; $c -le $nColsRec; $c++) {
            $sample = SafeStr $valsRec[2,$c]
            if ($sample -match "^\d{4}-\d{2}") { $rMesData = $c; break }
        }
        Write-Host "Recebíveis: coluna Mes=$rMes Ano=$rAno MesData=$rMesData" -ForegroundColor Cyan

        $contadorRec = 0
        for ($r = 2; $r -le $nRowsRec; $r++) {
            $emp = ""
            if ($null -ne $rEmpresa) { $emp = (SafeStr $valsRec[$r,$rEmpresa]).Trim() }
            if ($emp -eq "") { continue }

            $mes = 0
            if ($null -ne $rMes) {
                $mesRaw = $valsRec[$r,$rMes]
                if ($null -ne $mesRaw) { try { $mes = [int][double]$mesRaw } catch { $mes = 0 } }
            }
            if ($mes -lt 1 -or $mes -gt 12) { continue }

            $valor = 0.0
            if ($null -ne $rValor) {
                $valorRaw = $valsRec[$r,$rValor]
                if ($null -ne $valorRaw) { try { $valor = [double]$valorRaw } catch { $valor = 0.0 } }
            }
            if ($valor -lt 0) { $valor = -$valor }

            $ci = "Venda de imoveis"
            if ($null -ne $rCI) {
                $ciRaw = (SafeStr $valsRec[$r,$rCI]).Trim()
                if ($ciRaw -ne "") { $ci = $ciRaw }
            }
            $ci = $ci -replace '\s*\(Permutante\)\s*$', ''
            $ci = $ci -replace '\s*\(permutante\)\s*$', ''

            $tipo = "Operacional"
            if ($null -ne $rTipo) {
                $tipoRaw = (SafeStr $valsRec[$r,$rTipo]).Trim()
                if ($tipoRaw -ne "") { $tipo = $tipoRaw }
            }

            $pr = "P"
            if ($null -ne $rPR) {
                $prRaw = (SafeStr $valsRec[$r,$rPR]).Trim().ToUpper()
                if ($prRaw -eq "R") { $pr = "R" }
            }

            # Determinar ano do lançamento: coluna Ano (primário) → YYYY-MM texto (fallback) → ano atual
            $anoRec = 0
            if ($null -ne $rAno) {
                $anoRaw = $valsRec[$r,$rAno]
                if ($null -ne $anoRaw) { try { $anoRec = [int][double]$anoRaw } catch { $anoRec = 0 } }
            }
            $dataStr = ""
            if ($anoRec -gt 0) {
                $dataStr = "$anoRec-$($mes.ToString('D2'))-01"
            } elseif ($null -ne $rMesData) {
                $mesDataRaw = SafeStr $valsRec[$r,$rMesData]
                if ($mesDataRaw -match "^(\d{4})-(\d{2})") {
                    $dataStr = "$($Matches[1])-$($Matches[2])-01"
                }
            }
            if ($dataStr -eq "") {
                $dataStr = "$(( Get-Date).Year)-$($mes.ToString('D2'))-01"
            }

            # Recebíveis não têm CN — fallback por empresa
            $empFinal = $emp
            if ($empFallbackMap.ContainsKey($emp)) { $empFinal = $empFallbackMap[$emp] }

            # Ignorar Permutantes
            if ($permutantesEmp2Set.Contains($empFinal)) { continue }

            $valorStr = $valor.ToString("F2", [System.Globalization.CultureInfo]::InvariantCulture)
            $linha = "{`"Empresa`":`"$(Esc $empFinal)`",`"Mes`":$mes,`"Valor`":$valorStr,`"Class Interna`":`"$(Esc $ci)`",`"Tipo`":`"$(Esc $tipo)`",`"PR`":`"$pr`",`"Data`":`"$dataStr`"}"
            $linhas.Add($linha)
            $contadorRec++
        }
        Write-Host "Recebiveis: $contadorRec lancamentos." -ForegroundColor Cyan
    } else {
        Write-Host "Aba Recebiveis nao encontrada - ignorada." -ForegroundColor Yellow
    }

    # ─────────────────────────────────────────────────────────────────────────
    # 4. Ler aba Saldo inicial de caixa
    # ─────────────────────────────────────────────────────────────────────────
    $wsSI = $null
    foreach ($s in $wb.Worksheets) {
        if ((Norm $s.Name) -match "saldo.{0,10}inicial") { $wsSI = $s; break }
    }

    $saldoInicialMap = @{}
    if ($null -ne $wsSI) {
        Write-Host "Lendo aba: $($wsSI.Name)..." -ForegroundColor Cyan
        $usadoSI = $wsSI.UsedRange
        $nRowsSI = $usadoSI.Rows.Count
        $nColsSI = $usadoSI.Columns.Count
        $valsSI  = $usadoSI.Value2

        $hMapSI = @{}
        for ($c = 1; $c -le $nColsSI; $c++) {
            $raw = SafeStr $valsSI[1,$c]
            $key = Norm $raw
            if ($key -ne "") { $hMapSI[$key] = $c }
        }

        $siEmpCol   = if ($hMapSI.ContainsKey("empresa")) { $hMapSI["empresa"] } else { 1 }
        $siSaldoCol = $null
        foreach ($k in $hMapSI.Keys) {
            if ($k -match "saldo" -or $k -match "valor") { $siSaldoCol = $hMapSI[$k]; break }
        }
        if ($null -eq $siSaldoCol) { $siSaldoCol = if ($nColsSI -ge 2) { 2 } else { 1 } }

        for ($r = 2; $r -le $nRowsSI; $r++) {
            $emp = (SafeStr $valsSI[$r,$siEmpCol]).Trim()
            if ($emp -eq "") { continue }
            $sv = 0.0
            $svRaw = $valsSI[$r,$siSaldoCol]
            if ($null -ne $svRaw) { try { $sv = [double]$svRaw } catch { $sv = 0.0 } }
            # Remapear para Empresa 2 via fallback (resolve split por CN)
            $empSI = $emp
            if ($empFallbackMap.ContainsKey($emp)) { $empSI = $empFallbackMap[$emp] }
            $saldoInicialMap[$empSI] = ($saldoInicialMap[$empSI] -as [double]) + $sv
        }
        Write-Host "Saldo inicial: $($saldoInicialMap.Count) empreendimentos." -ForegroundColor Cyan
    } else {
        Write-Host "Aba 'Saldo inicial de caixa' nao encontrada - ignorada." -ForegroundColor Yellow
    }

    $wb.Close($false)

    # ─────────────────────────────────────────────────────────────────────────
    # 5. Gerar dados.js
    # ─────────────────────────────────────────────────────────────────────────
    $ts    = Get-Date -Format "dd/MM/yyyy HH:mm"
    $corpo = $linhas -join ",`r`n  "

    $siEntries = @()
    foreach ($k in $saldoInicialMap.Keys) {
        if ($permutantesEmp2Set.Contains($k)) { continue }
        # Ignorar empresas não reconhecidas (não mapeadas via Unidades 2)
        if (-not $empresasTipoMap.ContainsKey($k)) {
            Write-Host "AVISO: Saldo inicial ignorado para empresa nao mapeada: '$k'" -ForegroundColor Yellow
            continue
        }
        $sv = $saldoInicialMap[$k].ToString("F2", [System.Globalization.CultureInfo]::InvariantCulture)
        $siEntries += "`"$(Esc $k)`":$sv"
    }
    $siJson = "{" + ($siEntries -join ",") + "}"

    $tipoEntries = @()
    foreach ($k in $empresasTipoMap.Keys) {
        $tipoEntries += "`"$(Esc $k)`":`"$(Esc $empresasTipoMap[$k])`""
    }
    $tipoJson = "{" + ($tipoEntries -join ",") + "}"

    $grupoEntries = @()
    foreach ($k in $empresasGrupoMap.Keys) {
        if ($empresasGrupoMap[$k] -ne "") {
            $grupoEntries += "`"$(Esc $k)`":`"$(Esc $empresasGrupoMap[$k])`""
        }
    }
    $grupoJson = "{" + ($grupoEntries -join ",") + "}"

    $regimeEntries = @()
    foreach ($k in $empresasRegimeMap.Keys) {
        if ($empresasRegimeMap[$k] -ne "") {
            $regimeEntries += "`"$(Esc $k)`":`"$(Esc $empresasRegimeMap[$k])`""
        }
    }
    $regimeJson = "{" + ($regimeEntries -join ",") + "}"

    $jsContent = "var DADOS_DFC_GERADO = '$ts';`r`nvar DADOS_DFC = [`r`n  $corpo`r`n];`r`nvar SALDO_INICIAL_CAIXA = $siJson;`r`nvar EMPRESAS_TIPO = $tipoJson;`r`nvar EMPRESAS_GRUPO = $grupoJson;`r`nvar EMPRESAS_REGIME = $regimeJson;"

    # Gravar sem BOM (UTF-8 puro) — BOM em arquivos .js pode impedir carregamento no browser
    $utf8NoBom = [System.Text.UTF8Encoding]::new($false)
    [System.IO.File]::WriteAllText($ArquivoJS, $jsContent, $utf8NoBom)

    Write-Host ""
    Write-Host "Concluido! $($linhas.Count) lancamentos gravados em:" -ForegroundColor Green
    Write-Host "  $ArquivoJS" -ForegroundColor Green

} catch {
    Write-Host ""
    Write-Host "ERRO: $_" -ForegroundColor Red
    Read-Host "Pressione Enter para sair"
    exit 1
} finally {
    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}
    }
}
