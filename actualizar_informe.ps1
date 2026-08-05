# Genera el informe NPS de COBRA (Zona Punta Arenas) a partir de NPS GENERAL.xlsm
# y lo publica en GitHub Pages.
# Uso: doble clic en "Actualizar_Informe.bat", o ejecutar este script en PowerShell.
#      Editar $Mes / $Anio abajo antes de correr para un mes distinto.

param(
    [int]$Mes = 7,
    [int]$Anio = 2026
)

$ErrorActionPreference = "Stop"

$ExcelPath = "C:\Users\jvodn\Downloads\Web\BBDD\NPS GENERAL.xlsm"
$RepoDir   = "C:\Users\jvodn\Downloads\Web\informe-nps"
$TemplatePath = Join-Path $RepoDir "template.html"
$OutputPath   = Join-Path $RepoDir "index.html"

$MESES_ES = @('','Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre')

Write-Host "==> Leyendo Excel: $ExcelPath"
if (-not (Test-Path $ExcelPath)) {
    Write-Host "ERROR: no se encontro el archivo Excel en esa ruta." -ForegroundColor Red
    exit 1
}

function SafeCount($collection) {
    if ($null -eq $collection) { return 0 }
    return @($collection).Count
}

# ---------- 1. Extraer el .xlsm (es un .zip) ----------
$extractDir = Join-Path $env:TEMP "nps_informe_extract"
if (Test-Path $extractDir) { Remove-Item -Recurse -Force $extractDir }
New-Item -ItemType Directory -Force -Path $extractDir | Out-Null
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::ExtractToDirectory($ExcelPath, $extractDir)

# ---------- 2. Shared strings ----------
Write-Host "==> Cargando shared strings..."
[xml]$ss = New-Object System.Xml.XmlDocument
$ss.Load("$extractDir\xl\sharedStrings.xml")
$siNodes = $ss.sst.si
$strings = New-Object string[] $siNodes.Count
$idx = 0
foreach ($si in $siNodes) {
    if ($si.t -ne $null) {
        if ($si.t -is [string]) { $strings[$idx] = [string]$si.t } else { $strings[$idx] = [string]$si.t.'#text' }
    } elseif ($si.r) {
        $sb = New-Object System.Text.StringBuilder
        foreach ($r in $si.r) { if ($r.t -is [string]) { [void]$sb.Append($r.t) } else { [void]$sb.Append($r.t.'#text') } }
        $strings[$idx] = $sb.ToString()
    } else { $strings[$idx] = '' }
    $idx++
}

function Col-ToNum($colRef) {
    $col = $colRef -replace '[0-9]', ''
    $num = 0
    foreach ($c in $col.ToCharArray()) { $num = $num * 26 + ([int][char]$c - [int][char]'A' + 1) }
    return $num
}

# ---------- 3. Encontrar la hoja "BBDD" ----------
[xml]$wb = New-Object System.Xml.XmlDocument
$wb.Load("$extractDir\xl\workbook.xml")
$bbddSheet = $wb.workbook.sheets.sheet | Where-Object { $_.name -eq 'BBDD' }
if (-not $bbddSheet) { Write-Host "ERROR: no se encontro la hoja 'BBDD'." -ForegroundColor Red; exit 1 }
$sheetRid = ($bbddSheet.Attributes | Where-Object { $_.Name -eq 'r:id' }).Value
[xml]$wbRels = New-Object System.Xml.XmlDocument
$wbRels.Load("$extractDir\xl\_rels\workbook.xml.rels")
$rel = $wbRels.Relationships.Relationship | Where-Object { $_.Id -eq $sheetRid }
$sheetFile = Join-Path "$extractDir\xl" $rel.Target

# ---------- 4. Leer y filtrar filas ----------
Write-Host "==> Procesando hoja BBDD (esto puede tardar ~2 min por el tamano del archivo)..."
$reader = [System.Xml.XmlReader]::Create($sheetFile)
$currentRow = $null
$rowIdx = 0
$matched = New-Object System.Collections.Generic.List[object]
$mesStr = "$Mes"

function Process-Row($r) {
    if ($r[28] -eq 'COBRA' -and $r[32] -eq 'ACTIVO' -and $r[39] -eq '0' -and $r[26] -eq $mesStr) {
        $fechaSerial = [double]$r[11]
        $anioFila = [System.DateTime]::FromOADate($fechaSerial).Year
        if ($anioFila -ne $Anio) { return }
        $script:matched.Add([pscustomobject]@{
            tecnico = [string]$r[6]
            dia     = [int]$r[33]
            nota    = if ($r[13] -and [string]$r[13] -ne 'NULL') { [string]$r[13] } else { $null }
            obs     = if ($r[14] -and [string]$r[14] -ne 'NULL') { [string]$r[14] } else { $null }
            nps     = [string]$r[29]
            tipo    = [string]$r[31]
        })
    }
}

while ($reader.Read()) {
    if ($reader.NodeType -eq [System.Xml.XmlNodeType]::Element) {
        if ($reader.Name -eq 'row') {
            if ($currentRow -ne $null -and $rowIdx -gt 0) { Process-Row $currentRow }
            $currentRow = @{}
            $rowIdx++
        } elseif ($reader.Name -eq 'c') {
            $script:cellRef = $reader.GetAttribute('r')
            $script:cellType = $reader.GetAttribute('t')
        } elseif ($reader.Name -eq 'v') {
            $val = $reader.ReadElementContentAsString()
            if ($script:cellType -eq 's') { $val = $strings[[int]$val] }
            $colNum = Col-ToNum $script:cellRef
            $currentRow[$colNum] = $val
        }
    }
}
if ($currentRow -ne $null) { Process-Row $currentRow }
$reader.Close()

$enviadas = $matched.Count
$respondidas = $matched | Where-Object { $_.nota -ne $null }
Write-Host "==> Enviadas=$enviadas Respondidas=$($respondidas.Count)"
if ($respondidas.Count -eq 0) { Write-Host "ERROR: no hay encuestas respondidas para ese periodo/filtro." -ForegroundColor Red; exit 1 }

# ---------- 5. Agregados ----------
function GroupNps($group) {
    $p = SafeCount ($group | Where-Object nps -eq 'PROMOTOR')
    $n = SafeCount ($group | Where-Object nps -eq 'NEUTRO')
    $d = SafeCount ($group | Where-Object nps -eq 'DETRACTOR')
    $t = @($group).Count
    [pscustomobject]@{ P=$p; N=$n; D=$d; total=$t; nps=[math]::Round((($p-$d)/$t)*100,1) }
}

$daily = @($respondidas | Group-Object dia | Sort-Object { [int]$_.Name } | ForEach-Object {
    $agg = GroupNps $_.Group
    [pscustomobject]@{ dia=[int]$_.Name; P=$agg.P; N=$agg.N; D=$agg.D; total=$agg.total; nps=$agg.nps }
})

$weekly = @($respondidas | Group-Object -Property { [math]::Ceiling($_.dia/7.0) } | Sort-Object { [int]$_.Name } | ForEach-Object {
    $agg = GroupNps $_.Group
    [pscustomobject]@{ semana=[int]$_.Name; P=$agg.P; N=$agg.N; D=$agg.D; total=$agg.total; nps=$agg.nps }
})

$tipo = @($respondidas | Group-Object tipo | ForEach-Object {
    $agg = GroupNps $_.Group
    [pscustomobject]@{ tipo=$_.Name; P=$agg.P; N=$agg.N; D=$agg.D; total=$agg.total; nps=$agg.nps }
})

$tecnicos = @($respondidas | Group-Object tecnico | Where-Object { $_.Count -ge 5 } | ForEach-Object {
    $agg = GroupNps $_.Group
    [pscustomobject]@{ tecnico=$_.Name; P=$agg.P; N=$agg.N; D=$agg.D; total=$agg.total; nps=$agg.nps }
})

$totalP = SafeCount ($respondidas | Where-Object nps -eq 'PROMOTOR')
$totalN = SafeCount ($respondidas | Where-Object nps -eq 'NEUTRO')
$totalD = SafeCount ($respondidas | Where-Object nps -eq 'DETRACTOR')
$npsGeneral = [math]::Round((($totalP-$totalD)/$respondidas.Count)*100,1)

# ---------- 6. Highlights (mejor / menor desempeno) ----------
function KeywordTags($quotes) {
    $text = ($quotes -join ' ').ToLowerInvariant()
    $tags = New-Object System.Collections.Generic.List[string]
    if ($text -match 'dia|demor|espera|tiempo') { $tags.Add('Tiempos de respuesta') }
    if ($text -match 'limpi|residuo|orden') { $tags.Add('Orden/limpieza post-trabajo') }
    if ($text -match 'mal educad|trato|grose') { $tags.Add('Trato al cliente') }
    if ($text -match 'clarid|explic|dud') { $tags.Add('Claridad al explicar la solucion') }
    if ($text -match 'intermitenc|sigue|persist|no funciona') { $tags.Add('Calidad tecnica (falla persiste)') }
    if ($tags.Count -eq 0) { $tags.Add('Ver comentarios de clientes') }
    $result = @($tags | Select-Object -Unique | Select-Object -First 2)
    return ,$result
}

$topCandidates = @($tecnicos | Where-Object { $_.D -eq 0 -and $_.total -ge 10 } | Sort-Object nps, total -Descending | Select-Object -First 4)
$bottomCandidates = @($tecnicos | Sort-Object nps | Select-Object -First 4)

$topHighlights = @($topCandidates | ForEach-Object {
    $tName = $_.tecnico
    $quotesPool = $respondidas | Where-Object { $_.tecnico -eq $tName -and $_.obs -ne $null -and $_.nps -ne 'DETRACTOR' } | Sort-Object { $_.obs.Length } -Descending
    $quotes = @($quotesPool | Select-Object -First 1 -ExpandProperty obs | Where-Object { $_ -ne $null })
    $maxTotal = ($topCandidates | Measure-Object total -Maximum).Maximum
    $tag = if ($_.total -eq $maxTotal) { 'Mayor volumen del equipo, sin detractores' } elseif ($_.total -ge 20) { 'Buen volumen + consistencia' } else { 'Cero friccion reportada' }
    [pscustomobject]@{ tecnico=$_.tecnico; nps=$_.nps; total=$_.total; N=$_.N; D=$_.D; tags=@($tag); quotes=$quotes }
})

$bottomHighlights = @($bottomCandidates | ForEach-Object {
    $tName = $_.tecnico
    $quotesPool = $respondidas | Where-Object { $_.tecnico -eq $tName -and $_.obs -ne $null -and $_.nps -ne 'PROMOTOR' } | Sort-Object { $_.obs.Length } -Descending
    $quotes = @($quotesPool | Select-Object -First 2 -ExpandProperty obs | Where-Object { $_ -ne $null })
    $tags = KeywordTags $quotes
    [pscustomobject]@{ tecnico=$_.tecnico; nps=$_.nps; total=$_.total; N=$_.N; D=$_.D; tags=$tags; quotes=$quotes }
})

# ---------- 7. Conclusiones (auto-generadas) ----------
$primera = $weekly[0]; $ultima = $weekly[-1]
$pendiente = if ($weekly.Count -gt 1) { [math]::Round(($ultima.nps - $primera.nps) / ($weekly.Count - 1), 1) } else { 0 }
$tipoSorted = $tipo | Sort-Object nps -Descending
$brecha = [math]::Round($tipoSorted[0].nps - $tipoSorted[-1].nps, 1)
$tasaResp = [math]::Round($respondidas.Count / $enviadas * 100, 1)
$topNombres = ($topHighlights | ForEach-Object { (Get-Culture).TextInfo.ToTitleCase($_.tecnico.ToLower()) }) -join ', '
$bottomNombres = ($bottomHighlights | ForEach-Object { (Get-Culture).TextInfo.ToTitleCase($_.tecnico.ToLower()) }) -join ', '

$conclusiones = @(
    "<strong>Tendencia del mes:</strong> el NPS paso de $($primera.nps)% (semana 1) a $($ultima.nps)% (semana $($weekly.Count)), con una pendiente promedio de $(if($pendiente -ge 0){'+'})$pendiente pts NPS por semana."
    "<strong>$($tipoSorted[-1].tipo) como palanca de mejora:</strong> con $($tipoSorted[-1].nps)% vs $($tipoSorted[0].nps)% de $($tipoSorted[0].tipo), priorizar auditorias de calidad y refuerzo de capacitacion en ese protocolo (brecha de $brecha pts)."
    "<strong>Coaching dirigido:</strong> priorizar retroalimentacion 1:1 con $bottomNombres, cuyos comentarios apuntan a causas concretas y accionables."
    "<strong>Reconocimiento y buenas practicas:</strong> $topNombres cerraron el mes sin detractores (o casi) pese a buen volumen -- usarlos como referentes en instancias de capacitacion interna."
    "<strong>Tasa de respuesta ($tasaResp%):</strong> evaluar canal y timing de envio de la encuesta para mejorar la representatividad estadistica, especialmente en tecnicos con pocas respuestas."
)

# ---------- 8. Construir DATA y generar index.html ----------
$data = [ordered]@{
    periodo      = "$($MESES_ES[$Mes]) $Anio"
    zona         = "Punta Arenas"
    enviadas     = $enviadas
    respondidas  = $respondidas.Count
    promotores   = $totalP
    neutros      = $totalN
    detractores  = $totalD
    nps          = $npsGeneral
    daily        = $daily
    weekly       = $weekly
    tipo         = $tipo
    tecnicos     = $tecnicos
    highlights   = [ordered]@{ top=$topHighlights; bottom=$bottomHighlights }
    conclusiones = $conclusiones
    generadoEl   = (Get-Date -Format "yyyy-MM-dd HH:mm")
}

$json = $data | ConvertTo-Json -Depth 6 -Compress
if ($json -match '</script') { Write-Host "ERROR: los datos contienen una secuencia insegura, abortando." -ForegroundColor Red; exit 1 }

if (-not (Test-Path $TemplatePath)) { Write-Host "ERROR: no se encontro template.html en $RepoDir" -ForegroundColor Red; exit 1 }
$html = Get-Content $TemplatePath -Raw -Encoding UTF8
$final = $html.Replace('/*__DATA__*/', $json)
[System.IO.File]::WriteAllText($OutputPath, $final, (New-Object System.Text.UTF8Encoding($false)))
Write-Host "==> index.html regenerado ($((Get-Item $OutputPath).Length) bytes)"

# ---------- 9. Commit y push ----------
Push-Location $RepoDir
try {
    git add index.html | Out-Null
    $diff = git diff --cached --name-only
    if (-not $diff) {
        Write-Host "==> No hay cambios respecto a la ultima publicacion. Nada que subir."
    } else {
        $fecha = Get-Date -Format "yyyy-MM-dd HH:mm"
        git commit -q -m "Actualizar informe NPS ($fecha)"
        if ($LASTEXITCODE -ne 0) {
            Write-Host "ERROR: el commit fallo." -ForegroundColor Red
            exit 1
        }
        git push -q
        if ($LASTEXITCODE -ne 0) {
            Write-Host "ERROR: el push a GitHub fallo. Revisa tu conexion o 'gh auth status'." -ForegroundColor Red
            exit 1
        }
        Write-Host "==> Publicado. El sitio se actualizara en 1-2 minutos en:"
        Write-Host "    https://jhonasvk.github.io/informe-nps/" -ForegroundColor Cyan
    }
} finally {
    Pop-Location
}
