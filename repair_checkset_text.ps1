$ErrorActionPreference = "Stop"

function Remove-Diacritics([string]$text)
{
    if ([string]::IsNullOrWhiteSpace($text)) {
        return ""
    }

    $norm = $text.Normalize([Text.NormalizationForm]::FormD)
    $sb = New-Object Text.StringBuilder
    foreach ($ch in $norm.ToCharArray()) {
        $uc = [Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch)
        if ($uc -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$sb.Append($ch)
        }
    }

    return $sb.ToString()
}

function Get-MojibakeScore([string]$text)
{
    if ([string]::IsNullOrEmpty($text)) {
        return 0
    }

    return ([regex]::Matches($text, "Ã|Â|Æ|â|�")).Count
}

function Fix-Mojibake([string]$text)
{
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $text
    }

    $enc1252 = [Text.Encoding]::GetEncoding(1252)
    $current = $text

    for ($i = 0; $i -lt 5; $i++) {
        try {
            $candidate = [Text.Encoding]::UTF8.GetString($enc1252.GetBytes($current))
        }
        catch {
            break
        }

        if ($candidate -eq $current) {
            break
        }

        $currentScore = Get-MojibakeScore $current
        $candidateScore = Get-MojibakeScore $candidate

        if ($candidateScore -lt $currentScore) {
            $current = $candidate
        }
        else {
            break
        }
    }

    return $current
}

function Slug([string]$text)
{
    $s = Remove-Diacritics $text
    $s = $s -replace "[^A-Za-z0-9]+", "_"
    $s = $s.Trim("_")
    return $s
}

function Normalize-ParamKey([string]$text)
{
    $s = Fix-Mojibake $text
    $s = Remove-Diacritics $s
    $s = $s.ToUpperInvariant()
    $s = $s -replace "[^A-Z0-9]", ""
    return $s
}

$paramCodigoPartida = "ARG_C$([char]0x00D3)DIGO DE PARTIDA"
$paramDescripcionPartida = "ARG_DESCRIPCI$([char]0x00D3)N DE PARTIDA"
$paramCodigoPlano = "ARG_C$([char]0x00D3)DIGO DE PLANO"

function Canonical-ParameterName([string]$value)
{
    if ([string]::IsNullOrWhiteSpace($value)) {
        return $value
    }

    $key = Normalize-ParamKey $value

    if ($key -match "ARGSECTOR") { return "ARG_SECTOR" }
    if ($key -match "ARGNIVEL") { return "ARG_NIVEL" }
    if ($key -match "ARGSISTEMA") { return "ARG_SISTEMA" }
    if ($key -match "ARGUNIFORMAT") { return "ARG_UNIFORMAT" }
    if ($key -match "ARGESPECIALIDAD") { return "ARG_ESPECIALIDAD" }
    if ($key -match "ARGUNIDADDEPARTIDA") { return "ARG_UNIDAD DE PARTIDA" }
    if ($key -match "ARGDESCRIP" -and $key -match "PARTIDA") { return $paramDescripcionPartida }
    if ($key -match "ARGC" -and $key -match "DIGO" -and $key -match "PARTIDA") { return $paramCodigoPartida }
    if ($key -match "ARGC" -and $key -match "DIGO" -and $key -match "PLANO") { return $paramCodigoPlano }
    if ($key -match "ARGCOLABORADORES") { return "ARG_COLABORADORES" }
    if ($key -match "ARGNUMEROCOLEGIATURA") { return "ARG_NUMERO COLEGIATURA" }
    if ($key -match "ARGPROYECTISTARESPONSABLE") { return "ARG_PROYECTISTA RESPONSABLE" }
    if ($key -eq "NUMERO") { return "Numero" }
    if ($key -eq "NOMBRE") { return "Nombre" }

    return (Fix-Mojibake $value)
}

function Get-CategoryDisplayMap()
{
    $map = @{}

    if (Test-Path -LiteralPath "generate_general_checkset.ps1") {
        Get-Content -LiteralPath "generate_general_checkset.ps1" -Encoding UTF8 | ForEach-Object {
            if ($_ -match '^\s*"(.+)\|(OST_[A-Za-z0-9_]+)"\s*$') {
                $name = Remove-Diacritics (Fix-Mojibake $matches[1])
                $ost = $matches[2]
                $map[$ost] = $name
            }
        }
    }

    $extra = @{
        "Habitaciones" = "Habitaciones"
        "OST_ArcWallRectOpening" = "Hueco de muro en arco rectangular"
        "OST_IOSModelGroups" = "Grupos de modelo"
        "OST_DuctCurves" = "Conductos"
        "OST_DuctFitting" = "Uniones de conductos"
        "OST_DuctAccessory" = "Accesorios de conductos"
        "OST_DuctTerminal" = "Terminales de aire"
        "OST_FlexDuctCurves" = "Conductos flexibles"
        "OST_PipeCurves" = "Tuberias"
        "OST_PipeFitting" = "Uniones de tuberia"
        "OST_PipeAccessory" = "Accesorios de tuberia"
        "OST_FlexPipeCurves" = "Tuberias flexibles"
        "OST_Conduit" = "Tubos"
        "OST_ConduitFitting" = "Uniones de tubo"
        "OST_CableTray" = "Bandejas de cables"
        "OST_CableTrayFitting" = "Uniones de bandeja de cables"
        "OST_ElectricalFixtures" = "Aparatos electricos"
        "OST_ElectricalEquipment" = "Equipos electricos"
        "OST_LightingDevices" = "Dispositivos de iluminacion"
        "OST_PlumbingFixtures" = "Aparatos sanitarios"
        "OST_PlumbingEquipment" = "Equipo de fontaneria"
        "OST_MechanicalEquipment" = "Equipos mecanicos"
        "OST_StructuralFramingOther" = "Armazon estructural"
    }

    foreach ($k in $extra.Keys) {
        $map[$k] = $extra[$k]
    }

    return $map
}

function Resolve-CategoryDisplay([hashtable]$categoryMap, [string]$categoryCode)
{
    if ([string]::IsNullOrWhiteSpace($categoryCode)) {
        return ""
    }

    if ($categoryMap.ContainsKey($categoryCode)) {
        return $categoryMap[$categoryCode]
    }

    if ($categoryCode -like "OST_*") {
        $tail = $categoryCode.Substring(4)
        $tail = $tail -replace "([a-z])([A-Z])", '$1 $2'
        $tail = $tail -replace "_", " "
        return (Remove-Diacritics (Fix-Mojibake $tail)).Trim()
    }

    return (Remove-Diacritics (Fix-Mojibake $categoryCode)).Trim()
}

$files = @(
    "MCSettings_Basico_R2024.xml",
    "MCSettings_Estructuras_R2024.xml",
    "MCSettings_Electricas_R2024.xml",
    "MCSettings_Sanitarias_R2024.xml",
    "MCSettings_MEP_R2024.xml",
    "MCSettings_General_R2024.xml",
    "MCSettings_General_TodasCategorias_R2024.xml"
) | Where-Object { Test-Path -LiteralPath $_ }

$categoryMap = Get-CategoryDisplayMap

foreach ($file in $files) {
    [xml]$xml = Get-Content -LiteralPath $file -Raw -Encoding UTF8
    $changed = $false

    foreach ($check in $xml.SelectNodes("//Check")) {
        $catFilter = @($check.Filter | Where-Object { $_.Category -eq "Category" } | Select-Object -First 1)[0]
        $paramFilter = @($check.Filter | Where-Object { $_.Category -eq "Parameter" } | Select-Object -First 1)[0]
        $paramNameFilter = @($check.Filter | Where-Object { $_.Category -eq "Custom" -and $_.Property -eq "ParameterName" } | Select-Object -First 1)[0]

        $catName = ""
        if ($catFilter -ne $null) {
            $catName = Resolve-CategoryDisplay -categoryMap $categoryMap -categoryCode $catFilter.Property
        }

        $paramName = ""
        if ($paramFilter -ne $null) {
            $newParam = Canonical-ParameterName $paramFilter.Property
            if ($newParam -ne $paramFilter.Property) {
                $paramFilter.Property = $newParam
                $changed = $true
            }
            $paramName = $newParam
        }

        if ($paramNameFilter -ne $null) {
            $newParam = Canonical-ParameterName $paramNameFilter.Value
            if ($newParam -ne $paramNameFilter.Value) {
                $paramNameFilter.Value = $newParam
                $changed = $true
            }
            if ([string]::IsNullOrWhiteSpace($paramName)) {
                $paramName = $newParam
            }
        }

        if ($check.CheckType -eq "Custom" -and $catFilter -ne $null -and $paramFilter -ne $null -and ([string]$check.CheckName -like "*NoValue*")) {
            $newCheckName = ("{0}_{1}_NoValue" -f (Slug $catName), (Slug $paramName))
            $newDesc = ("PEB 4.11.3: parametro obligatorio '{0}' sin valor en categoria '{1}'" -f $paramName, $catName)
            $newFail = ("Hay elementos en '{0}' con '{1}' sin valor" -f $catName, $paramName)

            if ($check.CheckName -ne $newCheckName) { $check.CheckName = $newCheckName; $changed = $true }
            if ($check.Description -ne $newDesc) { $check.Description = $newDesc; $changed = $true }
            if ($check.FailureMessage -ne $newFail) { $check.FailureMessage = $newFail; $changed = $true }
        }

        if ($check.CheckType -eq "JeiAuditParameterExistence" -and -not [string]::IsNullOrWhiteSpace($paramName)) {
            $newCheckName = ("JeiAudit_Parameter_Existence_{0}" -f (Slug $paramName))
            $newDesc = ("JeiAudit: valida existencia exacta del parametro '{0}' segun PEB" -f $paramName)
            $newFail = ("No se encontro parametro '{0}' en el modelo." -f $paramName)

            if ($check.CheckName -ne $newCheckName) { $check.CheckName = $newCheckName; $changed = $true }
            if ($check.Description -ne $newDesc) { $check.Description = $newDesc; $changed = $true }
            if ($check.FailureMessage -ne $newFail) { $check.FailureMessage = $newFail; $changed = $true }
        }
    }

    if ($changed) {
        $settings = New-Object System.Xml.XmlWriterSettings
        $settings.Indent = $true
        $settings.Encoding = New-Object System.Text.UTF8Encoding($false)
        $writer = [System.Xml.XmlWriter]::Create((Join-Path (Resolve-Path ".") $file), $settings)
        $xml.Save($writer)
        $writer.Close()
        Write-Host "UPDATED $file"
    }
    else {
        Write-Host "NOCHANGE $file"
    }
}
