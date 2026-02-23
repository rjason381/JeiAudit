$ErrorActionPreference = "Stop"

$template = "MCSettings_Sanitarias_R2024.xml"
$out = "MCSettings_General_R2024.xml"

$categoryRows = @(
    "AdministraciÃƒÆ’Ã‚Â³n de vibraciÃƒÆ’Ã‚Â³n|OST_VibrationManagement"
    "Aislantes de vibraciones|OST_VibrationIsolators"
    "Amortiguadores de vibraciones|OST_VibrationDampers"
    "Aparatos elÃƒÆ’Ã‚Â©ctricos|OST_ElectricalFixtures"
    "Aparatos sanitarios|OST_PlumbingFixtures"
    "Arcos|OST_BridgeArches"
    "Armadura estructural|OST_Rebar"
    "ArmazÃƒÆ’Ã‚Â³n estructural|OST_StructuralFraming"
    "Balastres|OST_StairsRailingBaluster"
    "Barandales superiores|OST_RailingTopRail"
    "Barandillas|OST_StairsRailing"
    "Barridos de muro|OST_Cornices"
    "Bordes de losa|OST_EdgeSlab"
    "Cables de acero estructurales|OST_StructuralTendons"
    "Cables de puente|OST_BridgeCables"
    "Canalones|OST_Gutter"
    "Carreteras|OST_Roads"
    "Cielos rasos de cubierta|OST_RoofSoffit"
    "Cimentaciones de contrafuerte|OST_AbutmentFoundations"
    "Cimentaciones de pila|OST_BridgeFoundations"
    "CimentaciÃƒÆ’Ã‚Â³n estructural|OST_StructuralFoundation"
    "CirculaciÃƒÆ’Ã‚Â³n vertical|OST_VerticalCirculation"
    "Claraboya de masa|OST_MassSkylights"
    "Conexiones estructurales|OST_StructConnections"
    "Contrafuertes|OST_BridgeAbutments"
    "Cristalera de masa|OST_MassGlazing"
    "Cubierta de masa|OST_MassRoof"
    "Cubiertas|OST_Roofs"
    "Descansillos|OST_StairsLandings"
    "Diafragmas|OST_BridgeFramingDiaphragms"
    "Dispositivos audiovisuales|OST_AudioVisualDevices"
    "Dispositivos de alarma de incendios|OST_FireAlarmDevices"
    "Dispositivos de comunicaciÃƒÆ’Ã‚Â³n|OST_CommunicationDevices"
    "Dispositivos de control mecÃƒÆ’Ã‚Â¡nico|OST_MechanicalControlDevices"
    "Dispositivos de datos|OST_DataDevices"
    "Dispositivos de iluminaciÃƒÆ’Ã‚Â³n|OST_LightingDevices"
    "Dispositivos de seguridad|OST_SecurityDevices"
    "Dispositivos telefÃƒÆ’Ã‚Â³nicos|OST_TelephoneDevices"
    "Emplazamiento|OST_Site"
    "Entorno|OST_Entourage"
    "Equipo de fontanerÃƒÆ’Ã‚Â­a|OST_PlumbingEquipment"
    "Equipo de servicios alimentarios|OST_FoodServiceEquipment"
    "Equipo mÃƒÆ’Ã‚Â©dico|OST_MedicalEquipment"
    "Equipos elÃƒÆ’Ã‚Â©ctricos|OST_ElectricalEquipment"
    "Equipos especializados|OST_SpecialityEquipment"
    "Equipos mecÃƒÆ’Ã‚Â¡nicos|OST_MechanicalEquipment"
    "Escaleras|OST_Stairs"
    "Estructura de puente|OST_BridgeFraming"
    "Estructuras temporales|OST_TemporaryStructure"
    "Forma de armadura|OST_RebarShape"
    "Hueco de masa|OST_MassOpening"
    "Impostas|OST_Fascia"
    "Juntas de expansiÃƒÆ’Ã‚Â³n|OST_ExpansionJoints"
    "JÃƒÆ’Ã‚Â¡cenas|OST_BridgeGirders"
    "Losas de aproximaciÃƒÆ’Ã‚Â³n|OST_ApproachSlabs"
    "Luminarias|OST_LightingFixtures"
    "LÃƒÆ’Ã‚Â­neas de propiedad|OST_SiteProperty"
    "Mallazo de refuerzo estructural|OST_FabricReinforcement"
    "Masa|OST_Mass"
    "Mobiliario|OST_Furniture"
    "Modelos genÃƒÆ’Ã‚Â©ricos|OST_GenericModel"
    "Montantes de muro cortina|OST_CurtainWallMullions"
    "Muebles de obra|OST_Casework"
    "Muro exterior de masa|OST_MassExteriorWall"
    "Muro interior de masa|OST_MassInteriorWall"
    "Muros|OST_Walls"
    "Muros de contrafuerte|OST_AbutmentWalls"
    "Muros de pila|OST_PierWalls"
    "Paneles de muro cortina|OST_CurtainWallPanels"
    "Pasamanos|OST_RailingHandRail"
    "Pavimento|OST_Hardscape"
    "Piezas|OST_Parts"
    "Pilares|OST_Columns"
    "Pilares de pila|OST_PierColumns"
    "Pilares estructurales|OST_StructuralColumns"
    "Pilas|OST_BridgePiers"
    "Pilotes de contrafuerte|OST_AbutmentPiles"
    "Pilotes de pila|OST_PierPiles"
    "Portantes|OST_BridgeBearings"
    "ProtecciÃƒÆ’Ã‚Â³n contra incendios|OST_FireProtection"
    "Puertas|OST_Doors"
    "Rampas|OST_Ramps"
    "Refuerzo de ÃƒÆ’Ã‚Â¡rea estructural|OST_AreaRein"
    "Refuerzo estructural por camino|OST_PathRein"
    "Remates de pila|OST_PierCaps"
    "Rigidizadores estructurales|OST_StructuralStiffener"
    "Rociadores|OST_Sprinklers"
    "SeÃƒÆ’Ã‚Â±alizaciÃƒÆ’Ã‚Â³n|OST_Signage"
    "Sistemas de mobiliario|OST_FurnitureSystems"
    "Sistemas de muro cortina|OST_CurtaSystem"
    "Soportes|OST_RailingSupport"
    "Suelo de masa|OST_MassFloor"
    "Suelos|OST_Floors"
    "SÃƒÆ’Ã‚Â³lido topogrÃƒÆ’Ã‚Â¡fico|OST_Toposolid"
    "Tableros de puente|OST_BridgeDecks"
    "Techos|OST_Ceilings"
    "Terminaciones|OST_RailingTermination"
    "Timbres de enfermerÃƒÆ’Ã‚Â­a|OST_NurseCallDevices"
    "TopografÃƒÆ’Ã‚Â­a|OST_Topography"
    "Tornapunta transversal|OST_BridgeFramingCrossBracing"
    "Torres de pila|OST_BridgeTowers"
    "Tramos|OST_StairsRuns"
    "Ventanas|OST_Windows"
    "Vigas de celosÃƒÆ’Ã‚Â­a|OST_BridgeFramingTrusses"
    "Zona de masa|OST_MassZone"
)

$categories = foreach ($row in $categoryRows) {
    $parts = $row.Split("|")
    if ($parts.Count -ne 2) {
        throw "Fila invalida: $row"
    }

    [pscustomobject]@{
        Name = $parts[0].Trim()
        Ost = $parts[1].Trim()
    }
}

$allOst = Get-Content "all_ost_categories_2024.txt" | Where-Object { $_ -match "^OST_" }
$missing = $categories | Where-Object { $allOst -notcontains $_.Ost }
if ($missing.Count -gt 0) {
    Write-Host "OST no encontrados en all_ost_categories_2024.txt:" -ForegroundColor Red
    $missing | ForEach-Object { Write-Host (" - {0} | {1}" -f $_.Name, $_.Ost) -ForegroundColor Red }
    throw "Hay OST no validos en la lista."
}

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

[xml]$xml = Get-Content $template

$xml.MCSettings.Name = "PEB EIMI UP Model Checker R2024 (ES) - General"
$xml.MCSettings.Description = "Reglas basadas en PEB EIMI-AITEC-DS-000-ZZZ-PEB-ZZZ-ZZZ-001 - version categorias General"
$xml.MCSettings.Heading.HeadingText = "PEB_EIMI_UP_GENERAL"
$xml.MCSettings.Heading.Description = "Auditoria automatica basada en secciones 4.10.9 y 4.11 del PEB (General)"

$categories = $categories | ForEach-Object {
    [pscustomobject]@{
        Name = (Remove-Diacritics (Fix-Mojibake $_.Name)).Trim()
        Ost = $_.Ost
    }
}

$sec00 = $xml.MCSettings.Heading.Section | Where-Object { $_.SectionName -eq "00_Generales_Modelos" }
$sec03 = $xml.MCSettings.Heading.Section | Where-Object { $_.SectionName -eq "03_Metrados_Modelos" }
$sec04 = $xml.MCSettings.Heading.Section | Where-Object { $_.SectionName -eq "04_Nomenclatura_Parametros_PEB" }
$sec10 = $xml.MCSettings.Heading.Section | Where-Object { $_.SectionName -eq "10_Nomenclatura_Planos" }

$sec00.Description = "PEB 4.11.3: parametros generales obligatorios en categorias General seleccionadas"
$sec03.Description = "PEB 4.11.3: parametros de metrados obligatorios en categorias General seleccionadas"

@($sec00.SelectNodes("Check")) | ForEach-Object { [void]$sec00.RemoveChild($_) }
@($sec03.SelectNodes("Check")) | ForEach-Object { [void]$sec03.RemoveChild($_) }
if ($sec04 -ne $null) {
    @($sec04.Check | Where-Object {
            [string]$_.CheckName -like "*ARG_ESPECIALIDAD*" -or
            [string]$_.Description -like "*ARG_ESPECIALIDAD*" -or
            [string]$_.FailureMessage -like "*ARG_ESPECIALIDAD*"
        }) | ForEach-Object { [void]$sec04.RemoveChild($_) }
}

if ($sec10 -ne $null) {
    $sheetRegex = '^EIMI-[^-]+-[^-]+-[^-]+-[^-]+-PLN-[^-]+-[^-]+-[^-]+$'
    @($sec10.Check.Filter | Where-Object { $_.Property -eq "SheetNameRegex" }) | ForEach-Object {
        $_.Value = $sheetRegex
        $_.CaseInsensitive = "False"
    }
}

$params00 = @("ARG_SECTOR", "ARG_NIVEL", "ARG_SISTEMA")
$paramCodigo = "ARG_C$([char]0x00D3)DIGO DE PARTIDA"
$paramDescripcion = "ARG_DESCRIPCI$([char]0x00D3)N DE PARTIDA"
$params03 = @("ARG_UNIFORMAT", $paramCodigo, "ARG_UNIDAD DE PARTIDA", $paramDescripcion)
function Add-Check([xml]$doc, $section, [string]$catName, [string]$ost, [string]$param)
{
    $check = $doc.CreateElement("Check")
    $check.SetAttribute("ID", [guid]::NewGuid().ToString())
    $check.SetAttribute("CheckName", ("{0}_{1}_NoValue" -f (Slug $catName), (Slug $param)))
    $check.SetAttribute("Description", ("PEB 4.11.3: parametro obligatorio '{0}' sin valor en categoria '{1}'" -f $param, $catName))
    $check.SetAttribute("FailureMessage", ("Hay elementos en '{0}' con '{1}' sin valor" -f $catName, $param))
    $check.SetAttribute("ResultCondition", "FailMatchingElements")
    $check.SetAttribute("CheckType", "Custom")
    $check.SetAttribute("IsRequired", "True")
    $check.SetAttribute("IsChecked", "True")

    $f1 = $doc.CreateElement("Filter")
    $f1.SetAttribute("ID", [guid]::NewGuid().ToString())
    $f1.SetAttribute("Operator", "And")
    $f1.SetAttribute("Category", "Category")
    $f1.SetAttribute("Property", $ost)
    $f1.SetAttribute("Condition", "Included")
    $f1.SetAttribute("Value", "True")
    $f1.SetAttribute("CaseInsensitive", "False")
    $f1.SetAttribute("Unit", "None")
    $f1.SetAttribute("UnitClass", "None")
    $f1.SetAttribute("FieldTitle", "")
    $f1.SetAttribute("UserDefined", "False")
    $f1.SetAttribute("Validation", "None")

    $f2 = $doc.CreateElement("Filter")
    $f2.SetAttribute("ID", [guid]::NewGuid().ToString())
    $f2.SetAttribute("Operator", "And")
    $f2.SetAttribute("Category", "Parameter")
    $f2.SetAttribute("Property", $param)
    $f2.SetAttribute("Condition", "HasNoValue")
    $f2.SetAttribute("Value", "True")
    $f2.SetAttribute("CaseInsensitive", "False")
    $f2.SetAttribute("Unit", "None")
    $f2.SetAttribute("UnitClass", "None")
    $f2.SetAttribute("FieldTitle", "")
    $f2.SetAttribute("UserDefined", "False")
    $f2.SetAttribute("Validation", "None")

    $f3 = $doc.CreateElement("Filter")
    $f3.SetAttribute("ID", [guid]::NewGuid().ToString())
    $f3.SetAttribute("Operator", "And")
    $f3.SetAttribute("Category", "TypeOrInstance")
    $f3.SetAttribute("Property", "Is Element Type")
    $f3.SetAttribute("Condition", "Equal")
    $f3.SetAttribute("Value", "False")
    $f3.SetAttribute("CaseInsensitive", "False")
    $f3.SetAttribute("Unit", "None")
    $f3.SetAttribute("UnitClass", "None")
    $f3.SetAttribute("FieldTitle", "")
    $f3.SetAttribute("UserDefined", "False")
    $f3.SetAttribute("Validation", "None")

    [void]$check.AppendChild($f1)
    [void]$check.AppendChild($f2)
    [void]$check.AppendChild($f3)
    [void]$section.AppendChild($check)
}

foreach ($c in $categories) {
    foreach ($p in $params00) {
        Add-Check -doc $xml -section $sec00 -catName $c.Name -ost $c.Ost -param $p
    }
}

foreach ($c in $categories) {
    foreach ($p in $params03) {
        Add-Check -doc $xml -section $sec03 -catName $c.Name -ost $c.Ost -param $p
    }
}

$settings = New-Object System.Xml.XmlWriterSettings
$settings.Indent = $true
$settings.Encoding = New-Object System.Text.UTF8Encoding($false)
$writer = [System.Xml.XmlWriter]::Create((Join-Path (Resolve-Path ".") $out), $settings)
$xml.Save($writer)
$writer.Close()

[xml]$verify = Get-Content $out
$s00c = ($verify.MCSettings.Heading.Section | Where-Object { $_.SectionName -eq "00_Generales_Modelos" }).Check.Count
$s03c = ($verify.MCSettings.Heading.Section | Where-Object { $_.SectionName -eq "03_Metrados_Modelos" }).Check.Count

Write-Host "created=$out"
Write-Host "categories=$($categories.Count)"
Write-Host "checks00=$s00c"
Write-Host "checks03=$s03c"
Write-Host "totalCustomChecks=$($s00c + $s03c)"

