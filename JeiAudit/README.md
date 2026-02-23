# JeiAudit

Revit 2024 plugin that runs:

1. **Parameter existence audit** by category (independent of value).
2. **Family naming audit** using the PEB pattern (`ARG_XXX_NOMBRE`).
3. **Subproject/workset naming audit** using the PEB prefix rule (`ARG_...`).
4. **Model Checker XML audit** by reading `MCSettings*.xml` and evaluating compatible checks.

## What it checks

- Required parameters:
  - `ARG_SECTOR`
  - `ARG_NIVEL`
  - `ARG_SISTEMA`
  - `ARG_UNIFORMAT`
  - `ARG_CÓDIGO DE PARTIDA`
  - `ARG_UNIDAD DE PARTIDA`
  - `ARG_DESCRIPCIÓN DE PARTIDA`
- Categories:
  - `Muros`
  - `Puertas`
  - `Ventanas`
  - `Suelos`
  - `Habitaciones`
- Family naming:
  - Required prefix: `ARG_`
  - Required discipline code: `ARQ|EST|SAN|ELEC|ILUM|COM|MEC|ACI|DACI|SEG|EQP`
  - Required structure: `ARG_XXX_NOMBRE_FAMILIA`
- Subproject/workset naming:
  - Required prefix: `ARG_`
  - Required content after prefix
  - Warns when no second separator exists (recommended structure check)

## Output

- Folder on desktop: `JeiAudit_yyyyMMdd_HHmmss`
- Files:
  - `JeiAudit_Results.xlsx` (principal para dashboards/Power BI; incluye hojas por cada auditoría)
    - `00_Meta`
    - `01_Parameter_Existence`
    - `02_Family_Naming`
    - `03_Subproject_Naming`
    - `04_MC_Summary`
    - `05_MC_Failed_Elements`
    - `10_KPI_Summary` (KPIs consolidados)
    - `11_Fact_QualityChecks` (tabla limpia consolidada)
    - `12_Fact_FailedElements` (fallas por elemento)
    - `13_Pivot_By_Section` (base pivot por sección)
    - `14_Pivot_By_Category` (base pivot por categoría)
    - `15_Pivot_Top_FailChecks` (ranking de checks con más fallas)
  - `01_parameter_existence.csv`
  - `02_family_naming.csv`
  - `03_subproject_naming.csv`
  - `04_modelchecker_summary.csv` (if XML was selected)
  - `05_modelchecker_failed_elements.csv` (if XML was selected)

## UI workflow (Model Checker style)

- `Configuracion`:
  - Opens a checkset editor window.
  - Reads XML metadata (`Name`, `Date`, `Author`, `Description`).
  - Shows `Heading > Section > Check` tree with checkboxes.
  - Saves `IsChecked=True/False` back into the same XML.
- `Ejecutar`:
  - Runs JeiAudit checks + XML checks.
  - Produces Excel and CSV outputs.
- `Ver Reporte`:
  - Opens report viewer window (summary + tree).
  - Bottom actions: `Copia`, `Html`, `Excel`, `AVT`, `Cerrar`.

`Model Checker XML` support in this version:
- `ResultCondition`: `FailMatchingElements`, `FailNoMatchingElements`, `FailNoElements`, `CountOnly`, `CountAndList`
- `Operator`: `And`, `Or`, `Exclude`, `AndNot`, `OrNot`
- `Category` filter conditions:
  - `Included`, `Excluded`, `Equal`, `NotEqual`
- `TypeOrInstance` filter:
  - Property aliases: `Is Element Type`, `Is Type`
  - Conditions: `Equal`, `NotEqual`, `Included`, `Excluded`
- Extra filter categories:
  - `Family`, `Type`, `Workset`, `APIType`
  - `APIParameter` (alias of `Parameter`)
  - `Level`, `PhaseCreated`, `PhaseDemolished`, `PhaseStatus`
  - `DesignOption`, `View`, `StructuralType`
  - `Host`, `HostParameter`
  - `Room`, `Space` (best-effort)
- `Parameter` filter conditions:
  - `HasNoValue`, `HasValue`, `Exists`, `DoesNotExist`, `Defined`, `Undefined`
  - `Equal`, `NotEqual`, `Unequal`, `Included`
  - `Contains`, `NotContains`, `StartsWith`, `NotStartsWith`, `EndsWith`, `NotEndsWith`
  - `MatchesRegex`, `NotMatchesRegex`, `WildCard`, `WildCardNoMatch`
  - `MatchesParameter`, `DoesNotMatchParameter`
  - `MatchesHostParameter`, `DoesNotMatchHostParameter`
  - `Duplicated`
  - `InList`, `NotInList` (split by `;`, `,`, `|`)
  - `GreaterThan`, `GreaterOrEqual`, `LessThan`, `LessOrEqual`
  - `IsTrue`, `IsFalse`
- Numeric comparisons support `Unit` and `UnitClass` from XML (length/area/volume/angle conversion to Revit internal units).
- Built-in property support (without custom parameter binding):
  - `Family Name`, `Type Name`, `Workset/Subproject`, `Category`, `Element Id`, `Name`

Example (family/subproject naming directly from XML):

```xml
<Check CheckName="FamilyPrefix_ARG" ResultCondition="FailMatchingElements" IsChecked="True">
  <Filter Operator="And" Category="Parameter" Property="Family Name" Condition="NotStartsWith" Value="ARG_" CaseInsensitive="False" />
</Check>

<Check CheckName="SubprojectPrefix_ARG" ResultCondition="FailMatchingElements" IsChecked="True">
  <Filter Operator="And" Category="Parameter" Property="Subproject" Condition="NotStartsWith" Value="ARG_" CaseInsensitive="False" />
</Check>
```

Tip for strict parameter-name validation:
- Use `Category="Parameter"` with `Condition="DoesNotExist"` to flag elements where the exact parameter name is missing.

Unsupported XML patterns are still reported as `UNSUPPORTED` in summary for traceability.

## Install

1. Open PowerShell.
2. Run:

```powershell
cd "c:\Users\user\OneDrive\Documentos\Programacion\BIM\Jason\IT\JeiAudit\tools"
.\Install-JeiAudit.ps1
```

3. Restart Revit 2024.
4. In ribbon, open tab `JeiAudit` > panel `Model Checker`.
5. Click `Setup` once to select your `MCSettings*.xml`.
6. Click `Run` to execute audits.
7. Click `View Report` to open the latest output folder.

## Uninstall

```powershell
cd "c:\Users\user\OneDrive\Documentos\Programacion\BIM\Jason\IT\JeiAudit\tools"
.\Uninstall-JeiAudit.ps1
```
