using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ClosedXML.Excel;
using WinForms = System.Windows.Forms;

namespace JeiAudit
{
    [Transaction(TransactionMode.ReadOnly)]
    public class RunAuditCommand : IExternalCommand
    {
        private const int MaxFailedElementsPerCheckExport = 200;
        private static readonly Encoding Latin1252Encoding = Encoding.GetEncoding(1252);

        private sealed class CategoryTarget
        {
            public CategoryTarget(string label, BuiltInCategory builtInCategory)
            {
                Label = label;
                BuiltInCategory = builtInCategory;
            }

            public string Label { get; }
            public BuiltInCategory BuiltInCategory { get; }
        }

        private sealed class BindingInfo
        {
            public BindingInfo(string bindingType, HashSet<long> categoryIds, string guid)
            {
                BindingType = bindingType;
                CategoryIds = categoryIds;
                Guid = guid;
            }

            public string BindingType { get; }
            public HashSet<long> CategoryIds { get; }
            public string Guid { get; }
        }

        private sealed class ChecksetMetadata
        {
            public string Title { get; set; } = "-";
            public string Date { get; set; } = "-";
            public string Author { get; set; } = "-";
            public string Description { get; set; } = "-";
        }

        private sealed class ParameterAuditRow
        {
            public string ParameterName { get; set; } = string.Empty;
            public string CategoryName { get; set; } = string.Empty;
            public string CategoryAvailable { get; set; } = string.Empty;
            public string ExistsByExactName { get; set; } = string.Empty;
            public string BoundToCategory { get; set; } = string.Empty;
            public string BindingTypes { get; set; } = string.Empty;
            public string ParameterGuids { get; set; } = string.Empty;
            public string SimilarNames { get; set; } = string.Empty;
            public string Status { get; set; } = string.Empty;
        }

        private sealed class ModelCheckerFilterDef
        {
            public string Operator { get; set; } = string.Empty;
            public string FilterCategory { get; set; } = string.Empty;
            public string Property { get; set; } = string.Empty;
            public string Condition { get; set; } = string.Empty;
            public string Value { get; set; } = string.Empty;
            public bool CaseInsensitive { get; set; }
            public string Unit { get; set; } = string.Empty;
            public string UnitClass { get; set; } = string.Empty;
            public string Validation { get; set; } = string.Empty;
        }

        private sealed class ModelCheckerCheckDef
        {
            public string Heading { get; set; } = string.Empty;
            public string Section { get; set; } = string.Empty;
            public string Name { get; set; } = string.Empty;
            public string Description { get; set; } = string.Empty;
            public string FailureMessage { get; set; } = string.Empty;
            public string ResultCondition { get; set; } = string.Empty;
            public string CheckType { get; set; } = string.Empty;
            public bool IsChecked { get; set; }
            public List<ModelCheckerFilterDef> Filters { get; } = new List<ModelCheckerFilterDef>();
        }

        private sealed class ModelCheckerSummaryRow
        {
            public string Heading { get; set; } = string.Empty;
            public string Section { get; set; } = string.Empty;
            public string CheckName { get; set; } = string.Empty;
            public string Status { get; set; } = string.Empty;
            public int CandidateElements { get; set; }
            public int FailedElements { get; set; }
            public string Reason { get; set; } = string.Empty;
        }

        private sealed class ModelCheckerFailureRow
        {
            public string Heading { get; set; } = string.Empty;
            public string Section { get; set; } = string.Empty;
            public string CheckName { get; set; } = string.Empty;
            public long ElementId { get; set; }
            public string Category { get; set; } = string.Empty;
            public string FamilyOrType { get; set; } = string.Empty;
            public string Name { get; set; } = string.Empty;
        }

        private sealed class FamilyNamingRow
        {
            public string FamilyName { get; set; } = string.Empty;
            public string HasArgPrefix { get; set; } = string.Empty;
            public string DisciplineCode { get; set; } = string.Empty;
            public string DisciplineCodeValid { get; set; } = string.Empty;
            public string HasNamePart { get; set; } = string.Empty;
            public string Status { get; set; } = string.Empty;
            public string Reason { get; set; } = string.Empty;
        }

        private sealed class SubprojectNamingRow
        {
            public string WorksetName { get; set; } = string.Empty;
            public string HasArgPrefix { get; set; } = string.Empty;
            public string HasNameAfterPrefix { get; set; } = string.Empty;
            public string HasSecondSeparator { get; set; } = string.Empty;
            public string Status { get; set; } = string.Empty;
            public string Reason { get; set; } = string.Empty;
        }

        private sealed class QualityCheckFactRow
        {
            public string AuditType { get; set; } = string.Empty;
            public string Heading { get; set; } = string.Empty;
            public string Section { get; set; } = string.Empty;
            public string RuleName { get; set; } = string.Empty;
            public string EntityType { get; set; } = string.Empty;
            public string EntityName { get; set; } = string.Empty;
            public string Category { get; set; } = string.Empty;
            public string Status { get; set; } = string.Empty;
            public string Severity { get; set; } = string.Empty;
            public int CandidateElements { get; set; }
            public int FailedElements { get; set; }
            public string Reason { get; set; } = string.Empty;
        }

        private static readonly CategoryTarget[] CategoryTargets =
        {
            new CategoryTarget("Muros", BuiltInCategory.OST_Walls),
            new CategoryTarget("Puertas", BuiltInCategory.OST_Doors),
            new CategoryTarget("Ventanas", BuiltInCategory.OST_Windows),
            new CategoryTarget("Suelos", BuiltInCategory.OST_Floors),
            new CategoryTarget("Habitaciones", BuiltInCategory.OST_Rooms)
        };

        private static readonly string[] RequiredParameters =
        {
            "ARG_SECTOR",
            "ARG_NIVEL",
            "ARG_SISTEMA",
            "ARG_UNIFORMAT",
            "ARG_C\u00D3DIGO DE PARTIDA",
            "ARG_UNIDAD DE PARTIDA",
            "ARG_DESCRIPCI\u00D3N DE PARTIDA"
        };

        private static readonly HashSet<string> FamilyDisciplineCodes = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "ARQ", "EST", "SAN", "ELEC", "ILUM", "COM", "MEC", "ACI", "DACI", "SEG", "EQP"
        };

        private static readonly Dictionary<string, string> CategoryAliases = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            { "areas", "OST_Areas" },
            { "habitaciones", "OST_Rooms" },
            { "ceilings", "OST_Ceilings" },
            { "cielorrasos", "OST_Ceilings" },
            { "doors", "OST_Doors" },
            { "puertas", "OST_Doors" },
            { "duct accessories", "OST_DuctAccessory" },
            { "accesorios de conducto", "OST_DuctAccessory" },
            { "duct fittings", "OST_DuctFitting" },
            { "accesorios de conductos", "OST_DuctFitting" },
            { "ducts", "OST_DuctCurves" },
            { "conductos", "OST_DuctCurves" },
            { "electrical equipment", "OST_ElectricalEquipment" },
            { "equipos electricos", "OST_ElectricalEquipment" },
            { "floors", "OST_Floors" },
            { "suelos", "OST_Floors" },
            { "furniture", "OST_Furniture" },
            { "mobiliario", "OST_Furniture" },
            { "generic models", "OST_GenericModel" },
            { "modelos genericos", "OST_GenericModel" },
            { "mechanical equipment", "OST_MechanicalEquipment" },
            { "equipos mecanicos", "OST_MechanicalEquipment" },
            { "model groups", "OST_IOSModelGroups" },
            { "grupos de modelo", "OST_IOSModelGroups" },
            { "parking", "OST_Parking" },
            { "estacionamientos", "OST_Parking" },
            { "pipe accessories", "OST_PipeAccessory" },
            { "accesorios de tuberia", "OST_PipeAccessory" },
            { "pipe fittings", "OST_PipeFitting" },
            { "accesorios de tuberias", "OST_PipeFitting" },
            { "pipes", "OST_PipeCurves" },
            { "tuberias", "OST_PipeCurves" },
            { "planting", "OST_Planting" },
            { "plumbing fixtures", "OST_PlumbingFixtures" },
            { "artefactos sanitarios", "OST_PlumbingFixtures" },
            { "project information", "OST_ProjectInformation" },
            { "informacion de proyecto", "OST_ProjectInformation" },
            { "railings", "OST_Railings" },
            { "barandas", "OST_Railings" },
            { "barandillas", "OST_Railings" },
            { "barandilla", "OST_Railings" },
            { "rampas", "OST_Ramps" },
            { "ramps", "OST_Ramps" },
            { "roofs", "OST_Roofs" },
            { "techos", "OST_Roofs" },
            { "rooms", "OST_Rooms" },
            { "shaft openings", "OST_ShaftOpening" },
            { "huecos", "OST_ShaftOpening" },
            { "specialty equipment", "OST_SpecialityEquipment" },
            { "equipo especial", "OST_SpecialityEquipment" },
            { "sprinklers", "OST_Sprinklers" },
            { "rociadores", "OST_Sprinklers" },
            { "stairs", "OST_Stairs" },
            { "escaleras", "OST_Stairs" },
            { "structural columns", "OST_StructuralColumns" },
            { "columnas estructurales", "OST_StructuralColumns" },
            { "structural foundations", "OST_StructuralFoundation" },
            { "cimentaciones estructurales", "OST_StructuralFoundation" },
            { "structural framing", "OST_StructuralFraming" },
            { "entramado estructural", "OST_StructuralFraming" },
            { "toposolid", "OST_Toposolid" },
            { "walls", "OST_Walls" },
            { "muros", "OST_Walls" },
            { "windows", "OST_Windows" },
            { "ventanas", "OST_Windows" },
            { "sheets", "OST_Sheets" },
            { "planos", "OST_Sheets" },
            { "laminas", "OST_Sheets" }
        };

        private static readonly Dictionary<string, string[]> BuiltInCategoryFallbacks =
            new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase)
            {
                { "OST_Railings", new[] { "OST_StairsRailing", "OST_RailingSystem" } },
                { "OST_StairsRailing", new[] { "OST_Railings", "OST_RailingSystem" } },
                { "Railings", new[] { "OST_Railings", "OST_StairsRailing", "OST_RailingSystem" } },
                { "Barandillas", new[] { "OST_Railings", "OST_StairsRailing", "OST_RailingSystem" } },
                { "Barandilla", new[] { "OST_Railings", "OST_StairsRailing", "OST_RailingSystem" } },
                { "Barandas", new[] { "OST_Railings", "OST_StairsRailing", "OST_RailingSystem" } }
            };

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument? uiDoc = uiApp.ActiveUIDocument;
            if (uiDoc == null)
            {
                TaskDialog.Show("JeiAudit", "No hay un modelo activo para ejecutar la comprobacion.");
                return Result.Cancelled;
            }

            Document doc = uiDoc.Document;

            try
            {
                string modelCheckerXmlPath = ResolveModelCheckerXmlPathForRun(doc);
                if (string.IsNullOrWhiteSpace(modelCheckerXmlPath) || !File.Exists(modelCheckerXmlPath))
                {
                    return Result.Cancelled;
                }

                PluginState.SaveLastXmlPath(modelCheckerXmlPath);

                ChecksetMetadata metadata = LoadChecksetMetadata(modelCheckerXmlPath);
                List<RunAuditModelItem> runModels = BuildRunModelItems(doc);
                List<RunAuditModelItem> selectedModels;

                using (var runForm = new RunAuditSetupForm(
                    new RunAuditChecksetMetadata
                    {
                        Title = metadata.Title,
                        Date = metadata.Date,
                        Author = metadata.Author,
                        Description = metadata.Description
                    },
                    runModels))
                {
                    WinForms.DialogResult runDialogResult = runForm.ShowDialog();
                    if (runDialogResult != WinForms.DialogResult.OK)
                    {
                        return Result.Cancelled;
                    }

                    selectedModels = runForm.GetSelectedModels();
                }

                if (!selectedModels.Any(v => ReferenceEquals(v.Document, doc)))
                {
                    TaskDialog.Show("JeiAudit", "Debes incluir el modelo activo en la lista de ejecucion.");
                    return Result.Cancelled;
                }

                bool hasAdditionalModels = selectedModels.Any(v => !ReferenceEquals(v.Document, doc));
                if (hasAdditionalModels)
                {
                    TaskDialog.Show("JeiAudit", "En esta version, la ejecucion corre sobre el modelo activo. Los enlaces seleccionados se ignoraran por ahora.");
                }

                string outputDir = PrepareOutputFolder();

                Dictionary<string, List<BindingInfo>> bindings = CollectParameterBindings(doc);
                List<ModelCheckerCheckDef> checks = LoadModelCheckerChecks(modelCheckerXmlPath);
                List<string> auditedCategoryNames = CollectAuditedCategoryNamesForParameterAudit(checks);
                List<ParameterAuditRow> parameterRows = RunParameterExistenceAudit(doc, bindings, auditedCategoryNames);
                string parameterReportPath = WriteParameterAuditCsv(outputDir, parameterRows);

                List<FamilyNamingRow> familyRows = RunFamilyNamingAudit(doc);
                string familyReportPath = WriteFamilyNamingCsv(outputDir, familyRows);

                List<SubprojectNamingRow> subprojectRows = RunSubprojectNamingAudit(doc);
                string subprojectReportPath = WriteSubprojectNamingCsv(outputDir, subprojectRows);

                List<ModelCheckerSummaryRow> modelCheckerSummary = new List<ModelCheckerSummaryRow>();
                List<ModelCheckerFailureRow> modelCheckerFailures = new List<ModelCheckerFailureRow>();
                EvaluateModelCheckerChecks(doc, checks, familyRows, subprojectRows, bindings, modelCheckerSummary, modelCheckerFailures);
                string modelCheckerSummaryPath = WriteModelCheckerSummaryCsv(outputDir, modelCheckerSummary, modelCheckerXmlPath);
                string modelCheckerFailuresPath = WriteModelCheckerFailuresCsv(outputDir, modelCheckerFailures);

                string excelReportPath = WriteAuditWorkbook(
                    outputDir,
                    doc,
                    parameterRows,
                    familyRows,
                    subprojectRows,
                    modelCheckerSummary,
                    modelCheckerFailures,
                    modelCheckerXmlPath);

                PluginState.SaveLastOutputFolder(outputDir);

                int parameterMissing = parameterRows.Count(r => string.Equals(r.Status, "MISSING", StringComparison.Ordinal));
                int parameterOk = parameterRows.Count - parameterMissing;

                int familyFail = familyRows.Count(r => string.Equals(r.Status, "FAIL", StringComparison.Ordinal));
                int familyPass = familyRows.Count(r => string.Equals(r.Status, "PASS", StringComparison.Ordinal));
                int familySkipped = familyRows.Count(r => string.Equals(r.Status, "SKIPPED", StringComparison.Ordinal));

                int subprojectFail = subprojectRows.Count(r => string.Equals(r.Status, "FAIL", StringComparison.Ordinal));
                int subprojectWarn = subprojectRows.Count(r => string.Equals(r.Status, "WARN", StringComparison.Ordinal));
                int subprojectPass = subprojectRows.Count(r => string.Equals(r.Status, "PASS", StringComparison.Ordinal));
                int subprojectSkipped = subprojectRows.Count(r => string.Equals(r.Status, "SKIPPED", StringComparison.Ordinal));

                int modelCheckerFail = modelCheckerSummary.Count(r => string.Equals(r.Status, "FAIL", StringComparison.Ordinal));
                int modelCheckerPass = modelCheckerSummary.Count(r => string.Equals(r.Status, "PASS", StringComparison.Ordinal));
                int modelCheckerCount = modelCheckerSummary.Count(r => string.Equals(r.Status, "COUNT", StringComparison.Ordinal));
                int modelCheckerUnsupported = modelCheckerSummary.Count(r => string.Equals(r.Status, "UNSUPPORTED", StringComparison.Ordinal));
                int modelCheckerSkipped = modelCheckerSummary.Count(r => string.Equals(r.Status, "SKIPPED", StringComparison.Ordinal));

                var dialog = new TaskDialog("JeiAudit")
                {
                    MainInstruction = "JeiAudit completed.",
                    MainContent =
                        $"Output folder:{Environment.NewLine}{outputDir}{Environment.NewLine}{Environment.NewLine}" +
                        $"Parameter existence checks: {parameterRows.Count}{Environment.NewLine}" +
                        $" - OK: {parameterOk}{Environment.NewLine}" +
                        $" - Missing: {parameterMissing}{Environment.NewLine}{Environment.NewLine}" +
                        $"Family naming checks: {familyRows.Count}{Environment.NewLine}" +
                        $" - PASS: {familyPass}{Environment.NewLine}" +
                        $" - FAIL: {familyFail}{Environment.NewLine}" +
                        $" - SKIPPED: {familySkipped}{Environment.NewLine}{Environment.NewLine}" +
                        $"Subproject naming checks: {subprojectRows.Count}{Environment.NewLine}" +
                        $" - PASS: {subprojectPass}{Environment.NewLine}" +
                        $" - WARN: {subprojectWarn}{Environment.NewLine}" +
                        $" - FAIL: {subprojectFail}{Environment.NewLine}" +
                        $" - SKIPPED: {subprojectSkipped}{Environment.NewLine}{Environment.NewLine}" +
                        $"Chequeos XML de checkset: {modelCheckerSummary.Count}{Environment.NewLine}" +
                        $" - PASS: {modelCheckerPass}{Environment.NewLine}" +
                        $" - FAIL: {modelCheckerFail}{Environment.NewLine}" +
                        $" - COUNT (report-only): {modelCheckerCount}{Environment.NewLine}" +
                        $" - UNSUPPORTED: {modelCheckerUnsupported}{Environment.NewLine}" +
                        $" - SKIPPED (unchecked): {modelCheckerSkipped}{Environment.NewLine}{Environment.NewLine}" +
                        $"Files:{Environment.NewLine}" +
                        $" - {Path.GetFileName(excelReportPath)}{Environment.NewLine}" +
                        $" - {Path.GetFileName(parameterReportPath)}{Environment.NewLine}" +
                        $" - {Path.GetFileName(familyReportPath)}{Environment.NewLine}" +
                        $" - {Path.GetFileName(subprojectReportPath)}{Environment.NewLine}" +
                        $" - {Path.GetFileName(modelCheckerSummaryPath)}{Environment.NewLine}" +
                        $" - {Path.GetFileName(modelCheckerFailuresPath)}{Environment.NewLine}{Environment.NewLine}" +
                        "Open output folder now?",
                    CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No,
                    DefaultButton = TaskDialogResult.Yes
                };

                TaskDialogResult result = dialog.Show();
                if (result == TaskDialogResult.Yes)
                {
                    Process.Start(new ProcessStartInfo(outputDir) { UseShellExecute = true });
                }

                return Result.Succeeded;
            }
            catch (OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                TaskDialog.Show("JeiAudit", $"Audit failed.{Environment.NewLine}{ex.Message}");
                return Result.Failed;
            }
        }

        private static string PrepareOutputFolder()
        {
            string initialDirectory = ResolveInitialOutputDirectory();
            using (var dialog = new WinForms.FolderBrowserDialog())
            {
                dialog.Description = "Selecciona donde guardar los reportes de JeiAudit.";
                dialog.ShowNewFolderButton = true;
                if (!string.IsNullOrWhiteSpace(initialDirectory) && Directory.Exists(initialDirectory))
                {
                    dialog.SelectedPath = initialDirectory;
                }

                WinForms.DialogResult result = dialog.ShowDialog();
                if (result != WinForms.DialogResult.OK || string.IsNullOrWhiteSpace(dialog.SelectedPath))
                {
                    throw new OperationCanceledException("Audit cancelled by user.");
                }

                string rootFolder = dialog.SelectedPath.Trim();
                string folder = Path.Combine(rootFolder, $"JeiAudit_{DateTime.Now:yyyyMMdd_HHmmss}");
                Directory.CreateDirectory(folder);
                return folder;
            }
        }

        private static string ResolveInitialOutputDirectory()
        {
            string lastOutputFolder = PluginState.LoadLastOutputFolder();
            if (!string.IsNullOrWhiteSpace(lastOutputFolder) && Directory.Exists(lastOutputFolder))
            {
                try
                {
                    var info = new DirectoryInfo(lastOutputFolder);
                    if (info.Name.StartsWith("JeiAudit_", StringComparison.OrdinalIgnoreCase) && info.Parent != null)
                    {
                        return info.Parent.FullName;
                    }

                    return info.FullName;
                }
                catch
                {
                    // Ignore and fallback below.
                }
            }

            string desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            if (!string.IsNullOrWhiteSpace(desktop) && Directory.Exists(desktop))
            {
                return desktop;
            }

            string documents = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            return string.IsNullOrWhiteSpace(documents) ? string.Empty : documents;
        }

        private static string ResolveModelCheckerXmlPathForRun(Document doc)
        {
            string lastPath = PluginState.LoadLastXmlPath();
            string documentDir = string.IsNullOrWhiteSpace(doc.PathName)
                ? string.Empty
                : (Path.GetDirectoryName(doc.PathName) ?? string.Empty);
            string defaultPath = PluginState.ResolveDefaultXmlPath(lastPath, documentDir);

            if (!string.IsNullOrWhiteSpace(defaultPath) && File.Exists(defaultPath))
            {
                return defaultPath;
            }

            return PromptForModelCheckerXmlPath(defaultPath);
        }

        private static string PromptForModelCheckerXmlPath(string defaultPath)
        {
            using (var dialog = new WinForms.OpenFileDialog())
            {
                dialog.Title = "Seleccionar XML de checkset (MCSettings*.xml)";
                dialog.Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*";
                dialog.CheckFileExists = true;
                dialog.CheckPathExists = true;
                dialog.Multiselect = false;

                if (!string.IsNullOrWhiteSpace(defaultPath) && File.Exists(defaultPath))
                {
                    dialog.InitialDirectory = Path.GetDirectoryName(defaultPath);
                    dialog.FileName = Path.GetFileName(defaultPath);
                }
                else
                {
                    dialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                }

                WinForms.DialogResult result = dialog.ShowDialog();
                if (result == WinForms.DialogResult.OK && File.Exists(dialog.FileName))
                {
                    return dialog.FileName;
                }
            }

            TaskDialogResult proceed = TaskDialog.Show(
                "JeiAudit",
                "No se encontro un XML de checkset configurado. JeiAudit ejecutara solo auditorias internas.",
                TaskDialogCommonButtons.Ok | TaskDialogCommonButtons.Cancel);

            if (proceed == TaskDialogResult.Ok)
            {
                return string.Empty;
            }

            throw new OperationCanceledException("Audit cancelled by user.");
        }

        private static ChecksetMetadata LoadChecksetMetadata(string xmlPath)
        {
            var metadata = new ChecksetMetadata();
            XDocument xml = XDocument.Load(xmlPath);
            XElement? root = xml.Root;
            if (root == null)
            {
                return metadata;
            }

            metadata.Title = GetAttribute(root, "Name");
            metadata.Date = GetAttribute(root, "Date");
            metadata.Author = GetAttribute(root, "Author");
            metadata.Description = GetAttribute(root, "Description");

            if (string.IsNullOrWhiteSpace(metadata.Title))
            {
                metadata.Title = Path.GetFileNameWithoutExtension(xmlPath);
            }

            if (string.IsNullOrWhiteSpace(metadata.Date))
            {
                metadata.Date = DateTime.Now.ToString("dddd, dd 'de' MMMM 'de' yyyy", CultureInfo.GetCultureInfo("es-ES"));
            }

            if (string.IsNullOrWhiteSpace(metadata.Author))
            {
                metadata.Author = "-";
            }

            if (string.IsNullOrWhiteSpace(metadata.Description))
            {
                metadata.Description = "-";
            }

            return metadata;
        }

        private static List<RunAuditModelItem> BuildRunModelItems(Document hostDoc)
        {
            var items = new List<RunAuditModelItem>
            {
                new RunAuditModelItem
                {
                    Document = hostDoc,
                    DisplayName = hostDoc.Title,
                    Path = hostDoc.PathName,
                    IsLink = false,
                    IsSelected = true
                }
            };

            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                hostDoc.PathName ?? string.Empty
            };

            IEnumerable<RevitLinkInstance> links = new FilteredElementCollector(hostDoc)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>();

            foreach (RevitLinkInstance link in links)
            {
                Document? linkDoc = link.GetLinkDocument();
                if (linkDoc == null)
                {
                    continue;
                }

                string path = linkDoc.PathName ?? string.Empty;
                string dedupeKey = string.IsNullOrWhiteSpace(path) ? linkDoc.Title : path;
                if (seen.Contains(dedupeKey))
                {
                    continue;
                }

                seen.Add(dedupeKey);
                items.Add(new RunAuditModelItem
                {
                    Document = linkDoc,
                    DisplayName = linkDoc.Title,
                    Path = path,
                    IsLink = true,
                    IsSelected = false
                });
            }

            return items;
        }

        private static Dictionary<string, List<BindingInfo>> CollectParameterBindings(Document doc)
        {
            var map = new Dictionary<string, List<BindingInfo>>(StringComparer.Ordinal);
            BindingMap bindingMap = doc.ParameterBindings;
            DefinitionBindingMapIterator iterator = bindingMap.ForwardIterator();
            iterator.Reset();

            while (iterator.MoveNext())
            {
                Definition definition = iterator.Key;
                Binding? binding = iterator.Current as Binding;

                if (definition == null || binding == null)
                {
                    continue;
                }

                string parameterName = definition.Name ?? string.Empty;
                if (parameterName.Length == 0)
                {
                    continue;
                }

                var categoryIds = new HashSet<long>();
                if (binding is ElementBinding elementBinding && elementBinding.Categories != null)
                {
                    foreach (Category category in elementBinding.Categories)
                    {
                        if (category != null)
                        {
                            categoryIds.Add(category.Id.Value);
                        }
                    }
                }

                string bindingType = binding is InstanceBinding
                    ? "Instance"
                    : binding is TypeBinding
                        ? "Type"
                        : binding.GetType().Name;

                string guid = string.Empty;
                if (definition is ExternalDefinition externalDefinition)
                {
                    guid = externalDefinition.GUID.ToString();
                }

                if (!map.TryGetValue(parameterName, out List<BindingInfo> records))
                {
                    records = new List<BindingInfo>();
                    map[parameterName] = records;
                }

                records.Add(new BindingInfo(bindingType, categoryIds, guid));
            }

            return map;
        }

        private static List<ParameterAuditRow> RunParameterExistenceAudit(
            Document doc,
            Dictionary<string, List<BindingInfo>> bindings,
            List<string> auditedCategoryNames)
        {
            var rows = new List<ParameterAuditRow>();
            var allParameterNames = bindings.Keys.ToList();
            var categoryLookup = BuildLocalizedCategoryLookup(doc);

            var existingAuditedCategories = new List<Category>();
            foreach (string categoryName in auditedCategoryNames)
            {
                if (!TryResolveCategory(doc, categoryLookup, categoryName, out Category category, out _))
                {
                    continue;
                }

                if (!CategoryHasModelElements(doc, category.Id))
                {
                    continue;
                }

                if (existingAuditedCategories.All(c => c.Id.Value != category.Id.Value))
                {
                    existingAuditedCategories.Add(category);
                }
            }

            foreach (string requiredParameter in RequiredParameters)
            {
                bool existsByExactName = bindings.TryGetValue(requiredParameter, out List<BindingInfo> records);
                records ??= new List<BindingInfo>();

                string similarNames = string.Join(
                    "; ",
                    allParameterNames
                        .Where(name => IsNormalizedMatch(name, requiredParameter))
                        .OrderBy(name => name, StringComparer.Ordinal));

                foreach (Category category in existingAuditedCategories)
                {
                    bool boundToCategory = existsByExactName && records.Any(r => r.CategoryIds.Contains(category.Id.Value));

                    rows.Add(new ParameterAuditRow
                    {
                        ParameterName = requiredParameter,
                        CategoryName = category.Name ?? "-",
                        CategoryAvailable = "YES",
                        ExistsByExactName = existsByExactName ? "YES" : "NO",
                        BoundToCategory = boundToCategory ? "YES" : "NO",
                        BindingTypes = existsByExactName
                            ? string.Join("/", records.Select(r => r.BindingType).Distinct().OrderBy(v => v, StringComparer.Ordinal))
                            : "-",
                        ParameterGuids = existsByExactName
                            ? string.Join(";", records.Select(r => r.Guid).Where(v => !string.IsNullOrWhiteSpace(v)).Distinct().OrderBy(v => v, StringComparer.Ordinal))
                            : "-",
                        SimilarNames = string.IsNullOrWhiteSpace(similarNames) ? "-" : similarNames,
                        Status = existsByExactName && boundToCategory ? "OK" : "MISSING"
                    });
                }
            }

            return rows;
        }

        private static List<string> CollectAuditedCategoryNamesForParameterAudit(List<ModelCheckerCheckDef> checks)
        {
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (ModelCheckerCheckDef check in checks)
            {
                bool isGeneralSection = string.Equals(check.Section, "00_Generales_Modelos", StringComparison.OrdinalIgnoreCase);
                bool isMetradosSection = string.Equals(check.Section, "03_Metrados_Modelos", StringComparison.OrdinalIgnoreCase);
                if (!check.IsChecked || (!isGeneralSection && !isMetradosSection))
                {
                    continue;
                }

                foreach (ModelCheckerFilterDef filter in check.Filters)
                {
                    if (!string.Equals(filter.FilterCategory, "Category", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    if (!IsIncludeCategoryFilter(filter))
                    {
                        continue;
                    }

                    if (string.IsNullOrWhiteSpace(filter.Property))
                    {
                        continue;
                    }

                    names.Add(filter.Property.Trim());
                }
            }

            return names.OrderBy(v => v, StringComparer.OrdinalIgnoreCase).ToList();
        }

        private static bool CategoryHasModelElements(Document doc, ElementId categoryId)
        {
            ElementId firstElement = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .WherePasses(new ElementCategoryFilter(categoryId))
                .FirstElementId();
            return firstElement != ElementId.InvalidElementId;
        }
        private static string WriteParameterAuditCsv(string outputDir, List<ParameterAuditRow> rows)
        {
            string reportPath = Path.Combine(outputDir, "01_parameter_existence.csv");

            using (var writer = new StreamWriter(reportPath, false, new UTF8Encoding(true)))
            {
                writer.WriteLine("Parameter,Category,CategoryAvailable,ExistsByExactName,BoundToCategory,BindingTypes,ParameterGuids,SimilarNames,Status");

                foreach (ParameterAuditRow row in rows)
                {
                    writer.WriteLine(string.Join(
                        ",",
                        EscapeCsv(row.ParameterName),
                        EscapeCsv(row.CategoryName),
                        EscapeCsv(row.CategoryAvailable),
                        EscapeCsv(row.ExistsByExactName),
                        EscapeCsv(row.BoundToCategory),
                        EscapeCsv(row.BindingTypes),
                        EscapeCsv(row.ParameterGuids),
                        EscapeCsv(row.SimilarNames),
                        EscapeCsv(row.Status)));
                }
            }

            return reportPath;
        }

        private static List<FamilyNamingRow> RunFamilyNamingAudit(Document doc)
        {
            var rows = new List<FamilyNamingRow>();
            List<Family> families = new FilteredElementCollector(doc)
                .OfClass(typeof(Family))
                .Cast<Family>()
                .OrderBy(f => f.Name, StringComparer.Ordinal)
                .ToList();

            if (families.Count == 0)
            {
                rows.Add(new FamilyNamingRow
                {
                    FamilyName = "-",
                    HasArgPrefix = "-",
                    DisciplineCode = "-",
                    DisciplineCodeValid = "-",
                    HasNamePart = "-",
                    Status = "SKIPPED",
                    Reason = "No Family elements found in model."
                });
                return rows;
            }

            foreach (Family family in families)
            {
                string name = (family.Name ?? string.Empty).Trim();
                string[] parts = name.Split(new[] { '_' }, StringSplitOptions.None);
                bool hasArgPrefix = name.StartsWith("ARG_", StringComparison.OrdinalIgnoreCase);
                string disciplineCode = parts.Length >= 2 ? parts[1].Trim() : string.Empty;
                bool disciplineCodeValid = hasArgPrefix && FamilyDisciplineCodes.Contains(disciplineCode);
                bool hasNamePart = parts.Length >= 3 && !string.IsNullOrWhiteSpace(string.Join("_", parts.Skip(2)).Trim());

                string status;
                string reason;

                if (string.IsNullOrWhiteSpace(name))
                {
                    status = "FAIL";
                    reason = "Family name is empty.";
                }
                else if (!hasArgPrefix)
                {
                    status = "FAIL";
                    reason = "Missing required ARG_ prefix.";
                }
                else if (!disciplineCodeValid)
                {
                    status = "FAIL";
                    reason = $"Discipline code '{disciplineCode}' is not in allowed list ({string.Join("/", FamilyDisciplineCodes.OrderBy(v => v, StringComparer.Ordinal))}).";
                }
                else if (!hasNamePart)
                {
                    status = "FAIL";
                    reason = "Missing family name segment. Expected format: ARG_XXX_NOMBRE_FAMILIA.";
                }
                else
                {
                    status = "PASS";
                    reason = "-";
                }

                rows.Add(new FamilyNamingRow
                {
                    FamilyName = name,
                    HasArgPrefix = hasArgPrefix ? "YES" : "NO",
                    DisciplineCode = string.IsNullOrWhiteSpace(disciplineCode) ? "-" : disciplineCode,
                    DisciplineCodeValid = disciplineCodeValid ? "YES" : "NO",
                    HasNamePart = hasNamePart ? "YES" : "NO",
                    Status = status,
                    Reason = reason
                });
            }

            return rows;
        }

        private static string WriteFamilyNamingCsv(string outputDir, List<FamilyNamingRow> rows)
        {
            string reportPath = Path.Combine(outputDir, "02_family_naming.csv");
            using (var writer = new StreamWriter(reportPath, false, new UTF8Encoding(true)))
            {
                writer.WriteLine("FamilyName,HasARGPrefix,DisciplineCode,DisciplineCodeValid,HasNamePart,Status,Reason");
                foreach (FamilyNamingRow row in rows)
                {
                    writer.WriteLine(string.Join(
                        ",",
                        EscapeCsv(row.FamilyName),
                        EscapeCsv(row.HasArgPrefix),
                        EscapeCsv(row.DisciplineCode),
                        EscapeCsv(row.DisciplineCodeValid),
                        EscapeCsv(row.HasNamePart),
                        EscapeCsv(row.Status),
                        EscapeCsv(row.Reason)));
                }
            }

            return reportPath;
        }

        private static List<SubprojectNamingRow> RunSubprojectNamingAudit(Document doc)
        {
            var rows = new List<SubprojectNamingRow>();

            if (!doc.IsWorkshared)
            {
                rows.Add(new SubprojectNamingRow
                {
                    WorksetName = "-",
                    HasArgPrefix = "-",
                    HasNameAfterPrefix = "-",
                    HasSecondSeparator = "-",
                    Status = "SKIPPED",
                    Reason = "Model is not workshared; no user worksets/subprojects to audit."
                });
                return rows;
            }

            List<Workset> userWorksets = new FilteredWorksetCollector(doc)
                .OfKind(WorksetKind.UserWorkset)
                .ToWorksets()
                .OrderBy(w => w.Name, StringComparer.Ordinal)
                .ToList();

            if (userWorksets.Count == 0)
            {
                rows.Add(new SubprojectNamingRow
                {
                    WorksetName = "-",
                    HasArgPrefix = "-",
                    HasNameAfterPrefix = "-",
                    HasSecondSeparator = "-",
                    Status = "SKIPPED",
                    Reason = "No user worksets found."
                });
                return rows;
            }

            foreach (Workset workset in userWorksets)
            {
                string name = (workset.Name ?? string.Empty).Trim();
                bool hasArgPrefix = name.StartsWith("ARG_", StringComparison.OrdinalIgnoreCase);
                bool hasNameAfterPrefix = hasArgPrefix && name.Length > 4 && !string.IsNullOrWhiteSpace(name.Substring(4));
                bool hasSecondSeparator = hasNameAfterPrefix && name.IndexOf('_', 4) >= 0;

                string status;
                string reason;

                if (string.IsNullOrWhiteSpace(name))
                {
                    status = "FAIL";
                    reason = "Workset name is empty.";
                }
                else if (!hasArgPrefix || !hasNameAfterPrefix)
                {
                    status = "FAIL";
                    reason = "Subproject naming must start with ARG_ and include text after prefix.";
                }
                else if (!hasSecondSeparator)
                {
                    status = "WARN";
                    reason = "Recommended pattern: ARG_SISTEMA_ESP (or equivalent structured separator).";
                }
                else
                {
                    status = "PASS";
                    reason = "-";
                }

                rows.Add(new SubprojectNamingRow
                {
                    WorksetName = name,
                    HasArgPrefix = hasArgPrefix ? "YES" : "NO",
                    HasNameAfterPrefix = hasNameAfterPrefix ? "YES" : "NO",
                    HasSecondSeparator = hasSecondSeparator ? "YES" : "NO",
                    Status = status,
                    Reason = reason
                });
            }

            return rows;
        }

        private static string WriteSubprojectNamingCsv(string outputDir, List<SubprojectNamingRow> rows)
        {
            string reportPath = Path.Combine(outputDir, "03_subproject_naming.csv");
            using (var writer = new StreamWriter(reportPath, false, new UTF8Encoding(true)))
            {
                writer.WriteLine("WorksetName,HasARGPrefix,HasNameAfterPrefix,HasSecondSeparator,Status,Reason");
                foreach (SubprojectNamingRow row in rows)
                {
                    writer.WriteLine(string.Join(
                        ",",
                        EscapeCsv(row.WorksetName),
                        EscapeCsv(row.HasArgPrefix),
                        EscapeCsv(row.HasNameAfterPrefix),
                        EscapeCsv(row.HasSecondSeparator),
                        EscapeCsv(row.Status),
                        EscapeCsv(row.Reason)));
                }
            }

            return reportPath;
        }

        private static List<ModelCheckerCheckDef> LoadModelCheckerChecks(string xmlPath)
        {
            XDocument xml = XDocument.Load(xmlPath);
            var checks = new List<ModelCheckerCheckDef>();

            foreach (XElement headingElement in xml.Descendants().Where(e => e.Name.LocalName == "Heading"))
            {
                string headingName = GetAttribute(headingElement, "HeadingText");

                foreach (XElement sectionElement in headingElement.Elements().Where(e => e.Name.LocalName == "Section"))
                {
                    string sectionName = GetAttribute(sectionElement, "SectionName");

                    foreach (XElement checkElement in sectionElement.Elements().Where(e => e.Name.LocalName == "Check"))
                    {
                        var check = new ModelCheckerCheckDef
                        {
                            Heading = headingName,
                            Section = sectionName,
                            Name = GetAttribute(checkElement, "CheckName"),
                            Description = GetAttribute(checkElement, "Description"),
                            FailureMessage = GetAttribute(checkElement, "FailureMessage"),
                            ResultCondition = GetAttribute(checkElement, "ResultCondition"),
                            CheckType = GetAttribute(checkElement, "CheckType"),
                            IsChecked = ParseTrueFalse(GetAttribute(checkElement, "IsChecked"), true)
                        };

                        foreach (XElement filterElement in checkElement.Elements().Where(e => e.Name.LocalName == "Filter"))
                        {
                            check.Filters.Add(new ModelCheckerFilterDef
                            {
                                Operator = GetAttribute(filterElement, "Operator"),
                                FilterCategory = GetAttribute(filterElement, "Category"),
                                Property = GetAttribute(filterElement, "Property"),
                                Condition = GetAttribute(filterElement, "Condition"),
                                Value = GetAttribute(filterElement, "Value"),
                                CaseInsensitive = ParseTrueFalse(GetAttribute(filterElement, "CaseInsensitive"), false),
                                Unit = GetAttribute(filterElement, "Unit"),
                                UnitClass = GetAttribute(filterElement, "UnitClass"),
                                Validation = GetAttribute(filterElement, "Validation")
                            });
                        }

                        checks.Add(check);
                    }
                }
            }

            return checks;
        }

        private static void EvaluateModelCheckerChecks(
            Document doc,
            List<ModelCheckerCheckDef> checks,
            List<FamilyNamingRow> familyRows,
            List<SubprojectNamingRow> subprojectRows,
            Dictionary<string, List<BindingInfo>> bindings,
            List<ModelCheckerSummaryRow> summaryRows,
            List<ModelCheckerFailureRow> failureRows)
        {
            var categoryLookup = BuildLocalizedCategoryLookup(doc);
            var elementCache = new Dictionary<string, List<Element>>();

            foreach (ModelCheckerCheckDef check in checks)
            {
                var summary = new ModelCheckerSummaryRow
                {
                    Heading = check.Heading,
                    Section = check.Section,
                    CheckName = check.Name
                };

                if (!check.IsChecked)
                {
                    summary.Status = "SKIPPED";
                    summary.Reason = "Check IsChecked=False in XML.";
                    summaryRows.Add(summary);
                    continue;
                }

                if (TryEvaluateJeiAuditCustomCheck(doc, check, familyRows, subprojectRows, bindings, categoryLookup, summary, failureRows))
                {
                    summaryRows.Add(summary);
                    continue;
                }

                if (!TryGetCandidateElementsForCheck(doc, check, categoryLookup, elementCache, out List<Element> candidates, out string unsupportedReason))
                {
                    summary.Status = "UNSUPPORTED";
                    summary.Reason = unsupportedReason;
                    summaryRows.Add(summary);
                    continue;
                }

                summary.CandidateElements = candidates.Count;
                string mode = string.IsNullOrWhiteSpace(check.ResultCondition)
                    ? "FailMatchingElements"
                    : check.ResultCondition;
                string normalizedMode = NormalizeForCompare(mode);
                bool hasIncludeCategoryFilter = check.Filters
                    .Any(f =>
                        string.Equals(f.FilterCategory, "Category", StringComparison.OrdinalIgnoreCase) &&
                        IsIncludeCategoryFilter(f));

                if (summary.CandidateElements == 0 &&
                    hasIncludeCategoryFilter &&
                    normalizedMode == NormalizeForCompare("FailMatchingElements"))
                {
                    summary.Status = "SKIPPED";
                    summary.FailedElements = 0;
                    summary.Reason = "No hay elementos en la categoria evaluada; chequeo no ejecutado.";
                    summaryRows.Add(summary);
                    continue;
                }

                if (!TryBuildCheckEvaluationContext(doc, check, candidates, out CheckEvaluationContext context, out string contextReason))
                {
                    summary.Status = "UNSUPPORTED";
                    summary.Reason = contextReason;
                    summaryRows.Add(summary);
                    continue;
                }

                var matched = new List<Element>();
                bool unsupported = false;
                string unsupportedReasonInCheck = string.Empty;

                foreach (Element element in candidates)
                {
                    if (!TryEvaluateCheckOnElement(doc, element, check, context, categoryLookup, out bool isMatch, out string evalReason))
                    {
                        unsupported = true;
                        unsupportedReasonInCheck = evalReason;
                        break;
                    }

                    if (isMatch)
                    {
                        matched.Add(element);
                    }
                }

                if (unsupported)
                {
                    summary.Status = "UNSUPPORTED";
                    summary.Reason = unsupportedReasonInCheck;
                    summaryRows.Add(summary);
                    continue;
                }

                bool isFailure = false;
                bool exportMatchedList = false;

                if (normalizedMode == NormalizeForCompare("FailMatchingElements"))
                {
                    isFailure = matched.Count > 0;
                    summary.FailedElements = matched.Count;
                }
                else if (normalizedMode == NormalizeForCompare("FailNoMatchingElements") ||
                    normalizedMode == NormalizeForCompare("FailNoElements"))
                {
                    isFailure = matched.Count == 0;
                    summary.FailedElements = isFailure ? 1 : 0;
                }
                else if (normalizedMode == NormalizeForCompare("CountOnly"))
                {
                    summary.Status = "COUNT";
                    summary.FailedElements = matched.Count;
                    summary.Reason = $"CountOnly: {matched.Count} matching elements.";
                    summaryRows.Add(summary);
                    continue;
                }
                else if (normalizedMode == NormalizeForCompare("CountAndList"))
                {
                    summary.Status = "COUNT";
                    summary.FailedElements = matched.Count;
                    summary.Reason = $"CountAndList: {matched.Count} matching elements.";
                    summaryRows.Add(summary);
                    exportMatchedList = true;
                }
                else
                {
                    summary.Status = "UNSUPPORTED";
                    summary.Reason = $"Unsupported ResultCondition '{mode}'.";
                    summaryRows.Add(summary);
                    continue;
                }

                if (summary.Status != "COUNT")
                {
                    summary.Status = isFailure ? "FAIL" : "PASS";
                    if (isFailure)
                    {
                        if (TryBuildMissingParameterReasonForNoValueCheck(doc, check, candidates, out string missingParamReason))
                        {
                            summary.Reason = missingParamReason;
                        }
                        else
                        {
                            summary.Reason = string.IsNullOrWhiteSpace(check.FailureMessage)
                                ? $"Check failed ({mode})."
                                : check.FailureMessage;
                        }
                    }
                    else
                    {
                        summary.Reason = "-";
                    }
                    summaryRows.Add(summary);
                }

                if (isFailure && normalizedMode == NormalizeForCompare("FailMatchingElements"))
                {
                    exportMatchedList = true;
                }

                if (exportMatchedList)
                {
                    foreach (Element element in matched.Take(MaxFailedElementsPerCheckExport))
                    {
                        failureRows.Add(new ModelCheckerFailureRow
                        {
                            Heading = check.Heading,
                            Section = check.Section,
                            CheckName = check.Name,
                            ElementId = element.Id.Value,
                            Category = element.Category?.Name ?? "-",
                            FamilyOrType = ResolveTypeName(doc, element),
                            Name = element.Name ?? "-"
                        });
                    }
                }
            }
        }

        private static bool TryEvaluateJeiAuditCustomCheck(
            Document doc,
            ModelCheckerCheckDef check,
            List<FamilyNamingRow> familyRows,
            List<SubprojectNamingRow> subprojectRows,
            Dictionary<string, List<BindingInfo>> bindings,
            Dictionary<string, Category> categoryLookup,
            ModelCheckerSummaryRow summary,
            List<ModelCheckerFailureRow> failureRows)
        {
            bool isFamilyCheck = IsFamilyNamingCustomCheck(check);
            bool isSubprojectCheck = IsSubprojectNamingCustomCheck(check);
            bool isParameterExistenceCheck = IsParameterExistenceCustomCheck(check);
            bool isFileNamingCheck = IsFileNamingCustomCheck(check);
            bool isViewNamingCheck = IsViewNamingCustomCheck(check);
            bool isSheetNamingCheck = IsSheetNamingCustomCheck(check);

            if (!isFamilyCheck && !isSubprojectCheck && !isParameterExistenceCheck && !isFileNamingCheck && !isViewNamingCheck && !isSheetNamingCheck)
            {
                return false;
            }

            if (isFamilyCheck)
            {
                List<FamilyNamingRow> candidates = familyRows
                    .Where(r => !string.Equals(r.Status, "SKIPPED", StringComparison.OrdinalIgnoreCase))
                    .ToList();
                int failed = candidates.Count(r => string.Equals(r.Status, "FAIL", StringComparison.OrdinalIgnoreCase));
                summary.CandidateElements = candidates.Count;
                summary.FailedElements = failed;

                if (candidates.Count == 0)
                {
                    summary.Status = "SKIPPED";
                    summary.Reason = "Nomenclatura de familias: sin datos a evaluar.";
                    return true;
                }

                if (failed > 0)
                {
                    string sample = string.Join(
                        ", ",
                        candidates
                            .Where(r => string.Equals(r.Status, "FAIL", StringComparison.OrdinalIgnoreCase))
                            .Select(r => r.FamilyName)
                            .Where(v => !string.IsNullOrWhiteSpace(v))
                            .Take(3));
                    summary.Status = "FAIL";
                    summary.Reason = $"{failed} familias con nomenclatura invalida de {candidates.Count}."
                        + (string.IsNullOrWhiteSpace(sample) ? string.Empty : $" Ejemplos: {sample}.");
                }
                else
                {
                    summary.Status = "PASS";
                    summary.Reason = "-";
                }

                return true;
            }

            if (isFileNamingCheck)
            {
                return EvaluateFileNamingCustomCheck(doc, check, summary, failureRows);
            }

            if (isParameterExistenceCheck)
            {
                return EvaluateParameterExistenceCustomCheck(doc, check, bindings, categoryLookup, summary, failureRows);
            }

            if (isViewNamingCheck)
            {
                return EvaluateViewNamingCustomCheck(doc, check, summary, failureRows);
            }

            if (isSheetNamingCheck)
            {
                return EvaluateSheetNamingCustomCheck(doc, check, summary, failureRows);
            }

            List<SubprojectNamingRow> subprojectCandidates = subprojectRows
                .Where(r => !string.Equals(r.Status, "SKIPPED", StringComparison.OrdinalIgnoreCase))
                .ToList();
            int failedOrWarn = subprojectCandidates.Count(r =>
                string.Equals(r.Status, "FAIL", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(r.Status, "WARN", StringComparison.OrdinalIgnoreCase));
            summary.CandidateElements = subprojectCandidates.Count;
            summary.FailedElements = failedOrWarn;

            if (subprojectCandidates.Count == 0)
            {
                summary.Status = "SKIPPED";
                summary.Reason = "Nomenclatura de subproyectos: sin datos a evaluar.";
                return true;
            }

            if (failedOrWarn > 0)
            {
                int failCount = subprojectCandidates.Count(r => string.Equals(r.Status, "FAIL", StringComparison.OrdinalIgnoreCase));
                int warnCount = subprojectCandidates.Count(r => string.Equals(r.Status, "WARN", StringComparison.OrdinalIgnoreCase));
                string sample = string.Join(
                    ", ",
                    subprojectCandidates
                        .Where(r =>
                            string.Equals(r.Status, "FAIL", StringComparison.OrdinalIgnoreCase) ||
                            string.Equals(r.Status, "WARN", StringComparison.OrdinalIgnoreCase))
                        .Select(r => r.WorksetName)
                        .Where(v => !string.IsNullOrWhiteSpace(v))
                        .Take(3));
                summary.Status = "FAIL";
                summary.Reason = $"Subproyectos con alerta: FAIL={failCount}, WARN={warnCount}, total={subprojectCandidates.Count}."
                    + (string.IsNullOrWhiteSpace(sample) ? string.Empty : $" Ejemplos: {sample}.");
            }
            else
            {
                summary.Status = "PASS";
                summary.Reason = "-";
            }

            return true;
        }

        private static bool IsFamilyNamingCustomCheck(ModelCheckerCheckDef check)
        {
            string normalizedCheckName = NormalizeForCompare(check.Name);
            string normalizedSection = NormalizeForCompare(check.Section);
            string normalizedType = NormalizeForCompare(check.CheckType);

            if (normalizedType == NormalizeForCompare("JeiAuditFamilyNaming"))
            {
                return true;
            }

            if (normalizedCheckName.Contains(NormalizeForCompare("JeiAudit_Family_Naming")) ||
                normalizedCheckName.Contains(NormalizeForCompare("Family_Naming")) ||
                normalizedCheckName.Contains(NormalizeForCompare("Nomenclatura_Familias")))
            {
                return true;
            }

            return normalizedSection.Contains(NormalizeForCompare("Nomenclatura_Familias"));
        }

        private static bool IsSubprojectNamingCustomCheck(ModelCheckerCheckDef check)
        {
            string normalizedCheckName = NormalizeForCompare(check.Name);
            string normalizedSection = NormalizeForCompare(check.Section);
            string normalizedType = NormalizeForCompare(check.CheckType);

            if (normalizedType == NormalizeForCompare("JeiAuditSubprojectNaming"))
            {
                return true;
            }

            if (normalizedCheckName.Contains(NormalizeForCompare("JeiAudit_Subproject_Naming")) ||
                normalizedCheckName.Contains(NormalizeForCompare("Subproject_Naming")) ||
                normalizedCheckName.Contains(NormalizeForCompare("Nomenclatura_Subproyectos")))
            {
                return true;
            }

            return normalizedSection.Contains(NormalizeForCompare("Nomenclatura_Subproyectos"));
        }

        private static bool IsParameterExistenceCustomCheck(ModelCheckerCheckDef check)
        {
            string normalizedCheckName = NormalizeForCompare(check.Name);
            string normalizedSection = NormalizeForCompare(check.Section);
            string normalizedType = NormalizeForCompare(check.CheckType);

            if (normalizedType == NormalizeForCompare("JeiAuditParameterExistence"))
            {
                return true;
            }

            if (normalizedCheckName.Contains(NormalizeForCompare("JeiAudit_Parameter_Existence")) ||
                normalizedCheckName.Contains(NormalizeForCompare("Parameter_Existence")))
            {
                return true;
            }

            return normalizedSection.Contains(NormalizeForCompare("Nomenclatura_Parametros"));
        }

        private static bool IsFileNamingCustomCheck(ModelCheckerCheckDef check)
        {
            string normalizedCheckName = NormalizeForCompare(check.Name);
            string normalizedSection = NormalizeForCompare(check.Section);
            string normalizedType = NormalizeForCompare(check.CheckType);

            if (normalizedType == NormalizeForCompare("JeiAuditFileNaming"))
            {
                return true;
            }

            if (normalizedCheckName.Contains(NormalizeForCompare("JeiAudit_File_Naming")) ||
                normalizedCheckName.Contains(NormalizeForCompare("File_Naming")) ||
                normalizedCheckName.Contains(NormalizeForCompare("Nomenclatura_Archivo")))
            {
                return true;
            }

            return normalizedSection.Contains(NormalizeForCompare("Nomenclatura_Archivo"));
        }

        private static bool IsViewNamingCustomCheck(ModelCheckerCheckDef check)
        {
            string normalizedCheckName = NormalizeForCompare(check.Name);
            string normalizedSection = NormalizeForCompare(check.Section);
            string normalizedType = NormalizeForCompare(check.CheckType);

            if (normalizedType == NormalizeForCompare("JeiAuditViewNaming"))
            {
                return true;
            }

            if (normalizedCheckName.Contains(NormalizeForCompare("JeiAudit_View_Naming")) ||
                normalizedCheckName.Contains(NormalizeForCompare("View_Naming")) ||
                normalizedCheckName.Contains(NormalizeForCompare("Nomenclatura_Vistas")))
            {
                return true;
            }

            return normalizedSection.Contains(NormalizeForCompare("Nomenclatura_Vistas"));
        }

        private static bool IsSheetNamingCustomCheck(ModelCheckerCheckDef check)
        {
            string normalizedCheckName = NormalizeForCompare(check.Name);
            string normalizedSection = NormalizeForCompare(check.Section);
            string normalizedType = NormalizeForCompare(check.CheckType);

            if (normalizedType == NormalizeForCompare("JeiAuditSheetNaming"))
            {
                return true;
            }

            if (normalizedCheckName.Contains(NormalizeForCompare("JeiAudit_Sheet_Naming")) ||
                normalizedCheckName.Contains(NormalizeForCompare("Sheet_Naming")) ||
                normalizedCheckName.Contains(NormalizeForCompare("Nomenclatura_Planos")))
            {
                return true;
            }

            return normalizedSection.Contains(NormalizeForCompare("Nomenclatura_Planos"));
        }

        private static bool EvaluateParameterExistenceCustomCheck(
            Document doc,
            ModelCheckerCheckDef check,
            Dictionary<string, List<BindingInfo>> bindings,
            Dictionary<string, Category> categoryLookup,
            ModelCheckerSummaryRow summary,
            List<ModelCheckerFailureRow> failureRows)
        {
            string parameterName = CanonicalizeKnownParameterName(ResolveCustomValueFromCheck(check, "ParameterName", string.Empty));
            if (string.IsNullOrWhiteSpace(parameterName))
            {
                summary.CandidateElements = 1;
                summary.FailedElements = 1;
                summary.Status = "UNSUPPORTED";
                summary.Reason = "Check de existencia de parametro sin filtro Custom ParameterName.";
                return true;
            }

            string expectedBinding = ResolveCustomValueFromCheck(check, "ExpectedBinding", "Any");
            List<string> targetCategoryNames = check.Filters
                .Where(f => string.Equals(f.FilterCategory, "Category", StringComparison.OrdinalIgnoreCase))
                .Where(f => IsIncludeCategoryFilter(f))
                .Select(f => SanitizeLoadedText((f.Property ?? string.Empty).Trim()))
                .Where(v => !string.IsNullOrWhiteSpace(v))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            bool existsByExactName = bindings.TryGetValue(parameterName, out List<BindingInfo> records);
            records ??= new List<BindingInfo>();

            bool bindingTypeOk = true;
            if (existsByExactName &&
                !string.IsNullOrWhiteSpace(expectedBinding) &&
                !string.Equals(expectedBinding, "Any", StringComparison.OrdinalIgnoreCase))
            {
                bindingTypeOk = records.Any(r => string.Equals(r.BindingType, expectedBinding, StringComparison.OrdinalIgnoreCase));
            }

            var unresolvedCategories = new List<string>();
            var missingBoundCategories = new List<string>();

            if (existsByExactName && targetCategoryNames.Count > 0)
            {
                foreach (string categoryName in targetCategoryNames)
                {
                    if (!TryResolveCategory(doc, categoryLookup, categoryName, out Category category, out _))
                    {
                        unresolvedCategories.Add(categoryName);
                        continue;
                    }

                    bool isBound = records.Any(r => r.CategoryIds.Contains(category.Id.Value));
                    if (!isBound)
                    {
                        missingBoundCategories.Add(categoryName);
                    }
                }
            }

            int candidateCount = targetCategoryNames.Count > 0 ? targetCategoryNames.Count : 1;
            bool hasFailures = !existsByExactName || !bindingTypeOk || unresolvedCategories.Count > 0 || missingBoundCategories.Count > 0;

            summary.CandidateElements = candidateCount;
            summary.FailedElements = hasFailures
                ? (targetCategoryNames.Count > 0
                    ? Math.Max(1, unresolvedCategories.Count + missingBoundCategories.Count + (!bindingTypeOk ? 1 : 0) + (!existsByExactName ? targetCategoryNames.Count : 0))
                    : 1)
                : 0;

            if (!hasFailures)
            {
                summary.Status = "PASS";
                summary.Reason = "-";
                return true;
            }

            summary.Status = "FAIL";

            if (!existsByExactName)
            {
                string similar = string.Join(
                    ", ",
                    bindings.Keys
                        .Where(name => IsNormalizedMatch(name, parameterName))
                        .OrderBy(name => name, StringComparer.Ordinal)
                        .Take(5));

                string exactMissingMessage = $"No se encontro parametro exacto '{parameterName}' en el modelo.";
                if (string.IsNullOrWhiteSpace(check.FailureMessage))
                {
                    summary.Reason = exactMissingMessage;
                }
                else if (check.FailureMessage.IndexOf("No se encontro parametro", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    summary.Reason = check.FailureMessage;
                }
                else
                {
                    summary.Reason = $"{check.FailureMessage} {exactMissingMessage}";
                }

                if (!string.IsNullOrWhiteSpace(similar))
                {
                    summary.Reason += $" Similares: {similar}.";
                }

                failureRows.Add(new ModelCheckerFailureRow
                {
                    Heading = check.Heading,
                    Section = check.Section,
                    CheckName = check.Name,
                    ElementId = 0,
                    Category = "Parametros",
                    FamilyOrType = "Definicion de parametro",
                    Name = parameterName
                });

                return true;
            }

            var reasonParts = new List<string>();

            if (!bindingTypeOk)
            {
                reasonParts.Add($"No esta vinculado como '{expectedBinding}'.");
                failureRows.Add(new ModelCheckerFailureRow
                {
                    Heading = check.Heading,
                    Section = check.Section,
                    CheckName = check.Name,
                    ElementId = 0,
                    Category = "Parametros",
                    FamilyOrType = "Tipo de vinculacion",
                    Name = $"{parameterName} (esperado: {expectedBinding})"
                });
            }

            if (unresolvedCategories.Count > 0)
            {
                reasonParts.Add($"Categorias no resueltas: {string.Join(", ", unresolvedCategories)}.");
                foreach (string unresolved in unresolvedCategories.Take(MaxFailedElementsPerCheckExport))
                {
                    failureRows.Add(new ModelCheckerFailureRow
                    {
                        Heading = check.Heading,
                        Section = check.Section,
                        CheckName = check.Name,
                        ElementId = 0,
                        Category = unresolved,
                        FamilyOrType = "Categoria no disponible",
                        Name = parameterName
                    });
                }
            }

            if (missingBoundCategories.Count > 0)
            {
                reasonParts.Add($"No esta vinculado a categorias: {string.Join(", ", missingBoundCategories)}.");
                foreach (string missingCategory in missingBoundCategories.Take(MaxFailedElementsPerCheckExport))
                {
                    failureRows.Add(new ModelCheckerFailureRow
                    {
                        Heading = check.Heading,
                        Section = check.Section,
                        CheckName = check.Name,
                        ElementId = 0,
                        Category = missingCategory,
                        FamilyOrType = "Vinculacion de parametro",
                        Name = parameterName
                    });
                }
            }

            if (reasonParts.Count == 0)
            {
                reasonParts.Add($"No se pudo validar correctamente '{parameterName}'.");
            }

            summary.Reason = string.IsNullOrWhiteSpace(check.FailureMessage)
                ? $"{parameterName}: {string.Join(" ", reasonParts)}"
                : $"{check.FailureMessage} {string.Join(" ", reasonParts)}";

            return true;
        }

        private static bool IsIncludeCategoryFilter(ModelCheckerFilterDef filter)
        {
            string condition = NormalizeForCompare(filter.Condition);
            bool expected = ParseTrueFalse(filter.Value, true);

            if (condition == NormalizeForCompare("Included") || condition == NormalizeForCompare("Equal"))
            {
                return expected;
            }

            if (condition == NormalizeForCompare("Excluded") || condition == NormalizeForCompare("NotEqual"))
            {
                return !expected;
            }

            return false;
        }

        private static bool EvaluateFileNamingCustomCheck(
            Document doc,
            ModelCheckerCheckDef check,
            ModelCheckerSummaryRow summary,
            List<ModelCheckerFailureRow> failureRows)
        {
            string fileName = ResolveModelFileName(doc);
            if (string.IsNullOrWhiteSpace(fileName))
            {
                summary.CandidateElements = 1;
                summary.FailedElements = 1;
                summary.Status = "FAIL";
                summary.Reason = "No se pudo obtener el nombre del archivo del modelo.";
                return true;
            }

            string regexPattern = ResolveCustomRegexFromCheck(
                check,
                "FileNameRegex",
                @"^EIMI(?:-[A-Z0-9]+){8}(?:\.(RVT|rvt))?$");
            bool ignoreCase = ResolveCustomRegexIgnoreCase(check, "FileNameRegex");
            RegexOptions options = RegexOptions.CultureInvariant;
            if (ignoreCase)
            {
                options |= RegexOptions.IgnoreCase;
            }

            bool isMatch;
            try
            {
                isMatch = Regex.IsMatch(fileName, regexPattern, options);
            }
            catch (ArgumentException ex)
            {
                summary.CandidateElements = 1;
                summary.FailedElements = 1;
                summary.Status = "UNSUPPORTED";
                summary.Reason = $"Regex invalido en XML para nomenclatura de archivo: {ex.Message}";
                return true;
            }

            summary.CandidateElements = 1;
            summary.FailedElements = isMatch ? 0 : 1;

            if (isMatch)
            {
                summary.Status = "PASS";
                summary.Reason = "-";
                return true;
            }

            summary.Status = "FAIL";
            summary.Reason = string.IsNullOrWhiteSpace(check.FailureMessage)
                ? $"Nombre de archivo no cumple patron PEB. Archivo: '{fileName}'. Patron: '{regexPattern}'."
                : $"{check.FailureMessage} Archivo: '{fileName}'. Patron: '{regexPattern}'.";

            failureRows.Add(new ModelCheckerFailureRow
            {
                Heading = check.Heading,
                Section = check.Section,
                CheckName = check.Name,
                ElementId = 0,
                Category = "Archivo",
                FamilyOrType = "Nombre de archivo",
                Name = fileName
            });

            return true;
        }

        private static bool EvaluateViewNamingCustomCheck(
            Document doc,
            ModelCheckerCheckDef check,
            ModelCheckerSummaryRow summary,
            List<ModelCheckerFailureRow> failureRows)
        {
            string groupingParameterName = ResolveCustomValueFromCheck(check, "ViewGroupingParameter", string.Empty).Trim();
            string groupingAllowedValuesRaw = ResolveCustomValueFromCheck(check, "ViewGroupingAllowedValues", string.Empty).Trim();
            bool useGroupingValidation = !string.IsNullOrWhiteSpace(groupingParameterName) &&
                !string.IsNullOrWhiteSpace(groupingAllowedValuesRaw);

            string regexPattern = ResolveCustomRegexFromCheck(
                check,
                "ViewNameRegex",
                @"^[A-Z0-9_ \-\.]+$");
            bool ignoreCase = ResolveCustomRegexIgnoreCase(check, "ViewNameRegex");
            RegexOptions options = RegexOptions.CultureInvariant;
            if (ignoreCase)
            {
                options |= RegexOptions.IgnoreCase;
            }

            List<View> views = new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(v =>
                    !v.IsTemplate &&
                    v.ViewType != ViewType.DrawingSheet &&
                    v.ViewType != ViewType.ProjectBrowser &&
                    v.ViewType != ViewType.SystemBrowser &&
                    v.ViewType != ViewType.Internal)
                .ToList();

            summary.CandidateElements = views.Count;
            if (views.Count == 0)
            {
                summary.Status = "SKIPPED";
                summary.FailedElements = 0;
                summary.Reason = "No hay vistas/leyendas/tablas para auditar.";
                return true;
            }

            if (useGroupingValidation)
            {
                bool groupingIgnoreCase = ParseTrueFalse(
                    ResolveCustomValueFromCheck(check, "ViewGroupingIgnoreCase", "False"),
                    false);
                bool requireAllAllowedValues = ParseTrueFalse(
                    ResolveCustomValueFromCheck(check, "ViewGroupingRequireAllValues", "True"),
                    true);

                StringComparer comparer = groupingIgnoreCase ? StringComparer.OrdinalIgnoreCase : StringComparer.Ordinal;
                List<string> allowedValues = SplitList(groupingAllowedValuesRaw)
                    .Select(v => v?.Trim() ?? string.Empty)
                    .Where(v => !string.IsNullOrWhiteSpace(v))
                    .Distinct(comparer)
                    .ToList();

                if (allowedValues.Count == 0)
                {
                    summary.Status = "UNSUPPORTED";
                    summary.FailedElements = 1;
                    summary.Reason = $"Check de agrupacion de vistas sin valores permitidos para '{groupingParameterName}'.";
                    return true;
                }

                var allowedSet = new HashSet<string>(allowedValues, comparer);
                var presentValues = new HashSet<string>(comparer);
                var failedViews = new List<(View View, string CurrentValue)>();

                foreach (View view in views)
                {
                    FilterValue groupingValue = GetFilterValue(doc, view, groupingParameterName);
                    string currentValue = (groupingValue.Text ?? string.Empty).Trim();
                    bool isAllowed = groupingValue.Exists &&
                        !string.IsNullOrWhiteSpace(currentValue) &&
                        allowedSet.Contains(currentValue);

                    if (isAllowed)
                    {
                        presentValues.Add(currentValue);
                    }
                    else
                    {
                        failedViews.Add((view, currentValue));
                    }
                }

                List<string> missingAllowedValues = requireAllAllowedValues
                    ? allowedValues.Where(v => !presentValues.Contains(v)).ToList()
                    : new List<string>();

                summary.FailedElements = failedViews.Count + missingAllowedValues.Count;
                if (summary.FailedElements == 0)
                {
                    summary.Status = "PASS";
                    summary.Reason = "-";
                    return true;
                }

                summary.Status = "FAIL";
                string allowedText = string.Join(", ", allowedValues);
                var reasonParts = new List<string>();

                if (failedViews.Count > 0)
                {
                    string invalidExamples = string.Join(
                        ", ",
                        failedViews
                            .Select(v => string.IsNullOrWhiteSpace(v.CurrentValue)
                                ? $"{v.View.Name} [sin valor]"
                                : $"{v.View.Name} [{v.CurrentValue}]")
                            .Take(3));

                    reasonParts.Add(
                        $"{failedViews.Count} vistas/leyendas/tablas tienen '{groupingParameterName}' fuera del set permitido ({allowedText})." +
                        (string.IsNullOrWhiteSpace(invalidExamples) ? string.Empty : $" Ejemplos: {invalidExamples}."));
                }

                if (missingAllowedValues.Count > 0)
                {
                    reasonParts.Add($"Agrupaciones requeridas sin presencia en el modelo: {string.Join(", ", missingAllowedValues)}.");
                }

                string detailReason = string.Join(" ", reasonParts);
                summary.Reason = string.IsNullOrWhiteSpace(check.FailureMessage)
                    ? detailReason
                    : $"{check.FailureMessage} {detailReason}";

                foreach ((View view, string currentValue) in failedViews.Take(MaxFailedElementsPerCheckExport))
                {
                    string currentLabel = string.IsNullOrWhiteSpace(currentValue) ? "sin valor" : currentValue;
                    failureRows.Add(new ModelCheckerFailureRow
                    {
                        Heading = check.Heading,
                        Section = check.Section,
                        CheckName = check.Name,
                        ElementId = view.Id.Value,
                        Category = ResolveViewBucket(view),
                        FamilyOrType = view.ViewType.ToString(),
                        Name = $"{view.Name} [{groupingParameterName}: {currentLabel}]"
                    });
                }

                foreach (string missingValue in missingAllowedValues.Take(MaxFailedElementsPerCheckExport))
                {
                    failureRows.Add(new ModelCheckerFailureRow
                    {
                        Heading = check.Heading,
                        Section = check.Section,
                        CheckName = check.Name,
                        ElementId = 0,
                        Category = "Vistas",
                        FamilyOrType = "Agrupacion faltante",
                        Name = $"{groupingParameterName}: {missingValue}"
                    });
                }

                return true;
            }

            var failedViewsByName = new List<View>();
            foreach (View view in views)
            {
                string name = view.Name ?? string.Empty;
                bool isMatch;
                try
                {
                    isMatch = Regex.IsMatch(name, regexPattern, options);
                }
                catch (ArgumentException ex)
                {
                    summary.Status = "UNSUPPORTED";
                    summary.FailedElements = 1;
                    summary.Reason = $"Regex invalido en XML para nomenclatura de vistas: {ex.Message}";
                    return true;
                }

                if (!isMatch)
                {
                    failedViewsByName.Add(view);
                }
            }

            summary.FailedElements = failedViewsByName.Count;
            if (failedViewsByName.Count == 0)
            {
                summary.Status = "PASS";
                summary.Reason = "-";
                return true;
            }

            summary.Status = "FAIL";
            string examples = string.Join(", ", failedViewsByName.Select(v => v.Name).Where(v => !string.IsNullOrWhiteSpace(v)).Take(3));
            summary.Reason = string.IsNullOrWhiteSpace(check.FailureMessage)
                ? $"{failedViewsByName.Count} vistas/leyendas/tablas no cumplen nomenclatura de {views.Count}."
                : $"{check.FailureMessage} {failedViewsByName.Count} de {views.Count} no cumplen."
                    + (string.IsNullOrWhiteSpace(examples) ? string.Empty : $" Ejemplos: {examples}.")
                    + $" Patron: '{regexPattern}'.";

            foreach (View view in failedViewsByName.Take(MaxFailedElementsPerCheckExport))
            {
                failureRows.Add(new ModelCheckerFailureRow
                {
                    Heading = check.Heading,
                    Section = check.Section,
                    CheckName = check.Name,
                    ElementId = view.Id.Value,
                    Category = ResolveViewBucket(view),
                    FamilyOrType = view.ViewType.ToString(),
                    Name = view.Name ?? "-"
                });
            }

            return true;
        }

        private static bool EvaluateSheetNamingCustomCheck(
            Document doc,
            ModelCheckerCheckDef check,
            ModelCheckerSummaryRow summary,
            List<ModelCheckerFailureRow> failureRows)
        {
            string regexPattern = ResolveCustomRegexFromCheck(
                check,
                "SheetNameRegex",
                @"^[A-Z0-9_ \-\.]+$");
            bool ignoreCase = ResolveCustomRegexIgnoreCase(check, "SheetNameRegex");
            RegexOptions options = RegexOptions.CultureInvariant;
            if (ignoreCase)
            {
                options |= RegexOptions.IgnoreCase;
            }

            List<ViewSheet> sheets = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>()
                .Where(s => !s.IsPlaceholder)
                .ToList();

            summary.CandidateElements = sheets.Count;
            if (sheets.Count == 0)
            {
                summary.Status = "SKIPPED";
                summary.FailedElements = 0;
                summary.Reason = "No hay planos para auditar.";
                return true;
            }

            var failedSheets = new List<ViewSheet>();
            foreach (ViewSheet sheet in sheets)
            {
                string name = sheet.Name ?? string.Empty;
                bool isMatch;
                try
                {
                    isMatch = Regex.IsMatch(name, regexPattern, options);
                }
                catch (ArgumentException ex)
                {
                    summary.Status = "UNSUPPORTED";
                    summary.FailedElements = 1;
                    summary.Reason = $"Regex invalido en XML para nomenclatura de planos: {ex.Message}";
                    return true;
                }

                if (!isMatch)
                {
                    failedSheets.Add(sheet);
                }
            }

            summary.FailedElements = failedSheets.Count;
            if (failedSheets.Count == 0)
            {
                summary.Status = "PASS";
                summary.Reason = "-";
                return true;
            }

            summary.Status = "FAIL";
            string examples = string.Join(", ", failedSheets.Select(s => s.SheetNumber + " - " + s.Name).Take(3));
            summary.Reason = string.IsNullOrWhiteSpace(check.FailureMessage)
                ? $"{failedSheets.Count} planos no cumplen nomenclatura de {sheets.Count}."
                : $"{check.FailureMessage} {failedSheets.Count} de {sheets.Count} no cumplen."
                    + (string.IsNullOrWhiteSpace(examples) ? string.Empty : $" Ejemplos: {examples}.")
                    + $" Patron: '{regexPattern}'.";

            foreach (ViewSheet sheet in failedSheets.Take(MaxFailedElementsPerCheckExport))
            {
                failureRows.Add(new ModelCheckerFailureRow
                {
                    Heading = check.Heading,
                    Section = check.Section,
                    CheckName = check.Name,
                    ElementId = sheet.Id.Value,
                    Category = "Planos",
                    FamilyOrType = "ViewSheet",
                    Name = $"{sheet.SheetNumber} - {sheet.Name}"
                });
            }

            return true;
        }

        private static string ResolveViewBucket(View view)
        {
            if (view == null)
            {
                return "Vistas";
            }

            if (view.ViewType == ViewType.Legend)
            {
                return "Leyendas";
            }

            if (view is ViewSchedule)
            {
                return "Tablas de Planificacion";
            }

            return "Vistas";
        }

        private static string ResolveModelFileName(Document doc)
        {
            string path = doc.PathName ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(path))
            {
                return Path.GetFileName(path);
            }

            string title = doc.Title ?? string.Empty;
            return string.IsNullOrWhiteSpace(title) ? string.Empty : title + ".rvt";
        }

        private static string ResolveCustomValueFromCheck(ModelCheckerCheckDef check, string propertyName, string defaultValue)
        {
            ModelCheckerFilterDef filter = check.Filters.FirstOrDefault(f =>
                string.Equals(f.FilterCategory, "Custom", StringComparison.OrdinalIgnoreCase) &&
                string.Equals(f.Property, propertyName, StringComparison.OrdinalIgnoreCase));

            if (filter == null || string.IsNullOrWhiteSpace(filter.Value))
            {
                return defaultValue;
            }

            return SanitizeLoadedText(filter.Value.Trim());
        }

        private static string ResolveCustomRegexFromCheck(ModelCheckerCheckDef check, string propertyName, string defaultPattern)
        {
            return ResolveCustomValueFromCheck(check, propertyName, defaultPattern);
        }

        private static bool ResolveCustomRegexIgnoreCase(ModelCheckerCheckDef check, string propertyName)
        {
            ModelCheckerFilterDef filter = check.Filters.FirstOrDefault(f =>
                string.Equals(f.FilterCategory, "Custom", StringComparison.OrdinalIgnoreCase) &&
                string.Equals(f.Property, propertyName, StringComparison.OrdinalIgnoreCase));

            return filter != null && filter.CaseInsensitive;
        }

        private static bool TryGetCandidateElementsForCheck(
            Document doc,
            ModelCheckerCheckDef check,
            Dictionary<string, Category> categoryLookup,
            Dictionary<string, List<Element>> cache,
            out List<Element> candidates,
            out string reason)
        {
            candidates = new List<Element>();
            reason = string.Empty;

            if (check.Filters.Count == 0)
            {
                reason = "Check has no filters.";
                return false;
            }

            var includeCategoryIds = new HashSet<long>();
            var excludeCategoryIds = new HashSet<long>();

            List<ModelCheckerFilterDef> categoryFilters = check.Filters
                .Where(f => string.Equals(f.FilterCategory, "Category", StringComparison.OrdinalIgnoreCase))
                .ToList();

            foreach (ModelCheckerFilterDef filter in categoryFilters)
            {
                if (!TryResolveCategory(doc, categoryLookup, filter.Property, out Category category, out string categoryReason))
                {
                    reason = $"Category not resolved: '{filter.Property}'. {categoryReason}";
                    return false;
                }

                string condition = NormalizeForCompare(filter.Condition);
                bool expected = ParseTrueFalse(filter.Value, true);
                bool include;

                if (condition == NormalizeForCompare("Included") || condition == NormalizeForCompare("Equal"))
                {
                    include = expected;
                }
                else if (condition == NormalizeForCompare("Excluded") || condition == NormalizeForCompare("NotEqual"))
                {
                    include = !expected;
                }
                else
                {
                    reason = $"Unsupported Category condition '{filter.Condition}'.";
                    return false;
                }

                if (include)
                {
                    includeCategoryIds.Add(category.Id.Value);
                    excludeCategoryIds.Remove(category.Id.Value);
                }
                else
                {
                    excludeCategoryIds.Add(category.Id.Value);
                    includeCategoryIds.Remove(category.Id.Value);
                }
            }

            var ids = new HashSet<long>();

            if (includeCategoryIds.Count == 0)
            {
                foreach (Element element in GetAllElementsAnyType(doc, cache))
                {
                    if (element.Category == null)
                    {
                        continue;
                    }

                    if (excludeCategoryIds.Contains(element.Category.Id.Value))
                    {
                        continue;
                    }

                    if (ids.Add(element.Id.Value))
                    {
                        candidates.Add(element);
                    }
                }
            }
            else
            {
                foreach (long categoryIdValue in includeCategoryIds)
                {
                    var categoryId = new ElementId(categoryIdValue);
                    foreach (Element element in GetElementsForCategoryAnyType(doc, categoryId, cache))
                    {
                        if (element.Category == null)
                        {
                            continue;
                        }

                        if (excludeCategoryIds.Contains(element.Category.Id.Value))
                        {
                            continue;
                        }

                        if (ids.Add(element.Id.Value))
                        {
                            candidates.Add(element);
                        }
                    }
                }
            }

            if (!TryApplyTypeOrInstanceScope(check.Filters, candidates, out List<Element> scopedCandidates, out string typeScopeReason))
            {
                reason = typeScopeReason;
                return false;
            }

            candidates = scopedCandidates;

            return true;
        }

        private static bool TryApplyTypeOrInstanceScope(
            List<ModelCheckerFilterDef> filters,
            List<Element> source,
            out List<Element> scoped,
            out string reason)
        {
            scoped = source;
            reason = string.Empty;

            List<ModelCheckerFilterDef> typeFilters = filters
                .Where(f => string.Equals(f.FilterCategory, "TypeOrInstance", StringComparison.OrdinalIgnoreCase))
                .ToList();

            if (typeFilters.Count == 0)
            {
                return true;
            }

            var result = new List<Element>(source.Count);
            foreach (Element element in source)
            {
                if (!TryElementMatchesTypeOrInstanceFilters(element, typeFilters, out bool matches, out string elementReason))
                {
                    reason = elementReason;
                    return false;
                }

                if (matches)
                {
                    result.Add(element);
                }
            }

            scoped = result;
            return true;
        }

        private static bool TryElementMatchesTypeOrInstanceFilters(
            Element element,
            List<ModelCheckerFilterDef> typeFilters,
            out bool matches,
            out string reason)
        {
            matches = true;
            reason = string.Empty;
            bool isType = element is ElementType;

            foreach (ModelCheckerFilterDef filter in typeFilters)
            {
                string typeProperty = NormalizeForCompare(filter.Property);
                if (typeProperty != NormalizeForCompare("Is Element Type") &&
                    typeProperty != NormalizeForCompare("Is Type") &&
                    typeProperty != NormalizeForCompare("Es Tipo De Elemento"))
                {
                    reason = $"Unsupported TypeOrInstance property '{filter.Property}'.";
                    return false;
                }

                bool expected = ParseTrueFalse(filter.Value, false);
                string condition = NormalizeForCompare(filter.Condition);
                bool currentMatch;

                if (condition == NormalizeForCompare("Equal") || condition == NormalizeForCompare("Included"))
                {
                    currentMatch = isType == expected;
                }
                else if (condition == NormalizeForCompare("NotEqual") || condition == NormalizeForCompare("Excluded"))
                {
                    currentMatch = isType != expected;
                }
                else
                {
                    reason = $"Unsupported TypeOrInstance condition '{filter.Condition}'.";
                    return false;
                }

                if (!currentMatch)
                {
                    matches = false;
                    return true;
                }
            }

            return true;
        }

        private static bool TryResolveCategory(
            Document doc,
            Dictionary<string, Category> localizedCategories,
            string categoryName,
            out Category category,
            out string reason)
        {
            category = null!;
            reason = string.Empty;

            string raw = categoryName?.Trim() ?? string.Empty;
            string normalized = NormalizeForCompare(raw);

            if (localizedCategories.TryGetValue(normalized, out Category localizedCategory))
            {
                category = localizedCategory;
                return true;
            }

            var candidateNames = new List<string>();
            if (!string.IsNullOrWhiteSpace(raw))
            {
                candidateNames.Add(raw);
            }

            if (TryGetAliasBuiltInCategoryName(raw, out string aliasBuiltInCategory))
            {
                candidateNames.Add(aliasBuiltInCategory);
            }

            foreach (string candidate in ExpandCategoryResolutionCandidates(candidateNames))
            {
                if (TryResolveCategoryFromBuiltInName(doc, candidate, out Category resolvedFromBuiltIn))
                {
                    category = resolvedFromBuiltIn;
                    return true;
                }

                if (TryGetAliasBuiltInCategoryName(candidate, out string candidateAliasBuiltInCategory) &&
                    TryResolveCategoryFromBuiltInName(doc, candidateAliasBuiltInCategory, out resolvedFromBuiltIn))
                {
                    category = resolvedFromBuiltIn;
                    return true;
                }

                string normalizedCandidate = NormalizeForCompare(candidate);
                if (localizedCategories.TryGetValue(normalizedCandidate, out Category localizedFromCandidate))
                {
                    category = localizedFromCandidate;
                    return true;
                }
            }

            reason = "No matching Revit category found for current model language/context.";
            return false;
        }

        private static IEnumerable<string> ExpandCategoryResolutionCandidates(IEnumerable<string> seedCandidates)
        {
            var queue = new Queue<string>(
                seedCandidates
                    .Where(s => !string.IsNullOrWhiteSpace(s))
                    .Select(s => s.Trim()));
            var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            while (queue.Count > 0)
            {
                string current = queue.Dequeue();
                if (!visited.Add(current))
                {
                    continue;
                }

                yield return current;

                if (BuiltInCategoryFallbacks.TryGetValue(current, out string[] fallbacks))
                {
                    foreach (string fallback in fallbacks)
                    {
                        if (!string.IsNullOrWhiteSpace(fallback))
                        {
                            queue.Enqueue(fallback.Trim());
                        }
                    }
                }
            }
        }

        private static bool TryGetAliasBuiltInCategoryName(string categoryAlias, out string builtInCategoryName)
        {
            builtInCategoryName = string.Empty;
            if (string.IsNullOrWhiteSpace(categoryAlias))
            {
                return false;
            }

            string normalizedAlias = NormalizeForCompare(categoryAlias);
            KeyValuePair<string, string> aliasPair = CategoryAliases.FirstOrDefault(kvp =>
                string.Equals(NormalizeForCompare(kvp.Key), normalizedAlias, StringComparison.Ordinal));

            if (string.IsNullOrWhiteSpace(aliasPair.Value))
            {
                return false;
            }

            builtInCategoryName = aliasPair.Value.Trim();
            return true;
        }

        private static bool TryResolveCategoryFromBuiltInName(
            Document doc,
            string builtInCategoryName,
            out Category category)
        {
            category = null!;
            if (string.IsNullOrWhiteSpace(builtInCategoryName))
            {
                return false;
            }

            if (!Enum.TryParse(builtInCategoryName.Trim(), ignoreCase: true, out BuiltInCategory bic))
            {
                return false;
            }

            Category resolved = Category.GetCategory(doc, bic);
            if (resolved == null)
            {
                return false;
            }

            category = resolved;
            return true;
        }

        private static Dictionary<string, Category> BuildLocalizedCategoryLookup(Document doc)
        {
            var map = new Dictionary<string, Category>(StringComparer.Ordinal);
            Categories categories = doc.Settings.Categories;

            foreach (Category category in categories)
            {
                if (category == null || string.IsNullOrWhiteSpace(category.Name))
                {
                    continue;
                }

                string key = NormalizeForCompare(category.Name);
                if (!map.ContainsKey(key))
                {
                    map[key] = category;
                }
            }

            return map;
        }

        private static List<Element> GetElementsForCategoryAnyType(
            Document doc,
            ElementId categoryId,
            Dictionary<string, List<Element>> cache)
        {
            string cacheKey = $"{categoryId.Value}|ANY";
            if (cache.TryGetValue(cacheKey, out List<Element> cached))
            {
                return cached;
            }

            FilteredElementCollector collector = new FilteredElementCollector(doc)
                .WherePasses(new ElementCategoryFilter(categoryId));

            List<Element> elements = collector.ToElements().ToList();
            cache[cacheKey] = elements;
            return elements;
        }

        private static List<Element> GetAllElementsAnyType(
            Document doc,
            Dictionary<string, List<Element>> cache)
        {
            const string cacheKey = "ALL|ANY";
            if (cache.TryGetValue(cacheKey, out List<Element> cached))
            {
                return cached;
            }

            List<Element> instances = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .ToElements()
                .ToList();

            List<Element> types = new FilteredElementCollector(doc)
                .WhereElementIsElementType()
                .ToElements()
                .ToList();

            List<Element> all = instances
                .Concat(types)
                .ToList();

            cache[cacheKey] = all;
            return all;
        }

        private sealed class CheckEvaluationContext
        {
            public Dictionary<string, HashSet<string>> DuplicateTokensByFilterKey { get; } =
                new Dictionary<string, HashSet<string>>(StringComparer.Ordinal);
        }

        private static bool TryBuildCheckEvaluationContext(
            Document doc,
            ModelCheckerCheckDef check,
            List<Element> candidates,
            out CheckEvaluationContext context,
            out string reason)
        {
            context = new CheckEvaluationContext();
            reason = string.Empty;

            foreach (ModelCheckerFilterDef filter in check.Filters)
            {
                if (NormalizeForCompare(filter.Condition) != NormalizeForCompare("Duplicated"))
                {
                    continue;
                }

                var counts = new Dictionary<string, int>(StringComparer.Ordinal);
                foreach (Element candidate in candidates)
                {
                    if (!TryGetFilterValueForFilterCategory(doc, candidate, filter, out FilterValue value, out string valueReason))
                    {
                        reason = string.IsNullOrWhiteSpace(valueReason)
                            ? $"Unsupported filter category '{filter.FilterCategory}' for Duplicated condition."
                            : valueReason;
                        return false;
                    }

                    string token = GetDuplicateComparableToken(value, filter.CaseInsensitive);
                    if (string.IsNullOrWhiteSpace(token))
                    {
                        continue;
                    }

                    if (!counts.TryGetValue(token, out int count))
                    {
                        counts[token] = 1;
                    }
                    else
                    {
                        counts[token] = count + 1;
                    }
                }

                var duplicates = new HashSet<string>(
                    counts.Where(kvp => kvp.Value > 1).Select(kvp => kvp.Key),
                    StringComparer.Ordinal);

                context.DuplicateTokensByFilterKey[GetFilterDuplicateContextKey(filter)] = duplicates;
            }

            return true;
        }

        private static bool TryEvaluateCheckOnElement(
            Document doc,
            Element element,
            ModelCheckerCheckDef check,
            CheckEvaluationContext context,
            Dictionary<string, Category> categoryLookup,
            out bool isMatch,
            out string reason)
        {
            bool hasResult = false;
            bool result = false;
            reason = string.Empty;

            foreach (ModelCheckerFilterDef filter in check.Filters)
            {
                if (!TryEvaluateSingleFilter(doc, element, filter, context, categoryLookup, out bool filterResult, out string evalReason))
                {
                    isMatch = false;
                    reason = evalReason;
                    return false;
                }

                if (!hasResult)
                {
                    result = filterResult;
                    hasResult = true;
                    continue;
                }

                string op = string.IsNullOrWhiteSpace(filter.Operator) ? "And" : filter.Operator;
                string normalizedOp = NormalizeForCompare(op).Replace(" ", string.Empty);

                if (normalizedOp == "AND")
                {
                    result = result && filterResult;
                }
                else if (normalizedOp == "OR")
                {
                    result = result || filterResult;
                }
                else if (normalizedOp == "ANDNOT")
                {
                    result = result && !filterResult;
                }
                else if (normalizedOp == "ORNOT")
                {
                    result = result || !filterResult;
                }
                else if (normalizedOp == "EXCLUDE")
                {
                    result = result && !filterResult;
                }
                else
                {
                    isMatch = false;
                    reason = $"Unsupported operator '{op}'. Supported: And/Or/AndNot/OrNot/Exclude.";
                    return false;
                }
            }

            isMatch = hasResult && result;
            return true;
        }

        private static bool TryEvaluateSingleFilter(
            Document doc,
            Element element,
            ModelCheckerFilterDef filter,
            CheckEvaluationContext context,
            Dictionary<string, Category> categoryLookup,
            out bool result,
            out string reason)
        {
            result = false;
            reason = string.Empty;
            string filterCategoryNormalized = NormalizeForCompare(filter.FilterCategory);

            if (string.Equals(filter.FilterCategory, "Category", StringComparison.OrdinalIgnoreCase))
            {
                if (!TryResolveCategory(doc, categoryLookup, filter.Property, out Category category, out string categoryReason))
                {
                    reason = $"Category not resolved: '{filter.Property}'. {categoryReason}";
                    return false;
                }

                bool isInCategory = element.Category != null && element.Category.Id.Value == category.Id.Value;
                bool expected = ParseTrueFalse(filter.Value, true);
                string condition = NormalizeForCompare(filter.Condition);

                if (condition == NormalizeForCompare("Included"))
                {
                    result = expected ? isInCategory : !isInCategory;
                    return true;
                }

                if (condition == NormalizeForCompare("Excluded"))
                {
                    result = expected ? !isInCategory : isInCategory;
                    return true;
                }

                if (condition == NormalizeForCompare("Equal"))
                {
                    result = isInCategory == expected;
                    return true;
                }

                if (condition == NormalizeForCompare("NotEqual"))
                {
                    result = isInCategory != expected;
                    return true;
                }

                reason = $"Unsupported Category condition '{filter.Condition}'.";
                return false;
            }

            if (string.Equals(filter.FilterCategory, "TypeOrInstance", StringComparison.OrdinalIgnoreCase))
            {
                string typeProperty = NormalizeForCompare(filter.Property);
                if (typeProperty != NormalizeForCompare("Is Element Type") &&
                    typeProperty != NormalizeForCompare("Is Type") &&
                    typeProperty != NormalizeForCompare("Es Tipo De Elemento"))
                {
                    reason = $"Unsupported TypeOrInstance property '{filter.Property}'.";
                    return false;
                }

                bool isType = element is ElementType;
                bool expected = ParseTrueFalse(filter.Value, false);
                string condition = NormalizeForCompare(filter.Condition);

                if (condition == NormalizeForCompare("Equal") || condition == NormalizeForCompare("Included"))
                {
                    result = isType == expected;
                    return true;
                }

                if (condition == NormalizeForCompare("NotEqual") || condition == NormalizeForCompare("Excluded"))
                {
                    result = isType != expected;
                    return true;
                }

                reason = $"Unsupported TypeOrInstance condition '{filter.Condition}'.";
                return false;
            }

            if (TryGetFilterValueForFilterCategory(doc, element, filter, out FilterValue value, out string valueReason))
            {
                if (!TryEvaluateParameterCondition(doc, element, context, value, filter, out bool conditionResult, out string conditionReason))
                {
                    reason = conditionReason;
                    return false;
                }

                result = conditionResult;
                return true;
            }

            if (!string.IsNullOrWhiteSpace(valueReason))
            {
                reason = valueReason;
                return false;
            }

            reason = $"Unsupported filter category '{filter.FilterCategory}'.";
            return false;
        }

        private static bool TryGetFilterValueForFilterCategory(
            Document doc,
            Element element,
            ModelCheckerFilterDef filter,
            out FilterValue value,
            out string reason)
        {
            reason = string.Empty;
            string filterCategoryNormalized = NormalizeForCompare(filter.FilterCategory);

            if (filterCategoryNormalized == NormalizeForCompare("Parameter") ||
                filterCategoryNormalized == NormalizeForCompare("APIParameter"))
            {
                value = GetFilterValue(doc, element, filter.Property);
                return true;
            }

            if (filterCategoryNormalized == NormalizeForCompare("Family"))
            {
                value = GetFilterValue(doc, element, "Family Name");
                return true;
            }

            if (filterCategoryNormalized == NormalizeForCompare("Type"))
            {
                value = GetFilterValue(doc, element, "Type Name");
                return true;
            }

            if (filterCategoryNormalized == NormalizeForCompare("Workset"))
            {
                value = GetFilterValue(doc, element, "Workset Name");
                return true;
            }

            if (filterCategoryNormalized == NormalizeForCompare("APIType"))
            {
                string apiType = element.GetType().Name ?? string.Empty;
                value = CreateTextFilterValue(apiType);
                return true;
            }

            if (filterCategoryNormalized == NormalizeForCompare("Level"))
            {
                value = CreateTextFilterValue(ResolveLevelName(doc, element));
                return true;
            }

            if (filterCategoryNormalized == NormalizeForCompare("PhaseCreated"))
            {
                value = CreateTextFilterValue(ResolveElementNameByIdParameter(doc, element, "PHASE_CREATED"));
                return true;
            }

            if (filterCategoryNormalized == NormalizeForCompare("PhaseDemolished"))
            {
                value = CreateTextFilterValue(ResolveElementNameByIdParameter(doc, element, "PHASE_DEMOLISHED"));
                return true;
            }

            if (filterCategoryNormalized == NormalizeForCompare("PhaseStatus"))
            {
                value = CreateTextFilterValue(GetBuiltInParameterText(element, "PHASE_STATUS"));
                return true;
            }

            if (filterCategoryNormalized == NormalizeForCompare("DesignOption"))
            {
                value = CreateTextFilterValue(ResolveDesignOptionName(doc, element));
                return true;
            }

            if (filterCategoryNormalized == NormalizeForCompare("View"))
            {
                value = CreateTextFilterValue(ResolveOwnerViewName(doc, element));
                return true;
            }

            if (filterCategoryNormalized == NormalizeForCompare("StructuralType"))
            {
                value = CreateTextFilterValue(ResolveStructuralTypeName(element));
                return true;
            }

            if (filterCategoryNormalized == NormalizeForCompare("Host") ||
                filterCategoryNormalized == NormalizeForCompare("HostParameter"))
            {
                Element? host = GetHostElement(doc, element);
                if (host == null)
                {
                    value = new FilterValue
                    {
                        Exists = false,
                        HasValue = false,
                        Text = string.Empty
                    };
                    return true;
                }

                if (filterCategoryNormalized == NormalizeForCompare("HostParameter"))
                {
                    value = GetFilterValue(doc, host, filter.Property);
                    return true;
                }

                string propertyNormalized = NormalizeForCompare(filter.Property);
                if (string.IsNullOrWhiteSpace(filter.Property) ||
                    propertyNormalized == NormalizeForCompare("Host") ||
                    propertyNormalized == NormalizeForCompare("Name") ||
                    propertyNormalized == NormalizeForCompare("Element Name"))
                {
                    value = CreateTextFilterValue(host.Name ?? string.Empty);
                    return true;
                }

                if (propertyNormalized == NormalizeForCompare("Category") ||
                    propertyNormalized == NormalizeForCompare("Category Name"))
                {
                    value = CreateTextFilterValue(host.Category?.Name ?? string.Empty);
                    return true;
                }

                if (propertyNormalized == NormalizeForCompare("Element Id") ||
                    propertyNormalized == NormalizeForCompare("Id"))
                {
                    value = new FilterValue
                    {
                        Exists = true,
                        HasValue = true,
                        Text = host.Id.Value.ToString(CultureInfo.InvariantCulture),
                        Number = host.Id.Value
                    };
                    return true;
                }

                value = GetFilterValue(doc, host, filter.Property);
                return true;
            }

            if (filterCategoryNormalized == NormalizeForCompare("Room"))
            {
                string roomValue = GetBuiltInParameterText(element, "ROOM_NAME");
                if (string.IsNullOrWhiteSpace(roomValue))
                {
                    roomValue = GetBuiltInParameterText(element, "ROOM_NUMBER");
                }

                if (!string.IsNullOrWhiteSpace(roomValue))
                {
                    value = CreateTextFilterValue(roomValue);
                    return true;
                }

                value = GetFilterValue(doc, element, filter.Property);
                return true;
            }

            if (filterCategoryNormalized == NormalizeForCompare("Space"))
            {
                FilterValue direct = GetFilterValue(doc, element, filter.Property);
                if (direct.Exists)
                {
                    value = direct;
                    return true;
                }

                string spaceValue = GetBuiltInParameterText(element, "RBS_SPACE_NAME_PARAM");
                if (string.IsNullOrWhiteSpace(spaceValue))
                {
                    spaceValue = GetBuiltInParameterText(element, "RBS_SPACE_NUMBER_PARAM");
                }

                value = CreateTextFilterValue(spaceValue);
                return true;
            }

            if (filterCategoryNormalized == NormalizeForCompare("Redundant"))
            {
                reason = "Filter category 'Redundant' is not implemented yet.";
                value = new FilterValue();
                return false;
            }

            value = new FilterValue();
            return false;
        }

        private static FilterValue CreateTextFilterValue(string text)
        {
            string value = text ?? string.Empty;
            return new FilterValue
            {
                Exists = !string.IsNullOrWhiteSpace(value),
                HasValue = !string.IsNullOrWhiteSpace(value),
                Text = value
            };
        }

        private static bool TryGetBuiltInParameter(Element element, string builtInParameterName, out Parameter parameter)
        {
            parameter = null!;

            if (!Enum.TryParse(builtInParameterName, ignoreCase: false, out BuiltInParameter builtInParameter))
            {
                return false;
            }

            try
            {
                parameter = element.get_Parameter(builtInParameter);
            }
            catch
            {
                parameter = null!;
            }

            return parameter != null;
        }

        private static string GetBuiltInParameterText(Element element, string builtInParameterName)
        {
            if (!TryGetBuiltInParameter(element, builtInParameterName, out Parameter parameter))
            {
                return string.Empty;
            }

            return parameter.AsString() ?? parameter.AsValueString() ?? string.Empty;
        }

        private static string ResolveElementNameByIdParameter(Document doc, Element element, string builtInParameterName)
        {
            if (!TryGetBuiltInParameter(element, builtInParameterName, out Parameter parameter))
            {
                return string.Empty;
            }

            if (parameter.StorageType == StorageType.ElementId)
            {
                ElementId id = parameter.AsElementId();
                if (id != null && id != ElementId.InvalidElementId)
                {
                    return doc.GetElement(id)?.Name ?? parameter.AsValueString() ?? string.Empty;
                }
            }

            return parameter.AsString() ?? parameter.AsValueString() ?? string.Empty;
        }

        private static string ResolveLevelName(Document doc, Element element)
        {
            ElementId levelId = GetElementIdPropertyValue(element, "LevelId");
            if (levelId != ElementId.InvalidElementId)
            {
                return doc.GetElement(levelId)?.Name ?? string.Empty;
            }

            string fromPrimary = ResolveElementNameByIdParameter(doc, element, "LEVEL_PARAM");
            if (!string.IsNullOrWhiteSpace(fromPrimary))
            {
                return fromPrimary;
            }

            string fromFamily = ResolveElementNameByIdParameter(doc, element, "FAMILY_LEVEL_PARAM");
            if (!string.IsNullOrWhiteSpace(fromFamily))
            {
                return fromFamily;
            }

            return ResolveElementNameByIdParameter(doc, element, "SCHEDULE_LEVEL_PARAM");
        }

        private static ElementId GetElementIdPropertyValue(object instance, string propertyName)
        {
            if (instance == null)
            {
                return ElementId.InvalidElementId;
            }

            PropertyInfo property = instance.GetType().GetProperty(propertyName);
            if (property == null || property.PropertyType != typeof(ElementId))
            {
                return ElementId.InvalidElementId;
            }

            object value = property.GetValue(instance);
            return value is ElementId id ? id : ElementId.InvalidElementId;
        }

        private static string ResolveDesignOptionName(Document doc, Element element)
        {
            string fromParam = ResolveElementNameByIdParameter(doc, element, "DESIGN_OPTION_ID");
            if (!string.IsNullOrWhiteSpace(fromParam))
            {
                return fromParam;
            }

            Parameter parameter = element.LookupParameter("Design Option");
            return parameter?.AsString() ?? parameter?.AsValueString() ?? string.Empty;
        }

        private static string ResolveOwnerViewName(Document doc, Element element)
        {
            if (element is View view)
            {
                return view.Name ?? string.Empty;
            }

            ElementId ownerViewId = element.OwnerViewId;
            if (ownerViewId != ElementId.InvalidElementId)
            {
                return doc.GetElement(ownerViewId)?.Name ?? string.Empty;
            }

            return GetBuiltInParameterText(element, "VIEW_NAME");
        }

        private static string ResolveStructuralTypeName(Element element)
        {
            if (element is FamilyInstance familyInstance)
            {
                return familyInstance.StructuralType.ToString();
            }

            return GetBuiltInParameterText(element, "INSTANCE_STRUCT_USAGE_PARAM");
        }

        private static Element? GetHostElement(Document doc, Element element)
        {
            if (element is FamilyInstance familyInstance && familyInstance.Host != null)
            {
                return familyInstance.Host;
            }

            if (TryGetBuiltInParameter(element, "HOST_ID_PARAM", out Parameter hostParam) &&
                hostParam.StorageType == StorageType.ElementId)
            {
                ElementId hostId = hostParam.AsElementId();
                if (hostId != null && hostId != ElementId.InvalidElementId)
                {
                    return doc.GetElement(hostId);
                }
            }

            ElementId reflectedHostId = GetElementIdPropertyValue(element, "HostId");
            if (reflectedHostId != ElementId.InvalidElementId)
            {
                return doc.GetElement(reflectedHostId);
            }

            return null;
        }

        private static string GetFilterDuplicateContextKey(ModelCheckerFilterDef filter)
        {
            return string.Join(
                "|",
                NormalizeForCompare(filter.FilterCategory),
                NormalizeForCompare(filter.Property),
                filter.CaseInsensitive ? "CI1" : "CI0");
        }

        private static string GetDuplicateComparableToken(FilterValue value, bool caseInsensitive)
        {
            if (!value.Exists || !value.HasValue)
            {
                return string.Empty;
            }

            if (value.Boolean.HasValue)
            {
                return value.Boolean.Value ? "B:1" : "B:0";
            }

            if (value.Number.HasValue)
            {
                return "N:" + value.Number.Value.ToString("R", CultureInfo.InvariantCulture);
            }

            string text = value.Text ?? string.Empty;
            if (string.IsNullOrWhiteSpace(text))
            {
                return string.Empty;
            }

            return caseInsensitive
                ? "S:" + NormalizeForCompare(text)
                : "S:" + text.Trim();
        }

        private sealed class FilterValue
        {
            public bool Exists { get; set; }
            public bool HasValue { get; set; }
            public string Text { get; set; } = string.Empty;
            public double? Number { get; set; }
            public bool? Boolean { get; set; }
        }

        private static FilterValue GetFilterValue(Document doc, Element element, string property)
        {
            string normalized = NormalizeForCompare(property);

            if (normalized == NormalizeForCompare("Family Name") ||
                normalized == NormalizeForCompare("FamilyName") ||
                normalized == NormalizeForCompare("Nombre Familia"))
            {
                string familyName = ResolveFamilyName(doc, element);
                return new FilterValue
                {
                    Exists = !string.IsNullOrWhiteSpace(familyName),
                    HasValue = !string.IsNullOrWhiteSpace(familyName),
                    Text = familyName
                };
            }

            if (normalized == NormalizeForCompare("Type Name") ||
                normalized == NormalizeForCompare("TypeName") ||
                normalized == NormalizeForCompare("Nombre Tipo"))
            {
                string typeName = ResolveTypeName(doc, element);
                return new FilterValue
                {
                    Exists = !string.IsNullOrWhiteSpace(typeName) && typeName != "-",
                    HasValue = !string.IsNullOrWhiteSpace(typeName) && typeName != "-",
                    Text = typeName
                };
            }

            if (normalized == NormalizeForCompare("Workset Name") ||
                normalized == NormalizeForCompare("Workset") ||
                normalized == NormalizeForCompare("Subproject") ||
                normalized == NormalizeForCompare("Subproyecto"))
            {
                string worksetName = ResolveWorksetName(doc, element);
                return new FilterValue
                {
                    Exists = !string.IsNullOrWhiteSpace(worksetName),
                    HasValue = !string.IsNullOrWhiteSpace(worksetName),
                    Text = worksetName
                };
            }

            if (normalized == NormalizeForCompare("Category") ||
                normalized == NormalizeForCompare("Category Name"))
            {
                string categoryName = element.Category?.Name ?? string.Empty;
                return new FilterValue
                {
                    Exists = !string.IsNullOrWhiteSpace(categoryName),
                    HasValue = !string.IsNullOrWhiteSpace(categoryName),
                    Text = categoryName
                };
            }

            if (normalized == NormalizeForCompare("Element Id") ||
                normalized == NormalizeForCompare("Id"))
            {
                string idText = element.Id.Value.ToString(CultureInfo.InvariantCulture);
                return new FilterValue
                {
                    Exists = true,
                    HasValue = true,
                    Text = idText,
                    Number = element.Id.Value
                };
            }

            if (normalized == NormalizeForCompare("Name") ||
                normalized == NormalizeForCompare("Element Name") ||
                normalized == NormalizeForCompare("Nombre"))
            {
                string name = element.Name ?? string.Empty;
                return new FilterValue
                {
                    Exists = !string.IsNullOrWhiteSpace(name),
                    HasValue = !string.IsNullOrWhiteSpace(name),
                    Text = name
                };
            }

            Parameter? parameter = FindParameterByExactName(element, property);
            if (parameter == null)
            {
                ElementId typeId = element.GetTypeId();
                if (typeId != ElementId.InvalidElementId)
                {
                    Element typeElement = doc.GetElement(typeId);
                    if (typeElement != null)
                    {
                        parameter = FindParameterByExactName(typeElement, property);
                    }
                }
            }

            if (parameter == null)
            {
                return new FilterValue
                {
                    Exists = false,
                    HasValue = false,
                    Text = string.Empty
                };
            }

            var value = new FilterValue
            {
                Exists = true,
                HasValue = parameter.HasValue
            };

            switch (parameter.StorageType)
            {
                case StorageType.String:
                    value.Text = parameter.AsString() ?? parameter.AsValueString() ?? string.Empty;
                    value.HasValue = !string.IsNullOrWhiteSpace(value.Text);
                    break;
                case StorageType.ElementId:
                    ElementId id = parameter.AsElementId();
                    value.Text = id == ElementId.InvalidElementId ? string.Empty : id.Value.ToString(CultureInfo.InvariantCulture);
                    value.Number = id == ElementId.InvalidElementId ? null : id.Value;
                    value.HasValue = id != ElementId.InvalidElementId;
                    break;
                case StorageType.Double:
                    double dbl = parameter.AsDouble();
                    value.Number = dbl;
                    value.Text = parameter.AsValueString() ?? dbl.ToString(CultureInfo.InvariantCulture);
                    value.HasValue = true;
                    break;
                case StorageType.Integer:
                    int integer = parameter.AsInteger();
                    value.Number = integer;
                    value.Text = parameter.AsValueString() ?? integer.ToString(CultureInfo.InvariantCulture);
                    value.Boolean = integer == 0 ? false : integer == 1 ? true : (bool?)null;
                    value.HasValue = true;
                    break;
                default:
                    value.Text = parameter.AsValueString() ?? string.Empty;
                    value.HasValue = value.HasValue && !string.IsNullOrWhiteSpace(value.Text);
                    break;
            }

            if (!value.Boolean.HasValue && TryParseBoolean(value.Text, out bool parsedBool))
            {
                value.Boolean = parsedBool;
            }

            if (!value.Number.HasValue && TryParseDoubleFlexible(value.Text, out double parsedDouble))
            {
                value.Number = parsedDouble;
            }

            return value;
        }

        private static bool TryEvaluateParameterCondition(
            Document doc,
            Element element,
            CheckEvaluationContext context,
            FilterValue value,
            ModelCheckerFilterDef filter,
            out bool result,
            out string reason)
        {
            result = false;
            reason = string.Empty;

            string condition = NormalizeForCompare(filter.Condition);
            bool ignoreCase = filter.CaseInsensitive;
            string expected = filter.Value ?? string.Empty;

            if (condition == NormalizeForCompare("Duplicated"))
            {
                string contextKey = GetFilterDuplicateContextKey(filter);
                if (!context.DuplicateTokensByFilterKey.TryGetValue(contextKey, out HashSet<string> duplicates))
                {
                    result = false;
                    return true;
                }

                string token = GetDuplicateComparableToken(value, filter.CaseInsensitive);
                result = !string.IsNullOrWhiteSpace(token) && duplicates.Contains(token);
                return true;
            }

            if (IsNoValueCondition(condition))
            {
                // Stricter behavior for PEB checks: if parameter does not exist, treat as no-value failure.
                result = !value.Exists || !value.HasValue;
                return true;
            }

            if (condition == NormalizeForCompare("HasValue") ||
                condition == NormalizeForCompare("NotEmpty"))
            {
                result = value.Exists && value.HasValue;
                return true;
            }

            if (condition == NormalizeForCompare("Exists") ||
                condition == NormalizeForCompare("HasParameter") ||
                condition == NormalizeForCompare("HasProperty") ||
                condition == NormalizeForCompare("Defined"))
            {
                result = value.Exists;
                return true;
            }

            if (condition == NormalizeForCompare("DoesNotExist") ||
                condition == NormalizeForCompare("NotExists") ||
                condition == NormalizeForCompare("Missing") ||
                condition == NormalizeForCompare("Undefined"))
            {
                result = !value.Exists;
                return true;
            }

            if (condition == NormalizeForCompare("Equal") ||
                condition == NormalizeForCompare("Equals"))
            {
                if (!TryCompareEqual(value, expected, ignoreCase, filter, out bool equalResult, out string compareReason))
                {
                    reason = compareReason;
                    return false;
                }

                result = equalResult;
                return true;
            }

            if (condition == NormalizeForCompare("NotEqual") ||
                condition == NormalizeForCompare("NotEquals") ||
                condition == NormalizeForCompare("Unequal"))
            {
                if (!TryCompareEqual(value, expected, ignoreCase, filter, out bool equalResult, out string compareReason))
                {
                    reason = compareReason;
                    return false;
                }

                result = !equalResult;
                return true;
            }

            if (condition == NormalizeForCompare("Included"))
            {
                if (!TryCompareEqual(value, expected, ignoreCase, filter, out bool equalResult, out string compareReason))
                {
                    reason = compareReason;
                    return false;
                }

                result = equalResult;
                return true;
            }

            if (condition == NormalizeForCompare("Contains"))
            {
                result = value.Exists && ContainsText(value.Text, expected, ignoreCase);
                return true;
            }

            if (condition == NormalizeForCompare("NotContains") ||
                condition == NormalizeForCompare("DoesNotContain"))
            {
                result = value.Exists && !ContainsText(value.Text, expected, ignoreCase);
                return true;
            }

            if (condition == NormalizeForCompare("StartsWith"))
            {
                result = value.Exists && StartsWithText(value.Text, expected, ignoreCase);
                return true;
            }

            if (condition == NormalizeForCompare("NotStartsWith") ||
                condition == NormalizeForCompare("DoesNotStartWith"))
            {
                result = value.Exists && !StartsWithText(value.Text, expected, ignoreCase);
                return true;
            }

            if (condition == NormalizeForCompare("EndsWith"))
            {
                result = value.Exists && EndsWithText(value.Text, expected, ignoreCase);
                return true;
            }

            if (condition == NormalizeForCompare("NotEndsWith") ||
                condition == NormalizeForCompare("DoesNotEndWith"))
            {
                result = value.Exists && !EndsWithText(value.Text, expected, ignoreCase);
                return true;
            }

            if (condition == NormalizeForCompare("MatchesRegex") ||
                condition == NormalizeForCompare("Regex") ||
                condition == NormalizeForCompare("WildCard"))
            {
                try
                {
                    RegexOptions options = RegexOptions.CultureInvariant;
                    if (ignoreCase)
                    {
                        options |= RegexOptions.IgnoreCase;
                    }

                    string pattern = condition == NormalizeForCompare("WildCard")
                        ? WildcardToRegex(expected)
                        : (expected ?? string.Empty);
                    result = value.Exists && Regex.IsMatch(value.Text ?? string.Empty, pattern, options);
                    return true;
                }
                catch (Exception ex)
                {
                    reason = $"Invalid pattern '{expected}': {ex.Message}";
                    return false;
                }
            }

            if (condition == NormalizeForCompare("NotMatchesRegex") ||
                condition == NormalizeForCompare("DoesNotMatchRegex") ||
                condition == NormalizeForCompare("WildCardNoMatch"))
            {
                try
                {
                    RegexOptions options = RegexOptions.CultureInvariant;
                    if (ignoreCase)
                    {
                        options |= RegexOptions.IgnoreCase;
                    }

                    string pattern = condition == NormalizeForCompare("WildCardNoMatch")
                        ? WildcardToRegex(expected)
                        : (expected ?? string.Empty);
                    result = value.Exists && !Regex.IsMatch(value.Text ?? string.Empty, pattern, options);
                    return true;
                }
                catch (Exception ex)
                {
                    reason = $"Invalid pattern '{expected}': {ex.Message}";
                    return false;
                }
            }

            if (condition == NormalizeForCompare("MatchesParameter"))
            {
                FilterValue other = GetFilterValue(doc, element, expected);
                result = value.Exists && other.Exists && CompareFilterValues(value, other, ignoreCase);
                return true;
            }

            if (condition == NormalizeForCompare("DoesNotMatchParameter"))
            {
                FilterValue other = GetFilterValue(doc, element, expected);
                result = value.Exists && other.Exists && !CompareFilterValues(value, other, ignoreCase);
                return true;
            }

            if (condition == NormalizeForCompare("MatchesHostParameter"))
            {
                Element? host = GetHostElement(doc, element);
                if (host == null)
                {
                    result = false;
                    return true;
                }

                FilterValue other = GetFilterValue(doc, host, expected);
                result = value.Exists && other.Exists && CompareFilterValues(value, other, ignoreCase);
                return true;
            }

            if (condition == NormalizeForCompare("DoesNotMatchHostParameter"))
            {
                Element? host = GetHostElement(doc, element);
                if (host == null)
                {
                    result = false;
                    return true;
                }

                FilterValue other = GetFilterValue(doc, host, expected);
                result = value.Exists && other.Exists && !CompareFilterValues(value, other, ignoreCase);
                return true;
            }

            if (condition == NormalizeForCompare("In") ||
                condition == NormalizeForCompare("InList") ||
                condition == NormalizeForCompare("OneOf") ||
                condition == NormalizeForCompare("AnyOf"))
            {
                List<string> values = SplitList(expected).ToList();
                result = values.Any(v => CompareEqual(value, v, ignoreCase));
                return true;
            }

            if (condition == NormalizeForCompare("NotIn") ||
                condition == NormalizeForCompare("NotInList"))
            {
                List<string> values = SplitList(expected).ToList();
                result = values.Count > 0 && values.All(v => !CompareEqual(value, v, ignoreCase));
                return true;
            }

            if (condition == NormalizeForCompare("GreaterThan") ||
                condition == NormalizeForCompare("Greater"))
            {
                if (!TryGetExpectedNumeric(filter, out double expectedNumber, out string numericReason))
                {
                    reason = numericReason;
                    return false;
                }

                result = value.Number.HasValue && value.Number.Value > expectedNumber;
                return true;
            }

            if (condition == NormalizeForCompare("GreaterOrEqual") ||
                condition == NormalizeForCompare("GreaterThanOrEqual"))
            {
                if (!TryGetExpectedNumeric(filter, out double expectedNumber, out string numericReason))
                {
                    reason = numericReason;
                    return false;
                }

                result = value.Number.HasValue && value.Number.Value >= expectedNumber;
                return true;
            }

            if (condition == NormalizeForCompare("LessThan") ||
                condition == NormalizeForCompare("Less"))
            {
                if (!TryGetExpectedNumeric(filter, out double expectedNumber, out string numericReason))
                {
                    reason = numericReason;
                    return false;
                }

                result = value.Number.HasValue && value.Number.Value < expectedNumber;
                return true;
            }

            if (condition == NormalizeForCompare("LessOrEqual") ||
                condition == NormalizeForCompare("LessThanOrEqual"))
            {
                if (!TryGetExpectedNumeric(filter, out double expectedNumber, out string numericReason))
                {
                    reason = numericReason;
                    return false;
                }

                result = value.Number.HasValue && value.Number.Value <= expectedNumber;
                return true;
            }

            if (condition == NormalizeForCompare("IsTrue"))
            {
                result = value.Boolean.HasValue && value.Boolean.Value;
                return true;
            }

            if (condition == NormalizeForCompare("IsFalse"))
            {
                result = value.Boolean.HasValue && !value.Boolean.Value;
                return true;
            }

            reason = $"Unsupported Parameter condition '{filter.Condition}'.";
            return false;
        }

        private static bool TryBuildMissingParameterReasonForNoValueCheck(
            Document doc,
            ModelCheckerCheckDef check,
            List<Element> candidates,
            out string reason)
        {
            reason = string.Empty;
            if (candidates == null || candidates.Count == 0)
            {
                return false;
            }

            List<string> noValueParameterNames = check.Filters
                .Where(f =>
                    (string.Equals(f.FilterCategory, "Parameter", StringComparison.OrdinalIgnoreCase) ||
                     string.Equals(f.FilterCategory, "APIParameter", StringComparison.OrdinalIgnoreCase)) &&
                    IsNoValueCondition(NormalizeForCompare(f.Condition)) &&
                    !string.IsNullOrWhiteSpace(f.Property))
                .Select(f => CanonicalizeKnownParameterName(f.Property.Trim()))
                .Distinct(StringComparer.Ordinal)
                .ToList();

            if (noValueParameterNames.Count == 0)
            {
                return false;
            }

            List<string> missingParameters = noValueParameterNames
                .Where(parameterName => !candidates.Any(e => GetFilterValue(doc, e, parameterName).Exists))
                .ToList();

            if (missingParameters.Count == 0)
            {
                return false;
            }

            if (missingParameters.Count == 1)
            {
                reason = $"No se encontro parametro '{missingParameters[0]}' en la categoria evaluada.";
                return true;
            }

            reason = "No se encontraron parametros en la categoria evaluada: "
                + string.Join(", ", missingParameters.Select(v => $"'{v}'")) + ".";
            return true;
        }

        private static bool TryCompareEqual(
            FilterValue actual,
            string expectedRaw,
            bool ignoreCase,
            ModelCheckerFilterDef filter,
            out bool result,
            out string reason)
        {
            reason = string.Empty;
            result = false;

            if (!actual.Exists)
            {
                return true;
            }

            string expected = expectedRaw ?? string.Empty;

            if (actual.Boolean.HasValue && TryParseBoolean(expected, out bool expectedBool))
            {
                result = actual.Boolean.Value == expectedBool;
                return true;
            }

            if (actual.Number.HasValue)
            {
                if (TryGetExpectedNumeric(filter, out double expectedNumeric, out string numericReason))
                {
                    result = Math.Abs(actual.Number.Value - expectedNumeric) < 0.0000001;
                    return true;
                }

                if (HasConfiguredUnits(filter))
                {
                    reason = numericReason;
                    return false;
                }
            }

            result = CompareText(actual.Text, expected, ignoreCase);
            return true;
        }

        private static bool TryGetExpectedNumeric(ModelCheckerFilterDef filter, out double expectedNumber, out string reason)
        {
            expectedNumber = 0;
            reason = string.Empty;

            string raw = filter.Value ?? string.Empty;
            if (!TryParseDoubleFlexible(raw, out double parsed))
            {
                Match match = Regex.Match(raw, @"[-+]?\d+(?:[.,]\d+)?", RegexOptions.CultureInvariant);
                if (match.Success)
                {
                    if (!TryParseDoubleFlexible(match.Value, out parsed))
                    {
                        reason = $"Condition '{filter.Condition}' requires numeric value. Received '{raw}'.";
                        return false;
                    }
                }
                else
                {
                    reason = $"Condition '{filter.Condition}' requires numeric value. Received '{raw}'.";
                    return false;
                }
            }

            if (!TryConvertConfiguredUnitToInternal(parsed, filter.UnitClass, filter.Unit, out expectedNumber, out string unitReason))
            {
                reason = unitReason;
                return false;
            }

            return true;
        }

        private static bool HasConfiguredUnits(ModelCheckerFilterDef filter)
        {
            string unit = NormalizeUnitToken(filter.Unit);
            string unitClass = NormalizeUnitToken(filter.UnitClass);
            return !(string.IsNullOrWhiteSpace(unit) || unit == "NONE" || unit == "DEFAULT") ||
                !(string.IsNullOrWhiteSpace(unitClass) || unitClass == "NONE");
        }

        private static bool TryConvertConfiguredUnitToInternal(
            double value,
            string unitClassRaw,
            string unitRaw,
            out double converted,
            out string reason)
        {
            converted = value;
            reason = string.Empty;

            string unitClass = NormalizeUnitToken(unitClassRaw);
            string unit = NormalizeUnitToken(unitRaw);

            if (string.IsNullOrWhiteSpace(unitClass) || unitClass == "NONE")
            {
                unitClass = InferUnitClassFromUnit(unit);
            }

            if (string.IsNullOrWhiteSpace(unit) || unit == "NONE" || unit == "DEFAULT")
            {
                converted = value;
                return true;
            }

            if (unitClass == "NONE")
            {
                converted = value;
                return true;
            }

            if (unitClass == "LENGTH")
            {
                switch (unit)
                {
                    case "FEET":
                    case "FOOT":
                    case "FT":
                        converted = value;
                        return true;
                    case "INCHES":
                    case "INCH":
                    case "IN":
                        converted = value / 12.0;
                        return true;
                    case "METERS":
                    case "METER":
                    case "M":
                        converted = value / 0.3048;
                        return true;
                    case "DECIMETERS":
                    case "DECIMETER":
                    case "DECIMENTERS":
                    case "DM":
                        converted = value / 3.048;
                        return true;
                    case "CENTIMETERS":
                    case "CENTIMETER":
                    case "CM":
                        converted = value / 30.48;
                        return true;
                    case "MILLIMETERS":
                    case "MILLIMETER":
                    case "MM":
                        converted = value / 304.8;
                        return true;
                }
            }
            else if (unitClass == "AREA")
            {
                switch (unit)
                {
                    case "SQUAREFEET":
                    case "SQFT":
                    case "FT2":
                        converted = value;
                        return true;
                    case "SQUAREINCHES":
                    case "IN2":
                        converted = value / 144.0;
                        return true;
                    case "SQUAREMETERS":
                    case "M2":
                        converted = value * 10.76391041671;
                        return true;
                    case "SQUARECENTIMETERS":
                    case "CM2":
                        converted = value / 929.0304;
                        return true;
                    case "SQUAREMILLIMETERS":
                    case "MM2":
                        converted = value / 929030.4;
                        return true;
                    case "ACRES":
                        converted = value * 43560.0;
                        return true;
                    case "HECTARES":
                    case "HECTARE":
                    case "HECTACRES":
                        converted = value * 107639.1041671;
                        return true;
                }
            }
            else if (unitClass == "VOLUME")
            {
                switch (unit)
                {
                    case "CUBICFEET":
                    case "FT3":
                        converted = value;
                        return true;
                    case "CUBICYARDS":
                    case "YD3":
                        converted = value * 27.0;
                        return true;
                    case "CUBICINCHES":
                    case "IN3":
                        converted = value / 1728.0;
                        return true;
                    case "CUBICMETERS":
                    case "M3":
                        converted = value * 35.31466672149;
                        return true;
                    case "CUBICCENTIMETERS":
                    case "CM3":
                        converted = value / 28316.846592;
                        return true;
                    case "CUBICMILLIMETERS":
                    case "MM3":
                        converted = value / 28316846.592;
                        return true;
                    case "LITERS":
                    case "LITER":
                    case "L":
                        converted = value / 28.316846592;
                        return true;
                    case "USGALLONS":
                    case "GALLONS":
                    case "GAL":
                        converted = value / 7.48051948;
                        return true;
                }
            }
            else if (unitClass == "ANGLE")
            {
                switch (unit)
                {
                    case "RADIANS":
                    case "RAD":
                        converted = value;
                        return true;
                    case "DEGREES":
                    case "DEG":
                        converted = value * Math.PI / 180.0;
                        return true;
                    case "GRADS":
                    case "GRAD":
                        converted = value * Math.PI / 200.0;
                        return true;
                }
            }

            reason = $"Unsupported unit conversion UnitClass='{unitClassRaw}' Unit='{unitRaw}'.";
            return false;
        }

        private static string InferUnitClassFromUnit(string unit)
        {
            if (string.IsNullOrWhiteSpace(unit) || unit == "NONE" || unit == "DEFAULT")
            {
                return "NONE";
            }

            if (unit.StartsWith("SQUARE", StringComparison.Ordinal) || unit.EndsWith("2", StringComparison.Ordinal))
            {
                return "AREA";
            }

            if (unit.StartsWith("CUBIC", StringComparison.Ordinal) ||
                unit.EndsWith("3", StringComparison.Ordinal) ||
                unit == "LITERS" ||
                unit == "LITER" ||
                unit == "L" ||
                unit == "USGALLONS" ||
                unit == "GALLONS" ||
                unit == "GAL")
            {
                return "VOLUME";
            }

            if (unit == "DEGREES" || unit == "DEG" || unit == "RADIANS" || unit == "RAD" || unit == "GRADS" || unit == "GRAD")
            {
                return "ANGLE";
            }

            return "LENGTH";
        }

        private static string NormalizeUnitToken(string input)
        {
            return NormalizeForCompare(input)
                .Replace(" ", string.Empty)
                .Replace("_", string.Empty)
                .Replace("-", string.Empty);
        }

        private static bool CompareEqual(FilterValue actual, string expectedRaw, bool ignoreCase)
        {
            if (!actual.Exists)
            {
                return false;
            }

            string expected = expectedRaw ?? string.Empty;

            if (actual.Boolean.HasValue && TryParseBoolean(expected, out bool expectedBool))
            {
                return actual.Boolean.Value == expectedBool;
            }

            if (actual.Number.HasValue && TryParseDoubleFlexible(expected, out double expectedNumber))
            {
                return Math.Abs(actual.Number.Value - expectedNumber) < 0.0000001;
            }

            return CompareText(actual.Text, expected, ignoreCase);
        }

        private static bool CompareFilterValues(FilterValue left, FilterValue right, bool ignoreCase)
        {
            if (!left.Exists || !right.Exists)
            {
                return false;
            }

            if (left.Boolean.HasValue && right.Boolean.HasValue)
            {
                return left.Boolean.Value == right.Boolean.Value;
            }

            if (left.Number.HasValue && right.Number.HasValue)
            {
                return Math.Abs(left.Number.Value - right.Number.Value) < 0.0000001;
            }

            return CompareText(left.Text, right.Text, ignoreCase);
        }

        private static string WildcardToRegex(string wildcardPattern)
        {
            string pattern = wildcardPattern ?? string.Empty;
            return "^" + Regex.Escape(pattern)
                .Replace("\\*", ".*")
                .Replace("\\?", ".") + "$";
        }

        private static bool CompareText(string left, string right, bool ignoreCase)
        {
            string l = left ?? string.Empty;
            string r = right ?? string.Empty;
            StringComparison comparison = ignoreCase
                ? StringComparison.OrdinalIgnoreCase
                : StringComparison.Ordinal;
            return string.Equals(l, r, comparison);
        }

        private static bool ContainsText(string source, string fragment, bool ignoreCase)
        {
            string src = source ?? string.Empty;
            string part = fragment ?? string.Empty;
            if (part.Length == 0)
            {
                return false;
            }

            StringComparison comparison = ignoreCase
                ? StringComparison.OrdinalIgnoreCase
                : StringComparison.Ordinal;
            return src.IndexOf(part, comparison) >= 0;
        }

        private static bool StartsWithText(string source, string prefix, bool ignoreCase)
        {
            string src = source ?? string.Empty;
            string pre = prefix ?? string.Empty;
            if (pre.Length == 0)
            {
                return false;
            }

            StringComparison comparison = ignoreCase
                ? StringComparison.OrdinalIgnoreCase
                : StringComparison.Ordinal;
            return src.StartsWith(pre, comparison);
        }

        private static bool EndsWithText(string source, string suffix, bool ignoreCase)
        {
            string src = source ?? string.Empty;
            string suf = suffix ?? string.Empty;
            if (suf.Length == 0)
            {
                return false;
            }

            StringComparison comparison = ignoreCase
                ? StringComparison.OrdinalIgnoreCase
                : StringComparison.Ordinal;
            return src.EndsWith(suf, comparison);
        }

        private static bool TryParseDoubleFlexible(string input, out double value)
        {
            value = 0;
            if (string.IsNullOrWhiteSpace(input))
            {
                return false;
            }

            string candidate = input.Trim();
            const NumberStyles styles = NumberStyles.Float | NumberStyles.AllowThousands;

            if (double.TryParse(candidate, styles, CultureInfo.InvariantCulture, out value))
            {
                return true;
            }

            if (double.TryParse(candidate, styles, CultureInfo.CurrentCulture, out value))
            {
                return true;
            }

            string noSpace = candidate.Replace(" ", string.Empty);
            string commaToDot = noSpace.Replace(",", ".");
            if (double.TryParse(commaToDot, styles, CultureInfo.InvariantCulture, out value))
            {
                return true;
            }

            string dotToComma = noSpace.Replace(".", ",");
            return double.TryParse(dotToComma, styles, CultureInfo.GetCultureInfo("es-PE"), out value);
        }

        private static bool TryParseBoolean(string input, out bool value)
        {
            value = false;
            if (string.IsNullOrWhiteSpace(input))
            {
                return false;
            }

            string trimmed = input.Trim();
            if (bool.TryParse(trimmed, out bool parsed))
            {
                value = parsed;
                return true;
            }

            string normalized = NormalizeForCompare(trimmed);
            if (normalized == "1" ||
                normalized == "Y" ||
                normalized == "YES" ||
                normalized == "SI" ||
                normalized == "S" ||
                normalized == "TRUE" ||
                normalized == "VERDADERO" ||
                normalized == "VERDADERA")
            {
                value = true;
                return true;
            }

            if (normalized == "0" ||
                normalized == "N" ||
                normalized == "NO" ||
                normalized == "FALSE" ||
                normalized == "FALSO" ||
                normalized == "FALSA")
            {
                value = false;
                return true;
            }

            return false;
        }

        private static IEnumerable<string> SplitList(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw))
            {
                return Enumerable.Empty<string>();
            }

            return raw
                .Split(new[] { ';', ',', '|', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(v => v.Trim())
                .Where(v => v.Length > 0);
        }

        private static string ResolveFamilyName(Document doc, Element element)
        {
            if (element is Family family)
            {
                return family.Name ?? string.Empty;
            }

            if (element is ElementType elementType && !string.IsNullOrWhiteSpace(elementType.FamilyName))
            {
                return elementType.FamilyName;
            }

            ElementId typeId = element.GetTypeId();
            if (typeId != ElementId.InvalidElementId)
            {
                ElementType? type = doc.GetElement(typeId) as ElementType;
                if (type != null && !string.IsNullOrWhiteSpace(type.FamilyName))
                {
                    return type.FamilyName;
                }
            }

            return string.Empty;
        }

        private static string ResolveWorksetName(Document doc, Element element)
        {
            if (element.WorksetId == null || element.WorksetId == WorksetId.InvalidWorksetId)
            {
                return string.Empty;
            }

            WorksetTable table = doc.GetWorksetTable();
            if (table == null)
            {
                return string.Empty;
            }

            Workset workset = table.GetWorkset(element.WorksetId);
            return workset?.Name ?? string.Empty;
        }

        private static string ResolveTypeName(Document doc, Element element)
        {
            if (element is ElementType elementType)
            {
                return elementType.Name ?? "-";
            }

            ElementId typeId = element.GetTypeId();
            if (typeId == ElementId.InvalidElementId)
            {
                return "-";
            }

            ElementType? type = doc.GetElement(typeId) as ElementType;
            return type?.Name ?? "-";
        }

        private static string WriteModelCheckerSummaryCsv(string outputDir, List<ModelCheckerSummaryRow> rows, string xmlPath)
        {
            string reportPath = Path.Combine(outputDir, "04_modelchecker_summary.csv");
            using (var writer = new StreamWriter(reportPath, false, new UTF8Encoding(true)))
            {
                writer.WriteLine($"\"XML Source\",{EscapeCsv(xmlPath)}");
                writer.WriteLine("Heading,Section,CheckName,Status,CandidateElements,FailedElements,Reason");

                foreach (ModelCheckerSummaryRow row in rows)
                {
                    writer.WriteLine(string.Join(
                        ",",
                        EscapeCsv(row.Heading),
                        EscapeCsv(row.Section),
                        EscapeCsv(row.CheckName),
                        EscapeCsv(row.Status),
                        EscapeCsv(row.CandidateElements.ToString(CultureInfo.InvariantCulture)),
                        EscapeCsv(row.FailedElements.ToString(CultureInfo.InvariantCulture)),
                        EscapeCsv(row.Reason)));
                }
            }

            return reportPath;
        }

        private static string WriteModelCheckerFailuresCsv(string outputDir, List<ModelCheckerFailureRow> rows)
        {
            string reportPath = Path.Combine(outputDir, "05_modelchecker_failed_elements.csv");
            using (var writer = new StreamWriter(reportPath, false, new UTF8Encoding(true)))
            {
                writer.WriteLine("Heading,Section,CheckName,ElementId,Category,FamilyOrType,ElementName");

                foreach (ModelCheckerFailureRow row in rows)
                {
                    writer.WriteLine(string.Join(
                        ",",
                        EscapeCsv(row.Heading),
                        EscapeCsv(row.Section),
                        EscapeCsv(row.CheckName),
                        EscapeCsv(row.ElementId.ToString(CultureInfo.InvariantCulture)),
                        EscapeCsv(row.Category),
                        EscapeCsv(row.FamilyOrType),
                        EscapeCsv(row.Name)));
                }
            }

            return reportPath;
        }

        private static string WriteAuditWorkbook(
            string outputDir,
            Document doc,
            List<ParameterAuditRow> parameterRows,
            List<FamilyNamingRow> familyRows,
            List<SubprojectNamingRow> subprojectRows,
            List<ModelCheckerSummaryRow> modelCheckerSummary,
            List<ModelCheckerFailureRow> modelCheckerFailures,
            string modelCheckerXmlPath)
        {
            string reportPath = Path.Combine(outputDir, "JeiAudit_Results.xlsx");

            using (var workbook = new XLWorkbook())
            {
                WriteMetaSheet(workbook, doc, outputDir, modelCheckerXmlPath);
                WriteParameterSheet(workbook, parameterRows);
                WriteFamilySheet(workbook, familyRows);
                WriteSubprojectSheet(workbook, subprojectRows);
                WriteModelCheckerSummarySheet(workbook, modelCheckerSummary, modelCheckerXmlPath);
                WriteModelCheckerFailuresSheet(workbook, modelCheckerFailures, modelCheckerXmlPath);
                List<QualityCheckFactRow> qualityFacts = BuildQualityCheckFacts(parameterRows, familyRows, subprojectRows, modelCheckerSummary);
                WriteKpiSummarySheet(workbook, qualityFacts, modelCheckerFailures, doc, modelCheckerXmlPath);
                WriteFactQualityChecksSheet(workbook, qualityFacts);
                WriteFactFailedElementsSheet(workbook, modelCheckerFailures, modelCheckerXmlPath);
                WritePivotBySectionSheet(workbook, modelCheckerSummary);
                WritePivotByCategorySheet(workbook, modelCheckerFailures);
                WritePivotTopFailChecksSheet(workbook, modelCheckerSummary);
                workbook.SaveAs(reportPath);
            }

            return reportPath;
        }

        private static void WriteMetaSheet(XLWorkbook workbook, Document doc, string outputDir, string modelCheckerXmlPath)
        {
            IXLWorksheet sheet = workbook.Worksheets.Add("00_Meta");
            string[] headers = { "Field", "Value" };
            WriteHeaderRow(sheet, headers);

            int row = 2;
            WriteMetaRow(sheet, row++, "GeneratedAtLocal", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture));
            WriteMetaRow(sheet, row++, "ModelTitle", doc.Title ?? string.Empty);
            WriteMetaRow(sheet, row++, "ModelPath", doc.PathName ?? string.Empty);
            WriteMetaRow(sheet, row++, "OutputFolder", outputDir);
            WriteMetaRow(sheet, row++, "ModelCheckerXmlPath", modelCheckerXmlPath ?? string.Empty);
            WriteMetaRow(sheet, row++, "Plugin", "JeiAudit");

            FinalizeWorksheet(sheet, headers.Length, row - 1);
        }

        private static void WriteMetaRow(IXLWorksheet sheet, int row, string field, string value)
        {
            sheet.Cell(row, 1).Value = field;
            sheet.Cell(row, 2).Value = value ?? string.Empty;
        }

        private static void WriteParameterSheet(XLWorkbook workbook, List<ParameterAuditRow> rows)
        {
            IXLWorksheet sheet = workbook.Worksheets.Add("01_Parameter_Existence");
            string[] headers =
            {
                "Parameter",
                "Category",
                "CategoryAvailable",
                "ExistsByExactName",
                "BoundToCategory",
                "BindingTypes",
                "ParameterGuids",
                "SimilarNames",
                "Status"
            };
            WriteHeaderRow(sheet, headers);

            int row = 2;
            foreach (ParameterAuditRow item in rows)
            {
                sheet.Cell(row, 1).Value = item.ParameterName;
                sheet.Cell(row, 2).Value = item.CategoryName;
                sheet.Cell(row, 3).Value = item.CategoryAvailable;
                sheet.Cell(row, 4).Value = item.ExistsByExactName;
                sheet.Cell(row, 5).Value = item.BoundToCategory;
                sheet.Cell(row, 6).Value = item.BindingTypes;
                sheet.Cell(row, 7).Value = item.ParameterGuids;
                sheet.Cell(row, 8).Value = item.SimilarNames;
                sheet.Cell(row, 9).Value = item.Status;
                row++;
            }

            FinalizeWorksheet(sheet, headers.Length, row - 1);
        }

        private static void WriteFamilySheet(XLWorkbook workbook, List<FamilyNamingRow> rows)
        {
            IXLWorksheet sheet = workbook.Worksheets.Add("02_Family_Naming");
            string[] headers =
            {
                "FamilyName",
                "HasARGPrefix",
                "DisciplineCode",
                "DisciplineCodeValid",
                "HasNamePart",
                "Status",
                "Reason"
            };
            WriteHeaderRow(sheet, headers);

            int row = 2;
            foreach (FamilyNamingRow item in rows)
            {
                sheet.Cell(row, 1).Value = item.FamilyName;
                sheet.Cell(row, 2).Value = item.HasArgPrefix;
                sheet.Cell(row, 3).Value = item.DisciplineCode;
                sheet.Cell(row, 4).Value = item.DisciplineCodeValid;
                sheet.Cell(row, 5).Value = item.HasNamePart;
                sheet.Cell(row, 6).Value = item.Status;
                sheet.Cell(row, 7).Value = item.Reason;
                row++;
            }

            FinalizeWorksheet(sheet, headers.Length, row - 1);
        }

        private static void WriteSubprojectSheet(XLWorkbook workbook, List<SubprojectNamingRow> rows)
        {
            IXLWorksheet sheet = workbook.Worksheets.Add("03_Subproject_Naming");
            string[] headers =
            {
                "WorksetName",
                "HasARGPrefix",
                "HasNameAfterPrefix",
                "HasSecondSeparator",
                "Status",
                "Reason"
            };
            WriteHeaderRow(sheet, headers);

            int row = 2;
            foreach (SubprojectNamingRow item in rows)
            {
                sheet.Cell(row, 1).Value = item.WorksetName;
                sheet.Cell(row, 2).Value = item.HasArgPrefix;
                sheet.Cell(row, 3).Value = item.HasNameAfterPrefix;
                sheet.Cell(row, 4).Value = item.HasSecondSeparator;
                sheet.Cell(row, 5).Value = item.Status;
                sheet.Cell(row, 6).Value = item.Reason;
                row++;
            }

            FinalizeWorksheet(sheet, headers.Length, row - 1);
        }

        private static void WriteModelCheckerSummarySheet(XLWorkbook workbook, List<ModelCheckerSummaryRow> rows, string modelCheckerXmlPath)
        {
            IXLWorksheet sheet = workbook.Worksheets.Add("04_MC_Summary");
            string[] headers =
            {
                "XmlSource",
                "Heading",
                "Section",
                "CheckName",
                "Status",
                "CandidateElements",
                "FailedElements",
                "Reason"
            };
            WriteHeaderRow(sheet, headers);

            int row = 2;
            foreach (ModelCheckerSummaryRow item in rows)
            {
                sheet.Cell(row, 1).Value = modelCheckerXmlPath ?? string.Empty;
                sheet.Cell(row, 2).Value = item.Heading;
                sheet.Cell(row, 3).Value = item.Section;
                sheet.Cell(row, 4).Value = item.CheckName;
                sheet.Cell(row, 5).Value = item.Status;
                sheet.Cell(row, 6).Value = item.CandidateElements;
                sheet.Cell(row, 7).Value = item.FailedElements;
                sheet.Cell(row, 8).Value = item.Reason;
                row++;
            }

            FinalizeWorksheet(sheet, headers.Length, row - 1);
        }

        private static void WriteModelCheckerFailuresSheet(XLWorkbook workbook, List<ModelCheckerFailureRow> rows, string modelCheckerXmlPath)
        {
            IXLWorksheet sheet = workbook.Worksheets.Add("05_MC_Failed_Elements");
            string[] headers =
            {
                "XmlSource",
                "Heading",
                "Section",
                "CheckName",
                "ElementId",
                "Category",
                "FamilyOrType",
                "ElementName"
            };
            WriteHeaderRow(sheet, headers);

            int row = 2;
            foreach (ModelCheckerFailureRow item in rows)
            {
                sheet.Cell(row, 1).Value = modelCheckerXmlPath ?? string.Empty;
                sheet.Cell(row, 2).Value = item.Heading;
                sheet.Cell(row, 3).Value = item.Section;
                sheet.Cell(row, 4).Value = item.CheckName;
                sheet.Cell(row, 5).Value = item.ElementId;
                sheet.Cell(row, 6).Value = item.Category;
                sheet.Cell(row, 7).Value = item.FamilyOrType;
                sheet.Cell(row, 8).Value = item.Name;
                row++;
            }

            FinalizeWorksheet(sheet, headers.Length, row - 1);
        }

        private static List<QualityCheckFactRow> BuildQualityCheckFacts(
            List<ParameterAuditRow> parameterRows,
            List<FamilyNamingRow> familyRows,
            List<SubprojectNamingRow> subprojectRows,
            List<ModelCheckerSummaryRow> modelCheckerSummary)
        {
            var rows = new List<QualityCheckFactRow>(parameterRows.Count + familyRows.Count + subprojectRows.Count + modelCheckerSummary.Count);

            foreach (ParameterAuditRow item in parameterRows)
            {
                string status = item.Status ?? string.Empty;
                string severity = NormalizeSeverity(status);

                rows.Add(new QualityCheckFactRow
                {
                    AuditType = "ParameterExistence",
                    Heading = "JeiAudit",
                    Section = "Parameter_Existence",
                    RuleName = $"{item.ParameterName}@{item.CategoryName}",
                    EntityType = "CategoryParameter",
                    EntityName = $"{item.CategoryName}:{item.ParameterName}",
                    Category = item.CategoryName,
                    Status = status,
                    Severity = severity,
                    CandidateElements = 1,
                    FailedElements = string.Equals(severity, "FAIL", StringComparison.OrdinalIgnoreCase) ? 1 : 0,
                    Reason = string.Equals(severity, "FAIL", StringComparison.OrdinalIgnoreCase) ? "Missing parameter or category binding." : "-"
                });
            }

            foreach (FamilyNamingRow item in familyRows)
            {
                string status = item.Status ?? string.Empty;
                string severity = NormalizeSeverity(status);

                rows.Add(new QualityCheckFactRow
                {
                    AuditType = "FamilyNaming",
                    Heading = "JeiAudit",
                    Section = "Family_Naming",
                    RuleName = "FamilyName_PEB",
                    EntityType = "Family",
                    EntityName = item.FamilyName,
                    Category = "Family",
                    Status = status,
                    Severity = severity,
                    CandidateElements = 1,
                    FailedElements = string.Equals(severity, "FAIL", StringComparison.OrdinalIgnoreCase) ? 1 : 0,
                    Reason = item.Reason
                });
            }

            foreach (SubprojectNamingRow item in subprojectRows)
            {
                string status = item.Status ?? string.Empty;
                string severity = NormalizeSeverity(status);

                rows.Add(new QualityCheckFactRow
                {
                    AuditType = "SubprojectNaming",
                    Heading = "JeiAudit",
                    Section = "Subproject_Naming",
                    RuleName = "SubprojectName_PEB",
                    EntityType = "Workset",
                    EntityName = item.WorksetName,
                    Category = "Workset",
                    Status = status,
                    Severity = severity,
                    CandidateElements = 1,
                    FailedElements = string.Equals(severity, "FAIL", StringComparison.OrdinalIgnoreCase) ? 1 : 0,
                    Reason = item.Reason
                });
            }

            foreach (ModelCheckerSummaryRow item in modelCheckerSummary)
            {
                string status = item.Status ?? string.Empty;
                string severity = NormalizeSeverity(status);

                rows.Add(new QualityCheckFactRow
                {
                    AuditType = "ModelChecker",
                    Heading = item.Heading,
                    Section = item.Section,
                    RuleName = item.CheckName,
                    EntityType = "Check",
                    EntityName = item.CheckName,
                    Category = "-",
                    Status = status,
                    Severity = severity,
                    CandidateElements = item.CandidateElements,
                    FailedElements = item.FailedElements,
                    Reason = item.Reason
                });
            }

            return rows;
        }

        private static void WriteKpiSummarySheet(
            XLWorkbook workbook,
            List<QualityCheckFactRow> qualityFacts,
            List<ModelCheckerFailureRow> modelCheckerFailures,
            Document doc,
            string modelCheckerXmlPath)
        {
            IXLWorksheet sheet = workbook.Worksheets.Add("10_KPI_Summary");
            string[] headers = { "Metric", "Value" };
            WriteHeaderRow(sheet, headers);

            int row = 2;
            WriteKpiRow(sheet, ref row, "GeneratedAtLocal", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture));
            WriteKpiRow(sheet, ref row, "ModelTitle", doc.Title ?? string.Empty);
            WriteKpiRow(sheet, ref row, "ModelPath", doc.PathName ?? string.Empty);
            WriteKpiRow(sheet, ref row, "ModelCheckerXmlPath", modelCheckerXmlPath ?? string.Empty);

            WriteKpiRow(sheet, ref row, "Facts_TotalRows", qualityFacts.Count.ToString(CultureInfo.InvariantCulture));
            WriteKpiRow(sheet, ref row, "Facts_Pass", CountFactsBySeverity(qualityFacts, "PASS").ToString(CultureInfo.InvariantCulture));
            WriteKpiRow(sheet, ref row, "Facts_Fail", CountFactsBySeverity(qualityFacts, "FAIL").ToString(CultureInfo.InvariantCulture));
            WriteKpiRow(sheet, ref row, "Facts_Warn", CountFactsBySeverity(qualityFacts, "WARN").ToString(CultureInfo.InvariantCulture));
            WriteKpiRow(sheet, ref row, "Facts_Skipped", CountFactsBySeverity(qualityFacts, "SKIPPED").ToString(CultureInfo.InvariantCulture));
            WriteKpiRow(sheet, ref row, "Facts_CountOnly", CountFactsBySeverity(qualityFacts, "COUNT").ToString(CultureInfo.InvariantCulture));
            WriteKpiRow(sheet, ref row, "Facts_Unsupported", CountFactsBySeverity(qualityFacts, "UNSUPPORTED").ToString(CultureInfo.InvariantCulture));

            int executed = CountFactsBySeverity(qualityFacts, "PASS") + CountFactsBySeverity(qualityFacts, "FAIL");
            double passRateExecuted = executed > 0
                ? (double)CountFactsBySeverity(qualityFacts, "PASS") * 100.0 / executed
                : 0.0;
            WriteKpiRow(sheet, ref row, "Facts_PassRateExecutedPct", passRateExecuted.ToString("0.00", CultureInfo.InvariantCulture));

            WriteKpiRow(sheet, ref row, "MC_FailureRows_Exported", modelCheckerFailures.Count.ToString(CultureInfo.InvariantCulture));
            WriteKpiRow(
                sheet,
                ref row,
                "MC_FailureDistinctElements",
                modelCheckerFailures.Select(v => v.ElementId).Distinct().Count().ToString(CultureInfo.InvariantCulture));

            FinalizeWorksheet(sheet, headers.Length, row - 1);
        }

        private static void WriteKpiRow(IXLWorksheet sheet, ref int row, string metric, string value)
        {
            sheet.Cell(row, 1).Value = metric;
            sheet.Cell(row, 2).Value = value;
            row++;
        }

        private static int CountFactsBySeverity(List<QualityCheckFactRow> rows, string severity)
        {
            return rows.Count(r => string.Equals(r.Severity, severity, StringComparison.OrdinalIgnoreCase));
        }

        private static void WriteFactQualityChecksSheet(XLWorkbook workbook, List<QualityCheckFactRow> rows)
        {
            IXLWorksheet sheet = workbook.Worksheets.Add("11_Fact_QualityChecks");
            string[] headers =
            {
                "AuditType",
                "Heading",
                "Section",
                "RuleName",
                "EntityType",
                "EntityName",
                "Category",
                "Status",
                "Severity",
                "CandidateElements",
                "FailedElements",
                "Reason"
            };
            WriteHeaderRow(sheet, headers);

            int row = 2;
            foreach (QualityCheckFactRow item in rows)
            {
                sheet.Cell(row, 1).Value = item.AuditType;
                sheet.Cell(row, 2).Value = item.Heading;
                sheet.Cell(row, 3).Value = item.Section;
                sheet.Cell(row, 4).Value = item.RuleName;
                sheet.Cell(row, 5).Value = item.EntityType;
                sheet.Cell(row, 6).Value = item.EntityName;
                sheet.Cell(row, 7).Value = item.Category;
                sheet.Cell(row, 8).Value = item.Status;
                sheet.Cell(row, 9).Value = item.Severity;
                sheet.Cell(row, 10).Value = item.CandidateElements;
                sheet.Cell(row, 11).Value = item.FailedElements;
                sheet.Cell(row, 12).Value = item.Reason;
                row++;
            }

            FinalizeWorksheet(sheet, headers.Length, row - 1);
        }

        private static void WriteFactFailedElementsSheet(XLWorkbook workbook, List<ModelCheckerFailureRow> rows, string modelCheckerXmlPath)
        {
            IXLWorksheet sheet = workbook.Worksheets.Add("12_Fact_FailedElements");
            string[] headers =
            {
                "XmlSource",
                "Heading",
                "Section",
                "CheckName",
                "ElementId",
                "Category",
                "FamilyOrType",
                "ElementName",
                "FailureKey"
            };
            WriteHeaderRow(sheet, headers);

            int row = 2;
            foreach (ModelCheckerFailureRow item in rows)
            {
                sheet.Cell(row, 1).Value = modelCheckerXmlPath ?? string.Empty;
                sheet.Cell(row, 2).Value = item.Heading;
                sheet.Cell(row, 3).Value = item.Section;
                sheet.Cell(row, 4).Value = item.CheckName;
                sheet.Cell(row, 5).Value = item.ElementId;
                sheet.Cell(row, 6).Value = item.Category;
                sheet.Cell(row, 7).Value = item.FamilyOrType;
                sheet.Cell(row, 8).Value = item.Name;
                sheet.Cell(row, 9).Value = $"{item.CheckName}|{item.ElementId}";
                row++;
            }

            FinalizeWorksheet(sheet, headers.Length, row - 1);
        }

        private static void WritePivotBySectionSheet(XLWorkbook workbook, List<ModelCheckerSummaryRow> rows)
        {
            IXLWorksheet sheet = workbook.Worksheets.Add("13_Pivot_By_Section");
            string[] headers =
            {
                "Heading",
                "Section",
                "Checks_Total",
                "Pass",
                "Fail",
                "CountOnly",
                "Unsupported",
                "Skipped",
                "CandidateElements_Sum",
                "FailedElements_Sum",
                "PassRate_ExecutedPct"
            };
            WriteHeaderRow(sheet, headers);

            int row = 2;
            IEnumerable<IGrouping<string, ModelCheckerSummaryRow>> groups = rows
                .GroupBy(v => $"{v.Heading}|{v.Section}")
                .OrderBy(g => g.Key, StringComparer.Ordinal);

            foreach (IGrouping<string, ModelCheckerSummaryRow> group in groups)
            {
                string[] tokens = group.Key.Split('|');
                string heading = tokens.Length > 0 ? tokens[0] : string.Empty;
                string section = tokens.Length > 1 ? tokens[1] : string.Empty;
                int pass = group.Count(v => string.Equals(v.Status, "PASS", StringComparison.OrdinalIgnoreCase));
                int fail = group.Count(v => string.Equals(v.Status, "FAIL", StringComparison.OrdinalIgnoreCase));
                int countOnly = group.Count(v => string.Equals(v.Status, "COUNT", StringComparison.OrdinalIgnoreCase));
                int unsupported = group.Count(v => string.Equals(v.Status, "UNSUPPORTED", StringComparison.OrdinalIgnoreCase));
                int skipped = group.Count(v => string.Equals(v.Status, "SKIPPED", StringComparison.OrdinalIgnoreCase));
                int candidateSum = group.Sum(v => v.CandidateElements);
                int failedSum = group.Sum(v => v.FailedElements);
                int executed = pass + fail;
                double passRate = executed > 0 ? (double)pass * 100.0 / executed : 0.0;

                sheet.Cell(row, 1).Value = heading;
                sheet.Cell(row, 2).Value = section;
                sheet.Cell(row, 3).Value = group.Count();
                sheet.Cell(row, 4).Value = pass;
                sheet.Cell(row, 5).Value = fail;
                sheet.Cell(row, 6).Value = countOnly;
                sheet.Cell(row, 7).Value = unsupported;
                sheet.Cell(row, 8).Value = skipped;
                sheet.Cell(row, 9).Value = candidateSum;
                sheet.Cell(row, 10).Value = failedSum;
                sheet.Cell(row, 11).Value = passRate.ToString("0.00", CultureInfo.InvariantCulture);
                row++;
            }

            FinalizeWorksheet(sheet, headers.Length, row - 1);
        }

        private static void WritePivotByCategorySheet(XLWorkbook workbook, List<ModelCheckerFailureRow> rows)
        {
            IXLWorksheet sheet = workbook.Worksheets.Add("14_Pivot_By_Category");
            string[] headers =
            {
                "Category",
                "FailureRows",
                "DistinctElements",
                "DistinctChecks"
            };
            WriteHeaderRow(sheet, headers);

            int row = 2;
            IEnumerable<IGrouping<string, ModelCheckerFailureRow>> groups = rows
                .GroupBy(v => string.IsNullOrWhiteSpace(v.Category) ? "-" : v.Category)
                .OrderByDescending(g => g.Count())
                .ThenBy(g => g.Key, StringComparer.Ordinal);

            foreach (IGrouping<string, ModelCheckerFailureRow> group in groups)
            {
                sheet.Cell(row, 1).Value = group.Key;
                sheet.Cell(row, 2).Value = group.Count();
                sheet.Cell(row, 3).Value = group.Select(v => v.ElementId).Distinct().Count();
                sheet.Cell(row, 4).Value = group.Select(v => v.CheckName).Distinct().Count();
                row++;
            }

            FinalizeWorksheet(sheet, headers.Length, row - 1);
        }

        private static void WritePivotTopFailChecksSheet(XLWorkbook workbook, List<ModelCheckerSummaryRow> rows)
        {
            IXLWorksheet sheet = workbook.Worksheets.Add("15_Pivot_Top_FailChecks");
            string[] headers =
            {
                "Rank",
                "Heading",
                "Section",
                "CheckName",
                "Status",
                "CandidateElements",
                "FailedElements",
                "FailureRatePct",
                "Reason"
            };
            WriteHeaderRow(sheet, headers);

            int row = 2;
            List<ModelCheckerSummaryRow> ordered = rows
                .Where(v => v.FailedElements > 0)
                .OrderByDescending(v => v.FailedElements)
                .ThenBy(v => v.Heading, StringComparer.Ordinal)
                .ThenBy(v => v.Section, StringComparer.Ordinal)
                .ThenBy(v => v.CheckName, StringComparer.Ordinal)
                .ToList();

            int rank = 1;
            foreach (ModelCheckerSummaryRow item in ordered)
            {
                double failRate = item.CandidateElements > 0
                    ? (double)item.FailedElements * 100.0 / item.CandidateElements
                    : 0.0;

                sheet.Cell(row, 1).Value = rank++;
                sheet.Cell(row, 2).Value = item.Heading;
                sheet.Cell(row, 3).Value = item.Section;
                sheet.Cell(row, 4).Value = item.CheckName;
                sheet.Cell(row, 5).Value = item.Status;
                sheet.Cell(row, 6).Value = item.CandidateElements;
                sheet.Cell(row, 7).Value = item.FailedElements;
                sheet.Cell(row, 8).Value = failRate.ToString("0.00", CultureInfo.InvariantCulture);
                sheet.Cell(row, 9).Value = item.Reason;
                row++;
            }

            FinalizeWorksheet(sheet, headers.Length, row - 1);
        }

        private static string NormalizeSeverity(string status)
        {
            if (string.IsNullOrWhiteSpace(status))
            {
                return "UNKNOWN";
            }

            string normalized = status.Trim().ToUpperInvariant();
            if (normalized == "OK" || normalized == "PASS")
            {
                return "PASS";
            }

            if (normalized == "MISSING" || normalized == "FAIL")
            {
                return "FAIL";
            }

            if (normalized == "WARN")
            {
                return "WARN";
            }

            if (normalized == "SKIPPED")
            {
                return "SKIPPED";
            }

            if (normalized == "COUNT")
            {
                return "COUNT";
            }

            if (normalized == "UNSUPPORTED")
            {
                return "UNSUPPORTED";
            }

            return normalized;
        }

        private static void WriteHeaderRow(IXLWorksheet sheet, string[] headers)
        {
            for (int i = 0; i < headers.Length; i++)
            {
                IXLCell cell = sheet.Cell(1, i + 1);
                cell.Value = headers[i];
                cell.Style.Font.Bold = true;
                cell.Style.Fill.BackgroundColor = XLColor.LightGray;
            }
        }

        private static void FinalizeWorksheet(IXLWorksheet sheet, int columnCount, int lastDataRow)
        {
            int lastRow = Math.Max(1, lastDataRow);
            IXLRange range = sheet.Range(1, 1, lastRow, columnCount);
            range.SetAutoFilter();
            sheet.SheetView.FreezeRows(1);

            for (int column = 1; column <= columnCount; column++)
            {
                sheet.Column(column).AdjustToContents(1, lastRow, 70);
            }
        }

        private static bool ParseTrueFalse(string value, bool defaultValue)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return defaultValue;
            }

            if (bool.TryParse(value, out bool parsed))
            {
                return parsed;
            }

            if (string.Equals(value, "1", StringComparison.Ordinal))
            {
                return true;
            }

            if (string.Equals(value, "0", StringComparison.Ordinal))
            {
                return false;
            }

            return defaultValue;
        }

        private static string CanonicalizeKnownParameterName(string raw)
        {
            string sanitized = SanitizeLoadedText(raw);
            if (string.IsNullOrWhiteSpace(sanitized))
            {
                return string.Empty;
            }

            if (!ContainsMojibakeCharacters(raw))
            {
                return sanitized;
            }

            string token = Regex.Replace(NormalizeForCompare(sanitized), "[^A-Z0-9]", string.Empty);
            if (token.Length == 0)
            {
                return sanitized;
            }

            if (token.Contains("ARGSECTOR"))
            {
                return "ARG_SECTOR";
            }

            if (token.Contains("ARGNIVEL"))
            {
                return "ARG_NIVEL";
            }

            if (token.Contains("ARGSISTEMA"))
            {
                return "ARG_SISTEMA";
            }

            if (token.Contains("ARGUNIFORMAT"))
            {
                return "ARG_UNIFORMAT";
            }

            if (token.Contains("ARGESPECIALIDAD"))
            {
                return "ARG_ESPECIALIDAD";
            }

            if (token.Contains("ARGUNIDADDEPARTIDA"))
            {
                return "ARG_UNIDAD DE PARTIDA";
            }

            if (token.Contains("ARGDESCRIP") && token.Contains("PARTIDA"))
            {
                return "ARG_DESCRIPCIÓN DE PARTIDA";
            }

            if (token.Contains("ARGC") && token.Contains("DIGO") && token.Contains("PARTIDA"))
            {
                return "ARG_CÓDIGO DE PARTIDA";
            }

            if (token.Contains("ARGC") && token.Contains("DIGO") && token.Contains("PLANO"))
            {
                return "ARG_CÓDIGO DE PLANO";
            }

            if (token.Contains("ARGCOLABORADORES"))
            {
                return "ARG_COLABORADORES";
            }

            if (token.Contains("ARGNUMEROCOLEGIATURA"))
            {
                return "ARG_NUMERO COLEGIATURA";
            }

            if (token.Contains("ARGPROYECTISTARESPONSABLE"))
            {
                return "ARG_PROYECTISTA RESPONSABLE";
            }

            if (token == "NUMERO")
            {
                return "Numero";
            }

            if (token == "NOMBRE")
            {
                return "Nombre";
            }

            return sanitized;
        }

        private static string SanitizeLoadedText(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
            {
                return string.Empty;
            }

            string current = input.Trim();
            if (!ContainsMojibakeCharacters(current))
            {
                return current;
            }

            int currentScore = GetMojibakeScore(current);
            for (int i = 0; i < 5; i++)
            {
                string candidate;
                try
                {
                    candidate = Encoding.UTF8.GetString(Latin1252Encoding.GetBytes(current));
                }
                catch
                {
                    break;
                }

                if (string.Equals(candidate, current, StringComparison.Ordinal))
                {
                    break;
                }

                int candidateScore = GetMojibakeScore(candidate);
                if (candidateScore < currentScore)
                {
                    current = candidate;
                    currentScore = candidateScore;
                    continue;
                }

                break;
            }

            return current.Trim();
        }

        private static bool ContainsMojibakeCharacters(string value)
        {
            return !string.IsNullOrEmpty(value) && GetMojibakeScore(value) > 0;
        }

        private static int GetMojibakeScore(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return 0;
            }

            int score = 0;
            foreach (char c in value)
            {
                if (c == 'Ã' || c == 'Â' || c == 'Æ' || c == 'â' || c == 'ƒ' || c == '‚' || c == '€' || c == '™' || c == '\uFFFD')
                {
                    score++;
                }
            }

            return score;
        }

        private static string GetAttribute(XElement element, string attributeName)
        {
            XAttribute attribute = element.Attributes().FirstOrDefault(a => string.Equals(a.Name.LocalName, attributeName, StringComparison.Ordinal));
            return SanitizeLoadedText(attribute?.Value ?? string.Empty);
        }

        private static bool IsNormalizedMatch(string candidate, string target)
        {
            if (string.Equals(candidate, target, StringComparison.Ordinal))
            {
                return true;
            }

            return string.Equals(
                NormalizeForCompare(candidate),
                NormalizeForCompare(target),
                StringComparison.Ordinal);
        }

        private static bool IsNoValueCondition(string normalizedCondition)
        {
            return normalizedCondition == NormalizeForCompare("HasNoValue") ||
                   normalizedCondition == NormalizeForCompare("IsEmpty") ||
                   normalizedCondition == NormalizeForCompare("Blank");
        }

        private static Parameter? FindParameterByExactName(Element element, string parameterName)
        {
            if (element == null || string.IsNullOrWhiteSpace(parameterName))
            {
                return null;
            }

            foreach (Parameter parameter in element.Parameters)
            {
                string? name = parameter?.Definition?.Name;
                if (!string.IsNullOrWhiteSpace(name) &&
                    string.Equals(name, parameterName, StringComparison.OrdinalIgnoreCase))
                {
                    return parameter;
                }
            }

            return null;
        }

        private static string NormalizeForCompare(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
            {
                return string.Empty;
            }

            string normalized = input.Trim().ToUpperInvariant().Normalize(NormalizationForm.FormD);
            var builder = new StringBuilder(normalized.Length);

            foreach (char c in normalized)
            {
                UnicodeCategory category = CharUnicodeInfo.GetUnicodeCategory(c);
                if (category != UnicodeCategory.NonSpacingMark)
                {
                    builder.Append(c);
                }
            }

            return builder.ToString();
        }

        private static string EscapeCsv(string value)
        {
            if (value == null)
            {
                return "\"\"";
            }

            string escaped = value.Replace("\"", "\"\"");
            return $"\"{escaped}\"";
        }
    }
}

