using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using ClosedXML.Excel;

namespace JeiAudit
{
    internal static class AuditReportLoader
    {
        private const string WorkbookName = "JeiAudit_Results.xlsx";
        private static readonly Encoding Latin1252Encoding = Encoding.GetEncoding(1252);

        internal static AuditReportData LoadFromOutputFolder(string outputFolder)
        {
            if (string.IsNullOrWhiteSpace(outputFolder) || !Directory.Exists(outputFolder))
            {
                throw new DirectoryNotFoundException("Output folder was not found.");
            }

            string excelPath = Path.Combine(outputFolder, WorkbookName);
            if (!File.Exists(excelPath))
            {
                throw new FileNotFoundException("Excel report file was not found.", excelPath);
            }

            var data = new AuditReportData
            {
                OutputFolder = outputFolder,
                ExcelPath = excelPath
            };

            using (var workbook = new XLWorkbook(excelPath))
            {
                ReadMetaSheet(workbook, data);
                ReadParameterSheet(workbook, data);
                ReadFamilyNamingSheet(workbook, data);
                ReadSubprojectNamingSheet(workbook, data);
                ReadSummarySheet(workbook, data);
                ReadFailedElementsSheet(workbook, data);
            }

            ReadChecksetMetadataAndDescriptions(data);
            return data;
        }

        private static void ReadMetaSheet(XLWorkbook workbook, AuditReportData data)
        {
            if (!workbook.TryGetWorksheet("00_Meta", out IXLWorksheet sheet))
            {
                return;
            }

            int lastRow = sheet.LastRowUsed()?.RowNumber() ?? 1;
            for (int row = 2; row <= lastRow; row++)
            {
                string field = ReadCellText(sheet.Cell(row, 1));
                string value = ReadCellText(sheet.Cell(row, 2));

                if (field.Length == 0)
                {
                    continue;
                }

                switch (field)
                {
                    case "GeneratedAtLocal":
                        if (DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out DateTime parsed))
                        {
                            data.GeneratedAtLocal = parsed;
                        }
                        break;
                    case "ModelTitle":
                        data.ModelTitle = value;
                        break;
                    case "ModelPath":
                        data.ModelPath = value;
                        break;
                    case "ModelCheckerXmlPath":
                        data.ModelCheckerXmlPath = value;
                        break;
                }
            }
        }

        private static void ReadSummarySheet(XLWorkbook workbook, AuditReportData data)
        {
            if (!workbook.TryGetWorksheet("04_MC_Summary", out IXLWorksheet sheet))
            {
                return;
            }

            int lastRow = sheet.LastRowUsed()?.RowNumber() ?? 1;
            if (lastRow < 2)
            {
                return;
            }

            Dictionary<string, int> columns = BuildColumnMap(sheet);
            if (!columns.TryGetValue("Heading", out int headingCol) ||
                !columns.TryGetValue("Section", out int sectionCol) ||
                !columns.TryGetValue("CheckName", out int checkNameCol) ||
                !columns.TryGetValue("Status", out int statusCol))
            {
                return;
            }

            int candidateCol = columns.TryGetValue("CandidateElements", out int c) ? c : -1;
            int failedCol = columns.TryGetValue("FailedElements", out int f) ? f : -1;
            int reasonCol = columns.TryGetValue("Reason", out int r) ? r : -1;

            for (int row = 2; row <= lastRow; row++)
            {
                string checkName = ReadCellText(sheet.Cell(row, checkNameCol));
                string status = ReadCellText(sheet.Cell(row, statusCol));

                if (checkName.Length == 0 && status.Length == 0)
                {
                    continue;
                }

                data.ModelCheckerSummaryRows.Add(new ModelCheckerSummaryItem
                {
                    Heading = ReadCellText(sheet.Cell(row, headingCol)),
                    Section = ReadCellText(sheet.Cell(row, sectionCol)),
                    CheckName = checkName,
                    Status = status,
                    CandidateElements = candidateCol > 0 ? ReadIntCell(sheet.Cell(row, candidateCol)) : 0,
                    FailedElements = failedCol > 0 ? ReadIntCell(sheet.Cell(row, failedCol)) : 0,
                    Reason = reasonCol > 0 ? ReadCellText(sheet.Cell(row, reasonCol)) : string.Empty
                });
            }
        }

        private static void ReadParameterSheet(XLWorkbook workbook, AuditReportData data)
        {
            if (!workbook.TryGetWorksheet("01_Parameter_Existence", out IXLWorksheet sheet))
            {
                return;
            }

            int lastRow = sheet.LastRowUsed()?.RowNumber() ?? 1;
            if (lastRow < 2)
            {
                return;
            }

            Dictionary<string, int> columns = BuildColumnMap(sheet);
            if (!columns.TryGetValue("Parameter", out int parameterCol))
            {
                return;
            }

            int categoryCol = columns.TryGetValue("Category", out int category) ? category : -1;
            int categoryAvailableCol = columns.TryGetValue("CategoryAvailable", out int categoryAvailable) ? categoryAvailable : -1;
            int existsByExactNameCol = columns.TryGetValue("ExistsByExactName", out int existsByExactName) ? existsByExactName : -1;
            int boundToCategoryCol = columns.TryGetValue("BoundToCategory", out int boundToCategory) ? boundToCategory : -1;
            int bindingTypesCol = columns.TryGetValue("BindingTypes", out int bindingTypes) ? bindingTypes : -1;
            int guidsCol = columns.TryGetValue("ParameterGuids", out int guids) ? guids : -1;
            int similarNamesCol = columns.TryGetValue("SimilarNames", out int similarNames) ? similarNames : -1;
            int statusCol = columns.TryGetValue("Status", out int status) ? status : -1;

            for (int row = 2; row <= lastRow; row++)
            {
                string parameter = ReadCellText(sheet.Cell(row, parameterCol));
                if (parameter.Length == 0)
                {
                    continue;
                }

                data.ParameterExistenceRows.Add(new ParameterExistenceReportItem
                {
                    Parameter = parameter,
                    Category = categoryCol > 0 ? ReadCellText(sheet.Cell(row, categoryCol)) : string.Empty,
                    CategoryAvailable = categoryAvailableCol > 0 ? ReadCellText(sheet.Cell(row, categoryAvailableCol)) : string.Empty,
                    ExistsByExactName = existsByExactNameCol > 0 ? ReadCellText(sheet.Cell(row, existsByExactNameCol)) : string.Empty,
                    BoundToCategory = boundToCategoryCol > 0 ? ReadCellText(sheet.Cell(row, boundToCategoryCol)) : string.Empty,
                    BindingTypes = bindingTypesCol > 0 ? ReadCellText(sheet.Cell(row, bindingTypesCol)) : string.Empty,
                    ParameterGuids = guidsCol > 0 ? ReadCellText(sheet.Cell(row, guidsCol)) : string.Empty,
                    SimilarNames = similarNamesCol > 0 ? ReadCellText(sheet.Cell(row, similarNamesCol)) : string.Empty,
                    Status = statusCol > 0 ? ReadCellText(sheet.Cell(row, statusCol)) : string.Empty
                });
            }
        }

        private static void ReadFamilyNamingSheet(XLWorkbook workbook, AuditReportData data)
        {
            if (!workbook.TryGetWorksheet("02_Family_Naming", out IXLWorksheet sheet))
            {
                return;
            }

            int lastRow = sheet.LastRowUsed()?.RowNumber() ?? 1;
            if (lastRow < 2)
            {
                return;
            }

            Dictionary<string, int> columns = BuildColumnMap(sheet);
            if (!columns.TryGetValue("FamilyName", out int familyNameCol))
            {
                return;
            }

            int hasArgPrefixCol = columns.TryGetValue("HasARGPrefix", out int argPrefix) ? argPrefix : -1;
            int disciplineCodeCol = columns.TryGetValue("DisciplineCode", out int disciplineCode) ? disciplineCode : -1;
            int disciplineCodeValidCol = columns.TryGetValue("DisciplineCodeValid", out int disciplineValid) ? disciplineValid : -1;
            int hasNamePartCol = columns.TryGetValue("HasNamePart", out int namePart) ? namePart : -1;
            int statusCol = columns.TryGetValue("Status", out int status) ? status : -1;
            int reasonCol = columns.TryGetValue("Reason", out int reason) ? reason : -1;

            for (int row = 2; row <= lastRow; row++)
            {
                string familyName = ReadCellText(sheet.Cell(row, familyNameCol));
                if (familyName.Length == 0)
                {
                    continue;
                }

                data.FamilyNamingRows.Add(new FamilyNamingReportItem
                {
                    FamilyName = familyName,
                    HasArgPrefix = hasArgPrefixCol > 0 ? ReadCellText(sheet.Cell(row, hasArgPrefixCol)) : string.Empty,
                    DisciplineCode = disciplineCodeCol > 0 ? ReadCellText(sheet.Cell(row, disciplineCodeCol)) : string.Empty,
                    DisciplineCodeValid = disciplineCodeValidCol > 0 ? ReadCellText(sheet.Cell(row, disciplineCodeValidCol)) : string.Empty,
                    HasNamePart = hasNamePartCol > 0 ? ReadCellText(sheet.Cell(row, hasNamePartCol)) : string.Empty,
                    Status = statusCol > 0 ? ReadCellText(sheet.Cell(row, statusCol)) : string.Empty,
                    Reason = reasonCol > 0 ? ReadCellText(sheet.Cell(row, reasonCol)) : string.Empty
                });
            }
        }

        private static void ReadSubprojectNamingSheet(XLWorkbook workbook, AuditReportData data)
        {
            if (!workbook.TryGetWorksheet("03_Subproject_Naming", out IXLWorksheet sheet))
            {
                return;
            }

            int lastRow = sheet.LastRowUsed()?.RowNumber() ?? 1;
            if (lastRow < 2)
            {
                return;
            }

            Dictionary<string, int> columns = BuildColumnMap(sheet);
            if (!columns.TryGetValue("WorksetName", out int worksetNameCol))
            {
                return;
            }

            int hasArgPrefixCol = columns.TryGetValue("HasArgPrefix", out int argPrefix) ? argPrefix : -1;
            int hasNameAfterPrefixCol = columns.TryGetValue("HasNameAfterPrefix", out int nameAfterPrefix) ? nameAfterPrefix : -1;
            int hasSecondSeparatorCol = columns.TryGetValue("HasSecondSeparator", out int secondSeparator) ? secondSeparator : -1;
            int statusCol = columns.TryGetValue("Status", out int status) ? status : -1;
            int reasonCol = columns.TryGetValue("Reason", out int reason) ? reason : -1;

            for (int row = 2; row <= lastRow; row++)
            {
                string worksetName = ReadCellText(sheet.Cell(row, worksetNameCol));
                if (worksetName.Length == 0)
                {
                    continue;
                }

                data.SubprojectNamingRows.Add(new SubprojectNamingReportItem
                {
                    WorksetName = worksetName,
                    HasArgPrefix = hasArgPrefixCol > 0 ? ReadCellText(sheet.Cell(row, hasArgPrefixCol)) : string.Empty,
                    HasNameAfterPrefix = hasNameAfterPrefixCol > 0 ? ReadCellText(sheet.Cell(row, hasNameAfterPrefixCol)) : string.Empty,
                    HasSecondSeparator = hasSecondSeparatorCol > 0 ? ReadCellText(sheet.Cell(row, hasSecondSeparatorCol)) : string.Empty,
                    Status = statusCol > 0 ? ReadCellText(sheet.Cell(row, statusCol)) : string.Empty,
                    Reason = reasonCol > 0 ? ReadCellText(sheet.Cell(row, reasonCol)) : string.Empty
                });
            }
        }

        private static void ReadFailedElementsSheet(XLWorkbook workbook, AuditReportData data)
        {
            if (!workbook.TryGetWorksheet("05_MC_Failed_Elements", out IXLWorksheet sheet))
            {
                return;
            }

            int lastRow = sheet.LastRowUsed()?.RowNumber() ?? 1;
            if (lastRow < 2)
            {
                return;
            }

            Dictionary<string, int> columns = BuildColumnMap(sheet);
            if (!columns.TryGetValue("Heading", out int headingCol) ||
                !columns.TryGetValue("Section", out int sectionCol) ||
                !columns.TryGetValue("CheckName", out int checkNameCol))
            {
                return;
            }

            int elementIdCol = columns.TryGetValue("ElementId", out int idCol) ? idCol : -1;
            int categoryCol = columns.TryGetValue("Category", out int category) ? category : -1;
            int typeCol = columns.TryGetValue("FamilyOrType", out int type) ? type : -1;
            int nameCol = columns.TryGetValue("ElementName", out int elementName) ? elementName : -1;

            for (int row = 2; row <= lastRow; row++)
            {
                string heading = ReadCellText(sheet.Cell(row, headingCol));
                string section = ReadCellText(sheet.Cell(row, sectionCol));
                string checkName = ReadCellText(sheet.Cell(row, checkNameCol));
                if (heading.Length == 0 && section.Length == 0 && checkName.Length == 0)
                {
                    continue;
                }

                data.ModelCheckerFailedElements.Add(new ModelCheckerFailedElementItem
                {
                    Heading = heading,
                    Section = section,
                    CheckName = checkName,
                    ElementId = elementIdCol > 0 ? ReadLongCell(sheet.Cell(row, elementIdCol)) : 0,
                    Category = categoryCol > 0 ? ReadCellText(sheet.Cell(row, categoryCol)) : string.Empty,
                    FamilyOrType = typeCol > 0 ? ReadCellText(sheet.Cell(row, typeCol)) : string.Empty,
                    ElementName = nameCol > 0 ? ReadCellText(sheet.Cell(row, nameCol)) : string.Empty
                });
            }
        }

        private static void ReadChecksetMetadataAndDescriptions(AuditReportData data)
        {
            if (string.IsNullOrWhiteSpace(data.ModelCheckerXmlPath))
            {
                return;
            }

            if (!File.Exists(data.ModelCheckerXmlPath))
            {
                return;
            }

            XDocument xml;
            try
            {
                xml = XDocument.Load(data.ModelCheckerXmlPath);
            }
            catch
            {
                return;
            }

            XElement? root = xml.Root;
            if (root == null)
            {
                return;
            }

            data.ChecksetTitle = ReadAttribute(root, "Name", Path.GetFileNameWithoutExtension(data.ModelCheckerXmlPath));
            data.ChecksetDate = ReadAttribute(root, "Date", data.GeneratedAtLocal?.ToString("dddd, dd 'de' MMMM 'de' yyyy", CultureInfo.GetCultureInfo("es-ES")) ?? "-");
            data.ChecksetAuthor = ReadAttribute(root, "Author", "-");
            data.ChecksetDescription = ReadAttribute(root, "Description", "-");

            var descriptionMap = new Dictionary<string, (string HeadingDescription, string SectionDescription, string CheckDescription)>(StringComparer.OrdinalIgnoreCase);
            foreach (XElement heading in root.Elements().Where(e => string.Equals(e.Name.LocalName, "Heading", StringComparison.OrdinalIgnoreCase)))
            {
                string headingName = ReadAttribute(heading, "HeadingText", string.Empty);
                string headingDescription = ReadAttribute(heading, "Description", string.Empty);

                foreach (XElement section in heading.Elements().Where(e => string.Equals(e.Name.LocalName, "Section", StringComparison.OrdinalIgnoreCase)))
                {
                    string sectionName = ReadAttribute(section, "SectionName", string.Empty);
                    string sectionDescription = ReadAttribute(section, "Description", string.Empty);

                    foreach (XElement check in section.Elements().Where(e => string.Equals(e.Name.LocalName, "Check", StringComparison.OrdinalIgnoreCase)))
                    {
                        string checkName = ReadAttribute(check, "CheckName", string.Empty);
                        if (checkName.Length == 0)
                        {
                            continue;
                        }

                        string key = BuildCheckKey(headingName, sectionName, checkName);
                        if (!descriptionMap.ContainsKey(key))
                        {
                            descriptionMap[key] = (headingDescription, sectionDescription, ReadAttribute(check, "Description", string.Empty));
                        }
                    }
                }
            }

            foreach (ModelCheckerSummaryItem item in data.ModelCheckerSummaryRows)
            {
                string key = BuildCheckKey(item.Heading, item.Section, item.CheckName);
                if (descriptionMap.TryGetValue(key, out var match))
                {
                    item.HeadingDescription = match.HeadingDescription;
                    item.SectionDescription = match.SectionDescription;
                    item.CheckDescription = match.CheckDescription;
                }
            }
        }

        private static string ReadCellText(IXLCell cell)
        {
            return CleanText(cell.GetString()).Trim();
        }

        private static string CleanText(string input)
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

            return current;
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

        private static Dictionary<string, int> BuildColumnMap(IXLWorksheet sheet)
        {
            var map = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            IXLRow headerRow = sheet.Row(1);
            int lastColumn = sheet.LastColumnUsed()?.ColumnNumber() ?? 1;

            for (int col = 1; col <= lastColumn; col++)
            {
                string name = ReadCellText(headerRow.Cell(col));
                if (name.Length == 0 || map.ContainsKey(name))
                {
                    continue;
                }

                map[name] = col;
            }

            return map;
        }

        private static int ReadIntCell(IXLCell cell)
        {
            if (cell.TryGetValue(out int intValue))
            {
                return intValue;
            }

            if (cell.TryGetValue(out double numberValue))
            {
                return Convert.ToInt32(Math.Round(numberValue, 0), CultureInfo.InvariantCulture);
            }

            string text = ReadCellText(cell);
            return int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed) ? parsed : 0;
        }

        private static long ReadLongCell(IXLCell cell)
        {
            if (cell.TryGetValue(out long longValue))
            {
                return longValue;
            }

            if (cell.TryGetValue(out int intValue))
            {
                return intValue;
            }

            if (cell.TryGetValue(out double numberValue))
            {
                return Convert.ToInt64(Math.Round(numberValue, 0), CultureInfo.InvariantCulture);
            }

            string text = ReadCellText(cell);
            return long.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out long parsed) ? parsed : 0;
        }

        private static string BuildCheckKey(string heading, string section, string check)
        {
            return (heading ?? string.Empty).Trim() + "\u001f" +
                   (section ?? string.Empty).Trim() + "\u001f" +
                   (check ?? string.Empty).Trim();
        }

        private static string ReadAttribute(XElement element, string name, string fallback)
        {
            XAttribute? attribute = element.Attributes().FirstOrDefault(a => string.Equals(a.Name.LocalName, name, StringComparison.OrdinalIgnoreCase));
            if (attribute == null || string.IsNullOrWhiteSpace(attribute.Value))
            {
                return fallback;
            }

            return CleanText(attribute.Value).Trim();
        }
    }
}

