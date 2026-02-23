using System;
using System.Collections.Generic;
using System.Globalization;

namespace JeiAudit
{
    internal sealed class AuditReportData
    {
        public string OutputFolder { get; set; } = string.Empty;
        public string ExcelPath { get; set; } = string.Empty;
        public string ModelTitle { get; set; } = string.Empty;
        public string ModelPath { get; set; } = string.Empty;
        public string ModelCheckerXmlPath { get; set; } = string.Empty;
        public string ChecksetTitle { get; set; } = "-";
        public string ChecksetDate { get; set; } = "-";
        public string ChecksetAuthor { get; set; } = "-";
        public string ChecksetDescription { get; set; } = "-";
        public DateTime? GeneratedAtLocal { get; set; }
        public List<ModelCheckerSummaryItem> ModelCheckerSummaryRows { get; } = new List<ModelCheckerSummaryItem>();
        public List<ModelCheckerFailedElementItem> ModelCheckerFailedElements { get; } = new List<ModelCheckerFailedElementItem>();
        public List<ParameterExistenceReportItem> ParameterExistenceRows { get; } = new List<ParameterExistenceReportItem>();
        public List<FamilyNamingReportItem> FamilyNamingRows { get; } = new List<FamilyNamingReportItem>();
        public List<SubprojectNamingReportItem> SubprojectNamingRows { get; } = new List<SubprojectNamingReportItem>();

        public int PassCount => CountByStatus("PASS");
        public int FailCount => CountByStatus("FAIL");
        public int CountOnlyCount => CountByStatus("COUNT");
        public int UnsupportedCount => CountByStatus("UNSUPPORTED");
        public int SkippedCount => CountByStatus("SKIPPED");
        public int TotalChecks => ModelCheckerSummaryRows.Count;
        public int ExecutedChecks => PassCount + FailCount;

        public double PassRatePercent
        {
            get
            {
                if (ExecutedChecks <= 0)
                {
                    return 0.0;
                }

                return (double)PassCount * 100.0 / ExecutedChecks;
            }
        }

        public string PassRateText => PassRatePercent.ToString("0", CultureInfo.InvariantCulture) + "%";

        private int CountByStatus(string status)
        {
            int count = 0;
            foreach (ModelCheckerSummaryItem item in ModelCheckerSummaryRows)
            {
                if (string.Equals(item.Status, status, StringComparison.OrdinalIgnoreCase))
                {
                    count++;
                }
            }

            return count;
        }
    }

    internal sealed class ModelCheckerSummaryItem
    {
        public string Heading { get; set; } = string.Empty;
        public string HeadingDescription { get; set; } = string.Empty;
        public string Section { get; set; } = string.Empty;
        public string SectionDescription { get; set; } = string.Empty;
        public string CheckName { get; set; } = string.Empty;
        public string CheckDescription { get; set; } = string.Empty;
        public string Status { get; set; } = string.Empty;
        public int CandidateElements { get; set; }
        public int FailedElements { get; set; }
        public string Reason { get; set; } = string.Empty;
    }

    internal sealed class ModelCheckerFailedElementItem
    {
        public string Heading { get; set; } = string.Empty;
        public string Section { get; set; } = string.Empty;
        public string CheckName { get; set; } = string.Empty;
        public long ElementId { get; set; }
        public string Category { get; set; } = string.Empty;
        public string FamilyOrType { get; set; } = string.Empty;
        public string ElementName { get; set; } = string.Empty;
    }

    internal sealed class FamilyNamingReportItem
    {
        public string FamilyName { get; set; } = string.Empty;
        public string HasArgPrefix { get; set; } = string.Empty;
        public string DisciplineCode { get; set; } = string.Empty;
        public string DisciplineCodeValid { get; set; } = string.Empty;
        public string HasNamePart { get; set; } = string.Empty;
        public string Status { get; set; } = string.Empty;
        public string Reason { get; set; } = string.Empty;
    }

    internal sealed class ParameterExistenceReportItem
    {
        public string Parameter { get; set; } = string.Empty;
        public string Category { get; set; } = string.Empty;
        public string CategoryAvailable { get; set; } = string.Empty;
        public string ExistsByExactName { get; set; } = string.Empty;
        public string BoundToCategory { get; set; } = string.Empty;
        public string BindingTypes { get; set; } = string.Empty;
        public string ParameterGuids { get; set; } = string.Empty;
        public string SimilarNames { get; set; } = string.Empty;
        public string Status { get; set; } = string.Empty;
    }

    internal sealed class SubprojectNamingReportItem
    {
        public string WorksetName { get; set; } = string.Empty;
        public string HasArgPrefix { get; set; } = string.Empty;
        public string HasNameAfterPrefix { get; set; } = string.Empty;
        public string HasSecondSeparator { get; set; } = string.Empty;
        public string Status { get; set; } = string.Empty;
        public string Reason { get; set; } = string.Empty;
    }
}
