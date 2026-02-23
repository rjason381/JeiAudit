using System;
using System.Diagnostics;
using System.IO;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace JeiAudit
{
    [Transaction(TransactionMode.ReadOnly)]
    public class ViewLastReportCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                string folder = PluginState.LoadLastOutputFolder();
                if (string.IsNullOrWhiteSpace(folder) || !Directory.Exists(folder))
                {
                    folder = PluginState.FindLatestOutputFolder();
                }

                if (string.IsNullOrWhiteSpace(folder) || !Directory.Exists(folder))
                {
                    TaskDialog.Show(
                        "JeiAudit - View Report",
                        "No previous report folder was found.\nRun the audit first.");
                    return Result.Cancelled;
                }

                try
                {
                    AuditReportData data = AuditReportLoader.LoadFromOutputFolder(folder);
                    using (var form = new ReportViewerForm(data))
                    {
                        form.ShowDialog();
                    }
                }
                catch (FileNotFoundException)
                {
                    TaskDialog.Show(
                        "JeiAudit - View Report",
                        "No se encontro JeiAudit_Results.xlsx en esa carpeta.\nSe abrira la carpeta para revision manual.");
                    Process.Start(new ProcessStartInfo(folder) { UseShellExecute = true });
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                TaskDialog.Show("JeiAudit", $"Unable to open report folder.{Environment.NewLine}{ex.Message}");
                return Result.Failed;
            }
        }
    }
}
