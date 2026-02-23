using System;
using System.IO;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using WinForms = System.Windows.Forms;

namespace JeiAudit
{
    [Transaction(TransactionMode.ReadOnly)]
    public class ChecksetEditorCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIDocument uiDoc = commandData.Application.ActiveUIDocument;
                Document? doc = uiDoc?.Document;

                string lastPath = PluginState.LoadLastXmlPath();
                string documentDir = doc == null || string.IsNullOrWhiteSpace(doc.PathName)
                    ? string.Empty
                    : (Path.GetDirectoryName(doc.PathName) ?? string.Empty);
                string defaultPath = PluginState.ResolveDefaultXmlPath(lastPath, documentDir);

                using (var form = new ChecksetEditorForm(defaultPath))
                {
                    WinForms.DialogResult result = form.ShowDialog();
                    if (result != WinForms.DialogResult.OK)
                    {
                        return Result.Cancelled;
                    }

                    if (!string.IsNullOrWhiteSpace(form.SelectedXmlPath) && File.Exists(form.SelectedXmlPath))
                    {
                        PluginState.SaveLastXmlPath(form.SelectedXmlPath);
                    }
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                TaskDialog.Show("JeiAudit", $"Editor failed.{Environment.NewLine}{ex.Message}");
                return Result.Failed;
            }
        }
    }
}
