using System;
using System.Linq;
using System.Reflection;
using Autodesk.Revit.UI;

namespace JeiAudit
{
    public class App : IExternalApplication
    {
        private const string TabName = "JeiAudit";
        private const string PanelName = "Herramienta de auditor\u00EDa JeiAudit";

        public Result OnStartup(UIControlledApplication application)
        {
            try
            {
                try
                {
                    application.CreateRibbonTab(TabName);
                }
                catch
                {
                    // Ribbon tab already exists.
                }

                RibbonPanel panel = application
                    .GetRibbonPanels(TabName)
                    .FirstOrDefault(p => string.Equals(p.Name, PanelName, StringComparison.Ordinal))
                    ?? application.CreateRibbonPanel(TabName, PanelName);

                string assemblyPath = Assembly.GetExecutingAssembly().Location;

                var setupButtonData = new PushButtonData(
                    "JeiAudit.SetupModelChecker",
                    "Configuracion",
                    assemblyPath,
                    "JeiAudit.SetupModelCheckerCommand")
                {
                    ToolTip = "Abrir y configurar el XML de chequeos (MCSettings*.xml). Desarrollado por Jason Rojas Estrada - Coordinador BIM, Inspirada en herramientas de Autodesk."
                };

                var runButtonData = new PushButtonData(
                    "JeiAudit.RunAudit",
                    "Ejecutar",
                    assemblyPath,
                    "JeiAudit.RunAuditCommand")
                {
                    ToolTip = "Ejecutar auditoria con el XML configurado y las reglas extendidas de JeiAudit."
                };

                var viewReportButtonData = new PushButtonData(
                    "JeiAudit.ViewLastReport",
                    "Ver\nReporte",
                    assemblyPath,
                    "JeiAudit.ViewLastReportCommand")
                {
                    ToolTip = "Abrir visor de reporte con resumen, arbol y exportaciones (Html/Excel/AVT)."
                };

                var checksetEditorButtonData = new PushButtonData(
                    "JeiAudit.ChecksetEditor",
                    "Editor\nCheckset",
                    assemblyPath,
                    "JeiAudit.ChecksetEditorCommand")
                {
                    ToolTip = "Crear, cargar, copiar y guardar checksets XML de auditor\u00EDa JeiAudit."
                };

                panel.AddItem(setupButtonData);
                panel.AddItem(runButtonData);
                panel.AddItem(viewReportButtonData);
                panel.AddItem(checksetEditorButtonData);
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("JeiAudit", $"Failed to initialize plugin.{Environment.NewLine}{ex.Message}");
                return Result.Failed;
            }
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }
    }
}
