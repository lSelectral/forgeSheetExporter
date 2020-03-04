using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using DesignAutomationFramework;

namespace sheetExporterRevit
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class Command : IExternalDBApplication
    {
        string OUTPUT_FILE = "OutputFile.rvt";

        public ExternalDBApplicationResult OnShutdown(ControlledApplication application)
        {
            return ExternalDBApplicationResult.Succeeded;
        }

        public ExternalDBApplicationResult OnStartup(ControlledApplication application)
        {
            DesignAutomationBridge.DesignAutomationReadyEvent += DesignAutomationBridge_DesignAutomationReadyEvent;
            return ExternalDBApplicationResult.Succeeded;
        }

        private void DesignAutomationBridge_DesignAutomationReadyEvent(object sender, DesignAutomationReadyEventArgs e)
        {
            LogTrace("Design Automation Ready event triggered...");
            e.Succeeded = true;
            DesignAutomationData data = e.DesignAutomationData;
            Application rvtApp = data.RevitApp;
            string modelPath = data.FilePath;
            Document doc = data.RevitDoc;

            using (Transaction t = new Transaction(doc))
            {
                // Start the transaction
                t.Start("Export Sheets DWG");
                try
                {
                    var viewSheets = new FilteredElementCollector(doc).OfClass(typeof(ViewSheet)).Cast<ViewSheet>().Where(vw =>
                           vw.ViewType == ViewType.DrawingSheet && !vw.IsTemplate);

                    // create DWG export options
                    DWGExportOptions dwgOptions = new DWGExportOptions();
                    dwgOptions.MergedViews = true;
                    dwgOptions.SharedCoords = true;
                    dwgOptions.FileVersion = ACADVersion.R2018;

                    List<ElementId> views = new List<ElementId>();

                    foreach (var sheet in viewSheets)
                    {
                        if (!sheet.IsPlaceholder)
                        {
                            views.Add(sheet.Id);
                        }
                    }

                    // For Web Deployment
                    //doc.Export(@"D:\sheetExporterLocation", "TEST", views, dwgOptions);
                    // For Local
                    doc.Export(Directory.GetCurrentDirectory() + "//exportedDwgs", "rvt", views, dwgOptions);

                }
                catch (Exception ex)
                {
                }
                // Commit the transaction
                t.Commit();
            }
        }

        /// <summary>
        /// This will appear on the Design Automation output
        /// </summary>
        private static void LogTrace(string format, params object[] args) { System.Console.WriteLine(format, args); }
    }
}