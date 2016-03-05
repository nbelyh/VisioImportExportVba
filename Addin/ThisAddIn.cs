using System;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using VisioImportExportVba.Properties;
using Visio = Microsoft.Office.Interop.Visio;

namespace VisioImportExportVba
{
    public partial class ThisAddIn
    {
        private readonly AddinUI AddinUI = new AddinUI();

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return AddinUI;
        }

        public void ExportVBA(string folder, Settings settings)
        {
            var window = Application.ActiveWindow;
            if (window == null)
                return;
            
            switch ((Visio.VisWinTypes)window.Type)
            {
                case Visio.VisWinTypes.visDrawing:

                    ExportVBA(window.Document, folder);

                    if (settings.IncludeStencils)
                    {
                        Array stencilNames;
                        window.DockedStencils(out stencilNames);
                        foreach (var stencilName in stencilNames)
                        {
                            var stencilDoc = Application.Documents[stencilName];
                            ExportVBA(stencilDoc, Path.Combine(folder, stencilDoc.Name));
                        }
                    }
                    break;

                case Visio.VisWinTypes.visStencil:
                    ExportVBA(window.Document, folder);
                    break;
            }
        }

        string GetComponentFileExtension(dynamic component)
        {
            int componentType = Convert.ToInt32(component.Type);
            switch (componentType)
            {
                case 1: return ".bas";
                case 2: return ".cls";
                case 3: return ".frm";
                default:
                    return null;
            }
        }

        public void ExportVBA(Visio.Document doc, string path)
        {
            if (doc == null)
                return;

            var project = doc.VBProject;

            if (project == null)
                return;

            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);

            ExportThisDocumentVBA(path, project);

            foreach (var component in project.VBComponents)
            {
                var fileExtension = GetComponentFileExtension(component);
                if (fileExtension == null)
                    continue;

                component.Export(Path.Combine(path, component.Name + fileExtension));
            }
        }

        private static void ExportThisDocumentVBA(string path, dynamic project)
        {
            var thisDocumentComponent = project.VBComponents["ThisDocument"];
            if (thisDocumentComponent != null)
            {
                var codeModule = thisDocumentComponent.CodeModule;
                var countOfLines = Convert.ToInt32(codeModule.CountOfLines);
                if (countOfLines > 0)
                {
                    var lines = codeModule.Lines(1, countOfLines);
                    File.WriteAllText(Path.Combine(path, "ThisDocument.bas"), lines);
                }
            }
        }

        /// <summary>
        /// A command to demonstrate conditionally enabling/disabling.
        /// The command gets enabled only when a shape is selected
        /// </summary>

        public void ImportVBA(string folder, Settings settings)
        {
            var window = Application.ActiveWindow;
            switch ((Visio.VisWinTypes)window.Type)
            {
                case Visio.VisWinTypes.visDrawing:
                    {
                        ImportVBA(window.Document, folder, settings);

                        if (settings.IncludeStencils)
                        {
                            Array stencilNames;
                            window.DockedStencils(out stencilNames);
                            foreach (string stencilName in stencilNames)
                            {
                                var stencilDoc = Application.Documents[stencilName];

                                var readOnly = stencilDoc.ReadOnly != 0;
                                if (readOnly)
                                {
                                    stencilDoc.Close();
                                    short flags = (short) Visio.VisOpenSaveArgs.visOpenDocked
                                                | (short) Visio.VisOpenSaveArgs.visOpenRW;

                                    stencilDoc = Application.Documents.OpenEx(stencilName, flags);
                                }

                                ImportVBA(stencilDoc, Path.Combine(folder, stencilDoc.Name), settings);
                            }
                        }
                    }
                    break;

                case Visio.VisWinTypes.visStencil:
                    ImportVBA(window.Document, folder, settings);
                    break;
            }
        }

        void ImportThisDocumentVBA(string filePath, dynamic project)
        {
            var thisDocumentComponent = project.VBComponents["ThisDocument"];
            if (thisDocumentComponent != null)
            {
                var codeModule = thisDocumentComponent.CodeModule;
                var countOfLines = Convert.ToInt32(codeModule.CountOfLines);
                if (countOfLines > 0)
                    codeModule.DeleteLines(1, countOfLines);

                codeModule.AddFromString(File.ReadAllText(filePath));
            }
        }

        void ImportVBA(Visio.Document doc, string path, Settings settings)
        {
            var project = doc.VBProject;

            if (project == null)
                return;

            var files = Directory.GetFiles(path);

            foreach (var component in project.VBComponents)
            {
                var fileExtension = GetComponentFileExtension(component);
                if (fileExtension == null)
                    continue;

                if (settings.ClearBeforeImport || files.Any(f =>
                    string.Compare(Path.GetFileNameWithoutExtension(f), component.Name, StringComparison.OrdinalIgnoreCase) == 0))
                {
                    project.VBComponents.Remove(component);
                }
            }

            foreach (var file in files)
            {
                if (Path.GetFileName(file) == "ThisDocument.bas")
                {
                    ImportThisDocumentVBA(file, project);
                    continue;
                }

                var extension = Path.GetExtension(file);
                if (extension == null)
                    continue;

                switch (extension.ToLower())
                {
                    case ".cls":
                    case ".frm":
                    case ".bas":
                        project.VBComponents.Import(file);
                        break;
                }
            }
        }

        /// <summary>
        /// Callback called by the UI manager when user clicks a button
        /// Should do something meaningful when corresponding action is called.
        /// </summary>
        public void OnCommand(string commandId)
        {
            try
            {
                var doc = Application.ActiveDocument;

                var settings = SettingsManager.LoadOrCreate(doc);

                switch (commandId)
                {
                    case "ExportVBA":
                        if (string.IsNullOrEmpty(settings.TargetFolder))
                        {
                            OnCommand("ExportVBAFolder");
                            return;
                        }

                        ExportVBA(settings.TargetFolder, settings);
                        MessageBox.Show(
                            string.Format("The VBA code was successfully exported from the document {0} to the folder {1} ", doc.Name, settings.TargetFolder),
                            "VBA Import Export", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        break;

                    case "ExportVBAFolder":
                        var exportFolderBrowser = new FolderBrowser2
                        {
                            DirectoryPath = settings.TargetFolder
                        };
                        if (exportFolderBrowser.ShowDialog(null) == DialogResult.OK)
                        {
                            settings.TargetFolder = exportFolderBrowser.DirectoryPath;
                            SettingsManager.Store(doc, settings);

                            OnCommand("ExportVBA");
                        }
                        return;

                    case "ImportVBA":
                        if (string.IsNullOrEmpty(settings.TargetFolder))
                        {
                            OnCommand("ImportVBAFolder");
                            return;
                        }

                        ImportVBA(settings.TargetFolder, settings);
                        MessageBox.Show(
                            string.Format("The VBA code was successfully imported from the folder {0} to the document {1}", settings.TargetFolder, doc.Name),
                            "VBA Import Export", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        break;

                    case "ImportVBAFolder":
                        var importFolderBrowser = new FolderBrowser2
                        {
                            DirectoryPath = settings.TargetFolder
                        };
                        if (importFolderBrowser.ShowDialog(null) == DialogResult.OK)
                        {
                            settings.TargetFolder = importFolderBrowser.DirectoryPath;
                            SettingsManager.Store(doc, settings);

                            OnCommand("ImportVBA");
                        }
                        return;

                    case "ClearBeforeImport":
                        {
                            settings.ClearBeforeImport = !settings.ClearBeforeImport;
                            SettingsManager.Store(doc, settings);
                            UpdateUI();
                        }
                        break;

                    case "IncludeStencils":
                        {
                            settings.IncludeStencils = !settings.IncludeStencils;
                            SettingsManager.Store(doc, settings);
                            UpdateUI();
                        }
                        break;
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "VBA Import Export", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        /// <summary>
        /// Callback called by UI manager.
        /// Should return if corresponding command should be enabled in the user interface.
        /// By default, all commands are enabled.
        /// </summary>
        public bool IsCommandEnabled(string commandId)
        {
            switch (commandId)
            {
                case "ddExportVBA":
                case "ddImportVBA":
                case "ExportVBA":
                case "ImportVBA":
                case "ExportVBAFolder":
                case "ImportVBAFolder":
                case "ClearBeforeImport":
                    return Application != null && Application.ActiveWindow != null;

                case "IncludeStencils":
                    return Application != null && Application.ActiveDocument != null &&
                           Application.ActiveWindow.Type == (short)Visio.VisWinTypes.visDrawing;

                default:
                    return true;
            }
        }

        /// <summary>
        /// Callback called by UI manager.
        /// Should return if corresponding command (button) is pressed or not (makes sense for toggle buttons)
        /// </summary>
        public bool IsCommandChecked(string commandId)
        {
            var settings = SettingsManager.LoadOrCreate(Application.ActiveDocument);

            switch (commandId)
            {
                case "ClearBeforeImport":
                    return settings.ClearBeforeImport;

                case "IncludeStencils":
                    return settings.IncludeStencils;

                default:
                    return false;
            }
        }
        /// <summary>
        /// Callback called by UI manager.
        /// Returns a label associated with given command.
        /// We assume for simplicity taht command labels are named simply named as [commandId]_Label (see resources)
        /// </summary>
        public string GetCommandLabel(string command)
        {
            return Resources.ResourceManager.GetString(command + "_Label");
        }

        /// <summary>
        /// Returns a bitmap associated with given command.
        /// We assume for simplicity that bitmap ids are named after command id.
        /// </summary>
        public Bitmap GetCommandBitmap(string id)
        {
            return (Bitmap)Resources.ResourceManager.GetObject(id);
        }

        internal void UpdateUI()
        {
            AddinUI.UpdateCommandBars();
            AddinUI.UpdateRibbon();
        }

        private void Application_DocumentListChanged(Visio.Document window)
        {
            UpdateUI();
        }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            var version = int.Parse(Application.Version, NumberStyles.AllowDecimalPoint);
            if (version < 14)
                AddinUI.StartupCommandBars("VisioImportExportVba", new[]
                {
                    "ExportVBA", 
                    "ExportVBAFolder", 
                    "",
                    "ImportVBA",
                    "ImportVBAFolder",
                    "",
                    "ClearBeforeImport",
                    "IncludeStencils"
                });

            Application.DocumentOpened += Application_DocumentListChanged;
            Application.BeforeDocumentClose += Application_DocumentListChanged;
            Application.DocumentCreated += Application_DocumentListChanged;
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            AddinUI.ShutdownCommandBars();

            Application.DocumentOpened -= Application_DocumentListChanged;
            Application.BeforeDocumentClose -= Application_DocumentListChanged;
            Application.DocumentCreated -= Application_DocumentListChanged;
        }


        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }

    }
}
