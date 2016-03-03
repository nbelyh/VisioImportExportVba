using System;
using System.Drawing;
using System.Globalization;
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

        /// <summary>
        /// A simple command
        /// </summary>
        public void Command1()
        {
            MessageBox.Show(
                "Hello from command 1!",
                "VisioImportExportVba");
        }

        /// <summary>
        /// A command to demonstrate conditionally enabling/disabling.
        /// The command gets enabled only when a shape is selected
        /// </summary>
        public void Command2()
        {
            if (Application == null || Application.ActiveWindow == null || Application.ActiveWindow.Selection == null)
                return;

            MessageBox.Show(
                string.Format("Hello from (conditional) command 2! You have {0} shapes selected.", Application.ActiveWindow.Selection.Count),
                "VisioImportExportVba");
        }

        /// <summary>
        /// Callback called by the UI manager when user clicks a button
        /// Should do something meaningful when corresponding action is called.
        /// </summary>
        public void OnCommand(string commandId)
        {
            switch (commandId)
            {
                case "Command1":
                    Command1();
                    return;

                case "Command2":
                    Command2();
                    return;
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
                case "Command1":    // make command1 always enabled
                    return true;

                case "Command2":    // make command2 enabled only if a drawing is opened
                    return Application != null
                        && Application.ActiveWindow != null
                        && Application.ActiveWindow.Selection.Count > 0;
                default:
                    return true;
            }
        }

        /// <summary>
        /// Callback called by UI manager.
        /// Should return if corresponding command (button) is pressed or not (makes sense for toggle buttons)
        /// </summary>
        public bool IsCommandChecked(string command)
        {
            return false;
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

        private void Application_SelectionChanged(Visio.Window window)
        {
            UpdateUI();
        }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            var version = int.Parse(Application.Version, NumberStyles.AllowDecimalPoint);
            if (version < 14)
                AddinUI.StartupCommandBars("VisioImportExportVba", new[] { "Command1", "Command2" });
            Application.SelectionChanged += Application_SelectionChanged;

        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            AddinUI.ShutdownCommandBars();
            Application.SelectionChanged -= Application_SelectionChanged;

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
