﻿using System.Drawing;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

namespace VisioImportExportVba
{
    /// <summary>
    /// User interface manager for Visio 2010 and above
    /// Creates and controls ribbon UI
    /// </summary>
    /// 

    [ComVisible(true)]
    public partial class AddinUI : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI _ribbon;

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonId)
        {
            return Properties.Resources.Ribbon;
        }

        #endregion

        #region Ribbon Callbacks

        public bool IsRibbonCommandEnabled(Office.IRibbonControl ctrl)
        {
            return Globals.ThisAddIn.IsCommandEnabled(ctrl.Id);
        }

        public bool IsRibbonCommandChecked(Office.IRibbonControl ctrl)
        {
            return Globals.ThisAddIn.IsCommandChecked(ctrl.Id);
        }

        public void OnRibbonButtonCheckClick(Office.IRibbonControl control, bool pressed)
        {
            Globals.ThisAddIn.OnCommand(control.Id);
        }

        public void OnRibbonButtonClick(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.OnCommand(control.Id);
        }

        public string OnGetRibbonLabel(Office.IRibbonControl control)
        {
            return Globals.ThisAddIn.GetCommandLabel(control.Id);
        }

        public string OnGetRibbonScreentip(Office.IRibbonControl control)
        {
            return Globals.ThisAddIn.GetCommandScreentip(control.Id);
        }

        public string OnGetRibbonSupertip(Office.IRibbonControl control)
        {
            return Globals.ThisAddIn.GetCommandSupertip(control.Id);
        }

        public void OnRibbonLoad(Office.IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;
        }

        public Bitmap GetRibbonImage(Office.IRibbonControl control)
        {
            return Globals.ThisAddIn.GetCommandBitmap(control.Id);
        }

        #endregion

        public void UpdateRibbon()
        {
            if (_ribbon != null)
                _ribbon.Invalidate();
        }
    }
}
