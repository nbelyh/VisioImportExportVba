using System;
using System.IO;
using System.Linq;
using Visio = Microsoft.Office.Interop.Visio;

namespace ImportExportVbaLib
{
    public static class VisioVBA
    {
        static Visio.Window FindDocumentWindow(Visio.Application app, int docId)
        {
            return app.Windows
                .Cast<Visio.Window>()
                .FirstOrDefault(w => w.Document != null && w.Document.ID == docId);
        }

        public static void ImportVBA(Visio.Document doc, string folder, Settings settings)
        {
            if (doc.Type != Visio.VisDocumentTypes.visTypeDrawing &&
                doc.Type != Visio.VisDocumentTypes.visTypeTemplate &&
                doc.Type != Visio.VisDocumentTypes.visTypeStencil)
                return;

            VBA.ImportOneDocumentVBA(doc.VBProject, folder, settings);

            if (!settings.IncludeStencils)
                return;

            if (doc.Type != Visio.VisDocumentTypes.visTypeDrawing && 
                doc.Type != Visio.VisDocumentTypes.visTypeTemplate)
                return;

            var app = doc.Application;

            var window = FindDocumentWindow(app, doc.ID);
            if (window == null)
                return;

            Array stencilNames;
            window.DockedStencils(out stencilNames);
            foreach (string stencilName in stencilNames)
            {
                var stencilDoc = app.Documents[stencilName];

                var readOnly = stencilDoc.ReadOnly != 0;
                if (readOnly)
                {
                    stencilDoc.Close();
                    short flags = (short) Visio.VisOpenSaveArgs.visOpenDocked | (short) Visio.VisOpenSaveArgs.visOpenRW;

                    stencilDoc = app.Documents.OpenEx(stencilName, flags);
                }

                VBA.ImportOneDocumentVBA(stencilDoc.VBProject, Path.Combine(folder, stencilDoc.Name), settings);
            }
        }
        
        public static void ExportVBA(Visio.Document doc, string folder, Settings settings)
        {
            if (doc.Type != Visio.VisDocumentTypes.visTypeDrawing &&
                doc.Type != Visio.VisDocumentTypes.visTypeTemplate &&
                doc.Type != Visio.VisDocumentTypes.visTypeStencil)
                return;

            VBA.ExportDocumentVBA(doc.VBProject, folder);

            if (!settings.IncludeStencils)
                return;

            if (doc.Type != Visio.VisDocumentTypes.visTypeDrawing &&
                doc.Type != Visio.VisDocumentTypes.visTypeTemplate)
                return;

            var app = doc.Application;

            var window = FindDocumentWindow(app, doc.ID);
            if (window == null)
                return;

            Array stencilNames;
            window.DockedStencils(out stencilNames);
            foreach (var stencilName in stencilNames)
            {
                var stencilDoc = app.Documents[stencilName];
                VBA.ExportDocumentVBA(stencilDoc.VBProject, Path.Combine(folder, stencilDoc.Name));
            }
        }
    }
}
