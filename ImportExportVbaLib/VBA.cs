using System;
using System.IO;
using System.Linq;

namespace ImportExportVbaLib
{
    public static class VBA
    {
        static void ImportThisDocumentVBA(string filePath, dynamic project)
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

        public static void ImportOneDocumentVBA(dynamic project, string path, Settings settings)
        {
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

        static string GetComponentFileExtension(dynamic component)
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

        public static void ExportDocumentVBA(dynamic project, string path)
        {
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
    }
}
