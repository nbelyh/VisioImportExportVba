using CommandLine;
using CommandLine.Text;
using System;
using System.Collections.Generic;
using ImportExportVbaLib;
using Visio = Microsoft.Office.Interop.Visio;
using System.IO;

namespace vba_import_export
{
    [Verb("import", HelpText = "Import VBA code from the specified folder.")]
    public class ImportOptions
    {
        [Value(0, Required = true, HelpText = "List of Visio files to process, VBA code will be added to each of these files.")]
        public IEnumerable<string> InputFiles { get; set; }

        [Option('i', "input-directory", Required = false, HelpText = "Input directory that contains VBA modules to import. Current directory by default.")]
        public string InputDirectory { get; set; }

        [Option('s', "include-stencils", Required = false, HelpText = "Import also code into docked stencils.")]
        public bool IncludeStencils { get; set; }

        [Option('c', "clear", Required = false, HelpText = "Before importing new VBA code, remove all existing VBA code.")]
        public bool ClearBeforeImport { get; set; }
    }

    [Verb("export", HelpText = "Export the VBA to the specified folder")]
    public class ExportOptions
    {
        [Option('o', "output-directory", Required = false, HelpText = "Target directory to export VBA modules. If it does not exist, it will be created. Current directory by default.")]
        public string OutputDirectory { get; set; }

        [Option('s', "include-stencils", Required = false, HelpText = "Include docked stencils in export.")]
        public bool IncludeStencils { get; set; }

        [Value(0, Required = true, HelpText = "Visio file to process (source file).")]
        public string InputFile { get; set; }

        [Usage]
        public static IEnumerable<Example> Examples
        {
            get
            {
                var settings = new UnParserSettings
                {
                    PreferShortName = true
                };

                yield return new Example("Export to current directory", settings, new ExportOptions { InputFile = "file1.vsd" });
                yield return new Example("Export to other directory", settings, new ExportOptions { InputFile = "file1.vsd", OutputDirectory = "c:\\dir\\other"});
            }
        }
    }

    class Program
    {
        static int Main(string[] args)
        {
            return Parser.Default.ParseArguments<ImportOptions, ExportOptions>(args).MapResult(
                (ImportOptions opt) => Import(opt),
                (ExportOptions opt) => Export(opt),
                errs => 1);
        }
        
        private static int Export(ExportOptions opt)
        {
            var app = new Visio.InvisibleApp();

            var settings = new Settings
            {
                IncludeStencils = opt.IncludeStencils
            };

            var doc = app.Documents.OpenEx(opt.InputFile,
                (short)Visio.VisOpenSaveArgs.visOpenCopy | (short)Visio.VisOpenSaveArgs.visOpenRO);

            var path = string.IsNullOrEmpty(opt.OutputDirectory)
                ? Environment.CurrentDirectory
                : Path.IsPathRooted(opt.OutputDirectory)
                ? opt.OutputDirectory
                : Path.Combine(Environment.CurrentDirectory, opt.OutputDirectory);

            VisioVBA.ExportVBA(doc, path, settings);

            app.Quit();
            return 0;
        }

        private static int Import(ImportOptions opt)
        {
            var app = new Visio.InvisibleApp();

            var settings = new Settings
            {
                ClearBeforeImport = opt.ClearBeforeImport,
                IncludeStencils = opt.IncludeStencils
            };

            foreach (var inputFile in opt.InputFiles)
            {
                var doc = app.Documents.OpenEx(inputFile,
                    (short)Visio.VisOpenSaveArgs.visOpenRW);

                var path = string.IsNullOrEmpty(opt.InputDirectory)
                    ? Environment.CurrentDirectory
                    : Path.IsPathRooted(opt.InputDirectory)
                    ? opt.InputDirectory
                    : Path.Combine(Environment.CurrentDirectory, opt.InputDirectory);

                VisioVBA.ImportVBA(doc, path, settings);
            }

            app.Quit();
            return 0;
        }
    }
}
