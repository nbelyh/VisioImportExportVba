using System;

namespace ImportExportVbaLib
{
    [Serializable]
    public class Settings
    {
        public bool ClearBeforeImport { get; set; }
        public bool IncludeStencils { get; set; }
        public string TargetFolder { get; set; }
    }
}