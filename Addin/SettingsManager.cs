using System.IO;
using System.Xml;
using System.Xml.Serialization;
using Microsoft.Office.Interop.Visio;

namespace ImportExportVbaLib
{
    public class SettingsManager
    {
        private const string SolutionXmlNamespace = "http://unmanagedvisio.com";
        private const string SolutionXmlElementName = "VisioImportExportVBA";

        [XmlRoot("SolutionXML")]
        public class SolutionXmlElement
        {
            [XmlAttribute]
            public string Name { get; set; }

            [XmlElement(Namespace = SolutionXmlNamespace)]
            public Settings Options { get; set; }
        }

        public static Settings LoadOrCreate(Document doc)
        {
            return Load(doc) ?? new Settings
            {
                IncludeStencils = false,
                ClearBeforeImport = false
            };
        }

        private static Settings Load(Document doc)
        {
            if (doc == null)
                return null;

            if (!doc.SolutionXMLElementExists[SolutionXmlElementName])
                return null;

            var xml = doc.SolutionXMLElement[SolutionXmlElementName];

            if (xml == null)
                return null;

            using (var sr = new StringReader(xml))
            {
                var serializer = new XmlSerializer(typeof(SolutionXmlElement));

                var solutionXml = serializer.Deserialize(sr) as SolutionXmlElement;
                if (solutionXml == null)
                    return null;

                return solutionXml.Options;
            }
        }

        public static void Store(Document doc, Settings options)
        {
            if (doc == null)
                return;

            var solutionXml = new SolutionXmlElement
            {
                Name = SolutionXmlElementName,
                Options = options
            };

            var ns = new XmlSerializerNamespaces();
            ns.Add("uv", SolutionXmlNamespace);

            using (var sw = new StringWriter())
            {
                var xw = XmlWriter.Create(sw, new XmlWriterSettings
                {
                    OmitXmlDeclaration = true
                });

                var serializer = new XmlSerializer(typeof(SolutionXmlElement));

                serializer.Serialize(xw, solutionXml, ns);

                var xml = sw.ToString();
                doc.SolutionXMLElement[SolutionXmlElementName] = xml;
            }
        }
    }
}