using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace TSVCEO.OOXML.Packaging
{
    public class PackageContentTypes : PackageFile
    {
        public override byte[] Data
        {
            get
            {
                return base.Data;
            }
            set
            {
                throw new InvalidOperationException();
            }
        }

        public override XDocument XmlDocument
        {
            get
            {
                return new XDocument(
                    new XElement(xmlns.contentTypes + "Types",
                        new XAttribute("xmlns", xmlns.contentTypes.NamespaceName),
                        new XElement(xmlns.contentTypes + "Default",
                            new XAttribute("Extension", "rels"),
                            new XAttribute("ContentType", PackageContentTypes.relationships)
                        ),
                        new XElement(xmlns.contentTypes + "Default",
                            new XAttribute("Extension", "xml"),
                            new XAttribute("ContentType", "application/xml")
                        ),
                        Package.GetAllFiles("xml").Select(f =>
                            f.ContentType == null ? null : new XElement(xmlns.contentTypes + "Override",
                                new XAttribute("PartName", f.Path),
                                new XAttribute("ContentType", f.ContentType)
                            )
                        )
                    )
                );
            }
        }

        public static string stylesWithEffects = "application/vnd.ms-word.stylesWithEffects+xml";
        public static string customXmlProperties = "application/vnd.openxmlformats-officedocument.customXmlProperties+xml";
        public static string extendedProperties = "application/vnd.openxmlformats-officedocument.extended-properties+xml";
        public static string theme = "application/vnd.openxmlformats-officedocument.theme+xml";
        public static string document = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
        public static string endnotes = "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml";
        public static string fontTable = "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml";
        public static string footer = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml";
        public static string footnotes = "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml";
        public static string header = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml";
        public static string numbering = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml";
        public static string settings = "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml";
        public static string styles = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";
        public static string template = "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml";
        public static string webSettings = "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml";
        public static string coreProperties = "application/vnd.openxmlformats-package.core-properties+xml";
        public static string relationships = "application/vnd.openxmlformats-package.relationships+xml";
        public static string xml = "application/xml";
    }
}
