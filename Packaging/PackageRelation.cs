using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace TSVCEO.OOXML.Packaging
{
    public class PackageRelation
    {
        public string Id { get; set; }
        public string Type { get; set; }
        public string TargetName { get; set; }
        public PackageFile Target { get; set; }
        public bool IsExternal { get; set; }

        public static PackageRelation FromXElement(XElement relation, PackageEntry packageentry)
        {
            string id = relation.Attribute("Id").Value;
            string type = relation.Attribute("Type").Value;
            string targetname = relation.Attribute("Target").Value;
            string targetmode = relation.Attributes("TargetMode").Select(a => a.Value).SingleOrDefault();
            PackageFile target = null;

            if (targetmode != "External")
            {
                target = packageentry.Parent[targetname] as PackageFile;
            }

            return new PackageRelation
            {
                Id = id,
                Type = type,
                TargetName = targetname,
                Target = target,
                IsExternal = targetmode == "External"
            };
        }

        public XElement ToXElement()
        {
            return new XElement(xmlns.relationships + "Relationship",
                new XAttribute("Id", this.Id),
                new XAttribute("Type", this.Type),
                new XAttribute("Target", this.TargetName),
                this.IsExternal ? new XAttribute("TargetMode", "External") : null
            );
        }

        public static string stylesWithEffects = "http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects";
        public static string customXml = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml";
        public static string customXmlProps = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps";
        public static string endnotes = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes";
        public static string extendedProperties = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
        public static string fontTable = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable";
        public static string footer = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";
        public static string footnotes = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes";
        public static string header = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";
        public static string numbering = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";
        public static string officeDocument = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
        public static string settings = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings";
        public static string styles = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
        public static string theme = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
        public static string webSettings = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings";
        public static string coreProperties = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
    }
}
