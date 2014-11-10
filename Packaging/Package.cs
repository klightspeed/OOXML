using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Ionic.Zip;
using System.IO;

namespace TSVCEO.OOXML.Packaging
{
    public class Package : PackageDirectory
    {
        protected Dictionary<string, string> DefaultContentTypes { get; set; }

        protected void LoadContentTypes()
        {
            DefaultContentTypes = new Dictionary<string, string>();
            PackageFile file = this["[Content_Types].xml"] as PackageFile;
            XDocument doc = file.XmlDocument;

            foreach (XElement el in doc.Root.Elements(xmlns.contentTypes + "Default"))
            {
                string extension = el.Attribute("Extension").Value;
                string type = el.Attribute("ContentType").Value;
                DefaultContentTypes[extension] = type;
            }

            foreach (XElement el in doc.Root.Elements(xmlns.contentTypes + "Override"))
            {
                string name = el.Attribute("PartName").Value;
                string type = el.Attribute("ContentType").Value;

                this[name].ContentType = type;
            }

            Entries["[content_types].xml"] = new PackageContentTypes
            {
                Package = this,
                Parent = this,
                Name = "[Content_Types].xml",
                ContentType = null
            };
        }

        public static Package Load(string filename)
        {
            using (Stream stream = File.OpenRead(filename))
            {
                return Load(stream);
            }
        }

        public static Package Load(Stream stream)
        {
            Package pkg = new Package();
            pkg.Package = pkg;
            pkg.Parent = pkg;

            using (ZipFile zip = ZipFile.Read(stream))
            {
                foreach (var entry in zip.Entries)
                {
                    string[] pathcomponents = new string[] { "" }.Concat(entry.FileName.Split('\\', '/').Where(s => s != "")).ToArray();
                    pkg.LoadEntry(entry, pathcomponents);
                }
            }

            pkg.LoadContentTypes();
            pkg.LoadRelations();

            return pkg;
        }

        public void SaveAs(string filename)
        {
            using (Stream stream = File.Create(filename))
            {
                Save(stream);
            }
        }

        public void Save(Stream stream)
        {
            using (ZipFile zip = new ZipFile())
            {
                this.Save(zip, null);

                zip.Save(stream);
            }
        }
    }
}
