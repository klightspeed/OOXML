using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Ionic.Zip;

namespace TSVCEO.OOXML.Packaging
{
    public abstract class PackageEntry
    {
        public string Name { get; set; }
        public string ContentType { get; set; }
        public Package Package { get; set; }
        public PackageDirectory Parent { get; set; }
        public PackageRelation[] Relations { get; set; }

        public string Path
        {
            get
            {
                return (Parent == Package ? "" : Parent.Path) + "/" + Name;
            }
        }

        public abstract void LoadEntry(ZipEntry entry, string[] pathcomponents);
        public abstract void Save(ZipFile zip, string path);

        protected static XDocument LoadXDocument(byte[] data)
        {
            return XDocument.Load(new MemoryStream(data));
        }

        protected static byte[] GetBytes(XDocument doc)
        {
            using (MemoryStream stream = new MemoryStream())
            {
                using (XmlTextWriter writer = new XmlTextWriter(stream, new UTF8Encoding(false)))
                {
                    writer.Formatting = Formatting.None;
                    doc.Save(writer);
                }
                return stream.ToArray();
            }
        }

        public void LoadRelations(PackageFile relationsFile)
        {
            XDocument doc = relationsFile.XmlDocument;
            List<PackageRelation> relations = new List<PackageRelation>();

            foreach (XElement relation in doc.Root.Elements(xmlns.relationships + "Relationship"))
            {
                relations.Add(PackageRelation.FromXElement(relation, this));
            }

            this.Relations = relations.ToArray();
        }

        protected void SaveRelations(ZipFile zip, string path, string name)
        {
            if (this.Relations != null)
            {
                XDocument relations = new XDocument(
                    new XElement(xmlns.relationships + "Relationships",
                        new XAttribute("xmlns", xmlns.relationships.NamespaceName),
                        this.Relations.Select(rel => rel.ToXElement())
                    )
                );

                zip.AddEntry(path + "/_rels/" + name + ".rels", GetBytes(relations));
            }
        }

        public IEnumerable<PackageFile> GetRelations(string type)
        {
            return Relations.Where(r => r.Type == type).Select(r => r.Target);
        }

        public PackageFile GetRelation(string type)
        {
            return GetRelations(type).Single();
        }

        public void Delete()
        {
            Parent.Delete(this.Name);
        }
    }
}
