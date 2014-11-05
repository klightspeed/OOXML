using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using System.IO;
using Ionic.Zip;

namespace TSVCEO.OOXML.Packaging
{
    public class PackageFile : PackageEntry
    {
        protected XDocument _XmlDocument;
        protected byte[] _Data;

        public virtual byte[] Data
        {
            get
            {
                if (_Data != null)
                {
                    return _Data;
                }
                else
                {
                    return GetBytes(this.XmlDocument);
                }
            }
            set
            {
                _XmlDocument = null;
                _Data = value;
            }
        }

        public virtual XDocument XmlDocument
        {
            get
            {
                if (_XmlDocument != null)
                {
                    return _XmlDocument;
                }
                else if (_Data != null)
                {
                    _XmlDocument = LoadXDocument(_Data);
                    _Data = null;
                    return _XmlDocument;
                }
                else
                {
                    throw new InvalidOperationException();
                }
            }
        }

        public override void LoadEntry(ZipEntry entry, string[] pathcomponents)
        {
            Name = pathcomponents[0];

            using (MemoryStream memstrm = new MemoryStream())
            {
                entry.Extract(memstrm);
                _Data = memstrm.ToArray();
            }
        }

        public override void Save(ZipFile zip, string path)
        {
            string name = (String.IsNullOrEmpty(path) ? "" : path + "/") + this.Name;

            zip.AddEntry(name, Data);
            this.SaveRelations(zip, path, this.Name);
        }
    }
}
