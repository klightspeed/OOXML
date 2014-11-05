using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ionic.Zip;

namespace TSVCEO.OOXML.Packaging
{
    public class PackageDirectory : PackageEntry, IEnumerable<PackageEntry>
    {
        protected Dictionary<string, PackageEntry> Entries { get; set; }

        public PackageDirectory()
        {
            Entries = new Dictionary<string, PackageEntry>();
        }

        public override void LoadEntry(ZipEntry entry, string[] pathcomponents)
        {
            Name = pathcomponents[0];

            if (Entries.ContainsKey(pathcomponents[1].ToLower()))
            {
                Entries[pathcomponents[1].ToLower()].LoadEntry(entry, pathcomponents.Skip(1).ToArray());
            }
            else
            {
                if (pathcomponents.Length == 2)
                {
                    Entries[pathcomponents[1].ToLower()] = new PackageFile { Package = this.Package, Parent = this };
                    Entries[pathcomponents[1].ToLower()].LoadEntry(entry, pathcomponents.Skip(1).ToArray());
                }
                else
                {
                    Entries[pathcomponents[1].ToLower()] = new PackageDirectory { Name = pathcomponents[1], Package = this.Package, Parent = this };
                    Entries[pathcomponents[1].ToLower()].LoadEntry(entry, pathcomponents.Skip(1).ToArray());
                }
            }
        }

        public override void Save(ZipFile zip, string path)
        {
            path = (String.IsNullOrEmpty(path) ? "" : path + "/") + this.Name;
            this.SaveRelations(zip, path, "");

            foreach (PackageEntry entry in Entries.Values)
            {
                entry.Save(zip, path);
            }
        }

        protected PackageEntry GetEntry(string[] pathcomponents)
        {
            if (pathcomponents[0] == "..")
            {
                return Parent.GetEntry(pathcomponents.Skip(1).ToArray());
            }

            PackageEntry entry = Entries[pathcomponents[0].ToLower()];

            if (pathcomponents.Length == 1)
            {
                return entry;
            }
            else if (entry is PackageDirectory)
            {
                return ((PackageDirectory)entry).GetEntry(pathcomponents.Skip(1).ToArray());
            }
            else
            {
                throw new KeyNotFoundException();
            }
        }

        protected bool ContainsEntry(string[] pathcomponents)
        {
            if (Entries.ContainsKey(pathcomponents[0].ToLower()))
            {
                if (pathcomponents.Length == 1)
                {
                    return true;
                }
                else if (Entries[pathcomponents[0].ToLower()] is PackageDirectory)
                {
                    return ((PackageDirectory)Entries[pathcomponents[0].ToLower()]).ContainsEntry(pathcomponents.Skip(1).ToArray());
                }
            }

            return false;
        }

        public PackageEntry this[string name]
        {
            get
            {
                string[] pathcomponents = name.Split('\\', '/').Where(s => s != "").ToArray();

                if (name.StartsWith("/") || name.StartsWith("\\"))
                {
                    return Package.GetEntry(pathcomponents);
                }
                else
                {
                    return GetEntry(pathcomponents);
                }
            }
        }

        public IEnumerable<PackageFile> GetAllFiles(string extension)
        {
            foreach (PackageEntry entry in this.Entries.Values)
            {
                if (entry is PackageDirectory)
                {
                    foreach (PackageFile pkgfile in ((PackageDirectory)entry).GetAllFiles(extension))
                    {
                        yield return pkgfile;
                    }
                }
                else if (entry is PackageFile)
                {
                    if (entry.Name.EndsWith(extension))
                    {
                        yield return (PackageFile)entry;
                    }
                }
            }
        }

        public bool ContainsEntry(string name)
        {
            string[] pathcomponents = name.Split('\\', '/').Where(s => s != "").ToArray();
            return ContainsEntry(pathcomponents);
        }

        public IEnumerator<PackageEntry> GetEnumerator()
        {
            foreach (PackageEntry entry in Entries.Values)
            {
                yield return entry;
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public void LoadRelations()
        {
            if (this.ContainsEntry("_rels/.rels"))
            {
                PackageFile relationsfile = this["_rels/.rels"] as PackageFile;
                LoadRelations(relationsfile);
                relationsfile.Delete();
            }

            foreach (PackageEntry pkgent in this)
            {
                if (pkgent is PackageFile)
                {
                    PackageFile pkgfile = (PackageFile)pkgent;
                    if (this.ContainsEntry("_rels/" + pkgfile.Name + ".rels"))
                    {
                        PackageFile relationsfile = this["_rels/" + pkgfile.Name + ".rels"] as PackageFile;
                        pkgfile.LoadRelations(relationsfile);
                        relationsfile.Delete();
                    }
                }
                else if (pkgent is PackageDirectory)
                {
                    ((PackageDirectory)pkgent).LoadRelations();
                }
            }
        }

        public void Delete(string name)
        {
            if (Entries.ContainsKey(name.ToLower()))
            {
                Entries.Remove(name.ToLower());
            }
            else if (this.ContainsEntry(name))
            {
                this[name].Delete();
            }
        }
    }
}
