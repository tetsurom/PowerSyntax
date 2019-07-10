using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace PowerSyntax
{
    public static class ManifestResource
    {
        public static IEnumerable<string> GetFiles(string path, string type)
        {
            string normalized = path.Replace('/', '.').Replace('\\', '.');
            if(type.StartsWith("*"))
            {
                type = type.Substring(1);
            }
            return Assembly.GetExecutingAssembly().GetManifestResourceNames().Where(x => x.StartsWith(normalized) && x.EndsWith(type));
        }
        public static string ReadAllText(string path)
        {
            string normalized = path.Replace('/', '.').Replace('\\', '.');
            using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(normalized))
            {
                if (stream != null)
                {
                    using (var reader = new StreamReader(stream))
                    {
                        return reader.ReadToEnd();
                    }
                }
                throw new FileNotFoundException(path);
            }
        }
    }

    public class ManifestResourceProvider : IResourceProvider
    {
        public ManifestResourceProvider(string root)
        {
            this.root = root;
        }
        public IEnumerable<string> GetFiles(string path, string type)
        {
            return ManifestResource.GetFiles(root + path, type);
        }

        public string ReadAllText(string path)
        {
            return ManifestResource.ReadAllText(root + path);
        }

        private string root;
    }
}
