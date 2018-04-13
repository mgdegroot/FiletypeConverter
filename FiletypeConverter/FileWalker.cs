using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FiletypeConverter
{
    class FileWalker
    {
        public static List<string> WalkDir(string rootDir, string pattern="*", bool recursive = true)
        {
            List<string> foundFiles = new List<string>();

            try
            {
                foreach (string f in Directory.GetFiles(rootDir, pattern, recursive == true ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly))
                {
                    foundFiles.Add(f);
                }
            }
            catch (Exception ex)
            {
                // TODO: logging
            }

            return foundFiles;
        }

        public string convertToText(string srcPath)
        {
            return string.Empty;
        }
    }
}
