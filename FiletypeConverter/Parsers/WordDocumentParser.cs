using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SautinSoft.Document;

namespace FiletypeConverter
{
    class WordDocumentParser
    {
        public string Path { get; private set; }

        public string DocAsString { get; private set; }
        public string AsString => DocAsString;

        private DocumentCore document = null;

        public WordDocumentParser()
        {

        }

        public WordDocumentParser(string path):this()
        {
            Path = path;
        }

        public bool parse()
        {
            if (string.IsNullOrEmpty(Path))
            {
                throw new ArgumentException("Path not set");
            }
            return parse(Path);
        }

        public bool parse(string path)
        {
            return parseUsingSautinSoft(path);
        }

        private bool parseUsingSautinSoft(string path)
        {
            if (!File.Exists(path))
            {
                // TODO: logging
                return false;
            }

            DocumentCore doc = null;

            try
            {
                doc = DocumentCore.Load(path);
                string tempPath = System.IO.Path.GetTempPath();
                FileInfo fi = new FileInfo(path);

                string tmpFilename = System.IO.Path.ChangeExtension(System.IO.Path.Combine(tempPath, fi.Name), ".txt");

                doc.Save(tmpFilename, SaveOptions.TxtDefault);

                DocAsString = File.ReadAllText(tmpFilename);
                File.Delete(tmpFilename);
            }
            catch (Exception ex)
            {
                return false;
            }

            return true;
        }

        /// <summary>
        /// Replace trial text in generated text.
        /// </summary>
        /// <param name="orig"></param>
        /// <returns></returns>
        private string replaceTrialText(string orig) => orig.Replace(@"Created by the trial version of Document .Net 3.3.3.26!
The trial version sometimes inserts ""trial"" into random places.
Get the full version of Document .Net.
Created by the trial version of Document .Net 3.3.3.26!
The trial version sometimes inserts ""trial"" into random places.
Get the full version of Document .Net.
", string.Empty);

    }
}
