using FiletypeConverter.Interfaces;
using FiletypeConverter.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FiletypeConverter.Parsers
{
    public abstract class FileParser : IFileParser
    {
        public FileParser()
        {
        }

        public abstract string ContentAsString { get; }

        public string Path { get; set; }
        public List<string> Journal { get; private set; } = new List<string>();
        public List<string> Errors { get; private set; } = new List<string>();
        public IOutputSupplier Output { get; set; } = new OutputSupplier();

        public delegate void OutputAdded(string message);
        public event OutputAdded ErrorAdded;
        public event OutputAdded JournalAdded;

        public abstract bool Parse();
        public abstract bool Parse(string path);
    }
}
