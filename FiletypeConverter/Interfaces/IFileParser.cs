using FiletypeConverter.ParsedContent;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FiletypeConverter.Interfaces
{
    public interface IFileParser
    {
        string Path { get; set; }
        //string ContentAsString { get; }
        IOutputSupplier Output { get; set; }
        IList<IParsedContent> ParsedContent { get; }
        bool Parse();
        bool Parse(string path);


    }
}
