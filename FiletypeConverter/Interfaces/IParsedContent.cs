using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FiletypeConverter.Interfaces
{
    public interface IParsedContent
    {
        string ToString();
        string ContentAsString { get; }
        string FolderName { get; }
        string IdentifyingName { get; }
    }
}
