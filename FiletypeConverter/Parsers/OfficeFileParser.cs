using FiletypeConverter.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FiletypeConverter.Parsers
{
    public class OfficeFileParser : FileParser
    {
        public override bool Parse(string path)
        {
            // Office files do not need to be parsed for now -->
            return true;
        }
    }
}
