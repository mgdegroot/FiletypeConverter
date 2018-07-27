using FiletypeConverter.Converters;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FiletypeConverter.Utils
{
    public class Util
    {
        public static string PathSeparator = @"\";

        public static IDictionary<FileType, string[]> FileExtensions { get; set; } = new Dictionary<FileType, string[]>
        {
            { FileType.EXCEL, new []{ "*.xls?" } },
            { FileType.WORD, new [] { "*.doc?" } },
            { FileType.POWERPOINT, new [] { "*.ppt?" } },
            { FileType.OUTLOOK_MSG, new [] { "*.msg" } },
            { FileType.OUTLOOK_PST, new [] { "*.pst?" } },
            { FileType.PDF, new [] { "*.pdf" } },
            { FileType.IMAGES, new [] { "*.jpg", "*.jpeg", "*.png", "*.gif", "*.bmp", "*.pdf" } },
        };
    }
}
