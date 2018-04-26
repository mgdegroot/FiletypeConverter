using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FiletypeConverter
{
    interface IFileConverter
    {
        Task processInBackgroundAsync(FileConverter.ConvertConfig config);
    }
}
