using FiletypeConverter.Converters;
using OfficeConverter;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FiletypeConverter.Interfaces
{
    interface IFileConverter
    {
        IOutputSupplier Output { get; set; }
        ConvertTarget TargetFormat { get; set; }
        IConverter Converter { get; set; }
        Task ProcessInBackgroundAsync(FileConverter.ConvertConfig config);
    }
}
