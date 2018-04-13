using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeConverter;

namespace FiletypeConverter
{
    class Tester
    {
        private OfficeConverter.Converter converter;

        public static void TestWord()
        {
            var inputFile = @"E:\test\test.doc";
            var outputFile = @"E:\test\convertWordOutput.pdf";
            new Converter().Convert(inputFile, outputFile);
        }

        public static void TestPowerpoint()
        {
            var inputFile = @"E:\test\test.ppt";
            var outputFile = @"E:\test\convertPowerpointOutput.pdf";
            new Converter().Convert(inputFile, outputFile);
        }
    }
}
