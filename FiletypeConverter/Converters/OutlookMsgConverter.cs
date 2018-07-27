using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FiletypeConverter.Interfaces;
using FiletypeConverter.Parsers;
using FiletypeConverter.Utils;
using log4net;
using OfficeConverter;

namespace FiletypeConverter.Converters
{
    public class OutlookMsgConverter : FileConverter
    {
        public OutlookMsgConverter(IFileParser fileParser, IOutputSupplier outputSupplier) : base(fileParser, outputSupplier)
        {
        }

        public override async Task ProcessInBackgroundAsync(ConvertConfig config)
        {
            KeepIntermediateFiles = config.KeepIntermediateFiles;

            if (!Directory.Exists(config.OutputDir))
            {
                Directory.CreateDirectory(config.OutputDir);
            }

            Output.AddJournalEntry("Converting Outlook message files.");
            await processMsgFiles(config.RootDir, config.OutputDir);
        }

        private async Task processMsgFiles(string rootPath, string outputPath)
        {
            await Task.Run(async () => {
                FileAttributes attr = File.GetAttributes(rootPath);

                if ((attr & FileAttributes.Directory) == FileAttributes.Directory)
                {
                    if (!outputPath.EndsWith(System.IO.Path.DirectorySeparatorChar.ToString()))
                    {
                        outputPath += System.IO.Path.DirectorySeparatorChar;
                    }

                    List<string> matchingFiles = FileWalker.WalkDir(rootPath, Util.FileExtensions[FileType.OUTLOOK_MSG], true);
                    foreach (var filename in matchingFiles)
                    {
                        string nwFilename = filename.Replace(rootPath, outputPath);

                        processSingleMsgFile(filename, nwFilename);
                    }
                }
                else
                {
                    string nwFilename = rootPath.Replace(rootPath, outputPath);
                    processSingleMsgFile(rootPath, nwFilename);
                }
            });
        }

        private void processSingleMsgFile(string inFile, string outFile)
        {
            FileInfo nwFileInfo = new FileInfo(outFile);
            if (!nwFileInfo.Exists)
            {
                Directory.CreateDirectory(nwFileInfo.Directory.FullName);
            }

            string outFileTxt = outFile + ".txt",
                outFilePdf = outFile + ".pdf";

            Output.AddJournalEntry($"Original: {inFile}. New: {outFile}");
            var parser = new OutlookMsgParser(inFile);
            if (parser.Parse())
            {
                IParsedContent parsedContent = parser.ParsedContent.Single<IParsedContent>();
                string result = parsedContent.ContentAsString;
                File.WriteAllText(outFileTxt, result);
                Converter.ConvertFromCom(outFileTxt, outFilePdf);

                if (!KeepIntermediateFiles)
                {
                    File.Delete(outFileTxt);
                }

                //extractAndConvertAttachements(inFile, outFile);
            }
            else
            {
                Output.AddLogEntry($"Failed to parse file {inFile}", true);
            }
        }

        public async void extractAndConvertAttachements(string origMsgPath, string destMsgPath)
        {
            MsgReader.Reader reader = new MsgReader.Reader();

            string msgFile = origMsgPath;
            string attachementDestDir = destMsgPath + "_bijlages";

            if (!Directory.Exists(attachementDestDir))
            {
                Directory.CreateDirectory(attachementDestDir);
            }

            string[] outputFiles = reader.ExtractToFolder(msgFile, attachementDestDir);

            FileConverter.ConvertConfig convertConfig = new FileConverter.ConvertConfig()
            {
                SourceFiles = FileType.OUTLOOK_MSG | FileType.WORD | FileType.POWERPOINT | FileType.EXCEL | FileType.IMAGES,
                RootDir = attachementDestDir,
                OutputDir = attachementDestDir + "_pdf",
                Filter = "*",
            };


            FileConverter outlookFileConverter = new OutlookMsgConverter(new OutlookMsgParser(), Output);
            await outlookFileConverter.ProcessInBackgroundAsync(convertConfig);

            // TODO: refactor to enable dependency injection -->
            FileConverter officeFileConverter = new OfficeFileConverter(new OfficeFileParser(), Output);

            await officeFileConverter.ProcessInBackgroundAsync(convertConfig);

            FileConverter fileTransferrer = new ImageFileConverter(new ImageFileParser(), Output);
            await fileTransferrer.ProcessInBackgroundAsync(convertConfig);
        }

    }

}
