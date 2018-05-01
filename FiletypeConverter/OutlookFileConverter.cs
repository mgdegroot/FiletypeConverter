using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using log4net;
using OfficeConverter;

namespace FiletypeConverter
{
    public class OutlookFileConverter : FileConverter, IFileConverter
    {
        private Converter converter = new Converter();


        public bool KeepMsgTxtFiles { get; set; } = false;


        public OutlookFileConverter(ILog log) : base(log)
        {
        }

        public override async Task processInBackgroundAsync(ConvertConfig config)
        {
            if (!Directory.Exists(config.OutputDir))
            {
                Directory.CreateDirectory(config.OutputDir);
            }

            if (config.ProcessOutlookMsg)
            {
                updateLogAndJournal("Converting Outlook message files.", null);
                await processMsgFiles(config.RootDir, config.OutputDir);
            }
        }

        private async Task processMsgFiles(string rootPath, string outputPath)
        {
            await Task.Run(async () => {
                FileAttributes attr = File.GetAttributes(rootPath);

                if ((attr & FileAttributes.Directory) == FileAttributes.Directory)
                {
                    if (!outputPath.EndsWith(Path.DirectorySeparatorChar.ToString()))
                    {
                        outputPath += Path.DirectorySeparatorChar;
                    }

                    List<string> matchingFiles = FileWalker.WalkDir(rootPath, "*.msg", true);
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

            updateLogAndJournal($"Original: {inFile}. New: {outFile}", null);
            var parser = new OutlookMsgParser(inFile, log);
            if (parser.parse())
            {
                string result = parser.MsgAsString;
                File.WriteAllText(outFileTxt, result);
                converter.Convert(outFileTxt, outFilePdf);

                if (!KeepMsgTxtFiles)
                {
                    File.Delete(outFileTxt);
                }

                //extractAndConvertAttachements(inFile, outFile);
            }
            else
            {
                updateLogAndJournal(null, $"Failed to parse file {inFile}", true);
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
                ProcessOutlookMsg = true,
                ProcessWord = true,
                ProcessPowerpoint = true,
                ProcessExcel = true,
                ProcessImages = true,
                RootDir = attachementDestDir,
                OutputDir = attachementDestDir + "_pdf",
                Filter = "*",
            };


            FileConverter outlookFileConverter = new OutlookFileConverter(log);
            await outlookFileConverter.processInBackgroundAsync(convertConfig);

            FileConverter officeFileConverter = new OfficeFileConverter(log);

            await officeFileConverter.processInBackgroundAsync(convertConfig);

            FileConverter fileTransferrer = new ImageFileConverter(log);
            await fileTransferrer.processInBackgroundAsync(convertConfig);
        }
    }

}
