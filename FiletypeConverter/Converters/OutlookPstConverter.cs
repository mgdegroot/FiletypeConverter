using FiletypeConverter.Interfaces;
using FiletypeConverter.ParsedContent;
using FiletypeConverter.Parsers;
using log4net;
using OfficeConverter;

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;
using FiletypeConverter.Utils;

namespace FiletypeConverter.Converters
{
    public interface IOutlookPstHandler
    {
        void ProcessFile(string inFile, string outDir);
    }

    class OutlookPstConverterDirect : IOutlookPstHandler
    {
        private IFileParser _outlookPstParser;
        private IConverter _converter;
        private ConvertTarget _convertTarget;
        private bool _keepIntermediateFiles;

        internal OutlookPstConverterDirect(IFileParser fileParser, IConverter converter, ConvertTarget convertTarget, bool keepIntermediateFiles = false)
        {
            _outlookPstParser = fileParser;
            _converter = converter;
            _convertTarget = convertTarget;
            _keepIntermediateFiles = keepIntermediateFiles;
        }

        public void ProcessFile(string inFile, string outDir)
        {
            bool result = _outlookPstParser.Parse(inFile);

            if (result)
            {
                foreach (var message in _outlookPstParser.ParsedContent)
                {
                    string nwPath = outDir + System.IO.Path.DirectorySeparatorChar.ToString() + message.FolderName;
                    DirectoryInfo nwDirInfo = new DirectoryInfo(nwPath);
                    if (!nwDirInfo.Exists)
                    {
                        Directory.CreateDirectory(nwDirInfo.FullName);
                    }

                    string filePart = message.IdentifyingName;
                    foreach (char invalidChar in System.IO.Path.GetInvalidFileNameChars())
                    {
                        filePart = filePart.Replace(invalidChar, '_');
                    }
                    string nwFilename = $"{nwPath}{System.IO.Path.DirectorySeparatorChar}{filePart}";

                    string nwFilenameTxt = nwFilename + ".txt";

                    File.WriteAllText(nwFilenameTxt, message.ContentAsString);

                    if (_convertTarget == ConvertTarget.PDF)
                    {
                        string nwFilenamePdf = nwFilename + ".pdf";
                        _converter.ConvertFromCom(nwFilenameTxt, nwFilenamePdf);
                    }

                    // Remove intermediate txt files if target is not text -->
                    if (!_keepIntermediateFiles && _convertTarget != ConvertTarget.TXT)
                    {
                        File.Delete(nwFilenameTxt);
                    }
                }
            }

        }
    }

    public class OutlookPstConverterMsg : IOutlookPstHandler
    {
        private IFileParser _outlookMsgParser;
        private IConverter _converter;
        private ConvertTarget _convertTarget;
        private bool _keepIntermediateFiles;

        internal OutlookPstConverterMsg(IFileParser fileParser, IConverter converter, ConvertTarget convertTarget, bool keepIntermediateFiles=false)
        {
            _outlookMsgParser = fileParser;
            _converter = converter;
            _convertTarget = convertTarget;
            _keepIntermediateFiles = keepIntermediateFiles;
        }

        public void ProcessFile(string inFile, string outDir)
        {
            processSingleOutlookPstFileUsingMsgFiles(inFile, outDir);
        }

        private void processSingleOutlookPstFileUsingMsgFiles(string pstFilePath, string destPath)
        {
            Application app = new Application();
            app.Session.AddStore(pstFilePath);

            Folders folders = app.GetNamespace("MAPI").Folders;

            foreach (MAPIFolder folder in folders)
            {
                Console.WriteLine(folder.Name);
                Items items = folder.Items;
                foreach (object item in items)
                {
                    //TODO: Is there ever a mailitem here? -->
                    if (item is MailItem mailItem)
                    {
                        WriteOutputFiles(mailItem, destPath);
                    }
                    else if (item is Folder subFolder)
                    {
                        string subPath = (destPath.EndsWith(Util.PathSeparator) ? destPath : destPath + Util.PathSeparator) + subFolder.Name + Util.PathSeparator;
                        WalkPstFolderAndSaveAsMsg(subFolder, subPath);
                    }
                }
                foreach (Folder subFolder in folder.Folders)
                {
                    string subPath = (destPath.EndsWith(Util.PathSeparator) ? destPath : destPath + Util.PathSeparator) + subFolder.Name + Util.PathSeparator;
                    WalkPstFolderAndSaveAsMsg(subFolder, subPath);
                }
            }
        }

        private void WalkPstFolderAndSaveAsMsg(Folder folder, string destPath)
        {
            Items items = folder.Items;
            foreach (object item in items)
            {
                if (item is MailItem mailItem)
                {
                    WriteOutputFiles(mailItem, destPath);
                }
                else if (item is Folder subFolder)
                {
                    string subPath = (destPath.EndsWith(Util.PathSeparator) ? destPath : destPath + Util.PathSeparator) + subFolder.Name + Util.PathSeparator;
                    // Recursively walk directory structure -->
                    WalkPstFolderAndSaveAsMsg(subFolder, subPath);
                }
            }
        }

        private void WriteOutputFiles(MailItem mailItem, string destPath)
        {
            var outputFileName = $"{mailItem.SenderEmailAddress} - {mailItem.Subject}.msg";
            // Save the message to disk in MSG format
            // File name may contain invalid characters [\ / : * ? " < > |]
            foreach (char invalidChar in System.IO.Path.GetInvalidFileNameChars())
            {
                outputFileName = outputFileName.Replace(invalidChar, '_');
            }

            Directory.CreateDirectory(destPath);
            string destinationFilenameMsg = destPath + outputFileName;
            string destinationFilenameTxt = destinationFilenameMsg + ".txt";
            string destinationFilenamePdf = destinationFilenameMsg + ".pdf";

            mailItem.SaveAs(destinationFilenameMsg, OlSaveAsType.olMSG);

            if (_outlookMsgParser.Parse(destinationFilenameMsg))
            {
                IParsedContent parsedContent = _outlookMsgParser.ParsedContent.Single<IParsedContent>();

                File.WriteAllText(destinationFilenameTxt, parsedContent.ContentAsString);
                _converter.ConvertFromCom(destinationFilenameTxt, destinationFilenamePdf);

                if (!_keepIntermediateFiles)
                {
                    File.Delete(destinationFilenameTxt);
                    File.Delete(destinationFilenameMsg);
                }
            }

        }
    }

    public class OutlookPstConverter : FileConverter
    {

        public enum PstConvertMethod
        {
            DIRECT,
            INTERMEDIATE_MSG,
        }

        public OutlookPstConverter(IFileParser fileParser, IOutputSupplier outputSupplier) : base(fileParser, outputSupplier)
        {
        }

        public PstConvertMethod ConvertMethod { get; set; } = PstConvertMethod.INTERMEDIATE_MSG;

        public override async Task ProcessInBackgroundAsync(ConvertConfig config)
        {
            KeepIntermediateFiles = config.KeepIntermediateFiles;

            if (!Directory.Exists(config.OutputDir))
            {
                Directory.CreateDirectory(config.OutputDir);
            }

            await processOutlookPstFiles(config.RootDir, config.OutputDir, Util.FileExtensions[FileType.OUTLOOK_PST]);
        }


        private async Task processOutlookPstFiles(string rootPath, string outputPath, string[] extensions)
        {
            // TODO: refactor so all background work is done in one thread instead of multiple async methods-->
            await Task.Run(async () =>
            {
                FileAttributes attr = File.GetAttributes(rootPath);

                if ((attr & FileAttributes.Directory) == FileAttributes.Directory)
                {
                    if (!outputPath.EndsWith(System.IO.Path.DirectorySeparatorChar.ToString()))
                    {
                        outputPath += System.IO.Path.DirectorySeparatorChar;
                    }

                    List<string> matchingFiles = FileWalker.WalkDir(rootPath, extensions, true);

                    foreach (var filename in matchingFiles)
                    {
                        Output.AddJournalEntry($"Found matching file: {filename}");

                        string nwFilename = filename.Replace(rootPath, outputPath);

                        processSingleOutlookPstFile(filename, nwFilename);
                    }
                }
                else
                {
                    Output.AddJournalEntry($"Processing single file: {rootPath}");
                    string nwFilename = rootPath.Replace(rootPath, outputPath);
                    processSingleOutlookPstFile(rootPath, nwFilename);
                }
            });
        }

        private void processSingleOutlookPstFile(string filename, string nwDirName)
        {
            switch(ConvertMethod)
            {
                case PstConvertMethod.DIRECT:
                    new OutlookPstConverterDirect(new OutlookPstParser(), Converter, TargetFormat, KeepIntermediateFiles)
                        .ProcessFile(filename, nwDirName);
                    break;
                case PstConvertMethod.INTERMEDIATE_MSG:
                    new OutlookPstConverterMsg(new OutlookMsgParser(), Converter, TargetFormat, KeepIntermediateFiles)
                        .ProcessFile(filename, nwDirName);
                    break;
                default:
                    throw new NotImplementedException("Nog niet");
            }
        }
    }
}
