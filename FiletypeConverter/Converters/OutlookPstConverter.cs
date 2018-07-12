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

namespace FiletypeConverter
{
    public class OutlookPstConverter : FileConverter, IFileConverter
    {
        // TODO: interface based -->
        private OutlookPstParser outlookPstParser;

        public OutlookPstConverter()
        {
            outlookPstParser = new OutlookPstParser();
        }

        public override async Task processInBackgroundAsync(ConvertConfig config)
        {
            KeepIntermediateFiles = config.KeepIntermediateFiles;

            if (!Directory.Exists(config.OutputDir))
            {
                Directory.CreateDirectory(config.OutputDir);
            }

            if (config.ProcessOutlookPst)
            {

                await processOutlookPstFiles(config.RootDir, config.OutputDir, "*.pst?");
            }
        }


        private async Task processOutlookPstFiles(string rootPath, string outputPath, string extension)
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

                    List<string> matchingFiles = FileWalker.WalkDir(rootPath, extension, true);

                    foreach (var filename in matchingFiles)
                    {
                        Output.AddJournalEntry($"Found matching file: {filename}");

                        string nwFilename = filename.Replace(rootPath, outputPath);
                        /*
                         * instantiate parser, parse file, convert resulting messages
                         * 
                         * nwPath = nwPath + System.IO.Path.DirectorySeparatorChar.ToString() + folderName;
                         */

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
            bool result = outlookPstParser.Parse(filename);

            if (result)
            {
                foreach(var message in outlookPstParser.ParsedMessages)
                {
                    string nwPath = nwDirName + System.IO.Path.DirectorySeparatorChar.ToString() + message.FolderName;
                    DirectoryInfo nwDirInfo = new DirectoryInfo(nwPath);
                    if (!nwDirInfo.Exists)
                    {
                        Directory.CreateDirectory(nwDirInfo.FullName);
                    }

                    string filePart = $"{message.Sender} - {message.Subject}";
                    foreach(char invalidChar in System.IO.Path.GetInvalidFileNameChars())
                    {
                        filePart = filePart.Replace(invalidChar, '_');
                    }
                    string nwFilename = $"{nwPath}{System.IO.Path.DirectorySeparatorChar}{filePart}";

                    string nwFilenameTxt = nwFilename + ".txt";
                    

                    File.WriteAllText(nwFilenameTxt, message.ContentAsString);

                    if (TargetFormat == ConvertTarget.PDF)
                    {
                        string nwFilenamePdf = nwFilename + ".pdf";
                        Converter.ConvertFromCom(nwFilenameTxt, nwFilenamePdf);
                    }

                    // Remove intermediate txt files if target is not text -->
                    if (!KeepIntermediateFiles && TargetFormat != ConvertTarget.TXT)
                    {
                        File.Delete(nwFilenameTxt);
                    }
                }
            }

        }



        private static int msgCounter = 0;

        private void writeToFile(string filename, string txt)
        {
            using (StreamWriter outputFile = new StreamWriter(filename, false))
            {
                string[] lines = txt.Split(new []{ "\r\n"}, StringSplitOptions.None);
                foreach (string line in lines)
                    outputFile.WriteLine(line);
            }
        }
    }
}
