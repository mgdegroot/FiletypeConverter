using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging.Console;
using Microsoft.Extensions.Logging.Debug;
using MsgReader.Outlook;
using OfficeConverter;
using MessageBox = System.Windows.MessageBox;

namespace FiletypeConverter
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public static readonly ILogger _logger;
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private Converter converter = new Converter();

        private readonly SynchronizationContext synchronizationContext;

        public MainWindow()
        {
            InitializeComponent();
            synchronizationContext = SynchronizationContext.Current;

        }

        private void btnTest_Click(object sender, RoutedEventArgs e)
        {


            startInBg();


            //            Tester.TestWord();
            //            Tester.TestPowerpoint();
            return;

////            OutlookMsgParser outlookMsgParser = new OutlookMsgParser();
//            var path = txtPath.Text;
////            if (outlookMsgParser.parse(path))
////            {
////                txtDebug.Text = outlookMsgParser.MsgAsString;
////            }
//
//            if (chkOutlookMsg.IsChecked ?? false)
//            {
//                string msgAsText = string.Empty;
//
//                using (var msg = new MsgReader.Outlook.Storage.Message(path))
//                {
//                    var from = msg.Sender.DisplayName + "<" + msg.Sender.Email + ">";
//                    var sentOn = msg.SentOn;
//                    var recipientsTo = msg.GetEmailRecipients(Storage.Recipient.RecipientType.To, false, false);
//                    var recipientsCC = msg.GetEmailRecipients(Storage.Recipient.RecipientType.Cc, false, false);
//                    var subject = msg.Subject;
//                    var htmlBody = msg.BodyHtml;
//                    var rtfBody = msg.BodyRtf;
//                    var textBody = msg.BodyText;
//                    var attachements = msg.Attachments;
//                    var attachementNames = msg.GetAttachmentNames();
//                    var creationTime = msg.CreationTime;
//                    var headers = msg.Headers;
//
//                    var receivedOn = msg.ReceivedOn ?? DateTime.MinValue;
//
//
//                    var lastModificationTime = msg.LastModificationTime;
//
//                    msgAsText = $@"
//FROM: {from}
//SENT ON: {sentOn}
//TO: {recipientsTo}
//CC: {recipientsCC}
//SUBJECT: {subject}
//HTMLBODY: {htmlBody}
//RTFBODY: {rtfBody}
//TXTBODY: {textBody}
//ATTN: {attachementNames}
//CREATIONTIME: {creationTime}
//RECV_ON: {receivedOn}
//MOD_DATE: {lastModificationTime}";
//                }
//
//                txtDebug.Text = msgAsText;
//            }
//            else if (chkWord.IsChecked ?? false)
//            {
//                WordDocumentParser wdp = new WordDocumentParser(path);
//                bool result = wdp.parse();
//                if (result)
//                {
//                    string asTxt = wdp.DocAsString;
//                    txtDebug.Text = asTxt;
//                }
//
//            }

        }


        private void writeToJournalFile(string text)
        {
            FileInfo fi = new FileInfo(txtJournalFilename.Text);
            if (!fi.Exists)
            {
                Directory.CreateDirectory(fi.Directory.FullName);
            }

            File.AppendAllText(txtJournalFilename.Text, text);
        }

        struct ConvertConfig
        {
            public bool ProcessOutlookMsg { get; set; }
            public bool ProcessWord { get; set; }
            public bool ProcessPowerpoint { get; set; }
            public bool ProcessExcel { get; set; }
            public string RootDir { get; set; }
            public string OutputDir { get; set; }
            public string Filter { get; set; }
        }

        private void btnWalkdir_Click(object sender, RoutedEventArgs e)
        {
            startInBg();
        }


        private TaskScheduler scheduler = null;

        private async Task startInBg()
        {

            var outputDir = txtOutputRootDir.Text;

            if (!outputDir.EndsWith(System.IO.Path.DirectorySeparatorChar.ToString()))
            {
                outputDir += System.IO.Path.DirectorySeparatorChar;
            }

            ConvertConfig convertConfig = new ConvertConfig()
            {
                ProcessOutlookMsg = chkOutlookMsg.IsChecked ?? false,
                ProcessWord = chkWord.IsChecked ?? false,
                ProcessPowerpoint = chkPowerpoint.IsChecked ?? false,
                ProcessExcel = chkExcel.IsChecked ?? false,
                RootDir = txtRootDir.Text,
                OutputDir = outputDir,
                Filter = txtWalkdirFilter.Text
            };

            Task<int> result = processInBackgroundAsync(convertConfig);
            
        }

        async Task<int> processInBackgroundAsync(ConvertConfig config)
        {
            lblIsRunning.Content = "RUNNING";
            lblIsRunning.Background = Brushes.DarkRed;
            lblIsRunning.Foreground = Brushes.LightSeaGreen;
            txtDebug.Text += "\r\nstarted...";
//            Test async -->
            await Task.Delay(10);

            if (!Directory.Exists(config.OutputDir))
            {
                Directory.CreateDirectory(config.OutputDir);
            }

            if (config.ProcessOutlookMsg)
            {
                processMsgFiles(config.RootDir, config.OutputDir);
            }

            if (config.ProcessWord)
            {
                processWordFiles(config.RootDir, config.OutputDir);
            }

            if (config.ProcessPowerpoint)
            {
                processPowerpointFiles(config.RootDir, config.OutputDir);
            }

            if (config.ProcessExcel)
            {
                processExcelFiles(config.RootDir, config.OutputDir);
            }



            return 0;
        }

        private async Task processOfficeFiles(string rootDir, string outputDir, string extension)
        {
            await Task.Run(async () =>
            {
                List<string> matchingFiles = FileWalker.WalkDir(rootDir, extension, true);

                foreach (var filename in matchingFiles)
                {
                    logAndUpdateUI($"Found matching file: {filename}", string.Empty);
                    string nwFilename = filename.Replace(rootDir, outputDir) + ".pdf";


                    logAndUpdateUI($"ORIGINAL: {filename}\tNEW: {nwFilename}", string.Empty);


                    FileInfo nwFileInfo = new FileInfo(nwFilename);
                    if (!nwFileInfo.Exists)
                    {
                        Directory.CreateDirectory(nwFileInfo.Directory.FullName);
                    }

                    log.Info($"Converting {filename} to {nwFilename}");
                    try
                    {
                        converter.Convert(filename, nwFilename);
                    }
                    catch (Exception ex)
                    {
                        string errMsg = $"ERROR: {filename}: {ex.Message}";
                        logAndUpdateUI(string.Empty, errMsg, true);
                    }
                }
            });
            lblIsRunning.Content = "IDLE";
            lblIsRunning.Background = Brushes.DimGray;
            lblIsRunning.Foreground = Brushes.GreenYellow;
            txtDebug.Text += "\r\ndone...";
        }

        private void logAndUpdateUI(string outputText, string dbgText, bool isError = false)
        {
            if (!string.IsNullOrEmpty(outputText))
            {
                synchronizationContext.Post(new SendOrPostCallback(o => { txtOutput.Text += (string)o + Environment.NewLine; txtOutput.ScrollToEnd();}),outputText);
            }

            if (!string.IsNullOrEmpty(dbgText))
            {
                
                if (isError)
                {
                    log.Error(dbgText);
                }
                else
                {
                    log.Info(dbgText);
                }

                synchronizationContext.Post(new SendOrPostCallback(o => { txtDebug.Text += (string)o + Environment.NewLine; txtDebug.ScrollToEnd();}), dbgText);
            }
            
        }

        private async void processPowerpointFiles(string rootDir, string outputDir)
        {
            log.Info("Converting Powerpoint documents.");

            await processOfficeFiles(rootDir, outputDir, "*.ppt?");

        }

        private async void processWordFiles(string rootDir, string outputDir)
        {
            log.Info("Converting Word documents." );

            await processOfficeFiles(rootDir, outputDir, "*.doc?");

        }

        private async void processExcelFiles(string rootDir, string outputDir)
        {
            await processOfficeFiles(rootDir, outputDir, "*.xls?");
        }

        private void processMsgFiles(string rootDir, string outputDir)
        {
            List<string> matchingFiles = FileWalker.WalkDir(rootDir, "*.msg", true);
            foreach (var filename in matchingFiles)
            {
                string nwFilename = filename.Replace(rootDir, outputDir);

                txtOutput.Text = txtOutput.Text + Environment.NewLine + "ORIG: " + filename + "\t" + "NEW: " +
                                 nwFilename + Environment.NewLine;

                FileInfo nwFileInfo = new FileInfo(nwFilename);
                if (!nwFileInfo.Exists)
                {
                    Directory.CreateDirectory(nwFileInfo.Directory.FullName);
                }
                //                    File.Create(nwFilename);
                if (chkOutputTxt.IsChecked ?? false)
                {
                    nwFilename += ".txt";
                    var parser = new OutlookMsgParser(filename);
                    if (parser.parse())
                    {
                        var result = parser.MsgAsString;
                        File.WriteAllText(nwFilename, result);
                    }
                    else
                    {
                        txtDebug.Text = txtDebug.Text + Environment.NewLine + "ERROR: " + filename;
                    }
                }
            }
        }



        private void txtRootDir_OnMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            txtRootDir.Text = showDirectoryPicker(txtRootDir.Text);
        }

        private void txtOutputRootDir_OnMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            txtOutputRootDir.Text = showDirectoryPicker(txtOutputRootDir.Text);
            if (chkJournalInOutputRootDir.IsChecked ?? false)
            {
                txtJournalFilename.Text = txtOutputRootDir.Text;
            }
        }

        private void txtJournalFilename_OnMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (!chkJournalInOutputRootDir.IsChecked ?? false)
            {
                txtJournalFilename.Text = showDirectoryPicker(txtJournalFilename.Text);
            }
        }

        private string showDirectoryPicker(string startLoc)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.SelectedPath = startLoc;
            fbd.ShowNewFolderButton = true;

            if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                return fbd.SelectedPath;
            }
            else
            {
                return startLoc;
            }
        }

        private void btnDummy_OnClick(object sender, RoutedEventArgs e)
        {
            txtDebug.Text += "\r\n\r\nDUMMY DUUMMY DUMMY\r\n\r\n";
            txtDebug.ScrollToEnd();
        }

        private void chkJournalInOutputRootDir_OnChecked(object sender, RoutedEventArgs e)
        {
            txtJournalFilename.Text = txtOutputRootDir.Text;
            txtJournalFilename.IsReadOnly = true;
        }

        private void chkJournalInOutputRootDir_OnUnchecked(object sender, RoutedEventArgs e)
        {
            txtJournalFilename.IsReadOnly = false;
        }

        private void txtOutputRootDir_OnTextInput(object sender, RoutedEventArgs routedEventArgs)
        {
            if (chkJournalInOutputRootDir.IsChecked ?? false)
            {
                txtJournalFilename.Text = txtOutputRootDir.Text;
            }
        }
    }
}
