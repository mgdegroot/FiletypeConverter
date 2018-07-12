using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using static FiletypeConverter.FileConverter;
using static FiletypeConverter.OutlookPstConverter;
using Timer = System.Threading.Timer;

namespace FiletypeConverter
{
    /// <summary>
    /// </summary>
    public partial class MainWindow : Window
    {
        private static log4net.ILog log;
        private readonly SynchronizationContext synchronizationContext;

        public MainWindow()
        {
            InitializeComponent();
            synchronizationContext = SynchronizationContext.Current;
        }

        private void setJournalFilename(string filename)
        {
            FileInfo fi = new FileInfo(filename);

            if (!fi.Exists)
            {
                Directory.CreateDirectory(fi.Directory.FullName);
            }

            log4net.GlobalContext.Properties["LogFilename"] = filename;
            log4net.Config.XmlConfigurator.Configure();
            log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
            log.Info("Journal started");
            //            IAppender[] appenders = log.Logger.Repository.GetAppenders();
            //
            //            foreach (IAppender appender in appenders)
            //            {
            //                if (appender is log4net.Appender.RollingFileAppender)
            //                {
            //                    RollingFileAppender fsAppender = (RollingFileAppender)appender;
            //                    fsAppender.File = filename;
            //                    log.Info("Journal started");
            //                }
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

        private void btnWalkdir_Click(object sender, RoutedEventArgs e)
        {
            setJournalFilename(txtJournalFilename.Text);
            startInBg();
        }


        private TaskScheduler scheduler = null;
        private Task backgroundConversionTask = null;

        private async Task startInBg()
        {
            var outputDir = txtOutputRootDir.Text;
            
            ConvertConfig convertConfig = new ConvertConfig()
            {
                ProcessOutlookMsg = chkOutlookMsg.IsChecked ?? false,
                ProcessWord = chkWord.IsChecked ?? false,
                ProcessPowerpoint = chkPowerpoint.IsChecked ?? false,
                ProcessExcel = chkExcel.IsChecked ?? false,
                ProcessImages = chkCopyImages.IsChecked ?? false,
                ProcessOutlookPst = chkOutlookPst.IsChecked ?? false,
                KeepIntermediateFiles = chkOutputTxt.IsChecked ?? false,
                RootDir = txtRootDir.Text,
                OutputDir = outputDir,
                Filter = txtWalkdirFilter.Text
            };

            await Task.Run(async () =>
            {
                if (!Directory.Exists(convertConfig.OutputDir))
                {
                    Directory.CreateDirectory(convertConfig.OutputDir);
                }


                if (convertConfig.ProcessOutlookMsg)
                {
                    FileConverter outlookFileConverter = new OutlookFileConverter();
                    outlookFileConverter.JournalAdded += journalEntryAdded;
                    outlookFileConverter.ErrorAdded += errorEntryAdded;

                    await outlookFileConverter.processInBackgroundAsync(convertConfig);
                }

                if (convertConfig.ProcessWord || convertConfig.ProcessPowerpoint || convertConfig.ProcessExcel)
                {
                    FileConverter officeFileConverter = new OfficeFileConverter();
                    officeFileConverter.JournalAdded += journalEntryAdded;
                    officeFileConverter.ErrorAdded += errorEntryAdded;

                    await officeFileConverter.processInBackgroundAsync(convertConfig);
                }

                if (convertConfig.ProcessImages)
                {
                    FileConverter fileTransferrer = new ImageFileConverter();
                    fileTransferrer.JournalAdded += journalEntryAdded;
                    fileTransferrer.ErrorAdded += errorEntryAdded;

                    await fileTransferrer.processInBackgroundAsync(convertConfig);
                }

                if (convertConfig.ProcessOutlookPst)
                {
                    OutlookPstConverter pstConverter = new OutlookPstConverter();
                    pstConverter.JournalAdded += journalEntryAdded;
                    pstConverter.ErrorAdded += errorEntryAdded;

                    await pstConverter.processInBackgroundAsync(convertConfig);
                }
            });
        }

        public void journalEntryAdded(string message)
        {
            logAndUpdateUI(message, null, false);
        }

        public void errorEntryAdded(string message)
        {
            logAndUpdateUI(null, message, true);
        }



        private void logAndUpdateUI(string outputText, string dbgText, bool isError = false)
        {
            if (!string.IsNullOrEmpty(outputText))
            {
                log.Info("journal: " + outputText);
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
                string journalPath = showDirectoryPicker(txtJournalFilename.Text);

                if (!journalPath.EndsWith(@"\"))
                {
                    journalPath += @"\";
                }

                journalPath += $"journal_{DateTime.Now.ToString("yyyy-MM-dd")}";

                txtJournalFilename.Text = journalPath;
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

        private string showFilePicker(string startLoc)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.InitialDirectory = startLoc;

            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                return ofd.FileName;
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

        private void btnPickFile_OnClick(object sender, RoutedEventArgs e)
        {
            txtRootDir.Text = showFilePicker(txtRootDir.Text);
        }

        private void chkJournalInOutputRootDir_OnChecked(object sender, RoutedEventArgs e)
        {
            string journalPath = txtOutputRootDir.Text;
            if (!journalPath.EndsWith(@"\"))
            {
                journalPath += @"\";
            }

            journalPath += $"journal_{DateTime.Now.ToString("yyyy-MM-dd")}";

            txtJournalFilename.Text = journalPath;
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
                string journalPath = txtOutputRootDir.Text;

                if (!journalPath.EndsWith(@"\"))
                {
                    journalPath += @"\";
                }

                journalPath += $"journal_{DateTime.Now.ToString("yyyy-MM-dd")}";

                txtJournalFilename.Text = journalPath;
            }
        }


        private void btnTest_Click(object sender, RoutedEventArgs e)
        {
            setJournalFilename(@"D:\test\test.log");
            string path = txtPath.Text;

            //Tester.TestTxtToPdf();
            OutlookPstConverter outlookPstConverter = new OutlookPstConverter()
            {
                Path = path,
            };

            //outlookPstConverter.parse(path);
            string res = string.Empty;
            //foreach(ParsedPstMessage parsedMessage in outlookPstConverter.ParsedMessages)
            //{
            //    res += parsedMessage.MsgAsString + "\r\n\r\n";
            //}

            txtDebug.Text = res;

        }


    }
}
