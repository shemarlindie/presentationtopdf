using Microsoft.Win32;
using System;
using System.ComponentModel;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.AccessControl;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Animation;
using System.Windows.Shell;

using PP = Microsoft.Office.Interop.PowerPoint;
using Word = Microsoft.Office.Interop.Word;
using ppToPdf = PresentationToPDF.Properties;

/*
 * TODO
 * option for multithreaded conversion
 * remove file button
 * ability stop conversion while in progress
 * check for duplicates when adding files
 * Add files from directory (with subfolder option)
 */

namespace PresentationToPDF {
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window {
        public MainWindow() {
            InitializeComponent();

            dialogAddFiles = new OpenFileDialog();
            dialogAddFiles.Filter = ppToPdf.Resources.AddFilesFilter;
            dialogAddFiles.Title = ppToPdf.Resources.AddFilesTitle;
            dialogAddFiles.Multiselect = true;
            dialogAddFiles.FileOk += dialogAddFiles_FileOk;

            firstAdd = true;
            conversionRunning = false;
            cancelConversion = null;

            // log global unhandled exceptions
            Dispatcher.UnhandledException += (s, e) => {
                Logging.Logger.LogAsync(e.Exception);
                e.Handled = true;

                var mRes = MessageBox.Show("A fatal error has occured. The application will try to restart... Continue?",
                    "Closing application", MessageBoxButton.YesNo);

                if (mRes == MessageBoxResult.Yes) {
                    try {
                        System.Diagnostics.Process.Start(Assembly.GetExecutingAssembly().Location);
                    }
                    catch { }
                }
                Application.Current.Shutdown();
            };
        }

        private readonly OpenFileDialog dialogAddFiles;
        private bool firstAdd;
        private bool conversionRunning;
        private CancellationTokenSource cancelConversion;



        private void btnAddFiles_Click(object sender, RoutedEventArgs e) {
            dialogAddFiles.ShowDialog(this);
        }

        private async void dialogAddFiles_FileOk(object sender, CancelEventArgs e) {
            ResetProgress(); // reset the progress bar
            UIEnabled(false);

            try {
                // add files to listbox
                await AddFilesAsync(dialogAddFiles.FileNames);
            }
            catch (PathTooLongException ex) {
                MessageBox.Show(string.Format(ppToPdf.Resources.MsgDirCreationFailed_Path, ex.Data["Path"]), "Path too long",
                    MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK);
                Console.WriteLine(ex.Message);

                return;
            }
            catch (UnauthorizedAccessException ex) {
                MessageBox.Show(string.Format(ppToPdf.Resources.MsgDirCreationFailed_Access, ex.Data["Path"]), "Access is denied",
                    MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK);
                Console.WriteLine(ex.Message);

                return;
            }
            catch (Exception ex) {
                Logging.Logger.LogAsync(ex);
                MessageBox.Show(ppToPdf.Resources.UnknownError, "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK);

                return;
            }
            finally {
                UIEnabled(true);
            }
        }

        private async void btnConvert_Click(object sender, RoutedEventArgs e) {
            if (listboxFiles.Items.Count < 1 || firstAdd) { // is the listbox empty
                return;
            }

            if (!conversionRunning) {
                cancelConversion = new CancellationTokenSource();
                conversionRunning = true;
                btnCancel.IsEnabled = true;

                try {
                    await startConversionAsync(new Progress<ProgressInfo>(UpdateProgress), cancelConversion.Token);
                }
                catch (OperationCanceledException) {
                    ResetProgress();
                    lblProgress.Content = "Conversion canceled.";
                }
                finally {
                    // dispose CancellationTokenSource
                    cancelConversion.Dispose();
                    cancelConversion = null;

                    btnCancel.IsEnabled = false;
                    UIEnabled(true);
                    conversionRunning = false;

                    TaskbarItemInfo.ProgressState = TaskbarItemProgressState.None;
                }
            }
        }



        private void btnBrowse_Click(object sender, RoutedEventArgs e) {
            // bring up folder selection dialog
            var dialogOutput = new System.Windows.Forms.FolderBrowserDialog();
            dialogOutput.Description = "Select a folder to save converted files...";
            var res = dialogOutput.ShowDialog();

            // get selected folder from dialog
            if (res == System.Windows.Forms.DialogResult.OK) {
                txtDestDir.Text = dialogOutput.SelectedPath;
            }
        }

        private void btnClear_Click(object sender, RoutedEventArgs e) {
            if (listboxFiles.Items.Count > 0) {
                // remove all files and reset progress bar
                ResetProgress();
                listboxFiles.Items.Clear();
                Util.CleanupUnusedMemoryAsync();
            }
        }

        private void listboxFiles_DragEnter(object sender, DragEventArgs e) {
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) {
                e.Effects = DragDropEffects.Move;
            }
            else {
                e.Effects = DragDropEffects.None;
            }
        }

        private async void listboxFiles_Drop(object sender, DragEventArgs e) {

            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            string outDir = txtDestDir.Text;

            UIEnabled(false);

            try {
                await AddFilesAsync(files);
            }
            catch (PathTooLongException ex) {
                MessageBox.Show(string.Format(ppToPdf.Resources.MsgDirCreationFailed_Path, ex.Data["Path"]), "Path too long",
                    MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK);

                return;
            }
            catch (UnauthorizedAccessException ex) {
                MessageBox.Show(string.Format(ppToPdf.Resources.MsgDirCreationFailed_Access, ex.Data["Path"]), "Access is denied",
                    MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK);

                return;
            }
            catch (Exception ex) {
                Logging.Logger.LogAsync(ex);
                MessageBox.Show(ppToPdf.Resources.UnknownError, "Something went wrong",
                    MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK);

                return;
            }
            finally {
                UIEnabled(true);
            }
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e) {
            if (conversionRunning) {
                TaskbarItemInfo.ProgressState = TaskbarItemProgressState.Indeterminate;
                lblProgress.Content = "Cancelling...";
                cancelConversion.Cancel();
                btnCancel.IsEnabled = false;
            }
        }




        /// <summary>
        /// Starts converting the presentations in the onversion queue
        /// </summary>
        /// <param name="cancel">Cancellation token used to stop the conversion process.</param>
        /// <returns></returns>
        private async Task startConversionAsync(IProgress<ProgressInfo> onProgressChanged, CancellationToken cancel) {
            UIEnabled(false);

            TaskbarItemInfo.ProgressState = TaskbarItemProgressState.Indeterminate;
            lblProgress.Content = "Starting conversion...";

            // set progress bar limits
            progressConversion.Minimum = 0;
            progressConversion.Maximum = listboxFiles.Items.Count;

            string destDir = txtDestDir.Text; // output folder
            int index = 0;  // current progress index
            PP.Application ppApp = null;
            Word.Application wordApp = null;

            try {
                await Task.Run(() => {
                    // create the output directory if it does not exist
                    if (!Directory.Exists(destDir)) {
                        // set file system access rules
                        FileSystemAccessRule rule = new FileSystemAccessRule(
                                                            "Everyone",
                                                            FileSystemRights.FullControl,
                                                            InheritanceFlags.ContainerInherit |
                                                            InheritanceFlags.ObjectInherit,
                                                            PropagationFlags.None,
                                                            AccessControlType.Allow);

                        bool success;
                        var sec = new DirectorySecurity();
                        sec.ModifyAccessRule(AccessControlModification.Add, rule, out success);

                        Directory.CreateDirectory(destDir, sec); // create directory                    
                    }
                });
            }
            catch (PathTooLongException) {
                MessageBox.Show(string.Format(ppToPdf.Resources.MsgDirCreationFailed_Path, destDir), "Path too long",
                    MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK);
                lblProgress.Content = "Something went wrong.";
                UIEnabled(true);

                return;
            }
            catch (IOException) {
                MessageBox.Show(string.Format(ppToPdf.Resources.MsgDirCreationFailed_IO, destDir), "Cannot create folder",
                    MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK);
                lblProgress.Content = "Something went wrong.";
                UIEnabled(true);

                return;
            }
            catch (UnauthorizedAccessException) {
                MessageBox.Show(string.Format(ppToPdf.Resources.MsgDirCreationFailed_Access, destDir), "Access is denied",
                    MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK);
                lblProgress.Content = "Something went wrong.";
                UIEnabled(true);

                return;
            }
            catch (NotSupportedException) {
                MessageBox.Show(string.Format(ppToPdf.Resources.MsgDirCreationFailed_Invalid, destDir), "Invalid output path",
                    MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK);
                lblProgress.Content = "Something went wrong.";
                UIEnabled(true);

                return;
            }
            catch (Exception ex) {
                Logging.Logger.LogAsync(ex);
                MessageBox.Show(ppToPdf.Resources.UnknownError, "Something went wrong",
                    MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK);
                lblProgress.Content = "Something went wrong.";
                UIEnabled(true);

                return;
            }
            finally {

            }

            // start conversion
            foreach (OfficeFileInfo file in listboxFiles.Items) {
                // update progress
                onProgressChanged.Report(new ProgressInfo(++index, listboxFiles.Items.Count, Path.GetFileName(file.Path)));

                try {
                    if (file is PptInfo) {
                        if (ppApp == null) {
                            ppApp = await StartPowerPointAsync();
                        }

                        await ppConvertAsync(ppApp, file as PptInfo, destDir);
                    }
                    else if (file is DocInfo) {
                        if (wordApp == null) {
                            wordApp = await StartWordAsync();
                        }

                        await wordConvertAsync(wordApp, file as DocInfo, destDir);
                    }
                }
                // seems to indicate conversion failure but conversion doesn't really fail, so ignore for now
                catch (COMException) { }
                catch (Exception ex) {
                    Logging.Logger.LogAsync(ex);
                    MessageBox.Show(ppToPdf.Resources.UnknownError, "Conversion error",
                        MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK);
                    UIEnabled(true);

                    return;
                }

                try {
                    cancel.ThrowIfCancellationRequested();
                }
                catch (OperationCanceledException) {
                    if (ppApp != null) {
                        ppApp.Quit();
                        Util.ReleaseComObject(ppApp);
                    }

                    if (wordApp != null) {
                        ((Word._Application)wordApp).Quit(Word.WdSaveOptions.wdDoNotSaveChanges);
                        Util.ReleaseComObject(wordApp);
                    }

                    throw;
                }
            }

            if (ppApp != null) {
                ppApp.Quit();
                Util.ReleaseComObject(ppApp);
            }

            if (wordApp != null) {
                ((Word._Application)wordApp).Quit(Word.WdSaveOptions.wdDoNotSaveChanges);
                Util.ReleaseComObject(wordApp);
            }

            lblProgress.Content = "Done!";
        }

        /// <summary>
        /// Converts ppts to PDF documents asynchronously
        /// </summary>
        /// <param name="ppApp">The PowerPoint application</param>
        /// <param name="info">A PptInfo object of the file to be converted to pdf</param>
        /// <param name="newPath">Full path of the converted file</param>
        private Task ppConvertAsync(PP.Application ppApp, PptInfo info, string destDir) {
            return Task.Run(() => {
                // open presentation in PP in the bg
                PP.Presentation pres = ppApp.Presentations.Open(info.Path, WithWindow: Microsoft.Office.Core.MsoTriState.msoFalse);

                // build path for converted file
                string newPath = Path.Combine(new[] { destDir, Path.ChangeExtension(Path.GetFileName(info.Path), "pdf") });

                // convert to pdf
                pres.SaveAs(newPath, PP.PpSaveAsFileType.ppSaveAsPDF);
                pres.Close();
                Util.ReleaseComObject(pres);
            });
        }

        /// <summary>
        /// Converts docs to PDF documents asynchronously
        /// </summary>
        /// <param name="ppApp">The Word application</param>
        /// <param name="info">A DocInfo object of the file to be converted to pdf</param>
        /// <param name="newPath">Full path of the converted file</param>
        private Task wordConvertAsync(Word.Application wordApp, DocInfo info, string destDir) {
            return Task.Run(() => {
                // open doc in Word in the bg
                Word.Document doc = wordApp.Documents.Open(info.Path, AddToRecentFiles: false, Visible: false);

                // build path for converted file
                string newPath = Path.Combine(new[] { destDir, Path.ChangeExtension(Path.GetFileName(info.Path), "pdf") });

                // save as pdf
                doc.SaveAs2(newPath, Word.WdSaveFormat.wdFormatPDF);
                ((Word._Document)doc).Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                Util.ReleaseComObject(doc);
            });
        }

        private Task<PP.Application> StartPowerPointAsync() {
            return Task.Run<PP.Application>(() => new PP.Application());
        }

        private Task<Word.Application> StartWordAsync() {
            return Task.Run<Word.Application>(() => new Word.Application());
        }

        /// <summary>
        /// Add PowerPoint presentations to the conversion queue
        /// </summary>
        /// <param name="files">Path to files to be added</param>
        /// <returns></returns>
        private async Task AddFilesAsync(string[] files) {
            if (firstAdd) { // remove "drag & drop" placeholder
                listboxFiles.Items.Clear();
                firstAdd = false;
            }

            // add presentations
            foreach (string s in files) {
                OfficeFileInfo file = null;
                await Task.Run(() => {
                    try {
                        if (s.Contains(".ppt")) {
                            file = new PptInfo(s);
                        }
                        else if (s.Contains(".doc")) {
                            file = new DocInfo(s);
                        }
                    }
                    catch (PathTooLongException ex) {
                        ex.Data.Add("Path", s);
                        throw;
                    }
                    catch (UnauthorizedAccessException ex) {
                        ex.Data.Add("Path", s);
                        throw;
                    }
                });

                if (file != null) {
                    listboxFiles.Items.Add(file);
                }
            }

            // create path to default output folder
            if (listboxFiles.Items.Count > 0) { // were files added?
                // if not use default
                string path = Path.GetDirectoryName((listboxFiles.Items[0] as OfficeFileInfo).Path);
                string dest = Path.Combine(new[] { path, "PDF-Conversions" });
                txtDestDir.Text = dest;

                // show listbox file count
                lblProgress.Content = string.Format("{0} documents(s).", listboxFiles.Items.Count);
            }
        }

        /// <summary>
        /// Updates progress bar and text
        /// </summary>
        /// <param name="currentIndex">Progress index of current file</param>
        /// <param name="fileName">Name of current file</param>
        private void UpdateProgress(ProgressInfo prog) {
            TaskbarItemInfo.ProgressState = TaskbarItemProgressState.Normal;
            TaskbarItemInfo.ProgressValue = (double)prog.CurrentIndex / prog.MaxIndex;

            // create animation params
            Duration dur = new Duration(TimeSpan.FromSeconds(0.5));
            DoubleAnimation ani = new DoubleAnimation(prog.CurrentIndex, dur);

            // animate to new progress
            progressConversion.BeginAnimation(ProgressBar.ValueProperty, ani);

            // update progress text
            lblProgress.Content = string.Format("Converting... {0} / {1} : {2}",
                prog.CurrentIndex, prog.MaxIndex, prog.FileName);
        }

        /// <summary>
        /// Reset progress bar to it's "no progress" state
        /// </summary>
        private void ResetProgress() {
            // create animation params
            Duration dur = new Duration(TimeSpan.FromSeconds(1));
            DoubleAnimation ani = new DoubleAnimation(0, dur);

            // animate to reset state
            progressConversion.BeginAnimation(ProgressBar.ValueProperty, ani);

            lblProgress.Content = "-"; // reset progress text
        }

        /// <summary>
        /// Enables or disable primary UI controls. 
        /// Useful when performing certain long operations.
        /// </summary>
        /// <param name="state">Enable: true; Disable: false</param>
        private void UIEnabled(bool state) {
            //listboxFiles.IsEnabled = state;
            btnAddFiles.IsEnabled = state;
            btnBrowse.IsEnabled = state;
            btnClear.IsEnabled = state;
            btnConvert.IsEnabled = state;
            txtDestDir.IsEnabled = state;
        }

        private void listboxFiles_KeyDown(object sender, System.Windows.Input.KeyEventArgs e) {
            //switch (e.Key) {
            //    case Key.Delete:
            //        if (!firstAdd) {
            //            listboxFiles.Items.Remove(listboxFiles.SelectedItem);
            //            //object[] toRemove = new object[listboxFiles.SelectedItems.Count];
            //            //foreach (object o in toRemove) {
            //            //    listboxFiles.Items.Remove(o);
            //            //}
            //        }
            //        break;
            //}
        }

    }
}
