using Microsoft.Win32;
using System;
using System.ComponentModel;
using System.Windows;
using System.Windows.Media.Animation;
using System.Windows.Controls;
using System.IO;
using System.Threading;
using System.Security.AccessControl;
using System.Reflection;

using ppToPdf = PresentationToPDF.Properties;
using PP = Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Shell;
using System.Windows.Input;

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
                Logging.Logger.Log(e.Exception);
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
                Logging.Logger.Log(ex);
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
                catch(OperationCanceledException) {
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
                txtDestPath.Text = dialogOutput.SelectedPath;
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
            string outDir = txtDestPath.Text;

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
                Logging.Logger.Log(ex);
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
        private async Task startConversionAsync(IProgress<ProgressInfo> onProgressChanged,  CancellationToken cancel) {
            UIEnabled(false);
            
            TaskbarItemInfo.ProgressState = TaskbarItemProgressState.Indeterminate;
            lblProgress.Content = "Starting conversion...";

            // set progress bar limits
            progressConversion.Minimum = 0;
            progressConversion.Maximum = listboxFiles.Items.Count;

            string destPath = txtDestPath.Text; // output folder
            int index = 0;  // current progress index
            PP.Application ppApp = null;

            try {
                await Task.Run(() => {
                    ppApp = new PP.Application();

                    // create the output directory if it does not exist
                    if (!Directory.Exists(destPath)) {
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

                        Directory.CreateDirectory(destPath, sec); // create directory                    
                    }
                });
            }
            catch (PathTooLongException ex) {
                MessageBox.Show(string.Format(ppToPdf.Resources.MsgDirCreationFailed_Path, destPath), "Path too long",
                    MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK);
                lblProgress.Content = "Something went wrong.";
                UIEnabled(true);

                return;
            }
            catch (IOException ex) {
                MessageBox.Show(string.Format(ppToPdf.Resources.MsgDirCreationFailed_IO, destPath), "Cannot create folder",
                    MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK);
                lblProgress.Content = "Something went wrong.";
                UIEnabled(true);

                return;
            }
            catch (UnauthorizedAccessException ex) {
                MessageBox.Show(string.Format(ppToPdf.Resources.MsgDirCreationFailed_Access, destPath), "Access is denied",
                    MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK);
                lblProgress.Content = "Something went wrong.";
                UIEnabled(true);

                return;
            }
            catch (NotSupportedException ex) {
                MessageBox.Show(string.Format(ppToPdf.Resources.MsgDirCreationFailed_Invalid, destPath), "Invalid output path",
                    MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK);
                lblProgress.Content = "Something went wrong.";
                UIEnabled(true);

                return;
            }
            catch (Exception ex) {
                Logging.Logger.Log(ex);
                MessageBox.Show(ppToPdf.Resources.UnknownError, "Something went wrong",
                    MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK);
                lblProgress.Content = "Something went wrong.";
                UIEnabled(true);

                return;
            }
            finally {
                
            }

            // start conversion
            foreach (PptInfo p in listboxFiles.Items) {
                // update progress
                onProgressChanged.Report(new ProgressInfo(++index, listboxFiles.Items.Count, Path.GetFileName(p.Path)));

                // build path for converted file
                string newPath = Path.Combine(new[] { destPath, Path.ChangeExtension(Path.GetFileName(p.Path), "pdf") });

                try {
                    await ppConvertAsync(ppApp, p, newPath); // convert asynchronously
                }
                catch (COMException ex) {
                    Console.WriteLine(ex.Message);
                }
                catch (Exception ex) {
                    Logging.Logger.Log(ex);
                    MessageBox.Show(ppToPdf.Resources.UnknownError, "Conversion error",
                        MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK);
                    UIEnabled(true);

                    return;
                }

                try {
                    cancel.ThrowIfCancellationRequested();
                }
                catch (OperationCanceledException) {
                    ppApp.Quit();
                    Util.ReleaseComObject(ppApp);
                    throw;
                }
            }

            ppApp.Quit();
            Util.ReleaseComObject(ppApp);

            lblProgress.Content = "Done!";
        }

        /// <summary>
        /// Converts ppts to PDF documents asynchronously
        /// </summary>
        /// <param name="ppApp">The PowerPoint application</param>
        /// <param name="p">A PptInfo object to convert to pdf</param>
        /// <param name="newPath">Full path of the converted file</param>
        /// <returns></returns>
        private Task ppConvertAsync(PP.Application ppApp, PptInfo p, string newPath) {
            return Task.Run(() => {
                // open presentation in PP in the bg
                PP.Presentation pres = ppApp.Presentations.Open(p.Path, WithWindow: Microsoft.Office.Core.MsoTriState.msoFalse);

                // convert to pdf
                pres.SaveAs(newPath, PP.PpSaveAsFileType.ppSaveAsPDF);
                pres.Close();
                Util.ReleaseComObject(pres);
            });
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
                PptInfo pres = null;
                await Task.Run(() => {
                    if (s.Contains(".ppt")) {
                        try {
                            pres = new PptInfo(s);
                        }
                        catch (PathTooLongException ex) {
                            ex.Data.Add("Path", Path.GetDirectoryName(s));
                            throw;
                        }
                        catch (UnauthorizedAccessException ex) {
                            ex.Data.Add("Path", s);
                            throw;
                        }
                    }
                });

                if (pres != null) {
                    listboxFiles.Items.Add(pres);
                }
            }

            // create path to default output folder
            if (listboxFiles.Items.Count > 0) { // were files added?
                // if not use default
                string path = Path.GetDirectoryName((listboxFiles.Items[0] as PptInfo).Path);
                string dest = Path.Combine(new[] { path, "PDF-Presentations" });
                txtDestPath.Text = dest;

                // show listbox file count
                lblProgress.Content = string.Format("{0} presentation(s).", listboxFiles.Items.Count);
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
            txtDestPath.IsEnabled = state;
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
