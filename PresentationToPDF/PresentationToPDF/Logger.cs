using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows;

namespace PresentationToPDF.Logging {
    static class Logger {
        static readonly long MaxFileSize = 10 * 1024 * 1024;            // 10 MB - Limit on the log file size
        static readonly string LogDir = Environment.CurrentDirectory;   // Containing directory of the log file           
        static readonly string LogFileName = "PresentationToPDF_ErrorLog.txt";  // Name of the log file
        static readonly string StringTemplate = "----\r\nPresentation To PDF {0}\r\nTimestamp: {1}\r\nException: {2}\r\nSource object: {3}\r\nMessage: {4}\r\nStack trace: {5}\r\n----\r\n\r\n";

        /// <summary>
        /// Initializes the Logger class.
        /// </summary>
        static Logger() {
            // limit log file size
            try {
                if (CheckFileSize() > MaxFileSize) {
                    File.Delete(LogPath);
                }
            }
            catch { }
        }

        /// <summary>
        /// Full path to the log file.
        /// </summary>
        public static string LogPath {
            get {
                return Path.Combine(LogDir, LogFileName);
            }
        }

        /// <summary>
        /// Logs the details of an exception.
        /// </summary>
        /// <param name="ex"></param>
        public static void Log(Exception ex) {
            Task.Factory.StartNew(() => {
                try {
                    File.AppendAllText(LogPath, string.Format(StringTemplate, AppInfo.Version, DetailedDateTimeString(),
                        ex.GetType().ToString(), ex.Source, ex.Message, ex.StackTrace), Encoding.UTF8);

                    if (!(ex is AggregateException) && ex.InnerException != null) {
                        Log(ex.InnerException);
                    }
                    else {
                        var aex = ex as AggregateException;
                        if (aex.InnerExceptions != null && aex.InnerExceptions.Count > 0) {
                            Console.WriteLine("**** InnerExceptions");
                            foreach (Exception e in aex.InnerExceptions) {
                                Log(e);
                            }
                            Console.WriteLine("****");
                        }
                    }
                }
                catch { }
            }, TaskCreationOptions.LongRunning);
        }

        /// <summary>
        /// Formatted date/time string of the current time and date.
        /// </summary>
        /// <returns></returns>
        public static string DetailedDateTimeString() {
            return string.Format("{0:MMM dd, yyyy hh:mm:ss.FFF tt}", DateTime.Now);
        }

        /// <summary>
        /// Size of the log file.
        /// </summary>
        /// <returns></returns>
        public static long CheckFileSize() {
            return new FileInfo(LogPath).Length;
        }
    }
}
