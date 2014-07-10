using System;
using System.IO;

namespace PresentationToPDF {
    /// <summary>
    /// Base class for MS Office files' file information; Path, Size, Pages, etc.
    /// </summary>
    abstract class OfficeFileInfo {
        private string _path;

        public OfficeFileInfo() { }

        public OfficeFileInfo(string filePath) {
            using (FileStream f = File.OpenRead(filePath)) {
                Size = f.Length;
            }

            Path = filePath;
            Pages = CountPages();
        }

        /// <summary>
        /// Gets the file name.
        /// </summary>
        public virtual string Name { get; protected set; }

        /// <summary>
        /// Gets the number of pages in the file
        /// </summary>
        public virtual int Pages { get; protected set; }

        /// <summary>
        /// Gets the file path.
        /// </summary>
        public virtual string Path {
            get { return _path; }
            protected set {
                _path = value;
                Name = System.IO.Path.GetFileName(_path);
            }
        }

        /// <summary>
        /// Gets the file size (in bytes).
        /// </summary>
        public virtual long Size { get; protected set; }

        /// <summary>
        /// Gets the size formatted as a string
        /// </summary>
        public virtual string SizeString {
            get {
                if (Size < Math.Pow(1024, 2)) { // less than 1 MB
                    return string.Format("{0:N1} KB", Size / 1024.0f); // show in KB
                }
                else {
                    return string.Format("{0:N1} MB", Size / 1024.0f / 1024.0f); // show in MB
                }
            }
            protected set { } // do nothing
        }

        /// <summary>
        /// Gets the number of pages formatted as a string
        /// </summary>
        public abstract string PageString { get; protected set; }

        /// <summary>
        /// Counts the number of pages in the file
        /// </summary>
        /// <param name="countHidden">Include hidden pages in count</param>
        /// <returns>Number of slides</returns>
        protected abstract int CountPages();
    }
}
