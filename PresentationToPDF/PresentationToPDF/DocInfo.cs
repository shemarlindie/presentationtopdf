using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;

namespace PresentationToPDF {
    /// <summary>
    /// Instantiates a new instance of the <code>PresentationToPdf.DocInfo</code> class.
    /// </summary>
    /// <param name="filePath">Path to the MS Word document.</param>
    /// <exception cref="System.IO.PathTooLongException"></exception>
    /// <exception cref="System.IO.UnauthorizedAccessException"></exception>
    class DocInfo : OfficeFileInfo {
        public DocInfo(string filePath)
            : base(filePath) {
        }

        public override string PageString {
            get {
                if (Pages >= 0) {
                    return String.Format("| {0} Page(s)", Pages);
                }
                else {
                    return string.Empty;
                }
            }

            protected set { } // do nothing
        }

        protected override int CountPages() {
            int count = 0;
            bool compatMode = System.IO.Path.GetExtension(Path) != ".docx";

            if (compatMode) { // use PP to get slide count (for older ppt files)
                // TODO - improve slide count speed for ppt files (slide count disable for now)
                //var ppApp = new Application();
                //Presentation p =  ppApp.Presentations.Open(Path, WithWindow: Microsoft.Office.Core.MsoTriState.msoFalse);
                //count = p.Slides.Count;
                //p.Close();
                //ppApp.Quit();
                return -1;
            }
            else { // use OpenXML method
                using (var doc = WordprocessingDocument.Open(Path, false)) {
                    string countString = doc.ExtendedFilePropertiesPart.Properties.Pages.Text;

                    if (!int.TryParse(countString, out count)) {
                        count = 0;
                    }
                }
            }

            return count;
        }
    }
}
