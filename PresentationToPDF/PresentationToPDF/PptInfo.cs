using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Microsoft.Office.Interop.PowerPoint;
using DocumentFormat.OpenXml.Packaging;

namespace PresentationToPDF {
    /// <summary>
    /// Contains basic information about a PowerPoint presentation; Path, Size, Slides, etc.
    /// </summary>
    public class PptInfo {
        /// <summary>
        /// Instantiates a new instance of the <code>PresentationToPdf.PptInfo</code> class
        /// </summary>
        public PptInfo() {
            // So it can be instantiated in XAML to show a preview of the ListBox ItemTemplate
        }

        /// <summary>
        /// Instantiates a new instance of the <code>PresentationToPdf.PptInfo</code> class
        /// </summary>
        /// <param name="filePath">Path to the PowerPoint presentation</param>
        /// <exception cref="System.IO.PathTooLongException"></exception>
        /// <exception cref="System.IO.UnauthorizedAccessException"></exception>
        public PptInfo(string filePath) {
            using (FileStream f = File.OpenRead(filePath)) {
                Size = f.Length;
            }
            
            Name = System.IO.Path.GetFileName(filePath);
            Path = filePath;
            Slides = CountSlides(false);
        }

        /// <summary>
        /// Size (in bytes) of the presentation
        /// </summary>
        public long Size { get; set; }

        /// <summary>
        /// Size formatted as a string
        /// </summary>
        public string SizeString {
            get {
                if (Size < Math.Pow(1024, 2)) { // less than 1 MB
                    return string.Format("{0:N1} KB", Size / 1024.0f); // show in KB
                }
                else {
                    return string.Format("{0:N1} MB", Size / 1024.0f / 1024.0f); // show in MB
                }
            }
        }

        /// <summary>
        /// File name
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Full file path
        /// </summary>
        public string Path { get; set; }

        /// <summary>
        /// Number of slides in the presentation
        /// </summary>
        public int Slides { get; set; }

        /// <summary>
        /// Slide count formatted as a string
        /// </summary>
        public string SlideString {
            get {
                if (Slides >= 0) {
                    return String.Format("| {0} Slide(s)", Slides);
                }
                else {
                    return string.Empty;
                }
            }
        }


        /// <summary>
        /// Counts the number of slides in the presentation
        /// </summary>
        /// <param name="countHidden">Include hidden slide in count</param>
        /// <returns>Number of slides</returns>
        private int CountSlides(bool countHidden) {
            int count = 0;
            bool compatMode = System.IO.Path.GetExtension(Path) != ".pptx";

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
                using (var doc = PresentationDocument.Open(Path, false)) {
                    PresentationPart pres = doc.PresentationPart;

                    if (pres != null) {
                        if (countHidden) {
                            count = pres.SlideParts.Count();
                        }
                        else {
                            count = (from s in pres.SlideParts
                                     where (s.Slide.Show == null) ||
                                     (s.Slide.Show.HasValue && s.Slide.Show.Value)
                                     select s).Count();
                        }
                    }
                }
            }            

            return count;
        }
    }
}
