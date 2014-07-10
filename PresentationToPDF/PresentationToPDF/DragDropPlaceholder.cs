using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PresentationToPDF {
    class DragDropPlaceholder : OfficeFileInfo {
        public DragDropPlaceholder() {
            // so it can be instantiated in XAML
            
            Name = "Drag & drop ...";
            PageString = "";
            Path = "Drop your documents here.";
            SizeString = "";
        }

        public override string Name { get; protected set; }

        public override string PageString { get; protected set; }

        public override string Path { get; protected set; }

        public override string SizeString { get; protected set; }

        protected override int CountPages() {
            throw new NotImplementedException();
        }
    }
}
