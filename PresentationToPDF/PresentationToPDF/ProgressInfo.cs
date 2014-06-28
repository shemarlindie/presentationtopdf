using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PresentationToPDF {
    class ProgressInfo {
        public int CurrentIndex;
        public int MaxIndex;
        public string FileName;

        public ProgressInfo(int currentIndex, int maxIndex, string fileName) {
            CurrentIndex = currentIndex;
            MaxIndex = maxIndex;
            FileName = fileName;
        }
    }
}
