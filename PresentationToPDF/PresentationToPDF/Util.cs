using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

namespace PresentationToPDF {
    static class Util {
        public static void ReleaseComObject(object o) {
            try {
                Marshal.FinalReleaseComObject(o);
            }
            catch { }
        }

        public static Task CleanupUnusedMemoryAsync() {
            return Task.Run(() => {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            });
        }
    }
}
