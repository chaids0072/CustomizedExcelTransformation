using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelTransform
{
    class UpdateFunction
    {
        public static void UpdateProcessedRow() {
            Court.LABEL.Text = "Done: " + Excel.processedRows;
        }
    }
}
