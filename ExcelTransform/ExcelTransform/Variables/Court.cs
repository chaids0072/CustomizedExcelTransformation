using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelTransform
{
    class Court
    {
        public static System.Windows.Forms.Label LABEL;
        public static Excel excel;

        public static string currentExtraction = "";
        public static string currentExtractionSymbol = "";
        public static Boolean patternSelected = false;

        public static Boolean lineSeparated = false;
        public static Boolean numberSeparated = false;

        public static string expressionNumberExtraction = @"[\d]{1,4}([.,][\d]{0,10}[%]?)?";
        public static string expressionNumberExtractionSymbol = @"^[\u4e00-\u9fa5]{0,100}$";
    }
}
