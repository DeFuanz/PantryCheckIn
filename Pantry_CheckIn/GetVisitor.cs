using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace pantry_CheckIn
{
    public class GetVisitor
    {
        public static int visitorid { get; set; }
        public static string firstname { get; set; }
        public static string lastname { get; set; }
        public static string middleinit { get; set; }
        public static string visitdate { get; set; }
        public static string formneeded { get; set; }
        public static int systemNo { get; set; }
    }
    public class NewVisitorInfo
    {
        public static int newid { get; set; }
        public static string newfn { get; set; }
        public static string newln { get; set; }
        public static string newmi { get; set; }

    }
    public class PreviousVisitorDate
    {
        public static string prevdate { get; set; }
    }
}
