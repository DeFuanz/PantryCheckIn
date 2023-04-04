using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

namespace pantry_CheckIn
{
    public class FormInput
    {
        public static string searchCriteria { get; set; }
        public static int idNum { get; set; }
        public static string Name { get; set; }
        public static string fname { get; set; }
        public static string lname { get; set; }

        public static int lastEnteredNum { get; set; }

        public static int visitorCount { get; set; }

        public static string NewDayDate { get; set; }

        public static bool InNewWorkDay = false;
    }
}
