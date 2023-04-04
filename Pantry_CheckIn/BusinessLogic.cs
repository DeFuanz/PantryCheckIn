using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace pantry_CheckIn
{
    public class BusinessLogic
    {

        public static DataTable FilterSearch()
        {
            
            //filters search entry to decide on query type
            if (int.TryParse(FormInput.searchCriteria, out _))
            {
                FormInput.idNum = int.Parse(FormInput.searchCriteria);
                /*if (DataAccess.SearchVisitorsID().Rows.Count > 0)
                {*/
                try
                {
                    GetVisitor.systemNo = Convert.ToInt32(DataAccess.SearchVisitorsID().Rows[0]["systemNo"]);
                    return DataAccess.SearchIDWithNearbyNames();
                }
                catch
                {
                    return DataAccess.SearchVisitorsID();
                }
                /*}
                else
                {
                    return DataAccess.SearchVisitorsID();
                }*/
            }
            else if (FormInput.searchCriteria.Trim() == FormInput.searchCriteria && FormInput.searchCriteria.IndexOf(" ") > -1)
            {
                string[] splitnames = FormInput.searchCriteria.Split(' '); 

                FormInput.fname = splitnames[1];
                FormInput.lname = splitnames[0];

                return DataAccess.SearchVisitorsNameMultiple();
            }
            else
            {
                FormInput.Name = FormInput.searchCriteria;
                return DataAccess.SearchVisitorsName();
            }
        }
        public static string CheckForForm()
        {
            //parses dates to check if form was filed for this year
            DateTime parsedDate = DateTime.Parse(GetVisitor.formneeded);
            DateTime today = DateTime.Now;
            if (today.Year == parsedDate.Year)
            {
                return "On File";
            }
            else
            {
                return "Missing";
            }
        }
    }
}
