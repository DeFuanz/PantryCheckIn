using Dapper;
using DocumentFormat.OpenXml.Office2010.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SQLite;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace pantry_CheckIn
{
    public class DataAccess
    {
        public static string LoadConnection(string id = "Default")
        {
            return ConfigurationManager.ConnectionStrings[id].ConnectionString;
        }
        public static DataTable LoadVisitors()
        {
            using (SQLiteConnection conn = new SQLiteConnection(LoadConnection()))
            {
                SQLiteCommand cmd = new SQLiteCommand("select visitors.ID, Visitors.'Last Name', visitors.'First Name', visitors.'Middle Name', v.'Visit Date', form.'Form Date', visitors.systemNo from visitors inner join (select visits.visitor, visits.'Visit Date', max(visits.entry) from visits group by visits.visitor)v on v.visitor = visitors.systemNo inner join form on form.visitor = visitors.systemNo group by visitors.systemNo order by visitors.'Last Name' asc, visitors.'First Name' asc limit 100", conn); 
                conn.Open();
                DataTable dt = new DataTable();
                SQLiteDataReader dr = cmd.ExecuteReader();
                dt.Load(dr);
                return dt;
            }
        }
        public static DataTable SearchVisitorsID()
        {
            using (SQLiteConnection conn = new SQLiteConnection(LoadConnection()))
            {
                SQLiteCommand cmd = new SQLiteCommand("select visitors.ID, Visitors.'Last Name', visitors.'First Name', visitors.'Middle Name', v.'Visit Date', form.'Form Date', visitors.systemNo from visitors inner join (select visits.visitor, visits.'Visit Date', max(visits.entry) from visits group by visits.visitor)v on v.visitor = visitors.systemNo inner join form on form.visitor = visitors.systemNo where visitors.ID = @id group by visitors.systemNo order by visitors.ID asc", conn);
                conn.Open();
                cmd.Parameters.AddWithValue("@id", FormInput.idNum);
                DataTable dt = new DataTable();
                SQLiteDataReader dr = cmd.ExecuteReader();
                dt.Load(dr);
                return dt;
            }
        }
        public static DataTable SearchIDWithNearbyNames()
        {
            using (SQLiteConnection conn = new SQLiteConnection(LoadConnection()))
            {
                SQLiteCommand cmd = new SQLiteCommand("select visitors.ID, Visitors.'Last Name', visitors.'First Name', visitors.'Middle Name', v.'Visit Date', form.'Form Date', visitors.systemNo from visitors inner join (select visits.visitor, visits.'Visit Date', max(visits.entry) from visits group by visits.visitor)v on v.visitor = visitors.systemNo inner join form on form.visitor = visitors.systemNo where visitors.ID = @id group by visitors.systemNo UNION ALL select * from (select visitors.ID, Visitors.'Last Name', visitors.'First Name', visitors.'Middle Name', v.'Visit Date', form.'Form Date', visitors.systemNo from visitors inner join (select visits.visitor, visits.'Visit Date', max(visits.entry) from visits group by visits.visitor)v on v.visitor = visitors.systemNo inner join form on form.visitor = visitors.systemNo where visitors.systemNo > @sysNo and visitors.ID != @id group by visitors.systemNo limit 3) UNION ALL select * from (select visitors.ID, Visitors.'Last Name', visitors.'First Name', visitors.'Middle Name', v.'Visit Date', form.'Form Date', visitors.systemNo from visitors inner join (select visits.visitor, visits.'Visit Date', max(visits.entry) from visits group by visits.visitor)v on v.visitor = visitors.systemNo inner join form on form.visitor = visitors.systemNo where visitors.systemNo < @sysNo and visitors.ID != @id group by visitors.systemNo limit 3)", conn);
                cmd.Parameters.AddWithValue("@id", FormInput.idNum);
                cmd.Parameters.AddWithValue("@sysNo", GetVisitor.systemNo);
                conn.Open();
                DataTable dt = new DataTable();
                SQLiteDataReader dr = cmd.ExecuteReader();
                dt.Load(dr);
                return dt;
            }
        }
        public static DataTable SearchVisitorsName()
        {
            using (SQLiteConnection conn = new SQLiteConnection(LoadConnection()))
            {
                SQLiteCommand cmd = new SQLiteCommand("select visitors.ID, Visitors.'Last Name', visitors.'First Name', visitors.'Middle Name', v.'Visit Date', form.'Form Date', visitors.systemNo from visitors inner join (select visits.visitor, visits.'Visit Date', max(visits.entry) from visits group by visits.visitor)v on v.visitor = visitors.systemNo inner join form on form.visitor = visitors.systemNo where visitors.'Last Name' like @name group by visitors.systemNo order by visitors.'Last Name' asc, visitors.'First Name' asc", conn);
                conn.Open();
                cmd.Parameters.AddWithValue("@name", FormInput.Name + "%");
                DataTable dt = new DataTable();
                SQLiteDataReader dr = cmd.ExecuteReader();
                dt.Load(dr);
                return dt;
            }
        }
        public static DataTable SearchVisitorsNameMultiple()
        {
            using (SQLiteConnection conn = new SQLiteConnection(LoadConnection()))
            {
                SQLiteCommand cmd = new SQLiteCommand("select visitors.ID, Visitors.'Last Name', visitors.'First Name', visitors.'Middle Name', v.'Visit Date', form.'Form Date', visitors.systemNo from visitors inner join (select visits.visitor, visits.'Visit Date', max(visits.entry) from visits group by visits.visitor)v on v.visitor = visitors.systemNo inner join form on form.visitor = visitors.systemNo where visitors.'Last Name' like @lastname and visitors.'First Name' like @firstname group by visitors.systemNo order by visitors.'Last Name' asc, visitors.'First Name' asc", conn);
                conn.Open();
                cmd.Parameters.AddWithValue("@lastname", FormInput.lname + "%");
                cmd.Parameters.AddWithValue("@firstname", FormInput.fname + "%");
                DataTable dt = new DataTable();
                SQLiteDataReader dr = cmd.ExecuteReader();
                dt.Load(dr);
                return dt;
            }
        }
        public static void CheckInVisitor()
        {
            using (SQLiteConnection conn = new SQLiteConnection(LoadConnection()))
            {
                SQLiteCommand cmd = new SQLiteCommand("insert into visits (visitor, 'Visit Date') values (@id, @today)", conn);
                conn.Open();
                cmd.Parameters.AddWithValue("@id", GetVisitor.systemNo);
                cmd.Parameters.AddWithValue("@today", DateTime.Today.ToString("MM/dd/yyyy"));
                cmd.ExecuteNonQuery();
                SQLiteCommand cmd2 = new SQLiteCommand("insert into pastvisitors (dateoflog, visitornames, visitorid, sysNo) values (@date, @name, @id, @sysNo)", conn);
                cmd2.Parameters.AddWithValue("@date", DateTime.Today.ToString("MM/dd/yyyy"));
                cmd2.Parameters.AddWithValue("@name", GetVisitor.firstname.ToUpper() + " " + GetVisitor.lastname.ToUpper());
                cmd2.Parameters.AddWithValue("id", GetVisitor.visitorid);
                cmd2.Parameters.AddWithValue("sysNo", GetVisitor.systemNo);
                cmd2.ExecuteNonQuery();
                SQLiteCommand cmd3 = new SQLiteCommand("insert into totalvisits ([visitorcount], [date]) values (@tally, @date)", conn);
                cmd3.Parameters.AddWithValue("@tally", 1);
                cmd3.Parameters.AddWithValue("@date", DateTime.Today.ToString("MM/dd/yyyy"));
                cmd3.ExecuteNonQuery();
            }
        }
        public static void UpdateVisitorInfo()
        {
            using (SQLiteConnection conn = new SQLiteConnection(LoadConnection()))
            {
                SQLiteCommand cmd = new SQLiteCommand("update visitors set ID = @id, 'First Name' = @fn, 'Last Name' = @ln, 'Middle Name' = @mi where systemNo = @sysNo", conn);
                conn.Open();
                cmd.Parameters.AddWithValue("@id", GetVisitor.visitorid);
                cmd.Parameters.AddWithValue("@fn", GetVisitor.firstname.ToUpper());
                cmd.Parameters.AddWithValue("@ln", GetVisitor.lastname.ToUpper());
                cmd.Parameters.AddWithValue("@mi", GetVisitor.middleinit.ToUpper());
                cmd.Parameters.AddWithValue("@sysNo", GetVisitor.systemNo);
                cmd.ExecuteNonQuery();
            }
        }
        public static DataTable ShowVisitorHistory()
        {
            using (SQLiteConnection conn = new SQLiteConnection(LoadConnection()))
            {
                SQLiteCommand cmd = new SQLiteCommand("Select visits.'Visit Date' from visits where visits.visitor = @sysNo and visits.'Visit Date' != '01/00/1900' order by entry desc", conn);
                conn.Open();
                cmd.Parameters.AddWithValue("@sysNo", GetVisitor.systemNo);
                DataTable dt = new DataTable();
                SQLiteDataReader dr = cmd.ExecuteReader();
                dt.Load(dr);
                return dt;
            }
        }
        public static void UpdateFormDate()
        {
            using (SQLiteConnection conn = new SQLiteConnection(LoadConnection()))
            {
                SQLiteCommand cmd = new SQLiteCommand("update form set 'Form Date' = @fd where visitor = @sysNo", conn);
                conn.Open();
                cmd.Parameters.AddWithValue("@fd", DateTime.Today.ToString("MM/dd/yyyy"));
                cmd.Parameters.AddWithValue("@sysNo", GetVisitor.systemNo);
                cmd.ExecuteNonQuery();
            }
        }
        public static void InsertNewVisitor()
        {
            int sysNo = 0;
            using (SQLiteConnection conn = new SQLiteConnection(LoadConnection()))
            {
                SQLiteCommand cmd = new SQLiteCommand("insert into visitors([ID], [First Name], [Last Name], [Middle Name]) values (@id, @fn, @ln, @mi)", conn);
                conn.Open();
                cmd.Parameters.AddWithValue("@id", NewVisitorInfo.newid);
                cmd.Parameters.AddWithValue("@fn", NewVisitorInfo.newfn.ToUpper());
                cmd.Parameters.AddWithValue("@ln", NewVisitorInfo.newln.ToUpper());
                cmd.Parameters.AddWithValue("@mi", NewVisitorInfo.newmi.ToUpper());
                cmd.ExecuteNonQuery();
                SQLiteCommand cmd2 = new SQLiteCommand("select max(visitors.systemNo) from visitors", conn);
                sysNo = Convert.ToInt32(cmd2.ExecuteScalar());
                SQLiteCommand cmd3 = new SQLiteCommand("insert into form([visitor], [Form Date]) values (@sysNo, @fd)", conn);
                cmd3.Parameters.AddWithValue("@sysNo", sysNo);
                cmd3.Parameters.AddWithValue("@fd", DateTime.Today.ToString("MM/dd/yyyy"));
                cmd3.ExecuteNonQuery();
                SQLiteCommand cmd4 = new SQLiteCommand("insert into visits([visitor], [Visit Date]) values (@sysNo, @vd)", conn);
                cmd4.Parameters.AddWithValue("@sysNo", sysNo);
                cmd4.Parameters.AddWithValue("@vd", DateTime.Today.ToString("MM/dd/yyyy"));
                cmd4.ExecuteNonQuery();
                SQLiteCommand cmd5 = new SQLiteCommand("insert into pastvisitors (dateoflog, visitornames, visitorid, sysNo) values (@date, @name, @id, @sysNo)", conn);
                cmd5.Parameters.AddWithValue("@date", DateTime.Today.ToString("MM/dd/yyyy"));
                cmd5.Parameters.AddWithValue("@name", NewVisitorInfo.newfn.ToUpper() + " " + NewVisitorInfo.newln.ToUpper());
                cmd5.Parameters.AddWithValue("id", NewVisitorInfo.newid);
                cmd5.Parameters.AddWithValue("@sysNo", sysNo);
                cmd5.ExecuteNonQuery();
                SQLiteCommand cmd6 = new SQLiteCommand("insert into totalvisits ([visitorcount], [date]) values (@tally, @date)", conn);
                cmd6.Parameters.AddWithValue("@tally", 1);
                cmd6.Parameters.AddWithValue("@date", DateTime.Today.ToString("MM/dd/yyyy"));
                cmd6.ExecuteNonQuery();
            }
        }
        public static void DeleteVisitorRecords()
        {
            using (SQLiteConnection conn = new SQLiteConnection(LoadConnection()))
            {
                SQLiteCommand cmd = new SQLiteCommand("delete from form where visitor = @sysNo; delete from visits where visitor = @sysNo; delete from visitors where systemNo = @sysNo", conn);
                conn.Open();
                cmd.Parameters.AddWithValue("@sysNo", GetVisitor.systemNo);
                cmd.ExecuteNonQuery();
            }
        }
        public static DataTable GetVisitorCount()
        {
            using (SQLiteConnection conn = new SQLiteConnection(LoadConnection()))
            {
                SQLiteCommand cmd = new SQLiteCommand("select date, sum(visitorcount) from totalvisits group by date order by date(date) desc", conn);
                conn.Open();
                DataTable dt = new DataTable();
                SQLiteDataReader dr = cmd.ExecuteReader();
                dt.Load(dr);
                return dt;
            }
        }
        public static int SumOfVisits()
        {
            using (SQLiteConnection conn = new SQLiteConnection(LoadConnection()))
            {
                SQLiteCommand cmd = new SQLiteCommand("select sum(visitorcount) from totalvisits where date = @date", conn);
                conn.Open();
                cmd.Parameters.AddWithValue("@date", DateTime.Today.ToString("MM/dd/yyyy"));
                SQLiteDataReader dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    if (dr["sum(visitorcount)"].ToString().Length > 0)
                    {
                        FormInput.visitorCount = Convert.ToInt32(dr["sum(visitorcount)"]);
                    }
                    else
                    {
                        FormInput.visitorCount = 0;
                    }
                }
                return FormInput.visitorCount;
            }
        }
        public static DataTable GetPreviousVisitors()
        {
            using (SQLiteConnection conn = new SQLiteConnection(LoadConnection()))
            {
                SQLiteCommand cmd = new SQLiteCommand("select * from pastvisitors where dateoflog = @date", conn);
                conn.Open();
                cmd.Parameters.AddWithValue("@date", PreviousVisitorDate.prevdate);
                DataTable dt = new DataTable();
                SQLiteDataReader dr = cmd.ExecuteReader();
                dt.Load(dr);
                return dt;
            }
        }
        public static DataTable ExportExcelFormat()
        {
            using (SQLiteConnection conn = new SQLiteConnection(LoadConnection()))
            {
                SQLiteCommand cmd = new SQLiteCommand("select visitors.ID, Visitors.'Last Name', visitors.'First Name', visitors.'Middle Name', form.'Form Date', v.'Visit Date' from visitors inner join (select visits.visitor, visits.'Visit Date', max(visits.entry) from visits group by visits.visitor)v on v.visitor = visitors.systemNo inner join form on form.visitor = visitors.systemNo group by visitors.systemNo order by visitors.'Last Name' asc", conn);
                conn.Open();
                DataTable dt = new DataTable();
                SQLiteDataReader dr = cmd.ExecuteReader();
                dt.Load(dr);
                return dt;
            }
        }
        public static DataTable ExportVisitorsDB()
        {
            using (SQLiteConnection conn = new SQLiteConnection(LoadConnection()))
            {
                SQLiteCommand cmd = new SQLiteCommand("select * from visitors", conn);
                conn.Open();
                DataTable dt = new DataTable();
                SQLiteDataReader dr = cmd.ExecuteReader();
                dt.Load(dr);
                return dt;
            }
        }
        public static DataTable ExportFormDB()
        {
            using (SQLiteConnection conn = new SQLiteConnection(LoadConnection()))
            {
                SQLiteCommand cmd = new SQLiteCommand("select * from form", conn);
                conn.Open();
                DataTable dt = new DataTable();
                SQLiteDataReader dr = cmd.ExecuteReader();
                dt.Load(dr);
                return dt;
            }
        }
        public static DataTable ExportvisitsDB()
        {
            using (SQLiteConnection conn = new SQLiteConnection(LoadConnection()))
            {
                SQLiteCommand cmd = new SQLiteCommand("select * from visits", conn);
                conn.Open();
                DataTable dt = new DataTable();
                SQLiteDataReader dr = cmd.ExecuteReader();
                dt.Load(dr);
                return dt;
            }
        }
        public static DataTable ExportPastVisitorsDB()
        {
            using (SQLiteConnection conn = new SQLiteConnection(LoadConnection()))
            {
                SQLiteCommand cmd = new SQLiteCommand("select * from pastvisitors", conn);
                conn.Open();
                DataTable dt = new DataTable();
                SQLiteDataReader dr = cmd.ExecuteReader();
                dt.Load(dr);
                return dt;
            }
        }
        public static DataTable ExportTotalVisitsDB()
        {
            using (SQLiteConnection conn = new SQLiteConnection(LoadConnection()))
            {
                SQLiteCommand cmd = new SQLiteCommand("select * from totalvisits", conn);
                conn.Open();
                DataTable dt = new DataTable();
                SQLiteDataReader dr = cmd.ExecuteReader();
                dt.Load(dr);
                return dt;
            }
        }
        public static DataTable ExportLastIDNumDB()
        {
            using (SQLiteConnection conn = new SQLiteConnection(LoadConnection()))
            {
                SQLiteCommand cmd = new SQLiteCommand("select * from newidnum", conn);
                conn.Open();
                DataTable dt = new DataTable();
                SQLiteDataReader dr = cmd.ExecuteReader();
                dt.Load(dr);
                return dt;
            }
        }
        public static void DeletePreviousVisitor()
        {
            using (SQLiteConnection conn = new SQLiteConnection(LoadConnection()))
            {
                SQLiteCommand cmd = new SQLiteCommand("delete from pastvisitors where sysNo = @sysNo and dateoflog = @date", conn);
                conn.Open();
                cmd.Parameters.AddWithValue("@sysNo", GetVisitor.systemNo);
                cmd.Parameters.AddWithValue("@date", DateTime.Today.ToString("MM/dd/yyyy"));
                cmd.ExecuteNonQuery();
                SQLiteCommand cmd2 = new SQLiteCommand("delete from totalvisits where daynum = (select max(daynum) from totalvisits)", conn);
                cmd2.ExecuteNonQuery();
                SQLiteCommand cmd3 = new SQLiteCommand("delete from visits where visitor = @sysNo and visits.entry = (select max(entry) from visits where visitor = @sysNo)", conn);
                cmd3.Parameters.AddWithValue("@sysNo", GetVisitor.systemNo);
                cmd3.ExecuteNonQuery();
            }
        }
        public static int GetNewVisitorNum()
        {
            using (SQLiteConnection conn = new SQLiteConnection(LoadConnection()))
            {
                SQLiteCommand cmd = new SQLiteCommand("select LastID from newidnum", conn);
                conn.Open();
                SQLiteDataReader dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    FormInput.lastEnteredNum = Convert.ToInt32(dr["LastID"].ToString());
                }
                return FormInput.lastEnteredNum;
            }
        }
        public static void UpdateNewVisitorNum()
        {
            using (SQLiteConnection conn = new SQLiteConnection(LoadConnection()))
            {
                SQLiteCommand cmd = new SQLiteCommand("update newidnum set LastID = @lastID + 1", conn);
                conn.Open();
                cmd.Parameters.AddWithValue("@lastID", FormInput.lastEnteredNum);
                cmd.ExecuteNonQuery();
            }
        }
        public static void CheckInVisitorNewDay()
        {
            using (SQLiteConnection conn = new SQLiteConnection(LoadConnection()))
            {
                SQLiteCommand cmd = new SQLiteCommand("insert into visits (visitor, 'Visit Date') values (@id, @date)", conn);
                conn.Open();
                cmd.Parameters.AddWithValue("@id", GetVisitor.systemNo);
                cmd.Parameters.AddWithValue("@date", FormInput.NewDayDate);
                cmd.ExecuteNonQuery();
                SQLiteCommand cmd2 = new SQLiteCommand("insert into pastvisitors (dateoflog, visitornames, visitorid, sysNo) values (@date, @name, @id, @sysNo)", conn);
                cmd2.Parameters.AddWithValue("@date", FormInput.NewDayDate);
                cmd2.Parameters.AddWithValue("@name", GetVisitor.firstname.ToUpper() + " " + GetVisitor.lastname.ToUpper());
                cmd2.Parameters.AddWithValue("id", GetVisitor.visitorid);
                cmd2.Parameters.AddWithValue("sysNo", GetVisitor.systemNo);
                cmd2.ExecuteNonQuery();
                SQLiteCommand cmd3 = new SQLiteCommand("insert into totalvisits ([visitorcount], [date]) values (@tally, @date)", conn);
                cmd3.Parameters.AddWithValue("@tally", 1);
                cmd3.Parameters.AddWithValue("@date", FormInput.NewDayDate);
                cmd3.ExecuteNonQuery();
            }
        }
        public static void InsertNewVisitorNewDay()
        {
            int sysNo = 0;
            using (SQLiteConnection conn = new SQLiteConnection(LoadConnection()))
            {
                SQLiteCommand cmd = new SQLiteCommand("insert into visitors([ID], [First Name], [Last Name], [Middle Name]) values (@id, @fn, @ln, @mi)", conn);
                conn.Open();
                cmd.Parameters.AddWithValue("@id", NewVisitorInfo.newid);
                cmd.Parameters.AddWithValue("@fn", NewVisitorInfo.newfn.ToUpper());
                cmd.Parameters.AddWithValue("@ln", NewVisitorInfo.newln.ToUpper());
                cmd.Parameters.AddWithValue("@mi", NewVisitorInfo.newmi.ToUpper());
                cmd.ExecuteNonQuery();
                SQLiteCommand cmd2 = new SQLiteCommand("select max(visitors.systemNo) from visitors", conn);
                sysNo = Convert.ToInt32(cmd2.ExecuteScalar());
                SQLiteCommand cmd3 = new SQLiteCommand("insert into form([visitor], [Form Date]) values (@sysNo, @fd)", conn);
                cmd3.Parameters.AddWithValue("@sysNo", sysNo);
                cmd3.Parameters.AddWithValue("@fd", FormInput.NewDayDate);
                cmd3.ExecuteNonQuery();
                SQLiteCommand cmd4 = new SQLiteCommand("insert into visits([visitor], [Visit Date]) values (@sysNo, @vd)", conn);
                cmd4.Parameters.AddWithValue("@sysNo", sysNo);
                cmd4.Parameters.AddWithValue("@vd", FormInput.NewDayDate);
                cmd4.ExecuteNonQuery();
                SQLiteCommand cmd5 = new SQLiteCommand("insert into pastvisitors (dateoflog, visitornames, visitorid, sysNo) values (@date, @name, @id, @sysNo)", conn);
                cmd5.Parameters.AddWithValue("@date", FormInput.NewDayDate);
                cmd5.Parameters.AddWithValue("@name", NewVisitorInfo.newfn.ToUpper() + " " + NewVisitorInfo.newln.ToUpper());
                cmd5.Parameters.AddWithValue("id", NewVisitorInfo.newid);
                cmd5.Parameters.AddWithValue("@sysNo", sysNo);
                cmd5.ExecuteNonQuery();
                SQLiteCommand cmd6 = new SQLiteCommand("insert into totalvisits ([visitorcount], [date]) values (@tally, @date)", conn);
                cmd6.Parameters.AddWithValue("@tally", 1);
                cmd6.Parameters.AddWithValue("@date", FormInput.NewDayDate);
                cmd6.ExecuteNonQuery();
            }
        }
        public static void UpdateFormDateNewDay()
        {
            using (SQLiteConnection conn = new SQLiteConnection(LoadConnection()))
            {
                SQLiteCommand cmd = new SQLiteCommand("update form set 'Form Date' = @fd where visitor = @sysNo", conn);
                conn.Open();
                cmd.Parameters.AddWithValue("@fd", FormInput.NewDayDate);
                cmd.Parameters.AddWithValue("@sysNo", GetVisitor.systemNo);
                cmd.ExecuteNonQuery();
            }
        }
        public static int SumOfVisitsNewDay()
        {
            using (SQLiteConnection conn = new SQLiteConnection(LoadConnection()))
            {
                SQLiteCommand cmd = new SQLiteCommand("select sum(visitorcount) from totalvisits where date = @date", conn);
                conn.Open();
                cmd.Parameters.AddWithValue("@date", FormInput.NewDayDate);
                SQLiteDataReader dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    if (dr["sum(visitorcount)"].ToString().Length > 0)
                    {
                        FormInput.visitorCount = Convert.ToInt32(dr["sum(visitorcount)"]);
                    }
                    else
                    {
                        FormInput.visitorCount = 0;
                    }
                }
                return FormInput.visitorCount;
            }
        }
        public static void DeletePreviousVisitorNewDay()
        {
            using (SQLiteConnection conn = new SQLiteConnection(LoadConnection()))
            {
                SQLiteCommand cmd = new SQLiteCommand("delete from pastvisitors where sysNo = @sysNo and dateoflog = @date", conn);
                conn.Open();
                cmd.Parameters.AddWithValue("@sysNo", GetVisitor.systemNo);
                cmd.Parameters.AddWithValue("@date", FormInput.NewDayDate);
                cmd.ExecuteNonQuery();
                SQLiteCommand cmd2 = new SQLiteCommand("delete from totalvisits where daynum = (select max(daynum) from totalvisits where date = @date)", conn);
                cmd2.Parameters.AddWithValue("@date", FormInput.NewDayDate);
                cmd2.ExecuteNonQuery();

                SQLiteCommand cmd3 = new SQLiteCommand("delete from visits where visitor = @sysNo and visits.'Visit Date' = @date", conn);
                cmd3.Parameters.AddWithValue("@sysNo", GetVisitor.systemNo);
                cmd3.Parameters.AddWithValue("@date", FormInput.NewDayDate);
                cmd3.ExecuteNonQuery();
            }
        }
    }
}
