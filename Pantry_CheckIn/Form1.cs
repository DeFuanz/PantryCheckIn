using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Office2021.DocumentTasks;
using System;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Windows.Forms;
using System.Data;
using System.Threading.Tasks;

namespace pantry_CheckIn
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            //Hides and Shows certain data on startup
            try
            {
                InitializeComponent();

                dataGridView_visitors.DataSource = DataAccess.LoadVisitors();
                dataGridView_visitors.Columns[4].Visible = false;
                dataGridView_visitors.Columns[5].Visible = false;
                dataGridView_visitors.Columns[6].Visible = false;
                label_datetoday.Text = DateTime.Today.ToString("D");
                label_VisitorCount.Text = DataAccess.SumOfVisits().ToString();
                panel_updateinfo.Hide();
                panel_vistorhistory.Hide();
                panel_addvisitor.Hide();
                panel_counthistory.Hide();
                panel_previousvisitors.Hide();
                panel_settings.Hide();
                panel_searchresults.Show();
                panel_newdaydate.Hide();
                panel_newworkdayalert.Hide();
                label_lastvisitorrecorded.Text = "";
                label_selectedvisitortodelete.Text = "";
                label_countdate.Text = "";
                label_newworkdate.Text = "";
                label_loading.Hide();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void SearchButton_Click(object sender, EventArgs e)
        {
            //Takes search text and returns data queried
            button_search.Enabled = false;
            label_loading.Show();
            label_loading.Refresh();
            FormInput.searchCriteria = textBox_searchCriteria.Text.ToString();
            dataGridView_visitors.DataSource = null;
            dataGridView_visitors.DataSource = BusinessLogic.FilterSearch();
            label_loading.Hide();
            button_search.Enabled = true;
            dataGridView_visitors.ClearSelection();
            dataGridView_visitors.Columns[4].Visible = false;
            dataGridView_visitors.Columns[5].Visible = false;
            dataGridView_visitors.Columns[6].Visible = false;
        }

        private void button_clear_Click(object sender, EventArgs e)
        {
            //Clears labels and search results - just as would show on loadup
            textBox_searchCriteria.Text = "";
            dataGridView_visitors.DataSource = null;
            dataGridView_visitors.DataSource = DataAccess.LoadVisitors();
            dataGridView_visitors.Columns[4].Visible = false;
            dataGridView_visitors.Columns[5].Visible = false;
            dataGridView_visitors.Columns[6].Visible = false;
            label_selectedid.Text = "";
            label_selectedfirstname.Text = "";
            label_selectedlastname.Text = "";
            label_selectedmidde.Text = "";
            label_selecteddate.Text = "";
            label_selectedform.Text = "";
            label_selectedform.BackColor = Color.Transparent;
            label_selecteddate.BackColor = Color.Transparent;
            label_selectedid.BackColor = Color.Transparent;
            label_selectedfirstname.BackColor = Color.Transparent;
            label_selectedlastname.BackColor = Color.Transparent;
            label_selectedmidde.BackColor = Color.Transparent;
            button_checkin.BackColor = Color.White;
        }

        private void Checkin_button_Click(object sender, EventArgs e)
        {
            //Checks to verify no information is missing - checks in visitors or gives message on reason it could not
            if (label_selectedid.Text == "")
            {
                MessageBox.Show("Please select a visitor");
                return;
            }
            else if (label_selectedid.Text == "Missing ID")
            {
                DialogResult = MessageBox.Show("Visitors is missing ID. Please update information before proceeding", "alert", MessageBoxButtons.OKCancel);
                if (DialogResult == DialogResult.OK)
                {
                    panel_searchresults.Hide();
                    panel_updateinfo.Show();
                    textBox_updateid.Text = label_selectedid.Text;
                    textBox_updatefn.Text = label_selectedfirstname.Text;
                    textBox_updateln.Text = label_selectedlastname.Text;
                    textBox_updatemi.Text = label_selectedmidde.Text;
                    button_deletevisitor.Enabled = false;
                    button4.Enabled = false;
                    button_viewhistory.Enabled = false;
                    return;
                }
                else
                {
                    return;
                }
            }
            else if (label_selecteddate.Text == DateTime.Today.ToString("MM/dd/yyyy"))
            {
                MessageBox.Show("Unable to Check in. Visitor Already checked in Today.");
                return;
            }
            else
            {
                if (label_selectedform.Text == "On File" && label_selecteddate.BackColor == Color.LightGreen)
                {
                    if (FormInput.InNewWorkDay == true)
                    {
                        DataAccess.CheckInVisitorNewDay();
                    }
                    else
                    {
                        DataAccess.CheckInVisitor();
                    }
                    MessageBox.Show("Visitor Checked In");
                    label_lastvisitorrecorded.Text = $"{GetVisitor.firstname} {GetVisitor.lastname} - {GetVisitor.visitorid}";
                    button_checkin.BackColor = Color.White;
                }
                else
                {
                    if (label_selectedform.Text == "Missing")
                    {
                        DialogResult = MessageBox.Show("Did visitor fill out a form today?", "ALERT", MessageBoxButtons.YesNo);
                        if (DialogResult == DialogResult.Yes)
                        {
                            if (FormInput.InNewWorkDay == true)
                            {
                                DataAccess.CheckInVisitorNewDay();
                                DataAccess.UpdateFormDateNewDay();

                            }
                            else
                            {
                                DataAccess.UpdateFormDate();
                                DataAccess.CheckInVisitor();
                            }
                            MessageBox.Show("Visitor Checked In");
                            label_lastvisitorrecorded.Text = $"{GetVisitor.firstname} {GetVisitor.lastname} - {GetVisitor.visitorid}";
                            button_checkin.BackColor = Color.White;
                        }
                        else
                        {
                            MessageBox.Show("Please have the visitor fill out a form");
                            return;
                        }
                    }
                    else
                    {
                        DialogResult = MessageBox.Show("Visitor may be too early. Do you want to still check them in?", "ALERT", MessageBoxButtons.YesNo);
                        if (DialogResult == DialogResult.Yes)
                        {
                            if (FormInput.InNewWorkDay == true)
                            {
                                DataAccess.CheckInVisitorNewDay();
                                DataAccess.UpdateFormDateNewDay();

                            }
                            else
                            {
                                DataAccess.UpdateFormDate();
                                DataAccess.CheckInVisitor();
                            }
                            MessageBox.Show("Visitor Checked In");
                            label_lastvisitorrecorded.Text = $"{GetVisitor.firstname} {GetVisitor.lastname} - {GetVisitor.visitorid}";
                            button_checkin.BackColor = Color.White;
                        }
                        else
                        {
                            MessageBox.Show("Please inform visitor of next eligible date");
                            return;
                        }
                    }
                }
            }
            int currentRow = dataGridView_visitors.CurrentCell.RowIndex;
            dataGridView_visitors.DataSource = null;
            dataGridView_visitors.DataSource = DataAccess.LoadVisitors();
            dataGridView_visitors.Columns[4].Visible = false;
            dataGridView_visitors.Columns[5].Visible = false;
            dataGridView_visitors.Columns[6].Visible = false;
            dataGridView_visitors.Rows[currentRow].Selected = true;
            label_selectedid.Text = "";
            label_selectedfirstname.Text = "";
            label_selectedlastname.Text = "";
            label_selectedmidde.Text = "";
            label_selecteddate.Text = "";
            label_selectedform.Text = "";
            textBox_newvisitorid.Text = "";
            textBox_newvisitorfirst.Text = "";
            textBox_newvisitorlast.Text = "";
            textBox_newvisitormiddle.Text = "";
            label_selectedform.BackColor = Color.Transparent;
            label_selecteddate.BackColor = Color.Transparent;
            if (FormInput.InNewWorkDay == true)
            {
                label_VisitorCount.Text = DataAccess.SumOfVisitsNewDay().ToString();
            }
            else
            {
                label_VisitorCount.Text = DataAccess.SumOfVisits().ToString();
            }
            label_selectedid.BackColor = Color.Transparent;
            label_selectedfirstname.BackColor = Color.Transparent;
            label_selectedlastname.BackColor = Color.Transparent;
            label_selectedmidde.BackColor = Color.Transparent;
        }

        private void UpdateInfo_button_Click(object sender, EventArgs e)
        {
            //Verifies visitor is selected then opens information update page
            if (label_selectedid.Text == "")
            {
                MessageBox.Show("Please select a visitor to update");
                return;
            }
            else
            {
                panel_searchresults.Hide();
                panel_updateinfo.Show();
                textBox_updateid.Text = label_selectedid.Text;
                textBox_updatefn.Text = label_selectedfirstname.Text;
                textBox_updateln.Text = label_selectedlastname.Text;
                textBox_updatemi.Text = label_selectedmidde.Text;
                button_deletevisitor.Enabled = false;
                button4.Enabled = false;
                button_checkin.Enabled = false;
                button_deletevisitor.Enabled = false;
                button_updateinfo.Enabled = false;
                button_viewhistory.Enabled = false;
            }
        }

        private void button_saveinfo_Click(object sender, EventArgs e)
        {
            //Checks to make sure the information entered is valid, then updates it in database
            if (textBox_updateid.Text == "" || textBox_updatefn.Text == "" || textBox_updateln.Text == "")
            {
                MessageBox.Show("Please fill out all information for visitor");
                return;
            }
            else
            {
                if (int.TryParse(textBox_updateid.Text, out _))
                {
                    if (System.Text.RegularExpressions.Regex.IsMatch(textBox_updatefn.Text, "^[a-zA-Z ]"))
                    {
                        if (System.Text.RegularExpressions.Regex.IsMatch(textBox_updateln.Text, "^[a-zA-Z ]"))
                        {

                            GetVisitor.visitorid = Convert.ToInt32(textBox_updateid.Text);
                            GetVisitor.firstname = textBox_updatefn.Text.Trim();
                            GetVisitor.lastname = textBox_updateln.Text.Trim();
                            GetVisitor.middleinit = textBox_updatemi.Text.Trim();
                            DataAccess.UpdateVisitorInfo();
                            if (label_selecteddate.Text != DateTime.Today.ToString("MM/dd/yyyy"))
                            {
                                DialogResult = MessageBox.Show("Do you want to also check the visitor in?", "Check In", MessageBoxButtons.YesNo);
                                if (DialogResult == DialogResult.Yes)
                                {
                                    if (label_selectedform.Text == "Missing")
                                    {
                                        DialogResult = MessageBox.Show("Form Missing: Did visitor fill out a form today?", "Alert", MessageBoxButtons.YesNo);
                                        if (DialogResult == DialogResult.Yes)
                                        {
                                            if (FormInput.InNewWorkDay == true)
                                            {
                                                DataAccess.CheckInVisitorNewDay();
                                                DataAccess.UpdateFormDateNewDay();
                                                label_VisitorCount.Text = DataAccess.SumOfVisitsNewDay().ToString();

                                            }
                                            else
                                            {
                                                DataAccess.UpdateFormDate();
                                                DataAccess.CheckInVisitor();
                                                label_VisitorCount.Text = DataAccess.SumOfVisits().ToString();
                                            }
                                            label_lastvisitorrecorded.Text = $"{GetVisitor.firstname} {GetVisitor.lastname} - {GetVisitor.visitorid}";
                                            label_selectedid.Text = "";
                                            label_selectedfirstname.Text = "";
                                            label_selectedlastname.Text = "";
                                            label_selectedmidde.Text = "";
                                            label_selectedform.Text = "";
                                            label_selecteddate.Text = "";
                                            label_selectedid.BackColor = Color.Transparent;
                                            label_selectedfirstname.BackColor = Color.Transparent;
                                            label_selectedlastname.BackColor = Color.Transparent;
                                            label_selectedmidde.BackColor = Color.Transparent;
                                            label_selecteddate.BackColor = Color.Transparent;
                                            label_selectedform.BackColor = Color.Transparent;
                                        }
                                        else
                                        {
                                            MessageBox.Show("Please have visitor fill out a form before checking in");
                                        }
                                    }
                                    else
                                    {
                                        if (FormInput.InNewWorkDay == true)
                                        {
                                            DataAccess.CheckInVisitorNewDay();
                                            DataAccess.UpdateFormDateNewDay();
                                            label_VisitorCount.Text = DataAccess.SumOfVisitsNewDay().ToString();

                                        }
                                        else
                                        {
                                            DataAccess.UpdateFormDate();
                                            DataAccess.CheckInVisitor();
                                            label_VisitorCount.Text = DataAccess.SumOfVisits().ToString();
                                        }
                                        label_lastvisitorrecorded.Text = $"{GetVisitor.firstname} {GetVisitor.lastname} - {GetVisitor.visitorid}";
                                        label_selectedid.Text = "";
                                        label_selectedfirstname.Text = "";
                                        label_selectedlastname.Text = "";
                                        label_selectedmidde.Text = "";
                                        label_selectedform.Text = "";
                                        label_selecteddate.Text = "";
                                        label_selectedid.BackColor = Color.Transparent;
                                        label_selectedfirstname.BackColor = Color.Transparent;
                                        label_selectedlastname.BackColor = Color.Transparent;
                                        label_selectedmidde.BackColor = Color.Transparent;
                                        label_selecteddate.BackColor = Color.Transparent;
                                        label_selectedform.BackColor = Color.Transparent;
                                    }
                                }
                                else
                                {
                                    label_selectedid.Text = GetVisitor.visitorid.ToString();
                                    label_selectedfirstname.Text = GetVisitor.firstname.ToString();
                                    label_selectedlastname.Text = GetVisitor.lastname.ToString();
                                    label_selectedmidde.Text = GetVisitor.middleinit.ToString();
                                }
                            }
                            int currentRow = dataGridView_visitors.CurrentCell.RowIndex;
                            panel_updateinfo.Hide();
                            panel_searchresults.Show();
                            dataGridView_visitors.DataSource = DataAccess.LoadVisitors();
                            dataGridView_visitors.Rows[currentRow].Selected = true;
                            dataGridView_visitors.Columns[4].Visible = false;
                            dataGridView_visitors.Columns[5].Visible = false;
                            dataGridView_visitors.Columns[6].Visible = false;
                            button_viewhistory.Enabled = true;
                            button_deletevisitor.Enabled = true;
                            button4.Enabled = true;
                            button_checkin.Enabled = true;
                            button_updateinfo.Enabled = true;
                            button_viewhistory.Enabled = true;
                            MessageBox.Show("Visitor information updated");
                        }
                        else
                        {
                            MessageBox.Show("Please enter a valid Last Name");
                        }
                    }
                    else
                    {
                        MessageBox.Show("Please enter a valid First Name");
                    }
                }
                else
                {
                    MessageBox.Show("Visitor ID must be a number");
                    return;
                }
            }
        }

        private void ShowHistory_button_Click(object sender, EventArgs e)
        {
            //Opens panel to show recent visits if visitor is selected
            if (label_selectedid.Text == "")
            {
                MessageBox.Show("Select a visitor to view history");
                return;
            }
            else
            {
                panel_searchresults.Hide();
                label_visithistoryname.Text = $"{GetVisitor.firstname} {GetVisitor.lastname}";
                panel_vistorhistory.Show();
                dataGridView_history.DataSource = DataAccess.ShowVisitorHistory();
                button_checkin.Enabled = false;
                button_deletevisitor.Enabled = false;
                button_updateinfo.Enabled = false;
                button_viewhistory.Enabled = false;
            }
        }

        private void button_backnewvisitor_Click(object sender, EventArgs e)
        {
            //Cancels adding new visitor
            panel_addvisitor.Hide();
            panel_searchresults.Show();
            textBox_newvisitorid.Text = "";
            textBox_newvisitorfirst.Text = "";
            textBox_newvisitormiddle.Text = "";
            textBox_newvisitorlast.Text = "";
            button_checkin.Enabled = true;
            button_deletevisitor.Enabled = true;
            button_updateinfo.Enabled = true;
            button_viewhistory.Enabled = true;
            button4.Enabled = true;
        }

        private void button_submitnewvisitor_Click(object sender, EventArgs e)
        {
            //Checks if information is valid then submits new visitor into the database
            if (int.TryParse(textBox_newvisitorid.Text, out _))
            {
                if (System.Text.RegularExpressions.Regex.IsMatch(textBox_newvisitorfirst.Text, "^[a-zA-Z ]"))
                {
                    if (System.Text.RegularExpressions.Regex.IsMatch(textBox_newvisitorlast.Text, "^[a-zA-Z ]"))
                    {
                        FormInput.fname = textBox_newvisitorfirst.Text.Trim();
                        FormInput.lname = textBox_newvisitorlast.Text.Trim();
                        if (DataAccess.SearchVisitorsNameMultiple().Rows.Count > 0)
                        {
                            DialogResult = MessageBox.Show("Possible Match Found: Would you like to check if the visitor already exists?", "ALERT", MessageBoxButtons.YesNo);
                            if (DialogResult == DialogResult.Yes)
                            {
                                dataGridView_visitors.DataSource = null;
                                dataGridView_visitors.DataSource = DataAccess.SearchVisitorsNameMultiple();
                                panel_addvisitor.Visible = false;
                                panel_searchresults.Visible = true;
                                dataGridView_visitors.Columns[4].Visible = false;
                                dataGridView_visitors.Columns[5].Visible = false;
                                dataGridView_visitors.Columns[6].Visible = false;
                                button_checkin.Enabled = true;
                                button_deletevisitor.Enabled = true;
                                button_updateinfo.Enabled = true;
                                button_viewhistory.Enabled = true;
                                button4.Enabled = true;
                                return;
                            }
                        }
                        DialogResult = MessageBox.Show("Is everything correct? This will add and check in the visitor", "ALERT", MessageBoxButtons.YesNo);
                        if (DialogResult == DialogResult.Yes)
                        {
                            NewVisitorInfo.newid = Convert.ToInt32(textBox_newvisitorid.Text);
                            NewVisitorInfo.newfn = textBox_newvisitorfirst.Text.Trim();
                            NewVisitorInfo.newln = textBox_newvisitorlast.Text.Trim();
                            NewVisitorInfo.newmi = textBox_newvisitormiddle.Text.Trim();
                            if (FormInput.InNewWorkDay == true)
                            {
                                DataAccess.InsertNewVisitorNewDay();
                                label_VisitorCount.Text = DataAccess.SumOfVisitsNewDay().ToString();
                            }
                            else
                            {
                                FormInput.lastEnteredNum = Convert.ToInt32(textBox_newvisitorid.Text.Trim());
                                DataAccess.InsertNewVisitor();
                                label_VisitorCount.Text = DataAccess.SumOfVisits().ToString();
                                DataAccess.UpdateNewVisitorNum();
                            }
                            textBox_newvisitorid.Text = "";
                            textBox_newvisitorfirst.Text = "";
                            textBox_newvisitormiddle.Text = "";
                            textBox_newvisitorlast.Text = "";
                            MessageBox.Show("New visitor added and checked in");
                            panel_addvisitor.Hide();
                            dataGridView_visitors.DataSource = DataAccess.LoadVisitors();
                            dataGridView_visitors.Columns[4].Visible = false;
                            dataGridView_visitors.Columns[5].Visible = false;
                            dataGridView_visitors.Columns[6].Visible = false;
                            panel_searchresults.Show();
                            button_checkin.Enabled = true;
                            button_deletevisitor.Enabled = true;
                            button_updateinfo.Enabled = true;
                            button_viewhistory.Enabled = true;
                            button4.Enabled = true;
                            label_lastvisitorrecorded.Text = $"{NewVisitorInfo.newfn.ToUpper()} {NewVisitorInfo.newln.ToUpper()} - {NewVisitorInfo.newid}";
                        }
                    }
                    else
                    {
                        MessageBox.Show("Please enter a valid Last Name");
                    }
                }
                else
                {
                    MessageBox.Show("Please enter a valid First Name");
                }

            }
            else
            {
                MessageBox.Show("Please enter a numerical ID number");
            }
        }

        private void button_addnewvisitor(object sender, EventArgs e)
        {
            //opens add visitor panel
            panel_addvisitor.Show();
            panel_searchresults.Hide();
            button_checkin.Enabled = false;
            button_deletevisitor.Enabled = false;
            button_updateinfo.Enabled = false;
            button_viewhistory.Enabled = false;
            button4.Enabled = false;
            FormInput.lastEnteredNum = DataAccess.GetNewVisitorNum();
            if (FormInput.lastEnteredNum != 0)
            {
                textBox_newvisitorid.Text = FormInput.lastEnteredNum.ToString();
            }
            else
            {
                textBox_newvisitorid.Text = "";
            }
        }

        private void button_deletevisitor_Click(object sender, EventArgs e)
        {
            //Deletes selected visitor from database
            if (label_selectedid.Text == "")
            {
                MessageBox.Show("Select a visitor to delete");
            }
            else
            {
                DialogResult = MessageBox.Show($"Are you sure you want to delete {GetVisitor.firstname} {GetVisitor.lastname} from the system? This cannot be undone", "ALERT", MessageBoxButtons.YesNo);
                if (DialogResult == DialogResult.Yes)
                {
                    DataAccess.DeleteVisitorRecords();
                    label_selectedid.Text = "";
                    label_selectedfirstname.Text = "";
                    label_selectedlastname.Text = "";
                    label_selectedmidde.Text = "";
                    label_selecteddate.Text = "";
                    label_selectedform.Text = "";
                    label_selectedform.BackColor = Color.Transparent;
                    label_selecteddate.BackColor = Color.Transparent;
                    dataGridView_visitors.DataSource = DataAccess.LoadVisitors();
                    dataGridView_visitors.Columns[4].Visible = false;
                    dataGridView_visitors.Columns[5].Visible = false;
                    dataGridView_visitors.Columns[6].Visible = false;
                    label_selectedid.BackColor = Color.Transparent;
                    label_selectedfirstname.BackColor = Color.Transparent;
                    label_selectedlastname.BackColor = Color.Transparent;
                    label_selectedmidde.BackColor = Color.Transparent;
                }
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            //Verifies user wants to close out the application
            DialogResult = MessageBox.Show("Are you finished for the day?", "ALERT", MessageBoxButtons.YesNo);
            if (DialogResult == DialogResult.No)
            {
                e.Cancel = true;
            }
        }

        private void button_exitvisithist_Click(object sender, EventArgs e)
        {
            //Closes out previous visitors panel
            panel_vistorhistory.Hide();
            panel_searchresults.Show();
            button_checkin.Enabled = true;
            button_deletevisitor.Enabled = true;
            button_updateinfo.Enabled = true;
            button_viewhistory.Enabled = true;
        }

        private void button_showpreviousclicks(object sender, EventArgs e)
        {
            //Shows previous daily visitor counts panel
            panel_searchresults.Hide();
            panel_counthistory.Show();
            dataGridView_visitorcount.DataSource = DataAccess.GetVisitorCount();
            button_checkin.Enabled = false;
            button_deletevisitor.Enabled = false;
            button_updateinfo.Enabled = false;
            button_viewhistory.Enabled = false;
        }

        private void button_closevisitcount(object sender, EventArgs e)
        {
            //closes previous visitors and returns to search results
            panel_counthistory.Hide();
            panel_previousvisitors.Hide();
            panel_searchresults.Show();
            button_checkin.Enabled = true;
            button_deletevisitor.Enabled = true;
            button_updateinfo.Enabled = true;
            button_viewhistory.Enabled = true;
            label_selectedvisitortodelete.Text = "";
        }

        private void button_exitpreviousvisitors_Click(object sender, EventArgs e)
        {
            //exits previous visitors
            panel_previousvisitors.Hide();
            button_checkin.Enabled = true;
            button_deletevisitor.Enabled = true;
            button_updateinfo.Enabled = true;
            button_viewhistory.Enabled = true;
            label_selectedvisitortodelete.Text = "";
        }

        private void dataGridView_visitorcount_CellContentClick_1(object sender, DataGridViewCellEventArgs e)
        {
            //Chooses date to view history
            PreviousVisitorDate.prevdate = dataGridView_visitorcount.SelectedCells[0].Value.ToString();
            label_countdate.Text = dataGridView_visitorcount.SelectedCells[0].Value.ToString();

            dataGridView_previousvisitors.DataSource = null;
            dataGridView_previousvisitors.DataSource = DataAccess.GetPreviousVisitors();
            dataGridView_previousvisitors.Columns[0].Visible = false;
            dataGridView_previousvisitors.Columns[1].Visible = false;
            dataGridView_previousvisitors.Columns[4].Visible = false;
            panel_previousvisitors.Show();
            button_checkin.Enabled = false;
            button_deletevisitor.Enabled = false;
            button_updateinfo.Enabled = false;
            button_viewhistory.Enabled = false;
        }

        private void dataGridView_visitors_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            //Selects visitor from search and filters then displays information
            try
            {
                var testnull = dataGridView_visitors.SelectedCells[0].Value;
                if (testnull is DBNull)
                {
                    label_selectedid.Text = "Missing ID";
                }
                else
                {
                    GetVisitor.visitorid = Convert.ToInt32(dataGridView_visitors.SelectedCells[0].Value);
                    label_selectedid.Text = GetVisitor.visitorid.ToString();
                }
                GetVisitor.firstname = dataGridView_visitors.SelectedCells[2].Value.ToString();
                GetVisitor.lastname = dataGridView_visitors.SelectedCells[1].Value.ToString();
                GetVisitor.middleinit = dataGridView_visitors.SelectedCells[3].Value.ToString();
                label_selectedid.BackColor = Color.White;
                label_selectedfirstname.BackColor = Color.White;
                label_selectedlastname.BackColor = Color.White;
                label_selectedmidde.BackColor = Color.White;
                if (dataGridView_visitors.SelectedCells[4].Value.ToString() == "01/00/1900")
                {
                    GetVisitor.visitdate = "N/A";
                    label_selecteddate.BackColor = Color.LightGreen;
                }
                else
                {
                    GetVisitor.visitdate = dataGridView_visitors.SelectedCells[4].Value.ToString();
                    if ((DateTime.Today.Date - Convert.ToDateTime(dataGridView_visitors.SelectedCells[4].Value)).Days >= 30)
                    {
                        label_selecteddate.BackColor = Color.LightGreen;
                    }
                    else
                    {
                        label_selecteddate.BackColor = Color.Red;
                    }
                }
                GetVisitor.formneeded = dataGridView_visitors.SelectedCells[5].Value.ToString();
                GetVisitor.systemNo = Convert.ToInt32(dataGridView_visitors.SelectedCells[6].Value);
                label_selectedfirstname.Text = GetVisitor.firstname.ToString();
                label_selectedlastname.Text = GetVisitor.lastname.ToString();
                label_selectedmidde.Text = GetVisitor.middleinit.ToString();
                label_selecteddate.Text = GetVisitor.visitdate.ToString();
                label_selectedform.Text = BusinessLogic.CheckForForm();
                if (label_selectedform.Text == "On File")
                {
                    label_selectedform.BackColor = Color.LightGreen;
                }
                else
                {
                    label_selectedform.BackColor = Color.Red;
                }
                if (label_selecteddate.BackColor == Color.Red || label_selectedform.BackColor == Color.Red)
                {
                    button_checkin.BackColor = Color.Red;
                }
                else
                {
                    button_checkin.BackColor = Color.LightGreen;
                }
            }
            catch (Exception)
            {
                MessageBox.Show("No Results Found in Search");
            }
        }

        private void button_cancelupdateinfo_Click(object sender, EventArgs e)
        {
            //cancels update information page
            panel_updateinfo.Hide();
            panel_searchresults.Show();
            button_viewhistory.Enabled = true;
            button_deletevisitor.Enabled = true;
            button4.Enabled = true;
            button_checkin.Enabled = true;
            button_deletevisitor.Enabled = true;
            button_updateinfo.Enabled = true;
            button_viewhistory.Enabled = true;
        }

        private void button_settings_Click(object sender, EventArgs e)
        {
            //opens settings panel
            if (panel_settings.Visible == true)
            {
                panel_settings.Hide();
            }
            else
            {
                panel_settings.Show();
            }
        }

        private void button_exportExcel(object sender, EventArgs e)
        {
            //exports data similar to orignal excel spreadsheet for pantry
            using (SaveFileDialog sfd = new SaveFileDialog() { Filter = "Excel Workbook|*.xlsx" })
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        using (XLWorkbook workbook = new XLWorkbook())
                        {
                            workbook.Worksheets.Add(DataAccess.ExportExcelFormat(), "pantry Excel Mastersheet");
                            workbook.SaveAs(sfd.FileName);
                        }
                        MessageBox.Show("Database successfully Exported");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
            }
        }

        private void button_exportdb(object sender, EventArgs e)
        {
            //exports to excel file in db format incase of new program change
            DialogResult = MessageBox.Show("This will download 6 Excel Sheets that make up the main portion of the database", "", MessageBoxButtons.OKCancel);
            if (DialogResult == DialogResult.OK)
            {
                using (SaveFileDialog sfd = new SaveFileDialog() { Filter = "Excel Workbook|*.xlsx" })
                {
                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        try
                        {
                            using (XLWorkbook workbook = new XLWorkbook())
                            {
                                workbook.Worksheets.Add(DataAccess.ExportVisitorsDB(), "pantry DB Visitors");
                                workbook.Worksheets.Add(DataAccess.ExportFormDB(), "pantry DB Form Dates");
                                workbook.Worksheets.Add(DataAccess.ExportvisitsDB(), "pantry DB Visits");
                                workbook.Worksheets.Add(DataAccess.ExportTotalVisitsDB(), "pantry DB TotalVisits");
                                workbook.Worksheets.Add(DataAccess.ExportPastVisitorsDB(), "pantry DB PastVisitors");
                                workbook.Worksheets.Add(DataAccess.ExportLastIDNumDB(), "pantry DB LastIDNum");

                                workbook.SaveAs(sfd.FileName);
                            }
                            MessageBox.Show("All files Downloaded Succesfully");
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }
                    }
                    else
                    {
                        MessageBox.Show("Export Canceled");
                        return;
                    }
                }
            }
            else
            {
                MessageBox.Show("Export Canceled");
            }
        }

        private void button_deletePreviousVisitor_click(object sender, EventArgs e)
        {
            int currentRow = dataGridView_visitorcount.CurrentCell.RowIndex;
            if (label_selectedvisitortodelete.Text != "")
            {
                if ((FormInput.InNewWorkDay == true) && (dataGridView_visitorcount.SelectedCells[0].Value.ToString() == FormInput.NewDayDate))
                {
                    DialogResult = MessageBox.Show($"Are you sure you want to delete {dataGridView_previousvisitors.SelectedCells[2].Value} {dataGridView_previousvisitors.SelectedCells[3].Value} from the history?", "Alert", MessageBoxButtons.YesNo);
                    if (DialogResult == DialogResult.Yes)
                    {
                        DataAccess.DeletePreviousVisitorNewDay();
                        label_VisitorCount.Text = DataAccess.SumOfVisitsNewDay().ToString();
                        dataGridView_previousvisitors.DataSource = null;
                        label_selectedvisitortodelete.Text = "";
                        dataGridView_previousvisitors.DataSource = DataAccess.GetPreviousVisitors();
                        dataGridView_visitors.DataSource = null;
                        dataGridView_visitors.DataSource = DataAccess.LoadVisitors();
                        dataGridView_visitors.Columns[4].Visible = false;
                        dataGridView_visitors.Columns[5].Visible = false;
                        dataGridView_visitors.Columns[6].Visible = false;
                        if (dataGridView_previousvisitors.DataSource != null)
                        {
                            dataGridView_previousvisitors.Columns[0].Visible = false;
                            dataGridView_previousvisitors.Columns[1].Visible = false;
                            dataGridView_previousvisitors.Columns[4].Visible = false;
                        }
                        dataGridView_visitorcount.DataSource = null;
                        dataGridView_visitorcount.DataSource = DataAccess.GetVisitorCount();
                    }
                }
                else if ((FormInput.InNewWorkDay == false) && (dataGridView_visitorcount.SelectedCells[0].Value.ToString() == DateTime.Today.ToString("MM/dd/yyyy")))
                {
                    DialogResult = MessageBox.Show($"Are you sure you want to delete {dataGridView_previousvisitors.SelectedCells[2].Value} {dataGridView_previousvisitors.SelectedCells[3].Value} from the history?", "Alert", MessageBoxButtons.YesNo);
                    if (DialogResult == DialogResult.Yes)
                    {
                        DataAccess.DeletePreviousVisitor();
                        label_VisitorCount.Text = DataAccess.SumOfVisits().ToString();

                        dataGridView_previousvisitors.DataSource = null;
                        label_selectedvisitortodelete.Text = "";
                        dataGridView_previousvisitors.DataSource = DataAccess.GetPreviousVisitors();
                        dataGridView_visitors.DataSource = null;
                        dataGridView_visitors.DataSource = DataAccess.LoadVisitors();
                        dataGridView_visitors.Columns[4].Visible = false;
                        dataGridView_visitors.Columns[5].Visible = false;
                        dataGridView_visitors.Columns[6].Visible = false;
                        if (dataGridView_previousvisitors.DataSource != null)
                        {
                            dataGridView_previousvisitors.Columns[0].Visible = false;
                            dataGridView_previousvisitors.Columns[1].Visible = false;
                            dataGridView_previousvisitors.Columns[4].Visible = false;
                        }
                        dataGridView_visitorcount.DataSource = null;
                        dataGridView_visitorcount.DataSource = DataAccess.GetVisitorCount();
                    }
                }
                else
                {
                    MessageBox.Show("Unable to delete visitors from previous days");
                }
            }
            else
            {
                MessageBox.Show("Please select a visitor to delete");
            }
            dataGridView_visitorcount.CurrentCell = dataGridView_visitorcount.Rows[currentRow].Cells[0];
        }

        private void dataGridView_previousvisitors_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            GetVisitor.systemNo = Convert.ToInt32(dataGridView_previousvisitors.SelectedCells[4].Value);
            label_selectedvisitortodelete.Text = $"{dataGridView_previousvisitors.SelectedCells[2].Value} {dataGridView_previousvisitors.SelectedCells[3].Value}";
        }

        private void textBox_searchCriteria_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                SearchButton_Click(sender, e);
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        }

        private void button_AddNewDay(object sender, EventArgs e)
        {
            DialogResult = MessageBox.Show("Would you like to work under another business day?", "Alert", MessageBoxButtons.YesNo);
            if (DialogResult == DialogResult.Yes)
            {
                panel_newdaydate.Show();
                panel_counthistory.Hide();
                panel_previousvisitors.Hide();
            }
        }

        private void button_exitnewdayentry(object sender, EventArgs e)
        {
            panel_newdaydate.Hide();
            textBox_newdateentry.Text = "";
            panel_previousvisitors.Show();
            panel_counthistory.Show();
            panel_counthistory.Show();
            button_checkin.Enabled = true;
            button_updateinfo.Enabled = true;
            button_deletevisitor.Enabled = true;
            button_viewhistory.Enabled = true;

        }

        private void button_beginnewday_Click(object sender, EventArgs e)
        {
            if (DateTime.TryParseExact(textBox_newdateentry.Text.Trim(), "MM/dd/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out _))
            {
                if (Convert.ToDateTime(textBox_newdateentry.Text) < DateTime.Now.AddDays(-1))
                {
                    FormInput.NewDayDate = textBox_newdateentry.Text;
                    label_newworkdate.Text = FormInput.NewDayDate;
                    MessageBox.Show($"Now working in {textBox_newdateentry.Text}. Please be sure to exit the new day before working in the current day");
                    FormInput.InNewWorkDay = true;
                    if (DataAccess.SumOfVisitsNewDay() >= 0)
                    {
                        FormInput.visitorCount = DataAccess.SumOfVisitsNewDay();
                    }
                    else
                    {
                        FormInput.visitorCount = 0;
                    }
                    label_VisitorCount.Text = FormInput.visitorCount.ToString();
                    label_lastvisitorrecorded.Text = "";
                    panel_newworkdayalert.Show();
                    panel_newdaydate.Hide();
                    button_addDay.Hide();
                    panel_searchresults.Show();
                    button_checkin.Enabled = true;
                    button_updateinfo.Enabled = true;
                    button_deletevisitor.Enabled = true;
                    button_viewhistory.Enabled = true;
                    button_settings.Enabled = false;
                    label_datetoday.Hide();
                }
                else
                {
                    MessageBox.Show("Please enter a previous date to work in");
                }
            }
            else
            {
                MessageBox.Show("Invalid date format. Please enter the date in (MM/DD/YYYY) format");
            }
        }

        private void button_exitworkday_Click(object sender, EventArgs e)
        {
            DialogResult = MessageBox.Show($"Are you finished working in {FormInput.NewDayDate}?", "Alert", MessageBoxButtons.YesNo);
            if (DialogResult == DialogResult.Yes)
            {
                FormInput.InNewWorkDay = false;
                if (DataAccess.SumOfVisits() >= 0)
                {
                    FormInput.visitorCount = DataAccess.SumOfVisits();
                }
                else
                {
                    FormInput.visitorCount = 0;
                }
                label_VisitorCount.Text = FormInput.visitorCount.ToString();
                label_lastvisitorrecorded.Text = "";
                button_settings.Enabled = true;
                panel_newworkdayalert.Hide();
                button_addDay.Show();
                FormInput.NewDayDate = "";
                label_datetoday.Show();
            }
        }
    }
}
