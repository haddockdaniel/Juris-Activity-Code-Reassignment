using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Globalization;
using Gizmox.Controls;
using JDataEngine;
using JurisAuthenticator;
using JurisUtilityBase.Properties;
using System.Data.OleDb;

namespace JurisUtilityBase
{
    public partial class UtilityBaseMain : Form
    {
        #region Private  members

        private JurisUtility _jurisUtility;

        #endregion

        #region Public properties

        public string CompanyCode { get; set; }

        public string JurisDbName { get; set; }

        public string JBillsDbName { get; set; }

        public string fromActCode = "";

        public string toActCode = "";

        #endregion

        #region Constructor

        public UtilityBaseMain()
        {
            InitializeComponent();
            _jurisUtility = new JurisUtility();
        }

        #endregion

        #region Public methods

        public void LoadCompanies()
        {
            var companies = _jurisUtility.Companies.Cast<object>().Cast<Instance>().ToList();
//            listBoxCompanies.SelectedIndexChanged -= listBoxCompanies_SelectedIndexChanged;
            listBoxCompanies.ValueMember = "Code";
            listBoxCompanies.DisplayMember = "Key";
            listBoxCompanies.DataSource = companies;
//            listBoxCompanies.SelectedIndexChanged += listBoxCompanies_SelectedIndexChanged;
            var defaultCompany = companies.FirstOrDefault(c => c.Default == Instance.JurisDefaultCompany.jdcJuris);
            if (companies.Count > 0)
            {
                listBoxCompanies.SelectedItem = defaultCompany ?? companies[0];
            }
        }

        #endregion

        #region MainForm events

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void listBoxCompanies_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_jurisUtility.DbOpen)
            {
                _jurisUtility.CloseDatabase();
            }
            CompanyCode = "Company" + listBoxCompanies.SelectedValue;
            _jurisUtility.SetInstance(CompanyCode);
            JurisDbName = _jurisUtility.Company.DatabaseName;
            JBillsDbName = "JBills" + _jurisUtility.Company.Code;
            _jurisUtility.OpenDatabase();
            if (_jurisUtility.DbOpen)
            {
                ///GetFieldLengths();
            }

            try
            {
                string taskCode;
                cbFrom.ClearItems();
                string SQLTkpr = "select ActyCdCode + ' ' + ActyCdDesc as actcode from ActivityCode";
                DataSet myRSTkpr = _jurisUtility.RecordsetFromSQL(SQLTkpr);

                if (myRSTkpr.Tables[0].Rows.Count == 0)
                    cbFrom.SelectedIndex = 0;
                else
                {
                    foreach (DataTable table in myRSTkpr.Tables)
                    {

                        foreach (DataRow dr in table.Rows)
                        {
                            taskCode = dr["actcode"].ToString();
                            cbFrom.Items.Add(taskCode);
                        }
                    }

                }

                string TkprIndex2;
                cbTo.ClearItems();
                string SQLTkpr2 = "select ActyCdCode + ' ' + ActyCdDesc as actcode from ActivityCode";
                DataSet myRSTkpr2 = _jurisUtility.RecordsetFromSQL(SQLTkpr2);


                if (myRSTkpr2.Tables[0].Rows.Count == 0)
                    cbTo.SelectedIndex = 0;
                else
                {
                    foreach (DataTable table in myRSTkpr2.Tables)
                    {

                        foreach (DataRow dr in table.Rows)
                        {
                            TkprIndex2 = dr["actcode"].ToString();
                            cbTo.Items.Add(TkprIndex2);
                        }
                    }

                }


            }
            catch (Exception ex1)
            {
                MessageBox.Show("There was an error reading the activity codes." + "\r\n" + "This usually means this database has none", "Activity Code Read Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }



        #endregion

        #region Private methods

        private void DoDaFix()
        {
            // Enter your SQL code here
            // To run a T-SQL statement with no results, int RecordsAffected = _jurisUtility.ExecuteNonQueryCommand(0, SQL);
            // To get an ADODB.Recordset, ADODB.Recordset myRS = _jurisUtility.RecordsetFromSQL(SQL);
            DialogResult result = MessageBox.Show("This will change all Activity code references from " + fromActCode + "\r\n" + "to " + toActCode + ". Are you sure?","Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == System.Windows.Forms.DialogResult.Yes)
            {
                UpdateStatus("Updating AR Fee Task and Billed Time...", 1, 8);
                string SQL = "Select * from [ARFTaskAlloc] where ARFTActivityCd = '" + fromActCode + "'";
                //one record at a time. If exists with same data but new act code, update and conbine. If not, update
                
                DataSet fromArf = _jurisUtility.RecordsetFromSQL(SQL); // no matches so who cares
                if (fromArf.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in fromArf.Tables[0].Rows)
                    {
                        CompositKeyRecord record = new CompositKeyRecord();
                        record = getARFTaskAllocDataSet(dr);
                        SQL = "Select * from [ARFTaskAlloc] where ARFTActivityCd = '" + toActCode + "' and [ARFTBillNbr] = " + record.BillNo + " and [ARFTMatter] =" + record.Matter + " and [ARFTTkpr] = " + record.Tkpr + " and [ARFTTaskCd] = '" + record.TaskCode + "'";
                        //one record at a time. If exists with same data but new act code, update and conbine. If not, update

                        DataSet matches = _jurisUtility.RecordsetFromSQL(SQL);
                        if (matches.Tables[0].Rows.Count > 0) //we need to combine because key already exists
                        {
                            double WorkHrsBld = Double.Parse(matches.Tables[0].Rows[0]["ARFTWorkedHrsBld"].ToString().Trim());
                            double HrsBld = Double.Parse(matches.Tables[0].Rows[0]["ARFTHrsBld"].ToString().Trim());
                            double StdValueBld = Double.Parse(matches.Tables[0].Rows[0]["ARFTStdValueBld"].ToString().Trim());
                            double ActualValueBld = Double.Parse(matches.Tables[0].Rows[0]["ARFTActualValueBld"].ToString().Trim());
                            double ActualAmtBld = Double.Parse(matches.Tables[0].Rows[0]["ARFTActualAmtBld"].ToString().Trim());
                            double Rcvd = Double.Parse(matches.Tables[0].Rows[0]["ARFTRcvd"].ToString().Trim());
                            double Adj = Double.Parse(matches.Tables[0].Rows[0]["ARFTAdj"].ToString().Trim());
                            double Pend = Double.Parse(matches.Tables[0].Rows[0]["ARFTPend"].ToString().Trim());
                            //update record with correct task code
                            SQL = "update ARFTaskAlloc set [ARFTWorkedHrsBld] = Cast(" + WorkHrsBld + " + " + record.WorkHrsBld + " as decimal(12,2)),[ARFTHrsBld] = Cast(" + HrsBld + " + " + record.HrsBld + " as decimal(12,2)),[ARFTStdValueBld] = Cast(" + StdValueBld + " + " + record.StdValueBld + " as decimal(12,2)),[ARFTActualValueBld] = Cast(" + ActualValueBld + " + " + record.ActualValueBld + " as decimal(12,2)),[ARFTActualAmtBld] = Cast(" + ActualAmtBld + " + " + record.ActualAmtBld + " as decimal(12,2)),[ARFTRcvd] = Cast(" +Rcvd + " + " + record.Rcvd + " as decimal(12,2)),[ARFTAdj] = Cast(" +Adj + " + " + record.Adj + " as decimal(12,2)),[ARFTPend] = Cast(" + Pend + " + " + record.Pend + " as decimal(12,2)) where ARFTActivityCd = '" + toActCode + "' and [ARFTBillNbr] = " + record.BillNo + " and [ARFTMatter] =" + record.Matter + " and [ARFTTkpr] = " + record.Tkpr + " and [ARFTTaskCd] = '" + record.TaskCode + "'";
                            _jurisUtility.ExecuteNonQueryCommand(0, SQL);
                            //delete record with old task code
                            SQL = "delete from ARFTaskAlloc where ARFTActivityCd = '" + fromActCode + "' and [ARFTBillNbr] = " + record.BillNo + " and [ARFTMatter] =" + record.Matter + " and [ARFTTkpr] = " + record.Tkpr + " and [ARFTTaskCd] = '" + record.TaskCode + "'";
                            _jurisUtility.ExecuteNonQueryCommand(0, SQL);
                        }
                        else //no match so we can just update
                        {
                            SQL = "update ARFTaskAlloc set ARFTActivityCd ='" + toActCode + "' where ARFTActivityCd = '" + fromActCode + "' and [ARFTBillNbr] = " + record.BillNo + " and [ARFTMatter] =" + record.Matter + " and [ARFTTkpr] = " + record.Tkpr + " and [ARFTTaskCd] = '" + record.TaskCode + "'";
                            _jurisUtility.ExecuteNonQueryCommand(0, SQL);
                        }

                    }
                }

                fromArf.Clear();
                


                SQL = "update BilledTime set BTActivityCd ='" + toActCode + "' where BTActivityCd = '" + fromActCode + "'";
                _jurisUtility.ExecuteNonQueryCommand(0, SQL);

                //fave and most recent
                SQL = "delete from ActivityCodeMostRecent where Code ='" + fromActCode + "'";
                _jurisUtility.ExecuteNonQueryCommand(0, SQL);

                SQL = "delete from ActivityCodeFavorite where Code ='" + fromActCode + "'";
                _jurisUtility.ExecuteNonQueryCommand(0, SQL);

                UpdateStatus("Updating Cash Rec Fee Alloc...", 2, 8);




                SQL = "Select * from [CRFeeAlloc] where CRFActivityCd = '" + fromActCode + "'";
                //one record at a time. If exists with same data but new act code, update and conbine. If not, update

                fromArf = _jurisUtility.RecordsetFromSQL(SQL); // no matches so who cares
                if (fromArf.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in fromArf.Tables[0].Rows)
                    {
                        CompositKeyRecord record = new CompositKeyRecord();
                        record = getCRFeeAllocDataSet(dr);
                        SQL = "Select * from [CRFeeAlloc] where CRFActivityCd = '" + toActCode + "' and [CRFBillNbr] = " + record.BillNo + " and [CRFMatter] =" + record.Matter + " and [CRFTkpr] = " + record.Tkpr + " and [CRFTaskCd] = '" + record.TaskCode + "' and CRFBatch = " + record.Batch + " and CRFRecNbr = " + record.Record;
                        //one record at a time. If exists with same data but new act code, update and conbine. If not, update

                        DataSet matches = _jurisUtility.RecordsetFromSQL(SQL);
                        if (matches.Tables[0].Rows.Count > 0) //we need to combine because key already exists
                        {
                            double PrePost = Double.Parse(matches.Tables[0].Rows[0]["CRFPrePost"].ToString().Trim());
                            double Amount = Double.Parse(matches.Tables[0].Rows[0]["CRFAmount"].ToString().Trim());
                            //update record with correct task code
                            SQL = "update CRFeeAlloc set [CRFPrePost] = Cast(" +PrePost + " + " + record.PrePost + " as decimal(12,2)),[CRFAmount] = Cast(" +Amount + " + " + record.Amount + " as decimal(12,2)) where CRFActivityCd = '" + fromActCode + "' and [CRFBillNbr] = " + record.BillNo + " and [CRFMatter] =" + record.Matter + " and [CRFTkpr] = " + record.Tkpr + " and [CRFTaskCd] = '" + record.TaskCode + "' and CRFBatch = " + record.Batch + " and CRFRecNbr = " + record.Record;
                            _jurisUtility.ExecuteNonQueryCommand(0, SQL);
                            //delete record with old task code
                            SQL = "delete from CRFeeAlloc where CRFActivityCd = '" + fromActCode + "' and [CRFBillNbr] = " + record.BillNo + " and [CRFMatter] =" + record.Matter + " and [CRFTkpr] = " + record.Tkpr + " and [CRFTaskCd] = '" + record.TaskCode + "' and CRFBatch = " + record.Batch + " and CRFRecNbr = " + record.Record;
                            _jurisUtility.ExecuteNonQueryCommand(0, SQL);
                        }
                        else //no match so we can just update
                        {
                            SQL = "update CRFeeAlloc set CRFActivityCd ='" + toActCode + "' where CRFActivityCd = '" + fromActCode + "' and [CRFBillNbr] = " + record.BillNo + " and [CRFMatter] =" + record.Matter + " and [CRFTkpr] = " + record.Tkpr + " and [CRFTaskCd] = '" + record.TaskCode + "' and CRFBatch = " + record.Batch + " and CRFRecNbr = " + record.Record;
                            _jurisUtility.ExecuteNonQueryCommand(0, SQL);
                        }

                    }
                }

                fromArf.Clear();
                UpdateStatus("Updating Fee Sum By Period...", 3, 8);



                SQL = "Select * from [FeeSumByPrd] where FSPActivityCd = '" + fromActCode + "'";
                //one record at a time. If exists with same data but new act code, update and conbine. If not, update

                fromArf = _jurisUtility.RecordsetFromSQL(SQL); // no matches so who cares
                if (fromArf.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in fromArf.Tables[0].Rows)
                    {
                        CompositKeyRecord record = new CompositKeyRecord();
                        record = getFeeSumByPrdDataSet(dr);
                        SQL = "Select * from [FeeSumByPrd] where FSPActivityCd = '" + toActCode + "' and [FSPMatter] =" + record.Matter + " and [FSPTkpr] = " + record.Tkpr + " and [FSPTaskCd] = '" + record.TaskCode + "' and FSPPrdYear = " + record.Year + " and FSPPrdNbr = " + record.Period;
                        //one record at a time. If exists with same data but new act code, update and conbine. If not, update

                        DataSet matches = _jurisUtility.RecordsetFromSQL(SQL);
                        if (matches.Tables[0].Rows.Count > 0) //we need to combine because key already exists
                        {
                            double WorkedHrsBld = Double.Parse(matches.Tables[0].Rows[0]["FSPWorkedHrsBld"].ToString().Trim());
                            double HrsBilled = Double.Parse(matches.Tables[0].Rows[0]["FSPHrsBilled"].ToString().Trim());
                            double FeeBldStdValue = Double.Parse(matches.Tables[0].Rows[0]["FSPFeeBldStdValue"].ToString().Trim());
                            double FeeBldActualValue = Double.Parse(matches.Tables[0].Rows[0]["FSPFeeBldActualValue"].ToString().Trim());
                            double FeeBldActualAmt = Double.Parse(matches.Tables[0].Rows[0]["FSPFeeBldActualAmt"].ToString().Trim());
                            double FeeReceived = Double.Parse(matches.Tables[0].Rows[0]["FSPFeeReceived"].ToString().Trim());
                            double FeeAdjusted  = Double.Parse(matches.Tables[0].Rows[0]["FSPFeeAdjusted"].ToString().Trim());
                            double WorkedHrsEntered = Double.Parse(matches.Tables[0].Rows[0]["FSPWorkedHrsEntered"].ToString().Trim());
                            double NonBilHrsEntered = Double.Parse(matches.Tables[0].Rows[0]["FSPNonBilHrsEntered"].ToString().Trim());
                            double BilHrsEntered = Double.Parse(matches.Tables[0].Rows[0]["FSPBilHrsEntered"].ToString().Trim());
                            double FeeEnteredStdValue = Double.Parse(matches.Tables[0].Rows[0]["FSPFeeEnteredStdValue"].ToString().Trim());
                            double FeeEnteredActualValue = Double.Parse(matches.Tables[0].Rows[0]["FSPFeeEnteredActualValue"].ToString().Trim());
                            //update record with correct task code
                            SQL = "update FeeSumByPrd set [FSPWorkedHrsEntered] = Cast(" + WorkedHrsEntered + " + " + record.WorkedHrsEntered + " as decimal(12,2)) ,[FSPNonBilHrsEntered] = Cast(" +NonBilHrsEntered + " + " + record.NonBilHrsEntered + " as decimal(12,2)) ,[FSPBilHrsEntered] = Cast(" +BilHrsEntered + " + " + record.BilHrsEntered + " as decimal(12,2)) ,[FSPFeeEnteredStdValue] = Cast(" +FeeEnteredStdValue + " + " + record.FeeEnteredStdValue + " as decimal(12,2)) ,[FSPFeeEnteredActualValue] = Cast(" + FeeEnteredActualValue + " + " + record.FeeEnteredActualValue + " as decimal(12,2)),[FSPWorkedHrsBld] = Cast(" + WorkedHrsBld + " + " + record.WorkHrsBld + " as decimal(12,2)) ,[FSPHrsBilled] = Cast(" + HrsBilled + " + " + record.HrsBld + " as decimal(12,2)) ,[FSPFeeBldStdValue] = Cast(" + FeeBldStdValue + " + " + record.StdValueBld + " as decimal(12,2)) ,[FSPFeeBldActualValue] = Cast(" + FeeBldActualValue + " + " + record.ActualValueBld + " as decimal(12,2)) ,[FSPFeeBldActualAmt] = Cast(" + FeeBldActualAmt + " + " + record.ActualAmtBld + " as decimal(12,2)) ,[FSPFeeReceived] = Cast(" + FeeReceived + " + " + record.Rcvd + " as decimal(12,2)) ,[FSPFeeAdjusted] = Cast(" + FeeAdjusted + " + " + record.Adj + " as decimal(12,2)) where FSPActivityCd = '" + fromActCode + "' and [FSPMatter] =" + record.Matter + " and [FSPTkpr] = " + record.Tkpr + " and [FSPTaskCd] = '" + record.TaskCode + "' and FSPPrdYear = " + record.Year + " and FSPPrdNbr = " + record.Period;
                            _jurisUtility.ExecuteNonQueryCommand(0, SQL);
                            //delete record with old task code
                            SQL = "delete from FeeSumByPrd where FSPActivityCd = '" + fromActCode + "' and [FSPMatter] =" + record.Matter + " and [FSPTkpr] = " + record.Tkpr + " and [FSPTaskCd] = '" + record.TaskCode + "' and FSPPrdYear = " + record.Year + " and FSPPrdNbr = " + record.Period;
                            _jurisUtility.ExecuteNonQueryCommand(0, SQL);
                        }
                        else //no match so we can just update
                        {
                            SQL = "update FeeSumByPrd set FSPActivityCd ='" + toActCode + "' where FSPActivityCd = '" + fromActCode + "' and [FSPMatter] =" + record.Matter + " and [FSPTkpr] = " + record.Tkpr + " and [FSPTaskCd] = '" + record.TaskCode + "' and FSPPrdYear = " + record.Year + " and FSPPrdNbr = " + record.Period;
                            _jurisUtility.ExecuteNonQueryCommand(0, SQL);
                        }

                    }
                }
                fromArf.Clear();
                UpdateStatus("Updating Matter Fee Budget...", 4, 8);




                SQL = "update MatterFeeBudget set MFBActivityCode ='" + toActCode + "' where MFBActivityCode = '" + fromActCode + "'";
                _jurisUtility.ExecuteNonQueryCommand(0, SQL);
                UpdateStatus("Updating Time Batch Detail...", 5, 8);

                SQL = "update TimeBatchDetail set TBDActivityCd ='" + toActCode + "' where TBDActivityCd = '" + fromActCode + "'";
                _jurisUtility.ExecuteNonQueryCommand(0, SQL);
                UpdateStatus("Updating Time Entries...", 6, 8);

                SQL = "update TimeEntry set ActivityCode ='" + toActCode + "' where ActivityCode = '" + fromActCode + "'"; 
                _jurisUtility.ExecuteNonQueryCommand(0, SQL);
                UpdateStatus("Updating Unbilled Time...", 7, 8);

                SQL = "update UnbilledTime set UTActivityCd ='" + toActCode + "' where UTActivityCd = '" + fromActCode + "'";
                _jurisUtility.ExecuteNonQueryCommand(0, SQL);
                UpdateStatus("All tables updated.", 8, 8);

                //log tables?
                //xref?
                //fave and most recent?

                MessageBox.Show("The process is complete", "Finished", MessageBoxButtons.OK, MessageBoxIcon.None);
                toActCode = "";
                fromActCode = "";
                button1.Enabled = false;
            }
        }


        private CompositKeyRecord getARFTaskAllocDataSet(DataRow dr)
        {
            CompositKeyRecord rec = new CompositKeyRecord();
                rec.BillNo = dr["ARFTBillNbr"].ToString().Trim();
                rec.Matter = dr["ARFTMatter"].ToString().Trim();
                rec.Tkpr = dr["ARFTTkpr"].ToString().Trim();
                rec.TaskCode = dr["ARFTTaskCd"].ToString().Trim();
                rec.ActCode = dr["ARFTActivityCd"].ToString().Trim();
                rec.WorkHrsBld = Double.Parse(dr["ARFTWorkedHrsBld"].ToString().Trim());
                rec.HrsBld = Double.Parse(dr["ARFTHrsBld"].ToString().Trim());
                rec.StdValueBld = Double.Parse(dr["ARFTStdValueBld"].ToString().Trim());
                rec.ActualValueBld = Double.Parse(dr["ARFTActualValueBld"].ToString().Trim());
                rec.ActualAmtBld = Double.Parse(dr["ARFTActualAmtBld"].ToString().Trim());
                rec.Rcvd = Double.Parse(dr["ARFTRcvd"].ToString().Trim());
                rec.Adj = Double.Parse(dr["ARFTAdj"].ToString().Trim());
                rec.Pend = Double.Parse(dr["ARFTPend"].ToString().Trim());
            return rec;
        }

        private CompositKeyRecord getCRFeeAllocDataSet(DataRow dr)
        {
            CompositKeyRecord rec = new CompositKeyRecord();
            rec.BillNo = dr["CRFBillNbr"].ToString().Trim();
            rec.Matter = dr["CRFMatter"].ToString().Trim();
            rec.Tkpr = dr["CRFTkpr"].ToString().Trim();
            rec.TaskCode = dr["CRFTaskCd"].ToString().Trim();
            rec.ActCode = dr["CRFActivityCd"].ToString().Trim();
            rec.Batch = dr["CRFBatch"].ToString().Trim();
            rec.Record = dr["CRFRecNbr"].ToString().Trim();
            rec.PrePost = Double.Parse(dr["CRFPrePost"].ToString().Trim());
            rec.Amount = Double.Parse(dr["CRFAmount"].ToString().Trim());
            return rec;
        }

        private CompositKeyRecord getFeeSumByPrdDataSet(DataRow dr)
        {
            CompositKeyRecord rec = new CompositKeyRecord();
            rec.Year = dr["FSPPrdYear"].ToString().Trim();
            rec.Matter = dr["FSPMatter"].ToString().Trim();
            rec.Tkpr = dr["FSPTkpr"].ToString().Trim();
            rec.TaskCode = dr["FSPTaskCd"].ToString().Trim();
            rec.ActCode = dr["FSPActivityCd"].ToString().Trim();
            rec.Period = dr["FSPPrdNbr"].ToString().Trim();
            rec.WorkHrsBld = Double.Parse(dr["FSPWorkedHrsBld"].ToString().Trim());
            rec.HrsBld = Double.Parse(dr["FSPHrsBilled"].ToString().Trim());
            rec.StdValueBld = Double.Parse(dr["FSPFeeBldStdValue"].ToString().Trim());
            rec.ActualValueBld = Double.Parse(dr["FSPFeeBldActualValue"].ToString().Trim());
            rec.ActualAmtBld = Double.Parse(dr["FSPFeeBldActualAmt"].ToString().Trim());
            rec.Rcvd = Double.Parse(dr["FSPFeeReceived"].ToString().Trim());
            rec.Adj = Double.Parse(dr["FSPFeeAdjusted"].ToString().Trim());
            rec.WorkedHrsEntered = Double.Parse(dr["FSPWorkedHrsEntered"].ToString().Trim());
            rec.NonBilHrsEntered = Double.Parse(dr["FSPNonBilHrsEntered"].ToString().Trim());
            rec.BilHrsEntered = Double.Parse(dr["FSPBilHrsEntered"].ToString().Trim());
            rec.FeeEnteredStdValue = Double.Parse(dr["FSPFeeEnteredStdValue"].ToString().Trim());
            rec.FeeEnteredActualValue = Double.Parse(dr["FSPFeeEnteredActualValue"].ToString().Trim());
            return rec;
        }



        private bool VerifyFirmName()
        {
            //    Dim SQL     As String
            //    Dim rsDB    As ADODB.Recordset
            //
            //    SQL = "SELECT CASE WHEN SpTxtValue LIKE '%firm name%' THEN 'Y' ELSE 'N' END AS Firm FROM SysParam WHERE SpName = 'FirmName'"
            //    Cmd.CommandText = SQL
            //    Set rsDB = Cmd.Execute
            //
            //    If rsDB!Firm = "Y" Then
            return true;
            //    Else
            //        VerifyFirmName = False
            //    End If

        }

        private bool FieldExistsInRS(DataSet ds, string fieldName)
        {

            foreach (DataColumn column in ds.Tables[0].Columns)
            {
                if (column.ColumnName.Equals(fieldName, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            return false;
        }


        private static bool IsDate(String date)
        {
            try
            {
                DateTime dt = DateTime.Parse(date);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool IsNumeric(object Expression)
        {
            double retNum;

            bool isNum = Double.TryParse(Convert.ToString(Expression), System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out retNum);
            return isNum; 
        }

        private void WriteLog(string comment)
        {
            var sql =
                string.Format("Insert Into UtilityLog(ULTimeStamp,ULWkStaUser,ULComment) Values('{0}','{1}', '{2}')",
                    DateTime.Now, GetComputerAndUser(), comment);
            _jurisUtility.ExecuteNonQueryCommand(0, sql);
        }

        private string GetComputerAndUser()
        {
            var computerName = Environment.MachineName;
            var windowsIdentity = System.Security.Principal.WindowsIdentity.GetCurrent();
            var userName = (windowsIdentity != null) ? windowsIdentity.Name : "Unknown";
            return computerName + "/" + userName;
        }

        /// <summary>
        /// Update status bar (text to display and step number of total completed)
        /// </summary>
        /// <param name="status">status text to display</param>
        /// <param name="step">steps completed</param>
        /// <param name="steps">total steps to be done</param>
        private void UpdateStatus(string status, long step, long steps)
        {
            labelCurrentStatus.Text = status;

            if (steps == 0)
            {
                progressBar.Value = 0;
                labelPercentComplete.Text = string.Empty;
            }
            else
            {
                double pctLong = Math.Round(((double)step/steps)*100.0);
                int percentage = (int)Math.Round(pctLong, 0);
                if ((percentage < 0) || (percentage > 100))
                {
                    progressBar.Value = 0;
                    labelPercentComplete.Text = string.Empty;
                }
                else
                {
                    progressBar.Value = percentage;
                    labelPercentComplete.Text = string.Format("{0} percent complete", percentage);
                }
            }
        }

        private void DeleteLog()
        {
            string AppDir = Path.GetDirectoryName(Application.ExecutablePath);
            string filePathName = Path.Combine(AppDir, "VoucherImportLog.txt");
            if (File.Exists(filePathName + ".ark5"))
            {
                File.Delete(filePathName + ".ark5");
            }
            if (File.Exists(filePathName + ".ark4"))
            {
                File.Copy(filePathName + ".ark4", filePathName + ".ark5");
                File.Delete(filePathName + ".ark4");
            }
            if (File.Exists(filePathName + ".ark3"))
            {
                File.Copy(filePathName + ".ark3", filePathName + ".ark4");
                File.Delete(filePathName + ".ark3");
            }
            if (File.Exists(filePathName + ".ark2"))
            {
                File.Copy(filePathName + ".ark2", filePathName + ".ark3");
                File.Delete(filePathName + ".ark2");
            }
            if (File.Exists(filePathName + ".ark1"))
            {
                File.Copy(filePathName + ".ark1", filePathName + ".ark2");
                File.Delete(filePathName + ".ark1");
            }
            if (File.Exists(filePathName ))
            {
                File.Copy(filePathName, filePathName + ".ark1");
                File.Delete(filePathName);
            }

        }

            

        private void LogFile(string LogLine)
        {
            string AppDir = Path.GetDirectoryName(Application.ExecutablePath);
            string filePathName = Path.Combine(AppDir, "VoucherImportLog.txt");
            using (StreamWriter sw = File.AppendText(filePathName))
            {
                sw.WriteLine(LogLine);
            }	
        }
        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            DoDaFix();
        }

        private void buttonReport_Click(object sender, EventArgs e)
        {

            System.Environment.Exit(0);
          
        }

        private string getReportSQL()
        {
            string reportSQL = "";
            //if matter and billing timekeeper
            if (true)
                reportSQL = "select Clicode, Clireportingname, Matcode, Matreportingname,empinitials as CurrentBillingTimekeeper, 'DEF' as NewBillingTimekeeper" +
                        " from matter" +
                        " inner join client on matclinbr=clisysnbr" +
                        " inner join billto on matbillto=billtosysnbr" +
                        " inner join employee on empsysnbr=billtobillingatty" +
                        " where empinitials<>'ABC'";


            //if matter and originating timekeeper
            else if (false)
                reportSQL = "select Clicode, Clireportingname, Matcode, Matreportingname,empinitials as CurrentOriginatingTimekeeper, 'DEF' as NewOriginatingTimekeeper" +
                    " from matter" +
                    " inner join client on matclinbr=clisysnbr" +
                    " inner join matorigatty on matsysnbr=morigmat" +
                    " inner join employee on empsysnbr=morigatty" +
                    " where empinitials<>'ABC'";


            return reportSQL;
        }

        private void cbFrom_SelectedIndexChanged(object sender, EventArgs e)
        {
            fromActCode = cbFrom.Text;
            fromActCode = fromActCode.Split(' ')[0];
            if (!String.IsNullOrEmpty(toActCode))
                button1.Enabled = true;
        }

        private void cbTo_SelectedIndexChanged(object sender, EventArgs e)
        {
            toActCode = cbTo.Text;
            toActCode = toActCode.Split(' ')[0];
            if (!String.IsNullOrEmpty(fromActCode))
                button1.Enabled = true;
        }


    }
}
