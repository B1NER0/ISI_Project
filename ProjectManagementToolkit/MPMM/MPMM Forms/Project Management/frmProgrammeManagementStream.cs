using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.OleDb;
using Microsoft.Office.Interop;

namespace ProjectManagementToolkit.MPMM.MPMM_Forms.Project_Management
{
    public partial class frmProgrammeManagementStream : Form
    {
        public frmProgrammeManagementStream()
        {
            InitializeComponent();
        }

        private void btnBackToPlmsProg_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnPROGProgrammeWorkStreamPlanRemainingStage_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Stage Planning%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ProgrammeManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnPROGProgrammeWorkStreamEstimateFunding_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Submit Opportunities to Investment Portfolio Process'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ProgrammeManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnPROGProgrammeWorkStreamInitProject_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Initiate high-level project development planning'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ProgrammeManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnPROGProgrammeWorkStreamDefineProject_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Construct a preliminary Business Case and Risk Assessment'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ProgrammeManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnPROGProgrammeWorkStreamFundsAllocation_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Submit Opportunities to Investment Portfolio Process'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ProgrammeManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnPROGProgrammeWorkStreamDefineInvestment_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Develop a Business Case:%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ProgrammeManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnPROGProgrammeWorkStreamPrepareSupporting_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Compile Business Case'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ProgrammeManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnPROGProgrammeWorkStreamAssessProject_Click(object sender, EventArgs e)
        {

        }

        private void btnPROGProgrammeWorkStreamConfirmProject_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Submit Opportunities to Investment Portfolio Process'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ProgrammeManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnPROGProgrammeWorkStreamManagingStageBoundaries_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Stage Management:%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ProgrammeManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnPROGProgrammeWorkStreamMonitorControlPerformance_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Stage Management:%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ProgrammeManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnPROGProgrammeWorkStreamMonitorProgramme_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Stage Management:%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ProgrammeManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnPROGProgrammeWorkStreamReviewBusinessCase_Click(object sender, EventArgs e)
        {

        }

        private void btnPROGProgrammeWorkStreamRecordProject_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Close-Out Project'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ProgrammeManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnPROGProgrammeWorkStreamCloseProject_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE 'De-Commissioning A Project%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ProgrammeManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnPROGProgrammeWorkStreamAuditBenefit_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Audit Business Plan Benefits'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ProgrammeManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }
    }
}
