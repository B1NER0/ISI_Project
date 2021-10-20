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
    public partial class frmProcurementStream : Form
    {
        public frmProcurementStream()
        {
            InitializeComponent();
        }

        private void btnBackToPortfolioManagementStream_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnRFPResources_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Create the various execution plans: Create a procurement plan'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ProcurementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnEvaluateRFP_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Create the various execution plans: Create a procurement plan'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ProcurementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnRFPPrototypes_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Create the various execution plans: Create a procurement plan'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ProcurementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnGetTender_Click(object sender, EventArgs e)
        {

        }

        private void btnIssueRFPs_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Create the various execution plans: Create a procurement plan'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ProcurementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnEvaluate_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Create the various execution plans: Create a procurement plan'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ProcurementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnAwardContracts_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Create the various execution plans: Create a procurement plan'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ProcurementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnFeedbackTender_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Create the various execution plans: Create a procurement plan'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ProcurementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnConclude_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Create the various execution plans: Create a procurement plan'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ProcurementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnIdentifyContracts_Click(object sender, EventArgs e)
        {

        }

        private void btnEvaluatePerformance_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Audit business plan benefit realisation'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ProcurementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnMonitorCompliance_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Evaluate Technical Delivery.'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ProcurementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnProcessStage_Click(object sender, EventArgs e)
        {

        }

        private void btnCloseContracts_Click(object sender, EventArgs e)
        {

        }
    }
}
