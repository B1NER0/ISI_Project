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
    public partial class frmBenefitsRealisationStream : Form
    {
        public frmBenefitsRealisationStream()
        {
            InitializeComponent();
        }

        private void btnBackToPortfolioManagementStream_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnDefineBenfits_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Compile Business Case'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            BenefitsRealisationStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnIdentifyEarlyBenefits_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Compile Business Case'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            BenefitsRealisationStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnPrepareBenefits_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Develop a Business Case: Identify'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            BenefitsRealisationStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnIdentifyBenefits_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Develop a Business Case: Identify'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            BenefitsRealisationStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnUpdatePlan_Click(object sender, EventArgs e)
        {

        }

        private void btnMonitorBenefit_Click(object sender, EventArgs e)
        {

        }

        private void btnConductPlan_Click(object sender, EventArgs e)
        {

        }

        private void btnReviewPlan_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Audit business plan benefit realisation'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            BenefitsRealisationStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnProcessStageReport_Click(object sender, EventArgs e)
        {

        }
    }
}
