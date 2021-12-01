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
    public partial class frmFeasabilityWorkpackages : Form
    {
        public frmFeasabilityWorkpackages()
        {
            InitializeComponent();
        }

        private void btnFeasBackToFeasability_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnFeasProjectStartUp_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Project Start Up Definition%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            FeasabilityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnFeasDevProjectCharter_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Develop Project Charter:%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            FeasabilityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnFeasStructDCO_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Structure DCO & PMO Contract'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            FeasabilityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnFeasEstablishProjectManSupp_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Establish Project Management Support'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            FeasabilityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnFeasEstPMOParty_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Establish PMO & 3rd Party Structures'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            FeasabilityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnFeasEstClientOfficeManSupp_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Establish Client Office Management Support'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            FeasabilityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnFeasAssLegalReq_Click(object sender, EventArgs e)
        {

        }

        private void btnFeasComplieBusinessCase_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Compile Business Case'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            FeasabilityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnFeasPerformBasicDesign_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Perform Basic Design'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            FeasabilityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnFeasPreliminarySiteSelect_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Preliminary Site Selection'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            FeasabilityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnFeasEnvirImpactAss_Click(object sender, EventArgs e)
        {

        }

        private void btnFeasFinaliseSiteSelection_Click(object sender, EventArgs e)
        {

        }

        private void btnFeasExeEIAReg_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Execute EIA, Regulatory and Legal'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            FeasabilityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnFeasUndertakeFeasStudy_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Undertake a feasabillity study:%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            FeasabilityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnFeasDevContractStrat_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Develop Contracting Strategy'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            FeasabilityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnFeasProduceURS_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Produce (URS) User Requirement Specs.'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            FeasabilityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnFeasDealStructBankableReport_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Deal Structuring and Bankable Report'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            FeasabilityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }
    }
}
