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
    public partial class frmOpportunityScreeningWorkpackages : Form
    {
        public frmOpportunityScreeningWorkpackages()
        {
            InitializeComponent();
        }

        private void btnIdOppsForDev_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Screen Opportunities for strategic fit'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            OpportunityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnBackToOppScrn_Click(object sender, EventArgs e)
        {
            this.Hide();
        }

        private void btnIdNeedsOfStakeholders_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Identify Needs of Stakeholders'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            OpportunityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnScrnOppForStratFit_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Screen Opportunities for strategic fit'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            OpportunityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnReviewCollectSuppData_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Review and collect supporting Data'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            OpportunityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnDevPre_FeasabilityPlan_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Develop Pre-Feasibility Plan'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            OpportunityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnPriorOppApplyingScoringTool_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Prioritise Opportunities applying scoring tool'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            OpportunityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnSubmitOppTInvestPort_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Submit Opportunities to Investment Portfolio Process'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            OpportunityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnConStage_Click(object sender, EventArgs e)
        {
            
        }
    }
}
