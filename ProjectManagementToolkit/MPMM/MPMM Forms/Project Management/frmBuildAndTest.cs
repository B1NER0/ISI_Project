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
    public partial class frmBuildAndTest : Form
    {
        public frmBuildAndTest()
        {
            InitializeComponent();
        }

        private void btnBackToBuildAndTest_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnBuildTestWBuildExecute_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Build  Execute'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            BuildAndTestDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnBuildTestWConductTest_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Conduct tests%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            BuildAndTestDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnBuildTestWAcceptingDeliverables_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Accepting Deliverables%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            BuildAndTestDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnBuildTestWDeploySystem_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Deploy system%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            BuildAndTestDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnBuildTestWDeployPlanning_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Deployment Planning'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            BuildAndTestDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnBuildTestWDevelopTrainingConcept_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Develop training concept'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            BuildAndTestDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnBuildTestWDraftTrainingDoc_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Draft training documentation'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            BuildAndTestDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnBuildTestWTrainUsers_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Train users'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            BuildAndTestDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnBuildTestWImplementProtectionMeasures_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Implement protection measures'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            BuildAndTestDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnBuildTestWStartUpAndCommissioning_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Start-Up and Commissioning'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            BuildAndTestDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnBuildTestWHandOver_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Hand Over / Partial Hand Over'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            BuildAndTestDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnBuildTestWTransferSollution_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Transfer sollution%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            BuildAndTestDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnBuildTestWConsolidateDoc_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Consolidate Documentation'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            BuildAndTestDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }
    }
}
