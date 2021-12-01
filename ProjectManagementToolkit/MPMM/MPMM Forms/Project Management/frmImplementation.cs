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
    public partial class frmImplementation : Form
    {
        public frmImplementation()
        {
            InitializeComponent();
        }

        private void btnBackToImplementation_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnImpWSitePrepAndAccess_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Site Preparation and access'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ImplementationDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnImpWBuildConstructDeliverable_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Build Construct%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ImplementationDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnImpWManageDetailDesign_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Product delivery management%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ImplementationDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnImpWProductDeliveryManagement_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Product delivery management%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ImplementationDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnImpWConductTests_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Conduct tests%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ImplementationDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnImpWAcceptingDeliverables_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Accepting Deliverables%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ImplementationDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnImpWEstOperationalReadiness_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Establish Operational Readiness'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ImplementationDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnImpWDeploymentPlanning_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Deployment Planning'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ImplementationDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnImpWHandOver_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Hand Over / Partial Hand Over'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ImplementationDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnImpWDeploySystem_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Deploy system%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ImplementationDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnImpWStartUpAndCommissioning_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Start-Up and Commissioning'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ImplementationDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnImpWTransferSollution_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Transfer sollution%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ImplementationDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnImpWConcludeStage_Click(object sender, EventArgs e)
        {

        }
    }
}
