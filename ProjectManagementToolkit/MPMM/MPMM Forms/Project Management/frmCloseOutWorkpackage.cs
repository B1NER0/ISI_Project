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
    public partial class frmCloseOutWorkpackage : Form
    {
        public frmCloseOutWorkpackage()
        {
            InitializeComponent();
        }

        private void btnCloseOutCloseOutProject_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Close-Out Project'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            CloseOutDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnCloseOutInitiatePlantLifecycle_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Initiate Plant Lifecycle Plan - O&M'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            CloseOutDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnCloseOutConcludeQualityAssurance_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Conclude quality assurance'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            CloseOutDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnCloseOutConcludeProjectMarketing_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Conclude project marketing'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            CloseOutDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnCloseOutConcludeBusinessProcessModeling_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Conclude business process modeling'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            CloseOutDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnCloseOutConcludeRiskManagement_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Conclude risk management'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            CloseOutDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnCloseOutConcludeInfoSecurityAndDataProtection_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Conclude information security and data protection'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            CloseOutDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnCloseOutConcludeConfigManagement_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Conclude configuration management'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            CloseOutDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnCloseOutDeCommissioningProject_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%De-Commissioning A Project%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            CloseOutDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnBackToCloseOut_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
