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
    public partial class frmProjectManagementStream : Form
    {
        public frmProjectManagementStream()
        {
            InitializeComponent();
        }

        private void btnBackToProjectManagementStream_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnPWSProjectWorkStreamPlanNextSt_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Stage Planning%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ProjectManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnPWSProjectWorkStreamAquireTask_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Stage Planning%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ProjectManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnPWSProjectWorkStreamCreateInitDraft_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Project Implementation Plan'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ProjectManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnPWSProjectWorkStreamCreateFinalPPlan_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Project Execution Strategy'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ProjectManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnPWSProjectWorkStreamStartingUp_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Project Start Up Definition:%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ProjectManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnPWSProjectWorkStreamIdentifyProject_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Develop a Business Case:%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ProjectManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnPWSProjectWorkStreamCreateVariousExe_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Create the various execution plans : Schedule Project Work'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ProjectManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnPWSProjectWorkStreamArchiveProjectOutputs_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Consolidate Documentation'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ProjectManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnPWSProjectWorkStreamIntegratePlans_Click(object sender, EventArgs e)
        {

        }

        private void btnPWSProjectWorkStreamPlanNextStage_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Stage Planning%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ProjectManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnPWSProjectWorkStreamMonitorControlProject_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Stage Management%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ProjectManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }
    }
}
