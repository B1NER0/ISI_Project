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
    public partial class frmDetailedDesign : Form
    {
        public frmDetailedDesign()
        {
            InitializeComponent();
        }

        private void btnDetailDesignWDevDeployment_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Developing deployment and operational concepts'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            DetailedDesignDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnBackToDetailedDesign_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnDetailDesignWDevProcessOrganizational_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Develop process and organizational concept%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            DetailedDesignDGV.DataSource = ds.Tables[0];
            myconnection.Close();

        }

        private void btnDetailDesignWDeployPlanning_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Deployment Planning'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            DetailedDesignDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnDetailDesignWImplementProcessOrgConcept_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Implement process and organizational concept%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            DetailedDesignDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnDetailDesignWActivateProcesses_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Activate processes and organization'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            DetailedDesignDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnDetailDesignWPrepareProcesses_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Prepare processes and organization'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            DetailedDesignDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnDetailDesignWBackUpData_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Back up data'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            DetailedDesignDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnDetailDesignWConcludeStage_Click(object sender, EventArgs e)
        {

        }
    }
}
