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
    public partial class frmEvaluateWorkpackage : Form
    {
        public frmEvaluateWorkpackage()
        {
            InitializeComponent();
        }

        private void btnBackToEvaluate_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnEvalConfirmProjectComp_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Confirm Project Completion'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            EvaluateDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnEvalDeCommissionProject_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%De-Commissioning A Project%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            EvaluateDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnEvalEvalProjectGov_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Evaluate Project Governance and Operational Delivery.'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            EvaluateDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnEvalIdClosureAction_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Identify Closure Actions%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            EvaluateDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnEvalEvalTechDel_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Evaluate Technical Delivery.'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            EvaluateDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnEvalUndertakeClosureActions_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Underatke Closure Actions%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            EvaluateDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnEvalAuditBusinessPlan_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Audit Business Plan Benefits'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            EvaluateDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnEvalIdFollowOnActions_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Identifying Follow-On Actions'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            EvaluateDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }
    }
}
