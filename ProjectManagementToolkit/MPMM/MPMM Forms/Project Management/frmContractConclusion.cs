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
    public partial class frmContractConclusion : Form
    {
        public frmContractConclusion()
        {
            InitializeComponent();
        }

        private void btnBackToContractConclusion_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnContractConcWCompileBalanceOfEnquiries_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Compile Balance of Enquiries'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ContractConcDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnContractConcWUpdateProjectImplementationPlan_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Update Project Implementation Plan'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ContractConcDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnContractConcWConcludeStage_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Conclude Stage%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ContractConcDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }
    }
}
