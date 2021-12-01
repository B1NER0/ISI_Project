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
    public partial class frmArchitectureStream : Form
    {
        public frmArchitectureStream()
        {
            InitializeComponent();
        }

        private void btnBackToPortfolioManagementStream_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnPreparePrefeasiblity_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Perform base infrastructure assessment'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ArchitectureStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnPrepareConceptual_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Perform base infrastructure assessment'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ArchitectureStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnPrepareDesign_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Perform base infrastructure assessment'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ArchitectureStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnPreparePreContract_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Compile Business Case'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ArchitectureStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnPrepareDetailedDesign_Click(object sender, EventArgs e)
        {

        }

        private void btnPrepareTranferImplementation_Click(object sender, EventArgs e)
        {

        }
    }
}
