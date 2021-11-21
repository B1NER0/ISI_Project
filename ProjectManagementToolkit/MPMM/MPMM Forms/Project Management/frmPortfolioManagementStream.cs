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
    public partial class frmPortfolioManagementStream : Form
    {
        public frmPortfolioManagementStream()
        {
            InitializeComponent();
        }

        private void btnBackToPortfolioManagementStream_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnWorkstreamPortDirectingAProject_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Directing a project%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            PortfolioManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnWorkstreamPortScreenOppProblems_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Screen, prioritise and discard opportunities%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            PortfolioManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnWorkstreamPortFundsAcquisitionProcess_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Submit Opportunities to Investment Portfolio Process'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            PortfolioManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnWorkstreamPortIdBusinessRisks_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Develop a Business Case:%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            PortfolioManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnWorkstreamPortPriorOpportunitiesProblems_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Prioritise Opportunities applying scoring tool'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            PortfolioManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnWorkstreamPortDefineHighLevelBusinessBenf_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Develop a Business Case:%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            PortfolioManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnWorkstreamPortRegisterOpportunitiesProb_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Screen, prioritise and discard opportunities%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            PortfolioManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnWorkstreamPortConfirmBusinessBenefits_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Audit business plan benefit realisation'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            PortfolioManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnWorkstreamPortConductBenefitRealisation_Click(object sender, EventArgs e)
        {

        }

        private void btnWorkstreamPortMonitorBenefitRealisation_Click(object sender, EventArgs e)
        {

        }

        private void btnWorkstreamPortMonitorGovernanceCompliance_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Evaluate Project Governance and Operational Delivery.'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            PortfolioManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnWorkstreamPortAuditBenefitRealisation_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Audit Business Plan Benefits'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            PortfolioManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnWorkstreamPortProcessStageReport_Click(object sender, EventArgs e)
        {

        }
    }
}
