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
    public partial class frmPLSM_Workstreams : Form
    {
        public frmPLSM_Workstreams()
        {
            InitializeComponent();
        }

        private void btnWorkstreamDirectingAProject_Click(object sender, EventArgs e)
        {
            frmPortfolioManagementStream frmPMStream = new frmPortfolioManagementStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Directing a project%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPMStream.PortfolioManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPMStream.Show();
        }

        private void btnWorkstreamScreenOppProblems_Click(object sender, EventArgs e)
        {
            frmPortfolioManagementStream frmPMStream = new frmPortfolioManagementStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Screen, prioritise and discard opportunities%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPMStream.PortfolioManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPMStream.Show();
        }

        private void btnWorkstreamFundsAcquisitionProcess_Click(object sender, EventArgs e)
        {
            frmPortfolioManagementStream frmPMStream = new frmPortfolioManagementStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Submit Opportunities to Investment Portfolio Process'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPMStream.PortfolioManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPMStream.Show();
        }

        private void btnWorkstreamIdBusinessRisks_Click(object sender, EventArgs e)
        {
            frmPortfolioManagementStream frmPMStream = new frmPortfolioManagementStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Develop a Business Case:%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPMStream.PortfolioManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPMStream.Show();
        }

        private void btnWorkstreamPriorOpportunitiesProblems_Click(object sender, EventArgs e)
        {
            frmPortfolioManagementStream frmPMStream = new frmPortfolioManagementStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Prioritise Opportunities applying scoring tool'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPMStream.PortfolioManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPMStream.Show();
        }

        private void btnWorkstreamDefineHighLevelBusinessBenf_Click(object sender, EventArgs e)
        {
            frmPortfolioManagementStream frmPMStream = new frmPortfolioManagementStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Develop a Business Case:%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPMStream.PortfolioManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPMStream.Show();
            
        }

        private void btnWorkstreamRegisterOpportunitiesProb_Click(object sender, EventArgs e)
        {
            frmPortfolioManagementStream frmPMStream = new frmPortfolioManagementStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Screen, prioritise and discard opportunities'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPMStream.PortfolioManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPMStream.Show();
        }

        private void btnWorkstreamConfirmBusinessBenefits_Click(object sender, EventArgs e)
        {
            frmPortfolioManagementStream frmPMStream = new frmPortfolioManagementStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Audit business plan benefit realisation'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPMStream.PortfolioManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPMStream.Show();
        }

        private void btnWorkstreamConductBenefitRealisation_Click(object sender, EventArgs e)
        {
            frmPortfolioManagementStream frmPMStream = new frmPortfolioManagementStream();

            frmPMStream.Show();
        }

        private void btnWorkstreamMonitorBenefitRealisation_Click(object sender, EventArgs e)
        {
            frmPortfolioManagementStream frmPMStream = new frmPortfolioManagementStream();

            frmPMStream.Show();
        }

        private void btnWorkstreamMonitorGovernanceCompliance_Click(object sender, EventArgs e)
        {
            frmPortfolioManagementStream frmPMStream = new frmPortfolioManagementStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Evaluate Project Governance and Operational Delivery.'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPMStream.PortfolioManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPMStream.Show();
        }

        private void btnWorkstreamAuditBenefitRealisation_Click(object sender, EventArgs e)
        {
            frmPortfolioManagementStream frmPMStream = new frmPortfolioManagementStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Audit Business Plan Benefits'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPMStream.PortfolioManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPMStream.Show();
        }

        private void btnWorkstreamProcessStageReport_Click(object sender, EventArgs e)
        {
            frmPortfolioManagementStream frmPMStream = new frmPortfolioManagementStream();

            frmPMStream.Show();
        }
    }
}
