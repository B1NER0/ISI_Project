using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ProjectManagementToolkit.MPMM.MPMM_Document_Forms;
using ProjectManagementToolkit.MPMM.MPMM_Forms;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Excel;
using System.IO;

namespace ProjectManagementToolkit.MPMM.MPMM_Forms.Project_Management
{
    public partial class frmTemplates : Form
    {
        public frmTemplates()
        {
            InitializeComponent();
        }

        private void btnProjectInitianBusinessCase_Click(object sender, EventArgs e)
        {
            //string path = Directory.GetCurrentDirectory().ToString();
            //MessageBox.Show(path);
            var path = Path.Combine(Directory.GetCurrentDirectory(), @"Documents/Business Case.doc");
            Microsoft.Office.Interop.Word.Application ap = new Microsoft.Office.Interop.Word.Application();
            Document document = ap.Documents.Open(path);
        }

        private void btnFeasibilityStudy_Click(object sender, EventArgs e)
        {
            var path = Path.Combine(Directory.GetCurrentDirectory(), @"Documents/Feasibility Study.doc");
            Microsoft.Office.Interop.Word.Application ap = new Microsoft.Office.Interop.Word.Application();
            Document document = ap.Documents.Open(path);
        }

        private void btnTermsOfReference_Click(object sender, EventArgs e)
        {
            var path = Path.Combine(Directory.GetCurrentDirectory(), @"Documents/Terms of Reference.doc");
            Microsoft.Office.Interop.Word.Application ap = new Microsoft.Office.Interop.Word.Application();
            Document document = ap.Documents.Open(path);
        }

        private void btnJobDescriptions_Click(object sender, EventArgs e)
        {
            var path = Path.Combine(Directory.GetCurrentDirectory(), @"Documents/Job Description.doc");
            Microsoft.Office.Interop.Word.Application ap = new Microsoft.Office.Interop.Word.Application();
            Document document = ap.Documents.Open(path);
        }

        private void btnProjectOfficeChecklist_Click(object sender, EventArgs e)
        {
           var path = Path.Combine(Directory.GetCurrentDirectory(), @"Documents/Project Office Checklist.doc");
            Microsoft.Office.Interop.Word.Application ap = new Microsoft.Office.Interop.Word.Application();
            Document document = ap.Documents.Open(path);
        }

        private void btnPhaseReviewForm_Click(object sender, EventArgs e)
        {
            var path = Path.Combine(Directory.GetCurrentDirectory(), @"Documents/Phase Review Form - Planning");
            Microsoft.Office.Interop.Word.Application ap = new Microsoft.Office.Interop.Word.Application();
            Document document = ap.Documents.Open(path);
        }

        private void btnProjectPlan_Click(object sender, EventArgs e)
        {
            var path = Path.Combine(Directory.GetCurrentDirectory(), @"Documents/Project Plan.doc");
            Microsoft.Office.Interop.Word.Application ap = new Microsoft.Office.Interop.Word.Application();
            Document document = ap.Documents.Open(path);
        }

        private void btnResourcePlan_Click(object sender, EventArgs e)
        {
            var path = Path.Combine(Directory.GetCurrentDirectory(), @"Documents/Resource Plan.doc");
            Microsoft.Office.Interop.Word.Application ap = new Microsoft.Office.Interop.Word.Application();
            Document document = ap.Documents.Open(path);
        }

        private void btnFinancialPlan_Click(object sender, EventArgs e)
        {
            var path = Path.Combine(Directory.GetCurrentDirectory(), @"Documents/Financial Plan.doc");
            Microsoft.Office.Interop.Word.Application ap = new Microsoft.Office.Interop.Word.Application();
            Document document = ap.Documents.Open(path);
        }

        private void btnQualityPlan_Click(object sender, EventArgs e)
        {
            var path = Path.Combine(Directory.GetCurrentDirectory(), @"Documents/Quality Plan.doc");
            Microsoft.Office.Interop.Word.Application ap = new Microsoft.Office.Interop.Word.Application();
            Document document = ap.Documents.Open(path);
        }

        private void btnRiskPlan_Click(object sender, EventArgs e)
        {
            var path = Path.Combine(Directory.GetCurrentDirectory(), @"Documents/Risk Plan.doc");
            Microsoft.Office.Interop.Word.Application ap = new Microsoft.Office.Interop.Word.Application();
            Document document = ap.Documents.Open(path);
        }

        private void btnAcceptancePlan_Click(object sender, EventArgs e)
        {
            var path = Path.Combine(Directory.GetCurrentDirectory(), @"Documents/Acceptance Plan.doc");
            Microsoft.Office.Interop.Word.Application ap = new Microsoft.Office.Interop.Word.Application();
            Document document = ap.Documents.Open(path);
        }

        private void btnCommunicationsPlan_Click(object sender, EventArgs e)
        {
            var path = Path.Combine(Directory.GetCurrentDirectory(), @"Documents/Communications Plan.doc");
            Microsoft.Office.Interop.Word.Application ap = new Microsoft.Office.Interop.Word.Application();
            Document document = ap.Documents.Open(path);
        }

        private void btnProcurementPlan_Click(object sender, EventArgs e)
        {
            var path = Path.Combine(Directory.GetCurrentDirectory(), @"Documents/Procurement Plan.doc");
            Microsoft.Office.Interop.Word.Application ap = new Microsoft.Office.Interop.Word.Application();
            Document document = ap.Documents.Open(path);
        }

        private void btnTenderProcess_Click(object sender, EventArgs e)
        {
            var path = Path.Combine(Directory.GetCurrentDirectory(), @"Documents/Procurement Management Process.doc");
            Microsoft.Office.Interop.Word.Application ap = new Microsoft.Office.Interop.Word.Application();
            Document document = ap.Documents.Open(path);
        }

        private void btnStatementOfWork_Click(object sender, EventArgs e)
        {
            var path = Path.Combine(Directory.GetCurrentDirectory(), @"Documents/Statement of Work.doc");
            Microsoft.Office.Interop.Word.Application ap = new Microsoft.Office.Interop.Word.Application();
            Document document = ap.Documents.Open(path);
        }

        private void btnRequestForInformation_Click(object sender, EventArgs e)
        {
            var path = Path.Combine(Directory.GetCurrentDirectory(), @"Documents/Request for Information.doc");
            Microsoft.Office.Interop.Word.Application ap = new Microsoft.Office.Interop.Word.Application();
            Document document = ap.Documents.Open(path);
        }

        private void btnRequestForProposal_Click(object sender, EventArgs e)
        {
            var path = Path.Combine(Directory.GetCurrentDirectory(), @"Documents/Request for Proposal.doc");
            Microsoft.Office.Interop.Word.Application ap = new Microsoft.Office.Interop.Word.Application();
            Document document = ap.Documents.Open(path);
        }

        private void btnSupplierContract_Click(object sender, EventArgs e)
        {
            var path = Path.Combine(Directory.GetCurrentDirectory(), @"Documents/Supplier Contract.doc");
            Microsoft.Office.Interop.Word.Application ap = new Microsoft.Office.Interop.Word.Application();
            Document document = ap.Documents.Open(path);
        }

        private void btnTenderRegister_Click(object sender, EventArgs e)
        {
            var path = Path.Combine(Directory.GetCurrentDirectory(), @"Documents/Tender Register.xls");
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            Workbook workbook = app.Workbooks.Open(path, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, true, XlPlatform.xlWindows, Type.Missing, false, false, Type.Missing, false, Type.Missing, Type.Missing);
        }

        private void btnPhaseReview_Click(object sender, EventArgs e)
        {
            var path = Path.Combine(Directory.GetCurrentDirectory(), @"Documents/Phase Review Form - Planning");
            Microsoft.Office.Interop.Word.Application ap = new Microsoft.Office.Interop.Word.Application();
            Document document = ap.Documents.Open(path);
        }

        private void btnTimeManagementProcess_Click(object sender, EventArgs e)
        {
            var path = Path.Combine(Directory.GetCurrentDirectory(), @"Documents/Time Management Process.doc");
            Microsoft.Office.Interop.Word.Application ap = new Microsoft.Office.Interop.Word.Application();
            Document document = ap.Documents.Open(path);
        }

        private void btnCostManagementProcess_Click(object sender, EventArgs e)
        {
            var path = Path.Combine(Directory.GetCurrentDirectory(), @"Documents/Cost Management Process.doc");
            Microsoft.Office.Interop.Word.Application ap = new Microsoft.Office.Interop.Word.Application();
            Document document = ap.Documents.Open(path);
        }

        private void btnQualityManagementProcess_Click(object sender, EventArgs e)
        {
            var path = Path.Combine(Directory.GetCurrentDirectory(), @"Documents/Quality Management Process.doc");
            Microsoft.Office.Interop.Word.Application ap = new Microsoft.Office.Interop.Word.Application();
            Document document = ap.Documents.Open(path);
        }

        private void btnChangeManagementProcess_Click(object sender, EventArgs e)
        {
            var path = Path.Combine(Directory.GetCurrentDirectory(), @"Documents/Change Management Process.doc");
            Microsoft.Office.Interop.Word.Application ap = new Microsoft.Office.Interop.Word.Application();
            Document document = ap.Documents.Open(path);
        }

        private void btnRiskManagementProcess_Click(object sender, EventArgs e)
        {
            var path = Path.Combine(Directory.GetCurrentDirectory(), @"Documents/Risk Management Process.doc");
            Microsoft.Office.Interop.Word.Application ap = new Microsoft.Office.Interop.Word.Application();
            Document document = ap.Documents.Open(path);
        }

        private void btnIssueManagementProcess_Click(object sender, EventArgs e)
        {
            var path = Path.Combine(Directory.GetCurrentDirectory(), @"Documents/Issue Management Process.doc");
            Microsoft.Office.Interop.Word.Application ap = new Microsoft.Office.Interop.Word.Application();
            Document document = ap.Documents.Open(path);
        }

        private void btnProcurementManagementProcess_Click(object sender, EventArgs e)
        {
            var path = Path.Combine(Directory.GetCurrentDirectory(), @"Documents/Procurement Management Process.doc");
            Microsoft.Office.Interop.Word.Application ap = new Microsoft.Office.Interop.Word.Application();
            Document document = ap.Documents.Open(path);
        }

        private void btnAcceptanceManagementProcess_Click(object sender, EventArgs e)
        {
            var path = Path.Combine(Directory.GetCurrentDirectory(), @"Documents/Acceptance Management Process.doc");
            Microsoft.Office.Interop.Word.Application ap = new Microsoft.Office.Interop.Word.Application();
            Document document = ap.Documents.Open(path);
        }

        private void btnCommsManagementProcess_Click(object sender, EventArgs e)
        {
            var path = Path.Combine(Directory.GetCurrentDirectory(), @"Documents/Communications Management Process.doc");
            Microsoft.Office.Interop.Word.Application ap = new Microsoft.Office.Interop.Word.Application();
            Document document = ap.Documents.Open(path);
        }

        private void btnTimesheetForm_Click(object sender, EventArgs e)
        {
            var path = Path.Combine(Directory.GetCurrentDirectory(), @"Documents/Timesheet Form.doc");
            Microsoft.Office.Interop.Word.Application ap = new Microsoft.Office.Interop.Word.Application();
            Document document = ap.Documents.Open(path);
        }

        private void btnExpenseForm_Click(object sender, EventArgs e)
        {
            var path = Path.Combine(Directory.GetCurrentDirectory(), @"Documents/Expense Form.doc");
            Microsoft.Office.Interop.Word.Application ap = new Microsoft.Office.Interop.Word.Application();
            Document document = ap.Documents.Open(path);
        }

        private void btnQualityReviewForm_Click(object sender, EventArgs e)
        {
            var path = Path.Combine(Directory.GetCurrentDirectory(), @"Documents/Quality Review Form.doc");
            Microsoft.Office.Interop.Word.Application ap = new Microsoft.Office.Interop.Word.Application();
            Document document = ap.Documents.Open(path);
        }

        private void btnQuailtyReviewForm2_Click(object sender, EventArgs e)
        {
            var path = Path.Combine(Directory.GetCurrentDirectory(), @"Documents/Quality Review Form.doc");
            Microsoft.Office.Interop.Word.Application ap = new Microsoft.Office.Interop.Word.Application();
            Document document = ap.Documents.Open(path);
        }

        private void btnChangeRequestForm_Click(object sender, EventArgs e)
        {
            var path = Path.Combine(Directory.GetCurrentDirectory(), @"Documents/Change Request Form.doc");
            Microsoft.Office.Interop.Word.Application ap = new Microsoft.Office.Interop.Word.Application();
            Document document = ap.Documents.Open(path);
        }

        private void btnRiskForm_Click(object sender, EventArgs e)
        {
            var path = Path.Combine(Directory.GetCurrentDirectory(), @"Documents/Risk Form.doc");
            Microsoft.Office.Interop.Word.Application ap = new Microsoft.Office.Interop.Word.Application();
            Document document = ap.Documents.Open(path);
        }

        private void btnIssueForm_Click(object sender, EventArgs e)
        {
            var path = Path.Combine(Directory.GetCurrentDirectory(), @"Documents/Issue Form.doc");
            Microsoft.Office.Interop.Word.Application ap = new Microsoft.Office.Interop.Word.Application();
            Document document = ap.Documents.Open(path);
        }

        private void btnPurchaseOrderForm_Click(object sender, EventArgs e)
        {
            var path = Path.Combine(Directory.GetCurrentDirectory(), @"Documents/Purchase Order Form.doc");
            Microsoft.Office.Interop.Word.Application ap = new Microsoft.Office.Interop.Word.Application();
            Document document = ap.Documents.Open(path);
        }

        private void btnAcceptanceForm_Click(object sender, EventArgs e)
        {
            var path = Path.Combine(Directory.GetCurrentDirectory(), @"Documents/Acceptance Form.doc");
            Microsoft.Office.Interop.Word.Application ap = new Microsoft.Office.Interop.Word.Application();
            Document document = ap.Documents.Open(path);
        }

        private void btnProjectStatusReport_Click(object sender, EventArgs e)
        {
            var path = Path.Combine(Directory.GetCurrentDirectory(), @"Documents/Project Status Report.doc");
            Microsoft.Office.Interop.Word.Application ap = new Microsoft.Office.Interop.Word.Application();
            Document document = ap.Documents.Open(path);
        }

        private void btnPhaseReviewFormPE_Click(object sender, EventArgs e)
        {
            var path = Path.Combine(Directory.GetCurrentDirectory(), @"Documents/Phase Review Form - Execution.doc");
            Microsoft.Office.Interop.Word.Application ap = new Microsoft.Office.Interop.Word.Application();
            Document document = ap.Documents.Open(path);
        }

        private void btnTimesheetRegister_Click(object sender, EventArgs e)
        {
            TimesheetRegister timesheetRegister = new TimesheetRegister();
            timesheetRegister.Show();
        }

        private void btnExpenseRegister_Click(object sender, EventArgs e)
        {
            ExpenseRegister expenseRegister = new ExpenseRegister();
            expenseRegister.Show();
        }

        private void btnQualityRegister1_Click(object sender, EventArgs e)
        {
            QualityRegister qualityRegister = new QualityRegister();
            qualityRegister.Show();
        }

        private void btnQualityRegister2_Click(object sender, EventArgs e)
        {
            QualityRegister qualityRegister = new QualityRegister();
            qualityRegister.Show();
        }

        private void btnChangeRegister_Click(object sender, EventArgs e)
        {
            ChangeRegister changeRegister = new ChangeRegister();
            changeRegister.Show();
        }

        private void btnRiskRegisterPE_Click(object sender, EventArgs e)
        {
            RiskRegisterForm riskRegister = new RiskRegisterForm();
            riskRegister.Show();
        }

        private void btnIssueRegisterPE_Click(object sender, EventArgs e)
        {
            IssueRegisterForm issueRegister = new IssueRegisterForm();
            issueRegister.Show();
        }

        private void btnProcurementRegister_Click(object sender, EventArgs e)
        {
            ProcurementRegister procurementRegister = new ProcurementRegister();
            procurementRegister.Show();
        }

        private void btnAcceptanceRegister_Click(object sender, EventArgs e)
        {
            AcceptanceRegister acceptanceRegister = new AcceptanceRegister();
            acceptanceRegister.Show();
        }

        private void btnCommunicationsRegister_Click(object sender, EventArgs e)
        {
            CommunicationsRegister communicationsRegister = new CommunicationsRegister();
            communicationsRegister.Show();
        }

        private void btnClosureReport_Click(object sender, EventArgs e)
        {
            var path = Path.Combine(Directory.GetCurrentDirectory(), @"Documents/Project Closure Report.doc");
            Microsoft.Office.Interop.Word.Application ap = new Microsoft.Office.Interop.Word.Application();
            Document document = ap.Documents.Open(path);
        }

        private void btnPostImplementationReview_Click(object sender, EventArgs e)
        {
            var path = Path.Combine(Directory.GetCurrentDirectory(), @"Documents/Post Implementation Review.doc");
            Microsoft.Office.Interop.Word.Application ap = new Microsoft.Office.Interop.Word.Application();
            Document document = ap.Documents.Open(path);
        }
    }
}
