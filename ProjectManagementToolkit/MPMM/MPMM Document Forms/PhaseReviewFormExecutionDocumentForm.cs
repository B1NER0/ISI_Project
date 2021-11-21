﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ProjectManagementToolkit.MPMM.MPMM_Document_Models;
using ProjectManagementToolkit.Utility;
using ProjectManagementToolkit.Properties;
using Newtonsoft.Json;
using MoreLinq;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace ProjectManagementToolkit.MPMM.MPMM_Document_Forms
{
    public partial class PhaseReviewFormExecutionDocumentForm : Form
    {
        VersionControl<PhaseReviewFormExecutionModel> versionControl;
        PhaseReviewFormExecutionModel newPhaseReviewExeModel;
        PhaseReviewFormExecutionModel currentPhaseReviewExeModel;
        Color TABLE_HEADER_COLOR = Color.FromArgb(73, 173, 252);
        ProjectModel projectModel = new ProjectModel();

        public PhaseReviewFormExecutionDocumentForm()
        {
            InitializeComponent();
        }

        private void label19_Click(object sender, EventArgs e)
        {

        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {

        }

        private void PhaseReviewFormExecutionDocumentForm_Load(object sender, EventArgs e)
        {
            loadDocument();
        }

        private void loadDocument()
        {
            dgvReviewDetails.Columns.Add("colReviewCategory", "ReviewCategory");
            dgvReviewDetails.Columns.Add("colReviewQuestion", "ReviewQuestion");
            dgvReviewDetails.Columns.Add("colAnswer", "Answer");
            dgvReviewDetails.Columns.Add("colVariance", "Variance");

            string json = JsonHelper.loadDocument(Settings.Default.ProjectID, "PhaseReviewExe");
            List<string[]> documentInfo = new List<string[]>();
            newPhaseReviewExeModel = new PhaseReviewFormExecutionModel();
            currentPhaseReviewExeModel = new PhaseReviewFormExecutionModel();

            string jsonWord = JsonHelper.loadProjectInfo(Settings.Default.Username);
            List<ProjectModel> projectListModel = JsonConvert.DeserializeObject<List<ProjectModel>>(jsonWord);
            projectModel = projectModel.getProjectModel(Settings.Default.ProjectID, projectListModel);

            if (json != "")
            {
                versionControl = JsonConvert.DeserializeObject<VersionControl<PhaseReviewFormExecutionModel>>(json);
                newPhaseReviewExeModel = JsonConvert.DeserializeObject<PhaseReviewFormExecutionModel>(versionControl.getLatest(versionControl.DocumentModels));
                currentPhaseReviewExeModel = JsonConvert.DeserializeObject<PhaseReviewFormExecutionModel>(versionControl.getLatest(versionControl.DocumentModels));

                txtProjectName2.Text = currentPhaseReviewExeModel.ProjectName;
                txtProjectManager.Text = currentPhaseReviewExeModel.ProjectManager;
                txtProjectSponsor.Text = currentPhaseReviewExeModel.ProjectSponsor;
                txtReportPreparedBy.Text = currentPhaseReviewExeModel.ReportPreparedBy;
                txtReportPreperationDate.Text = currentPhaseReviewExeModel.ReportPreparationDate;
                txtReportingPeriod.Text = currentPhaseReviewExeModel.eportingPeriod;

                txtSummary.Text = currentPhaseReviewExeModel.Summary;
                txtProjectSchedule.Text = currentPhaseReviewExeModel.ProjectSchedule;
                txtProjectExpenses.Text = currentPhaseReviewExeModel.ProjectExpenses;
                txtProjectDeliverables.Text = currentPhaseReviewExeModel.ProjectDeliverables;
                txtProjectRisks.Text = currentPhaseReviewExeModel.ProjectRisks;
                txtProjectIssues.Text = currentPhaseReviewExeModel.ProjectIssues;
                txtProjectChanges.Text = currentPhaseReviewExeModel.ProjectChanges;

                txtSupportingDocumentation.Text = currentPhaseReviewExeModel.SupportingDocumentation;
                txtSignature.Text = currentPhaseReviewExeModel.Signature;
                txtDate.Text = currentPhaseReviewExeModel.SignatureDate;

                foreach (var row in currentPhaseReviewExeModel.ReviewDetials)
                {
                    dgvReviewDetails.Rows.Add(new string[] { row.ReviewCategory, row.ReviewQuestion, row.Answer, row.Varaince });
                }

            }
            else
            {
                versionControl = new VersionControl<PhaseReviewFormExecutionModel>();
                versionControl.DocumentModels = new List<VersionControl<PhaseReviewFormExecutionModel>.DocumentModel>();
                documentInfo.Add(new string[] { "Document ID", "" });
                documentInfo.Add(new string[] { "Document Owner", "" });
                documentInfo.Add(new string[] { "Issue Date", "" });
                documentInfo.Add(new string[] { "Last Save Date", "" });
                documentInfo.Add(new string[] { "File Name", "" });
                newPhaseReviewExeModel = new PhaseReviewFormExecutionModel();
                
            }            
        }

        private void saveDocument()
        {
            newPhaseReviewExeModel.ProjectName = txtProjectName2.Text;
            newPhaseReviewExeModel.ProjectManager = txtProjectManager.Text;
            newPhaseReviewExeModel.ProjectSponsor = txtProjectSponsor.Text;
            newPhaseReviewExeModel.ReportPreparedBy = txtReportPreparedBy.Text;
            newPhaseReviewExeModel.ReportPreparationDate = txtReportPreperationDate.Text;
            newPhaseReviewExeModel.eportingPeriod = txtReportingPeriod.Text;
            newPhaseReviewExeModel.PhaseReviewExeProgress = "DONE";
            newPhaseReviewExeModel.completedDate = DateTime.Now.ToString("yyyy/MM/dd");

            newPhaseReviewExeModel.Summary = txtSummary.Text;
            newPhaseReviewExeModel.ProjectSchedule = txtProjectSchedule.Text;
            newPhaseReviewExeModel.ProjectExpenses = txtProjectExpenses.Text;
            newPhaseReviewExeModel.ProjectDeliverables = txtProjectDeliverables.Text;
            newPhaseReviewExeModel.ProjectRisks = txtProjectRisks.Text;
            newPhaseReviewExeModel.ProjectIssues = txtProjectIssues.Text;
            newPhaseReviewExeModel.ProjectChanges = txtProjectChanges.Text;

            newPhaseReviewExeModel.SupportingDocumentation = txtSupportingDocumentation.Text;
            newPhaseReviewExeModel.Signature = txtSignature.Text;
            newPhaseReviewExeModel.SignatureDate = txtDate.Text;

            List<PhaseReviewFormExecutionModel.ReviewDetial> review = new List<PhaseReviewFormExecutionModel.ReviewDetial>();

            int reviewRowsCount = dgvReviewDetails.Rows.Count;

            for (int i = 0; i < reviewRowsCount - 1; i++)
            {
                PhaseReviewFormExecutionModel.ReviewDetial rev = new PhaseReviewFormExecutionModel.ReviewDetial();
                var ReviewCategory = dgvReviewDetails.Rows[i].Cells[0].Value?.ToString() ?? "";
                var ReviewQuestion = dgvReviewDetails.Rows[i].Cells[1].Value?.ToString() ?? "";
                var Answer = dgvReviewDetails.Rows[i].Cells[2].Value?.ToString() ?? "";
                var Varaince = dgvReviewDetails.Rows[i].Cells[3].Value?.ToString() ?? "";

                rev.ReviewCategory = ReviewCategory;
                rev.ReviewQuestion = ReviewQuestion;
                rev.Answer = Answer;
                rev.Varaince = Varaince;

                review.Add(rev);
            }
            newPhaseReviewExeModel.ReviewDetials = review;


            List<VersionControl<PhaseReviewFormExecutionModel>.DocumentModel> documentModels = versionControl.DocumentModels;

            if (!versionControl.isEqual(currentPhaseReviewExeModel, newPhaseReviewExeModel))
            {
                VersionControl<PhaseReviewFormExecutionModel>.DocumentModel documentModel = new VersionControl<PhaseReviewFormExecutionModel>.DocumentModel(newPhaseReviewExeModel, DateTime.Now, VersionControl<ProjectModel>.generateID());

                documentModels.Add(documentModel);

                versionControl.DocumentModels = documentModels;

                string json = JsonConvert.SerializeObject(versionControl);
                currentPhaseReviewExeModel = JsonConvert.DeserializeObject<PhaseReviewFormExecutionModel>(JsonConvert.SerializeObject(newPhaseReviewExeModel));

                JsonHelper.saveDocument(json, Settings.Default.ProjectID, "PhaseReviewExe");
                MessageBox.Show("Phase review execution saved successfully", "save", MessageBoxButtons.OK);
            }
        }

        private void exportToWord()
        {
            string path;
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                saveFileDialog.Filter = "Word 97-2003 Documents (*.doc)|*.doc|Word 2007 Documents (*.docx)|*.docx";
                saveFileDialog.FilterIndex = 2;
                saveFileDialog.RestoreDirectory = true;
                

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    path = saveFileDialog.FileName;
                    using (var document = DocX.Create(path))
                    {
                        document.InsertParagraph("Phase review form \nFor " + projectModel.ProjectName)
                           .Font("Arial")
                           .Bold(true)
                           .FontSize(22d).Alignment = Alignment.left;
                        document.InsertSectionPageBreak();


                        var p = document.InsertParagraph();
                        var title = p.InsertParagraphBeforeSelf("Table of Contents").Bold().FontSize(20);

                        var tocSwitches = new Dictionary<TableOfContentsSwitches, string>()
                        {
                            { TableOfContentsSwitches.O, "1-3"},
                            { TableOfContentsSwitches.U, ""},
                            { TableOfContentsSwitches.Z, ""},
                            { TableOfContentsSwitches.H, ""}
                        };

                        document.InsertTableOfContents(p, "", tocSwitches);
                        document.InsertParagraph().InsertPageBreakAfterSelf();
                        var CommunicationReqHeading = document.InsertParagraph("1 Project details")
                            .Bold()
                            .FontSize(14d)
                            .Color(Color.Black)
                            .Bold(true)
                            .Font("Arial");

                        CommunicationReqHeading.StyleId = "Heading1";

                        var ComPlanHeading = document.InsertParagraph("1.1 Project name")
                     .Bold()
                     .FontSize(12d)
                     .Color(Color.Black)
                     .Bold(true)
                     .Font("Arial");

                        document.InsertParagraph(currentPhaseReviewExeModel.ProjectName)
                             .FontSize(11d)
                             .Color(Color.Black)
                             .Font("Arial").Alignment = Alignment.left;


                        ComPlanHeading.StyleId = "Heading2";

                        var projIDHeading = document.InsertParagraph("1.2 Project ID")
                         .Bold()
                         .FontSize(12d)
                         .Color(Color.Black)
                         .Bold(true)
                         .Font("Arial");

                        document.InsertParagraph(Settings.Default.ProjectID)
                             .FontSize(11d)
                             .Color(Color.Black)
                             .Font("Arial").Alignment = Alignment.left;


                        projIDHeading.StyleId = "Heading2";

                        var promanagerHeading = document.InsertParagraph("1.3 Project manager")
                         .Bold()
                         .FontSize(12d)
                         .Color(Color.Black)
                         .Bold(true)
                         .Font("Arial");

                        document.InsertParagraph(currentPhaseReviewExeModel.ProjectManager)
                             .FontSize(11d)
                             .Color(Color.Black)
                             .Font("Arial").Alignment = Alignment.left;


                        promanagerHeading.StyleId = "Heading2";

                        var sponsorHeading = document.InsertParagraph("1.4 Project sponsor")
                         .Bold()
                         .FontSize(12d)
                         .Color(Color.Black)
                         .Bold(true)
                         .Font("Arial");

                        document.InsertParagraph(currentPhaseReviewExeModel.ProjectSponsor)
                             .FontSize(11d)
                             .Color(Color.Black)
                             .Font("Arial").Alignment = Alignment.left;


                        sponsorHeading.StyleId = "Heading2";

                        var reportHeading = document.InsertParagraph("1.5 Report prepared by")
                        .Bold()
                        .FontSize(12d)
                        .Color(Color.Black)
                        .Bold(true)
                        .Font("Arial");

                        document.InsertParagraph(currentPhaseReviewExeModel.ReportPreparedBy)
                             .FontSize(11d)
                             .Color(Color.Black)
                             .Font("Arial").Alignment = Alignment.left;


                        sponsorHeading.StyleId = "Heading2";

                        var prepdateHeading = document.InsertParagraph("1.6 Report preperation date")
                        .Bold()
                        .FontSize(12d)
                        .Color(Color.Black)
                        .Bold(true)
                        .Font("Arial");

                        document.InsertParagraph(currentPhaseReviewExeModel.ReportPreparationDate)
                             .FontSize(11d)
                             .Color(Color.Black)
                             .Font("Arial").Alignment = Alignment.left;


                        prepdateHeading.StyleId = "Heading2";

                        var rperHeading = document.InsertParagraph("1.7 Reporting Period")
                        .Bold()
                        .FontSize(12d)
                        .Color(Color.Black)
                        .Bold(true)
                        .Font("Arial");

                        document.InsertParagraph(currentPhaseReviewExeModel.eportingPeriod)
                             .FontSize(11d)
                             .Color(Color.Black)
                             .Font("Arial").Alignment = Alignment.left;


                        rperHeading.StyleId = "Heading2";

                       /* var recHeading = document.InsertParagraph("1.8 Reporting Recipients")
                        .Bold()
                        .FontSize(12d)
                        .Color(Color.Black)
                        .Bold(true)
                        .Font("Arial");

                        document.InsertParagraph(currentPhaseReviewExeModel.Recipients)
                             .FontSize(11d)
                             .Color(Color.Black)
                             .Font("Arial").Alignment = Alignment.left;


                        recHeading.StyleId = "Heading2"; */

                        var sumHeading = document.InsertParagraph("2 Overall status")
                            .Bold()
                            .FontSize(14d)
                            .Color(Color.Black)
                            .Bold(true)
                            .Font("Arial");

                        CommunicationReqHeading.StyleId = "Heading1";

                        var desHeading = document.InsertParagraph("2.1 Project Summary")
                     .Bold()
                     .FontSize(12d)
                     .Color(Color.Black)
                     .Bold(true)
                     .Font("Arial");

                        document.InsertParagraph(currentPhaseReviewExeModel.Summary)
                             .FontSize(11d)
                             .Color(Color.Black)
                             .Font("Arial").Alignment = Alignment.left;


                        desHeading.StyleId = "Heading2";




                        var scheHeading = document.InsertParagraph("2.2 Project schedule")
                         .Bold()
                         .FontSize(12d)
                         .Color(Color.Black)
                         .Bold(true)
                         .Font("Arial");

                        document.InsertParagraph(currentPhaseReviewExeModel.ProjectSchedule)
                             .FontSize(11d)
                             .Color(Color.Black)
                             .Font("Arial").Alignment = Alignment.left;


                        scheHeading.StyleId = "Heading2";

                        var expensesHeading = document.InsertParagraph("2.3 Project expenses")
                         .Bold()
                         .FontSize(12d)
                         .Color(Color.Black)
                         .Bold(true)
                         .Font("Arial");

                        document.InsertParagraph(currentPhaseReviewExeModel.ProjectExpenses)
                             .FontSize(11d)
                             .Color(Color.Black)
                             .Font("Arial").Alignment = Alignment.left;


                        expensesHeading.StyleId = "Heading2";

                        var delivHeading = document.InsertParagraph("2.4 Project Deliverables")
                        .Bold()
                        .FontSize(12d)
                        .Color(Color.Black)
                        .Bold(true)
                        .Font("Arial");

                        document.InsertParagraph(currentPhaseReviewExeModel.ProjectDeliverables)
                             .FontSize(11d)
                             .Color(Color.Black)
                             .Font("Arial").Alignment = Alignment.left;


                        delivHeading.StyleId = "Heading2";

                        var riskHeading = document.InsertParagraph("2.5 Project risks")
                        .Bold()
                        .FontSize(12d)
                        .Color(Color.Black)
                        .Bold(true)
                        .Font("Arial");

                        document.InsertParagraph(currentPhaseReviewExeModel.ProjectRisks)
                             .FontSize(11d)
                             .Color(Color.Black)
                             .Font("Arial").Alignment = Alignment.left;


                        riskHeading.StyleId = "Heading2";

                        var issuesHeading = document.InsertParagraph("2.6 Reporting Issues")
                        .Bold()
                        .FontSize(12d)
                        .Color(Color.Black)
                        .Bold(true)
                        .Font("Arial");

                        document.InsertParagraph(currentPhaseReviewExeModel.ProjectIssues)
                             .FontSize(11d)
                             .Color(Color.Black)
                             .Font("Arial").Alignment = Alignment.left;


                        issuesHeading.StyleId = "Heading2";

                        var changesHeading = document.InsertParagraph("2.7 Reporting Changes")
                        .Bold()
                        .FontSize(12d)
                        .Color(Color.Black)
                        .Bold(true)
                        .Font("Arial");

                        document.InsertParagraph(currentPhaseReviewExeModel.ProjectChanges)
                             .FontSize(11d)
                             .Color(Color.Black)
                             .Font("Arial").Alignment = Alignment.left;


                        changesHeading.StyleId = "Heading2";

                        var revHead = document.InsertParagraph("3 Review details")
                           .Bold()
                           .FontSize(14d)
                           .Color(Color.Black)
                           .Bold(true)
                           .Font("Arial");

                        revHead.StyleId = "Heading1";

                        var docReviewDetail = document.AddTable(currentPhaseReviewExeModel.ReviewDetials.Count + 1, 4);
                        docReviewDetail.Rows[0].Cells[0].Paragraphs[0].Append("ReviewCategory")
                            .Bold(true)
                            .Color(Color.White);
                        docReviewDetail.Rows[0].Cells[1].Paragraphs[0].Append("ReviewQuestion")
                            .Bold(true)
                            .Color(Color.White);
                        docReviewDetail.Rows[0].Cells[2].Paragraphs[0].Append("Answer")
                            .Bold(true)
                            .Color(Color.White);
                        docReviewDetail.Rows[0].Cells[3].Paragraphs[0].Append("Varaince")
                            .Bold(true)
                            .Color(Color.White);
                        docReviewDetail.Rows[0].Cells[0].FillColor = TABLE_HEADER_COLOR;
                        docReviewDetail.Rows[0].Cells[1].FillColor = TABLE_HEADER_COLOR;
                        docReviewDetail.Rows[0].Cells[2].FillColor = TABLE_HEADER_COLOR;
                        docReviewDetail.Rows[0].Cells[3].FillColor = TABLE_HEADER_COLOR;

                        for (int i = 1; i < currentPhaseReviewExeModel.ReviewDetials.Count + 1; i++)
                        {
                            docReviewDetail.Rows[i].Cells[0].Paragraphs[0].Append(currentPhaseReviewExeModel.ReviewDetials[i - 1].ReviewCategory);
                            docReviewDetail.Rows[i].Cells[1].Paragraphs[0].Append(currentPhaseReviewExeModel.ReviewDetials[i - 1].ReviewQuestion);
                            docReviewDetail.Rows[i].Cells[2].Paragraphs[0].Append(currentPhaseReviewExeModel.ReviewDetials[i - 1].Answer);
                            docReviewDetail.Rows[i].Cells[3].Paragraphs[0].Append(currentPhaseReviewExeModel.ReviewDetials[i - 1].Varaince);
                        }
                        docReviewDetail.SetWidths(new float[] { 493, 332, 508, 254 });
                        document.InsertTable(docReviewDetail);
                        document.InsertParagraph().InsertPageBreakAfterSelf();


                        var adHeading = document.InsertParagraph("4 Approval details")
                           .Bold()
                           .FontSize(14d)
                           .Color(Color.Black)
                           .Bold(true)
                           .Font("Arial");

                        CommunicationReqHeading.StyleId = "Heading1";

                        var sdHeading = document.InsertParagraph("4.1 Supporting documentation")
                     .Bold()
                     .FontSize(12d)
                     .Color(Color.Black)
                     .Bold(true)
                     .Font("Arial");

                        document.InsertParagraph(currentPhaseReviewExeModel.SupportingDocumentation)
                             .FontSize(11d)
                             .Color(Color.Black)
                             .Font("Arial").Alignment = Alignment.left;


                        sdHeading.StyleId = "Heading2";

                        var psSignature = document.InsertParagraph("4.2 Project signature")
                         .Bold()
                         .FontSize(12d)
                         .Color(Color.Black)
                         .Bold(true)
                         .Font("Arial");

                        document.InsertParagraph(currentPhaseReviewExeModel.Signature)
                             .FontSize(11d)
                             .Color(Color.Black)
                             .Font("Arial").Alignment = Alignment.left;


                        psSignature.StyleId = "Heading2";

                        var dateHeading = document.InsertParagraph("4.3 Date")
                         .Bold()
                         .FontSize(12d)
                         .Color(Color.Black)
                         .Bold(true)
                         .Font("Arial");

                        document.InsertParagraph(currentPhaseReviewExeModel.SignatureDate)
                             .FontSize(11d)
                             .Color(Color.Black)
                             .Font("Arial").Alignment = Alignment.left;


                        dateHeading.StyleId = "Heading2";

                        try
                        {
                            document.Save();
                        }
                        catch (Exception)
                        {
                            MessageBox.Show("The selected File is open.", "Close File", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            saveDocument();
        }

        private void PhaseReviewFormExecutionDocumentForm_Load_1(object sender, EventArgs e)
        {
            
        }

        private void Btn_Save_Document_Click(object sender, EventArgs e)
        {
            
        }

        private void btn_Export_Document_Click(object sender, EventArgs e)
        {
            
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            saveDocument();
        }

        private void btnExportWord_Click(object sender, EventArgs e)
        {
            exportToWord();
        }

        private void PhaseReviewFormExecutionDocumentForm_Load_2(object sender, EventArgs e)
        {
            string jsoni = JsonHelper.loadProjectInfo(Settings.Default.Username);
            List<ProjectModel> projectListModel = JsonConvert.DeserializeObject<List<ProjectModel>>(jsoni);
            projectModel = projectModel.getProjectModel(Settings.Default.ProjectID, projectListModel);
            txtProjectName.Text = projectModel.ProjectName;
            loadDocument();
        }

        private void btnSaveProgress_Click(object sender, EventArgs e)
        {
            newPhaseReviewExeModel.ProjectName = txtProjectName2.Text;
            newPhaseReviewExeModel.ProjectManager = txtProjectManager.Text;
            newPhaseReviewExeModel.ProjectSponsor = txtProjectSponsor.Text;
            newPhaseReviewExeModel.ReportPreparedBy = txtReportPreparedBy.Text;
            newPhaseReviewExeModel.ReportPreparationDate = txtReportPreperationDate.Text;
            newPhaseReviewExeModel.eportingPeriod = txtReportingPeriod.Text;
            newPhaseReviewExeModel.PhaseReviewExeProgress = "UNDONE";

            newPhaseReviewExeModel.Summary = txtSummary.Text;
            newPhaseReviewExeModel.ProjectSchedule = txtProjectSchedule.Text;
            newPhaseReviewExeModel.ProjectExpenses = txtProjectExpenses.Text;
            newPhaseReviewExeModel.ProjectDeliverables = txtProjectDeliverables.Text;
            newPhaseReviewExeModel.ProjectRisks = txtProjectRisks.Text;
            newPhaseReviewExeModel.ProjectIssues = txtProjectIssues.Text;
            newPhaseReviewExeModel.ProjectChanges = txtProjectChanges.Text;

            newPhaseReviewExeModel.SupportingDocumentation = txtSupportingDocumentation.Text;
            newPhaseReviewExeModel.Signature = txtSignature.Text;
            newPhaseReviewExeModel.SignatureDate = txtDate.Text;

            List<PhaseReviewFormExecutionModel.ReviewDetial> review = new List<PhaseReviewFormExecutionModel.ReviewDetial>();

            int reviewRowsCount = dgvReviewDetails.Rows.Count;

            for (int i = 0; i < reviewRowsCount - 1; i++)
            {
                PhaseReviewFormExecutionModel.ReviewDetial rev = new PhaseReviewFormExecutionModel.ReviewDetial();
                var ReviewCategory = dgvReviewDetails.Rows[i].Cells[0].Value?.ToString() ?? "";
                var ReviewQuestion = dgvReviewDetails.Rows[i].Cells[1].Value?.ToString() ?? "";
                var Answer = dgvReviewDetails.Rows[i].Cells[2].Value?.ToString() ?? "";
                var Varaince = dgvReviewDetails.Rows[i].Cells[3].Value?.ToString() ?? "";

                rev.ReviewCategory = ReviewCategory;
                rev.ReviewQuestion = ReviewQuestion;
                rev.Answer = Answer;
                rev.Varaince = Varaince;

                review.Add(rev);
            }
            newPhaseReviewExeModel.ReviewDetials = review;


            List<VersionControl<PhaseReviewFormExecutionModel>.DocumentModel> documentModels = versionControl.DocumentModels;

            if (!versionControl.isEqual(currentPhaseReviewExeModel, newPhaseReviewExeModel))
            {
                VersionControl<PhaseReviewFormExecutionModel>.DocumentModel documentModel = new VersionControl<PhaseReviewFormExecutionModel>.DocumentModel(newPhaseReviewExeModel, DateTime.Now, VersionControl<ProjectModel>.generateID());

                documentModels.Add(documentModel);

                versionControl.DocumentModels = documentModels;

                string json = JsonConvert.SerializeObject(versionControl);
                currentPhaseReviewExeModel = JsonConvert.DeserializeObject<PhaseReviewFormExecutionModel>(JsonConvert.SerializeObject(newPhaseReviewExeModel));

                JsonHelper.saveDocument(json, Settings.Default.ProjectID, "PhaseReviewExe");
                MessageBox.Show("Phase review execution saved successfully", "save", MessageBoxButtons.OK);
            }
        }
    }
}
