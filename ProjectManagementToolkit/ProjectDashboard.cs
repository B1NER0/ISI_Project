using Newtonsoft.Json;
using ProjectManagementToolkit.Properties;
using ProjectManagementToolkit.Utility;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace ProjectManagementToolkit
{
    public partial class ProjectDashboard : Form
    {
        ProjectModel projectModel = new ProjectModel();
        public ProjectDashboard()
        {
            InitializeComponent();
        }

        // Made variables public for 'Complete' buttons to work
        double initationProgressVal = 0;
        double planningProgressVal = 0;
        double executionProgressVal = 0;
        double closingProgressVal = 0;
        double initationPercentage = 0;
        double planningPercentage = 0;
        double executionPercentage = 0;
        double closingPercentage = 0;


        // Made lists public for 'Complete' buttons to work
        List<string> closingDocuments = new List<string>();
        List<string> initiationDocuments = new List<string>();
        List<string> planningDocuments = new List<string>();
        List<string> executionDocuments = new List<string>();

        string[] xValues1 = new string[4];
        double[] yValues1 = new double[4];
        double[] yValues2 = new double[4];

        private void ProjectDashboard_Load(object sender, EventArgs e)
        {
            string json = JsonHelper.loadProjectInfo(Settings.Default.Username);
            List<ProjectModel> projectListModel = JsonConvert.DeserializeObject<List<ProjectModel>>(json);
            projectModel = projectModel.getProjectModel(Settings.Default.ProjectID, projectListModel);

            List<string> localDocuments = getLocalDocuments();

            lblProjectName.Text = projectModel.ProjectName;

            chart1.ChartAreas[0].BackColor = Color.Transparent;
            chart1.Legends[0].BackColor = Color.Transparent;
            chart2.Legends[0].BackColor = Color.Transparent;

            if (localDocuments == null)
            {
                MessageBox.Show("No documents have been added yet.", "File Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                
                initiationDocuments.Add("BusinessCase");
                initiationDocuments.Add("FeasibilityStudy");
                initiationDocuments.Add("ProjectCharter");
                initiationDocuments.Add("JobDescription");
                initiationDocuments.Add("ProjectOfficeCheckList");
                initiationDocuments.Add("PhaseReviewFormInitiation");
                initiationDocuments.Add("TermOfReferenceDocument");

                lblInitiationProgress.Text = "Progress: 0%";
                pbarInitiation.Value = 0;
                pbarInitiation.Maximum = initiationDocuments.Count;
                for (int i = 0; i < initiationDocuments.Count; i++)
                {
                    dgvInitiation.Rows.Add();
                    dgvInitiation.Rows[i].Cells[0].Value = initiationDocuments[i];
                    if (localDocuments.Contains(initiationDocuments[i]))
                    {
                        initationProgressVal++;
                        dgvInitiation.Rows[i].Cells[1].Value = true;
                        pbarInitiation.Value = (int)initationProgressVal;
                        initationPercentage = ((initationProgressVal) / initiationDocuments.Count) * 100;
                        lblInitiationProgress.Text = "Progress: " + Math.Round(initationPercentage, 2) + "%";

                        xValues1[0] = "Initiation";
                        yValues1[0] = initationPercentage;

                        yValues2[0] = 100 - initationPercentage;

                        chart2.Series["Completed"].Points.DataBindXY(xValues1, yValues1);
                        chart2.Series["Uncompleted"].Points.DataBindXY(xValues1, yValues2);
                    }
                    else
                    {
                        dgvInitiation.Rows[i].Cells[1].Value = false;
                    }
                }

                
                planningDocuments.Add("ProjectPlan");
                planningDocuments.Add("ResourcePlan");
                planningDocuments.Add("FinancialPlan");
                planningDocuments.Add("QualityPlan");
                planningDocuments.Add("RiskPlan");
                planningDocuments.Add("AcceptancePlan");
                planningDocuments.Add("CommunicationPlan");
                planningDocuments.Add("ProcurementPlan");
                planningDocuments.Add("StatementOfWork");
                planningDocuments.Add("RequestForInformation");
                planningDocuments.Add("SupplierContract");
                planningDocuments.Add("RequestForProposal");
                planningDocuments.Add("PhaseReviewPlanning");

                lblPlanningProgress.Text = "Progress: 0%";
                pbarPlanning.Value = 0;
                pbarPlanning.Maximum = planningDocuments.Count;
                for (int i = 0; i < planningDocuments.Count; i++)
                {
                    dgvPlanning.Rows.Add();
                    dgvPlanning.Rows[i].Cells[0].Value = planningDocuments[i];
                    if (localDocuments.Contains(planningDocuments[i]))
                    {
                        planningProgressVal++;
                        dgvPlanning.Rows[i].Cells[1].Value = true;
                        pbarPlanning.Value = (int)planningProgressVal;
                        planningPercentage = ((planningProgressVal) / planningDocuments.Count) * 100;
                        lblPlanningProgress.Text = "Progress: " + Math.Round(planningPercentage, 2) + "%";

                        xValues1[1] = "Planning";
                        yValues1[1] = planningPercentage;

                        yValues2[1] = 100 - planningPercentage;

                        chart2.Series["Completed"].Points.DataBindXY(xValues1, yValues1);
                        chart2.Series["Uncompleted"].Points.DataBindXY(xValues1, yValues2);
                    }
                    else
                    {
                        dgvPlanning.Rows[i].Cells[1].Value = false;
                    }
                }

                
                executionDocuments.Add("BuildDeliverables");
                executionDocuments.Add("MonitorAndControl");
                executionDocuments.Add("TimeMangement");
                executionDocuments.Add("TimeSheet");
                executionDocuments.Add("TimeSheetRegister");
                executionDocuments.Add("CostManagementProcess");
                executionDocuments.Add("ExpenseForm");
                executionDocuments.Add("ExpenseRegister");
                executionDocuments.Add("QualityManagement");
                executionDocuments.Add("QualityReviewPlan");
                executionDocuments.Add("QualityReviewForm");
                executionDocuments.Add("QualityReviewRegister");
                executionDocuments.Add("ChangeManagementProcess");
                executionDocuments.Add("ChangeRequestForm");
                executionDocuments.Add("ChangeRequestRegister");
                executionDocuments.Add("RiskManagamentProcess");
                executionDocuments.Add("RiskForm");
                executionDocuments.Add("RiskRegister");
                executionDocuments.Add("IssueManagementProcess");
                executionDocuments.Add("IssueForm");
                executionDocuments.Add("IssueRegister");
                executionDocuments.Add("PurchaseOrder");
                executionDocuments.Add("ProcurementRegister");
                executionDocuments.Add("AcceptanceManagementProcess");
                executionDocuments.Add("AcceptanceForm");
                executionDocuments.Add("AcceptanceRegister");
                executionDocuments.Add("CommunicationsManagementProcess");
                executionDocuments.Add("ProjectStatusReport");
                executionDocuments.Add("CommunicationsRegister");
                executionDocuments.Add("PhaseReviewExe");

                lblExecutionProgress.Text = "Progress: 0%";
                pbarExecution.Value = 0;
                pbarExecution.Maximum = executionDocuments.Count;
                for (int i = 0; i < executionDocuments.Count; i++)
                {
                    dgvExecution.Rows.Add();
                    dgvExecution.Rows[i].Cells[0].Value = executionDocuments[i];
                    if (localDocuments.Contains(executionDocuments[i]))
                    {
                        executionProgressVal++;
                        dgvExecution.Rows[i].Cells[1].Value = true;
                        pbarExecution.Value = (int)executionProgressVal;
                        executionPercentage = ((executionProgressVal) / executionDocuments.Count) * 100;
                        lblExecutionProgress.Text = "Progress: " + Math.Round(executionPercentage, 2) + "%";

                        xValues1[2] = "Execution";
                        yValues1[2] = executionPercentage;

                        yValues2[2] = 100 - executionPercentage;

                        chart2.Series["Completed"].Points.DataBindXY(xValues1, yValues1);
                        chart2.Series["Uncompleted"].Points.DataBindXY(xValues1, yValues2);
                    }
                    else
                    {
                        dgvExecution.Rows[i].Cells[1].Value = false;
                    }
                }

                
                closingDocuments.Add("ProjectClosureReport");
                closingDocuments.Add("PostImplementationReview");

                lblClosingProgress.Text = "Progress: 0%";
                pbarClosing.Value = 0;
                pbarClosing.Maximum = closingDocuments.Count;
                for (int i = 0; i < closingDocuments.Count; i++)
                {
                    dgvClosing.Rows.Add();
                    dgvClosing.Rows[i].Cells[0].Value = closingDocuments[i];
                    if (localDocuments.Contains(closingDocuments[i]))
                    {
                        closingProgressVal++;
                        dgvClosing.Rows[i].Cells[1].Value = true;
                        pbarClosing.Value = (int)closingProgressVal;
                        closingPercentage = ((closingProgressVal) / closingDocuments.Count) * 100;
                        lblClosingProgress.Text = "Progress: " + Math.Round(closingPercentage, 2) + "%";

                        xValues1[3] = "Closing";
                        yValues1[3] = closingPercentage;

                        yValues2[3] = 100 - closingPercentage;

                        chart2.Series["Completed"].Points.DataBindXY(xValues1, yValues1);
                        chart2.Series["Uncompleted"].Points.DataBindXY(xValues1, yValues2);
                    }
                    else
                    {
                        dgvClosing.Rows[i].Cells[1].Value = false;
                        
                        xValues1[3] = "Closing";
                        yValues1[3] = 0;

                        yValues2[3] = 100;

                        chart2.Series["Completed"].Points.DataBindXY(xValues1, yValues1);
                        chart2.Series["Uncompleted"].Points.DataBindXY(xValues1, yValues2);
                        
                    }
                }

                double overallProgressVal = (initationProgressVal + planningProgressVal + executionProgressVal + closingProgressVal);
                pbarOverall.Value = (int)overallProgressVal;
                pbarOverall.Maximum = initiationDocuments.Count + planningDocuments.Count + executionDocuments.Count + closingDocuments.Count;
                double overallPercentage = ((overallProgressVal) / pbarOverall.Maximum) * 100;
                lblOverallProgress.Text = "Overall Progress: " + Math.Round(overallPercentage, 2) + "%";

                string[] xValues = { "Completed Tasks", "Uncompleted Tasks" };
                double[] yValues = { overallPercentage, 100 - overallPercentage };

                chart1.Series["Series1"].Points.DataBindXY(xValues, yValues);
                chart1.Series["Series1"].ChartType = SeriesChartType.Doughnut;

                chart1.Series["Series1"]["PieLabelStyle"] = "Disabled";
                chart1.Legends[0].Enabled = true;

                chart1.Series["Series1"].Points[0].Color = Color.LimeGreen;
                chart1.Series["Series1"].Points[1].Color = Color.Gray;


                foreach (DataPoint p in chart1.Series["Series1"].Points)
                {
                    p.Label = "#PERCENT\n#VALX";
                }
            }
        }

        private List<string> getLocalDocuments()
        {
            List<string> localDocuments = new List<string>();
            string projectPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ProjectManagementToolkit", Settings.Default.ProjectID);

            if (Directory.Exists(projectPath))
            {
                foreach (string documentPath in Directory.GetFiles(projectPath))
                {
                    string documentName = Path.GetFileNameWithoutExtension(documentPath);
                    localDocuments.Add(documentName);
                }
            }
            else
            {
                return null;
            }

            return localDocuments;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            initationProgressVal = 0;
            for (int i = 0; i < pbarInitiation.Maximum; i++) // Go through documents
            {
                initationProgressVal++; //Increase progress
                dgvInitiation.Rows[i].Cells[1].Value = true; // check each unchecked checkbox
                
            }
            pbarInitiation.Value = pbarInitiation.Maximum;
            lblInitiationProgress.Text = "Progress: 100%";

            //////////////////// Calculate overall progressbar percentage //////////////
            double overallProgressVal = (initationProgressVal + planningProgressVal + executionProgressVal + closingProgressVal);
            pbarOverall.Value = (int)overallProgressVal;
            pbarOverall.Maximum = initiationDocuments.Count + planningDocuments.Count + executionDocuments.Count + closingDocuments.Count;
            double overallPercentage = ((overallProgressVal) / pbarOverall.Maximum) * 100;
            lblOverallProgress.Text = "Overall Progress: " + Math.Round(overallPercentage, 2) + "%";
            ////////////////////////////////////////////////////////////////////////////
            string[] xValues = { "Completed Tasks", "Uncompleted Tasks" };
            double[] yValues = { overallPercentage, 100 - overallPercentage };

            chart1.Series["Series1"].Points.DataBindXY(xValues, yValues);
            chart1.Series["Series1"].ChartType = SeriesChartType.Doughnut;

            chart1.Series["Series1"]["PieLabelStyle"] = "Disabled";
            chart1.Legends[0].Enabled = true;

            chart1.Series["Series1"].Points[0].Color = Color.LimeGreen;
            chart1.Series["Series1"].Points[1].Color = Color.Gray;

            foreach (DataPoint p in chart1.Series["Series1"].Points)
            {
                p.Label = "#PERCENT\n#VALX";
            }

            initationPercentage = ((initationProgressVal) / initiationDocuments.Count) * 100;

            xValues1[0] = "Initiation";
            //yValues1[0] = 0;

            yValues2[0] = 100 - initationPercentage;

            chart2.Series["Completed"].Points.DataBindXY(xValues1, yValues1);
            chart2.Series["Uncompleted"].Points.DataBindXY(xValues1, yValues2);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            planningProgressVal = 0;
            for (int i = 0; i < pbarPlanning.Maximum; i++) // Go through documents
            {
                planningProgressVal++; //Increase progress
                dgvPlanning.Rows[i].Cells[1].Value = true; // check each unchecked checkbox
            }

            pbarPlanning.Value = pbarPlanning.Maximum;
            lblPlanningProgress.Text = "Progress: 100%";

            //////////////////// Calculate overall progressbar percentage //////////////
            double overallProgressVal = (initationProgressVal + planningProgressVal + executionProgressVal + closingProgressVal);
            pbarOverall.Value = (int)overallProgressVal;
            pbarOverall.Maximum = initiationDocuments.Count + planningDocuments.Count + executionDocuments.Count + closingDocuments.Count;
            double overallPercentage = ((overallProgressVal) / pbarOverall.Maximum) * 100;
            lblOverallProgress.Text = "Overall Progress: " + Math.Round(overallPercentage, 2) + "%";
            ////////////////////////////////////////////////////////////////////////////
            string[] xValues = { "Completed Tasks", "Uncompleted Tasks" };
            double[] yValues = { overallPercentage, 100 - overallPercentage };

            chart1.Series["Series1"].Points.DataBindXY(xValues, yValues);
            chart1.Series["Series1"].ChartType = SeriesChartType.Doughnut;

            chart1.Series["Series1"]["PieLabelStyle"] = "Disabled";
            chart1.Legends[0].Enabled = true;

            chart1.Series["Series1"].Points[0].Color = Color.LimeGreen;
            chart1.Series["Series1"].Points[1].Color = Color.Gray;


            foreach (DataPoint p in chart1.Series["Series1"].Points)
            {
                p.Label = "#PERCENT\n#VALX";
            }

            planningPercentage = ((planningProgressVal) / planningDocuments.Count) * 100;

            yValues1[1] = 100;
            yValues2[1] = 100 - planningPercentage;

            chart2.Series["Completed"].Points.DataBindXY(xValues1, yValues1);
            chart2.Series["Uncompleted"].Points.DataBindXY(xValues1, yValues2);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            executionProgressVal = 0;
            for (int i = 0; i < pbarExecution.Maximum; i++) // Go through documents
            {
                executionProgressVal++; //Increase progress
                dgvExecution.Rows[i].Cells[1].Value = true; // check each unchecked checkbox
            }
            pbarExecution.Value = pbarExecution.Maximum;
            lblExecutionProgress.Text = "Progress: 100%";

            //////////////////// Calculate overall progressbar percentage //////////////
            double overallProgressVal = (initationProgressVal + planningProgressVal + executionProgressVal + closingProgressVal);
            pbarOverall.Value = (int)overallProgressVal;
            pbarOverall.Maximum = initiationDocuments.Count + planningDocuments.Count + executionDocuments.Count + closingDocuments.Count;
            double overallPercentage = ((overallProgressVal) / pbarOverall.Maximum) * 100;
            lblOverallProgress.Text = "Overall Progress: " + Math.Round(overallPercentage, 2) + "%";
            ////////////////////////////////////////////////////////////////////////////
            ///
            string[] xValues = { "Completed Tasks", "Uncompleted Tasks" };
            double[] yValues = { overallPercentage, 100 - overallPercentage };

            chart1.Series["Series1"].Points.DataBindXY(xValues, yValues);
            chart1.Series["Series1"].ChartType = SeriesChartType.Doughnut;

            chart1.Series["Series1"]["PieLabelStyle"] = "Disabled";
            chart1.Legends[0].Enabled = true;

            chart1.Series["Series1"].Points[0].Color = Color.LimeGreen;
            chart1.Series["Series1"].Points[1].Color = Color.Gray;


            foreach (DataPoint p in chart1.Series["Series1"].Points)
            {
                p.Label = "#PERCENT\n#VALX";
            }

            executionPercentage = ((executionProgressVal) / executionDocuments.Count) * 100;

            yValues1[2] = 100;
            yValues2[2] = 100 - executionPercentage;

            chart2.Series["Completed"].Points.DataBindXY(xValues1, yValues1);
            chart2.Series["Uncompleted"].Points.DataBindXY(xValues1, yValues2);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            closingProgressVal = 0;
            for (int i = 0; i < pbarClosing.Maximum; i++) // Go through documents
            {
                closingProgressVal++; //Increase progress
                dgvClosing.Rows[i].Cells[1].Value = true; // check each unchecked checkbox
            }
            pbarClosing.Value = pbarClosing.Maximum;
            lblClosingProgress.Text = "Progress: 100%";

            //////////////////// Calculate overall progressbar percentage //////////////
            double overallProgressVal = (initationProgressVal + planningProgressVal + executionProgressVal + closingProgressVal);
            pbarOverall.Value = (int)overallProgressVal;
            pbarOverall.Maximum = initiationDocuments.Count + planningDocuments.Count + executionDocuments.Count + closingDocuments.Count;
            double overallPercentage = ((overallProgressVal) / pbarOverall.Maximum) * 100;
            lblOverallProgress.Text = "Overall Progress: " + Math.Round(overallPercentage, 2) + "%";
            ////////////////////////////////////////////////////////////////////////////
            ///
            string[] xValues = { "Completed Tasks", "Uncompleted Tasks" };
            double[] yValues = { overallPercentage, 100 - overallPercentage };

            chart1.Series["Series1"].Points.DataBindXY(xValues, yValues);
            chart1.Series["Series1"].ChartType = SeriesChartType.Doughnut;

            chart1.Series["Series1"]["PieLabelStyle"] = "Disabled";
            chart1.Legends[0].Enabled = true;

            chart1.Series["Series1"].Points[0].Color = Color.LimeGreen;
            chart1.Series["Series1"].Points[1].Color = Color.Gray;


            foreach (DataPoint p in chart1.Series["Series1"].Points)
            {
                p.Label = "#PERCENT\n#VALX";
            }

            closingPercentage = ((closingProgressVal) / closingDocuments.Count) * 100;

            yValues1[3] = 100;
            yValues2[3] = 100 - closingPercentage;

            chart2.Series["Completed"].Points.DataBindXY(xValues1, yValues1);
            chart2.Series["Uncompleted"].Points.DataBindXY(xValues1, yValues2);
        }
    }
}
