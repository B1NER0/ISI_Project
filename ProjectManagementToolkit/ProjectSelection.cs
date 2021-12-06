using ProjectManagementToolkit.Utility;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;
using System.Web;
using System.Net;
using System.IO;
using Newtonsoft.Json.Linq;
using System.Net.Http;
using ProjectManagementToolkit.Properties;
using System.Diagnostics;
using ProjectManagementToolkit.Classes;


namespace ProjectManagementToolkit.MPMM.MPMM_Document_Forms
{
    public partial class ProjectSelection : Form
    {
        List<ProjectModel> projectListModel = new List<ProjectModel>();
        private static readonly HttpClient client = new HttpClient();

        public ProjectSelection()
        {
            InitializeComponent();
        }

        private void btnSelectProject_Click(object sender, EventArgs e)
        {
            if (lstboxProject.SelectedIndex != -1)
            {
                int index = lstboxProject.SelectedIndex;
                Settings.Default.ProjectID = projectListModel[index].ProjectID;
                MainForm mainForm = new MainForm();
                mainForm.WindowState = FormWindowState.Maximized;
                mainForm.Show();
                this.Visible = false;

                string projectName = projectListModel[index].ProjectName;
                string result = Path.GetTempPath();
                //MessageBox.Show("@" + result + "ProjectNames.txt");
                StreamWriter projectNamesFile;
                string path = Path.Combine("@", result, "ProjectNames.txt");
                projectNamesFile = File.CreateText(path);
                projectNamesFile.WriteLine(projectName);
                projectNamesFile.Close();
            }
        }

        private void btnCreateProject_Click(object sender, EventArgs e)
        {           
            List<string> listMembers = new List<string>();
            if (!string.IsNullOrEmpty(txtProjectName.Text) && !string.IsNullOrEmpty(txtProjectSponsor.Text) && !string.IsNullOrEmpty(txtProjectManager.Text))
            {
                string defaultPassword = Settings.Default.Default_Password;
                string role = "admin";

                new clsRestAPIHandler().create_user(Settings.Default.Username, defaultPassword, role);
                new clsFileHandler().writeToFile(Settings.Default.Username, new clsFileHandler().get_user_file());
                

                listMembers.Add(txtProjectSponsor.Text);
               // if (new clsRestAPIHandler().create_user(listMembers[0], defaultPassword, role)) { MessageBox.Show("Success 1 !!!"); }
                listMembers.Add(txtProjectReviewGroup.Text);
                new clsRestAPIHandler().create_user(listMembers[1], defaultPassword, role);
                listMembers.Add(txtProjectManager.Text);
                new clsRestAPIHandler().create_user(listMembers[2], defaultPassword, role);
                listMembers.Add(txtQualityManager.Text);
                new clsRestAPIHandler().create_user(listMembers[3], defaultPassword, role);
                listMembers.Add(txtProcurementManager.Text);
                new clsRestAPIHandler().create_user(listMembers[4], defaultPassword, role);
                listMembers.Add(txtCommunicationsManager.Text);
                new clsRestAPIHandler().create_user(listMembers[5], defaultPassword, role);
                listMembers.Add(txtProjectOfficeManager.Text);
                new clsRestAPIHandler().create_user(listMembers[6], defaultPassword, role);
                string projectName = txtProjectName.Text;
                List<string> sprints = new List<string>(); 
               
                sprints.Add("Default_sprint");
                DateTime dateStart = new DateTime();
                dateStart = DateTime.Now;
                DateTime dateEnd = new DateTime(9998, 12, 31);
                string currentUser = Settings.Default.Username;

                if (validateProjAdd(projectName, sprints[0], listMembers, dateStart, dateEnd))
                {
                    
                    JObject obj_proj = new clsRestAPIHandler().create_project(projectName, listMembers, sprints);
                    update_user_projects(listMembers, projectName, currentUser);
                    get_updated_project_file();
                    if (obj_proj != null)
                    {
                        new clsRestAPIHandler().create_sprint(sprints[0], projectName, dateStart, dateEnd);
                        //lblOutput.Text = obj_proj["message"].ToString();
                    }
                    //lblOutput.Text = obj_proj["message"].ToString();
                    //AddTabPageResetControls();

                }

                ProjectModel newProject = new ProjectModel();
                string projectID = newProject.generateID();
                Settings.Default.ProjectID = projectID;

                newProject.ProjectID = projectID;
                newProject.ProjectName = txtProjectName.Text;
                newProject.ProjectSponsor = txtProjectSponsor.Text;
                newProject.ProjectReviewGroup = txtProjectReviewGroup.Text;
                newProject.ProjectManager = txtProjectManager.Text;
                newProject.QualityManager = txtQualityManager.Text;
                newProject.ProcurementManager = txtProcurementManager.Text;
                newProject.CommunicationsManager = txtCommunicationsManager.Text;
                newProject.OfficeManager = txtProjectOfficeManager.Text;

                projectListModel.Add(newProject);
                Settings.Default.ProjectID = newProject.ProjectID;
                string json = JsonConvert.SerializeObject(projectListModel);
                JsonHelper.saveProjectInfo(json, Settings.Default.Username);
                MainForm mainForm = new MainForm();
                mainForm.WindowState = FormWindowState.Maximized;
                mainForm.Show();
                this.Visible = false;
            }
            else
            {
                MessageBox.Show("Please ensure that you enter a project Name,ProjectSponsor and a Project Manager before continuing");
            }
        }


        private bool validateProjAdd(string proj_name, string sprint_name, List<string> list, DateTime dateStart, DateTime dateEnd)
        {
            if (proj_name.Length <= 0 || proj_name.Contains(" "))
            {
                //lblOutput.Text = "Please enter a valid project name eg: proj_name";
                //txtProjName.Focus();
                return false;
            }

            if (sprint_name.Length <= 0 || sprint_name.Contains(" "))
            {
                //lblOutput.Text = "Please enter a valid sprint name eg: sprint_1";
                //txtSprintName.Focus();
                return false;
            }

            if (list.Count <= 0)
            {
                //lblOutput.Text = "No members added to project";
                //list.Focus();
                return false;
            }

            if (DateTime.Compare(dateEnd, dateStart) < 0)
            {
                //lblOutput.Text = "Sprint end date is before the start date.";
                //dStart.Focus();
                return false;
            }

            if (DateTime.Compare(dateEnd, dateStart) == 0)
            {
                //lblOutput.Text = "Sprint start and date are on the same day!";
                //dEnd.Focus();
                return false;
            }

            //lblOutput.Text = "";
            return true;
        }

        public void update_user_projects(List<string> list, string project, string currentUser)
        {
            //Adds project to each user in the listbox to DB.
            for (int i = 0; i < list.Count; i++)
            {
                JObject obj = new clsRestAPIHandler().get_user_info(list[i].ToString());
                string projects = obj["user"][0]["projects"].ToString();
                JArray user_projects = JArray.Parse(projects);
                user_projects.Add(project);
                string json_payload = new clsRestAPIHandler().prepareJsonPayload("projects", user_projects);
                new clsRestAPIHandler().update_user(list[i].ToString(), json_payload);
            }

            //Adds project to the current user in DB.
            //string current_user = new clsFileHandler().readFromFile(new clsFileHandler().get_user_file());
            JObject obj_current_user = new clsRestAPIHandler().get_user_info(currentUser);
            string projects_current_user = obj_current_user["user"][0]["projects"].ToString();
            JArray projects_current_user_array = JArray.Parse(projects_current_user);
            projects_current_user_array.Add(project);
            string current_payload = new clsRestAPIHandler().prepareJsonPayload("projects", projects_current_user_array);
            new clsRestAPIHandler().update_user(currentUser, current_payload);
        }

        private void get_updated_project_file()
        {
            new clsFileHandler().deleteFile(new clsFileHandler().get_project_file());
            string current_user = new clsFileHandler().readFromFile(new clsFileHandler().get_user_file());
            JObject obj = new clsRestAPIHandler().get_user_info(current_user);
            string projects = obj["user"][0]["projects"].ToString();
            JArray user_projects = JArray.Parse(projects);
            new clsFileHandler().writeMutlipleLines(user_projects, new clsFileHandler().get_project_file());
        }


        private void ProjectSelection_Load(object sender, EventArgs e)
        {
            string json = JsonHelper.loadProjectInfo(Settings.Default.Username);
            if (json != "")
            {
                projectListModel = JsonConvert.DeserializeObject<List<ProjectModel>>(json);
            }

            foreach (var project in projectListModel)
            {
                lstboxProject.Items.Add(project.ProjectName);
                //projectNames.Add(project.ProjectName);    
            }


          
            txtProjectCode.Text = placeholderText;
            txtProjectCode.ForeColor = SystemColors.GrayText;

            //string result = Path.GetTempPath();
            //MessageBox.Show("@"+result+"ProjectNames.txt");
            //StreamWriter projectNamesFile;
            //string path = Path.Combine("@", result, "ProjectNames.txt");
            //projectNamesFile = File.CreateText(path);
            //foreach (var projectName in projectNames)
            //{
            //    projectNamesFile.WriteLine(projectName);
            //}
            //projectNamesFile.Close();

        }

        private void btnProjectCode_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;

            if (txtProjectCode.Text.Contains(" ") || txtProjectCode.Text.Contains(".") || txtProjectCode.Text == "")
            {
                MessageBox.Show("Incorrect Project ID.", "Sync Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtProjectCode.Text = "";
                return;
            }

            string projectCodeToAdd = txtProjectCode.Text;

            bool containsItem = projectListModel.Any(item => item.ProjectID == projectCodeToAdd);


            if (containsItem)
            {
                ProjectModel projectItem = projectListModel.Find(item => item.ProjectID == projectCodeToAdd);
                string projectName = projectItem.ProjectName;

                MessageBox.Show("Project Already added as " + projectName, "Project ID Error.", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Cursor.Current = Cursors.Default;
                return;
            }

            bool connectionSuccessful = attemptHttpConnection();

            if (!connectionSuccessful)
            {
                MessageBox.Show("Unable to connect to server.", "Server Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            string projectID = txtProjectCode.Text;

            //Get Project Config from server.
            try
            {
                HttpResponseMessage responseMessage = client.GetAsync(Settings.Default.URI + "/project/" + projectID).Result;
                var jsonResponse = responseMessage.Content.ReadAsStringAsync().Result;
                int statusCode = responseMessage.StatusCode.GetHashCode();

                if (jsonResponse == "[]" || jsonResponse == "")
                {
                    MessageBox.Show("Incorrect Project ID", "Project Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                JObject projectModel = JArray.Parse(jsonResponse)[0].ToObject<JObject>();

                ProjectModel newProject = new ProjectModel();



                newProject.ProjectID = projectModel["ProjectID"].ToString();
                newProject.ProjectName = projectModel["ProjectName"].ToString();
                newProject.ProjectSponsor = projectModel["ProjectSponsor"].ToString();
                newProject.ProjectReviewGroup = projectModel["ProjectReviewGroup"].ToString();
                newProject.ProjectManager = projectModel["ProjectManager"].ToString();
                newProject.QualityManager = projectModel["QualityManager"].ToString();
                newProject.ProcurementManager = projectModel["ProcurementManager"].ToString();
                newProject.CommunicationsManager = projectModel["CommunicationsManager"].ToString();
                newProject.OfficeManager = projectModel["OfficeManager"].ToString();
                newProject.LastDateTimeSynced = projectModel["LastDateTimeSynced"].ToObject<DateTime>();

                projectListModel.Add(newProject);

                string json = JsonConvert.SerializeObject(projectListModel);

                JsonHelper.saveProjectInfo(json, Settings.Default.Username);

                lstboxProject.Items.Clear();
                foreach (var project in projectListModel)
                {
                    lstboxProject.Items.Add(project.ProjectName);

                }
                txtProjectCode.Text = "";
                MessageBox.Show("Successfully added Project: " + projectModel["ProjectName"].ToString());
                Cursor.Current = Cursors.Default;
            }
            catch (AggregateException)
            {
                MessageBox.Show("An unexpected server error occurred.", "Server Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

        }

        private bool attemptHttpConnection()
        {

            try
            {
                Task<HttpResponseMessage> responseMessage = client.GetAsync(Settings.Default.URI + "/");
                HttpResponseMessage response = responseMessage.Result;
                int statusCode = response.StatusCode.GetHashCode();

                switch (statusCode)
                {
                    case 200:
                        return true;
                    case 404:
                        return false;
                    default:
                        break;
                }

                return false;
            }
            catch (AggregateException)
            {
                return false;
            }
        }

        private void ProjectSelection_FormClosing(object sender, FormClosingEventArgs e)
        {
            Application.Exit();
        }

        string placeholderText = "Enter Project Code";

        private void txtProjectCode_Enter(object sender, EventArgs e)
        {
            if(txtProjectCode.Text == placeholderText)
            {
                txtProjectCode.Text = "";
                txtProjectCode.ForeColor = SystemColors.WindowText;
            }
        }

        private void txtProjectCode_Leave(object sender, EventArgs e)
        {
            if (txtProjectCode.Text == "")
            {
                txtProjectCode.Text = placeholderText;
                txtProjectCode.ForeColor = SystemColors.GrayText;
            }
        }

        private void lstboxProject_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void btnCreateProject_Click_1(object sender, EventArgs e)
        {

        }
    }
}
