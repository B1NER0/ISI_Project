﻿using ProjectManagementToolkit.Utility;
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
            }
        }

        private void btnCreateProject_Click(object sender, EventArgs e)
        {
            if(!string.IsNullOrEmpty(txtProjectName.Text) && !string.IsNullOrEmpty(txtProjectSponsor.Text) && !string.IsNullOrEmpty(txtProjectManager.Text))
            {
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
            }

            string appPath = readFile();

            if (appPath != "empty")
            {
                this.appPath = appPath;
                btnRun.Enabled = true;
            }
        }

        private void btnProjectCode_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;

            if(txtProjectCode.Text.Contains(" ") || txtProjectCode.Text.Contains(".") || txtProjectCode.Text == "")
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

        //Instance variables for setting path
        private static string path = Path.GetTempPath();
        private static string APP_FILE_PATH = "\\Agile_app.txt";
        private string appPath = "";

        private string get_app_file()
        {
            return APP_FILE_PATH;
        }

        private bool writeToFile(string file, string content)
        {
            try
            {
                using (StreamWriter writer = new StreamWriter(path + file))
                {
                    writer.WriteLine(content);
                }
                return true;
            }
            catch (IOException e)
            {
                lblIndicator.Text = e.ToString();
            }

            return false;
        }

        private string readFile()
        {
            try
            {
                string line;
                using (StreamReader reader = new StreamReader(path + get_app_file()))
                {
                    line = reader.ReadToEnd();
                    line = line.Trim();
                    lblIndicator.Text = "External application path found.";
                    return line;
                }
            }
            catch (IOException e)
            {
                lblIndicator.Text = "No application set.";
            }
            return "empty";
        }

        private static Task<int> runApp(string app_path)
        {
            var tcs = new TaskCompletionSource<int>();

            var process = new Process
            {
                StartInfo = { FileName = app_path },
                EnableRaisingEvents = true
            };

            process.Exited += (sender, args) =>
            {
                tcs.SetResult(process.ExitCode);
                process.Dispose();
            };

            process.Start();

            return tcs.Task;
        }       

      
        private void ProjectSelection_FormClosing(object sender, FormClosingEventArgs e)
        {
            Application.Exit();
        }

        private void btnSetPath_Click(object sender, EventArgs e)
        {
            btnRun.Enabled = false;
            openAgile.InitialDirectory = @"C:\\";
            openAgile.Filter = "Exe Files (.exe)|*.exe|All Files (*.*)|*.*";
            openAgile.RestoreDirectory = true;

            if (openAgile.ShowDialog() == DialogResult.OK)
            {
                string appPath = openAgile.FileName;
                if (writeToFile(get_app_file(), appPath))
                {
                    lblIndicator.Text = "External application path saved";
                    this.appPath = appPath;
                    btnRun.Enabled = true;
                }
            }
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            runApp(this.appPath);
        }
    }
}
