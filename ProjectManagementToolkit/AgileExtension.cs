using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;

namespace ProjectManagementToolkit
{
    public partial class AgileExtension : Form
    {
        public AgileExtension()
        {
            InitializeComponent();
        }

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

        private void Form1_Load(object sender, EventArgs e)
        {
            string appPath = readFile();

            if (appPath != "empty")
            {
                this.appPath = appPath;
                btnRun.Enabled = true;
            }
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
            this.Close();
        }

        private void AgileExtension_Load(object sender, EventArgs e)
        {
            string appPath = readFile();

            if (appPath != "empty")
            {
                this.appPath = appPath;
                btnRun.Enabled = true;
            }
        }
    }
}
