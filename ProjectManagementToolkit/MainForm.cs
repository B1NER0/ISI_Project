﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ProjectManagementToolkit
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void governanceToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Governance governance = new Governance();
            governance.Show();
            governance.MdiParent = this;
        }
    }
}
