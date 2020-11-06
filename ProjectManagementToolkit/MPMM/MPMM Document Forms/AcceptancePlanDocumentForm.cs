﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ProjectManagementToolkit.MPMM.MPMM_Document_Forms
{
    public partial class AcceptancePlanDocumentForm : Form
    {
        public AcceptancePlanDocumentForm()
        {
            InitializeComponent();
        }

        private void AcceptancePlanDocumentForm_Load(object sender, EventArgs e)
        {

        }

        private void btnSaveProjectName_Click(object sender, EventArgs e)
        {
            string projectName = txtProjectName.Text;
        }

        private void btnCompanyOverview_Click(object sender, EventArgs e)
        {
            string addAssumptions = txtAssumptions.Text;
            listBoxAssumptions.Items.Add(addAssumptions);
        }

        private void btnConstraints_Click(object sender, EventArgs e)
        {
            string addConstraints = txtConstraints.Text;
            listBoxConstraints.Items.Add(addConstraints);
        }

        private void btnActivities_Click(object sender, EventArgs e)
        {
            string activities = txtActivities.Text;
            listBoxActivities.Items.Add(activities);
        }

        private void btnRoles_Click(object sender, EventArgs e)
        {
            string roles = txtRoles.Text;
            listBoxRoles.Items.Add(roles);
        }

        private void btnDocuments_Click(object sender, EventArgs e)
        {
            string documents = txtDocuments.Text;
            listBoxDocuments.Items.Add(documents);
        }
    }
}
