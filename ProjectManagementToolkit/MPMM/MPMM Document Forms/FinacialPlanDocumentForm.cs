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
    public partial class FinacialPlanDocumentForm : Form
    {
        public FinacialPlanDocumentForm()
        {
            InitializeComponent();
        }

        private void btnSaveAssumptions_Click(object sender, EventArgs e)
        {
            string assumptions = txtAssumptions.Text;
        }

        private void btnSaveConstraints_Click(object sender, EventArgs e)
        {
            string constraints = txtConstraints.Text;
        }

        private void btnSaveActivitiesRolesDocuments_Click(object sender, EventArgs e)
        {
            string activies = txtActivities.Text;
            string roles = txtRoles.Text;
            string documents = txtDocuments.Text;
        }

        private void btnSaveProjectName_Click(object sender, EventArgs e)
        {
            string projectName = txtProjectName.Text;
        }

        private void dataGridViewOther_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void tabControlFinancialExpense_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
