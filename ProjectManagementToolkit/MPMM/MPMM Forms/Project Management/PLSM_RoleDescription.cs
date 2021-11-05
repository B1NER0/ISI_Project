using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ProjectManagementToolkit.MPMM.MPMM_Forms.Project_Management
{
    public partial class PLSM_RoleDescription : Form
    {
        public PLSM_RoleDescription()
        {
            InitializeComponent();
        }

        private void BtnBackToPLCMFrontEND_Click(object sender, EventArgs e)
        {
            this.Hide();
        }

        private void BtnCostExpenseManagement_Click(object sender, EventArgs e)
        {
            txtProcessRolesDescription.Text = "Team Member \r\n" +
                                                "Complete regular Expense Forms to the level of detail required \r\n" +
                                                "Submit Expense forms timeously \r\n" +
                                                "Provide additional information regarding time spent when required \r\n \r\n" +

                                                "Project Manager \r\n" +
                                                "Informing all staff about the Cost Management process \r\n" +
                                                "Ensure timeously completion of Expense Forms throughout the duration of the project \r\n" +
                                                "Reviewing and approving all Expense Forms \r\n \r\n" +

                                                "Project Administrator \r\n" +
                                                "Manages day to day Expense process \r\n" +
                                                "Provides all staff with basic Expense Form templates \r\n" +
                                                "Ensures level of detail completion of Expense Forms \r\n" +
                                                "Ensures all Expense Forms have been signed off by the Project Manager \r\n" +
                                                "Keeping The Expense register up to date \r\n" +
                                                "Updating the Project Plan and identifying deviations \r\n" +
                                                "Arrange payment for approved expenses \r\n";
                                                
        }

        private void BtnRiskmanagement_Click(object sender, EventArgs e)
        {
            txtProcessRolesDescription.Text = "Team Member \r\n" +
                                                "Identify risks within the project \r\n" +
                                                "Completing a risk form for each identified risk \r\n" +
                                                "Submitting risk forms for review \r\n" +
                                                "Completing_implementing risks actions as identified by the PM \r\n" +

                                                "Project Manager \r\n" +
                                                "Review all risks to determine their priority \r\n" +
                                                "Implementing risk actions for low and medium priority risks \r\n" +
                                                "Reviewing the effectiveness of risk actions after implementation \r\n" +
                                                "Maintain a risk register \r\n" +

                                                "Project Board \r\n" +
                                                "Conducts a review of high priority risks as identified by the PM \r\n" +
                                                "Identify actions to take for: \r\n" +
                                                "Mitigate \r\n" +
                                                "Transfer \r\n" +
                                                "Avoid \r\n" +
                                                "Supports the PM with implementation of all risk actions \r\n";
                                                
        }
    }
}
