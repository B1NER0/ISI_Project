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
    public partial class frmProcessFlowOverview : Form
    {
        public frmProcessFlowOverview()
        {
            InitializeComponent();
        }

        private void BtnBackToPLSM_Click(object sender, EventArgs e)
        {
            this.Hide();
            
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            MessageBox.Show("This vertical grouping display the assingment of responsibilities in accordance with the Lifecycle and various Stages.  Click on the Vertical bar to the right to see the specific assignment.");
        }

        private void BtnObjectives_Click(object sender, EventArgs e)
        {
            MessageBox.Show("This vertical grouping display the assingment of objectives in accordance with the Lifecycle and various Stages.  Click on the Vertical bar to the right to see the specific assignment.");
        }

        private void BtnAccountable_Click(object sender, EventArgs e)
        {
            MessageBox.Show("This vertical grouping display the assingment of Accountabilities in accordance with the Lifecycle and various Stages.  Click on the Vertical bar to the right to see the specific assignment.");
        }

        private void Btnparics_Click(object sender, EventArgs e)
        {
            MessageBox.Show("This vertical grouping display the PATRICS matrix in accordance with the Lifecycle and various Stages.  Click on the Vertical bar to the right to see the specific assignment.");
        }

        private void PictureBox5_Click(object sender, EventArgs e)
        {

        }
    }
}
