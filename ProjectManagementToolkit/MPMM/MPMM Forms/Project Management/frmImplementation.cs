using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.OleDb;
using Microsoft.Office.Interop;

namespace ProjectManagementToolkit.MPMM.MPMM_Forms.Project_Management
{
    public partial class frmImplementation : Form
    {
        public frmImplementation()
        {
            InitializeComponent();
        }

        private void btnBackToImplementation_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
