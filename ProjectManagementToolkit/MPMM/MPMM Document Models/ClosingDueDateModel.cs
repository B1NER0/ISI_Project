using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProjectManagementToolkit.MPMM.MPMM_Document_Models
{
    class ClosingDueDateModel
    {
        //Start Date
        public string ProjectClosureReportSD { get; set; }
        public string PostImplementationReviewSD { get; set; }

        //Complete data for tasks
        public string ProjectClosureReportCD { get; set; }
        public string PostImplementationReviewCD { get; set; }

        //Due date
        public string ProjectClosureReportDD { get; set; }
        public string PostImplementationReviewDD { get; set; }

        //Planned Budget
        public string ProjectClosureReportPlannedBudget { get; set; }
        public string PostImplementationReviewPlannedBudget { get; set; }

        //Budget Used
        public string ProjectClosureReportBudget { get; set; }
        public string PostImplementationReviewBudget { get; set; }
    }
}
