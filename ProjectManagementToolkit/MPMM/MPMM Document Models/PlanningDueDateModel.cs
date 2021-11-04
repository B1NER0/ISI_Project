using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProjectManagementToolkit.MPMM.MPMM_Document_Models
{
    class PlanningDueDateModel
    {
        public string ProjectPlan { get; set; } = "";
        public string ResourcePlan { get; set; } = "";
        public string FinancialPlan { get; set; }
        public string QualityPlan { get; set; }
        public string RiskPlan { get; set; }
        public string AcceptancePlan { get; set; }
        public string CommunicationPlan { get; set; }
        public string ProcurementPlan { get; set; }
        public string StatementOfWork { get; set; }
        public string RequestForInformation { get; set; }
        public string SupplierContract { get; set; }
        public string RequestForProposal { get; set; }
        public string PhaseReviewPlanning { get; set; }

    }
}
