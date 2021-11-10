using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProjectManagementToolkit.MPMM.MPMM_Document_Models
{
    class PlanningDueDateModel
    {
        public string ProjectPlanDD { get; set; } = "";
        public string ResourcePlanDD { get; set; } = "";
        public string FinancialPlanDD { get; set; }
        public string QualityPlanDD { get; set; }
        public string RiskPlanDD { get; set; }
        public string AcceptancePlanDD { get; set; }
        public string CommunicationPlanDD { get; set; }
        public string ProcurementPlanDD { get; set; }
        public string StatementOfWorkDD { get; set; }
        public string RequestForInformationDD { get; set; }
        public string SupplierContractDD { get; set; }
        public string RequestForProposalDD { get; set; }
        public string PhaseReviewPlanningDD { get; set; }

        public string ProjectPlanBudget { get; set; } = "";
        public string ResourcePlanBudget { get; set; } = "";
        public string FinancialPlanBudget { get; set; }
        public string QualityPlanBudget { get; set; }
        public string RiskPlanBudget { get; set; }
        public string AcceptancePlanBudget { get; set; }
        public string CommunicationPlanBudget { get; set; }
        public string ProcurementPlanBudget { get; set; }
        public string StatementOfWorkBudget { get; set; }
        public string RequestForInformationBudget { get; set; }
        public string SupplierContractBudget { get; set; }
        public string RequestForProposalBudget { get; set; }
        public string PhaseReviewPlanningBudget { get; set; }

    }
}
