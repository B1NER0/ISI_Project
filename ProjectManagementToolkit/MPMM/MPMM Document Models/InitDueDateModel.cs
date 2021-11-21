using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProjectManagementToolkit.MPMM.MPMM_Document_Models
{
    class InitDueDateModel
    {
        // Start date for tasks
        public string BusinessCaseSD { get; set; }
        public string FeasibilityStudySD { get; set; }
        public string ProjectCharterSD { get; set; }
        public string JobDescriptionSD { get; set; }
        public string ProjectOfficeCheckListSD { get; set; }
        public string PhaseRevieFormInitiationSD { get; set; }
        public string TermOfReferenceDocumentSD { get; set; }

        // Complete date for tasks

        public string BusinessCaseCD { get; set; }
        public string FeasibilityStudyCD { get; set; }
        public string ProjectCharterCD { get; set; }
        public string JobDescriptionCD { get; set; }
        public string ProjectOfficeCheckListCD { get; set; }
        public string PhaseRevieFormInitiationCD { get; set; }
        public string TermOfReferenceDocumentCD { get; set; }

        // Due date for tasks
        public string BusinessCaseDD { get; set; }
        public string FeasibilityStudyDD { get; set; }
        public string ProjectCharterDD { get; set; }
        public string JobDescriptionDD { get; set; }
        public string ProjectOfficeCheckListDD { get; set; }
        public string PhaseRevieFormInitiationDD { get; set; }
        public string TermOfReferenceDocumentDD { get; set; }

        // Planned Budget for tasks
        public string BusinessCasePlannedBudget { get; set; }
        public string FeasibilityStudyPlannedBudget { get; set; }
        public string ProjectCharterPlannedBudget { get; set; }
        public string JobDescriptionPlannedBudget { get; set; }
        public string ProjectOfficeCheckListPlannedBudget { get; set; }
        public string PhaseRevieFormInitiationPlannedBudget { get; set; }
        public string TermOfReferenceDocumentPlannedBudget { get; set; }

        // Budget used for tasks
        public string BusinessCaseBudget { get; set; }
        public string FeasibilityStudyBudget { get; set; }
        public string ProjectCharterBudget { get; set; }
        public string JobDescriptionBudget { get; set; }
        public string ProjectOfficeCheckListBudget { get; set; }
        public string PhaseRevieFormInitiationBudget { get; set; }
        public string TermOfReferenceDocumentBudget { get; set; }


    }
}



