﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProjectManagementToolkit.MPMM.MPMM_Document_Models
{
    class CommunicationsPlanModel
    {
        public List<Stakeholder> StakeholderReq { get; set; }

        public string ProjectName { get; set; }
        public string DocumentID { get; set; }
        public string CommunicationPlanProgress{ get; set; }
        public string DocumentOwner { get; set; }
        public string IssueDate { get; set; }
        public string LastSavedDate { get; set; }
        public string FileName { get; set; }
        public string StakeholderList { get; set; }
        public List<DocumentHistory> DocumentHistories { get; set; }
        public List<DocumentApproval> DocumentApprovals { get; set; }
        public List<ProjectSchedule> ProSchedule { get; set; }

        public string ComPlan { get; set; }
        public string Assumptions { get; set; }

        public string ComProcess { get; set; }
        public string Activities { get; set; }
        public string Roles { get; set; }
        public string Documents { get; set; }

        public class ProjectSchedule
        {
            public string Deliverable { get; set; }
            public string ScheduledCompletionDate { get; set; }
            public string ActualCompletionDate { get; set; }
            public string ActualVariance { get; set; }
            public string ForecastCompletionDate { get; set; }
            public string ForecastVariance { get; set; }
            public string Summary { get; set; }
        }

        public class DocumentHistory
        {
            public string Version { get; set; }
            public string IssueDate { get; set; }
            public string Changes { get; set; }
        }
        public class DocumentApproval
        {
            public string Role { get; set; }
            public string Name { get; set; }
            public string Signature { get; set; }
            public string DateApproved { get; set; }
        }

        public class Stakeholder
        {
            public string StakeholderName { get; set; }
            public string StakeholderRole { get; set; }
            public string StakeholderOrganization { get; set; }
            public string InformationRequirement { get; set; }
        }

        public class Phase
        {
            public string PhaseTitle { get; set; }
            public string PhaseDescription { get; set; }
            public string PhaseSequence { get; set; }
        }
    }
}
