﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProjectManagementToolkit.Classes
{
    class PhaseReviewPlanningModel
    {
        public string PlanningPhase { get; set; }
        public string ProjectName { get; set; }
        public string PhaseReviewPlanningProgress { get; set; }
        public string ProjectManager { get; set; }
        public string RepportPreparedBy { get; set; }
        public string ReportPrepDate { get; set; }
        public string ReportingPeriod { get; set; }
        public string ProjectSponsor { get; set; }
        public string Summary { get; set; }
        public string ProjectSchedule { get; set; }
        public string ProjectExpense { get; set; }
        public string ProjectDeliverables { get; set; }
        public string ProjectRisks { get; set; }
        public string ProjectIssues { get; set; }
        public string ProjectChanges { get; set; }
        public string SupportingDocumentation { get; set; }
        public List<ReviewDetails> Reviews { get; set; }
    }

    class ReviewDetails
    {
        public string ReviewCategory { get; set; }
        public string ReviewQuestion { get; set; }
        public string Answer { get; set; }
        public string Variance { get; set; }
    }
}
