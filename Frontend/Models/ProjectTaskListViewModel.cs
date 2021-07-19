using Data.Models;
using Data.Models.Projects;
using System.Linq;
using Data.Models.KeyStagePerformance;

namespace Frontend.Models
{
    public class ProjectTaskListViewModel : ProjectViewModel
    {
        public EducationPerformance EducationPerformance { get; set; }
        public ProjectStatuses FeatureTransferStatus => GetFeatureTransferStatus();
        public ProjectStatuses TransferDatesStatus => GetTransferDatesStatus();
        public ProjectStatuses BenefitsAndOtherFactorsStatus => GetBenefitsAndOtherFactorsStatus();
        public ProjectStatuses RationaleStatus => GetRationaleStatus();
        public ProjectStatuses AcademyAndTrustInformationStatus() => GetAcademyAndTrustInformationStatus();

        public bool HasKeyStage5PerformanceInformation
        {
            get
            {
                return EducationPerformance.KeyStage5Performance != null &&
                       EducationPerformance.KeyStage5Performance.Any(ks5 =>
                           ks5.Academy != null && (
                               !string.IsNullOrEmpty(ks5.Academy.AcademicAverage) ||
                               !string.IsNullOrEmpty(ks5.Academy.AppliedGeneralAverage)
                           )
                       );
            }
        }

        private ProjectStatuses GetAcademyAndTrustInformationStatus()
        {
            var academyAndTrustInformation = Project.AcademyAndTrustInformation;
            return (string.IsNullOrEmpty(academyAndTrustInformation.Author) &&
                    (academyAndTrustInformation.Recommendation ==
                     TransferAcademyAndTrustInformation.RecommendationResult.Empty)) ? ProjectStatuses.NotStarted
                : (!string.IsNullOrEmpty(academyAndTrustInformation.Author) &&
                   (academyAndTrustInformation.Recommendation !=
                    TransferAcademyAndTrustInformation.RecommendationResult.Empty)) ? ProjectStatuses.Completed
                : ProjectStatuses.InProgress;
        }

        private ProjectStatuses GetFeatureTransferStatus()
        {
            if (Project.Features.WhoInitiatedTheTransfer == TransferFeatures.ProjectInitiators.Empty &&
                Project.Features.ReasonForTransfer.IsSubjectToRddOrEsfaIntervention == null &&
                Project.Features.TypeOfTransfer == TransferFeatures.TransferTypes.Empty)
                return ProjectStatuses.NotStarted;

            if (Project.Features.WhoInitiatedTheTransfer != TransferFeatures.ProjectInitiators.Empty &&
                Project.Features.ReasonForTransfer.IsSubjectToRddOrEsfaIntervention != null &&
                Project.Features.TypeOfTransfer != TransferFeatures.TransferTypes.Empty)
                return ProjectStatuses.Completed;

            return ProjectStatuses.InProgress;
        }

        private ProjectStatuses GetTransferDatesStatus()
        {
            if (string.IsNullOrEmpty(Project.Dates.FirstDiscussed) &&
                string.IsNullOrEmpty(Project.Dates.Target) &&
                string.IsNullOrEmpty(Project.Dates.Htb))
                return ProjectStatuses.NotStarted;

            if (!string.IsNullOrEmpty(Project.Dates.FirstDiscussed) &&
                !string.IsNullOrEmpty(Project.Dates.Target) &&
                !string.IsNullOrEmpty(Project.Dates.Htb))
                return ProjectStatuses.Completed;

            return ProjectStatuses.InProgress;
        }

        private ProjectStatuses GetBenefitsAndOtherFactorsStatus()
        {
            if ((Project.Benefits.IntendedBenefits == null || !Project.Benefits.IntendedBenefits.Any()) &&
                (Project.Benefits.OtherFactors == null || !Project.Benefits.OtherFactors.Any()))
                return ProjectStatuses.NotStarted;

            if ((Project.Benefits.IntendedBenefits != null && Project.Benefits.IntendedBenefits.Any()) &&
                (Project.Benefits.OtherFactors != null && Project.Benefits.OtherFactors.Any()))
                return ProjectStatuses.Completed;

            return ProjectStatuses.InProgress;
        }

        private ProjectStatuses GetRationaleStatus()
        {
            if (string.IsNullOrEmpty(Project.Rationale.Project) &&
                string.IsNullOrEmpty(Project.Rationale.Trust))
                return ProjectStatuses.NotStarted;

            if (!string.IsNullOrEmpty(Project.Rationale.Project) &&
                !string.IsNullOrEmpty(Project.Rationale.Trust))
                return ProjectStatuses.Completed;

            return ProjectStatuses.InProgress;
        }
    }
}