using System.Collections.Generic;
using System.Linq;
using Data.Models;
using Data.Models.Projects;
using Data.TRAMS.Mappers.Response;
using Data.TRAMS.Models;
using Data.TRAMS.Models.AcademyTransferProject;
using Helpers;
using Xunit;

namespace Data.TRAMS.Tests.Mappers.Response
{
    public class TramsProjectMapperTests
    {
        [Fact]
        public void GivenTramsProject_MapsCorrectly()
        {
            var toMap = new TramsProject
            {
                Benefits = new AcademyTransferProjectBenefits
                {
                    IntendedTransferBenefits = new IntendedTransferBenefits
                    {
                        SelectedBenefits = new List<string>
                        {
                            TransferBenefits.IntendedBenefit.ImprovingSafeguarding.ToString(),
                            TransferBenefits.IntendedBenefit.Other.ToString()
                        },
                        OtherBenefitValue = "Other benefit value"
                    },
                    OtherFactorsToConsider = new OtherFactorsToConsider
                    {
                        HighProfile = new OtherFactor
                            {FurtherSpecification = "High profile", ShouldBeConsidered = true},
                        FinanceAndDebt = new OtherFactor
                            {FurtherSpecification = "Finance", ShouldBeConsidered = true},
                        ComplexLandAndBuilding = new OtherFactor
                            {FurtherSpecification = "Complex land and building", ShouldBeConsidered = true}
                    }
                },
                Dates = new AcademyTransferProjectDates
                {
                    HtbDate = "01/01/2020",
                    TargetDateForTransfer = "02/01/2020",
                    TransferFirstDiscussed = "03/01/2020"
                },
                Features = new AcademyTransferProjectFeatures
                {
                    TypeOfTransfer = TransferFeatures.TransferTypes.TrustsMerging.ToString(),
                    OtherTransferTypeDescription = "Other",
                    RddOrEsfaIntervention = true,
                    WhoInitiatedTheTransfer = TransferFeatures.ProjectInitiators.Dfe.ToString(),
                    RddOrEsfaInterventionDetail = "Intervention details"
                },
                Rationale = new AcademyTransferProjectRationale
                {
                    ProjectRationale = "Project rationale",
                    TrustSponsorRationale = "Trust rationale"
                },
                GeneralInformation = new AcademyTransferProjectAcademyAndTrustInformation
                {
                    Author = "Author",
                    Recommendation = TransferAcademyAndTrustInformation.RecommendationResult.Approve.ToString()
                },
                State = "State",
                Status = "Status",
                OutgoingTrust = new TrustSummary
                {
                    GroupName = "Outgoing trust name",
                    Ukprn = "123"
                },
                ProjectUrn = "Urn",
                TransferringAcademies = new List<TransferringAcademy>
                {
                    new TransferringAcademy
                    {
                        IncomingTrust = new TrustSummary
                        {
                            GroupName = "Incoming trust",
                            Ukprn = "456"
                        },
                        OutgoingAcademy = new AcademySummary
                        {
                            Name = "Outgoing academy",
                            Urn = "789",
                            Ukprn = "987"
                        }
                    }
                },
                AcademyPerformanceAdditionalInformation = "AcademyPerformanceAdditionalInformation",
                PupilNumbersAdditionalInformation = "PupilNumbersAdditionalInformation",
                LatestOfstedJudgementAdditionalInformation = "LatestOfstedJudgementAdditionalInformation",
                KeyStage2PerformanceAdditionalInformation = "KeyStage2PerformanceAdditionalInformation",
                KeyStage4PerformanceAdditionalInformation = "KeyStage4PerformanceAdditionalInformation",
                KeyStage5PerformanceAdditionalInformation = "KeyStage5PerformanceAdditionalInformation",
                OutgoingTrustUkprn = "123"
            };

            var subject = new TramsProjectMapper();
            var result = subject.Map(toMap);

            Assert.Equal(toMap.State, result.State);
            Assert.Equal(toMap.Status, result.Status);
            Assert.Equal(toMap.ProjectUrn, result.Urn);
            Assert.Equal(toMap.OutgoingTrustUkprn, result.OutgoingTrustUkprn);
            Assert.Equal(toMap.OutgoingTrust.GroupName, result.OutgoingTrustName);
            Assert.Equal(toMap.AcademyPerformanceAdditionalInformation, result.AcademyPerformanceAdditionalInformation);
            Assert.Equal(toMap.PupilNumbersAdditionalInformation, result.PupilNumbersAdditionalInformation);
            Assert.Equal(toMap.LatestOfstedJudgementAdditionalInformation, result.LatestOfstedJudgementAdditionalInformation);
            Assert.Equal(toMap.KeyStage2PerformanceAdditionalInformation, result.KeyStage2PerformanceAdditionalInformation);
            Assert.Equal(toMap.KeyStage4PerformanceAdditionalInformation, result.KeyStage4PerformanceAdditionalInformation);
            Assert.Equal(toMap.KeyStage5PerformanceAdditionalInformation, result.KeyStage5PerformanceAdditionalInformation);

            AssertBenefitsCorrect(toMap, result);
            AssertDatesCorrect(toMap, result);
            AssertFeaturesCorrect(toMap, result);
            AssertRationaleCorrect(toMap, result);
            AssertGeneralInformationCorrect(toMap, result);
            AssertTransferringAcademiesCorrect(toMap, result);
        }

        private void AssertGeneralInformationCorrect(TramsProject toMap, Project result)
        {
            Assert.Equal(toMap.GeneralInformation.Author, result.AcademyAndTrustInformation.Author);
            var expectedRecommendation =
                EnumHelpers<TransferAcademyAndTrustInformation.RecommendationResult>.Parse(toMap.GeneralInformation.Recommendation);
            Assert.Equal(expectedRecommendation, result.AcademyAndTrustInformation.Recommendation);
        }

        private static void AssertTransferringAcademiesCorrect(TramsProject toMap, Project result)
        {
            var expectedTransfer = toMap.TransferringAcademies[0];
            Assert.Equal(expectedTransfer.IncomingTrust.GroupName, result.TransferringAcademies[0].IncomingTrustName);
            Assert.Equal(expectedTransfer.IncomingTrust.Ukprn, result.TransferringAcademies[0].IncomingTrustUkprn);
            Assert.Equal(expectedTransfer.OutgoingAcademy.Name, result.TransferringAcademies[0].OutgoingAcademyName);
            Assert.Equal(expectedTransfer.OutgoingAcademy.Ukprn, result.TransferringAcademies[0].OutgoingAcademyUkprn);
            Assert.Equal(expectedTransfer.OutgoingAcademy.Urn, result.TransferringAcademies[0].OutgoingAcademyUrn);
        }

        private static void AssertRationaleCorrect(TramsProject toMap, Project result)
        {
            Assert.Equal(toMap.Rationale.ProjectRationale, result.Rationale.Project);
            Assert.Equal(toMap.Rationale.TrustSponsorRationale, result.Rationale.Trust);
        }

        private static void AssertFeaturesCorrect(TramsProject toMap, Project result)
        {
            var expectedType = EnumHelpers<TransferFeatures.TransferTypes>.Parse(toMap.Features.TypeOfTransfer);
            var expectedInitiator =
                EnumHelpers<TransferFeatures.ProjectInitiators>.Parse(toMap.Features.WhoInitiatedTheTransfer);
            Assert.Equal(expectedType, result.Features.TypeOfTransfer);
            Assert.Equal(toMap.Features.OtherTransferTypeDescription, result.Features.OtherTypeOfTransfer);
            Assert.Equal(expectedInitiator, result.Features.WhoInitiatedTheTransfer);
            Assert.Equal(toMap.Features.RddOrEsfaIntervention,
                result.Features.ReasonForTransfer.IsSubjectToRddOrEsfaIntervention);
            Assert.Equal(toMap.Features.RddOrEsfaInterventionDetail,
                result.Features.ReasonForTransfer.InterventionDetails);
        }

        private static void AssertDatesCorrect(TramsProject toMap, Project result)
        {
            Assert.Equal(toMap.Dates.HtbDate, result.Dates.Htb);
            Assert.Equal(toMap.Dates.TransferFirstDiscussed, result.Dates.FirstDiscussed);
            Assert.Equal(toMap.Dates.TargetDateForTransfer, result.Dates.Target);
        }

        private static void AssertBenefitsCorrect(TramsProject toMap, Project result)
        {
            var expectedBenefitsList = toMap.Benefits.IntendedTransferBenefits.SelectedBenefits
                .Select(EnumHelpers<TransferBenefits.IntendedBenefit>.Parse)
                .ToList();
            Assert.Equal(expectedBenefitsList, result.Benefits.IntendedBenefits);
            Assert.Equal(toMap.Benefits.IntendedTransferBenefits.OtherBenefitValue,
                result.Benefits.OtherIntendedBenefit);
            Assert.Equal(toMap.Benefits.OtherFactorsToConsider.HighProfile.FurtherSpecification,
                result.Benefits.OtherFactors[TransferBenefits.OtherFactor.HighProfile]);
            Assert.Equal(toMap.Benefits.OtherFactorsToConsider.FinanceAndDebt.FurtherSpecification,
                result.Benefits.OtherFactors[TransferBenefits.OtherFactor.FinanceAndDebtConcerns]);
            Assert.Equal(toMap.Benefits.OtherFactorsToConsider.ComplexLandAndBuilding.FurtherSpecification,
                result.Benefits.OtherFactors[TransferBenefits.OtherFactor.ComplexLandAndBuildingIssues]);
        }
    }
}