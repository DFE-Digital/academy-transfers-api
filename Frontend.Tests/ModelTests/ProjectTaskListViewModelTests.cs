using Data.Models;
using Data.Models.Projects;
using Frontend.Models;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Data.Models.KeyStagePerformance;
using DocumentFormat.OpenXml.Office2013.PowerPoint.Roaming;
using Xunit;
using static Data.Models.Projects.TransferFeatures;

namespace Frontend.Tests.ModelTests
{
    public class ProjectTaskListViewModelTests
    {
        public class TransferFeatureTests
        {
            [Fact]
            public void GivenTransferFeatureDataIsEmpty_ReturnsNotStarted()
            {
                // Arrange
                var project = new Project
                {
                    Features = new TransferFeatures
                    {
                        WhoInitiatedTheTransfer = TransferFeatures.ProjectInitiators.Empty,
                        ReasonForTransfer = new ReasonForTransfer
                        {
                            IsSubjectToRddOrEsfaIntervention = null,
                        },
                        TypeOfTransfer = TransferFeatures.TransferTypes.Empty
                    }
                };
                var model = new ProjectTaskListViewModel {Project = project};

                // Act
                var result = model.FeatureTransferStatus;

                // Assert
                Assert.Equal(ProjectStatuses.NotStarted, result);
            }

            [Fact]
            public void GivenTransferFeatureDataIsFullyCompleted_ReturnCompleted()
            {
                // Arrange
                var project = new Project
                {
                    Features = new TransferFeatures
                    {
                        WhoInitiatedTheTransfer = TransferFeatures.ProjectInitiators.Dfe,
                        ReasonForTransfer = new ReasonForTransfer
                        {
                            IsSubjectToRddOrEsfaIntervention = false,
                        },
                        TypeOfTransfer = TransferFeatures.TransferTypes.MatClosure
                    }
                };
                var model = new ProjectTaskListViewModel {Project = project};

                // Act
                var result = model.FeatureTransferStatus;

                // Assert
                Assert.Equal(ProjectStatuses.Completed, result);
            }

            [Theory]
            [InlineData(ProjectInitiators.OutgoingTrust, null, TransferTypes.Empty)]
            [InlineData(ProjectInitiators.OutgoingTrust, false, TransferTypes.Empty)]
            [InlineData(ProjectInitiators.OutgoingTrust, null, TransferTypes.JoiningToFormMat)]
            [InlineData(ProjectInitiators.Empty, null, TransferTypes.JoiningToFormMat)]
            [InlineData(ProjectInitiators.Empty, true, TransferTypes.Empty)]
            [InlineData(ProjectInitiators.Empty, true, TransferTypes.JoiningToFormMat)]
            public void GivenTransferFeatureDataIsPartiallyCompleted_ReturnInProgress(
                ProjectInitiators projectInitiator,
                bool? isSubjectToRddOrEsfaIntervention,
                TransferTypes transferType)
            {
                // Arrange
                var project = new Project
                {
                    Features = new TransferFeatures
                    {
                        WhoInitiatedTheTransfer = projectInitiator,
                        ReasonForTransfer = new ReasonForTransfer
                            {IsSubjectToRddOrEsfaIntervention = isSubjectToRddOrEsfaIntervention},
                        TypeOfTransfer = transferType
                    }
                };
                var model = new ProjectTaskListViewModel {Project = project};

                // Act
                var result = model.FeatureTransferStatus;

                // Assert
                Assert.Equal(ProjectStatuses.InProgress, result);
            }
        }

        public class HasKeyStage5PerformanceDataTests
        {
            [Fact]
            public void GivenNull_ReturnFalse()
            {
                var model = new ProjectTaskListViewModel {EducationPerformance = new EducationPerformance()};
                Assert.False(model.HasKeyStage5PerformanceInformation);
            }

            [Fact]
            public void GivenEmptyList_ReturnFalse()
            {
                var model = new ProjectTaskListViewModel
                    {EducationPerformance = new EducationPerformance() {KeyStage5Performance = new List<KeyStage5>()}};
                Assert.False(model.HasKeyStage5PerformanceInformation);
            }

            [Fact]
            public void GivenListWithItemWithNoAcademy_ReturnFalse()
            {
                var model = new ProjectTaskListViewModel
                {
                    EducationPerformance = new EducationPerformance
                        {KeyStage5Performance = new List<KeyStage5> {new KeyStage5()}}
                };
                Assert.False(model.HasKeyStage5PerformanceInformation);
            }

            [Theory]
            [InlineData("12.12", null)]
            [InlineData("12.12", "")]
            [InlineData(null, "12.12")]
            [InlineData("", "12.12")]
            [InlineData("12.12", "12.12")]
            public void GivenListWithItemWithAcademyWithData_ReturnTrue(string academicAverage,
                string appliedGeneralAverage)
            {
                var model = new ProjectTaskListViewModel
                {
                    EducationPerformance = new EducationPerformance
                    {
                        KeyStage5Performance = new List<KeyStage5>
                        {
                            new KeyStage5
                            {
                                Academy = new KeyStage5Result
                                {
                                    AcademicAverage = academicAverage,
                                    AppliedGeneralAverage = appliedGeneralAverage
                                }
                            }
                        }
                    }
                };
                Assert.True(model.HasKeyStage5PerformanceInformation);
            }
        }

        public class TransferDatesTests
        {
            [Fact]
            public void GivenTransferDatesAreEmpty_ReturnNotStarted()
            {
                // Arrange
                var project = new Project
                {
                    Dates = new TransferDates
                    {
                        FirstDiscussed = string.Empty,
                        Target = string.Empty,
                        Htb = string.Empty
                    }
                };
                var model = new ProjectTaskListViewModel {Project = project};

                // Act
                var result = model.TransferDatesStatus;

                // Assert
                Assert.Equal(ProjectStatuses.NotStarted, result);
            }

            [Fact]
            public void GivenTransferDatesAreNull_ReturnNotStarted()
            {
                // Arrange
                var project = new Project
                {
                    Dates = new TransferDates
                    {
                        FirstDiscussed = null,
                        Target = null,
                        Htb = null
                    }
                };
                var model = new ProjectTaskListViewModel {Project = project};

                // Act
                var result = model.TransferDatesStatus;

                // Assert
                Assert.Equal(ProjectStatuses.NotStarted, result);
            }

            [Theory]
            [InlineData("test", "", "")]
            [InlineData("test", "test", "")]
            [InlineData("test", "", "test")]
            [InlineData("", "", "test")]
            [InlineData("", "test", "")]
            [InlineData("", "test", "test")]
            public void GivenTransferDatesArePartiallyCompleted_ReturnInProgress(string firstDiscussed,
                string targetDate, string htbDate)
            {
                // Arrange
                var project = new Project
                {
                    Dates = new TransferDates
                    {
                        FirstDiscussed = firstDiscussed,
                        Target = targetDate,
                        Htb = htbDate
                    }
                };
                var model = new ProjectTaskListViewModel {Project = project};

                // Act
                var result = model.TransferDatesStatus;

                // Assert
                Assert.Equal(ProjectStatuses.InProgress, result);
            }

            [Fact]
            public void GivenTransferDatesAreCompleted_ReturnCompleted()
            {
                // Arrange
                var project = new Project
                {
                    Dates = new TransferDates
                    {
                        FirstDiscussed = "test",
                        Target = "test",
                        Htb = "test"
                    }
                };
                var model = new ProjectTaskListViewModel {Project = project};

                // Act
                var result = model.TransferDatesStatus;

                // Assert
                Assert.Equal(ProjectStatuses.Completed, result);
            }
        }

        public class BenefitsAndOtherFactorsTests
        {
            [Fact]
            public void GivenBenefitsAreEmpty_ReturnNotStarted()
            {
                // Arrange
                var project = new Project
                {
                    Benefits = new TransferBenefits
                    {
                        IntendedBenefits = new List<TransferBenefits.IntendedBenefit>(),
                        OtherFactors = new Dictionary<TransferBenefits.OtherFactor, string>()
                    }
                };
                var model = new ProjectTaskListViewModel {Project = project};

                // Act
                var result = model.BenefitsAndOtherFactorsStatus;

                // Assert
                Assert.Equal(ProjectStatuses.NotStarted, result);
            }

            [Fact]
            public void GivenIntendedBenefitIsCompletedAndOtherFactorsIsEmpty_ReturnInProgress()
            {
                // Arrange
                var project = new Project
                {
                    Benefits = new TransferBenefits
                    {
                        IntendedBenefits = new List<TransferBenefits.IntendedBenefit>
                            {TransferBenefits.IntendedBenefit.CentralFinanceTeamAndSupport},
                        OtherFactors = new Dictionary<TransferBenefits.OtherFactor, string>()
                    }
                };
                var model = new ProjectTaskListViewModel {Project = project};

                // Act
                var result = model.BenefitsAndOtherFactorsStatus;

                // Assert
                Assert.Equal(ProjectStatuses.InProgress, result);
            }

            [Fact]
            public void GivenIntendedBenefitIsEmptyAndOtherFactorsIsCompleted_ReturnInProgress()
            {
                // Arrange
                var project = new Project
                {
                    Benefits = new TransferBenefits
                    {
                        IntendedBenefits = new List<TransferBenefits.IntendedBenefit>(),
                        OtherFactors = new Dictionary<TransferBenefits.OtherFactor, string>
                            {{TransferBenefits.OtherFactor.HighProfile, "test"}}
                    }
                };
                var model = new ProjectTaskListViewModel {Project = project};

                // Act
                var result = model.BenefitsAndOtherFactorsStatus;

                // Assert
                Assert.Equal(ProjectStatuses.InProgress, result);
            }

            [Fact]
            public void GivenBenefitsAreNull_ReturnNotStarted()
            {
                // Arrange
                var project = new Project
                {
                    Benefits = new TransferBenefits
                    {
                        IntendedBenefits = null,
                        OtherFactors = null
                    }
                };
                var model = new ProjectTaskListViewModel {Project = project};

                // Act
                var result = model.BenefitsAndOtherFactorsStatus;

                // Assert
                Assert.Equal(ProjectStatuses.NotStarted, result);
            }

            [Fact]
            public void GivenBenefitsAreCompleted_ReturnCompleted()
            {
                // Arrange
                var project = new Project
                {
                    Benefits = new TransferBenefits
                    {
                        IntendedBenefits = new List<TransferBenefits.IntendedBenefit>
                            {TransferBenefits.IntendedBenefit.CentralFinanceTeamAndSupport},
                        OtherFactors = new Dictionary<TransferBenefits.OtherFactor, string>
                            {{TransferBenefits.OtherFactor.HighProfile, "test"}}
                    }
                };
                var model = new ProjectTaskListViewModel {Project = project};

                // Act
                var result = model.BenefitsAndOtherFactorsStatus;

                // Assert
                Assert.Equal(ProjectStatuses.Completed, result);
            }
        }

        public class RationaleTests
        {
            [Fact]
            public void GivenRationaleIsEmpty_ReturnNotStarted()
            {
                // Arrange
                var project = new Project
                {
                    Rationale = new TransferRationale
                    {
                        Project = string.Empty,
                        Trust = string.Empty
                    }
                };
                var model = new ProjectTaskListViewModel {Project = project};

                // Act
                var result = model.RationaleStatus;

                // Assert
                Assert.Equal(ProjectStatuses.NotStarted, result);
            }

            [Theory]
            [InlineData("some value", "")]
            [InlineData("some value", null)]
            [InlineData("", "some value")]
            [InlineData(null, "some value")]
            public void GivenRationaleIsPartiallyCompleted_ReturnInProgress(string projectValue, string trustValue)
            {
                // Arrange
                var project = new Project
                {
                    Rationale = new TransferRationale
                    {
                        Project = projectValue,
                        Trust = trustValue
                    }
                };
                var model = new ProjectTaskListViewModel {Project = project};

                // Act
                var result = model.RationaleStatus;

                // Assert
                Assert.Equal(ProjectStatuses.InProgress, result);
            }

            [Fact]
            public void GivenRationaleIsCompleted_ReturnCompleted()
            {
                // Arrange
                var project = new Project
                {
                    Rationale = new TransferRationale
                    {
                        Project = "some value",
                        Trust = "some value"
                    }
                };
                var model = new ProjectTaskListViewModel {Project = project};

                // Act
                var result = model.RationaleStatus;

                // Assert
                Assert.Equal(ProjectStatuses.Completed, result);
            }
        }

        public class AcademyAndTrustInformationTests
        {
            [Theory]
            [InlineData(TransferAcademyAndTrustInformation.RecommendationResult.Empty, null,
                ProjectStatuses.NotStarted)]
            [InlineData(TransferAcademyAndTrustInformation.RecommendationResult.Approve, "test author",
                ProjectStatuses.Completed)]
            [InlineData(TransferAcademyAndTrustInformation.RecommendationResult.Empty, "test author",
                ProjectStatuses.InProgress)]
            [InlineData(TransferAcademyAndTrustInformation.RecommendationResult.Decline, null,
                ProjectStatuses.InProgress)]
            [InlineData(TransferAcademyAndTrustInformation.RecommendationResult.Decline, "",
                ProjectStatuses.InProgress)]
            public void GivenRelevantData_ReturnCorrectStatus(
                TransferAcademyAndTrustInformation.RecommendationResult recommendation, string author,
                ProjectStatuses expectedStatus)
            {
                var project = new Project
                {
                    AcademyAndTrustInformation = new TransferAcademyAndTrustInformation
                    {
                        Recommendation = recommendation,
                        Author = author
                    }
                };
                var model = new ProjectTaskListViewModel {Project = project};
                var result = model.AcademyAndTrustInformationStatus();

                Assert.Equal(expectedStatus, result);
            }
        }
    }
}