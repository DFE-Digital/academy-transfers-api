using System;
using Data;
using Data.Models;
using Frontend.Controllers.Projects;
using Frontend.Helpers;
using Frontend.Models;
using Frontend.Tests.Helpers;
using Helpers;
using Microsoft.AspNetCore.Mvc;
using Moq;
using Xunit;

namespace Frontend.Tests.ControllerTests.Projects
{
    public class TransferDatesControllerTests
    {
        private const string _errorWithGetByUrn = "errorUrn";
        private readonly TransferDatesController _subject;
        private readonly Mock<IProjects> _projectsRepository;
        private readonly Mock<IDateTimeProvider> _dateTimeProvider;
        private readonly Project _foundProject;

        public TransferDatesControllerTests()
        {
            _foundProject = new Project()
            {
                Urn = "0001"
            };

            _projectsRepository = new Mock<IProjects>();
            _dateTimeProvider = new Mock<IDateTimeProvider>();

            _projectsRepository.Setup(r => r.GetByUrn(It.IsAny<string>()))
                .ReturnsAsync(new RepositoryResult<Project> {Result = _foundProject});
            _projectsRepository.Setup(s => s.GetByUrn(_errorWithGetByUrn))
                .ReturnsAsync(
                  new RepositoryResult<Project>
                  {
                      Error = new RepositoryResultBase.RepositoryError
                      {
                          ErrorMessage = "Error"
                      }
                  });

            _projectsRepository.Setup(r => r.Update(It.IsAny<Project>()))
                .ReturnsAsync(new RepositoryResult<Project>());
            _dateTimeProvider.Setup(r => r.Today()).Returns(new DateTime(2020, 1, 1));

            _subject = new TransferDatesController(_projectsRepository.Object, _dateTimeProvider.Object);
        }

        public class IndexTests : TransferDatesControllerTests
        {
            [Fact]
            public async void GivenUrn_AssignsModelToTheView()
            {
                var result = await _subject.Index("0001");
                var viewModel = ControllerTestHelpers.GetViewModelFromResult<TransferDatesViewModel>(result);

                Assert.Equal(_foundProject.Urn, viewModel.Project.Urn);
            }

            [Fact]
            public async void GivenGetByUrnReturnsError_DisplayErrorPage()
            {
                var response = await _subject.Index(_errorWithGetByUrn);
                var viewResult = Assert.IsType<ViewResult>(response);

                Assert.Equal("ErrorPage", viewResult.ViewName);
                Assert.Equal("Error", viewResult.Model);
            }
        }

        public class FirstDiscussedTests : TransferDatesControllerTests
        {
            public class GetTests : FirstDiscussedTests
            {
                [Fact]
                public async void GivenUrn_AssignsModelToTheView()
                {
                    var result = await _subject.FirstDiscussed("0001");
                    var viewModel = ControllerTestHelpers.GetViewModelFromResult<TransferDatesViewModel>(result);

                    Assert.Equal(_foundProject.Urn, viewModel.Project.Urn);
                }

                [Fact]
                public async void GivenGetByUrnReturnsError_DisplayErrorPage()
                {
                    var response = await _subject.FirstDiscussed(_errorWithGetByUrn);
                    var viewResult = Assert.IsType<ViewResult>(response);

                    Assert.Equal("ErrorPage", viewResult.ViewName);
                    Assert.Equal("Error", viewResult.Model);
                }
            }

            public class PostTests : FirstDiscussedTests
            {
                [Theory]
                [InlineData("01", "01", "2020", "01/01/2020")]
                [InlineData("20", "02", "2021", "20/02/2021")]
                [InlineData("2", "2", "2021", "02/02/2021")]
                public async void GivenUrnAndFullDate_UpdatesTheProjectWithTheCorrectDate(string day, string month,
                    string year, string expectedDate)
                {
                    await _subject.FirstDiscussedPost("0001", day, month, year);

                    _projectsRepository.Verify(r =>
                        r.Update(It.Is<Project>(project => project.Dates.FirstDiscussed == expectedDate)));
                }

                [Fact]
                public async void GivenUrnAndFullDate_RedirectsToTheSummaryPage()
                {
                    var response = await _subject.FirstDiscussedPost("0001", "01", "01", "2020");
                    ControllerTestHelpers.AssertResultRedirectsToAction(response, "Index");
                }

                [Theory]
                [InlineData(null, "1", "2020")]
                [InlineData("1", null, "2020")]
                [InlineData("1", "1", null)]
                public async void GivenPartsOfDateMissing_SetErrorOnTheModel(string day, string month, string year)
                {
                    var response = await _subject.FirstDiscussedPost("0001", day, month, year);
                    var responseModel = ControllerTestHelpers.GetViewModelFromResult<TransferDatesViewModel>(response);

                    Assert.True(responseModel.FormErrors.HasErrors);
                    Assert.Equal("Please enter the date the transfer was first discussed",
                        responseModel.FormErrors.Errors[0].ErrorMessage);
                }

                [Theory]
                [InlineData("0", "1", "2020")]
                [InlineData("32", "1", "2020")]
                [InlineData("1", "0", "2020")]
                [InlineData("1", "13", "2020")]
                [InlineData("1", "13", "0")]
                public async void GivenInvalidDate_SetErrorOnTheModel(string day, string month, string year)
                {
                    var response = await _subject.FirstDiscussedPost("0001", day, month, year);
                    var responseModel = ControllerTestHelpers.GetViewModelFromResult<TransferDatesViewModel>(response);

                    Assert.True(responseModel.FormErrors.HasErrors);
                    Assert.Equal("Please enter a valid date", responseModel.FormErrors.Errors[0].ErrorMessage);
                }

                [Fact]
                public async void GivenGetByUrnReturnsError_DisplayErrorPage()
                {
                    var response = await _subject.FirstDiscussedPost(_errorWithGetByUrn, "01", "01", "2020");
                    var viewResult = Assert.IsType<ViewResult>(response);

                    Assert.Equal("ErrorPage", viewResult.ViewName);
                    Assert.Equal("Error", viewResult.Model);
                }

                [Fact]
                public async void GivenUpdatenReturnsError_DisplayErrorPage()
                {
                    _projectsRepository.Setup(s => s.Update(It.IsAny<Project>())).ReturnsAsync(
                       new RepositoryResult<Project>
                       {
                           Error = new RepositoryResultBase.RepositoryError
                           {
                               ErrorMessage = "Update error"
                           }
                       });

                    var response = await _subject.FirstDiscussedPost("0001", "01", "01", "2020");
                    var viewResult = Assert.IsType<ViewResult>(response);

                    Assert.Equal("ErrorPage", viewResult.ViewName);
                    Assert.Equal("Update error", viewResult.Model);
                }
            }
        }

        public class TargetDateTests : TransferDatesControllerTests
        {
            public class GetTests : TargetDateTests
            {
                [Fact]
                public async void GivenUrn_AssignsModelToTheView()
                {
                    var result = await _subject.TargetDate("0001");
                    var viewModel = ControllerTestHelpers.GetViewModelFromResult<TransferDatesViewModel>(result);

                    Assert.Equal(_foundProject.Urn, viewModel.Project.Urn);
                }

                [Fact]
                public async void GivenGetByUrnReturnsError_DisplayErrorPage()
                {
                    var response = await _subject.TargetDate(_errorWithGetByUrn);
                    var viewResult = Assert.IsType<ViewResult>(response);

                    Assert.Equal("ErrorPage", viewResult.ViewName);
                    Assert.Equal("Error", viewResult.Model);
                }
            }

            public class PostTests : TargetDateTests
            {
                [Theory]
                [InlineData("01", "01", "2020", "01/01/2020")]
                [InlineData("20", "02", "2021", "20/02/2021")]
                [InlineData("2", "2", "2021", "02/02/2021")]
                public async void GivenUrnAndFullDate_UpdatesTheProjectWithTheCorrectDate(string day, string month,
                    string year, string expectedDate)
                {
                    await _subject.TargetDatePost("0001", day, month, year);

                    _projectsRepository.Verify(r =>
                        r.Update(It.Is<Project>(project => project.Dates.Target == expectedDate)));
                }

                [Fact]
                public async void GivenUrnAndFullDate_RedirectsToTheSummaryPage()
                {
                    var response = await _subject.TargetDatePost("0001", "01", "01", "2020");
                    ControllerTestHelpers.AssertResultRedirectsToAction(response, "Index");
                }

                [Theory]
                [InlineData(null, "1", "2020")]
                [InlineData("1", null, "2020")]
                [InlineData("1", "1", null)]
                public async void GivenPartsOfDateMissing_SetErrorOnTheModel(string day, string month, string year)
                {
                    var response = await _subject.TargetDatePost("0001", day, month, year);
                    var responseModel = ControllerTestHelpers.GetViewModelFromResult<TransferDatesViewModel>(response);

                    Assert.True(responseModel.FormErrors.HasErrors);
                    Assert.Equal("Please enter the target date for the transfer",
                        responseModel.FormErrors.Errors[0].ErrorMessage);
                }

                [Theory]
                [InlineData("0", "1", "2020")]
                [InlineData("32", "1", "2020")]
                [InlineData("1", "0", "2020")]
                [InlineData("1", "13", "2020")]
                [InlineData("1", "13", "0")]
                public async void GivenInvalidDate_SetErrorOnTheModel(string day, string month, string year)
                {
                    var response = await _subject.TargetDatePost("0001", day, month, year);
                    var responseModel = ControllerTestHelpers.GetViewModelFromResult<TransferDatesViewModel>(response);

                    Assert.True(responseModel.FormErrors.HasErrors);
                    Assert.Equal("Please enter a valid date", responseModel.FormErrors.Errors[0].ErrorMessage);
                }

                [Fact]
                public async void GivenGetByUrnReturnsError_DisplayErrorPage()
                {
                    var response = await _subject.TargetDatePost(_errorWithGetByUrn, "01", "01", "2020");
                    var viewResult = Assert.IsType<ViewResult>(response);

                    Assert.Equal("ErrorPage", viewResult.ViewName);
                    Assert.Equal("Error", viewResult.Model);
                }

                [Fact]
                public async void GivenUpdatenReturnsError_DisplayErrorPage()
                {
                    _projectsRepository.Setup(s => s.Update(It.IsAny<Project>())).ReturnsAsync(
                       new RepositoryResult<Project>
                       {
                           Error = new RepositoryResultBase.RepositoryError
                           {
                               ErrorMessage = "Update error"
                           }
                       });

                    var response = await _subject.TargetDatePost("0001", "01", "01", "2020");
                    var viewResult = Assert.IsType<ViewResult>(response);

                    Assert.Equal("ErrorPage", viewResult.ViewName);
                    Assert.Equal("Update error", viewResult.Model);
                }
            }
        }

        public class HtbDateTests : TransferDatesControllerTests
        {
            public class GetTests : HtbDateTests
            {
                [Fact]
                public async void GivenUrn_AssignsModelToTheView()
                {
                    var result = await _subject.HtbDate("0001");
                    var viewModel = ControllerTestHelpers.GetViewModelFromResult<TransferDatesViewModel>(result);

                    Assert.Equal(_foundProject.Urn, viewModel.Project.Urn);
                }

                [Fact]
                public async void GivenGetByUrnReturnsError_DisplayErrorPage()
                {
                    var response = await _subject.HtbDate(_errorWithGetByUrn);
                    var viewResult = Assert.IsType<ViewResult>(response);

                    Assert.Equal("ErrorPage", viewResult.ViewName);
                    Assert.Equal("Error", viewResult.Model);
                }
            }

            public class PostTests : HtbDateTests
            {
                [Theory]
                [InlineData("01", "01", "2020", "01/01/2020")]
                [InlineData("20", "02", "2021", "20/02/2021")]
                [InlineData("2", "2", "2021", "02/02/2021")]
                public async void GivenUrnAndFullDate_UpdatesTheProjectWithTheCorrectDate(string day, string month,
                    string year, string expectedDate)
                {
                    await _subject.HtbDatePost("0001", day, month, year);

                    _projectsRepository.Verify(r =>
                        r.Update(It.Is<Project>(project => project.Dates.Htb == expectedDate)));
                }

                [Fact]
                public async void GivenUrnAndFullDate_RedirectsToTheSummaryPage()
                {
                    var response = await _subject.HtbDatePost("0001", "01", "01", "2020");
                    ControllerTestHelpers.AssertResultRedirectsToAction(response, "Index");
                }

                [Theory]
                [InlineData(null, "1", "2020")]
                [InlineData("1", null, "2020")]
                [InlineData("1", "1", null)]
                public async void GivenPartsOfDateMissing_SetErrorOnTheModel(string day, string month, string year)
                {
                    var response = await _subject.HtbDatePost("0001", day, month, year);
                    var responseModel = ControllerTestHelpers.GetViewModelFromResult<TransferDatesViewModel>(response);

                    Assert.True(responseModel.FormErrors.HasErrors);
                    Assert.Equal("Please enter the HTB date", responseModel.FormErrors.Errors[0].ErrorMessage);
                }

                [Theory]
                [InlineData("0", "1", "2020")]
                [InlineData("32", "1", "2020")]
                [InlineData("1", "0", "2020")]
                [InlineData("1", "13", "2020")]
                [InlineData("1", "13", "0")]
                public async void GivenInvalidDate_SetErrorOnTheModel(string day, string month, string year)
                {
                    var response = await _subject.HtbDatePost("0001", day, month, year);
                    var responseModel = ControllerTestHelpers.GetViewModelFromResult<TransferDatesViewModel>(response);

                    Assert.True(responseModel.FormErrors.HasErrors);
                    Assert.Equal("Please enter a valid date", responseModel.FormErrors.Errors[0].ErrorMessage);
                }

                [Fact]
                public async void GivenGetByUrnReturnsError_DisplayErrorPage()
                {
                    var response = await _subject.HtbDatePost(_errorWithGetByUrn, "01", "01", "2020");
                    var viewResult = Assert.IsType<ViewResult>(response);

                    Assert.Equal("ErrorPage", viewResult.ViewName);
                    Assert.Equal("Error", viewResult.Model);
                }

                [Fact]
                public async void GivenUpdatenReturnsError_DisplayErrorPage()
                {
                    _projectsRepository.Setup(s => s.Update(It.IsAny<Project>())).ReturnsAsync(
                       new RepositoryResult<Project>
                       {
                           Error = new RepositoryResultBase.RepositoryError
                           {
                               ErrorMessage = "Update error"
                           }
                       });

                    var response = await _subject.HtbDatePost("0001", "01", "01", "2020");
                    var viewResult = Assert.IsType<ViewResult>(response);

                    Assert.Equal("ErrorPage", viewResult.ViewName);
                    Assert.Equal("Update error", viewResult.Model);
                }
            }
        }
    }
}