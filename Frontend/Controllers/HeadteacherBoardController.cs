using System.Linq;
using System.Threading.Tasks;
using Frontend.Models;
using Frontend.Services.Interfaces;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace Frontend.Controllers
{
    [Authorize]
    [Route("project/{id}/headteacher-board")]
    public class HeadteacherBoardController : Controller
    {
        private readonly ICreateHtbDocument _createHtbDocument;
        private readonly IGetInformationForProject _getInformationForProject;

        public HeadteacherBoardController(ICreateHtbDocument createHtbDocument,
            IGetInformationForProject getInformationForProject)
        {
            _createHtbDocument = createHtbDocument;
            _getInformationForProject = getInformationForProject;
        }

        [Route("preview")]
        public async Task<IActionResult> Preview([FromRoute] string id)
        {
            var projectInformation = await _getInformationForProject.Execute(id);
            if (!projectInformation.IsValid)
            {
                return View("ErrorPage", projectInformation.ResponseError.ErrorMessage);
            }

            var model = new HeadTeacherBoardPreviewViewModel
            {
                Project = projectInformation.Project,
                OutgoingAcademy = projectInformation.OutgoingAcademy
            };

            return View(model);
        }

        [Route("download")]
        public async Task<IActionResult> Download([FromRoute] string id)
        {
            var projectInformation = await _getInformationForProject.Execute(id);
            if (!projectInformation.IsValid)
            {
                return View("ErrorPage", projectInformation.ResponseError.ErrorMessage);
            }

            var model = new ProjectViewModel()
            {
                Project = projectInformation.Project
            };

            return View(model);
        }

        public async Task<IActionResult> GenerateDocument([FromRoute] string id)
        {
            var document = await _createHtbDocument.Execute(id);
            if (!document.IsValid)
            {
                return View("ErrorPage", document.ResponseError.ErrorMessage);
            }

            return File(document.Document.ToArray(),
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "Test.docx");
        }
    }
}