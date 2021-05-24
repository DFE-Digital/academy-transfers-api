using System.Collections.Generic;
using System.Threading.Tasks;
using Data;
using Frontend.Models;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Routing;

namespace Frontend.Controllers.Projects
{
    [Authorize]
    [Route("project/{urn}/rationale")]
    public class RationaleController : Controller
    {
        private readonly IProjects _projectsRepository;

        public RationaleController(IProjects projectsRepository)
        {
            _projectsRepository = projectsRepository;
        }

        public async Task<IActionResult> Index(string urn)
        {
            var model = await GetModel(urn);
            return View(model);
        }

        [HttpGet("rationale")]
        public async Task<IActionResult> Project(string urn)
        {
            var model = await GetModel(urn);
            return View(model);
        }

        [ActionName("Project")]
        [HttpPost("rationale")]
        public async Task<IActionResult> ProjectPost(string urn, string rationale)
        {
            var model = await GetModel(urn);
            model.Project.Rationale.Project = rationale;
            
            if (string.IsNullOrEmpty(rationale))
            {
                model.FormErrors.AddError("rationale", "rationale", "Please enter a rationale");
                return View(model);
            }

            await _projectsRepository.Update(model.Project);
            return RedirectToAction("Index", new {urn});
        }

        [HttpGet("trust-or-sponsor")]
        public async Task<IActionResult> TrustOrSponsor(string urn)
        {
            var model = await GetModel(urn);
            return View(model);
        }

        [ActionName("TrustOrSponsor")]
        [HttpPost("trust-or-sponsor")]
        public async Task<IActionResult> TrustOrSponsorPost(string urn, string rationale)
        {
            var model = await GetModel(urn);
            model.Project.Rationale.Trust = rationale;
            
            if (string.IsNullOrEmpty(rationale))
            {
                model.FormErrors.AddError("rationale", "rationale", "Please enter a rationale");
                return View(model);
            }

            await _projectsRepository.Update(model.Project);
            return RedirectToAction("Index", new {urn});
        }

        private async Task<RationaleViewModel> GetModel(string urn)
        {
            var project = await _projectsRepository.GetByUrn(urn);

            var model = new RationaleViewModel
            {
                Project = project.Result
            };

            return model;
        }
    }
}