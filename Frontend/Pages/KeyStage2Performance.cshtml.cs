using System.Threading.Tasks;
using Data;
using Data.Models.KeyStagePerformance;
using Frontend.ExtensionMethods;
using Frontend.Models.Forms;
using Frontend.Services.Interfaces;
using Frontend.Services.Responses;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;

namespace Frontend.Pages
{
    public class KeyStage2Performance : PageModel
    {
        private readonly IGetInformationForProject _getInformationForProject;
        private readonly IProjects _projectRepository;
        public string ProjectUrn { get; private set; }
        public string OutgoingAcademyUrn { get; private set; }
        public string OutgoingAcademyName { get; private set; }
        public string LocalAuthorityName { get; private set; }
        public AdditionalInformationViewModel AdditionalInformation { get; private set; }
        public EducationPerformance EducationPerformance { get; private set; }

        public KeyStage2Performance(IGetInformationForProject getInformationForProject, IProjects projectRepository)
        {
            _getInformationForProject = getInformationForProject;
            _projectRepository = projectRepository;
        }
        
        public async Task<IActionResult> OnGetAsync(string id, bool addOrEditAdditionalInformation = false)
        {
            var projectInformation = await _getInformationForProject.Execute(id);

            if (!projectInformation.IsValid)
            {
                return this.View("ErrorPage", projectInformation.ResponseError.ErrorMessage);
            }

            BuildPageModel(projectInformation, addOrEditAdditionalInformation);

            return Page();
        }

        public async Task<IActionResult> OnPostAsync(string id, string additionalInformation)
        {
            var project = await _projectRepository.GetByUrn(id);

            if (!project.IsValid)
            {
                return this.View("ErrorPage", project.Error.ErrorMessage);
            }
            
            project.Result.KeyStage2PerformanceAdditionalInformation = additionalInformation;
            await _projectRepository.Update(project.Result);
            
            return new RedirectToPageResult(nameof(KeyStage2Performance), 
                "OnGetAsync", 
                new { id }, 
                "additional-information-hint");
        }

        private void BuildPageModel(GetInformationForProjectResponse projectInformation, bool addOrEditAdditionalInformation)
        {
            ProjectUrn = projectInformation.Project.Urn;
            OutgoingAcademyUrn = projectInformation.OutgoingAcademy.Urn;
            LocalAuthorityName = projectInformation.OutgoingAcademy.LocalAuthorityName;
            OutgoingAcademyName = projectInformation.OutgoingAcademy.Name;
            EducationPerformance = projectInformation.EducationPerformance;
            AdditionalInformation = new AdditionalInformationViewModel
            {
                AdditionalInformation = projectInformation.Project.KeyStage2PerformanceAdditionalInformation,
                HintText =
                    "This information will populate in your HTB template under the key stage performance tables section.",
                Urn = projectInformation.Project.Urn,
                AddOrEditAdditionalInformation = addOrEditAdditionalInformation
            };
        }
    }
}