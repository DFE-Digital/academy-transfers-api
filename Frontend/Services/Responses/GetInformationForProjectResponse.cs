using Data.Models;
using Data.Models.KeyStagePerformance;

namespace Frontend.Services.Responses
{
    public class GetInformationForProjectResponse
    {
        public GetInformationForProjectResponse()
        {
            ResponseError = new ServiceResponseError();
        }

        public Project Project { get; set; }
        public Academy OutgoingAcademy { get; set; }
        public EducationPerformance EducationPerformance { get; set; }
        public ServiceResponseError ResponseError { get; set; }
        public bool IsValid => string.IsNullOrEmpty(ResponseError.ErrorMessage);
    }
}