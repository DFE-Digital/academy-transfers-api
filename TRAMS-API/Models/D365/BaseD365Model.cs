﻿using Newtonsoft.Json;

namespace API.Models.D365
{
    public class BaseD365Model
    {
        [JsonProperty("statecode")]
        public int StateCode { get; set; }

        [JsonProperty("statuscode")]
        public int StatusCode { get; set; }
    }
}
