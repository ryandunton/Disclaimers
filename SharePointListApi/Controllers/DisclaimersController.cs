using Microsoft.AspNetCore.Mvc;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Microsoft.Win32.SafeHandles;

// This is a sample Web API controller that returns a list of disclaimers.
// A fetch call is used from the office taskpane javascript code.  This example is hard coded, but you could expand this to pull the data from another back-end source like a database, storage, json file, etc.
// You would need to handle security/credentials if you are pulling from a secure source (not covered here)

namespace SharePointListApi.Controllers
{

    public class Disclaimer
    {
        public required string Text { get; set; }
        public required string Description { get; set; }
        public required string Version { get; set; }
    }


    [Route("api/[controller]")]
    [ApiController]
    public class DisclaimerController : ControllerBase
    {

        private readonly IHttpClientFactory _clientFactory;
        private readonly IHostEnvironment _environment;

        public DisclaimerController(IHttpClientFactory clientFactory, IHostEnvironment environment)
        {
            _clientFactory = clientFactory;
            _environment = environment;
        }

        [HttpGet]
        public async Task<IActionResult> Get()
        {
            try
            {
                var items = new List<Disclaimer>
                {
                    new Disclaimer
                    {
                        Description = "Standard Disclaimer",
                        Text = "This is a standard disclaimer.",
                        Version = "v1.0"
                    },
                    new Disclaimer
                    {
                        Description = "GCCH Disclaimer",
                        Text = "Content subject to GCCH restrictions.",
                        Version = "v1.0"
                    },
                    new Disclaimer
                    {
                        Description = "ITAR Disclaimer",
                        Text = "Content subject to ITAR restrictions.",
                        Version = "v1.0"
                    },
                    new Disclaimer
                    {
                        Description = "Random Note",
                        Text = "This is a random note",
                        Version = "1.0"
                    }
                };

                return Ok(items);
            }
            catch (Exception ex)
            {
                // Provide a proper JSON response for errors
                var errorMessage = _environment.IsDevelopment()
                    ? $"An error occurred: {ex.Message}"
                    : "An error occurred while processing your request.";
                
                return StatusCode(500, new { error = errorMessage });
            }
        }
    }
}
