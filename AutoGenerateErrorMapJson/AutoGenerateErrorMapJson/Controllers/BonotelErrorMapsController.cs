using Microsoft.AspNetCore.Mvc;
using AutoGenerateErrorMapJson.ErrorMapping;

// For more information on enabling Web API for empty projects, visit https://go.microsoft.com/fwlink/?LinkID=397860

namespace AutoGenerateErrorMapJson.Controllers
{
    [Route("api/[controller]")]
    public class BonotelErrorMapsController : Controller
    {
        // GET: api/BonotelErrorMaps
        [HttpGet]
        public string Get()
        {
            return ErrorMapFileBuilder.BuildErrorMapsJson("ErrorCodes-message-Bonotel.xlsx", false);
        }
    }
}
