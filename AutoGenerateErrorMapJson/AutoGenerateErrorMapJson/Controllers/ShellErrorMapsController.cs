using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using AutoGenerateErrorMapJson.ErrorMapping;

// For more information on enabling Web API for empty projects, visit https://go.microsoft.com/fwlink/?LinkID=397860

namespace AutoGenerateErrorMapJson.Controllers
{
    [Route("api/[controller]")]
    public class ShellErrorMapsController : Controller
    {
        // GET: api/shellerrormaps
        [HttpGet]
        public string Get()
        {
            return ErrorMapFileBuilder.BuildErrorMapsJson("ErrorCodes-message-Shell.xlsx", true);
        }
    }
}
