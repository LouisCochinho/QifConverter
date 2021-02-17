using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using QifConverter.Api.Services;
using System.Threading.Tasks;

namespace QifConverter.Api.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class QifConverterController : ControllerBase
    {
        private readonly ILogger<QifConverterController> _logger;
        private readonly IQifConverterService _qifConverterService;

        public QifConverterController(ILogger<QifConverterController> logger, IQifConverterService qifConverterService)
        {
            _logger = logger;
            _qifConverterService = qifConverterService;
        }

        [HttpPost]
        [Route("file")]
        public async Task<IActionResult> ConvertToQif(string folderPath)
        {
            _logger.LogDebug($"Converting files from {folderPath} into qif file...");
            return Created("/path/", new object { });
        }
    }
}
