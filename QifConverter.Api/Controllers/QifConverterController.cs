using Microsoft.AspNetCore.Http;
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

        /// <summary>
        /// Convert excel file content to QIF file
        /// </summary>
        /// <param name="repertoryPath">The path to the repertory containing the excel files</param>
        /// <param name="initialAmount">An initial amount on which starting transactions in QIF file</param>
        /// <param name="onlyTransactions">if true, the excel file is only containing transactions amount. If false, it is only containing amounts.</param>
        /// <returns>QIF file</returns>
        [HttpPost("file")]
        [ProducesResponseType(typeof(byte[]), StatusCodes.Status200OK)]
        [ProducesResponseType(typeof(BadRequestResult), StatusCodes.Status400BadRequest)]
        public async Task<IActionResult> ConvertToQif(string repertoryPath, float initialAmount, bool onlyTransactions)
        {
            _logger.LogDebug($"Converting files from {repertoryPath} into qif file...");

            var fileNames = _qifConverterService.GetExcelFileNamesFromDirectoryPath(repertoryPath);

            if (fileNames == null || fileNames.Length == 0)
            {
                _logger.LogWarning("No file(s) to scan.");
                return new BadRequestObjectResult("Error : No files to scan.");
            }

            var rows = _qifConverterService.ConvertFilesContentToRows(fileNames);

            if(rows == null || rows.Count == 0)
            {
                _logger.LogWarning("Empty file(s).");
                return new BadRequestObjectResult("Empty file(s).");
            }

            var qif = _qifConverterService.ConvertRowsToQif(rows, initialAmount, onlyTransactions);

            if (string.IsNullOrEmpty(qif))
            {
                _logger.LogWarning("Empty qif.");
                return new BadRequestObjectResult("Empty qif.");
            }

            var fileName = "result.qif";
            var qifFileUri = $"{repertoryPath}/{fileName}";

            await _qifConverterService.WriteInFile(qif, qifFileUri);

            if (!System.IO.File.Exists(qifFileUri))
            {
                _logger.LogError($"Error when creating qif file.");
                return new BadRequestObjectResult("Error when creating qif file.");
            }

            return File(await System.IO.File.ReadAllBytesAsync(qifFileUri), "application/octet-stream", fileName);
        }
    }
}
