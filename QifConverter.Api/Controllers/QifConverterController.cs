using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using QifConverter.Api.Models;
using QifConverter.Api.Services;
using System;
using System.ComponentModel.DataAnnotations;
using System.IO;
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
        /// Convert excel file(s) content to the specified output bank file format.
        /// </summary>
        /// <param name="inputPath">The path to the repertory containing the input excel files.</param>
        /// <param name="outputPath">The path to the repertory containing the output bank files.</param>
        /// <param name="initialAmount">An initial amount on which starting transactions in QIF file.</param>
        /// <param name="onlyTransactions">If true, the excel file is only containing transactions amount. If false, it is only containing amounts.</param>
        /// <param name="inputFileType"> The input file type.</param>
        /// <param name="outputFileType"> The output bank file format.</param>
        /// <response code="200">OK.</response>
        /// <response code="400">BadRequest : invalid url provided or empty excel file(s).</response>
        /// <response code="404">No excel file(s) found.</response>
        /// <response code="500">Internal server error.</response>
        /// <response code="501">Not implemented.</response>
        /// <returns>QIF or OFX file.</returns>
        [HttpPost("file")]
        [ProducesResponseType(typeof(FileContentResult), StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status400BadRequest)]
        [ProducesResponseType(StatusCodes.Status404NotFound)]
        [ProducesResponseType(StatusCodes.Status500InternalServerError)]
        [ProducesResponseType(StatusCodes.Status501NotImplemented)]
        [Produces("application/json")]    
        public async Task<IActionResult> ConvertToQif(
            [Required] string inputPath,
            string outputPath,
            InputFileType inputFileType=InputFileType.ANY,
            OutputFileType outputFileType=OutputFileType.QIF,
            bool onlyTransactions=true,
            float initialAmount=0)
        {
            if(!inputFileType.Equals(InputFileType.ANY) || !outputFileType.Equals(OutputFileType.QIF))
            {
                _logger.LogError($"Not implemented input File type.");
                return new StatusCodeResult(501);
            }

            if(!Directory.Exists(inputPath) || (outputPath != null && !Directory.Exists(outputPath)))
            {
                _logger.LogWarning("Malformed url");
                return new BadRequestObjectResult("Malformed url");
            }

            _logger.LogDebug($"Converting files from {inputPath} into qif file...");

            var fileNames = _qifConverterService.GetExcelFileNamesFromDirectoryPath(inputPath);

            if (fileNames == null || fileNames.Length == 0)
            {
                _logger.LogWarning("No file(s) to scan.");
                return new NotFoundObjectResult("Error : No files to scan.");
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
            var qifFileUri = string.IsNullOrEmpty(outputPath) ? $"{inputPath}/{fileName}" : $"{outputPath}/{fileName}";

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
