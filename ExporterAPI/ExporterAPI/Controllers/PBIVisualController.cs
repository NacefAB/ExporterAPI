using ExporterAPI.Helpers;
using ExporterAPI.Model;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using System.Globalization;

namespace ExporterAPI.Controllers
{
    [ApiController]
    [Route("[controller]/[action]")]
    public class PBIVisualController : Controller
    {
        private readonly ILogger<PBIVisualController> _logger;

        public PBIVisualController(ILogger<PBIVisualController> logger)
        {
            _logger = logger;
        }
        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        [RequestSizeLimit(2147483648)]
        public ActionResult<string> ProcessExcelData([FromBody] object request)
        {
            ExportData result = JsonConvert.DeserializeObject<ExportData>(request.ToString());
            string filename = Utils.ExportDataToXlsx(result);
            string fileID = Path.GetFileNameWithoutExtension(filename);
            string exportFileName = Utils.CleanSpecialChars(CultureInfo.CurrentCulture.TextInfo.ToTitleCase(result.Title)) + ".xlsx";
            string url = $"{Request.Scheme}://{Request.Host}/PBIVisual/DownloadFile?id={fileID}&filename=" + exportFileName;
            ViewBag.Message = url;

            return url;
        }
        [HttpGet(Name = "DownloadFile")]
        public ActionResult DownloadFile(string id, string filename)
        {
            ViewBag.Message = $"/PBIVisual/Downloader?id={id}";
            ViewBag.DownloadFileName = filename;
            return View();
        }
        [HttpGet(Name = "Downloader")]
        public ActionResult Downloader(string id)
        {
            string downloadPath = Utils.GetExportFolder() + @"\" + id + ".xlsx";
            MemoryStream obj_stream = new MemoryStream();
            obj_stream = new MemoryStream(System.IO.File.ReadAllBytes(downloadPath));
            obj_stream.Seek(0, SeekOrigin.Begin);
            string contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            return File(obj_stream, contentType);
        }

    }



}
