using Microsoft.AspNetCore.Mvc;
using WebApplication1.Models;
using WebApplication1.Data;

namespace WebApplication1.Controllers
{
    [ApiController]
    public class GraphicWalkerController : ControllerBase
    {
        private readonly GraphicWalkerConnection _connection;
        private readonly IWebHostEnvironment _environment;

        public GraphicWalkerController(GraphicWalkerConnection connection, IWebHostEnvironment environment)
        {
            _connection = connection;
            _environment = environment;
        }

        #region Dashboard Endpoints
        [HttpPost]
        [Route("Dashboard")]
        [ProducesResponseType(StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status400BadRequest)]
        [ProducesResponseType(StatusCodes.Status500InternalServerError)]
        public async Task<IActionResult> SaveDashboard([FromBody] GraphicWalker.Dashboard dashboard)
        {
            if (dashboard == null || string.IsNullOrWhiteSpace(dashboard.DashboardName))
            {
                return BadRequest(new { error = "Dashboard name is required." });
            }

            try
            {
                await _connection.SaveDashboard(dashboard);
                return Ok(new { message = "Dashboard saved successfully" });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { error = $"Failed to save dashboard: {ex.Message}" });
            }
        }

        [HttpGet]
        [Route("Dashboard")]
        [ProducesResponseType(StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status500InternalServerError)]
        public async Task<IActionResult> GetDashboards()
        {
            try
            {
                var dashboards = await _connection.GetDashboards();
                return Ok(dashboards);
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { error = $"Failed to retrieve dashboards: {ex.Message}" });
            }
        }
        #endregion

        #region Dataset Endpoints
        [HttpGet]
        [Route("Dataset")]
        [ProducesResponseType(StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status500InternalServerError)]
        public async Task<IActionResult> GetDatasets()
        {
            try
            {
                var datasets = await _connection.GetDatasets();
                return Ok(datasets);
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { error = $"Failed to retrieve datasets: {ex.Message}" });
            }
        }

        [HttpPost]
        [Route("Dataset")]
        [ProducesResponseType(StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status400BadRequest)]
        [ProducesResponseType(StatusCodes.Status500InternalServerError)]
        public async Task<IActionResult> SaveDataset([FromBody] GraphicWalker.Dataset dataset)
        {
            if (dataset == null || string.IsNullOrWhiteSpace(dataset.DatasetName))
            {
                return BadRequest(new { error = "Dataset name is required" });
            }

            try
            {
                await _connection.SaveDataset(dataset);
                return Ok(new { message = "Dataset saved successfully" });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { error = $"Failed to save dataset: {ex.Message}" });
            }
        }
        #endregion

        #region Excel Endpoints
        [HttpGet]
        [Route("api/excel/read")]
        [ProducesResponseType(StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status404NotFound)]
        [ProducesResponseType(StatusCodes.Status500InternalServerError)]
        public IActionResult ReadExcelData([FromQuery] string excelPath)
        {
            if (string.IsNullOrWhiteSpace(excelPath))
            {
                return BadRequest(new { error = "Excel file path is required" });
            }

            try
            {
                var data = _connection.ReadExcelData(excelPath);
                if (data == null || data.Count == 0)
                {
                    return NotFound(new { error = "No data found or Excel file does not exist." });
                }
                return Ok(data);
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { error = $"Failed to read Excel file: {ex.Message}" });
            }
        }
        #endregion

        #region StoredProcedure Endpoints
        [HttpGet]
        [Route("api/storedprocedure/execute")]
        [ProducesResponseType(StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status404NotFound)]
        [ProducesResponseType(StatusCodes.Status500InternalServerError)]
        public async Task<IActionResult> ExecuteStoredProcedure([FromQuery] string storedProcedureName)
        {
            if (string.IsNullOrWhiteSpace(storedProcedureName))
            {
                return BadRequest(new { error = "Stored procedure name is required" });
            }

            try
            {
                var parameters = new Dictionary<string, object>();
                var data = await _connection.ExecuteStoredProcedure(storedProcedureName, parameters);
                if (data == null || data.Count == 0)
                {
                    return NotFound(new { error = "No data found or stored procedure did not return any data." });
                }
                return Ok(data);
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { error = $"Failed to execute stored procedure: {ex.Message}" });
            }
        }
        #endregion

        [HttpPost]
        [Route("Dataset/Upload")]
        [ProducesResponseType(StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status400BadRequest)]
        [ProducesResponseType(StatusCodes.Status500InternalServerError)]
        public async Task<IActionResult> UploadDataset([FromForm] IFormFile file, [FromForm] string datasetName)
        {
            if (file == null || file.Length == 0)
                return BadRequest(new { error = "No file uploaded" });

            if (string.IsNullOrEmpty(datasetName))
                return BadRequest(new { error = "Dataset name is required" });

            try
            {
                // Create uploads directory if it doesn't exist
                var uploadsFolder = Path.Combine(_environment.ContentRootPath, "uploads");
                Directory.CreateDirectory(uploadsFolder);

                // Generate unique filename
                var uniqueFileName = $"{DateTime.Now:yyyyMMddHHmmss}_{file.FileName}";
                var filePath = Path.Combine(uploadsFolder, uniqueFileName);

                // Save file
                using (var stream = new FileStream(filePath, FileMode.Create))
                {
                    await file.CopyToAsync(stream);
                }

                // Create dataset entry
                var dataset = new GraphicWalker.Dataset
                {
                    DatasetName = datasetName,
                    ExcelPath = filePath,
                    IsItFromExcel = true,
                    SP = ""
                };

                await _connection.SaveDataset(dataset);

                return Ok(new
                {
                    message = "File uploaded successfully",
                    filePath = filePath
                });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { error = $"Failed to upload file: {ex.Message}" });
            }
        }
    }

}
