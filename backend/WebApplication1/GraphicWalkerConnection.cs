using System.Data;
using System.Data.Odbc;
using ClosedXML.Excel;
using WebApplication1.Models;

namespace WebApplication1.Data
{
    public class GraphicWalkerConnection
    {
        private readonly string _connectionString;
        private readonly ILogger<GraphicWalkerConnection> _logger;

        public GraphicWalkerConnection(IConfiguration configuration, ILogger<GraphicWalkerConnection> logger)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection") ??
                throw new ArgumentNullException("DefaultConnection", "Connection string not found in configuration");
            _logger = logger;
        }

        #region Dashboard Operations
        public async Task<List<GraphicWalker.Dashboard>> GetDashboards()
        {
            var dashboards = new List<GraphicWalker.Dashboard>();
            using var connection = new OdbcConnection(_connectionString);

            try
            {
                await connection.OpenAsync();
                string query = "SELECT DashboardName, JsonFormat, IsMultiple, DatasetName FROM Dashboards";

                using var command = new OdbcCommand(query, connection);
                using var reader = await command.ExecuteReaderAsync();

                while (await reader.ReadAsync())
                {
                    dashboards.Add(new GraphicWalker.Dashboard
                    {
                        DashboardName = reader.GetString(0),
                        JsonFormat = reader.GetString(1),
                        IsMultiple = reader.GetBoolean(2),
                        DatasetName = reader.GetString(3)
                    });
                }
                _logger.LogInformation($"Retrieved {dashboards.Count} dashboards");
                return dashboards;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error retrieving dashboards");
                throw;
            }
        }

        public async Task SaveDashboard(GraphicWalker.Dashboard dashboard)
        {
            using var connection = new OdbcConnection(_connectionString);
            try
            {
                await connection.OpenAsync();

                // Check for existing dashboard
                string checkQuery = "SELECT COUNT(*) FROM Dashboards WHERE DashboardName = ?";
                using (var checkCommand = new OdbcCommand(checkQuery, connection))
                {
                    checkCommand.Parameters.Add("@DashboardName", OdbcType.NVarChar, 255).Value = dashboard.DashboardName;
                    int count = Convert.ToInt32(await checkCommand.ExecuteScalarAsync());
                    if (count > 0)
                    {
                        throw new Exception("Dashboard name already exists");
                    }
                }

                // Insert new dashboard
                string insertQuery = @"INSERT INTO Dashboards 
                             (DashboardName, JsonFormat, IsMultiple, DatasetName) 
                             VALUES (?, ?, ?, ?)";

                using var command = new OdbcCommand(insertQuery, connection);

                // Set parameters with explicit types and sizes
                var p1 = command.Parameters.Add("@DashboardName", OdbcType.NVarChar, 255);
                p1.Value = dashboard.DashboardName;

                var p2 = command.Parameters.Add("@JsonFormat", OdbcType.NVarChar);
                p2.Value = dashboard.JsonFormat ?? (object)DBNull.Value;

                var p3 = command.Parameters.Add("@IsMultiple", OdbcType.Bit);
                p3.Value = dashboard.IsMultiple;

                var p4 = command.Parameters.Add("@DatasetName", OdbcType.NVarChar, 50);
                p4.Value = dashboard.DatasetName ?? (object)DBNull.Value;

                await command.ExecuteNonQueryAsync();
            }
            catch (Exception ex)
            {
                throw new Exception($"Error saving dashboard: {ex.Message}", ex);
            }
        }

        #endregion

        #region Dataset Operations
        public async Task<List<GraphicWalker.Dataset>> GetDatasets()
        {
            var datasets = new List<GraphicWalker.Dataset>();
            using var connection = new OdbcConnection(_connectionString);

            try
            {
                await connection.OpenAsync();
                string query = "SELECT DatasetName, SP, Excelpath, IsItFromExcel FROM Dataset WHERE DatasetName IS NOT NULL";

                using var command = new OdbcCommand(query, connection);
                using var reader = await command.ExecuteReaderAsync();

                while (await reader.ReadAsync())
                {
                    string datasetName = !reader.IsDBNull(0) ? reader.GetString(0) : string.Empty;
                    if (!string.IsNullOrEmpty(datasetName))
                    {
                        datasets.Add(new GraphicWalker.Dataset
                        {
                            DatasetName = datasetName,
                            SP = !reader.IsDBNull(1) ? reader.GetString(1) : string.Empty,
                            ExcelPath = !reader.IsDBNull(2) ? reader.GetString(2) : string.Empty,
                            IsItFromExcel = !reader.IsDBNull(3) && reader.GetBoolean(3)
                        });
                    }
                }
                _logger.LogInformation($"Retrieved {datasets.Count} datasets");
                return datasets;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error retrieving datasets");
                throw;
            }
        }

        public async Task SaveDataset(GraphicWalker.Dataset dataset)
        {
            using var connection = new OdbcConnection(_connectionString);
            try
            {
                await connection.OpenAsync();

                string checkQuery = "SELECT COUNT(*) FROM Dataset WHERE DatasetName = ?";
                using (var checkCommand = new OdbcCommand(checkQuery, connection))
                {
                    checkCommand.Parameters.Add("@DatasetName", OdbcType.VarChar).Value = dataset.DatasetName;
                    int existingCount = Convert.ToInt32(await checkCommand.ExecuteScalarAsync());

                    if (existingCount > 0)
                    {
                        throw new Exception("Dataset name already exists");
                    }
                }

                string insertQuery = @"INSERT INTO Dataset (DatasetName, SP, ExcelPath, IsItFromExcel) 
                                     VALUES (?, ?, ?, ?)";

                using var command = new OdbcCommand(insertQuery, connection);
                command.Parameters.Add("@DatasetName", OdbcType.VarChar).Value = dataset.DatasetName;
                command.Parameters.Add("@SP", OdbcType.VarChar).Value = dataset.IsItFromExcel ? DBNull.Value : dataset.SP;
                command.Parameters.Add("@ExcelPath", OdbcType.VarChar).Value = dataset.ExcelPath;
                command.Parameters.Add("@IsItFromExcel", OdbcType.Bit).Value = dataset.IsItFromExcel;

                await command.ExecuteNonQueryAsync();
                _logger.LogInformation($"Dataset {dataset.DatasetName} saved successfully");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error saving dataset {dataset.DatasetName}");
                throw;
            }
        }
        #endregion

        #region Excel Operations
        public List<Dictionary<string, object>> ReadExcelData(string excelPath)
        {
            var dataList = new List<Dictionary<string, object>>();

            try
            {
                if (!File.Exists(excelPath))
                {
                    _logger.LogWarning($"Excel file does not exist: {excelPath}");
                    return dataList;
                }

                using var workbook = new XLWorkbook(excelPath);
                var worksheet = workbook.Worksheet(1);
                var rowCount = worksheet.RowsUsed().Count();
                var columnCount = worksheet.ColumnsUsed().Count();

                _logger.LogInformation($"Processing Excel file: Rows={rowCount}, Columns={columnCount}");

                // Get headers from first row
                var headers = new List<string>();
                for (int col = 1; col <= columnCount; col++)
                {
                    headers.Add(worksheet.Cell(1, col).GetValue<string>());
                }

                // Read data rows
                for (int row = 2; row <= rowCount; row++)
                {
                    var rowData = new Dictionary<string, object>();

                    for (int col = 1; col <= columnCount; col++)
                    {
                        var cell = worksheet.Cell(row, col);
                        var header = headers[col - 1];

                        if (cell.IsEmpty())
                        {
                            rowData[header] = null;
                            continue;
                        }

                        var value = cell.Value;
                        switch (cell.DataType)
                        {
                            case XLDataType.Number:
                                rowData[header] = cell.GetValue<double>();
                                break;
                            case XLDataType.DateTime:
                                rowData[header] = cell.GetValue<DateTime>();
                                break;
                            case XLDataType.Boolean:
                                rowData[header] = cell.GetValue<bool>();
                                break;
                            default:
                                rowData[header] = cell.GetValue<string>();
                                break;
                        }
                    }

                    dataList.Add(rowData);
                }

                _logger.LogInformation($"Successfully read {dataList.Count} rows from Excel file");
                return dataList;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error reading Excel file: {excelPath}");
                throw;
            }
        }
        #endregion

        #region StoredProcedure Operations
        public async Task<List<Dictionary<string, object>>> ExecuteStoredProcedure(string storedProcedureName, Dictionary<string, object> parameters)
        {
            var dataList = new List<Dictionary<string, object>>();
            using var connection = new OdbcConnection(_connectionString);

            try
            {
                await connection.OpenAsync();
                using var command = new OdbcCommand(storedProcedureName, connection)
                {
                    CommandType = CommandType.StoredProcedure
                };

                foreach (var param in parameters)
                {
                    command.Parameters.AddWithValue(param.Key, param.Value ?? DBNull.Value);
                }

                using var reader = await command.ExecuteReaderAsync();
                while (await reader.ReadAsync())
                {
                    var rowData = new Dictionary<string, object>();
                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        var columnName = reader.GetName(i);
                        rowData[columnName] = reader.IsDBNull(i) ? null : reader.GetValue(i);
                    }
                    dataList.Add(rowData);
                }

                _logger.LogInformation($"Stored procedure {storedProcedureName} executed successfully, returned {dataList.Count} rows");
                return dataList;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error executing stored procedure {storedProcedureName}");
                throw;
            }
        }
        #endregion
    }
}
