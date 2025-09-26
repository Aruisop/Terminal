using System.ComponentModel.DataAnnotations;

namespace WebApplication1.Models
{
    public class GraphicWalker
    {
        public class Dashboard
        {
            [Required(ErrorMessage = "Dashboard name is required")]
            [StringLength(100, ErrorMessage = "Dashboard name cannot exceed 100 characters")]
            public string DashboardName { get; set; } = string.Empty;

            [Required(ErrorMessage = "JSON format is required")]
            public string JsonFormat { get; set; } = string.Empty;

            public bool IsMultiple { get; set; }

            [Required(ErrorMessage = "Dataset name is required")]
            public string DatasetName { get; set; } = string.Empty;
        }
        public class Dataset
        {
            [Required(ErrorMessage = "Dataset name is required")]
            [StringLength(100, ErrorMessage = "Dataset name cannot exceed 100 characters")]
            public string DatasetName { get; set; } = string.Empty;

            [StringLength(100, ErrorMessage = "Stored procedure name cannot exceed 100 characters")]
            public string SP { get; set; } = string.Empty;

            [StringLength(500, ErrorMessage = "Excel path cannot exceed 500 characters")]
            public string ExcelPath { get; set; } = string.Empty;

            public bool IsItFromExcel { get; set; }
        }
        public class DatasetUploadRequest
        {
            [Required(ErrorMessage = "File is required")]
            public IFormFile File { get; set; }

            [Required(ErrorMessage = "Dataset name is required")]
            [StringLength(100, ErrorMessage = "Dataset name cannot exceed 100 characters")]
            public string DatasetName { get; set; }
        }
    }
}
