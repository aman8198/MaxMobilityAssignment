using System.Data;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System.Data.SqlClient;
using System.Text.RegularExpressions;
using System.Globalization;

namespace YourNamespace.Controllers
{
    public class HomeController : Controller
    {
        private readonly string _connectionString;

        public HomeController(IConfiguration configuration)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection");
        }

        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public async Task<IActionResult> Upload(IFormFile file)
        {
            if (file != null && file.Length > 0 &&
                (Path.GetExtension(file.FileName).Equals(".xls") || Path.GetExtension(file.FileName).Equals(".xlsx")))
            {
                string uploadsFolder = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/Uploads");
                if (!Directory.Exists(uploadsFolder))
                {
                    Directory.CreateDirectory(uploadsFolder);
                }

                string filePath = Path.Combine(uploadsFolder, Path.GetFileName(file.FileName));
                using (var stream = new FileStream(filePath, FileMode.Create))
                {
                    await file.CopyToAsync(stream);
                }

                DataTable dataTable = ReadExcelFile(filePath);

                if (ValidateDataTable(dataTable))
                {
                    SaveToDatabase(dataTable);
                    ViewBag.Message = "Upload Successful";
                    ViewBag.DownloadLink = Url.Content("~/Uploads/" + Path.GetFileName(file.FileName));
                }
                else
                {
                    ViewBag.Message = "Validation Failed";
                }
            }
            else
            {
                ViewBag.Message = "Invalid File Format";
            }

            return View("Index");
        }

        private DataTable ReadExcelFile(string filePath)
        {
            // Set the license context before using EPPlus
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            DataTable dataTable = new DataTable();

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.First();
                bool hasHeader = true; // adjust it accordingly

                // Add columns to DataTable
                foreach (var firstRowCell in worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column])
                {
                    dataTable.Columns.Add(hasHeader ? firstRowCell.Text : $"Column {firstRowCell.Start.Column}");
                }

                // Add rows to DataTable
                var startRow = hasHeader ? 2 : 1;
                for (var rowNum = startRow; rowNum <= worksheet.Dimension.End.Row; rowNum++)
                {
                    var wsRow = worksheet.Cells[rowNum, 1, rowNum, worksheet.Dimension.End.Column];

                    // Check if all cells in the row are empty (blank)
                    if (wsRow.All(cell => string.IsNullOrWhiteSpace(cell.Text)))
                    {
                        continue; // Skip blank rows
                    }

                    DataRow row = dataTable.NewRow();

                    foreach (var cell in wsRow)
                    {
                        row[cell.Start.Column - 1] = cell.Text;
                    }

                    dataTable.Rows.Add(row);
                }
            }

            return dataTable;
        }

        private bool ValidateDataTable(DataTable dataTable)
        {
            bool isValid = true;

            foreach (DataRow row in dataTable.Rows)
            {
                bool isRowValid = true;

                // Check each field in the row for validity
                if (string.IsNullOrEmpty(row["Name"].ToString()))
                {
                    isRowValid = false;
                }

                if (string.IsNullOrEmpty(row["Email"].ToString()) || !IsValidEmail(row["Email"].ToString()))
                {
                    isRowValid = false;
                }

                if (string.IsNullOrEmpty(row["Phone No"].ToString()))
                {
                    isRowValid = false;
                }

                if (string.IsNullOrEmpty(row["Address"].ToString()))
                {
                    isRowValid = false;
                }

                // If any field in the row is invalid, set isValid to false
                if (!isRowValid)
                {
                    isValid = false;
                    // Optionally, you can log or accumulate validation errors here
                }
            }

            return isValid;
        }
        private bool IsValidEmail(string email)
        {
           
            try
            {
                // Normalize the domain part of the email address
                email = Regex.Replace(email, @"(@)(.+)$", DomainMapper, RegexOptions.None, TimeSpan.FromMilliseconds(300));

                // Return true if MailAddress accepts the email address
                return new System.Net.Mail.MailAddress(email).Address == email;
            }
            catch (Exception)
            {
                return false;
            }
        }
        private static string DomainMapper(Match match)
        {
            // Use IdnMapping class to convert Unicode domain names.
            var idn = new IdnMapping();

            // Pull out and process domain name (throws ArgumentException on invalid)
            string domainName = idn.GetAscii(match.Groups[2].Value);

            return match.Groups[1].Value + domainName;
        }

        private void SaveToDatabase(DataTable dataTable)
        {
            using (SqlConnection conn = new SqlConnection(_connectionString))
            {
                using (SqlCommand cmd = new SqlCommand("usp_SaveExcelData", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    SqlParameter tvpParam = cmd.Parameters.AddWithValue("@ExcelData", dataTable);
                    tvpParam.SqlDbType = SqlDbType.Structured;
                    conn.Open();
                    cmd.ExecuteNonQuery();
                }
            }
        }
    }
}
