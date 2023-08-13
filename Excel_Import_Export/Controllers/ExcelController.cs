using ClosedXML.Excel;
using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Excel = Microsoft.Office.Interop.Excel;

namespace Excel_Import_Export.Controllers
{
    public class ExcelController : Controller
    {
        // GET: Excel
        public ActionResult Excel()
        {
            return View();
        }




        [HttpPost]
        public ActionResult UploadExcel(HttpPostedFileBase excelFile)
        {
            if (excelFile != null && excelFile.ContentLength > 0)
            {
                try
                {
                    // Read data from Excel file
                    string fileName = Path.GetFileNameWithoutExtension(excelFile.FileName); // Get the file name without extension
                    string fileExtension = Path.GetExtension(excelFile.FileName); // Get the file extension

                    // Append datetime stamp to the file name
                    string dateTimeStamp = DateTime.Now.ToString("yyyyMMddHHmmss");
                    string newFileName = $"{fileName}_{dateTimeStamp}{fileExtension}";

                    string path = Path.Combine(Server.MapPath("~/App_Data"), newFileName);
                    excelFile.SaveAs(path);

                    // Interact with Excel using COM interop
                    Excel.Application excelApp = new Excel.Application();
                    Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(path);
                    Excel.Worksheet excelWorksheet = excelWorkbook.Sheets[1]; // Assuming you're reading from the first sheet

                    int rowCount = excelWorksheet.UsedRange.Rows.Count;
                    int columnCount = excelWorksheet.UsedRange.Columns.Count;

                    // Transform and insert data into SQL table
                    string connectionString = ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;
                    using (SqlConnection connection = new SqlConnection(connectionString))
                    {
                        connection.Open();
                        SqlCommand cmd = connection.CreateCommand();
                        cmd.CommandType = CommandType.StoredProcedure;
                        SqlParameter[] parameters = new SqlParameter[]
                           {
                                new SqlParameter("@first_name", SqlDbType.NVarChar),
                                new SqlParameter("@last_name", SqlDbType.NVarChar),
                                new SqlParameter("@company_name", SqlDbType.NVarChar),
                                new SqlParameter("@address", SqlDbType.NVarChar),
                                new SqlParameter("@city", SqlDbType.NVarChar),
                                new SqlParameter("@county", SqlDbType.NVarChar),
                                new SqlParameter("@postal", SqlDbType.NVarChar),
                                new SqlParameter("@phone", SqlDbType.NVarChar),
                                new SqlParameter("@email", SqlDbType.NVarChar),
                                new SqlParameter("@web", SqlDbType.NVarChar)
                           };

                        for (int i = 2; i <= rowCount; i++) // Start from row 2 assuming row 1 is headers
                        {
                            cmd.CommandText = "InsertEmployee";

                            // Assign parameter values from Excel cells
                            for (int j = 0; j < columnCount; j++) // Note: j starts from 0
                            {
                                parameters[j].Value = (excelWorksheet.Cells[i, j + 1] as Excel.Range).Value;
                            }

                            // Add parameters to the command
                            cmd.Parameters.AddRange(parameters);

                            // Execute the stored procedure
                            cmd.ExecuteNonQuery();
                            cmd.Parameters.Clear();
                        }
                    }


                    excelWorkbook.Close(false);
                    excelApp.Quit();

                    ViewBag.Message = "File uploaded and data saved to SQL table successfully.";
                }
                catch (Exception ex)
                {
                    ViewBag.Message = "Error: " + ex.Message;
                }
            }
            else
            {
                ViewBag.Message = "Please select a file to upload.";
            }

            return View("Excel");
        }

        //public ActionResult UploadExcel(HttpPostedFileBase excelFile)
        //{
        //    if (excelFile != null && excelFile.ContentLength > 0)
        //    {
        //        try
        //        {
        //            // Read data from Excel file
        //            string fileName = Path.GetFileNameWithoutExtension(excelFile.FileName);
        //            string fileExtension = Path.GetExtension(excelFile.FileName);
        //            string dateTimeStamp = DateTime.Now.ToString("yyyyMMddHHmmss");
        //            string newFileName = $"{fileName}_{dateTimeStamp}{fileExtension}";

        //            string path = Path.Combine(Server.MapPath("~/App_Data"), newFileName);
        //            excelFile.SaveAs(path);

        //            using (var workbook = new XLWorkbook(path))
        //            {
        //                var worksheet = workbook.Worksheets.First();

        //                int rowCount = worksheet.RowsUsed().Count();
        //                int columnCount = worksheet.ColumnsUsed().Count();

        //                string connectionString = ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;
        //                using (SqlConnection connection = new SqlConnection(connectionString))
        //                {
        //                    connection.Open();
        //                    SqlCommand cmd = connection.CreateCommand();
        //                    cmd.CommandType = CommandType.StoredProcedure;

        //                    for (int i = 2; i <= rowCount; i++) // Start from row 2 assuming row 1 is headers
        //                    {
        //                        cmd.CommandText = "InsertEmployee";

        //                        for (int j = 1; j <= columnCount; j++) // Note: j starts from 1
        //                        {
        //                            cmd.Parameters.Clear(); // Clear previous parameters

        //                            // Construct the parameter name dynamically
        //                            string paramName = "@" + worksheet.Cell(1, j).Value.ToString().ToLower(); // Assuming headers are in the first row

        //                            // Add parameter and its value
        //                            cmd.Parameters.AddWithValue(paramName, worksheet.Cell(i, j).Value.ToString());
        //                        }

        //                        // Execute the stored procedure
        //                        cmd.ExecuteNonQuery();
        //                    }
        //                }
        //            }




        //            ViewBag.Message = "File uploaded and data saved to SQL table successfully.";
        //        }
        //        catch (Exception ex)
        //        {
        //            ViewBag.Message = "Error: " + ex.Message;
        //        }
        //    }
        //    else
        //    {
        //        ViewBag.Message = "Please select a file to upload.";
        //    }

        //    return View("Excel");
        //}

    }
}