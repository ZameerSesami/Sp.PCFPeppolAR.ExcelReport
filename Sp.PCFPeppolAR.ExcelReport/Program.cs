using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//using Excel = Microsoft.Office.Interop.Excel;
using System.Net.Mail;
using ClosedXML.Excel;

namespace Sp.PCFPeppolAR.ExcelReport
{
    public class Program
    {
        // Modify your DB connection string and base output folder
        private static string connectionString = System.Configuration.ConfigurationManager.ConnectionStrings["PCFCon"].ConnectionString;
        private static string ReportPath = System.Configuration.ConfigurationManager.AppSettings["ExcelFilePath"];
        private static string LogsPath = System.Configuration.ConfigurationManager.AppSettings["LogsFilePath"];
        private static string currDateFolderName = DateTime.Now.AddDays(-1).ToString("yyyyMMdd");
        private static string currDate = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd");

        static void Main(string[] args)
        {
            try
            {
                //string connectionString = System.Configuration.ConfigurationManager.ConnectionStrings["PCFCon"].ConnectionString;
                //string ReportPath = System.Configuration.ConfigurationManager.AppSettings["ExcelFilePath"];
                //string LogsPath = System.Configuration.ConfigurationManager.AppSettings["LogsFilePath"];

                if (!System.IO.Directory.Exists(ReportPath))
                {
                    System.IO.Directory.CreateDirectory(ReportPath);
                }
                if (!System.IO.Directory.Exists(LogsPath))
                {
                    System.IO.Directory.CreateDirectory(LogsPath);
                }
                System.Data.DataTable data = GetData();

                if (data.Rows.Count == 0)
                {
                    //Console.WriteLine("No data found to generate report.");
                    Log("No data found to generate report.");
                    //Console.ReadLine();
                }

                string todayFolder = Path.Combine(ReportPath, currDateFolderName);
                Directory.CreateDirectory(todayFolder);

                string filePath = Path.Combine(todayFolder, "PCFPeppolARReport_" + currDateFolderName + ".xlsx");
                SaveAsXlsx(data, filePath);
                //SaveAsExcelInterop(data, filePath);
                SendEmailWithAttachment(data, filePath);
                Log($"Report generated successfully: {filePath}");
                //Console.WriteLine("Report generated successfully.");
                //Console.ReadLine();
            }
            catch (Exception ex)
            {
                Log("Error: " + ex.ToString());
                //Console.WriteLine("Error occurred. Check the log." + ex.ToString());
                //Console.ReadLine();
            }
        }

        static DataTable GetData()
        {
            string connectionString = System.Configuration.ConfigurationManager.ConnectionStrings["PCFCon"].ConnectionString;
            string ReportPath = System.Configuration.ConfigurationManager.AppSettings["ExcelFilePath"];
            string LogsPath = System.Configuration.ConfigurationManager.AppSettings["LogsFilePath"];

            var dt = new DataTable();

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                string query = System.Configuration.ConfigurationManager.AppSettings["Query1"];

                using (SqlCommand cmd = new SqlCommand(query, conn))
                {
                    conn.Open();
                    using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                    {
                        adapter.Fill(dt);
                    }
                }
            }

            return dt;
        }

        static void SaveAsXlsx(DataTable data, string filePath)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Report");

                // Insert headers
                for (int i = 0; i < data.Columns.Count; i++)
                {
                    worksheet.Cell(1, i + 1).Value = data.Columns[i].ColumnName;
                    worksheet.Cell(1, i + 1).Style.Font.Bold = true;
                }

                // Insert data
                for (int row = 0; row < data.Rows.Count; row++)
                {
                    for (int col = 0; col < data.Columns.Count; col++)
                    {
                        worksheet.Cell(row + 2, col + 1).Value = data.Rows[row][col];
                    }
                }
                //this will autoadjust the columns 
                worksheet.Columns().AdjustToContents();
                workbook.SaveAs(filePath);
            }
        }

        static void SaveAsExcel(DataTable table, string filePath)
        {
            string connectionString = System.Configuration.ConfigurationManager.ConnectionStrings["PCFCon"].ConnectionString;
            string ReportPath = System.Configuration.ConfigurationManager.AppSettings["ExcelFilePath"];
            string LogsPath = System.Configuration.ConfigurationManager.AppSettings["LogsFilePath"];

            StringBuilder sb = new StringBuilder();

            sb.AppendLine("<table border='1' style='border-collapse:collapse;'>");

            // Header row
            sb.AppendLine("<tr>");
            foreach (DataColumn column in table.Columns)
            {
                sb.AppendFormat("<th>{0}</th>", column.ColumnName);
            }
            sb.AppendLine("</tr>");

            // Data rows
            foreach (DataRow row in table.Rows)
            {
                sb.AppendLine("<tr>");
                foreach (var item in row.ItemArray)
                {
                    sb.AppendFormat("<td>{0}</td>", item.ToString());
                }
                sb.AppendLine("</tr>");
            }

            sb.AppendLine("</table>");

            File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);
        }

        //static void SaveAsExcelInterop(DataTable table, string filePath)
        //{
        //    var excelApp = new Excel.Application();
        //    excelApp.Visible = false;
        //    Excel.Workbook workbook = excelApp.Workbooks.Add();
        //    Excel.Worksheet worksheet = workbook.Sheets[1];

        //    // Header
        //    for (int i = 0; i < table.Columns.Count; i++)
        //    {
        //        worksheet.Cells[1, i + 1] = table.Columns[i].ColumnName;
        //        ((Excel.Range)worksheet.Cells[1, i + 1]).Font.Bold = true;
        //    }

        //    // Data
        //    for (int r = 0; r < table.Rows.Count; r++)
        //    {
        //        for (int c = 0; c < table.Columns.Count; c++)
        //        {
        //            worksheet.Cells[r + 2, c + 1] = table.Rows[r][c]?.ToString();
        //        }
        //    }
        //    // Auto fit columns
        //    worksheet.Columns.AutoFit();

        //    // Save as .xlsx
        //    workbook.SaveAs(filePath, Excel.XlFileFormat.xlOpenXMLWorkbook);
        //    workbook.Close(false);
        //    excelApp.Quit();

        //    // Release COM objects
        //    System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
        //    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
        //    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
        //}

        static void SendEmailWithAttachment(DataTable dt,string attachmentPath)
        {
            try
            {
                var fromEmail = System.Configuration.ConfigurationManager.AppSettings["MailFrom"];
                var smtpHost = System.Configuration.ConfigurationManager.AppSettings["ServerHost"];
                var smtpPort = 25; // or 465/25 depending on your server

                using (MailMessage mail = new MailMessage())
                {
                    String Toaddressees = System.Configuration.ConfigurationManager.AppSettings["MailTo"];
                    String[] toaddr = Toaddressees.Split(',');
                    foreach (string ToEmail in toaddr)
                    {
                        mail.To.Add(new MailAddress(ToEmail));
                    }

                    String CCaddressees = System.Configuration.ConfigurationManager.AppSettings["MailCC"];
                    String[] ccaddr = CCaddressees.Split(',');
                    foreach (string ccEmail in ccaddr)
                    {
                        mail.CC.Add(new MailAddress(ccEmail));
                    }

                    mail.From = new MailAddress(fromEmail);
                    mail.Subject = "Report of PCF Peppol invoices received on " + currDate;
                    mail.Body = "Please find attached Report of PCF Peppol AR Invoices received on " + currDate + "<br/><br/> Total invoice received " + dt.Rows.Count.ToString();
                    mail.IsBodyHtml = true;
                    String addressees = System.Configuration.ConfigurationManager.AppSettings["MailTo"];
                    String[] addr = addressees.Split(',');
                    foreach (string MultiEmailId in addr)
                    {
                        mail.To.Add(new MailAddress(MultiEmailId));
                    }

                    if (!string.IsNullOrEmpty(attachmentPath) && File.Exists(attachmentPath))
                    {
                        mail.Attachments.Add(new Attachment(attachmentPath));
                    }

                    using (SmtpClient smtp = new SmtpClient(smtpHost, smtpPort))
                    {
                        //smtp.Credentials = new NetworkCredential(fromEmail, fromPassword);
                        //smtp.EnableSsl = true;
                        smtp.Send(mail);
                    }
                }

                Log("Email sent successfully.");
            }
            catch (Exception ex)
            {
                Log("Error sending email: " + ex.ToString());
            }
        }


        static void Log(string message)
        {
            string connectionString = System.Configuration.ConfigurationManager.ConnectionStrings["PCFCon"].ConnectionString;
            string ReportPath = System.Configuration.ConfigurationManager.AppSettings["ExcelFilePath"];
            string LogsPath = System.Configuration.ConfigurationManager.AppSettings["LogsFilePath"];

            try
            {
                //string logDir = Path.Combine(LogsPath, "logs");
                //Directory.CreateDirectory(logDir);

                string logFile = Path.Combine(LogsPath, $"log_{DateTime.Now:yyyyMMdd}.txt");
                string logEntry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}";

                File.AppendAllText(logFile, logEntry + Environment.NewLine);
            }
            catch
            {
                // Avoid crashing on logging errors
            }
        }

    }
}
