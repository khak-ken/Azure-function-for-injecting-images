using System;
using System.IO;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Extensions.Logging;
using System.Data;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Diagnostics;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;

namespace scrape
{
    public static class BlobTriggerScrapeImages
    {
        [FunctionName("BlobTriggerScrapeImages")]
        public static void Run([BlobTrigger("scapedimages/meta_data_intellispec/{name}", Connection = "AZURE_STORAGE_CONNECTION_STRING")] Stream myBlob, string name, ILogger log)
        {
            DataTable dt = new DataTable();
            List<string> Headers = new List<string>();
            log.LogInformation($"C# Blob trigger function Processed blob\n Name:{name} \n Size: {myBlob.Length} Bytes");
            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(myBlob, false))
            {
                WorkbookPart workbookPart = doc.WorkbookPart;
                SharedStringTablePart sstpart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
                SharedStringTable sst = sstpart.SharedStringTable;

                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                //Worksheet sheet = worksheetPart.Worksheet;
                var sheetName = "All Platforms HTML Link";
                string relationshipId = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name.Equals(sheetName))?.Id;

                Worksheet sheet = ((WorksheetPart)workbookPart.GetPartById(relationshipId)).Worksheet;

                SheetData sheetData = sheet.GetFirstChild<SheetData>();
                IEnumerable<Row> rows = sheetData.Descendants<Row>();
                bool firstRowIsHeader = true;
                int counter = 0;
                foreach (Row row in rows)
                {
                    counter = counter + 1;
                    //Read the first row as header
                    if (counter == 1)
                    {
                        var j = 1;
                        foreach (Cell cell in row.Descendants<Cell>())
                        {
                            if (cell.CellValue.InnerXml != null)
                            {
                                var colunmName = firstRowIsHeader ? GetCellValue(doc, cell) : "Field" + j++;

                                if (dt.Columns.Contains(colunmName))
                                {
                                    colunmName = colunmName + "_1";
                                }
                                Headers.Add(colunmName);
                                dt.Columns.Add(colunmName);
                            }
                        }
                    }
                    else
                    {
                        dt.Rows.Add();
                        int i = 0;
                        foreach (Cell cell in row.Descendants<Cell>())
                        {
                            try
                            {
                                dt.Rows[dt.Rows.Count - 1][i] = GetCellValue(doc, cell);
                                i++;
                            }
                            catch (NullReferenceException e)
                            {
                            }
                        }
                    }
                }

                string[] selectedColumns = new[] { "Link" };

                DataTable dt_link = new DataView(dt).ToTable(false, selectedColumns);
                dt_link.Columns.Add("report_id", typeof(System.Int32));

                foreach (DataRow dataRow in dt_link.Rows)
                {
                    if (!dataRow["Link"].ToString().Equals(""))
                    {
                        dataRow["report_id"] = dataRow["Link"].ToString().Substring(dataRow["Link"].ToString().LastIndexOf("/") + 1);
                        foreach (var item in dataRow.ItemArray)
                        {
                            Console.WriteLine(item);
                        }
                    }

                }

                //DB connection:
                var connectionString = Environment.GetEnvironmentVariable("ConnectionStrings:SQL_Connection");
                string result_class_from_report = "";
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    var class_from_report_from_db = "SELECT Distinct SUBSTRING(report_id,1,CHARINDEX('_',report_id)-1) as report_id_short from images;";
                    SqlCommand command1 = new SqlCommand(class_from_report_from_db, connection);
                    SqlDataReader reader = command1.ExecuteReader();

                    while (reader.Read())
                    {
                        result_class_from_report = result_class_from_report + reader[0] + ",";
                    }
                    result_class_from_report = result_class_from_report.Remove(result_class_from_report.Length - 1);
                }

                DataRow[] result = dt_link.Select("report_id not  in (" + result_class_from_report + ")");
                Console.WriteLine("result:" + result.Count());
                IWebDriver driver = new ChromeDriver("C:/Users/nhlr/Anaconda-22-Jun-2020/Scripts");

                
                foreach (DataRow row in result)
                {
                    Console.WriteLine("{0}, {1}", row[0], row[1]);
                    driver.Navigate().GoToUrl("http://localhost:7071/api/HttpTriggerSelenium?name=" + row[0]);
                }

            }

        }
        private static string GetCellValue(SpreadsheetDocument document, Cell cell)
        {
            var stringTablePart = document.WorkbookPart.SharedStringTablePart;
            var value = cell.CellValue.InnerXml;

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                return stringTablePart.SharedStringTable.ChildElements[int.Parse(value)].InnerText;

            return value;
        }

    }
}

