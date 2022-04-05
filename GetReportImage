using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using System.Data;
using System.Drawing;
using System.Text;
using System.Data.SqlClient;

namespace getReportImage
{
    public static class GetReportImageByReportId
    {
        [FunctionName("GetReportImageByReportId")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req,
            ILogger log,ExecutionContext context)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");

            string name = req.Query["name"];

            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            dynamic data = JsonConvert.DeserializeObject(requestBody);
            name = name ?? data?.name;

            string responseMessage = string.IsNullOrEmpty(name)
                ? "This HTTP triggered function executed successfully. Pass a name in the query string or in the request body for a personalized response."
                : $"Hello, {name}. This HTTP triggered function executed successfully.";
            

            /** Connection to DB: **/
            SqlConnection cnn;
            var connectionString = Environment.GetEnvironmentVariable("ConnectionStrings:SQL_Connection");
            cnn = new SqlConnection(connectionString);
            try
            {
                cnn.Open();
                responseMessage="Connection Open!";
                SqlCommand command;
                
                string sql = "select img from myimages where report_id='"+name+"'";
                //23468_1
                command = new SqlCommand(sql, cnn);
                byte[] img = (byte[])command.ExecuteScalar();
                MemoryStream str = new MemoryStream();
                str.Write(img, 0, img.Length);
                Bitmap bit = new Bitmap(str);
                cnn.Close();
                return new FileContentResult(ImageToByteArray(bit), "image/jpeg");
            }
            catch (Exception ex)
            {
                responseMessage="Can not open connection!"+ex;
            }

            return new OkObjectResult(responseMessage);
        }

        private static byte[] ImageToByteArray(Image image)
        {
            ImageConverter converter = new ImageConverter();
            return (byte[])converter.ConvertTo(image, typeof(byte[]));
        }
    }
}
