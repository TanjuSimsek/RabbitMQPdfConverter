using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using RabbitMQ_Word_To_Pdf.Producer.Models;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using RabbitMQ.Client;

namespace RabbitMQ_Word_To_Pdf.Producer.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly IConfiguration _configuration;

        public HomeController(ILogger<HomeController> logger, IConfiguration configuration)
        {
            _logger = logger;
            _configuration = configuration;
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult WordToPdfPage()
        {
            return View();
        }
        [HttpPost]
        public IActionResult WordToPdfPage(WordToPdf wordToPdf)
        {
            var factory = new ConnectionFactory();
            factory.Uri = new Uri(_configuration["ConnectionStrings:RabbitMQCloudString"]);

            using (var connection=factory.CreateConnection())
            {
                using (var channel=connection.CreateModel())
                {
                    

                    channel.ExchangeDeclare("convert-exchange",ExchangeType.Direct,true,false,null);

                    channel.QueueDeclare(queue: "File", durable: true, exclusive: false,autoDelete:false,arguments:null);


                    channel.QueueBind("File", "convert-exchange","WordToPdf");


                    MessageWordToPdf messageWordToPdf = new MessageWordToPdf();

                    using (MemoryStream ms=new MemoryStream())
                    {
                        wordToPdf.WordFile.CopyTo(ms);
                        messageWordToPdf.WordByte = ms.ToArray();

                    }

                    messageWordToPdf.Email = wordToPdf.Email;
                    messageWordToPdf.FileName = Path.GetFileNameWithoutExtension(wordToPdf.WordFile.FileName);

                    string serializeMessage = JsonConvert.SerializeObject(messageWordToPdf);


                    byte[] ByteMesssage = Encoding.UTF8.GetBytes(serializeMessage);


                    var properties = channel.CreateBasicProperties();

                    properties.Persistent = true;

                    channel.BasicPublish("convert-exchange",routingKey: "WordToPdf",basicProperties:properties,body:ByteMesssage);

                    ViewBag.Result =
                        "Word dosyanız pdf dosyasına dönüştürüldükten sonra size E-Mail olarak gönderilecektir...";




                }
            }

            return View();
        }
        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
