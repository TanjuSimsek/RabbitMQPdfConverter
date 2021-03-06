using System;
using System.IO;
using System.Net.Mail;
using System.Text;
using System.Text.Json.Serialization;
using Newtonsoft.Json;
using RabbitMQ.Client;
using RabbitMQ.Client.Events;
using Spire.Doc;

namespace RabbitMQ_Word_To_Pdf.Consumer
{
    class Program
    {
        
        public static bool EmailSend(string email, MemoryStream memoryStream,string fileName)
        {
            try
            {

                memoryStream.Position = 0;

                System.Net.Mime.ContentType ct = new System.Net.Mime.ContentType(System.Net.Mime.MediaTypeNames.Application.Pdf);

                Attachment attach = new Attachment(memoryStream, ct);

                attach.ContentDisposition.FileName = $"{fileName}.pdf";

                MailMessage mailMessage = new MailMessage();

                SmtpClient smtpClient = new SmtpClient();

                mailMessage.From = new MailAddress("admin@teknohub.net");

                mailMessage.To.Add(email);

                mailMessage.Subject = "Created Pdf File !!";

                mailMessage.Body = "You can find your pdf file on attach";

                mailMessage.IsBodyHtml = true;

                mailMessage.Attachments.Add(attach);

                #region Smtp Sunucu Ayarı

                smtpClient.Host = "mail.teknohub.net";
                smtpClient.Port = 587;
                smtpClient.Credentials = new System.Net.NetworkCredential("admin@teknohub.net", "Fatih1234");

                smtpClient.Send(mailMessage);


                #endregion
                Console.WriteLine($"Result :{email} send successfuly");
                

                memoryStream.Close();
                memoryStream.Dispose();

                return true;

            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error occurred while sending mail :{ex.InnerException}");
                return false;
            }
           


        }

        static void Main(string[] args)
        {

            bool result = false;
            
            var factory = new ConnectionFactory();
            factory.Uri = new Uri("amqps://crbdixlg:g9QiD8b2xSqzb7Ht0lZMocnmE_2XdT97@jaguar.rmq.cloudamqp.com/crbdixlg");

            using (var connection = factory.CreateConnection())
            {
                using (var channel = connection.CreateModel())
                {


                    channel.ExchangeDeclare("convert-exchange", ExchangeType.Direct, true, false, null);

                    channel.QueueBind(queue:"File",exchange: "convert-exchange", "WordToPdf");

                    channel.BasicQos(0,1,false);

                    var consumer = new EventingBasicConsumer(channel);

                    channel.BasicConsume("File", false, consumer);

                    consumer.Received += (model, ea) =>
                    {

                        try
                        {
                            Console.WriteLine("Kuyruktan bir mesaj alındı ve işleniyor..");

                            Document document = new Document();

                            string message = Encoding.UTF8.GetString(ea.Body);
                            MessageWordToPdf messageWordToPdf =
                                JsonConvert.DeserializeObject<MessageWordToPdf>(message);

                            document.LoadFromStream(new MemoryStream(messageWordToPdf.WordByte),FileFormat.Docx2013);


                            using (MemoryStream ms=new MemoryStream())
                            {
                                document.SaveToStream(ms,FileFormat.PDF);

                                result = EmailSend(messageWordToPdf.Email, ms, messageWordToPdf.FileName);

                            }



                        }
                        catch (Exception e)
                        {
                            Console.WriteLine("Hata Meydana Geldi : "+e.Message);
                            
                        }

                        if (result)
                        {
                            Console.WriteLine("Kuyruktan Mesak Başarıyla İşlendi..");
                            channel.BasicAck(ea.DeliveryTag,false);
                        }

                    };



                    Console.WriteLine("Çıkmak için tıklayınız ");
                    Console.ReadLine();



                }
            }
        }
        }
}
