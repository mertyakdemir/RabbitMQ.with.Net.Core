using Newtonsoft.Json;
using RabbitMQ.Client;
using RabbitMQ.Client.Events;
using Spire.Doc;
using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Text;

namespace RabbitMQ.WordToPDF.Consumer
{
    internal class Program
    {
        public static bool EmailSend(string email, MemoryStream memoryStream, string fileName)
        {
            try
            {
                memoryStream.Position = 0;

                System.Net.Mime.ContentType ct = new System.Net.Mime.ContentType(System.Net.Mime.MediaTypeNames.Application.Pdf);

                Attachment attach = new Attachment(memoryStream, ct);
                attach.ContentDisposition.FileName = $"{fileName}.pdf";

                MailMessage mailMessage = new MailMessage();

                SmtpClient smtpClient = new SmtpClient();

                mailMessage.From = new MailAddress("ozandemir777123@gmail.com");

                mailMessage.To.Add(email);

                mailMessage.Subject = "PDF file created";

                mailMessage.Body = "You'll find the attachment below.";

                mailMessage.IsBodyHtml = true;

                mailMessage.Attachments.Add(attach);

                smtpClient.Host = "smtp.gmail.com";
                smtpClient.Port = 587;
                smtpClient.EnableSsl = true;


                smtpClient.Credentials = new NetworkCredential("ozandemir777123@gmail.com", "Testpassword123");
                smtpClient.Send(mailMessage);
                Console.WriteLine($"Result: sent to { email}");

                memoryStream.Close();
                memoryStream.Dispose();
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while sending email:{ ex.InnerException}");
                return false;
            }
        }

        private static void Main(string[] args)
        {
            bool result = false;
            var factory = new ConnectionFactory();

            factory.Uri = new Uri("amqps://abbnslee:SFz_ia19C6--LYR8Z2t3LiYoa392dmfR@clam.rmq.cloudamqp.com/abbnslee");

            using (var connection = factory.CreateConnection())
            {
                using (var channel = connection.CreateModel())
                {
                    channel.ExchangeDeclare("convert-exchange", ExchangeType.Direct, true, false, null);

                    channel.QueueBind(queue: "File", exchange: "convert-exchange", "WordToPdf");

                    channel.BasicQos(0, 1, false);

                    var consumer = new EventingBasicConsumer(channel);

                    channel.BasicConsume("File", false, consumer);

                    consumer.Received += (model, ea) =>
                    {
                        try
                        {
                            Console.WriteLine("A message has been received from the queue and is being processed");

                            Document document = new Document();

                            var body = ea.Body.ToArray();
                            string message = Encoding.UTF8.GetString(body);

                            MessageWordToPdf messageWordToPdf = JsonConvert.DeserializeObject<MessageWordToPdf>(message);

                            document.LoadFromStream(new MemoryStream(messageWordToPdf.WordByte), FileFormat.Docx2013);

                            using (MemoryStream ms = new MemoryStream())
                            {
                                document.SaveToStream(ms, FileFormat.PDF);

                                result = EmailSend(messageWordToPdf.Email, ms, messageWordToPdf.FileName);
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Error occurred:" + ex.Message);
                        }

                        if (result)
                        {
                            Console.WriteLine("Message from the queue has been successfully processed");
                            channel.BasicAck(ea.DeliveryTag, false);
                        }
                    };

                    Console.WriteLine("Click to exit");
                    Console.ReadLine();
                }
            }
        }
    }
}
