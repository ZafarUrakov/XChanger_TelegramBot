using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Serialization;
using Telegram.Bot;
using Telegram.Bot.Types;
using Telegram.Bot.Types.Enums;
using XChanger_TelegramBot.Brokers;
using XChanger_TelegramBot.Models;

namespace XChanger_TelegramBot
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            var telegramBotClient =
                new TelegramBotClient("6486297329:AAFZeruXJHWP1FPh5hVtRAjIw_xCC0AByJk");

            telegramBotClient.StartReceiving(HandleMessage, HandleError);
            Console.ReadKey();
        }

        private static async Task HandleMessage(ITelegramBotClient client, Update update, CancellationToken token)
        {
            if (update.Message.Text is "/start")
            {
                await client.SendTextMessageAsync(
                    chatId: update.Message.Chat.Id,
                    text: "Excel file yuboring");
            }
            else if (update.Message.Type is MessageType.Document)
            {
                Telegram.Bot.Types.File file = await client.GetFileAsync(update.Message.Document.FileId);

                string folderPath = Path.Combine(Environment.CurrentDirectory, "Files");
                Directory.CreateDirectory(folderPath);
                string filePath = Path.Combine(folderPath, update.Message.Document.FileName);
                ExcelBroker excelBroker = new ExcelBroker();

                using (var fileStream = new FileStream(filePath, FileMode.Create))
                {
                    await client.DownloadFileAsync(file.FilePath, fileStream);
                }

                List<Person> persons =
                    await excelBroker.ReadAllExternalPersonsAsync(filePath);

                string xmlString = SerializeToXml(persons);
                byte[] bytes = Encoding.UTF8.GetBytes(xmlString);

                using (var stream = new MemoryStream(bytes))
                {
                    await client.SendDocumentAsync(
                        chatId: update.Message.Chat.Id,
                        document: InputFile.FromStream(stream, "Persons.xml"),
                        caption: "XML file");
                }
            }
        }

        private static string SerializeToXml(List<Person> persons)
        {
            XmlSerializer serializer = new XmlSerializer(typeof(List<Person>));
            using (StringWriter stringWriter = new StringWriter())
            {
                serializer.Serialize(stringWriter, persons);
                return stringWriter.ToString();
            }

        }

        private static async Task HandleError(ITelegramBotClient client, Exception exception, CancellationToken token)
        {
            Console.WriteLine("Error: " + exception.Message);
        }
    }
}

