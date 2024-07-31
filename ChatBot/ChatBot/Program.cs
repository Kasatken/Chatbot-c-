using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using OfficeOpenXml;
using Telegram.Bot;
using Telegram.Bot.Exceptions;
using Telegram.Bot.Polling;
using Telegram.Bot.Types;
using Telegram.Bot.Types.Enums;
using Telegram.Bot.Types.ReplyMarkups;

class Program
{
    private static ITelegramBotClient botClient;
    private const string ExcelFilePath = "C:\\Users\\kaptg\\Downloads\\Аурора.xlsx";

    static async Task Main()
    {

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        botClient = new TelegramBotClient("6999685472:AAHpgoghBIhT5xwchakdI14CRpwvAbTL4nk");

        using var cts = new CancellationTokenSource();


        var receiverOptions = new ReceiverOptions
        {
            AllowedUpdates = Array.Empty<UpdateType>()
        };

        botClient.StartReceiving(
            HandleUpdateAsync,
            HandleErrorAsync,
            receiverOptions,
            cancellationToken: cts.Token
        );

        var me = await botClient.GetMeAsync();
        Console.WriteLine($"Start listening for @{me.Username}");

        Console.ReadLine();


        cts.Cancel();
    }

    static async Task HandleUpdateAsync(ITelegramBotClient botClient, Update update, CancellationToken cancellationToken)
    {
        if (update.Type != UpdateType.Message || update.Message!.Type != MessageType.Text)
            return;

        var chatId = update.Message.Chat.Id;
        var messageText = update.Message.Text;

        Console.WriteLine($"Полученно '{messageText}' сообщение в чате {chatId}.");
        var username = update.Message.From.FirstName;
        if (messageText.ToLower() == "/start")
        {
            await botClient.SendTextMessageAsync(
                chatId: chatId,
                text: $"Привет, {username} 👋, этот бот для поиска данных в Excel. Напишите /search для поиска",
                cancellationToken: cancellationToken
            );
            await HandleStartCommandAsync(chatId, cancellationToken);
        }
        else if (messageText.StartsWith("/search ", StringComparison.OrdinalIgnoreCase))
        {
            var query = messageText.Substring(8).Trim();
            var response = GetExcelData(query);

            await botClient.SendTextMessageAsync(
                chatId: chatId,
                text: response,
                cancellationToken: cancellationToken
            );
        }
        else if (messageText.ToLower() == "/info")
        {
            await HandleInfoCommandAsync(chatId, cancellationToken);
        }

        else if (messageText.ToLower() == "/contacts")
        {
            await HandleContactsCommandAsync(chatId, cancellationToken);
        }
        else if (messageText.ToLower() == "info")
        {
            await HandleInfoCommandAsync(chatId, cancellationToken);
        }
        else if (messageText.ToLower() == "contacts")
        {
            await HandleContactsCommandAsync(chatId, cancellationToken);
        }
        else
        {
            await botClient.SendTextMessageAsync(
                chatId: chatId,
                text: "Пожалуйста используйте /search и слово которое вы хотите найти",
                cancellationToken: cancellationToken
            );
        }
    }
    static Task HandleErrorAsync(ITelegramBotClient botClient, Exception exception, CancellationToken cancellationToken)
    {
        var errorMessage = exception switch
        {
            ApiRequestException apiRequestException
                => $"Telegram API Ошибка:\n[{apiRequestException.ErrorCode}]\n{apiRequestException.Message}",
            _ => exception.ToString()
        };

        Console.WriteLine(errorMessage);
        return Task.CompletedTask;
    }
    private static string GetExcelData(string query)
    {
        if (!System.IO.File.Exists(ExcelFilePath))
        {
            return "Файла не существует.";
        }

        using (var package = new ExcelPackage(new FileInfo(ExcelFilePath)))
        {
            var worksheet = package.Workbook.Worksheets.FirstOrDefault();
            if (worksheet == null)
            {
                return "В файле Excel не найден рабочий лист.";
            }

            var rows = worksheet.Dimension.Rows;
            var cols = worksheet.Dimension.Columns;
            var data = "";
            int counter = 1;
            for (int row = 1; row <= rows; row++)
            {
                for (int col = 1; col <= cols; col++)
                {
                    var cellValue = worksheet.Cells[row, col].Text;
                    if (cellValue.Contains(query, StringComparison.OrdinalIgnoreCase))
                    {
                        data += $"{counter}) Нашел '{query}' в ряду {row}, колонка {col}: \n{cellValue}\n";
                        counter++;
                    }
                }
            }

            return string.IsNullOrEmpty(data) ? "Совпадение не было найдено." : data;
        }
    }

    private static async Task HandleInfoCommandAsync(long chatId, CancellationToken cancellationToken)
    {
        if (!System.IO.File.Exists(ExcelFilePath))
        {
            await botClient.SendTextMessageAsync(
                chatId: chatId,
                text: "Файла не существует.",
                cancellationToken: cancellationToken
            );
            return;
        }

        using (var package = new ExcelPackage(new FileInfo(ExcelFilePath)))
        {
            var worksheet = package.Workbook.Worksheets.FirstOrDefault();
            if (worksheet == null)
            {
                await botClient.SendTextMessageAsync(
                    chatId: chatId,
                    text: "В файле Excel не найден рабочий лист.",
                    cancellationToken: cancellationToken
                );
                return;
            }

            var rows = worksheet.Dimension.Rows;
            var cols = worksheet.Dimension.Columns;
            var columns = Enumerable.Range(1, cols).Select(col => worksheet.Cells[1, col].Text).ToList();

            await botClient.SendTextMessageAsync(
                chatId: chatId,
                text: $"Информация о файле ----- \n" +
                      $"Количество строк: {rows} \n" +
                      $"Количество столбцов: {cols} \n" +
                      $"Столбцы: {string.Join(", ", columns)}",
                      cancellationToken: cancellationToken
            );

        }

    }
    private static async Task HandleContactsCommandAsync(long chatId, CancellationToken cancellationToken)
    {
        await botClient.SendTextMessageAsync(
            chatId: chatId,
            text: "Контакты ----- \n" +
      "Разработчик ИИ: @EkatrinaSmith\n" +
      "Публичное представление: @MIR0_3\n" +
      "Создатель чатбота: @EvgeniyKasatkin \n",
            cancellationToken: cancellationToken
        );
    }
    private static async Task HandleStartCommandAsync(long chatId, CancellationToken cancellationToken)
    {
        var replyKeyboard = new ReplyKeyboardMarkup(new[]
        {
        new KeyboardButton[] { "info", "contacts" }
    })
        {
            ResizeKeyboard = true
        };

        await botClient.SendTextMessageAsync(
            chatId: chatId,
            text: "Выберите опцию ниже:",
            replyMarkup: replyKeyboard,
            cancellationToken: cancellationToken
        );
    }
}
