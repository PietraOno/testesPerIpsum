using System;
using System.IO;
using System.Net.Http;
using HtmlAgilityPack;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using System.Threading;
using System.Collections.Generic;

class Program
{
    static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
    static readonly string ApplicationName = "Google Sheets API";
    static readonly string SpreadsheetId = "1pK_oXYBKskRTNKUYo42qe0ZZODabBo6_9j72jkXBpHE"; //id da planilha que quero que coloque os dados (link+título, nesse caso)
    static readonly string sheet = "Página1";
    static SheetsService service;

    static void Main(string[] args)
    {
        GoogleCredential credential;

        using (var stream = new FileStream("credenciais.json", FileMode.Open, FileAccess.Read))
        {
            credential = GoogleCredential.FromStream(stream).CreateScoped(Scopes);
        }

        service = new SheetsService(new BaseClientService.Initializer()
        {
            HttpClientInitializer = credential,
            ApplicationName = ApplicationName,
        });

        ScrapeAndWriteToGoogleSheets().Wait();
    }

    static async System.Threading.Tasks.Task ScrapeAndWriteToGoogleSheets()
    {
        var url = "https://pt.wikipedia.org/wiki/Wikip%C3%A9dia:P%C3%A1gina_principal";
        using (HttpClient client = new HttpClient())
        {
            var response = await client.GetStringAsync(url);
            var doc = new HtmlDocument();
            doc.LoadHtml(response);

            var data = new List<IList<object>>();

            foreach (var link in doc.DocumentNode.SelectNodes("//a[@href]"))
            {
                var href = link.Attributes["href"].Value;
                var title = link.GetAttributeValue("title", string.Empty);

                data.Add(new List<object> { href, title });
            }

            WriteToGoogleSheets(data);
        }
    }

    static void WriteToGoogleSheets(IList<IList<object>> values)
    {
        var range = $"{sheet}!A:B";
        var valueRange = new ValueRange
        {
            Values = values
        };

        var appendRequest = service.Spreadsheets.Values.Append(valueRange, SpreadsheetId, range);
        appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.RAW;
        var appendResponse = appendRequest.Execute();

        Console.WriteLine("Dados inseridos na planilha com sucesso!");
    }
}
