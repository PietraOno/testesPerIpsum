using HtmlAgilityPack;
using System;
using System.Net.Http;
using System.Threading.Tasks;
using System.Collections.Generic;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using static System.Runtime.InteropServices.JavaScript.JSType;


namespace WebScrap
{
    class Program
    {
        static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static readonly string ApplicationName = "Google Sheets API";
        static readonly string SpreadsheetId = "1kSfs66OOPxknVh5cN9sKignAD-jNaITziql8_1aP_IE"; //id da planilha que quero que coloque os dados
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



        static async System.Threading.Tasks.Task ScrapeAndWriteToGoogleSheets() //Acho que tenho que fazer um desse para cada site que eu quero então?
        {
            //Site que eu quero pegar as informações
            var url = "https://scholarships.uq.edu.au/scholarships?status[31]=31&type[160]=160";
            var httpClient = new HttpClient();
            var html = await httpClient.GetStringAsync(url);
            var htmlDocument = new HtmlDocument();
            htmlDocument.LoadHtml(html);

            //Todos os elementos de bolsas individuais (cards de bolsa)
            var scholarshipCards = htmlDocument.DocumentNode.SelectNodes("//div[contains(@class, 'card__content__wrapper')]");
            var data = new List<IList<object>>();

            if (scholarshipCards != null)
            {
                foreach (var card in scholarshipCards)
                {
                    //Título da bolsa
                    var titleElement = card.SelectSingleNode(".//h3[contains(@class, 'card__title')]");
                    var titulo = titleElement != null ? titleElement.InnerText.Trim() : "Título não encontrado";
                    Console.WriteLine("Título: " + titulo);

                    //Descrição da bolsa
                    var descriptionElement = card.SelectSingleNode(".//div[contains(@class, 'field-short-description')]");
                    var description = descriptionElement != null ? descriptionElement.InnerText.Trim() : "Descrição não encontrada";
                    Console.WriteLine("Descrição: " + description);

                    //Datas de início e fechamento (se tiver, nem todas tem né)
                    var dateElement = card.SelectSingleNode(".//div[contains(@class, 'field-applications-open')]");
                    var closeDateElement = card.SelectSingleNode(".//div[contains(@class, 'field-applications-close')]");
                    var openDate = dateElement != null ? dateElement.InnerText.Trim() : "Data de início não encontrada";
                    var closeDate = closeDateElement != null ? closeDateElement.InnerText.Trim() : "Data de fechamento não encontrada";
                    Console.WriteLine("Data de início: " + openDate);
                    Console.WriteLine("Data de fechamento: " + closeDate);

                    //Palavras-chave (área de estudo e foco da bolsa)
                    var keywordsElement = card.SelectSingleNode(".//div[contains(@class, 'field-study-area')]");
                    var keywordElement = card.SelectSingleNode(".//div[contains(@class, 'field-scholarship-focus')]");
                    var keywords = keywordsElement != null ? keywordsElement.InnerText.Trim() : "Áreas de estudo não encontradas";
                    var keyword = keywordElement != null ? keywordElement.InnerText.Trim() : "Foco da bolsa não encontrado";

                    var key = keyword + keywords;


                    data.Add(new List<object> { titulo, description, openDate, closeDate, key });
                }

                WriteToGoogleSheets(data);
            }
            else //chora
            {
                Console.WriteLine("Nenhuma bolsa encontrada na página.");
            }
        }

        static void WriteToGoogleSheets(IList<IList<object>> values)
        {
            var range = $"{sheet}!A:E";
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
}
