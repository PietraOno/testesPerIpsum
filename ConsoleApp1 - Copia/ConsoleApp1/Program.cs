using System;
using System.Net.Http;
using HtmlAgilityPack;

class Program
{
    static async System.Threading.Tasks.Task Main(string[] args)
    {
        var url = "https://pt.wikipedia.org/wiki/Wikip%C3%A9dia:P%C3%A1gina_principal#:~:text=A%20Wikip%C3%A9dia%20%C3%A9%20um%20projeto%20de";
        using (HttpClient client = new HttpClient())
        {
            var response = await client.GetStringAsync(url);
            var doc = new HtmlDocument();
            doc.LoadHtml(response);

            foreach (var link in doc.DocumentNode.SelectNodes("//a[@href]"))
            {
                var href = link.Attributes["href"].Value;
                var title = link.GetAttributeValue("title", string.Empty);

                Console.WriteLine($"Link: {href}");
                if (!string.IsNullOrEmpty(title))
                {
                    Console.WriteLine($"Título: {title}");
                }
                Console.WriteLine();
            }
        }
    }
}
