using ClosedXML.Excel;
using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace WebCrawler
{
    class Program
    {
        static void Main(string[] args)
        {
            var cars = new Program().StartCrawlerAsync().Result.ToList();

            new Program().Export<Car>(cars, @"C:\Users\gpenchev\Downloads\NewCars.xlsx", "Cars");

            Console.WriteLine();

        }
        
        public async Task<List<Car>> StartCrawlerAsync()
        {
            var url = "https://www.cars.bg/carslist.php?subm=1&add_search=1&typeoffer=1&brandId=57&conditions%5B%5D=4&conditions%5B%5D=1";

            var httpClient = new HttpClient();
            var html = await httpClient.GetStringAsync(url);

            var htmlDocument = new HtmlDocument();
            htmlDocument.LoadHtml(html);

            var divs = htmlDocument.DocumentNode.Descendants("div")
                .Where(node => node.GetAttributeValue("class", "")
                .Equals("mdc-card__primary-action")).ToList();

            var cars = new List<Car>();

            foreach (var div in divs)
            {
                var name = div.Descendants("h5").FirstOrDefault().InnerText;
                var specifications = div.Descendants("div").Where(node => node.GetAttributeValue("class", "")
                .Equals("card__secondary mdc-typography mdc-typography--body1 black")).FirstOrDefault().InnerText;
                var description = div.Descendants("div").Where(node => node.GetAttributeValue("class", "")
                .Equals("card__secondary mdc-typography mdc-typography--body2")).FirstOrDefault().InnerText;
                var imageAttribute = div.Descendants("div").Where(node => node.GetAttributeValue("class", "")
                .Equals("mdc-card__media mdc-card__media--16-9")).FirstOrDefault().ChildAttributes("style").SingleOrDefault().Value;
                var price = div.Descendants("h6").Where(node => node.GetAttributeValue("class", "")
                .Equals("card__title mdc-typography mdc-typography--headline6 price")).FirstOrDefault().InnerText;
                var location = div.Descendants("div").Where(node => node.GetAttributeValue("class", "")
                .Equals("card__footer mdc-typography mdc-typography--body2 align-bottom")).FirstOrDefault().InnerText;


                var imgDetails = imageAttribute.Split(';').ToList();
                var imgSplitted = imgDetails[1].Split('&');
                var image = imgSplitted[0];

                var car = new Car
                {
                    Name = name.Trim(),
                    Specifications = specifications.Trim(),
                    Description = description.Trim(),
                    Price = price.Trim(),
                    Location = location.Trim(),
                    Image = image
                };

                cars.Add(car);
            }

            return cars;

        }
        public bool Export<T>(List<T> list, string file, string sheeName)
        {
            bool exported = false;
            using (IXLWorkbook workbook = new XLWorkbook())
            {
                workbook.AddWorksheet(sheeName).FirstCell().InsertTable<T>(list, false);

                workbook.SaveAs(file);
                exported = true;
            }
            return exported;
        }
    }
}
