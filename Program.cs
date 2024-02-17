using Microsoft.Playwright;
using HtmlAgilityPack;
using System.Net.Http;
using System.Text.RegularExpressions;
using System.Data;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using System.IO;

namespace TMK_RBC
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            var playwright = await Playwright.CreateAsync();

            await using var browser = await playwright.Chromium.LaunchAsync(new BrowserTypeLaunchOptions { Headless = false });

            var context = await browser.NewContextAsync();

            context.SetDefaultTimeout(120000);

            var page = await context.NewPageAsync();

            //var page = await browser.NewPageAsync();

            await page.GotoAsync("https://pro.rbc.ru/rbc500");

            //wait 1 seconds
            await page.WaitForTimeoutAsync(1000);

            await page.GetByText("2021", new() { Exact = true }).First.ClickAsync();

            //wait 1 seconds
            await page.WaitForTimeoutAsync(1000);

            await page.Locator(".js-filter-item-container > div").First.ClickAsync();

            //wait 1 seconds
            await page.WaitForTimeoutAsync(1000);

            await page.GetByText("Весь список").ClickAsync();

            //wait 1 seconds
            await page.WaitForTimeoutAsync(1000);

            // Now you can scrape the table or the entire page content
            var content = await page.ContentAsync();

            //Создаем таблицу с колонками для добавления данных.
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("Позиция", typeof(string));
            dt.Columns.Add("Название компании", typeof(string));
            dt.Columns.Add("Город", typeof(string));
            dt.Columns.Add("Индустрия", typeof(string));
            dt.Columns.Add("Выручка, млрд.", typeof(decimal));
            dt.Columns.Add("Прибыль, млрд.", typeof(decimal));

            var htmlDoc = new HtmlAgilityPack.HtmlDocument();

            htmlDoc.LoadHtml(content); // Assuming 'content' contains your HTML

            // Select the container for the table items
            var tableContainer = htmlDoc.DocumentNode.SelectSingleNode("//div[contains(@class, 'rating__table__list__inner')]");

            // Select all individual items within the container
            var tableItems = tableContainer.SelectNodes("./a");

            if (tableItems != null)
            {

                foreach (var item in tableItems)
                {

                    // Position
                    var position = item.SelectSingleNode(".//span[contains(@class, 'rating__company__position')]").InnerText.Trim();

                    position = position.Split(new[] { ' ', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault();

                    // Company Name
                    var companyName = item.SelectSingleNode(".//span[contains(@class, 'rating__company__name')]").InnerText.Trim();

                    //Если в названии компании есть слово Лента или слово Объединенная то пропускаем эту компанию
                    if (companyName.Contains("Лента") || companyName.Contains("Объединенная"))
                    {
                        //continue;
                    }


                    var companyNameClean = Regex.Match(companyName, @"^[^\n]+").Value.Trim();

                    // City
                    var city = item.SelectSingleNode(".//span[contains(@class, 'rating__company__city')]").InnerText.Trim();

                    // Industry
                    var industry = item.SelectSingleNode(".//span[contains(@class, 'rating__company__spec-name')]").InnerText.Trim();

                    var revenueValue = string.Empty;

                    var revenue = item.SelectNodes(".//span[contains(@class, 'js-rating-graph-number active')]").FirstOrDefault()?.InnerText.Trim();

                    var revenueMatch = Regex.Match(revenue, @"₽\s*([\d,\s]+)\s*млрд");

                    if (revenueMatch.Success)
                    {
                        revenueValue = revenueMatch.Groups[1].Value.Trim().Replace(" ", "");
                    }
                    else
                    {
                        revenueValue = "Не удалось извлечь данные";
                    }

                    // Net Gain (most recent year)
                    var netGain = item.SelectNodes(".//span[contains(@class, 'js-rating-graph-number active')]").LastOrDefault()?.InnerText.Trim();

                    var netGainMatch = Regex.Match(netGain, @"₽\s*([\d,\s]+)\s*млрд");

                    var netGainValue = string.Empty;

                    if (netGainMatch.Success)
                    {
                        netGainValue = netGainMatch.Groups[1].Value.Trim().Replace(" ", "");
                    }
                    else
                    {
                        netGainValue = "Не удалось извлечь данные";                       
                    }

                    // Add the data to the DataTable
                    dt.Rows.Add(position, companyNameClean, city, industry, decimal.Parse(revenueValue), decimal.Parse(netGainValue));

                }

                // EPPlus requires a license context for non-commercial or commercial use
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                var fileInfo = new FileInfo("E:\\Csharp\\source\\repos\\TMK_RBC\\RBC500.xlsx");

                using (var package = new ExcelPackage(fileInfo))
                {
                    var worksheet = package.Workbook.Worksheets.Add("Sheet1");

                    // Adding column headers
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        worksheet.Cells[1, i + 1].Value = dt.Columns[i].ColumnName;
                    }

                    // Dumping data
                    for (int row = 0; row < dt.Rows.Count; row++)
                    {
                        for (int col = 0; col < dt.Columns.Count; col++)
                        {
                            worksheet.Cells[row + 2, col + 1].Value = dt.Rows[row][col];
                        }
                    }

                    // Auto-fit columns
                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                    // Formatting headers - bold
                    using (var range = worksheet.Cells[1, 1, 1, dt.Columns.Count])
                    {
                        range.Style.Font.Bold = true;
                    }

                    // Formatting as Table
                    var tblRange = worksheet.Cells[1, 1, dt.Rows.Count + 1, dt.Columns.Count];

                    var table = worksheet.Tables.Add(tblRange, "DataTable");

                    table.TableStyle = TableStyles.Medium2;

                    // Freeze header row
                    worksheet.View.FreezePanes(2, 1);

                    // Adjust Text Alignment for the header
                    using (var range = worksheet.Cells[1, 1, 1, dt.Columns.Count])
                    {
                        range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    }

                    // Add borders to the table
                    tblRange.Style.Border.Top.Style = tblRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    tblRange.Style.Border.Left.Style = tblRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;

                    // Save and close
                    package.Save();
                }



                

                Console.WriteLine("Excel file created.");
            }
            else
            {
                Console.WriteLine("No table items found.");
            }
            


        }
    }
}
