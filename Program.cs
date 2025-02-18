using System;
using System.IO;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using OfficeOpenXml;
using System.Threading;

namespace EudamedManufacturersScraper
{
    class Program
    {
        static void Main(string[] args)
        {
            var options = new ChromeOptions();
            options.AddArguments("--no-sandbox", "--disable-dev-shm-usage", "--remote-debugging-port=9222", "--disable-gpu", "--window-size=1920,1080");
            var driver = new ChromeDriver(options);

            string url = "https://ec.europa.eu/tools/eudamed/#/screen/search-eo?actorTypeCode=refdata.actor-type.manufacturer&includeHistoricalVersion=true&submitted=true";
            driver.Navigate().GoToUrl(url);
            driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(40);

            SetEntriesPerPage(driver);

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Eudamed_Manufacturers.xlsx");
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Manufacturers");
                int rowIndex = 1;

                // Set Excel Headers
                sheet.Cells[rowIndex, 1].Value = "Actor ID/SRN";
                sheet.Cells[rowIndex, 2].Value = "Version";
                sheet.Cells[rowIndex, 3].Value = "Role";
                sheet.Cells[rowIndex, 4].Value = "Name";
                sheet.Cells[rowIndex, 5].Value = "Abbreviated Name";
                sheet.Cells[rowIndex, 6].Value = "City";
                sheet.Cells[rowIndex, 7].Value = "Country";
                rowIndex++;

                int totalPages = 1000; // Annahme: viele Seiten
                for (int currentPage = 1; currentPage <= totalPages; currentPage++)
                {
                    Console.WriteLine($"Scraping page {currentPage}...");
                    rowIndex = ScrapeManufacturers(driver, sheet, rowIndex);
                    package.SaveAs(new FileInfo(filePath));
                    if (!NavigateToNextPage(driver)) break;
                }
            }
            driver.Quit();
        }

        public static void SetEntriesPerPage(ChromeDriver driver)
        {
            try
            {
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(20));
                var dropdown = wait.Until(ExpectedConditions.ElementToBeClickable(By.CssSelector(".p-dropdown-trigger")));
                dropdown.Click();

                var option50 = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//li[text()='50']")));
                option50.Click();
                
                Console.WriteLine("Einträge pro Seite erfolgreich auf 50 gesetzt.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Fehler beim Setzen der Einträge pro Seite mit Klick: {ex.Message}");
                Console.WriteLine("Versuche, die Einträge pro Seite per JavaScript zu setzen...");

                try
                {
                    IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                    js.ExecuteScript("document.querySelector(\".p-dropdown\").click();");
                    js.ExecuteScript("document.querySelector(\"li[aria-label='50']\").click();");
                    Console.WriteLine("Einträge pro Seite per JavaScript erfolgreich auf 50 gesetzt.");
                    Thread.Sleep(3000); // Warten, bis die Änderungen greifen
                }
                catch (Exception jsEx)
                {
                    Console.WriteLine($"Auch per JavaScript nicht möglich: {jsEx.Message}");
                }
            }
        }

        public static int ScrapeManufacturers(ChromeDriver driver, ExcelWorksheet sheet, int rowIndex)
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(60));
            int retries = 3;  // Falls ein Fehler auftritt, bis zu 3-mal erneut versuchen

            while (retries > 0)
            {
                try
                {
                    var rows = wait.Until(d => d.FindElements(By.CssSelector("tbody tr")));
                    foreach (var row in rows)
                    {
                        var cells = row.FindElements(By.CssSelector("td"));
                        if (cells.Count < 7) continue; // Sicherheitscheck
                        
                        sheet.Cells[rowIndex, 1].Value = cells[0].Text;
                        sheet.Cells[rowIndex, 2].Value = cells[1].Text;
                        sheet.Cells[rowIndex, 3].Value = cells[2].Text;
                        sheet.Cells[rowIndex, 4].Value = cells[3].Text;
                        sheet.Cells[rowIndex, 5].Value = cells[4].Text;
                        sheet.Cells[rowIndex, 6].Value = cells[5].Text;
                        sheet.Cells[rowIndex, 7].Value = cells[6].Text;
                        rowIndex++;
                    }
                    return rowIndex;
                }
                catch (StaleElementReferenceException)
                {
                    Console.WriteLine("StaleElementReferenceException erkannt, versuche erneut...");
                    retries--;
                    Thread.Sleep(2000); // Kurze Pause, damit das Element neu geladen werden kann
                }
            }

            Console.WriteLine("Fehler: Konnte die Tabellenzeilen nicht stabil auslesen.");
            return rowIndex;
        }

        public static bool NavigateToNextPage(ChromeDriver driver)
        {
            int retries = 3; // Mehrere Versuche, falls die Navigation fehlschlägt
            while (retries > 0)
            {
                try
                {
                    var nextPageButtonXPath = "//button[contains(@class,'p-paginator-next')]";
                    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(60));
                    var nextPageButton = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(nextPageButtonXPath)));
                    ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", nextPageButton);
                    nextPageButton.Click();
                    Thread.Sleep(3000); // Warte kurz, damit die Seite sich aktualisieren kann
                    return true;
                }
                catch (StaleElementReferenceException)
                {
                    Console.WriteLine("StaleElementReferenceException beim Navigieren, versuche erneut...");
                    retries--;
                    Thread.Sleep(2000);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"No more pages or error navigating: {ex.Message}");
                    return false;
                }
            }

            Console.WriteLine("Fehler: Konnte nicht zur nächsten Seite navigieren.");
            return false;
        }
    }
}
