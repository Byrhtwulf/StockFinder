using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;
using System.Threading;
using OpenQA.Selenium.Support.UI;
using System.Diagnostics;
using System.IO;

namespace StockFinder
{
    public class StockScreener
    {
        IWebDriver driver;
        IWebDriver driver2;

        [SetUp]
        public void Initialize()
        {
            driver = new ChromeDriver();
            driver2 = new ChromeDriver();
        }

        [Test]
        public void OpenAppTest()
        {
            driver.Url = "https://finviz.com/screener.ashx?v=111&f=cap_smallover,fa_curratio_o1,fa_debteq_u0.5,fa_estltgrowth_o5,fa_pb_u2,fa_pe_u15,geo_usa&ft=2";
            List<IWebElement> links = driver.FindElements(By.ClassName("screener-link-primary")).ToList();

            List<StockWrapper> stocks = new List<StockWrapper>();
            string mydocpath =
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            // Write the string array to a new file named "WriteLines.txt".
            using (StreamWriter outputFile = new StreamWriter(mydocpath + @"\WriteLines.txt"))
            {
                for (int i = 0; i < 1; i++)
                {
                    StockWrapper newStock = new StockWrapper();
                    newStock.tickerSymbol = links[i].Text;
                    driver2.Url = "http://www.advfn.com/stock-market/NYSE/" + links[i].Text + "/financials?btn=start_date&mode=annual_reports";
                        //Get date selector
                        // select the drop down list
                        var dateDropDown = driver.FindElement(By.Id("start_dateid"));
                        //create select element object 
                        var selectElement = new SelectElement(dateDropDown);
                        // select by text
                        selectElement.SelectByText("2007/12");
                        Thread.Sleep(5000);
                        //Get old book value
                        IWebElement oldBookValueParent = driver2.FindElement(By.XPath("//td[contains(text(),'book value per share')]"));
                        IWebElement oldBookValue = oldBookValueParent.FindElement(By.XPath("following-sibling::*"));
                        Decimal.TryParse(oldBookValue.Text, out newStock.oldBookValue);

                        //Get oldest date, which may not be 10 years behind the current date
                        IWebElement oldYearEndParent = driver2.FindElement(By.XPath("//td[contains(text(),'year end date')]"));
                        IWebElement oldYearEnd = oldYearEndParent.FindElement(By.XPath("following-sibling::*"));
                        newStock.oldYear = oldYearEnd.Text;

                        //Set page to latest date
                        driver2.Url = "http://www.advfn.com/stock-market/NYSE/" + links[i].Text + "/financials?btn=start_date&mode=annual_reports";

                        //Get new book value
                        IWebElement newBookValueParent = driver2.FindElement(By.XPath("//td[contains(text(),'book value per share')]"));
                        IWebElement newBookValue = newBookValueParent.FindElements(By.XPath("following-sibling::*")).LastOrDefault();
                        Decimal.TryParse(newBookValue.Text, out newStock.newBookValue);

                        //Get newest year
                        IWebElement newYearEndParent = driver2.FindElement(By.XPath("//td[contains(text(),'year end date')]"));
                        IWebElement newYearEnd = newYearEndParent.FindElements(By.XPath("following-sibling::*")).LastOrDefault();
                        newStock.newYear = newYearEnd.Text;

                        //Get dividend
                        IWebElement dividendParent = driver2.FindElement(By.XPath("//td[contains(text(),'Dividends Paid Per Share')]"));
                        IWebElement dividend = dividendParent.FindElements(By.XPath("following-sibling::*")).LastOrDefault();
                        Decimal.TryParse(dividend.Text, out newStock.dividend);

                        //Set page to current info
                        driver2.Url = "http://www.advfn.com/stock-market/NYSE/" + links[i].Text + "/financials";
                        IWebElement priceText = driver2.FindElement(By.XPath("//span[contains(text(),'Price')]"));
                        IWebElement priceParent = priceText.FindElement(By.XPath("../.."));
                        IWebElement priceRow = priceParent.FindElement(By.XPath("following-sibling::*"));
                        IWebElement price = priceRow.FindElement(By.XPath("child::*"));
                        Decimal.TryParse(price.Text, out newStock.currentMarketValue);

                        outputFile.WriteLine(newStock.tickerSymbol + "- Book Value in " + newStock.oldYear + ": " +
                            newStock.oldBookValue + ", Book Value in " + newStock.newYear + ": " + newStock.newBookValue + ", Dividend: " +
                            newStock.dividend + ", Current Value: " + newStock.currentMarketValue);

                }

            }
        }

        [TearDown]
        public void Close(){
            driver.Close();
            driver2.Close();
        }

        private class StockWrapper
        {
            public string tickerSymbol;
            public decimal oldBookValue;
            public string oldYear;
            public decimal newBookValue;
            public string newYear;
            public decimal dividend;
            public decimal currentMarketValue; 
        }

    }
}
