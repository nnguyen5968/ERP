using Microsoft.Extensions.Configuration;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.UI; // Thêm dòng này
using NUnit.Framework;
using OfficeOpenXml; // Cần cài đặt thư viện EPPlus để đọc và ghi Excel
using System;
using System.IO;
using System.Linq;

namespace ERP
{
    public class Tests
    {
        private IWebDriver _driver;
        private string excelPath = @"C:\Users\Administrator\Downloads\ERP.xlsx";
        private IConfiguration _configuration;
        private string _baseUrl;

        [OneTimeSetUp]
        public void OneTimeSetUp()
        {
            var builder = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile(@"D:\C#\AutomationSTC\AutomationSTC\appsettings.json");

            _configuration = builder.Build();
            _baseUrl = _configuration["TestSettings:BaseUrl"];
            string browser = _configuration["TestSettings:Browser"];

            if (browser.Equals("Chrome", StringComparison.OrdinalIgnoreCase))
            {
                _driver = new ChromeDriver();
            }
            else if (browser.Equals("Firefox", StringComparison.OrdinalIgnoreCase))
            {
                _driver = new FirefoxDriver();
            }
            else
            {
                throw new ArgumentException($"Unsupported browser: {browser}");
            }

            _driver.Navigate().GoToUrl(_baseUrl);

            // Thực hiện đăng nhập một lần trước tất cả các test case
            PerformLogin("0000", "123456");
        }

        [SetUp]
        public void SetUp()
        {
            string createUrl = _configuration["TestSettings:BaseUrl"] + "SuppliesTypes";
            _driver.Navigate().GoToUrl(createUrl);
            _driver.Manage().Window.Maximize();
        }

        private void PerformLogin(string username, string password)
        {
            var usernameField = _driver.FindElement(By.Name("username"));
            usernameField.SendKeys(username);

            var passwordField = _driver.FindElement(By.Name("password"));
            passwordField.SendKeys(password);

            var loginButton = _driver.FindElement(By.Id("btnLogin"));
            loginButton.Click();
        }

        [Test]
        public void RunTestWithMultipleXPaths()
        {
            using (var package = new ExcelPackage(new FileInfo(excelPath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Lấy worksheet đầu tiên
                int rowCount = worksheet.Dimension.Rows; // Số lượng hàng

                for (int row = 2; row <= rowCount; row++) // Bắt đầu từ hàng 2 (bỏ qua tiêu đề)
                {
                    // Đọc giá trị từ cột TestCaseID
                    string testCaseID = worksheet.Cells[row, 1].Text;
                    if (string.IsNullOrEmpty(testCaseID))
                    {
                        Console.WriteLine($"TestCaseID is empty in row {row}. Skipping this row.");
                        continue; // Bỏ qua hàng này và tiếp tục với hàng tiếp theo
                    }

                    // Đọc giá trị từ các cột khác
                    string xpaths = worksheet.Cells[row, 2].Text;
                    string inputs = worksheet.Cells[row, 3].Text;
                    string expectedResultXPath = worksheet.Cells[row, 4].Text;
                    string expectedResultValue = worksheet.Cells[row, 5].Text;

                    var xpathList = xpaths.Split(';').Select(s => s.Trim()).ToArray();
                    var inputList = inputs.Split(';').Select(s => s.Trim()).ToArray();

                    if (xpathList.Length != inputList.Length && inputList.Length > 0)
                    {
                        Console.WriteLine($"Warning: Number of XPaths and Inputs do not match in row {row}.");
                    }

                    for (int i = 0; i < xpathList.Length; i++)
                    {
                        string xpath = xpathList[i];
                        string input = i < inputList.Length ? inputList[i] : null;

                        try
                        {
                            var element = _driver.FindElement(By.XPath(xpath));
                            if (!string.IsNullOrEmpty(input))
                            {
                                element.Clear();
                                element.SendKeys(input);
                            }
                            else
                            {
                                // Nếu không có input, thực hiện hành động khác, ví dụ click
                                element.Click();
                            }
                        }
                        catch (NoSuchElementException e)
                        {
                            Console.WriteLine($"Element not found for XPath: {xpath}. Exception: {e.Message}");
                        }
                    }

                    // Kiểm tra kết quả
                    bool testPassed = false;
                    if (!string.IsNullOrEmpty(expectedResultXPath))
                    {
                        var wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(10));
                        try
                        {
                            var resultElement = wait.Until(driver =>
                            {
                                var element = driver.FindElement(By.XPath(expectedResultXPath));
                                return element.Displayed ? element : null;
                            });

                            string resultValue = resultElement.Text.Trim();

                            // So sánh giá trị kết quả với giá trị mong đợi
                            testPassed = resultValue.Contains(expectedResultValue);
                            Console.WriteLine($"Test {testCaseID} executed. Expected result: {expectedResultValue}. Actual result: {resultValue}. Test Passed: {testPassed}");
                        }
                        catch (WebDriverTimeoutException e)
                        {
                            Console.WriteLine($"Expected result element not found or not visible in time for XPath: {expectedResultXPath}. Exception: {e.Message}");
                        }
                    }

                    // Ghi kết quả vào cột Kết quả trong Excel
                    worksheet.Cells[row, 6].Value = testPassed ? "True" : "False";

                    // Reload lại trang sau mỗi test case
                    _driver.Navigate().Refresh();
                }

                // Lưu file Excel sau khi cập nhật
                package.Save();
            }
        }

        [OneTimeTearDown]
        public void OneTimeTearDown()
        {
            if (_driver != null)
            {
                _driver.Quit();
            }
        }
    }
}
