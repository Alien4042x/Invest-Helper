using System;
using AppKit;
using Foundation;
using HtmlAgilityPack;
using System.Collections.Generic;
using System.Reflection.Emit;
using System.Linq;
using System.Security.Policy;
using CoreServices;
using System.IO;
using ClosedXML.Excel;
using System.Windows;
using System.Reflection;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Math;
using DocumentFormat.OpenXml.Spreadsheet;
using CoreFoundation;
using System.Globalization;
using System.Security.Cryptography;
using System.Net;
using CoreAudioKit;
using CoreVideo;

namespace InvestHelper
{
    public partial class ViewController : NSViewController
    {
        public ViewController(IntPtr handle) : base(handle)
        {
        }

        public override void ViewDidLoad()
        {
            base.ViewDidLoad();
            progress_indicator.Hidden = true;
            // Do any additional setup after loading the view.
        }

        public override NSObject RepresentedObject
        {
            get
            {
                return base.RepresentedObject;
            }
            set
            {
                base.RepresentedObject = value;
                // Update the view, if already loaded.
            }
        }

        private void message(string message, string text)
        {
            var alert = new NSAlert()
            {
                AlertStyle = NSAlertStyle.Warning,
                InformativeText = text,
                MessageText = message,
            };
            alert.RunModal();
        }

        partial void btn_calculate(NSObject sender)
        {
            try
            {
                if (CheckEmptyTextBoxes())
                {
                    SaveDataToExcel();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());

                var alert = new NSAlert()
                {
                    AlertStyle = NSAlertStyle.Critical,
                    InformativeText = ex.Message,
                    MessageText = "Error",
                };
                alert.RunModal();
            }
        }

        private bool CheckEmptyTextBoxes()
        {
            bool allFieldsFilled = true;
            if (string.IsNullOrEmpty(stock.StringValue))
            {
                stock_icon.Hidden = false;
                message("Warning", "The ticker is empty");
                allFieldsFilled = false;
            }
            else
            {
                stock_icon.Hidden = true;
            }

            if(string.IsNullOrEmpty(user_growth_rate_textbox.StringValue))
            {
                if (use_analytical_predictions_checkbox.State == NSCellStateValue.Off)
                {
                    user_growth_rate_icon.Hidden = false;
                    message("Warning", "Please enter a value in the text field or select 'Analytical Predictions'.");
                    allFieldsFilled = false;
                }
            }
            else
            {
                user_growth_rate_icon.Hidden = true;
            }

            if (string.IsNullOrEmpty(perpertual_growth_rate.StringValue))
            {
                perpertual_growth_rate_icon.Hidden = false;
                message("Warning", "The text field is empty");
                allFieldsFilled = false;
            }
            else
            {
                perpertual_growth_rate_icon.Hidden = true;
            }

            return allFieldsFilled;
        }

        public void SaveDataToExcel()
        {
            var savePanel = new NSSavePanel();
            savePanel.Title = "Save File As...";
            savePanel.Prompt = "Save";
            savePanel.AllowedFileTypes = new string[] { "xlsx" };

            var result = savePanel.RunModal();
            if (result == 1)
            {
                var selectedPath = savePanel.Url.Path;

                //ProgressManager.IsProcessing = true;

                progress_indicator.Hidden = false;
                btn_generate_click.Enabled = false;

                user_growth_rate_textbox.Window.MakeFirstResponder(null);
                user_discount_rate.Window.MakeFirstResponder(null);

                Excel_Parameters(selectedPath);
            }
        }

        private bool IsValidTicker(string ticker)
        {
            var url = $"https://finance.yahoo.com/quote/{ticker}?p={ticker}";
            var web = new HtmlWeb();
            var doc = web.Load(url);

            var row = doc.DocumentNode.SelectSingleNode("//fin-streamer[@class='Fw(b) Fz(36px) Mb(-4px) D(ib)']");
            if (row != null)
            {
                var values = row.GetAttributeValue("value", "0");
                Console.WriteLine(values);
                return true;
            }
            Console.WriteLine("Wrong Ticker");
            return false;
        }

        private async void Excel_Parameters(string selectedPath)
        {
            try
            {
                // Verify the validity of the ticker
                if (!IsValidTicker(stock.StringValue))
                {
                    // Display an error message to the user
                    InvokeOnMainThread(() =>
                    {
                        var alert = new NSAlert
                        {
                            AlertStyle = NSAlertStyle.Critical,
                            InformativeText = "The entered stock ticker is invalid. Please enter a valid ticker.",
                            MessageText = "Invalid Ticker",
                        };
                        alert.RunModal();
                    });
                    progress_indicator.Hidden = true;
                    progress_indicator.DoubleValue = 0;
                    btn_generate_click.Enabled = true;
                    return;
                }

                GetData data = new GetData(this);

                await data.DownloadAllData();
                //await data.free_cash_flow();

                //var assembly = IntrospectionExtensions.GetTypeInfo(typeof(InvestHelper.ViewController)).Assembly;

                //Stream stream = assembly.GetManifestResourceStream("InvestHelper.Resources.table.xlsx");
                using (FileStream stream = File.Open(Environment.CurrentDirectory + "/Stock_Table.xlsx", FileMode.Open, FileAccess.Read))
                {
                    using (var workbook = new XLWorkbook(stream))
                    {
                        var worksheet = workbook.Worksheet("Sheet1");

                        //Rename
                        worksheet.Name = string.Format("DCF - {0}", stock.StringValue);

                        //Date Update Free Cash Flow
                        if (data.years.Count >= 10)
                        {
                            worksheet.Cell("C35").Value = data.years[0];
                            worksheet.Cell("D35").Value = data.years[1];
                            worksheet.Cell("E35").Value = data.years[2];
                            worksheet.Cell("F35").Value = data.years[3];
                            worksheet.Cell("G35").Value = data.years[4];
                            worksheet.Cell("H35").Value = data.years[5];
                            worksheet.Cell("I35").Value = data.years[6];
                            worksheet.Cell("J35").Value = data.years[7];
                            worksheet.Cell("K35").Value = data.years[8];
                            worksheet.Cell("L35").Value = data.years[9];
                        }
                        else
                        {
                            Console.WriteLine("The 'data.years' list does not have a sufficient number of items.");
                        }

                        //Free cash flow
                        worksheet.Cell("C36").Value = data.freeCashFlowValues[0];
                        worksheet.Cell("D36").Value = data.freeCashFlowValues[1];
                        worksheet.Cell("E36").Value = data.freeCashFlowValues[2];
                        worksheet.Cell("F36").Value = data.freeCashFlowValues[3];
                        worksheet.Cell("G36").Value = data.freeCashFlowValues[4];
                        worksheet.Cell("H36").Value = data.freeCashFlowValues[5];
                        worksheet.Cell("I36").Value = data.freeCashFlowValues[6];
                        worksheet.Cell("J36").Value = data.freeCashFlowValues[7];
                        worksheet.Cell("K36").Value = data.freeCashFlowValues[8];
                        worksheet.Cell("L36").Value = data.freeCashFlowValues[9];

                        // Date Update for Future Free Cash Flow
                        int startColumn = 3; // Starting at Column C
                        int row = 40; // Starting at Row 40

                        int currentYear = DateTime.Now.Year; // Get the current year

                        for (int year = currentYear; year <= currentYear + 8; year++) // Add 8 years to the current year
                        {
                            // Write the year into the cell
                            worksheet.Cell(row, startColumn).Value = year;

                            // Move to the next column
                            startColumn++;
                        }

                        //Growth Rate
                        if (use_analytical_predictions_checkbox.State == NSCellStateValue.On)
                        {
                            await data.growth_estimates();

                            double finalGrowthRate = data.growth_rate / 100; // Convert to a fraction

                            if (conservative_rounding.State == NSCellStateValue.On)
                            {
                                finalGrowthRate = Math.Round(finalGrowthRate, 2); // Round to two decimal places (or whatever precision you prefer)
                            }

                            worksheet.Cell("G25").Value = finalGrowthRate;
                        }

                        //User Growth Rate
                        if (!string.IsNullOrEmpty(user_growth_rate_textbox.StringValue))
                        {
                            string cleanedValue = user_growth_rate_textbox.StringValue.Replace("%", "").Trim();

                            if (double.TryParse(cleanedValue, out double user_growth_rate_value))
                            {
                                if (conservative_rounding.State == NSCellStateValue.On)
                                {
                                    user_growth_rate_value = Math.Round(user_growth_rate_value, 2);
                                }

                                worksheet.Cell("G24").Value = user_growth_rate_value / 100;
                            }
                        }

                        //Perpertual Growth Rate
                        if (perpertual_growth_rate.StringValue.Contains('.'))
                            perpertual_growth_rate.StringValue = perpertual_growth_rate.StringValue.Replace(',', '.');

                        double perpertual_growth_rate_value;
                        if (double.TryParse(perpertual_growth_rate.StringValue, out perpertual_growth_rate_value))
                        {
                            worksheet.Cell("K23").Value = perpertual_growth_rate_value;
                        }

                        //User Discount Rate
                        if (!string.IsNullOrEmpty(user_discount_rate.StringValue))
                        {
                            string cleanedValue = user_discount_rate.StringValue.Replace("%", "").Trim();

                            if (double.TryParse(cleanedValue, out double user_discount_rate_value))
                            {
                                worksheet.Cell("K25").Value = user_discount_rate_value / 100;
                            }
                        }

                        //Stock
                        worksheet.Cell("E6").Value = stock.StringValue;

                        //Cash & Cash Equivalents
                        worksheet.Cell("C22").Value = data.cash_cash_equivalents;

                        //Total Debt
                        worksheet.Cell("G12").Value = data.total_debt;

                        //Shares Outstanding
                        worksheet.Cell("G21").Value = data.shares_outstanding;

                        //Interest Expense
                        worksheet.Cell("G11").Value = data.interest_expense;

                        //Income Tax Expense
                        worksheet.Cell("G14").Value = data.income_tax_expense;

                        //Income Before Tax
                        worksheet.Cell("G15").Value = data.income_before_tax;

                        //Risk Free Rate
                        worksheet.Cell("C11").Value = data.t_yeald_x_years;

                        //Market Cap
                        worksheet.Cell("B27").Value = data.market_cap;

                        //BETA
                        worksheet.Cell("C12").Value = data.beta;

                        //Market Return
                        worksheet.Cell("C13").Value = data.return_rate;

                        //EPS
                        worksheet.Cell("C20").Value = data.eps;

                        //PE
                        worksheet.Cell("K11").Value = data.p_e;

                        //PB
                        worksheet.Cell("K12").Value = data.p_b;

                        //roe
                        worksheet.Cell("K13").Value = data.roe;

                        //Current Ratio
                        worksheet.Cell("K14").Value = data.current_ratio;

                        //Revenue Growth
                        worksheet.Cell("K18").Value = data.revenue_growth;

                        //Profit Growth
                        worksheet.Cell("K19").Value = data.profit_growth;

                        //Dividend Yield
                        worksheet.Cell("K20").Value = data.dividend_yield;

                        //Save
                        workbook.SaveAs(selectedPath);

                        // Add a delay after saving the workbook
                        await Task.Delay(800); // Wait for 8 seconds

                        // Enable the button after the delay
                        btn_generate_click.Enabled = true;
                    }
                }
            }
            catch (WebException ex)
            {
                Console.WriteLine(ex);
            }
            catch (IOException ex)
            {
                Console.WriteLine(ex);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }

        partial void user_growth_rate_textbox_action(NSObject sender)
        {
            NSTextField textField = sender as NSTextField;
            if (textField != null && !textField.StringValue.EndsWith("%"))
            {
                textField.StringValue += "%";
            }

            if(textField.StringValue.Contains("."))
            {
                textField.StringValue = textField.StringValue.Replace(",", ".");
            }
        }

        partial void user_discount_rate_textbox(NSObject sender)
        {
            NSTextField textField = sender as NSTextField;
            if (textField != null && !textField.StringValue.EndsWith("%"))
            {
                textField.StringValue += "%";
            }

            if (textField.StringValue.Contains("."))
            {
                textField.StringValue = textField.StringValue.Replace(",", ".");
            }
        }

        partial void perpertual_growth_rate_textbox(NSObject sender)
        {
            NSTextField textField = sender as NSTextField;
            if (textField != null && !textField.StringValue.EndsWith("%"))
            {
                textField.StringValue += "%";
            }
            if (textField.StringValue.Contains("."))
            {
                textField.StringValue = textField.StringValue.Replace(",", ".");
            }
        }

        public async Task ProcessTasksWithProgressUpdate(List<Task> tasks)
        {
            var progressManager = new ProgressManager(this, tasks.Count);

            var runningTasks = tasks.ToList(); // Vytvořte kopii seznamu úkolů, protože budeme odstraňovat dokončené úkoly

            while (runningTasks.Count > 0)
            {
                var completedTask = await Task.WhenAny(runningTasks);
                runningTasks.Remove(completedTask);

                progressManager.IncrementStep();
            }
        }

        private class ProgressManager
        {
            private ViewController _viewController;
            private double currentStep = 0;
            private double totalSteps;
            private double incrementValue;

            public bool IsProcessing { get; private set; }

            public ProgressManager(ViewController viewController, int totalSteps)
            {
                IsProcessing = true;
                _viewController = viewController;
                this.totalSteps = totalSteps;
                this.incrementValue = 100.0 / totalSteps;
            }

            public void IncrementStep()
            {
                currentStep++;
                UpdateProgressBar(incrementValue);
            }

            public void UpdateProgressBar(double incrementValue)
            {
                if (!IsProcessing) return;

                _viewController.InvokeOnMainThread(() =>
                {
                    _viewController.progress_indicator.DoubleValue += incrementValue;
                });

                if (_viewController.progress_indicator.DoubleValue >= 100)
                {

                    _viewController.InvokeOnMainThread(() =>
                    {
                        _viewController.progress_indicator.Hidden = true;
                        _viewController.progress_indicator.DoubleValue = 0;
                    });

                    IsProcessing = false;
                }
            }
        }

        private class GetData
        {
            private ViewController _viewController;
            public List<double> freeCashFlowValues = new List<double>();
            public List<int> years = new List<int>();
            public double cash_cash_equivalents = 0;
            public double total_debt = 0;
            public double shares_outstanding = 0;
            public double interest_expense = 0;
            public double income_tax_expense = 0;
            public double income_before_tax = 0;
            public double market_cap = 0;
            public double beta = 0;
            public double return_rate = 0;
            public double t_yeald_x_years = 0;
            public double eps = 0;
            public double growth_rate = 0;
            public double p_e = 0;
            public double p_b = 0;
            public double dividend_yield = 0;
            public double revenue_growth = 0;
            public double profit_growth = 0;
            public double current_ratio = 0;
            public double roe = 0;

            public GetData(ViewController viewController)
            {
                _viewController = viewController;
            }

            /// <summary>
            /// Initiates the download and processing of various financial data for a stock.
            /// </summary>
            /// <remarks>
            /// This method aggregates multiple asynchronous tasks that fetch and process different financial metrics and data points for a stock.
            /// The tasks include retrieving free cash flow, cash equivalents, growth estimates, total debt, statistical data, financials, bond data, return rate, and the latest stock information.
            /// The tasks are processed concurrently using the `Task.WhenAll` method for efficiency.
            /// Additionally, the progress of these tasks is tracked and updated using the `_viewController.ProcessTasksWithProgressUpdate` method.
            /// </remarks>
            public async Task DownloadAllData()
            {
                var tasks = new List<Task>
                {
                    free_cash_flow(),
                    Cash_Equivalents(),
                    growth_estimates(),
                    Total_Debt(),
                    Statistics(),
                    Financials(),
                    Bonds(),
                    Return_Rate(),
                    Get_last_info()
                };

                await _viewController.ProcessTasksWithProgressUpdate(tasks);
                await Task.WhenAll(tasks);
            }

            /// <summary>
            /// Retrieves the free cash flow values for the past 10 years for a stock from Macrotrends.
            /// </summary>
            /// <remarks>
            /// This method navigates to the Macrotrends website for the specified stock ticker.
            /// It extracts the free cash flow values for the past 10 years from the website.
            /// The extracted year and corresponding free cash flow values are stored in the `years` and `freeCashFlowValues` class lists respectively.
            /// If the expected elements are not found on the webpage or the values cannot be parsed, 
            /// appropriate messages are logged.
            /// After extracting the free cash flow values, the method proceeds to retrieve the cash and cash equivalents.
            /// </remarks>
            public async Task free_cash_flow()
            {
                var url = $"https://www.macrotrends.net/stocks/charts/{_viewController.stock.StringValue}/{_viewController.stock.StringValue}/free-cash-flow";
                var web = new HtmlWeb();
                web.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36";
                var doc = await Task.Run(() => web.Load(url));

                var tbody = doc.DocumentNode.SelectSingleNode("//tbody");

                if (tbody != null)
                {
                    var rows = tbody.SelectNodes("./tr").Take(10);

                    foreach (var row in rows)
                    {
                        var yearNode = row.SelectSingleNode("td[1]");
                        var valueNode = row.SelectSingleNode("td[2]");

                        if (yearNode != null && valueNode != null)
                        {
                            string yearText = yearNode.InnerText.Trim();
                            string valueText = valueNode.InnerText.Trim().Replace(",", "").Split('.')[0]; // Removing commas and decimals

                            if (int.TryParse(yearText, out int year) && int.TryParse(valueText, out int value))
                            {
                                years.Add(year);
                                freeCashFlowValues.Add(value);
                            }
                            else
                            {
                                Console.WriteLine($"Failed to convert '{yearText}' or '{valueText}' to an integer.");
                            }
                        }
                        else
                        {
                            Console.WriteLine("Failed to extract year or value.");
                        }
                    }

                    years.Reverse();
                    freeCashFlowValues.Reverse();

                    await Cash_Equivalents();
                }
                else
                {
                    Console.WriteLine("The 'tbody' element was not found.");
                }
            }

            /// <summary>
            /// Retrieves the cash and cash equivalents value for a stock from Macrotrends.
            /// </summary>
            /// <remarks>
            /// This method navigates to the Macrotrends website for the specified stock ticker.
            /// It extracts the cash and cash equivalents value from the website.
            /// The extracted value is then stored in the `cash_cash_equivalents` class variable.
            /// If the expected elements are not found on the webpage or the value cannot be parsed, 
            /// appropriate messages are logged.
            /// After extracting the cash and cash equivalents, the method proceeds to retrieve the total debt.
            /// </remarks>
            public async Task Cash_Equivalents()
            {
                var url = $"https://www.macrotrends.net/stocks/charts/{_viewController.stock.StringValue}/{_viewController.stock.StringValue}/cash-on-hand";
                var web = new HtmlWeb();
                web.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36";
                var doc = await Task.Run(() => web.Load(url));

                var ul = doc.DocumentNode.SelectSingleNode("//ul[@style='margin-top:10px;']");
                if (ul == null)
                {
                    Console.WriteLine("The unordered list 'ul' was not found.");
                    return;
                }

                var li = ul.SelectNodes("./li")?.ElementAtOrDefault(1);
                if (li == null)
                {
                    Console.WriteLine("The 'li' item was not found.");
                    return;
                }

                var strongNode = li.SelectSingleNode("./strong");
                if (strongNode == null)
                {
                    Console.WriteLine("The 'strong' element was not found.");
                    return;
                }

                string rawValue = strongNode.InnerText.Trim().Replace("$", "");
                cash_cash_equivalents = ConvertValueToMillions(rawValue);

                await Total_Debt();
            }

            /// <summary>
            /// Converts a string value to its equivalent in millions.
            /// Handles values ending with 'B' (billions), 'M' (millions), 'T' (trillions), and '%' (percentage).
            /// </summary>
            /// <param name="value">The string value to convert.</param>
            /// <returns>The converted value in millions or the original value if no specific format is detected.</returns>
            private double ConvertValueToMillions(string value)
            {
                try
                {
                    // If the value ends with 'B', convert from billions to millions
                    if (value.EndsWith("B"))
                    {
                        return double.Parse(value.TrimEnd('B')) * 1000;
                    }
                    // If the value ends with 'M', it's already in millions
                    else if (value.EndsWith("M"))
                    {
                        return double.Parse(value.TrimEnd('M'));
                    }
                    // If the value ends with 'T', convert from trillions to millions
                    else if (value.EndsWith("T"))
                    {
                        return double.Parse(value.TrimEnd('T')) * 1000000;
                    }
                    // If the value ends with '%', just parse the percentage value
                    else if (value.EndsWith("%"))
                    {
                        return double.Parse(value.TrimEnd('%'));
                    }
                    // If no specific format is detected, assume the number is in the correct format
                    else
                    {
                        return double.Parse(value);
                    }
                }
                catch (FormatException)
                {
                    Console.WriteLine($"Invalid format for value: {value}");
                    return 0; // Return a default value or throw an exception
                }
            }

            /// <summary>
            /// Retrieves the growth estimate for the next 5 years (per annum) for a given stock from Yahoo Finance.
            /// </summary>
            /// <remarks>
            /// This method navigates to the Yahoo Finance analysis page for the specified stock and extracts the growth rate.
            /// If the growth rate is found, it is stored in the 'growth_rate' variable. If not, appropriate messages are logged.
            /// </remarks>
            public async Task growth_estimates()
            {
                try
                {
                    var url = string.Format("https://finance.yahoo.com/quote/{0}/analysis?p={0}", _viewController.stock.StringValue);
                    var web = new HtmlWeb();
                    web.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36";
                    var doc = await Task.Run(() => web.Load(url));

                    var row = doc.DocumentNode.SelectSingleNode("//tr[td/span[text()='Next 5 Years (per annum)']]");

                    if (row != null)
                    {
                        var values = row.SelectNodes(".//td");

                        if (values != null && values.Count > 1)
                        {
                            string value = values[1].InnerText;
                            value = value.TrimEnd('%').Replace(",", ".");

                            if (double.TryParse(value, out double growthRateValue))
                            {
                                growth_rate = growthRateValue;
                            }
                            else
                            {
                                Console.WriteLine("Cannot convert '" + value + "' to a number.");
                            }
                        }
                        else
                        {
                            Console.WriteLine("No values found for 'Next 5 Years (per annum)'.");
                        }
                    }
                    else
                    {
                        Console.WriteLine("Row with 'Next 5 Years (per annum)' not found.");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Exception: " + ex.Message);
                }
            }

            /// <summary>
            /// Retrieves the total debt for a given stock from Yahoo Finance's balance sheet page.
            /// </summary>
            /// <remarks>
            /// This method navigates to the Yahoo Finance balance sheet page for the specified stock and extracts the total debt value.
            /// If the total debt value is found, it is stored in the 'total_debt' variable. If not, appropriate messages are logged.
            /// Additionally, after retrieving the total debt, the method calls 'Statistics' to fetch the number of outstanding shares.
            /// </remarks>
            public async Task Total_Debt()
            {
                var url = string.Format("https://finance.yahoo.com/quote/{0}/balance-sheet?p={0}", _viewController.stock.StringValue);
                var web = new HtmlWeb();
                web.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36";
                var doc = await Task.Run(() => web.Load(url));

                // Search for the specific div containing the text "Total Debt"
                var totalDebtDiv = doc.DocumentNode.SelectSingleNode("//div[div[@title='Total Debt']]");

                if (totalDebtDiv != null)
                {
                    // Find the parent div that contains the entire row
                    var rowDiv = totalDebtDiv.ParentNode.ParentNode;

                    // Extract values from this row
                    var values = rowDiv.SelectNodes(".//div[@data-test='fin-col']/span");
                    if (values != null && values.Count > 0)
                    {
                        string rawValue = values[0].InnerText.Replace(",", ""); // Remove commas

                        // Remove trailing zeros
                        string trimmedValue = Regex.Replace(rawValue, "0+$", "");

                        double doubleValue = double.Parse(trimmedValue); // Convert to double

                        total_debt = (int)doubleValue;

                        await Statistics();
                    }
                    else
                    {
                        Console.WriteLine("No values found for 'Total Debt'.");
                    }
                }
                else
                {
                    Console.WriteLine("Div with 'Total Debt' not found.");
                }
            }

            /// <summary>
            /// Retrieves key financial statistics for a given stock from Yahoo Finance's key statistics page.
            /// </summary>
            /// <remarks>
            /// This method navigates to the Yahoo Finance key statistics page for the specified stock and extracts various financial metrics.
            /// The metrics include Shares Outstanding, Trailing P/E, Price/Book, Forward Annual Dividend Yield, Quarterly Revenue Growth, 
            /// Quarterly Earnings Growth, Current Ratio, and Return on Equity.
            /// Each metric is stored in its respective variable if found. If a particular metric is not found on the page, 
            /// an appropriate message is logged.
            /// </remarks>
            public async Task Statistics()
            {
                var url = string.Format("https://finance.yahoo.com/quote/{0}/key-statistics?p={0}", _viewController.stock.StringValue);
                var web = new HtmlWeb();
                web.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36";
                var doc = await Task.Run(() => web.Load(url));

                // Define the row titles you want to extract
                var rowTitles = new List<string> { "Shares Outstanding", "Trailing P/E", "Price/Book", "Forward Annual Dividend Yield",
                    "Quarterly Revenue Growth", "Quarterly Earnings Growth", "Current Ratio", "Return on Equity" };

                foreach (var title in rowTitles)
                {
                    var col = doc.DocumentNode.SelectSingleNode($"//td[span[text()='{title}']]");
                    if (col != null)
                    {
                        var valueNode = col.NextSibling;
                        if (valueNode != null)
                        {
                            string rawValue = valueNode.InnerText.Trim();
                            double convertedValue = ConvertValueToMillions(rawValue);

                            switch (title)
                            {
                                case "Shares Outstanding":
                                    shares_outstanding = convertedValue;
                                    break;
                                case "Trailing P/E":
                                    p_e = convertedValue / 100;
                                    break;
                                case "Price/Book":
                                    p_b = convertedValue / 100;
                                    break;
                                case "Forward Annual Dividend Yield":
                                    dividend_yield = convertedValue;
                                    break;
                                case "Quarterly Revenue Growth":
                                    revenue_growth = convertedValue;
                                    break;
                                case "Quarterly Earnings Growth":
                                    profit_growth = convertedValue;
                                    break;
                                case "Current Ratio":
                                    current_ratio = convertedValue;
                                    break;
                                case "Return on Equity":
                                    roe = convertedValue / 100;
                                    break;
                            }
                        }
                        else
                        {
                            Console.WriteLine($"No value found for '{title}'.");
                        }
                    }
                    else
                    {
                        Console.WriteLine($"Row with '{title}' not found.");
                    }
                }
            }

            /// <summary>
            /// Retrieves key financial data for a given stock from Yahoo Finance's financials page.
            /// </summary>
            /// <remarks>
            /// This method navigates to the Yahoo Finance financials page for the specified stock and extracts various financial metrics.
            /// The metrics include Interest Expense, Tax Provision, and Pretax Income.
            /// Each metric is stored in its respective variable if found. If a particular metric is not found on the page, 
            /// an appropriate message is logged. After extracting the metrics, the method proceeds to retrieve bond information.
            /// </remarks>
            public async Task Financials()
            {
                var url = string.Format("https://finance.yahoo.com/quote/{0}/financials?p={0}", _viewController.stock.StringValue);
                var web = new HtmlWeb();
                web.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36";
                var doc = await Task.Run(() => web.Load(url));

                // Define the row titles you want to extract
                var rowTitles = new List<string> { "Interest Expense", "Tax Provision", "Pretax Income" };

                foreach (var title in rowTitles)
                {
                    var col = doc.DocumentNode.SelectSingleNode($"//div[div[@title='{title}']]");

                    if (col != null)
                    {
                        // Search for the specific div containing the given text
                        var rowDiv = col.ParentNode.ParentNode;

                        // Find the parent div that contains the entire row
                        var values = rowDiv.SelectNodes(".//div[@data-test='fin-col']/span");
                        if (values != null && values.Count > 1)
                        {
                            // Extract values from this row
                            string trimmedValue = Regex.Replace(values[1].InnerText, "(,0+)$", ",");

                            string rawValue = trimmedValue.Replace(",", "");

                            if (double.TryParse(rawValue, out double financialValue))
                            {
                                switch (title)
                                {
                                    case "Interest Expense":
                                        interest_expense = financialValue;
                                        break;
                                    case "Tax Provision":
                                        income_tax_expense = financialValue;
                                        break;
                                    case "Pretax Income":
                                        income_before_tax = financialValue;
                                        break;
                                }
                            }
                            else
                            {
                                Console.WriteLine("Failed to convert the value to a number.");
                            }
                        }
                        else
                        {
                            Console.WriteLine("No values found for 'Interest Expense'.");
                        }
                    }
                    else
                    {
                        Console.WriteLine("Div with 'Interest Expense' not found.");
                    }
                }
                await Bonds();
            }

            /// <summary>
            /// Retrieves the bond yield for a specific bond from Yahoo Finance's bonds page.
            /// </summary>
            /// <remarks>
            /// This method navigates to the Yahoo Finance bonds page and extracts the bond yield for a specific bond.
            /// The bond yield is then stored in the variable 't_yeald_x_years' if successfully extracted.
            /// If the bond yield is not found or cannot be parsed, an appropriate message is logged.
            /// After extracting the bond yield, the method proceeds to retrieve the return rate.
            /// </remarks>
            public async Task Bonds()
            {
                var web = new HtmlWeb();
                web.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36";
                var doc = await Task.Run(() => web.Load("https://finance.yahoo.com/bonds"));

                var nodes = doc.DocumentNode.SelectNodes("//td[@aria-label='Last Price']/fin-streamer[@data-field='regularMarketPrice']");

                if (nodes != null)
                {
                    var node = nodes.ElementAtOrDefault(2);

                    if (node != null)
                    {
                        var valueAttribute = node.GetAttributeValue("value", string.Empty);
                        if (!string.IsNullOrEmpty(valueAttribute) && double.TryParse(valueAttribute, out double value))
                        {
                            t_yeald_x_years = Math.Round(value, 2) / 100;
                        }
                        else
                        {
                            Console.WriteLine("Failed to parse the value.");
                        }
                    }
                    else
                    {
                        Console.WriteLine("Node at index 2 not found.");
                    }
                    await Return_Rate();
                }
                else
                {
                    Console.WriteLine("Nodes not found.");
                }
            }

            /// <summary>
            /// Retrieves the annual return rate for the S&P 500 from the Macrotrends website.
            /// </summary>
            /// <remarks>
            /// This method navigates to the Macrotrends website that provides historical annual returns for the S&P 500.
            /// It extracts the most recent annual return rate and stores it in the 'return_rate' variable.
            /// If the return rate is not found or cannot be parsed, an appropriate message is logged.
            /// After extracting the return rate, the method proceeds to retrieve the last available information.
            /// </remarks>
            public async Task Return_Rate()
            {
                var web = new HtmlWeb();
                web.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36";
                var doc = await Task.Run(() => web.Load("https://www.macrotrends.net/2526/sp-500-historical-annual-returns"));

                var tbody = doc.DocumentNode.SelectSingleNode("//tbody");

                if (tbody != null)
                {
                    var rows = tbody.SelectNodes("./tr").Take(1);

                    foreach (var row in rows)
                    {
                        var valueNode = row.SelectSingleNode("td[7]");

                        string rawValue = valueNode.InnerText.Trim();

                        rawValue = Regex.Replace(rawValue, @"[.](?=\d+$)|[%]$", string.Empty);

                        if (valueNode != null)
                        {
                            if (double.TryParse(rawValue, out double value))
                            {
                                return_rate = value / 100;
                            }
                        }
                        else
                        {
                            Console.WriteLine("Failed to extract year or value.");
                        }
                    }

                    await Get_last_info();
                }
                else
                {
                    Console.WriteLine("The 'tbody' element was not found.");
                }

            }

            /// <summary>
            /// Retrieves specific financial information for a stock from Yahoo Finance.
            /// </summary>
            /// <remarks>
            /// This method navigates to the Yahoo Finance website for the specified stock ticker.
            /// It extracts key financial metrics such as "Market Cap", "Beta (5Y Monthly)", and "EPS (TTM)".
            /// The extracted values are then stored in the respective class variables.
            /// If a particular metric is not found or cannot be parsed, an appropriate message is logged.
            /// </remarks>
            public async Task Get_last_info()
            {
                var url = string.Format("https://finance.yahoo.com/quote/{0}?p={0}&.tsrc=fin-srch", _viewController.stock.StringValue);
                var web = new HtmlWeb();
                web.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36";
                var doc = await Task.Run(() => web.Load(url));

                // Define the row titles you want to extract
                var rowTitles = new List<string> { "Market Cap", "Beta (5Y Monthly)", "EPS (TTM)" };

                foreach (var title in rowTitles)
                {
                    var col = doc.DocumentNode.SelectSingleNode($"//td[span[text()='{title}']]");

                    if (col != null)
                    {
                        var valueNode = col.NextSibling;

                        if (valueNode != null)
                        {
                            string rawValue = valueNode.InnerText.Trim();

                            switch (title)
                            {
                                case "Market Cap":
                                    market_cap = ConvertValueToMillions(rawValue);
                                    break;
                                case "Beta (5Y Monthly)":
                                    if (double.TryParse(rawValue, out double betaValue))
                                    {
                                        beta = betaValue;
                                    }
                                    break;
                                case "EPS (TTM)":
                                    if (double.TryParse(rawValue, out double epsValue))
                                    {
                                        eps = epsValue;
                                    }
                                    break;
                            }
                        }
                        else
                        {
                            Console.WriteLine($"No value found for '{title}'.");
                        }
                    }
                    else
                    {
                        Console.WriteLine($"Row with '{title}' not found.");
                    }
                }
            }
        }
    }
}
