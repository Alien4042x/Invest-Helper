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

namespace InvestHelper
{
    public partial class ViewController : NSViewController
    {
        private bool isProcessing = false;

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

        public async void UpdateProgressBar(double value)
        {
            if (!isProcessing) return;

            InvokeOnMainThread(() =>
            {
                progress_indicator.DoubleValue = value;
            });

            if (progress_indicator.DoubleValue >= 100)
            {
                await Task.Delay(3000); // Wait for 3 second

                InvokeOnMainThread(() =>
                {
                    progress_indicator.Hidden = true;
                    progress_indicator.DoubleValue = 0;
                    btn_generate_click.Enabled = true;
                });

                isProcessing = false;
            }
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

                isProcessing = true;

                progress_indicator.Hidden = false;
                btn_generate_click.Enabled = false;

                UpdateProgressBar(0);

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

                await data.free_cash_flow();

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
                            worksheet.Cell("B5").Value = data.years[0];
                            worksheet.Cell("C5").Value = data.years[1];
                            worksheet.Cell("D5").Value = data.years[2];
                            worksheet.Cell("E5").Value = data.years[3];
                            worksheet.Cell("F5").Value = data.years[4];
                            worksheet.Cell("G5").Value = data.years[5];
                            worksheet.Cell("H5").Value = data.years[6];
                            worksheet.Cell("I5").Value = data.years[7];
                            worksheet.Cell("J5").Value = data.years[8];
                            worksheet.Cell("K5").Value = data.years[9];
                        }
                        else
                        {
                            Console.WriteLine("The 'data.years' list does not have a sufficient number of items.");
                        }

                        //Free cash flow
                        worksheet.Cell("B6").Value = data.freeCashFlowValues[0];
                        worksheet.Cell("C6").Value = data.freeCashFlowValues[1];
                        worksheet.Cell("D6").Value = data.freeCashFlowValues[2];
                        worksheet.Cell("E6").Value = data.freeCashFlowValues[3];
                        worksheet.Cell("F6").Value = data.freeCashFlowValues[4];
                        worksheet.Cell("G6").Value = data.freeCashFlowValues[5];
                        worksheet.Cell("H6").Value = data.freeCashFlowValues[6];
                        worksheet.Cell("I6").Value = data.freeCashFlowValues[7];
                        worksheet.Cell("J6").Value = data.freeCashFlowValues[8];
                        worksheet.Cell("K6").Value = data.freeCashFlowValues[9];

                        // Date Update for Future Free Cash Flow
                        int startColumn = 2; // Starting at Column B
                        int row = 9; // Starting at Row 10

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

                            worksheet.Cell("P6").Value = finalGrowthRate;
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

                                worksheet.Cell("P5").Value = user_growth_rate_value / 100;
                            }
                        }

                        //Perpertual Growth Rate
                        if (perpertual_growth_rate.StringValue.Contains('.'))
                            perpertual_growth_rate.StringValue = perpertual_growth_rate.StringValue.Replace(',', '.');

                        double perpertual_growth_rate_value;
                        if (double.TryParse(perpertual_growth_rate.StringValue, out perpertual_growth_rate_value))
                        {
                            worksheet.Cell("P9").Value = perpertual_growth_rate_value;
                        }

                        //User Discount Rate
                        if (!string.IsNullOrEmpty(user_discount_rate.StringValue))
                        {
                            string cleanedValue = user_discount_rate.StringValue.Replace("%", "").Trim();

                            if (double.TryParse(cleanedValue, out double user_discount_rate_value))
                            {
                                worksheet.Cell("P11").Value = user_discount_rate_value / 100;
                            }
                        }

                        //Stock
                        worksheet.Cell("A2").Value = stock.StringValue;

                        //Cash & Cash Equivalents
                        worksheet.Cell("B14").Value = data.cash_cash_equivalents;

                        //Total Debt
                        worksheet.Cell("B15").Value = data.total_debt;

                        //Shares Outstanding
                        worksheet.Cell("B17").Value = data.shares_outstanding;

                        //Interest Expense
                        worksheet.Cell("O14").Value = data.interest_expense;

                        //Income Tax Expense
                        worksheet.Cell("O17").Value = data.income_tax_expense;

                        //Income Before Tax
                        worksheet.Cell("O18").Value = data.income_before_tax;

                        //Risk Free Rate
                        worksheet.Cell("K14").Value = data.t_yeald_x_years;

                        //Market Cap
                        worksheet.Cell("J22").Value = data.market_cap;

                        //BETA
                        worksheet.Cell("K15").Value = data.beta;

                        //Market Return
                        worksheet.Cell("K16").Value = data.return_rate;

                        //EPS
                        worksheet.Cell("F13").Value = data.eps;

                        //Save
                        workbook.SaveAs(selectedPath);
                    }
                }
            }
            catch(Exception ex)
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

            public GetData(ViewController viewController)
            {
                _viewController = viewController;
            }

            public async Task free_cash_flow()
            {
                var url = $"https://www.macrotrends.net/stocks/charts/{_viewController.stock.StringValue}/{_viewController.stock.StringValue}/free-cash-flow";
                var web = new HtmlWeb();
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

                    _viewController.UpdateProgressBar(12.5); // Update Status
                    await Cash_Cash_Equivalents();
                }
                else
                {
                    Console.WriteLine("The 'tbody' element was not found.");
                }
            }

            public async Task Cash_Cash_Equivalents()
            {
                var url = $"https://www.macrotrends.net/stocks/charts/{_viewController.stock.StringValue}/{_viewController.stock.StringValue}/cash-on-hand";
                var web = new HtmlWeb();
                var doc = await Task.Run(() => web.Load(url));

                // Search for the specific unordered list with the given style attribute
                var ul = doc.DocumentNode.SelectSingleNode("//ul[@style='margin-top:10px;']");

                if (ul != null)
                {
                    // Get the second list item
                    var li = ul.SelectNodes("./li")[1];  // Index 1 for the second item

                    if (li != null)
                    {
                        var strongNode = li.SelectSingleNode("./strong");
                        if (strongNode != null)
                        {
                            string rawValue = strongNode.InnerText.Trim().Replace("$", "");
                            double doubleValue;

                            // Check if the value ends with 'B', 'M', or 'T'
                            if (rawValue.EndsWith("B"))
                            {
                                doubleValue = double.Parse(rawValue.TrimEnd('B')) * 1000; // Convert billions to millions
                            }
                            else if (rawValue.EndsWith("M"))
                            {
                                doubleValue = double.Parse(rawValue.TrimEnd('M')); // Value is already in millions
                            }
                            else if (rawValue.EndsWith("T"))
                            {
                                doubleValue = double.Parse(rawValue.TrimEnd('T')) * 1000000; // Convert trillions to millions
                            }
                            else
                            {
                                doubleValue = double.Parse(rawValue); // Assume the number is in the correct format
                            }

                            cash_cash_equivalents = doubleValue;

                            _viewController.UpdateProgressBar(25);//Update Status

                            await Total_Debt();
                        }
                        else
                        {
                            Console.WriteLine("The 'strong' element was not found.");
                        }
                    }
                    else
                    {
                        Console.WriteLine("The 'li' item was not found.");
                    }
                }
                else
                {
                    Console.WriteLine("The unordered list 'ul' was not found.");
                }
            }

            public async Task growth_estimates()
            {
                try
                {
                    var url = string.Format("https://finance.yahoo.com/quote/{0}/analysis?p={0}", _viewController.stock.StringValue);

                    var web = new HtmlWeb();
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

                    _viewController.UpdateProgressBar(0);//Update Status
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Exception: " + ex.Message);
                }
            }

            public async Task Total_Debt()
            {
                var url = string.Format("https://finance.yahoo.com/quote/{0}/balance-sheet?p={0}", _viewController.stock.StringValue);
                var web = new HtmlWeb();
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

                        _viewController.UpdateProgressBar(37.5); // Update Status

                        await Shares_Outstanding();
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

            public async Task Shares_Outstanding()
            {
                var url = string.Format("https://finance.yahoo.com/quote/{0}/key-statistics?p={0}", _viewController.stock.StringValue);
                var web = new HtmlWeb();
                var doc = await Task.Run(() => web.Load(url));

                // Find the specific div that contains the text "Shares Outstanding"
                var sharesOutstandingRow = doc.DocumentNode.SelectSingleNode("//td[span='Shares Outstanding']/following-sibling::td");

                if (sharesOutstandingRow != null)
                {
                    string rawValue = sharesOutstandingRow.InnerText.Trim();
                    double doubleValue;

                    // Check if the value ends with 'B', 'M', or 'T'
                    if (rawValue.EndsWith("B"))
                    {
                        doubleValue = double.Parse(rawValue.TrimEnd('B')) * 1000; // Convert billions to millions
                    }
                    else if (rawValue.EndsWith("M"))
                    {
                        doubleValue = double.Parse(rawValue.TrimEnd('M')); // Value is already in millions
                    }
                    else if (rawValue.EndsWith("T"))
                    {
                        doubleValue = double.Parse(rawValue.TrimEnd('T')) * 1000000; // Convert trillions to millions
                    }
                    else
                    {
                        doubleValue = double.Parse(rawValue); // Assume the number is in the correct format
                    }

                    shares_outstanding = (int)doubleValue;

                    _viewController.UpdateProgressBar(50);//Update Status

                    await Financials();

                }
                else
                {
                    Console.WriteLine("Div with 'Shares Outstanding' not found.");
                }
            }

            public async Task Financials()
            {
                var url = string.Format("https://finance.yahoo.com/quote/{0}/financials?p={0}", _viewController.stock.StringValue);
                var web = new HtmlWeb();
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

                _viewController.UpdateProgressBar(62.5);//Update Status
                await Bonds();
            }

            public async Task Bonds()
            {
                var web = new HtmlWeb();
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

                    _viewController.UpdateProgressBar(75); // Update Status
                    await Return_Rate();
                }
                else
                {
                    Console.WriteLine("Nodes not found.");
                }
            }

            public async Task Return_Rate()
            {
                var web = new HtmlWeb();
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
                    _viewController.UpdateProgressBar(87.5);//Update Status
                }
                else
                {
                    Console.WriteLine("The 'tbody' element was not found.");
                }

            }

            public async Task Get_last_info()
            {
                var url = string.Format("https://finance.yahoo.com/quote/{0}?p={0}&.tsrc=fin-srch", _viewController.stock.StringValue);
                var web = new HtmlWeb();
                var doc = await Task.Run(() => web.Load(url));

                // Define the row titles you want to extract
                var rowTitles = new List<string> { "Market Cap", "Beta (5Y Monthly)" , "EPS (TTM)" };

                foreach (var title in rowTitles)
                {
                    var col = doc.DocumentNode.SelectSingleNode($"//td[span[text()='{title}']]");

                    if (col != null)
                    {
                        var valueNode = col.NextSibling;

                        if (valueNode != null)
                        {
                            string rawValue = valueNode.InnerText.Trim();
                            char multiplierChar = rawValue.Last(); // Get the last character (B, M, T)

                            // Remove the last character and any decimal comma
                            rawValue = Regex.Replace(rawValue, @"[.](?=\d+$)|[A-Za-z]$", string.Empty);

                            if (double.TryParse(rawValue, out double value))
                            {
                                switch (title)
                                {
                                    case "Market Cap":
                                        switch (multiplierChar)
                                        {
                                            case 'B':
                                                market_cap = value * 1000; // Convert billions to millions
                                                break;
                                            case 'M':
                                                market_cap = value; // Value is already in millions
                                                break;
                                            case 'T':
                                                market_cap = value * 1000000; // Convert trillions to millions
                                                break;
                                            default:
                                                market_cap = value; // Assume the number is in the correct format
                                                break;
                                        }
                                        break;
                                    case "Beta (5Y Monthly)":
                                        beta = value / 100;
                                        break;

                                    case "EPS (TTM)":
                                        eps = value / 100;
                                        break;
                                }
                                _viewController.UpdateProgressBar(100);//Update Status
                            }
                            else
                            {
                                Console.WriteLine($"Failed to convert the value for '{title}' to a number.");

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
