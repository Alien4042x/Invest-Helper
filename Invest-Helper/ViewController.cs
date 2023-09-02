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
using System.Windows;
using ClosedXML.Excel;
using System.Reflection;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Math;
using DocumentFormat.OpenXml.Spreadsheet;

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
                InformativeText = message,
                MessageText = text,
            };
            alert.RunModal();
        }

        partial void btn_calculate(NSObject sender)
        {
            try
            {
                if (CheckEmptyTextBoxes())
                {
                    isProcessing = true;
                    UpdateProgressBar(0);
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
                message("Warning", "The text field is empty");
                allFieldsFilled = false;
            }
            else
            {
                stock_icon.Hidden = true;
            }

            if(string.IsNullOrEmpty(growth_rate.StringValue))
            {
                growth_rate_icon.Hidden = false;
                message("Warning", "The text field is empty");
                allFieldsFilled = false;
            }
            else
            {
                growth_rate_icon.Hidden = true;
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

            if (string.IsNullOrEmpty(discount_rate.StringValue))
            {
                discount_rate_icon.Hidden = false;
                message("Warning", "The text field is empty");
                allFieldsFilled = false;
            }
            else
            {
                discount_rate_icon.Hidden = true;
            }

            return allFieldsFilled;
        }

        public async void UpdateProgressBar(double value)
        {
            if (isProcessing)
            {
                progress_indicator.Hidden = false;
                btn_generate_click.Enabled = false;
            }

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

                Excel_Parameters(selectedPath);
            }
        }

        private async void Excel_Parameters(string selectedPath)
        {
            GetData data = new GetData(this);
            await data.cash_flow();

            var assembly = IntrospectionExtensions.GetTypeInfo(typeof(InvestHelper.ViewController)).Assembly;
            Stream stream = assembly.GetManifestResourceStream("InvestHelper.Resources.table.xlsx");
            using (var workbook = new XLWorkbook(stream))
            {
                var worksheet = workbook.Worksheet("Sheet1");

                //Rename
                worksheet.Name = string.Format("DCF - {0}",stock.StringValue);

                //Date Update
                worksheet.Cell("C6").Value = data.years[0];
                worksheet.Cell("D6").Value = data.years[1];
                worksheet.Cell("E6").Value = data.years[2];
                worksheet.Cell("F6").Value = data.years[3];

                //Growth Rate
                if (growth_rate.StringValue.Contains('.'))
                    growth_rate.StringValue = growth_rate.StringValue.Replace('.', ',');

                worksheet.Cell("M6").Value = growth_rate.StringValue;

                //Perpertual Growth Rate
                if (perpertual_growth_rate.StringValue.Contains('.'))
                    perpertual_growth_rate.StringValue = perpertual_growth_rate.StringValue.Replace('.', ',');

                worksheet.Cell("M8").Value = perpertual_growth_rate.StringValue;

                //Discount Rate
                if (discount_rate.StringValue.Contains('.'))
                    discount_rate.StringValue = discount_rate.StringValue.Replace('.', ',');

                worksheet.Cell("M9").Value = discount_rate.StringValue;

                // Start at cell C13
                int startColumn = 3; // Column C
                int row = 13; // Row 13

                for (int year = data.years[3]; year <= 8; year++)
                {
                    // Write the year into the cell
                    worksheet.Cell(row, startColumn).Value = year;

                    // Move to the next column
                    startColumn++;
                }

                //Stock
                worksheet.Cell("A2").Value = stock.StringValue;

                //Free cash flow
                worksheet.Cell("C7").Value = data.freeCashFlowValues[0];
                worksheet.Cell("D7").Value = data.freeCashFlowValues[1];
                worksheet.Cell("E7").Value = data.freeCashFlowValues[2];
                worksheet.Cell("F7").Value = data.freeCashFlowValues[3];

                //Cash & Cash Equivalents
                worksheet.Cell("C19").Value = data.cash_cash_equivalents;

                //Total Debt
                worksheet.Cell("C20").Value = data.total_debt;

                //Shares Outstanding
                worksheet.Cell("C22").Value = data.shares_outstanding;

                //Save
                workbook.SaveAs(selectedPath);
            }
        }

        partial void growth_rate_textbox(NSObject sender)
        {
            NSTextField textField = sender as NSTextField;
            if (textField != null && !textField.StringValue.EndsWith("%"))
            {
                textField.StringValue += "%";
            }

            if(textField.StringValue.Contains("."))
            {
                textField.StringValue = textField.StringValue.Replace(".", ",");
            }
        }

        partial void discount_rate_textbox(NSObject sender)
        {
            NSTextField textField = sender as NSTextField;
            if (textField != null && !textField.StringValue.EndsWith("%"))
            {
                textField.StringValue += "%";
            }
            if (textField.StringValue.Contains("."))
            {
                textField.StringValue = textField.StringValue.Replace(".", ",");
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
                textField.StringValue = textField.StringValue.Replace(".", ",");
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

            public GetData(ViewController viewController)
            {
                _viewController = viewController;
            }

            public async Task cash_flow()
            {
                var url = string.Format("https://finance.yahoo.com/quote/{0}/cash-flow?p={0}", _viewController.stock.StringValue);
                var web = new HtmlWeb();
                var doc = await Task.Run(() => web.Load(url));

                var parentDiv = doc.DocumentNode.SelectSingleNode("//div[@class='D(tbr) C($primaryColor)']");

                if (parentDiv != null)
                {
                    // Find all divs that contain dates, within the parent div
                    var dateDivs = parentDiv.SelectNodes(".//div[contains(@class, 'Ta(c)') and contains(@class, 'Py(6px)') and contains(@class, 'Bxz(bb)')]/span");

                    if (dateDivs != null)
                    {
                        // Skip the first element (ttm) and iterate over the remaining
                        for (int i = 1; i < dateDivs.Count; i++)
                        {
                            // Split the date and get the year
                            int year = int.Parse(dateDivs[i].InnerText.Split('/').Last());

                            // Add the year to the list
                            years.Add(year);
                        }
                    }
                    else
                    {
                        Console.WriteLine("No dates found.");
                    }
                }
                else
                {
                    Console.WriteLine("Parent div not found.");
                }

                years.Reverse();

                // Find the specific div that contains the text "Free Cash Flow"
                var freeCashFlowDiv = doc.DocumentNode.SelectSingleNode("//div[div[@title='Free Cash Flow']]");

                if (freeCashFlowDiv != null)
                {
                    // Find the parent div that the entire row
                    var rowDiv = freeCashFlowDiv.ParentNode.ParentNode;

                    // Print the values ​​from this line
                    var values = rowDiv.SelectNodes(".//div[@data-test='fin-col']/span");
                    if (values != null)
                    {
                        foreach (var value in values)
                        {
                            string rawValue = value.InnerText.Replace(",", ""); // Remove commas
                            double doubleValue;

                            // Check if the number ends with zeros
                            if (rawValue.EndsWith("000"))
                            {
                                doubleValue = double.Parse(rawValue) / 1000000; // Convert the number to a double with two decimal places
                            }
                            else
                            {
                                doubleValue = double.Parse(rawValue); // Keep the number as it is
                            }

                            // Multiply the double value by 1000 and convert to int
                            int intValue = (int)(doubleValue * 1000);
                            freeCashFlowValues.Add(intValue); // Add the value to the list

                            _viewController.UpdateProgressBar(25);//Update Status
                        }

                        freeCashFlowValues.RemoveAt(0); //Remove TTM
                        freeCashFlowValues.Reverse(); //Reverse list

                        await Cash_Cash_Equivalents();
                    }
                    else
                    {
                        Console.WriteLine("No values found for 'Free Cash Flow'.");
                    }
                }
                else
                {
                    Console.WriteLine("Div with 'Free Cash Flow' not found.");
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
                            var cashOnHandStr = strongNode.InnerText;

                            // Remove all non-numeric characters except for the dot
                            var cleanStr = Regex.Replace(cashOnHandStr, "[^0-9.]", "");

                            // Convert to a number
                            if (float.TryParse(cleanStr, out float cashOnHand))
                            {
                                // Convert from billions to millions if needed
                                if (cashOnHandStr.Contains("B"))
                                {
                                    cashOnHand *= 1000;
                                }

                                cash_cash_equivalents = Math.Round(cashOnHand);

                                _viewController.UpdateProgressBar(50);//Update Status

                                await Total_Debt();
                            }
                            else
                            {
                                Console.WriteLine("Failed to convert the value to a number.");
                            }
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
                        double doubleValue;

                        // Check if the number ends with zeros
                        if (rawValue.EndsWith("000"))
                        {
                            doubleValue = double.Parse(rawValue) / 1000; // Convert the number to thousands
                        }
                        else
                        {
                            doubleValue = double.Parse(rawValue); // Assume the number is already in the correct format
                        }

                        total_debt = (int)doubleValue;

                        _viewController.UpdateProgressBar(75);//Update Status

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

                    // Check if the value ends with 'B' or 'M'
                    if (rawValue.EndsWith("B"))
                    {
                        doubleValue = double.Parse(rawValue.TrimEnd('B')) * 1000; // Convert billions to millions
                    }
                    else if (rawValue.EndsWith("M"))
                    {
                        doubleValue = double.Parse(rawValue.TrimEnd('M')); // Value is already in millions
                    }
                    else
                    {
                        doubleValue = double.Parse(rawValue); // Assume the number is in the correct format
                    }

                    shares_outstanding = (int)doubleValue;

                    _viewController.UpdateProgressBar(100);//Update Status

                }
                else
                {
                    Console.WriteLine("Div with 'Shares Outstanding' not found.");
                }
            }
        }
    }
}