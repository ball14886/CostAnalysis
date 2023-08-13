using Cost_Analysis.Models;
using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media.Media3D;
using Excel = Microsoft.Office.Interop.Excel;

namespace Cost_Analysis
{
    public partial class MainWindow : System.Windows.Window
    {
        List<InputWeightModel> waybillWeights = new List<InputWeightModel>();
        List<InputModel> waybillCosts = new List<InputModel>();

        public MainWindow()
        {
            InitializeComponent();
        }

        private void ImportWeight_Button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();

            dlg.DefaultExt = ".xls|.xlsx";
            dlg.Filter = "(.xls)|*.xls|(.xlsx)|*.xlsx";

            bool? result = dlg.ShowDialog();

            if (result == true)
            {
                if (dlg.FileName.Length > 0)
                {
                    Excel.Application excelApp = new Excel.Application();
                    Workbook excelBook = excelApp.Workbooks.Open(dlg.FileName, 0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                    Worksheet excelSheet = (Worksheet)excelBook.Worksheets.get_Item(1); ;
                    Excel.Range excelRange = excelSheet.UsedRange;

                    if (excelSheet != null)
                    {
                        Excel.Range range = excelSheet.UsedRange;
                        int successRowCount = 0;

                        for (int i = 2; i <= excelRange.Rows.Count; i++)
                        {
                            InputWeightModel wayBillWeight = new InputWeightModel();
                            var wayBillNumberRaw = range.Cells[i, 1] as Excel.Range;
                            if (wayBillNumberRaw is null || wayBillNumberRaw.Value2 is null)
                            {
                                break;
                            }
                            string wayBillNumber = wayBillNumberRaw.Value2.ToString();
                            wayBillWeight.WayBillNumber = wayBillNumber;

                            var realWeightRaw = range.Cells[i, 2] as Excel.Range;
                            if (realWeightRaw is null || realWeightRaw.Value2 is null)
                            {
                                continue;
                            }
                            string realWeight = realWeightRaw.Value2.ToString();
                            wayBillWeight.RealWeight = Convert.ToDecimal(realWeight);

                            var dimensionWeightRaw = range.Cells[i, 3] as Excel.Range;
                            if (dimensionWeightRaw is null || dimensionWeightRaw.Value2 is null)
                            {
                                continue;
                            }
                            string dimensionWeight = dimensionWeightRaw.Value2.ToString();
                            wayBillWeight.DimensionWeight = Convert.ToDecimal(dimensionWeight);

                            var costedWeightRaw = range.Cells[i, 4] as Excel.Range;
                            if (costedWeightRaw is null || costedWeightRaw.Value2 is null)
                            {
                                continue;
                            }
                            string costedWeight = costedWeightRaw.Value2.ToString();
                            wayBillWeight.CostedWeight = Convert.ToDecimal(costedWeight);

                            var recipientAddressRaw = range.Cells[i, 5] as Excel.Range;
                            if (recipientAddressRaw is null || recipientAddressRaw.Value2 is null)
                            {
                                break;
                            }
                            string recipientAddress = recipientAddressRaw.Value2.ToString();
                            wayBillWeight.RecipientAddress = recipientAddress;

                            waybillWeights.Add(wayBillWeight);
                            successRowCount++;
                        }
                        StatusTextBox.Text += $"Weight {successRowCount} rows added\r\n";
                        ImportWeightCount_TextBox.Text = $"{waybillWeights.Count} records";
                    }
                    else
                    {
                        MessageBox.Show("Incorrect Weight file format. Correct format is [Waybill, Weight]",
                                        "Incorrect File Format",
                                        MessageBoxButton.OK,
                                        MessageBoxImage.Warning);
                    }

                    excelBook.Close(true, null, null);
                    excelApp.Quit();
                }
            }
        }

        private void ImportCost_Button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();

            dlg.DefaultExt = ".xls|.xlsx";
            dlg.Filter = "(.xls)|*.xls|(.xlsx)|*.xlsx";

            bool? result = dlg.ShowDialog();

            if (result == true)
            {
                if (dlg.FileName.Length > 0)
                {
                    Excel.Application excelApp = new Excel.Application();
                    Workbook excelBook = excelApp.Workbooks.Open(dlg.FileName, 0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                    Worksheet excelSheet = (Worksheet)excelBook.Worksheets.get_Item(1);
                    Excel.Range excelRange = excelSheet.UsedRange;

                    if (excelSheet != null)
                    {
                        Excel.Range range = excelSheet.UsedRange;
                        int successRowCount = 0;

                        for (int i = 2; i <= excelRange.Rows.Count; i++)
                        {
                            var wayBillNumberRaw = range.Cells[i, 1] as Excel.Range;
                            if (wayBillNumberRaw is null || wayBillNumberRaw.Value2 is null)
                            {
                                break;
                            }
                            string wayBillNumber = wayBillNumberRaw.Value2.ToString();

                            var costRaw = range.Cells[i, 2] as Excel.Range;
                            if (costRaw is null || costRaw.Value2 is null)
                            {
                                continue;
                            }
                            string cost = costRaw.Value2.ToString();

                            InputModel wayBillCost = new InputModel();
                            wayBillCost.WayBillNumber = wayBillNumber;
                            wayBillCost.Value = Convert.ToDecimal(cost);

                            waybillCosts.Add(wayBillCost);
                            successRowCount++;
                        }
                        StatusTextBox.Text += $"Cost {successRowCount} rows added\r\n";
                        ImportCostCount_TextBox.Text = $"{waybillCosts.Count} records";
                    }
                    else
                    {
                        MessageBox.Show("Incorrect Cost file format. Correct format is [Waybill, Cost]",
                                        "Incorrect File Format",
                                        MessageBoxButton.OK,
                                        MessageBoxImage.Warning);
                    }

                    excelBook.Close(true, null, null);
                    excelApp.Quit();
                }
            }
        }

        private void ResetButton_Click(object sender, RoutedEventArgs e)
        {
            StatusTextBox.Text += $"Cost {waybillCosts.Count()} rows cleared\r\n";
            StatusTextBox.Text += $"Weight {waybillWeights.Count()} rows cleared\r\n";
            waybillWeights.Clear();
            waybillCosts.Clear();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var costList = GenerateCostList();

            var merged = (from weight in waybillWeights
                          join cost in waybillCosts on weight.WayBillNumber equals cost.WayBillNumber
                          join expectedCost in costList on weight.ExpectedCalculateCostWeight equals expectedCost.Weight
                          select new WaybillIncomeCostWeightModel
                          {
                              WayBillNumber = weight.WayBillNumber,
                              RealWeight = weight.RealWeight,
                              DimensionWeight = weight.DimensionWeight,
                              CostedWeight = weight.CostedWeight,
                              Cost = (cost.Value * -1),
                              CostBKK = expectedCost.CostBKK,
                              CostUpcountry = weight.IsBKK ? 0 : expectedCost.CostUpcountry,
                              RecipientAddress = weight.RecipientAddress
                          }).ToList();

            Export(merged);
        }

        private async void Export(List<WaybillIncomeCostWeightModel> items)
        {
            var exportInformation = items.Where(x => x.CostBKK != x.ExpectedCostBkk
                                                     && x.CostUpcountry != x.ExpectedCostUpcountry
                                                     && x.Cost != x.ExpectedCostBkk
                                                     && x.Cost != x.ExpectedCostUpcountry)
                                         .Select(x => new
                                                      {
                                                          x.WayBillNumber,
                                                          x.RealWeight,
                                                          x.DimensionWeight,
                                                          x.CostedWeight,
                                                          x.Cost,
                                                          x.ExpectedCostBkk,
                                                          //x.CostBKK,
                                                          x.ExpectedCostUpcountry,
                                                          //x.CostUpcountry,
                                                          x.DifferenceBKK,
                                                          x.DifferenceUpcountry,
                                                          x.RecipientAddress
                                                      })
                                         .OrderBy(x => x.WayBillNumber)
                                         .ToList();

            var contentBytes = await GenerateCSVFile(exportInformation, true);

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.DefaultExt = "csv";
            saveFileDialog.FileName = $"CostAnalysis - {DateTime.Now.ToString("yyyyMMddHHmmss")}";
            if (saveFileDialog.ShowDialog() == true)
            {
                File.WriteAllBytes(saveFileDialog.FileName, contentBytes);
                MessageBox.Show("Save file completed",
                                "Save file",
                                MessageBoxButton.OK,
                                MessageBoxImage.None);
            }
        }

        public static async Task<byte[]> GenerateCSVFile<T>(List<T> datas, bool hasHeader)
        {
            Type dataType = typeof(T);
            var props = dataType.GetProperties(BindingFlags.Public | BindingFlags.Instance);
            MemoryStream stream = new MemoryStream();

            using (var writer = new StreamWriter(stream, Encoding.UTF8))
            {
                var builder = new StringBuilder();
                if (hasHeader)
                {
                    builder.Append(string.Join(",", props.Select(p => p.Name)));
                    builder.Append(Environment.NewLine);
                }

                foreach (var data in datas)
                {
                    builder.Append(string.Join(",", props.Select(p => p.GetValue(data, null))));
                    builder.Append(Environment.NewLine);
                }

                await writer.WriteAsync(builder.ToString());
                await writer.FlushAsync();
                stream.Position = 0;

                return stream.ToArray();
            }
        }

        private List<WeightCost> GenerateCostList()
        {
            var costList = new List<WeightCost>
            {
                new WeightCost(0.50m, 15, 18),
                new WeightCost(1, 25, 31),
                new WeightCost(2, 30, 34),
                new WeightCost(3, 35, 37),
                new WeightCost(4, 45, 40),
                new WeightCost(5, 50, 43),
                new WeightCost(6, 60, 77),
                new WeightCost(7, 70, 89),
                new WeightCost(8, 80, 101),
                new WeightCost(9, 90, 113),
                new WeightCost(10, 100, 125),
                new WeightCost(11, 110, 137),
                new WeightCost(12, 120, 149),
                new WeightCost(13, 130, 161),
                new WeightCost(14, 140, 173),
                new WeightCost(15, 150, 185),
                new WeightCost(16, 170, 210),
                new WeightCost(17, 190, 235),
                new WeightCost(18, 210, 260),
                new WeightCost(19, 230, 285),
                new WeightCost(20, 250, 310),
                new WeightCost(21, 270, 310),
                new WeightCost(22, 280, 310),
                new WeightCost(23, 290, 310),
                new WeightCost(24, 300, 310),
                new WeightCost(25, 310, 310),
                new WeightCost(26, 320, 310),
                new WeightCost(27, 330, 310),
                new WeightCost(28, 340, 310),
                new WeightCost(29, 350, 310),
                new WeightCost(30, 360, 310),
                new WeightCost(31, 370, 310),
                new WeightCost(32, 380, 322),
                new WeightCost(33, 390, 334),
                new WeightCost(34, 400, 346),
                new WeightCost(35, 410, 358),
                new WeightCost(36, 420, 370),
                new WeightCost(37, 430, 382),
                new WeightCost(38, 440, 394),
                new WeightCost(39, 450, 406),
                new WeightCost(40, 460, 418),
                new WeightCost(41, 470, 430),
                new WeightCost(42, 480, 442),
                new WeightCost(43, 490, 454),
                new WeightCost(44, 500, 466),
                new WeightCost(45, 510, 478),
                new WeightCost(46, 520, 490),
                new WeightCost(47, 530, 502),
                new WeightCost(48, 540, 514),
                new WeightCost(49, 550, 526),
                new WeightCost(50, 560, 538)
            };

            return costList;
        }
    }
}