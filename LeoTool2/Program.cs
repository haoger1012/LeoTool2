using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LeoTool2
{
    class Program
    {
        static void Main(string[] args)
        {
            string fileName = "input.xlsx";
            string extension = Path.GetExtension(fileName);
            if (extension == ".xlsx" || extension == ".xls")
            {
                if (!Directory.Exists("Result"))
                {
                    Directory.CreateDirectory("Result");
                }

                var inputs = new List<Input>();
                IWorkbook wbInput;
                using (var fs = new FileStream(fileName, FileMode.Open, FileAccess.Read))
                {
                    if (extension == ".xlsx")
                    {
                        wbInput = new XSSFWorkbook(fs);
                    }
                    else
                    {
                        wbInput = new HSSFWorkbook(fs);
                    }
                }

                // get data
                var sheetInput = wbInput.GetSheetAt(0);
                for (int i = sheetInput.FirstRowNum + 1; i <= sheetInput.LastRowNum; i++)
                {
                    var row = sheetInput.GetRow(i);
                    if (row.Cells.All(x => x.CellType == CellType.Blank)) continue;
                    inputs.Add(new Input
                    {
                        Company = row.GetCell(0).StringCellValue,
                        Contact = row.GetCell(1).StringCellValue,
                        Tel = row.GetCell(2).NumericCellValue,
                        TaxId = row.GetCell(3).NumericCellValue,
                        Address = row.GetCell(4).StringCellValue,
                        No = row.GetCell(5).NumericCellValue,
                        Product = row.GetCell(6).StringCellValue,
                        Quantity = row.GetCell(7).NumericCellValue,
                        Price = row.GetCell(8).NumericCellValue,
                    });
                }

                // group data
                var outputs = from i in inputs
                              group i by new { i.Company, i.Contact, i.Tel, i.TaxId, i.Address } into grp
                              select new
                              {
                                  grp.Key.Company,
                                  grp.Key.Contact,
                                  grp.Key.Tel,
                                  grp.Key.TaxId,
                                  grp.Key.Address,
                                  Details = (from g in grp
                                             select new
                                             {
                                                 g.No,
                                                 g.Product,
                                                 g.Quantity,
                                                 g.Price
                                             }).ToList()
                              };

                // fill data
                foreach (var item in outputs)
                {
                    IWorkbook wbOutput;
                    using (var fs = new FileStream("Template.xlsx", FileMode.Open, FileAccess.Read))
                    {
                        wbOutput = new XSSFWorkbook(fs);
                    }

                    var sheetOutput = wbOutput.GetSheetAt(0);
                    sheetOutput.ForceFormulaRecalculation = true;

                    sheetOutput.GetRow(12).GetCell(2).SetCellValue(item.Company);
                    sheetOutput.GetRow(13).GetCell(2).SetCellValue(item.Contact);
                    sheetOutput.GetRow(14).GetCell(2).SetCellValue(item.Tel);
                    sheetOutput.GetRow(16).GetCell(2).SetCellValue(item.TaxId);
                    sheetOutput.GetRow(17).GetCell(2).SetCellValue(item.Address);

                    int sourceRowNum = 21;
                    var sourceRow = sheetOutput.GetRow(sourceRowNum);
                    int detailsCount = item.Details.Count();

                    for (int i = 0; i < detailsCount; i++)
                    {
                        IRow row;
                        if (i == 0)
                        {
                            row = sheetOutput.GetRow(i + sourceRowNum);
                        }
                        else
                        {
                            // copy row
                            sheetOutput.ShiftRows(i + sourceRowNum, sheetOutput.LastRowNum, 1);                            
                            row = sheetOutput.CreateRow(i + sourceRowNum);
                            for (int j = sourceRow.FirstCellNum; j < sourceRow.LastCellNum; j++)
                            {
                                var oldCell = sourceRow.GetCell(j);
                                var newCell = row.CreateCell(j);

                                newCell.CellStyle = oldCell.CellStyle;
                                newCell.SetCellType(oldCell.CellType);
                            }
                        }

                        row.GetCell(1).SetCellValue(item.Details[i].No);
                        row.GetCell(2).SetCellValue(item.Details[i].Product);
                        row.GetCell(3).SetCellValue(item.Details[i].Quantity);
                        row.GetCell(4).SetCellValue(item.Details[i].Price);
                        // 單項總價
                        row.GetCell(5).SetCellFormula($"D{i + sourceRowNum + 1}*E{i + sourceRowNum + 1}");
                    }

                    // 小計
                    sheetOutput.GetRow(sourceRowNum + detailsCount).GetCell(5)
                        .SetCellFormula($"SUM(F{sourceRowNum + 1}:F{sourceRowNum + detailsCount})");
                    // 稅金
                    sheetOutput.GetRow(sourceRowNum + detailsCount + 2).GetCell(5)
                        .SetCellFormula($"F{sourceRowNum + detailsCount + 1}*F{sourceRowNum + detailsCount + 2}");
                    // 總計(含稅)
                    sheetOutput.GetRow(sourceRowNum + detailsCount + 3).GetCell(5)
                        .SetCellFormula($"F{sourceRowNum + detailsCount + 1}+F{sourceRowNum + detailsCount + 3}");

                    using (var fs = new FileStream($"Result/{item.Company}.xlsx", FileMode.Create, FileAccess.Write))
                    {
                        wbOutput.Write(fs);
                    }

                    Process.Start("Result");
                }
            }
            else
            {
                Console.WriteLine("File invalid");
            }
        }
    }

    public class Input
    {
        public string Company { get; set; }
        public string Contact { get; set; }
        public double Tel { get; set; }
        public double TaxId { get; set; }
        public string Address { get; set; }

        public double No { get; set; }
        public string Product { get; set; }
        public double Quantity { get; set; }
        public double Price { get; set; }
    }
}
