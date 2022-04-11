using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using System.IO;
using System;
using System.Data;

#nullable disable
namespace Serializable.Classes
{
    public delegate void SaveSucceseful(string mess);
    public delegate void ErrorHandler(int code, string msg);
    public class Excel
    {
        public List<List<string>> Table = new();

        public event SaveSucceseful SaveSucces;
        public event ErrorHandler ErrHandle;
        // copywrite with Github (because with my VS is doesn`t work). Creator ukushu. Original: https://github.com/ukushu/DataExporter
        // 
        public void FileOpen(string path)
        {

            var workbook = new XLWorkbook(path);
            var ws1 = workbook.Worksheet(1);
            try
            {
                foreach (var xlRow in ws1.RangeUsed().Rows())
                {
                    Table.Add(new List<string>());

                    foreach (var xlCell in xlRow.Cells())
                    {
                        var formula = xlCell.FormulaA1;
                        var value = xlCell.Value.ToString();

                        string targetCellValue = (formula.Length == 0) ? value : "=" + formula;

                        Table[^1].Add(targetCellValue);
                    }
                }
            }
            catch (Exception e)
            {
                ErrHandle(1, e.Message);
            }

        }

        public void FileSave(string path)
        {
            CreateDirIfNotExist(path, true);

            using (XLWorkbook wb = new())
            {
                var workSheet = wb.Worksheets.Add("Sample Sheet");

                for (int row = 0; row < Table.Count; row++)
                {
                    for (int col = 0; col < Table[row].Count; col++)
                    {
                        var cellAdress = GetExcelPos(row, col);

                        if (Table[row][col].StartsWith("="))
                        {
                            workSheet.Cell(cellAdress).FormulaA1 = Table[row][col];
                        }
                        else
                        {
                            workSheet.Cell(cellAdress).Value = Table[row][col];
                        }
                    }
                }

                wb.SaveAs(path);
                SaveSucces("Сохранение выполнено успешно");
            }
        }

        public void FileSave(string path, List<List<string>> otherData)
        {
            Table = otherData;

            CreateDirIfNotExist(path, true);

            using (XLWorkbook wb = new())
            {
                var workSheet = wb.Worksheets.Add("Sample Sheet");

                for (int row = 0; row < Table.Count; row++)
                {
                    for (int col = 0; col < Table[row].Count; col++)
                    {
                        var cellAdress = GetExcelPos(row, col);

                        if (Table[row][col].StartsWith("="))
                        {
                            workSheet.Cell(cellAdress).FormulaA1 = Table[row][col];
                        }
                        else
                        {
                            workSheet.Cell(cellAdress).Value = Table[row][col];
                        }
                    }
                }

                wb.SaveAs(path);
                SaveSucces("Сохранение выполнено успешно");
            }
        }

        public void AddRow(params string[] cells)
        {
            Table.Add(cells.ToList());
        }

        public static string GetExcelPos(int row, int cell)
        {
            char[] alph = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray();

            int count = cell / 26;
            string alphResult;

            alphResult = count > 0 ? alph[count] + alph[count % 26].ToString() : alph[cell].ToString();

            return alphResult + (row + 1);
        }

        private void CreateDirIfNotExist(string dirPath, bool removeFilename = false)
        {
            if (removeFilename)
            {
                dirPath = Directory.GetParent(dirPath).FullName;
            }

            if (!Directory.Exists(dirPath))
            {
                _ = Directory.CreateDirectory(dirPath);
            }
        }

        public static DataTable ToDataTable(List<List<string>> matrix)
        {
            DataTable res = new();

            for (int i = 0; i < matrix.Count; i++)
            {
                _ = res.Columns.Add($"{i + 1}");
            }

            for (int i = 0; i < matrix.Count; i++)
            {
                DataRow row = res.NewRow();

                for (int j = 0; j < matrix[i].Count; j++)
                {
                    row[j] = matrix[i][j];
                }

                res.Rows.Add(row);
            }

            return res;
        }
    }
}
