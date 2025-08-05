using OfficeOpenXml;
using System;
using System.IO;

namespace Demo;

class Program
{
    static void Main(string[] args)
    {
        var filePath = Path.Combine(  "EPPlus4-demo.xlsx");
        var package = new ExcelPackage();
        var diffWorkBook = package.Workbook;
        var summarySheet = diffWorkBook.Worksheets.Add("summary");

        try
        {
            // 填充10行10列的数据
            for (int row = 1; row <= 10; row++)
            {
                for (int col = 1; col <= 10; col++)
                {
                    // 可以填入不同的数据，这里以行列坐标为例
                    summarySheet.Cells[row, col].Value = $"Row{row}-Col{col}";
                }
            }

            diffWorkBook.View.ActiveTab = 0; // 设置默认表格为第一个sheet

            // 保存文件
            package.SaveAs(filePath);

            Console.WriteLine($"文件已保存到: {filePath}");
        }
        catch (Exception e)
        {
            Console.WriteLine($"保存文件时出错: {e.Message}");
        }
        finally
        {
            package.Dispose();
        }
    }
}