using System;
using System.IO;
using OfficeOpenXml;

namespace MetalCuttingOptimizer.Models
{
    public class CuttingResult
    {
        public int Id { get; set; }
        public int BilletId { get; set; }
        public SteelBillet? Billet { get; set; }
        public double WastePercentage { get; set; }
        public double UsefulArea { get; set; }
        public double WasteArea { get; set; }
        public DateTime CalculationDate { get; set; }
        public string? CuttingPattern { get; set; }
        public decimal TotalCost { get; set; }
        public int TotalPieces { get; set; }
        public double RequiredLength { get; set; }
        public double RequiredWidth { get; set; }

        public byte[] ExportToExcel(CuttingResult result)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Отчет по раскрою");

                // Ваш код для заполнения данных...

                // Сохраняем файл на диск для отладки
                var filePath = "C:\\temp\\cutting_optimization_debug.xlsx";
                File.WriteAllBytes(filePath, package.GetAsByteArray());
                Console.WriteLine($"Файл сохранен для отладки: {filePath}");

                return package.GetAsByteArray();
            }
        }
    }
} 