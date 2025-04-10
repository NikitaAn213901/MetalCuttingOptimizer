using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using MetalCuttingOptimizer.Models;

namespace MetalCuttingOptimizer.Services
{
    public class CuttingOptimizationService
    {
        private readonly List<CuttingResult> _results = new List<CuttingResult>();

        public CuttingResult OptimizeCutting(SteelBillet billet, double requiredLength, double requiredWidth)
        {
            // Простой алгоритм оптимизации
            double totalArea = billet.Length * billet.Width;
            double pieceArea = requiredLength * requiredWidth;
            
            int piecesLengthwise = (int)(billet.Length / requiredLength);
            int piecesWidthwise = (int)(billet.Width / requiredWidth);
            
            int totalPieces = piecesLengthwise * piecesWidthwise;
            double usefulArea = totalPieces * pieceArea;
            double wasteArea = totalArea - usefulArea;
            
            var result = new CuttingResult
            {
                BilletId = billet.Id,
                Billet = billet,
                WastePercentage = (wasteArea / totalArea) * 100,
                UsefulArea = usefulArea,
                WasteArea = wasteArea,
                CalculationDate = DateTime.Now,
                CuttingPattern = $"{piecesLengthwise}x{piecesWidthwise}",
                TotalCost = billet.CostPerUnit,
                TotalPieces = totalPieces,
                RequiredLength = requiredLength,
                RequiredWidth = requiredWidth
            };

            _results.Add(result);
            return result;
        }

       
        public byte[] ExportToExcel(CuttingResult result)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Отчет по раскрою");

                // Логирование начала создания файла
                Console.WriteLine("Начало создания файла Excel");

                // Стили для заголовков
                var headerStyle = worksheet.Cells[1, 1].Style;
                headerStyle.Font.Bold = true;
                headerStyle.Font.Size = 14;
                headerStyle.Fill.PatternType = ExcelFillStyle.Solid;
                headerStyle.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);

                // Стили для подзаголовков
                var subHeaderStyle = worksheet.Cells[3, 1].Style;
                subHeaderStyle.Font.Bold = true;
                subHeaderStyle.Font.Size = 12;

                // Заголовок отчета
                worksheet.Cells[1, 1].Value = "ОТЧЕТ ПО ОПТИМИЗАЦИИ РАСКРОЯ СТАЛЬНЫХ ЗАГОТОВОК";
                worksheet.Cells[1, 1, 1, 4].Merge = true;
                worksheet.Cells[1, 1, 1, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                // Дата отчета
                worksheet.Cells[2, 1].Value = $"Дата расчета: {result.CalculationDate:dd.MM.yyyy HH:mm:ss}";
                worksheet.Cells[2, 1, 2, 4].Merge = true;

                // Исходные данные заготовки
                worksheet.Cells[4, 1].Value = "1. ПАРАМЕТРЫ ИСХОДНОЙ ЗАГОТОВКИ";
                worksheet.Cells[4, 1, 4, 4].Merge = true;
                
                worksheet.Cells[5, 1].Value = "Длина заготовки:";
                worksheet.Cells[5, 2].Value = result.Billet?.Length ?? 0;
                worksheet.Cells[5, 3].Value = "мм";
                
                worksheet.Cells[6, 1].Value = "Ширина заготовки:";
                worksheet.Cells[6, 2].Value = result.Billet?.Width ?? 0;
                worksheet.Cells[6, 3].Value = "мм";
                
                worksheet.Cells[7, 1].Value = "Толщина заготовки:";
                worksheet.Cells[7, 2].Value = result.Billet?.Thickness ?? 0;
                worksheet.Cells[7, 3].Value = "мм";
                
                worksheet.Cells[8, 1].Value = "Марка стали:";
                worksheet.Cells[8, 2].Value = result.Billet?.Grade ?? "Неизвестно";
                
                worksheet.Cells[9, 1].Value = "Стоимость заготовки:";
                worksheet.Cells[9, 2].Value = result.Billet?.CostPerUnit ?? 0;
                worksheet.Cells[9, 3].Value = "руб.";

                // Требуемые размеры деталей
                worksheet.Cells[11, 1].Value = "2. ТРЕБУЕМЫЕ РАЗМЕРЫ ДЕТАЛЕЙ";
                worksheet.Cells[11, 1, 11, 4].Merge = true;
                
                worksheet.Cells[12, 1].Value = "Длина детали:";
                worksheet.Cells[12, 2].Value = result.RequiredLength;
                worksheet.Cells[12, 3].Value = "мм";
                
                worksheet.Cells[13, 1].Value = "Ширина детали:";
                worksheet.Cells[13, 2].Value = result.RequiredWidth;
                worksheet.Cells[13, 3].Value = "мм";

                // Результаты оптимизации
                worksheet.Cells[15, 1].Value = "3. РЕЗУЛЬТАТЫ ОПТИМИЗАЦИИ";
                worksheet.Cells[15, 1, 15, 4].Merge = true;
                
                worksheet.Cells[16, 1].Value = "Схема раскроя:";
                worksheet.Cells[16, 2].Value = result.CuttingPattern;
                
                worksheet.Cells[17, 1].Value = "Количество деталей:";
                worksheet.Cells[17, 2].Value = result.TotalPieces;
                worksheet.Cells[17, 3].Value = "шт.";
                
                worksheet.Cells[18, 1].Value = "Полезная площадь:";
                worksheet.Cells[18, 2].Value = Math.Round(result.UsefulArea, 2);
                worksheet.Cells[18, 3].Value = "мм²";
                
                worksheet.Cells[19, 1].Value = "Площадь отходов:";
                worksheet.Cells[19, 2].Value = Math.Round(result.WasteArea, 2);
                worksheet.Cells[19, 3].Value = "мм²";
                
                worksheet.Cells[20, 1].Value = "Процент отходов:";
                worksheet.Cells[20, 2].Value = Math.Round(result.WastePercentage, 2);
                worksheet.Cells[20, 3].Value = "%";

                // Экономические показатели
                worksheet.Cells[22, 1].Value = "4. ЭКОНОМИЧЕСКИЕ ПОКАЗАТЕЛИ";
                worksheet.Cells[22, 1, 22, 4].Merge = true;
                
                worksheet.Cells[23, 1].Value = "Стоимость одной детали:";
                worksheet.Cells[23, 2].Value = Math.Round((double)result.TotalCost / result.TotalPieces, 2);
                worksheet.Cells[23, 3].Value = "руб.";
                
                worksheet.Cells[24, 1].Value = "Общая стоимость отходов:";
                worksheet.Cells[24, 2].Value = Math.Round((double)result.TotalCost * (result.WastePercentage / 100), 2);
                worksheet.Cells[24, 3].Value = "руб.";

                // Форматирование
                var range = worksheet.Cells[1, 1, 24, 4];
                range.AutoFitColumns();
                
                // Добавляем границы для всех заполненных ячеек
                var borderRange = worksheet.Cells[1, 1, 24, 3];
                borderRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                borderRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                borderRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                borderRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                // Выравнивание
                worksheet.Cells[1, 1, 24, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells[1, 2, 24, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                worksheet.Cells[1, 3, 24, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                // Применяем стили к заголовкам разделов
                var sectionHeaders = new[] { 4, 11, 15, 22 };
                foreach (var row in sectionHeaders)
                {
                    var headerRange = worksheet.Cells[row, 1, row, 4];
                    headerRange.Style.Font.Bold = true;
                    headerRange.Style.Font.Size = 12;
                    headerRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    headerRange.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                }

                // Логирование завершения создания файла
                Console.WriteLine("Файл Excel создан успешно");

                return package.GetAsByteArray();
            }
        }

        internal CuttingResult GetCuttingResultById(int billetId)
        {
            throw new NotImplementedException();
        }
    }
} 