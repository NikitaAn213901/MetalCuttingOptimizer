using System;

namespace MetalCuttingOptimizer.Models
{
    public class SteelBillet
    {
        public int Id { get; set; }
        public double Length { get; set; }
        public double Width { get; set; }
        public double Thickness { get; set; }
        public double Weight { get; set; }
        public string? Grade { get; set; }
        public decimal CostPerUnit { get; set; }
    }
} 