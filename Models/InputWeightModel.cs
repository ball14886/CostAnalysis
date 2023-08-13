using System;

namespace Cost_Analysis.Models
{
    public class InputWeightModel
    {
        public string? WayBillNumber { get; set; }
        public decimal RealWeight { get; set; }
        public decimal DimensionWeight { get; set; }
        public decimal ExpectedCalculateCostWeight
        {
            get
            {
                if (RealWeight <= (decimal)0.5)
                {
                    return (decimal)0.5;
                }
                return Math.Ceiling(Math.Max(RealWeight, DimensionWeight));
            }
        }
        public decimal CostedWeight { get; set; }
        public string? RecipientAddress { get; set; }
        public bool IsBKK
        {
            get
            {
                return RecipientAddress.Contains("กรุงเทพ") || RecipientAddress.Contains("สมุทรปราการ");
            }
        }
    }
}