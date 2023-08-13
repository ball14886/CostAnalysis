namespace Cost_Analysis.Models
{
    public class WaybillIncomeCostWeightModel
    {
        public string? WayBillNumber { get; set; }
        public decimal? RealWeight { get; set; }
        public decimal? DimensionWeight { get; set; }
        public decimal? CostedWeight { get; set; }
        public decimal Cost { get; set; }

        public decimal ExpectedCostBkk
        {
            get
            {
                return CostBKK * (decimal)0.7;
            }
        }

        public decimal CostBKK { get; set; }

        public decimal ExpectedCostUpcountry
        {
            get
            {
                return CostUpcountry * (decimal)0.7;
            }
        }

        public decimal CostUpcountry { get; set; }

        public decimal DifferenceBKK
        {
            get
            {
                return ExpectedCostBkk - Cost;
            }
        }

        public decimal DifferenceUpcountry
        {
            get
            {
                return ExpectedCostUpcountry - Cost;
            }
        }

        public string? RecipientAddress { get; set; }
    }
}