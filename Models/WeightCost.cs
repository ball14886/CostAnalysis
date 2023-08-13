namespace Cost_Analysis.Models
{
    public class WeightCost
    {
        public decimal Weight { get; set; }
        public int CostBKK { get; set; }
        public int CostUpcountry { get; set; }

        public WeightCost(decimal weight, int costBKK, int costUpcountry)
        {
            Weight = weight;
            CostBKK = costBKK;
            CostUpcountry = costUpcountry;
        }
    }
}