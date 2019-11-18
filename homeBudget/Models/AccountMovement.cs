using System;

namespace homeBudget.Models
{
    public class AccountMovement
    {
        public int Id { get; set; }
        public DateTime DateTime { get; set; }
        public string Text { get; set; }
        public double Amount { get; set; }
        public string AcountName { get; set; }
        public SubCategory SubCategory { get; set; }
    }
}
