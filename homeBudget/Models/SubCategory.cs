using System.Collections.Generic;

namespace homeBudget.Models
{
    public class SubCategory
    {
        public int Id { get; set; }
        public string KeyWord { get; set; }
        public string Category { get; set; }
        public string SupPorject { get; set; }
        
        public ICollection<Transaction> Transactions { get; set; }

    }
}
