using System.Collections.Generic;

namespace homeBudget.Models
{
    public class SubCategory
    {
        public string KeyWord { get; set; }
        public string Category { get; set; }
        public string SubPorject { get; set; }
        public List<AccountMovement> accountMovements { get; set; }
    }
}
