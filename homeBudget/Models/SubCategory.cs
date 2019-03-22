using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace homeBudget.Models
{
    public class SubCategory
    {
        public string KeyWord { get; set; }
        public string Category { get; set; }
        public string SupPorject { get; set; }
        public List<AccountMovement> accountMovements { get; set; }
    }
}
