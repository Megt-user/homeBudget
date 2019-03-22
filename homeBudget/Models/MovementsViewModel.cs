using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace homeBudget.Models
{
    public class MovementsViewModel
    {
        public int Id { get; set; }
        public DateTime DateTime { get; set; }
        public string Text { get; set; }
        public double Amount { get; set; }
        public string AcountName { get; set; }
        public string KeyWord { get; set; }
        public string Category { get; set; }
        public string SupPorject { get; set; }

        public static List<string> excelColumns { get; set; }

        public MovementsViewModel()
        {
            excelColumns = new List<string>() {"Id", "DateTime", "Text","Amount","KeyWord", "Category","SupPorject", "AcountName" };
        }

    }
}
