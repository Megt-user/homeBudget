using System.Collections.Generic;

namespace homeBudget.Models
{
    public class SubCategory
    {
        public int Id { get; set; }
        public string KeyWord { get; set; }
        public string Category { get; set; }
        public string SubPorject { get; set; }

        public ICollection<KeeWord> KeeWord { get; set; }

    }
}
