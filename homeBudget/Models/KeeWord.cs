using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace homeBudget.Models
{
    public class KeeWord
    {
        public int Id { get; set; }
        [Required]
        public string Name { get; set; }
        public double XCoordinate { get; set; }
        public double YCoordinate { get; set; }
        public SubCategory SubCategory { get; set; }

        public ICollection<TransactionsKeewords> TransactionsKeewordses { get; set; }

        //public IEnumerable<Transaction> Transactions
    }
}
