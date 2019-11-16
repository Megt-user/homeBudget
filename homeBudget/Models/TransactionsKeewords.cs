using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace homeBudget.Models
{
    public class TransactionsKeewords
    {
        public int Id { get; set; }
        public string Description { get; set; }



        public int KeeWordId { get; set; }
        public int TransactionId { get; set; }

        public Transaction Transaction { get; set; }
        public KeeWord KeeWord { get; set; }


    }
}
