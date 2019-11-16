using System;
using System.Collections.Generic;

namespace homeBudget.Models
{
    public class Transaction
    {
        public int Id { get; set; }
        public DateTime DateTime { get; set; }
        public string Text { get; set; }
        public double Amount { get; set; }

        public BankAccount BankAccount { get; set; }
        public KeeWord KeeWord { get; set; }
    }
}
