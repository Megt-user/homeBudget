﻿using System;
using System.Collections.Generic;

namespace homeBudget.Models
{
    public class TransactionViewModel
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

        public TransactionViewModel()
        {
            excelColumns = new List<string>() {"Id", "DateTime", "Text","Amount","KeyWord", "Category","SubPorject", "AcountName" };
        }

    }
}
