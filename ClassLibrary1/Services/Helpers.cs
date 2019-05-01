using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Transactions.Services
{
    public class Helpers
    {
        public static IEnumerable<string> AddItemsToIenumeration(IEnumerable<string> Inumerables, List<string> items)
        {
            List<string> temp = Inumerables.ToList();
            foreach (var item in items)
            {
                temp.Add(item);
            }

            return temp;
        }
        public static IEnumerable<string> DeleteItemsfromIenumeration(IEnumerable<string> Inumerables, List<string> items)
        {
            List<string> temp = Inumerables.ToList();
            foreach (var item in items)
            {
                var itemToRemove = temp.Single(i => i == item);
                if (itemToRemove != null)
                    temp.Remove(item);
            }

            return temp;
        }

    }
}
