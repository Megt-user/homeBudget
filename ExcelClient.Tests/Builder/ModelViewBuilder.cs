using System;
using System.Collections.Generic;
using System.Text;
using Transactions.Models;

namespace ExcelClient.Tests.Builder
{
    public class ModelViewBuilder
    {
        private readonly MovementsViewModel _model;
        public ModelViewBuilder()
        {
            _model = new MovementsViewModel();
        }

        public MovementsViewModel Build()
        {
            return _model;
        }

        public List<MovementsViewModel> AddTextToMovemnt(List<string> texts)
        {
            List<MovementsViewModel> newList = new List<MovementsViewModel>();
            foreach (var text in texts)
            {
                newList.Add(new MovementsViewModel() { Text = text });
            }

            return newList;
        }
    }
}
