using System;
using System.Collections.Generic;
using System.Text;
using FluentAssertions;
using homeBudget.Models;
using homeBudget.Services;
using Xunit;

namespace homeBudget.Tests
{
   public class ModelOperationTests
    {
        [Fact]
        public void AverageforCategoryTestYearExtractions()
        {
            var jsonArray = TestsHelper.GetJonsArrayFromFile("TransactionsArray.json");
            List<AccountMovement> accountMovements = ModelConverter.GetAccountMovmentsFromJarray(jsonArray);
            accountMovements.Count.Should().Be(122);

            jsonArray = TestsHelper.GetJonsArrayFromFile("CategoriesArray.json");
            List<SubCategory> categorisModel = ModelConverter.GetCategoriesFromJarray(jsonArray);
            categorisModel.Count.Should().Be(105);

            var modementsViewModels = ModelConverter.CreateMovementsViewModels(accountMovements, categorisModel, "Felles");
            modementsViewModels[0].Category.Should().BeEquivalentTo("Altibox");

            string category = "Mat";
            int? year = null;
            int? month = null;
            bool justExtrations = true;

            var average = ModelOperation.AverageforCategory(modementsViewModels, category, 2017, month, justExtrations);

            average.Should().Be(466.8425);

        } 
        [Fact]
        public void AverageforCategoryTestYearIncams()
        {
            var jsonArray = TestsHelper.GetJonsArrayFromFile("TransactionsArray.json");
            List<AccountMovement> accountMovements = ModelConverter.GetAccountMovmentsFromJarray(jsonArray);
            accountMovements.Count.Should().Be(122);

            jsonArray = TestsHelper.GetJonsArrayFromFile("CategoriesArray.json");
            List<SubCategory> categorisModel = ModelConverter.GetCategoriesFromJarray(jsonArray);
            categorisModel.Count.Should().Be(105);

            var modementsViewModels = ModelConverter.CreateMovementsViewModels(accountMovements, categorisModel, "Felles");
            modementsViewModels[0].Category.Should().BeEquivalentTo("Altibox");

            string category = "Mat";
            int? year = null;
            int? month = null;
            bool justExtrations = false;

            var average = ModelOperation.AverageforCategory(modementsViewModels, category, 2017, month, justExtrations);

            average.Should().Be(117.5);

        }
        [Fact]
        public void AverageforCategoryTestMonth()
        {
            var jsonArray = TestsHelper.GetJonsArrayFromFile("TransactionsArray.json");
            List<AccountMovement> accountMovements = ModelConverter.GetAccountMovmentsFromJarray(jsonArray);
            accountMovements.Count.Should().Be(122);

            jsonArray = TestsHelper.GetJonsArrayFromFile("CategoriesArray.json");
            List<SubCategory> categorisModel = ModelConverter.GetCategoriesFromJarray(jsonArray);
            categorisModel.Count.Should().Be(105);

            var modementsViewModels = ModelConverter.CreateMovementsViewModels(accountMovements, categorisModel, "Felles");
            modementsViewModels[0].Category.Should().BeEquivalentTo("Altibox");

            var average = ModelOperation.AverageforCategory(modementsViewModels, "Mat", null, 6, true);

            average.Should().Be(117.5);

        }
    }
}
