using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FluentAssertions;
using homeBudget.Models;
using homeBudget.Services;
using Xunit;

namespace homeBudget.Tests
{
    public class ModelConverterTests
    {
        [Fact]
        public void GetJonsArrayFromTransactionFile()
        {
            var jsonArray = TestsHelper.GetJonsArrayFromFile("TransactionsArray.json");
            jsonArray.Count.Should().Be(122);
        }
        [Fact]
        public void GetAccountMovmentsFromTransactionJarray()
        {
            var jsonArray = TestsHelper.GetJonsArrayFromFile("TransactionsArray.json");
            List<AccountMovement> accountMovements = ModelConverter.GetAccountMovmentsFromJarray(jsonArray);

            accountMovements.Count.Should().Be(122);
        }
        [Fact]
        public void GetJonsArrayFromSubCategoriesFile()
        {
            var jsonArray = TestsHelper.GetJonsArrayFromFile("CategoriesArray.json");
            jsonArray.Count.Should().Be(105);
        }
        [Fact]
        public void GetAccountMovmentsFromSubCategoriesFile()
        {
            var jsonArray = TestsHelper.GetJonsArrayFromFile("CategoriesArray.json");
            List<SubCategory> categorisModel = ModelConverter.GetCategoriesFromJarray(jsonArray);
            categorisModel.Count.Should().Be(105);
        }
        [Fact]
        public void GetPropertyValueTest()
        {
            var jsonArray = TestsHelper.GetJonsArrayFromFile("TransactionsArray.json");
            List<AccountMovement> accountMovements = ModelConverter.GetAccountMovmentsFromJarray(jsonArray);
            accountMovements.Count.Should().Be(122);

            var movement = accountMovements[0];
            var propertyValue = ModelConverter.GetPropertyValue(movement, "Amount");
            propertyValue.Should().BeEquivalentTo(35);
        }

        [Fact]
        public void AddValuesToMovementsViewModelTests()
        {
            var jsonArray = TestsHelper.GetJonsArrayFromFile("TransactionsArray.json");
            List<AccountMovement> accountMovements = ModelConverter.GetAccountMovmentsFromJarray(jsonArray);
            accountMovements.Count.Should().Be(122);
            MovementsViewModel movementsViewModel = new MovementsViewModel();
            ModelConverter.AddValuesToMovementsViewModel(accountMovements[0], ref movementsViewModel);

            movementsViewModel.Amount.Should().Be(35);
        }

        [Fact]
        public void GetAccountMovmentsFromSubCategoriesFile1()
        {
            var jsonArray = TestsHelper.GetJonsArrayFromFile("TransactionsArray.json");
            List<AccountMovement> accountMovements = ModelConverter.GetAccountMovmentsFromJarray(jsonArray);
            accountMovements.Count.Should().Be(122);

            jsonArray = TestsHelper.GetJonsArrayFromFile("CategoriesArray.json");
            List<SubCategory> categorisModel = ModelConverter.GetCategoriesFromJarray(jsonArray);
            categorisModel.Count.Should().Be(105);

            var modementsViewModels = ModelConverter.CreateMovementsViewModels(accountMovements, categorisModel, "Felles");

            modementsViewModels[0].Category.Should().BeEquivalentTo("Altibox");

        }

        [Fact]
        public void AddSubcategoriesToMovementTest()
        {
            var jsonArray = TestsHelper.GetJonsArrayFromFile("CategoriesArray.json");
            List<SubCategory> categorisModel = ModelConverter.GetCategoriesFromJarray(jsonArray);
            categorisModel.Count.Should().Be(105);

            var noko = categorisModel.Where(c => (c.Category == "Mat" || c.Category == "Familly"));
            var subcategory = ModelConverter.AddSubcategoriesToMovement(noko);

            subcategory.SubPorject.Should().BeEquivalentTo("Mismatch");
        }
        [Fact]
        public void GetKeewordfromTransactionTextTest()
        {
            var jsonArray = TestsHelper.GetJonsArrayFromFile("CategoriesArray.json");
            List<SubCategory> categorisModel = ModelConverter.GetCategoriesFromJarray(jsonArray);

            var movementsViewModel = new MovementsViewModel()
            {
                Text = "Innskudd - ATM SPAREBANKEN DIN   TER4077"
                //Text = "Varekjøp BRUKÅS SPORT AS BØGATA 56 BØ I TELEMARK betal dato 2017-11-11"
                //Text = "Varekjøp REMA BØ LUNDEVEGEN 1 BØ I TELEMARK betal dato 2017-06-06"
                //Text = @"VISA VARE 429941XXXXXX9552 10.09 IZ *PETS OF NORWAY\OSLO"
                //Text = @"VISA VARE 429941XXXXXX5850 09.10  0,00 ITUNES.COM/BILL\80056952 Kurs  100,0000E"
            };
            var subcategory = ModelConverter.GetTransactionCategoryFromKeewordList(categorisModel, movementsViewModel);

            subcategory.Category.Should().NotBeEmpty();
        }
    }
}
