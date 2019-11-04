using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Transactions.Models;

namespace Transactions.Services
{
    public class ModelClassServices
    {
        public AccountMovement JsonToAccountMovement(JToken json)
        {
            AccountMovement accountMovement = new AccountMovement();
            var noko = ParseObjectProperties(accountMovement, json);
            return accountMovement;
        }

        public static List<SubCategory> GetSubCategoriesFromJarray(JArray jArray)
        {
            var subCategories = new List<SubCategory>();
            foreach (var item in jArray)
            {
                subCategories.Add(new ModelClassServices().JsonToSubCategory(item));
            }
            return subCategories;
        }

        public static List<AccountMovement> GetAccountMovmentsFromJarray(JArray jArray)
        {

            var accountmovments = new List<AccountMovement>();
            foreach (var item in jArray)
            {
                accountmovments.Add((AccountMovement)ModelClassServices.ParseObjectProperties(new AccountMovement(), item));
            }
            return accountmovments;
        }

        public SubCategory JsonToSubCategory(JToken json)
        {
            SubCategory subCategories = new SubCategory();
            var noko = ParseObjectProperties(subCategories, json);

            return subCategories;

        }

        public MovementsViewModel JsonToMovementsViewModels(JToken json)
        {
            var movementsViewModel = new MovementsViewModel();
            var noko = ParseObjectProperties(movementsViewModel, json);

            return movementsViewModel;
        }

        public static object ParseObjectProperties(Object model, JToken json)
        {
            var type = model.GetType();
            var typeProperties = type.GetProperties();

            foreach (var property in typeProperties)
            {
                var propertyName = property.Name;
                var jsonPropertyValue = json[propertyName];
                if (jsonPropertyValue != null)
                {
                    var jtokenValue = jsonPropertyValue.ToString();
                    var propertyType = property.PropertyType.Name;
                    var value = ParseObjectValue(propertyType, jtokenValue);
                    property.SetValue(model, value);
                }
            }
            return model;
        }

        public static object ParseObjectValue(string type, string value)
        {
            switch (type.ToLower())
            {
                case "string":
                    return value;
                case "datetime":
                    DateTime dateTime;
                    if (DateTime.TryParse(value, out dateTime))
                        return dateTime;
                    return null;

                case "int16":
                case "int32":
                case "int64":
                case "integer":
                    int integer;
                    if (int.TryParse(value, out integer))
                        return integer;
                    return null;
                case "double":
                    double doubleValue;
                    if (double.TryParse(value, out doubleValue))
                        return doubleValue;
                    return null;
                default:
                    return null;
            }
        }

        public static List<string> GetPropertiesNamesFromObject(object model)
        {
            var properties = model?.GetType().GetProperties();
            if (properties != null)
            {
                return properties.Select(prop => prop.Name).ToList();
            }
            return null;
        }

        public static List<MovementsViewModel> CreateMovementsViewModels(List<AccountMovement> accountMovements, List<SubCategory> subCategories, string acountName)
        {
            var moventsViewModel = new List<MovementsViewModel>();
            foreach (var movement in accountMovements)
            {
                MovementsViewModel movementViewModel = new MovementsViewModel() { AcountName = acountName };

                // Add values to model if it find same property name
                AddValueToMovementsModel(movement, ref movementViewModel);

                movementViewModel = UpdateMovementViewModelWithSubCategory(subCategories, movementViewModel);

                moventsViewModel.Add(movementViewModel);
            }
            AddUnspecifiedTransaction(ref moventsViewModel);
            return moventsViewModel;
        }

        private static void AddUnspecifiedTransaction(ref List<MovementsViewModel> moventsViewModel)
        {
            var listOfUnspecifiedTransaction = moventsViewModel.Where(mv => string.IsNullOrEmpty(mv.Category));
            foreach (var movent in listOfUnspecifiedTransaction)
            {
                movent.Category = "Unspecified";
            }
        }

        public static MovementsViewModel UpdateMovementViewModelWithSubCategory(List<SubCategory> subCategories, MovementsViewModel movementModel)
        {
            try
            {

                if (!string.IsNullOrEmpty(movementModel.Text))
                {
                    var subcategoriesMatch = subCategories.Where(sub => CultureInfo.InvariantCulture.CompareInfo.LastIndexOf(movementModel.Text, sub.KeyWord.ToLower(), CompareOptions.IgnoreCase) >= 0);

                    if (subcategoriesMatch != null && subcategoriesMatch.Count() > 0)
                    {
                        if (subcategoriesMatch.Count() == 1)
                            AddValueToMovementsModel(subcategoriesMatch.FirstOrDefault(), ref movementModel);
                        else
                            AddValueToMovementsModel(AddSubcategoriesToMovement(subcategoriesMatch), ref movementModel);
                    }
                }
            }
            catch
            {
                //
            }
            return movementModel;
        }

        public static List<string> GetExtractionCategories(IEnumerable<string> categories, List<MovementsViewModel> movementsModel)
        {
            List<string> categoriesExtractions = new List<string>();
            foreach (var category in categories)
            {
                if (movementsModel.Where(mv => mv.Category == category).Any(mv => mv.Amount < 0))
                    categoriesExtractions.Add(category);
            }
            return categoriesExtractions;
        }
        public static List<string> GetIncomsCategories(IEnumerable<string> categories, List<MovementsViewModel> movementsModel)
        {
            List<string> categoriesExtractions = new List<string>();
            foreach (var category in categories)
            {
                if (movementsModel.Where(mv => mv.Category == category).Any(mv => mv.Amount > 0))
                    categoriesExtractions.Add(category);
            }
            return categoriesExtractions;
        }

        private static SubCategory AddSubcategoriesToMovement(IEnumerable<SubCategory> subcategoriesMatch)
        {
            var subcategory = new SubCategory();
            string SupPorject = null;
            var moreThanOneCategory = subcategoriesMatch.Select(sub => sub.Category).Distinct().Count() > 1;
            if (moreThanOneCategory)
            {

                var subcategoryNames = subcategoriesMatch.Select(sub => sub.KeyWord).ToArray();
                var subcategoryCategories = subcategoriesMatch.Select(sub => sub.Category).ToArray();
                if (subcategoryCategories.Contains("Mat"))
                {
                    subcategory.KeyWord = string.Join(",", subcategoryNames);
                    subcategory.Category = "Mat";
                    SupPorject = "Mismatch";
                }
                else if (subcategoryCategories.Contains("Vinmonopolet"))
                {
                    subcategory.KeyWord = string.Join(",", subcategoryNames);
                    subcategory.Category = "Vinmonopolet";
                    SupPorject = "Mismatch";
                }
                else if (subcategoryNames.Contains("ffo"))
                {
                    subcategory.KeyWord = string.Join(",", subcategoryNames);
                    subcategory.Category = subcategoriesMatch.First(cat => cat.KeyWord == "ffo").Category;
                    SupPorject = "Mismatch";
                }
                else if (ArrayCointains(subcategoryNames, "Matias"))
                {
                    subcategory.KeyWord = string.Join(",", subcategoryNames);
                    subcategory.Category = subcategoriesMatch.First(cat => cat.KeyWord == "matias").Category;
                }
                else if (ArrayCointains(subcategoryNames, "Åse"))
                {
                    subcategory.KeyWord = string.Join(",", subcategoryNames);
                    subcategory.Category = subcategoriesMatch.First(cat => cat.KeyWord == "Åse").Category;
                }
                else
                {
                    subcategory.KeyWord = string.Join(",", subcategoryNames);
                    subcategory.Category = string.Join(",", subcategoryCategories);
                    SupPorject = "Mismatch";
                }
                subcategory.SupPorject = SupPorject;
            }
            else
            {
                subcategory = subcategoriesMatch.FirstOrDefault();
            }
            return subcategory;
        }

        private static bool ArrayCointains(string[] subcategoryNames, string name)
        {
            return subcategoryNames.Any(sub => CultureInfo.InvariantCulture.CompareInfo.LastIndexOf(name, sub, CompareOptions.IgnoreCase) >= 0);
        }

        //Loop through all the properties
        private static void AddValueToMovementsModel(object movement, ref MovementsViewModel movementsViewModel)
        {
            foreach (var property in movementsViewModel.GetType().GetProperties())
            {
                var propertyValue = GetPropertyValue(movement, property.Name);
                if (propertyValue != null)
                    SetPropertyValueToMovementsViewModel(property.Name, propertyValue, ref movementsViewModel);
            }
        }

        public static object GetPropertyValue(object model, string propertyName)
        {
            try
            {
                var properties = GetPropertiesNamesFromObject(model);
                if (properties.Contains(propertyName))
                {
                    object result = model.GetType().GetProperty(propertyName).GetValue(model, null);
                    return result;
                }
                return null;
            }
            catch
            {
                return null;
            }
        }

        public static bool SetPropertyValueToMovementsViewModel(string propertyName, object propertyValue, ref MovementsViewModel modelToUpdate)
        {
            try
            {
                var properties = GetPropertiesNamesFromObject(modelToUpdate);
                if (properties.Contains(propertyName))
                {
                    modelToUpdate.GetType().GetProperty(propertyName).SetValue(modelToUpdate, propertyValue);
                    return true;
                }
                return false;
            }
            catch
            {
                return false;
            }
        }

        public static double TotalforCategory(IEnumerable<MovementsViewModel> movements, string category, int? year = null, int? month = null, bool justExtrations = true)
        {

            var monthAndYaerMovements = GetMovementByMonthYear(movements, year, month);

            if (monthAndYaerMovements.Any())
            {
                var movementsByCategory = monthAndYaerMovements.Where(mov => mov.Category == category);
                return SumByType(justExtrations, movementsByCategory);
            }
            return 0;
        }

        public static double AverageforCategory(IEnumerable<MovementsViewModel> movements, string category, int? year = null, int? month = null, bool justExtrations = true)
        {
            var monthAndYaerMovements = GetMovementByMonthYear(movements, year, month);
            double average = 0;
            if (monthAndYaerMovements.Any())
            {
                //TODO get average for year / month for the category
                var movementsByCategory = monthAndYaerMovements.Where(mov => mov.Category == category && mov.Amount != 0);
                movementsByCategory = GetMovementsViewModelsByMovmentType(justExtrations, movementsByCategory);
                if (month != null && year == null)
                {
                    var result1 = movementsByCategory.GroupBy(mv => mv.DateTime.Year).Select(mov => new { Year = mov.Key, sum = mov.Sum(p => p.Amount) });
                    average = !result1.Any() ? 0 : result1.Average(r => r.sum);
                }

                if (year != null && month == null)
                {
                    var result1 = movementsByCategory.GroupBy(mv => mv.DateTime.Month).Select(mov => new { Month = mov.Key, sum = mov.Sum(p => p.Amount) });
                    average = !result1.Any() ? 0 : result1.Average(r => r.sum);
                }

                if (year == null && month == null)
                {
                    var group = movementsByCategory.GroupBy(mv => mv.DateTime.Day);
                    var count = group.Count();
                    var sum = group.Select(mov => new { Day = mov.Key, sum = mov.Sum(p => p.Amount) });
                    var result1 = movementsByCategory.GroupBy(mv => mv.DateTime.Day).Select(mov => new { Day = mov.Key, sum = mov.Sum(p => p.Amount) });
                    average = !result1.Any() ? 0 : result1.Average(r => r.sum);

                }

                return Math.Abs(average);
            }
            return Math.Abs(average);
        }

        private static IEnumerable<MovementsViewModel> GetMovementsViewModelsByMovmentType(bool justExtrations, IEnumerable<MovementsViewModel> monthAndYaerMovements)
        {
            if (justExtrations)
                return monthAndYaerMovements.Where(mv => mv.Amount < 0).ToList();
            else
                return monthAndYaerMovements.Where(mv => mv.Amount > 0).ToList();
        }

        private static IEnumerable<MovementsViewModel> GetMovementByMonthYear(IEnumerable<MovementsViewModel> movements, int? year = null, int? month = null)
        {
            IEnumerable<MovementsViewModel> movementsTemp = null;
            if (year != null)
                movementsTemp = movements.Where(move => move.DateTime.Year == year);
            if (month != null)
            {
                if (movementsTemp != null)
                    movementsTemp = movementsTemp.Where(move => move.DateTime.Month == month);
                else
                    movementsTemp = movements.Where(move => move.DateTime.Month == month);
            }
            return movementsTemp ?? movements;
        }

        public static double CategoriesMonthYearTotal(IEnumerable<MovementsViewModel> movements, int? year = null, int? month = null, bool justExtrations = true)
        {

            var monthAndYaerMovements = movements.Where(mov => !string.IsNullOrEmpty(mov.Category) && mov.DateTime.Year == year && mov.DateTime.Month == month);
            return SumByType(justExtrations, monthAndYaerMovements);
        }
        public static double MonthYearTotal(IEnumerable<MovementsViewModel> movements, int? year = null, int? month = null, bool justExtrations = true)
        {
            var monthAndYaerMovements = movements.Where(mov => mov.DateTime.Year == year && mov.DateTime.Month == month);

            return SumByType(justExtrations, monthAndYaerMovements);
        }


        private static double SumByType(bool justExtrations, IEnumerable<MovementsViewModel> monthAndYaerMovements)
        {
            double sum = 0;
            if (monthAndYaerMovements.Any())
            {
                if (justExtrations)
                {
                    var sum1 = monthAndYaerMovements.Where(mv => mv.Amount < 0).ToList();
                    sum = monthAndYaerMovements.Where(mv => mv.Amount < 0).Sum(cat => Math.Abs(cat.Amount));
                }
                else
                {
                    var sum1 = monthAndYaerMovements.Where(mv => mv.Amount > 0).ToList();
                    sum = monthAndYaerMovements.Where(mv => mv.Amount > 0).Sum(cat => Math.Abs(cat.Amount));
                }
            }

            return sum;
        }

        public static List<string> GetListOfCategories(List<MovementsViewModel> momvents)
        {
            var list = momvents.Where(m => m.SupPorject != "Mismatch" && !string.IsNullOrEmpty(m.Category))
                .Select(m => m.Category).Distinct().ToList();
            return list;
        }

    }
}
