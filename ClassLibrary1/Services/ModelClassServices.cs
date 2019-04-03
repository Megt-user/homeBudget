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

        public static List<MovementsViewModel> getListOfModementsViewModel(List<AccountMovement> accountMovements, List<SubCategory> subCategories, string acountName)
        {
            var moventsViewModel = new List<MovementsViewModel>();
            foreach (var movement in accountMovements)
            {
                MovementsViewModel movementViewModel = new MovementsViewModel() { AcountName = acountName };
                FillUpMovementViewModel(movement, ref movementViewModel);
                movementViewModel = UpdateMovementViewModelWithSubCategory(subCategories, movementViewModel);
                moventsViewModel.Add(movementViewModel);
            }
            return moventsViewModel;
        }

        private static MovementsViewModel UpdateMovementViewModelWithSubCategory(List<SubCategory> subCategories, MovementsViewModel movementModel)
        {
            try
            {

                if (!string.IsNullOrEmpty(movementModel.Text))
                {
                    var subcategoriesMatch = subCategories.Where(sub => CultureInfo.InvariantCulture.CompareInfo.LastIndexOf(movementModel.Text, sub.KeyWord.ToLower(), CompareOptions.IgnoreCase) > 0);

                    if (subcategoriesMatch != null && subcategoriesMatch.Count() > 0)
                    {
                        if (subcategoriesMatch.Count() == 1)
                            FillUpMovementViewModel(subcategoriesMatch.FirstOrDefault(), ref movementModel);
                        else
                            FillUpMovementViewModel(ChoseSubCategory(subcategoriesMatch), ref movementModel);
                    }

                }
            }
            catch
            {
                //
            }
            return movementModel;
        }



        private static SubCategory ChoseSubCategory(IEnumerable<SubCategory> subcategoriesMatch)
        {
            var subcategory = new SubCategory();
            var moreThanOneCategory = subcategoriesMatch.Select(sub => sub.Category).Distinct().Count() > 1;
            if (moreThanOneCategory)
            {
                var subcategoryNames = subcategoriesMatch.Select(sub => sub.KeyWord).ToArray();
                var subcategoryCategories = subcategoriesMatch.Select(sub => sub.Category).ToArray();
                subcategory.KeyWord = string.Join(",", subcategoryNames);
                subcategory.Category = string.Join(",", subcategoryCategories);
                subcategory.SupPorject = "Mismatch";
            }
            else
            {
                subcategory = subcategoriesMatch.FirstOrDefault();
            }
            return subcategory;
        }

        private static void FillUpMovementViewModel(object movement, ref MovementsViewModel movementsViewModel)
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

        public static IEnumerable<MovementsViewModel> GetAllMonthAndYaerMovements(List<MovementsViewModel> movements, int year, int month)
        {
            return movements.Where(move => move.DateTime.Year == year && move.DateTime.Month == month);
        }

        public static double? GetTotalforCategory(IEnumerable<MovementsViewModel> movements, string category, int? year = null, int? month = null)
        {

            var monthAndYaerMovements = movements.Where(move => move.DateTime.Year == year && move.DateTime.Month == month);

            if (monthAndYaerMovements.Any())
            {
                return monthAndYaerMovements.Where(mov => mov.Category == category).Sum(cat => Math.Abs(cat.Amount));
            }

            return null;
        }
        public static double? GetCategoriesMonthYearTotal(IEnumerable<MovementsViewModel> movements, int? year = null, int? month = null)
        {

            var monthAndYaerMovements = movements.Where(mov => !string.IsNullOrEmpty(mov.Category) && mov.DateTime.Year == year && mov.DateTime.Month == month);

            if (monthAndYaerMovements.Any())
            {
                return monthAndYaerMovements.Sum(cat => Math.Abs(cat.Amount));
            }

            return null;
        }
        public static double? GetMonthYearTotal(IEnumerable<MovementsViewModel> movements, int? year = null, int? month = null)
        {

            var monthAndYaerMovements = movements.Where(mov => mov.DateTime.Year == year && mov.DateTime.Month == month);

            if (monthAndYaerMovements.Any())
            {
                return monthAndYaerMovements.Sum(cat => Math.Abs(cat.Amount));
            }

            return null;
        }


    }
}
