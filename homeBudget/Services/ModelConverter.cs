using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using homeBudget.Models;
using Newtonsoft.Json.Linq;

namespace homeBudget.Services
{
    public class ModelConverter
    {
        public AccountMovement JsonToAccountMovement(JToken json)
        {
            AccountMovement accountMovement = new AccountMovement();
            var noko = ParseObjectProperties(accountMovement, json);
            return accountMovement;
        }

        public static List<SubCategory> GetCategoriesFromJarray(JArray jArray)
        {
            var subCategories = new List<SubCategory>();
            foreach (var item in jArray)
            {
                subCategories.Add(new ModelConverter().JsonToSubCategory(item));
            }
            return subCategories;
        }

        public static List<AccountMovement> GetAccountMovmentsFromJarray(JArray jArray)
        {
            return jArray.Select(item => (AccountMovement)ParseObjectProperties(new AccountMovement(), item)).ToList();
        }

        public SubCategory JsonToSubCategory(JToken json)
        {
            SubCategory subCategories = new SubCategory();
            var noko = ParseObjectProperties(subCategories, json);

            return subCategories;

        }

        public TransactionViewModel JsonToMovementsViewModels(JToken json)
        {
            var movementsViewModel = new TransactionViewModel();
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
            return properties?.Select(prop => prop.Name).ToList();
        }

        public static List<TransactionViewModel> CreateMovementsViewModels(List<AccountMovement> accountMovements, List<SubCategory> subCategories, string acountName)
        {
            var moventsViewModel = new List<TransactionViewModel>();
            foreach (var movement in accountMovements)
            {
                TransactionViewModel transactionViewModel = new TransactionViewModel() { AcountName = acountName };

                // Add values to model if it find same property name
                AddValuesToMovementsViewModel(movement, ref transactionViewModel);

                transactionViewModel = UpdateMovementViewModelWithSubCategory(subCategories, transactionViewModel);

                moventsViewModel.Add(transactionViewModel);
            }
            AddUnspecifiedTransaction(ref moventsViewModel);
            return moventsViewModel;
        }

        private static void AddUnspecifiedTransaction(ref List<TransactionViewModel> moventsViewModel)
        {
            var listOfUnspecifiedTransaction = moventsViewModel.Where(mv => string.IsNullOrEmpty(mv.Category));
            foreach (var movent in listOfUnspecifiedTransaction)
            {
                movent.Category = "Unspecified";
            }
        }

        public static TransactionViewModel UpdateMovementViewModelWithSubCategory(List<SubCategory> subCategories, TransactionViewModel transactionModel)
        {
            try
            {

                if (!string.IsNullOrEmpty(transactionModel.Text))
                {

                    var subcategoriesMatch = subCategories.Where(sub => CultureInfo.InvariantCulture.CompareInfo.LastIndexOf(transactionModel.Text, sub.KeyWord.ToLower(), CompareOptions.IgnoreCase) >= 0);

                    if (subcategoriesMatch != null && subcategoriesMatch.Count() > 0)
                    {
                        if (subcategoriesMatch.Count() == 1)
                            AddValuesToMovementsViewModel(subcategoriesMatch.FirstOrDefault(), ref transactionModel);
                        else
                            AddValuesToMovementsViewModel(AddSubcategoriesToMovement(subcategoriesMatch), ref transactionModel);
                    }
                }
            }
            catch
            {
                //
            }
            return transactionModel;
        }

        //TODO check name and transaction value to create a rule to place the transaction in the right category f.eks. Husly > 200 NOK = House not Social
        public static SubCategory AddSubcategoriesToMovement(IEnumerable<SubCategory> subcategoriesMatch)
        {
            var subcategory = new SubCategory();
            string subCategoryName = null;
            string subProject = null;

            var subCategories = subcategoriesMatch as SubCategory[] ?? subcategoriesMatch.ToArray();
            var moreThanOneCategory = subCategories.Select(sub => sub.Category).Distinct().Count() > 1;

            if (moreThanOneCategory)
            {
                var keeWords = subCategories.Select(sub => sub.KeyWord).Distinct().ToArray();
                var subcategoryCategories = subCategories.Select(sub => sub.Category).Distinct().ToArray();
                if (ArrayCointains(subcategoryCategories, "Mat"))
                {
                    subCategoryName = "Mat";
                    subProject = "Mismatch";
                }
                else if (ArrayCointains(subcategoryCategories, "Vinmonopolet"))
                {
                    subCategoryName = "Vinmonopolet";
                    subProject = "Mismatch";
                }
                else if (ArrayCointains(keeWords, "ffo"))
                {
                    subCategoryName = "ffo";
                    subProject = "Mismatch";
                }
                else if (ArrayCointains(keeWords, "Matias"))
                {
                    subcategory.Category = "Matias";
                }else if (ArrayCointains(keeWords, "Åse"))
                {
                    subcategory.Category = "Åse";

                }else if (ArrayCointains(keeWords, "Oscar"))
                {
                    subCategoryName = "Oscar";
                    subProject = "Mismatch";
                }else if (ArrayCointains(keeWords, "BRUKÅS"))
                {
                    subCategoryName = "Sport";
                    subProject = "Mismatch";
                }
                else if (ArrayCointains(keeWords, "Hermann Ivarson"))
                {
                    subCategoryName = "Utlaie";
                    subProject = "Mismatch";
                } else if (ArrayCointains(keeWords, "Forsikring"))
                {
                    subCategoryName = "Forsikring";
                    subProject = "Mismatch";
                }
                //TODO verifique cómo crear privilegios para configurar la subcategoría por ejemplo Hovden / Mat subcategoría cuando Mat debe ser comida pero otras causas Fritid
                else if (ArrayCointains(keeWords, "yx")) 
                {
                    subCategoryName = "Diesel";
                    subProject = "Mismatch";
                }else if (ArrayCointains(keeWords, "CIRCLE K")) 
                {
                    subCategoryName = "Diesel";
                    subProject = "Mismatch";
                }else if (ArrayCointains(keeWords, "Hovden")) 
                {
                    subCategoryName = "Fritid";
                    subProject = "Mismatch";
                }else if (ArrayCointains(keeWords, "cf")) 
                {
                    subCategoryName = "House";
                    subProject = "Mismatch";
                }else if (ArrayCointains(keeWords, "HVASSER")) 
                {
                    subCategoryName = "Fritid";
                    subProject = "Mismatch";
                }else if (ArrayCointains(keeWords, "Husly")) 
                {
                    subCategoryName = "House";
                    subProject = "Mismatch";
                }else if (ArrayCointains(keeWords, "SANDEN CAMPING")) 
                {
                    subCategoryName = "Fritid";
                    subProject = "Mismatch";
                }else if (ArrayCointains(keeWords, "SKARPHEDIN")) 
                {
                    subCategoryName = "Familly";
                    subProject = "Mismatch";
                }
                else
                {
                    subCategoryName = string.Join(",", subcategoryCategories);
                    subProject = "Mismatch";
                }
               
                subcategory.KeyWord = string.Join(",", keeWords);
                subcategory.Category = subCategoryName;
                subcategory.SubPorject = subProject;
            }
            else
            {
                subcategory = subCategories.FirstOrDefault();
            }
            return subcategory;
        }

        private static bool ArrayCointains(string[] subcategoryNames, string name)
        {
            return subcategoryNames.Any(sub => CultureInfo.InvariantCulture.CompareInfo.LastIndexOf(name, sub, CompareOptions.IgnoreCase) >= 0);
        }

        //Loop through all the properties
        public static void AddValuesToMovementsViewModel(object movement, ref TransactionViewModel transactionViewModel)
        {
            foreach (var property in transactionViewModel.GetType().GetProperties())
            {
                var propertyValue = GetPropertyValue(movement, property.Name);
                if (propertyValue != null)
                {
                    var properties = GetPropertiesNamesFromObject(transactionViewModel);
                    if (properties.Contains(property.Name))
                    {
                        transactionViewModel.GetType().GetProperty(property.Name)?.SetValue(transactionViewModel, propertyValue);
                    }
                }
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


        public static IEnumerable<TransactionViewModel> GetMovementsViewModelsByType(bool justExtrations, IEnumerable<TransactionViewModel> monthAndYaerMovements)
        {
            if (justExtrations)
                return monthAndYaerMovements.Where(mv => mv.Amount < 0).ToList();
            else
                return monthAndYaerMovements.Where(mv => mv.Amount > 0).ToList();
        }



        public static double CategoriesMonthYearTotal(IEnumerable<TransactionViewModel> movements, int? year = null, int? month = null, bool justExtrations = true)
        {

            var monthAndYaerMovements = movements.Where(mov => !string.IsNullOrEmpty(mov.Category) && mov.DateTime.Year == year && mov.DateTime.Month == month);
            return ModelOperation.SumByType(monthAndYaerMovements, justExtrations);
        }
        public static double MonthYearTotal(IEnumerable<TransactionViewModel> movements, int? year = null, int? month = null, bool justExtrations = true)
        {
            var monthAndYaerMovements = movements.Where(mov => mov.DateTime.Year == year && mov.DateTime.Month == month);

            return ModelOperation.SumByType(monthAndYaerMovements, justExtrations);
        }

        public static List<string> GetListOfCategories(List<TransactionViewModel> momvents)
        {
            var list = momvents.Where(m => m.SupPorject != "Mismatch" && !string.IsNullOrEmpty(m.Category))
                .Select(m => m.Category).Distinct().ToList();
            return list;

        }
    }
}
