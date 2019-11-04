using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using homeBudget.Models;
using homeBudget.Services;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

namespace homeBudget
{
    public class ModelOperation
    {
        /// <summary>
        /// Get all category movements from enumeration that mutch Year or/and month
        /// </summary>
        /// <param name="movements"></param>
        /// <param name="category"></param>
        /// <param name="year"></param>
        /// <param name="month"></param>
        /// <param name="justExtrations">'true' it is just extractions</param>
        /// <returns></returns>
        public static double GetTotalforCategory(IEnumerable<MovementsViewModel> movements, string category, int? year = null, int? month = null, bool justExtrations = true)
        {

            var monthAndYaerMovements = GetMovementByMonthYear(movements, year, month);

            if (monthAndYaerMovements.Any())
            {
                var movementsByCategory = monthAndYaerMovements.Where(mov => mov.Category == category);
                return SumByType(movementsByCategory, justExtrations);
            }
            return 0;
        }

        /// <summary>
        /// Get all movements from enumeration that mutch Year or/and month
        /// </summary>
        /// <param name="movements"></param>
        /// <param name="year"></param>
        /// <param name="month"></param>
        /// <returns></returns>
        public static IEnumerable<MovementsViewModel> GetMovementByMonthYear(IEnumerable<MovementsViewModel> movements, int? year = null, int? month = null)
        {
            IEnumerable<MovementsViewModel> movementsTemp = null;
            var movementsViewModels = movements as MovementsViewModel[] ?? movements.ToArray();
            if (year != null)
                movementsTemp = movementsViewModels.Where(move => move.DateTime.Year == year);
            if (month != null)
            {
                movementsTemp = movementsTemp != null ? movementsTemp.Where(move => move.DateTime.Month == month) : movementsViewModels.Where(move => move.DateTime.Month == month);
            }
            return movementsTemp ?? movements;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="categories"></param>
        /// <param name="movementsModel"></param>
        /// <returns></returns>
        public static List<string> GetExtractionCategories(IEnumerable<string> categories, List<MovementsViewModel> movementsModel)
        {
            return categories.Where(category => movementsModel.Where(mv => mv.Category == category).Any(mv => mv.Amount < 0)).ToList();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="categories"></param>
        /// <param name="movementsModel"></param>
        /// <returns></returns>
        public static List<string> GetIncomsCategories(IEnumerable<string> categories, List<MovementsViewModel> movementsModel)
        {
            return categories.Where(category => movementsModel.Where(mv => mv.Category == category).Any(mv => mv.Amount > 0)).ToList();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="justExtrations">'true' if it is just extractions, 'false' just incomes, 'null' both</param>
        /// <param name="movementsViewModels"></param>
        /// <returns></returns>
        public static double SumByType(IEnumerable<MovementsViewModel> movementsViewModels, bool? justExtrations = null)
        {
            double sum = 0;
            var viewModels = movementsViewModels as MovementsViewModel[] ?? movementsViewModels.ToArray();
            if (viewModels.Any())
            {
                if (justExtrations.HasValue)
                {
                    if (justExtrations.GetValueOrDefault(false))
                    {
                        var sum1 = viewModels.Where(mv => mv.Amount < 0).ToList();
                        sum = viewModels.Where(mv => mv.Amount < 0).Sum(cat => cat.Amount);
                    }
                    else
                    {
                        var sum1 = viewModels.Where(mv => mv.Amount > 0).ToList();
                        sum = viewModels.Where(mv => mv.Amount > 0).Sum(cat => cat.Amount);
                    }
                }
                else
                {
                    sum = viewModels.Sum(cat => cat.Amount);
                }
            }

            //return Math.Abs(sum);
            return sum;
        }



        public static double AverageforCategory(IEnumerable<MovementsViewModel> movements, string category, int? year = null, int? month = null, bool justExtrations = true)
        {
            var monthAndYaerMovements = ModelOperation.GetMovementByMonthYear(movements, year == 0 ? null : year, month == 0 ? null : month);

            double average = 0;

            //TODO get average for year / month for the category
            var movementsByCategory = monthAndYaerMovements.Where(mov => mov.Category == category && Math.Abs(mov.Amount) > 0);
            movementsByCategory = ModelConverter.GetMovementsViewModelsByType(justExtrations, movementsByCategory);

            if (movementsByCategory != null && movementsByCategory.Any())
            {
                //Years average
                if (month == null && year == null)
                {
                    var yearResult = movementsByCategory.GroupBy(mv => mv.DateTime.Year).Select(mov => new { Year = mov.Key, sum = mov.Sum(p => p.Amount) });
                    average = !yearResult.Any() ? 0 : yearResult.Average(r => r.sum);
                }

                // Get the months average
                if (month == 0)
                {
                    var monthResult = movementsByCategory.GroupBy(mv => mv.DateTime.Month).Select(mov => new { Month = mov.Key, sum = mov.Sum(p => p.Amount) });
                    average = !monthResult.Any() ? 0 : monthResult.Average(r => r.sum);
                }
                // Get the average of a specific month
                if (month != 0 && month != null)
                {
                    var monthResult = movementsByCategory.GroupBy(mv => mv.DateTime.Month).Select(mov => new { Month = mov.Key, sum = mov.Sum(p => p.Amount) });
                    average = !monthResult.Any() ? 0 : monthResult.Average(r => r.sum);
                }

                //if (year != null && month == 0)
                //{
                //    var monthResult = movementsByCategory.GroupBy(mv => mv.DateTime.Day).Select(mov => new { Month = mov.Key, sum = mov.Sum(p => p.Amount) });
                //    average = !monthResult.Any() ? 0 : monthResult.Average(r => r.sum);
                //}


                if (year != null && (month != null && month != 0))
                {
                    average = movementsByCategory.Average(m => m.Amount);
                }
            }
            //return Math.Abs(average);
            return average;
        }
    }
}
