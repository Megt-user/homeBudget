using System;
using Newtonsoft.Json.Linq;
using Transactions.Models;

namespace Transactions.Services
{
    public class JsonServices
    {
        public static SubCategory GetSubCategory(JToken json)
        {
            return json.ToObject<SubCategory>();
            //var subCategory = Newtonsoft.Json.JsonConvert.DeserializeObject<SubCategories>(json);
        }

        public AccountMovement GetAcountMovment(JToken movment)
        {
            var accountMovement = new AccountMovement();
            DateTime movementDateTime;

            if (DateTime.TryParse(movment["DateTime"].ToString(), out movementDateTime))
                accountMovement.DateTime = movementDateTime;

            movment.ToObject<AccountMovement>();
            return accountMovement;
        }

    }   
}
