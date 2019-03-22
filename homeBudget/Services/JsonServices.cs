using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using homeBudget.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace homeBudget.Services
{
    public class JsonServices
    {
        public SubCategory GetSubCategory(JToken json)
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
