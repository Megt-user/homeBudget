using System;
using System.Collections.Generic;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using homeBudget.Models;

namespace homeBudget.Services
{
    public class JsonServices
    {
        public static SubCategory GetSubCategory(JToken json)
        {
            return json.ToObject<SubCategory>();
            //var subCategory = Newtonsoft.Json.JsonConvert.DeserializeObject<SubCategories>(json);
        }

        public Transaction GetAcountMovment(JToken movment)
        {
            var accountMovement = new Transaction();
            DateTime movementDateTime;

            if (DateTime.TryParse(movment["DateTime"].ToString(), out movementDateTime))
                accountMovement.DateTime = movementDateTime;

            movment.ToObject<Transaction>();
            return accountMovement;
        }

        public static JsonSerializerSettings GetJsonSerializerSettings()
        {
            return new JsonSerializerSettings
            {
                Converters = new List<JsonConverter> { new JsonDateFixingConverter() },
                DateParseHandling = DateParseHandling.None,
                NullValueHandling = NullValueHandling.Ignore,
                MissingMemberHandling = MissingMemberHandling.Ignore
            };
        }

    }
}
