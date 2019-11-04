using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using homeBudget.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace homeBudget.Controllers
{
    [Produces("application/json")]
    [Route("api/[controller]")]
    public class ValuesController : ControllerBase
    {

        //GET api/values/5
        [HttpGet]
        public IEnumerable<string> Get()
        {
            return new string[] { "value 1", "Value 2"};
        }

        //GET api/values/5
        [HttpGet("{id}")]
        public IActionResult Get(int id, string query)
        {
            return Ok(new Transaction { Id= id, AcountName = "value"+id});
        }


        [HttpPost]
        public IActionResult Post([FromBody] Transaction value)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            return CreatedAtAction("Get", new { id = value.Id }, value);
        }

        
    }
}