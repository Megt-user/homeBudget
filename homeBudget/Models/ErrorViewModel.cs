using System;
using System.ComponentModel.DataAnnotations;

namespace homeBudget.Models
{
    public class ErrorViewModel
    {
        public string RequestId { get; set; }

        [MinLength(3)]
        public bool ShowRequestId => !string.IsNullOrEmpty(RequestId);
    }
}