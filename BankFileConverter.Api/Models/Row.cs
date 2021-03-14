using System;

namespace BankFileConverter.Api.Models
{
    public class Row
    {
        public DateTime Date { get; set; }
        public string Label { get; set; }
        public string Amount { get; set; }
    }
}
