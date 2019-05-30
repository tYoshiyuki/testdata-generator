using System;
using System.Collections.Generic;
using System.Text;

namespace TestdataGenerator.Models
{
    public class Item
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public int? Point { get; set; }
        public DateTime CreateDay { get; set; }
        public DateTime? UpdateDay { get; set; }
    }
}
