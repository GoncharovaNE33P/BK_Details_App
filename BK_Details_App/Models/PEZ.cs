using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BK_Details_App.Models
{
    public partial class PEZ
    {
        public int IdNumber { get; set; }
        public string Mark { get; set; }
        public string Name { get; set; }
        public int Quantity { get; set; }
        public string Color { get; set; }
        public string Matched { get; set; }
    }
}
