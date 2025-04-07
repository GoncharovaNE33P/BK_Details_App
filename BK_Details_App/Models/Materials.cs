using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BK_Details_App.Models
{
    public partial class Materials
    {
        public int IdNumber { get; set; }
        public string Name { get; set; }
        public string Measurement { get; set; }
        public string Analogs { get; set; }
        public string Note { get; set; }
        public int Group { get; set; }
        public int Category { get; set; }
        public virtual Groups GroupNavigation { get; set; } = null!;
        public virtual Category CategoryNavigation { get; set; } = null!;
    }
}
