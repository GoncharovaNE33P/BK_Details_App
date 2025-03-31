using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BK_Details_App.Models
{
    internal class Groups
    {
        public int GroupIdNumber { get; set; }
        public string Name { get; set; }
        public virtual ICollection<Materials> Materials { get; set; } = new List<Materials>();
    }
}
