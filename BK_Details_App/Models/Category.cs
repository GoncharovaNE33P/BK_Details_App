using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BK_Details_App.Models
{
    public partial class Category
    {
        public int CategoryId { get; set; }
        public string Name { get; set; }
        public int Group { get; set; }
        public virtual Groups GroupNavigation { get; set; } = null!;
        public virtual ICollection<Materials> Materials { get; set; } = new List<Materials>();
    }
}
