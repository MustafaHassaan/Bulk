using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Model
{
    public class Customer
    {
        public int Id { get; set; }
        public string CustName { get; set; }
        public string CustPhone { get; set; }
        public string AnotherPhone { get; set; }
        public string LinePhone { get; set; }
        public long UserId { get; set; }

        [ForeignKey("UserId")]
        public User Users { get; set; }
    }
}
