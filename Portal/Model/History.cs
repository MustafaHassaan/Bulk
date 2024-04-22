using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Model;

namespace Model
{
    public class History
    {
        public int Id { get; set; }

        [Required, StringLength(15)]
        public string SMSPhone { get; set; }

        [Required, StringLength(250)]
        public string SMSBody { get; set; }

        [Required, StringLength(50)]
        public string Languge { get; set; }

        [Required, StringLength(50)]
        public string Status { get; set; }

        [Required]
        public string Result { get; set; }

        [Required]
        public DateTime Date { get; set; }

        [Required]
        public string Time { get; set; }

        public long UserId { get; set; }
  
        [ForeignKey("UserId")]
        public User Users { get; set; }
    }
}
