using Model;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Model
{
    public class Memory
    {
        public int Id { get; set; }

        [Required, StringLength(500)]
        public string Message { get; set; }

        [Required, StringLength(50)]
        public string MemoryName { get; set; }

        [Required]
        public long UserId { get; set; }

        [ForeignKey("UserId")]
        public User Users { get; set; }
    }
}
