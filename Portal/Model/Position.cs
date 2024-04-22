using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Model
{
    public class Position
    {
        public int Id { get; set; }

        [Required, StringLength(50)]
        public string PositionName { get; set; }
    }
}
